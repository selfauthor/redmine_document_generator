# frozen_string_literal: true

class DocumentGeneratorController < ApplicationController
  include QueriesHelper
  
  # Находим проект и проверяем права пользователя перед выполнением действий контроллера.
  before_action :find_project
  before_action :authorize_document_generator

  # Проверяем доступность библиотек только перед запуском генерации.
  before_action :check_gems_loaded, only: [:export]

  # GET /projects/:project_id/document_generator/dialog
  # Подготавливает параметры фильтра и отображает диалог генерации документов.
  # Перед подготовкой диалога удаляет папки с результатами, срок хранения которых истёк.
  # @return [void]
  def dialog
    # Удаляем каталоги с результатами, срок хранения которых истёк.
    cleanup_expired_directories

    # Копируем логику из IssuesController#index
    use_session = true

    # Копируем retrieve_default_query
    unless params[:query_id].present? || api_request? || params[:set_filter]
      if params[:without_default].present?
        params[:set_filter] = 1
      elsif !params[:set_filter] && use_session && session[:issue_query]
        query_id, project_id = session[:issue_query].values_at(:id, :project_id)
        unless query_id && project_id == @project&.id && IssueQuery.exists?(id: query_id)
          # continue
        end
      end

      if default_query = IssueQuery.default(project: @project)
        params[:query_id] = default_query.id
      end
    end

    # Используем retrieve_query из QueriesHelper
    retrieve_query(IssueQuery, use_session)

    @record_count = @query.issue_count

    # Сохраняем параметры фильтра для передачи в export
    @filter_params = {
      'f' => @query.filters.keys,
      'op' => @query.filters.transform_values { |v| v[:operator] },
      'v' => @query.filters.transform_values { |v| v[:values] },
      'sort' => @query.sort_criteria.to_param,
      'group_by' => @query.group_by,
      'c' => @query.column_names
    }.compact

    # Добавляем query_id, если есть
    @filter_params['query_id'] = params[:query_id] if params[:query_id].present?

    respond_to do |format|
      format.js
    end
  end


  # POST /projects/:project_id/document_generator/export
  # Генерирует документ или ZIP-архив и возвращает ссылку на защищённое скачивание.
  # @return [void]
  def export
    # Сбрасываем кэш разрешения полей перед каждой генерацией,
    # чтобы изменения пользовательских полей Redmine сразу учитывались.
    DocumentGenerator::FieldResolver.reset_cache!

    @template_file = params[:template_file]
    @export_mode = params[:export_mode]
    @file_name = params[:file_name].to_s.strip
    @error_behavior = params[:error_behavior]

    # Режим Redmine-разметки разрешён только для Word.
    # Для Excel всегда используется обычная текстовая подстановка.
    @description_format =
      params[:description_format].to_s == 'redmine' ? 'redmine' : 'raw'

    @render_warnings = []

    # Проверяем имя файла до запуска генерации.
    if @file_name.blank? || @file_name =~ %r{[\\/:*?"<>|]}
      return render_export_error(I18n.t('document_generator.error_invalid_filename'))
    end

    # Проверяем наличие шаблона и допустимость его расширения.
    unless @template_file && valid_template_extension?(@template_file.original_filename)
      return render_export_error(I18n.t('document_generator.error_invalid_format'))
    end

    # Режим Redmine-разметки разрешён только для Word.
    # Для Excel всегда используется обычная текстовая подстановка.
    if File.extname(@template_file.original_filename).downcase == '.xlsx'
      @description_format = 'raw'
    else
      @description_format =
        params[:description_format].to_s == 'redmine' ? 'redmine' : 'raw'
    end

    # Получаем задачи, выбранные текущими параметрами фильтра.
    data_provider = DocumentGenerator::DataProvider.new(@project, User.current, params)
    @issues = data_provider.fetch_issues

    if @issues.empty?
      return render_export_error(I18n.t('document_generator.error_no_records'))
    end

    # Создаём уникальную папку для промежуточных и итоговых файлов текущей операции.
    export_uuid = SecureRandom.uuid
    export_dir = File.join(export_root_dir, export_uuid)
    FileUtils.mkdir_p(export_dir)

    begin
      # Сохраняем загруженный шаблон внутри блока обработки ошибок,
      # чтобы ошибка превышения допустимого размера корректно возвращалась пользователю.
      template_result = save_uploaded_template(@template_file)
      temp_template_path = template_result[:path]
      temp_template_dir = template_result[:dir]
      
      parser = DocumentGenerator::TemplateParser.new(temp_template_path)
      config = parser.parse

      if @export_mode == 'single'
        # Генерируем отдельные документы и помещаем их в ZIP с исходными именами.
        output_path = File.join(export_dir, "#{export_uuid}.zip")
        generate_single_mode_archive(temp_template_path, config, output_path, export_dir)
        download_filename = "#{@file_name}.zip"
        content_type = 'application/zip'
      else
        # Генерируем единый документ в папке текущей операции.
        ext = File.extname(temp_template_path).downcase
        output_path = File.join(export_dir, "#{export_uuid}#{ext}")
        generate_combined_document(temp_template_path, config, output_path)
        download_filename = "#{@file_name}#{ext}"
        content_type = mime_type_for(ext)
      end

      # Сохраняем метаданные, необходимые для проверки доступа и выдачи файла.
      metadata = {
        uuid: export_uuid,
        project_id: @project.id,
        user_id: User.current.id,
        filename: download_filename,
        content_type: content_type,
        created_at: Time.current.iso8601
      }

      File.write(
        File.join(export_dir, 'metadata.json'),
        JSON.pretty_generate(metadata),
        mode: 'w:UTF-8'
      )

      # Формируем адрес защищённого маршрута скачивания.
      download_url = download_document_generator_export_url(
        project_id: @project.identifier,
        uuid: export_uuid
      )

      touch_working_directory(export_dir)

      # Возвращаем браузеру адрес скачивания и предупреждения генератора.
      render json: {
        type: 'success',
        download_url: download_url,
        filename: download_filename,
        warnings: @render_warnings
      }, status: :ok
    rescue DocumentGenerator::TemplateError, DocumentGenerator::RenderError => e
      # Удаляем незавершённый результат при штатной ошибке генерации.
      FileUtils.rm_rf(export_dir) if export_dir && File.directory?(export_dir)
      render_export_error(e.message)
    rescue StandardError => e
      # Записываем техническую информацию в журнал, а пользователю возвращаем локализованное сообщение.
      Rails.logger.error "[DocumentGenerator] Export failed: #{e.message}\n#{e.backtrace&.join("\n")}"
      FileUtils.rm_rf(export_dir) if export_dir && File.directory?(export_dir)

      short_msg = e.message.to_s.truncate(200)
      render_export_error(
        I18n.t('document_generator.error_render_failed', message: short_msg)
      )
    ensure
      # Удаляем временный каталог шаблона целиком независимо от результата генерации.
      # Это предотвращает накопление пустых директорий в системном /tmp.
      if temp_template_dir && File.directory?(temp_template_dir)
        FileUtils.rm_rf(temp_template_dir)
      end
    end
  end

  # GET /projects/:project_id/document_generator/download/:uuid
  # Проверяет права пользователя и передаёт сформированный файл браузеру.
  # UUID определяет папку результата, а имя скачиваемого файла берётся из metadata.json.
  # @return [void]
  def download
    # Принимаем только UUID ожидаемого формата, исключая передачу произвольного пути.
    uuid = params[:uuid].to_s
    unless uuid.match?(/\A[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}\z/i)
      return render_404
    end

    # Формируем путь только внутри корневой директории выгрузок.
    export_dir = File.join(export_root_dir, uuid)
    metadata_path = File.join(export_dir, 'metadata.json')

    return render_404 unless File.file?(metadata_path)

    # Читаем метаданные результата.
    metadata = JSON.parse(File.read(metadata_path, mode: 'r:UTF-8'))

    # Проверяем, что UUID и проект в метаданных соответствуют адресу запроса.
    return render_404 unless metadata['uuid'] == uuid
    return render_404 unless metadata['project_id'].to_i == @project.id

    # Имя файла в заголовке берётся из метаданных, а не из пользовательского параметра.
    filename = metadata['filename'].to_s
    ext = File.extname(filename).downcase

    # Разрешаем только ожидаемые расширения документов.
    return render_404 unless %w[.docx .xlsx .zip].include?(ext)
    return render_404 if filename.blank? || filename =~ %r{[/:*?"<>|]}

    # Физическое имя формируется по UUID и расширению, без национальных символов.
    file_path = File.join(export_dir, "#{uuid}#{ext}")
    return render_404 unless File.file?(file_path)

    # Передаём файл через Rails, чтобы запрос проходил через авторизацию Redmine.
    send_file file_path,
              filename: filename,
              type: metadata['content_type'].presence || mime_type_for(ext),
              disposition: 'attachment'
  rescue JSON::ParserError, Errno::ENOENT
    # Повреждённые метаданные или отсутствующий файл не должны приводить к выдаче исключения пользователю.
    render_404
  end

  private

  # Возвращает ошибку экспорта в JSON-формате для единого JavaScript-обработчика.
  # @param message [String] Текст ошибки, который нужно показать пользователю.
  # @return [void]
  def render_export_error(message)
    render json: {
      type: 'error',
      message: message
    }, status: :unprocessable_entity
  end

  # Находит проект по числовому ID или строковому идентификатору из URL.
  #
  # Поддержка обоих вариантов нужна для обратной совместимости:
  # существующие ссылки с числовым ID продолжают работать,
  # а новые ссылки используют строковый project.identifier.
  #
  # @return [void]
  def find_project
    project_id = params[:project_id].to_s

    @project =
      if project_id.match?(/\A\d+\z/)
        Project.find_by(id: project_id)
      else
        Project.find_by(identifier: project_id)
      end

    render_404 unless @project
  end

  def authorize_document_generator
    return if @project.module_enabled?(:document_generator) &&
              User.current.allowed_to?(:use_document_generator, @project)
    deny_access
  end

  # Проверяет наличие библиотек, необходимых для генерации документов.
  # При AJAX-запросе возвращает ошибку в JSON, при обычном запросе использует redirect.
  def check_gems_loaded
    return if defined?(::DOCUMENT_GENERATOR_GEMS_LOADED) && ::DOCUMENT_GENERATOR_GEMS_LOADED

    message = I18n.t('error_gems_not_loaded')

    if request.xhr?
      render json: {
        type: 'error',
        message: message
      }, status: :unprocessable_entity
    else
      flash[:error] = message
      redirect_to project_issues_path(@project)
    end
  end

  # Сохраняет загруженный пользователем шаблон во временный каталог.
  #
  # @param uploaded_file [ActionDispatch::Http::UploadedFile] Загруженный файл шаблона.
  # @return [String] Полный путь к сохранённому временному файлу.
  # @raise [DocumentGenerator::TemplateError] Если размер файла превышает допустимый размер вложения.
  def save_uploaded_template(uploaded_file)
    # Получаем максимально допустимый размер файла из глобальной настройки Redmine.
    # Setting.attachment_max_size хранит значение в килобайтах.
    max_size_kb = Setting.attachment_max_size.to_i
    max_size_bytes = max_size_kb.kilobytes

    # Проверяем размер до записи файла во временный каталог.
    # Это позволяет отклонить слишком большой файл до его полного чтения и обработки.
    if uploaded_file.size.to_i > max_size_bytes
      raise DocumentGenerator::TemplateError,
            I18n.t(
              'document_generator.error_template_too_big',
              max_size: max_size_kb
            )
    end

    # Создаём отдельный временный каталог для шаблона.
    temp_dir = Dir.mktmpdir

    # Используем только имя файла без возможного пути, переданного клиентом.
    temp_path = File.join(
      temp_dir,
      File.basename(uploaded_file.original_filename.to_s)
    )

    begin
      # Перемещаем указатель временного файла в начало перед копированием.
      uploaded_file.tempfile.rewind
      # Копируем файл потоком, не загружая всё его содержимое в память Ruby.
      File.open(temp_path, 'wb') do |file|
        IO.copy_stream(uploaded_file.tempfile, file)
      end
      # Возвращаем хэш с путями для корректного последующего удаления всей директории
      { dir: temp_dir, path: temp_path }
    rescue StandardError
      # Если сохранение не удалось, удаляем временный каталог целиком.
      FileUtils.rm_rf(temp_dir)
      raise
    end
  end


  # Генерирует единый документ для всего набора задач и сохраняет его в папке выгрузки.
  # @param template_path [String] Путь к временному файлу шаблона.
  # @param config [Hash] Конфигурация, полученная из парсера шаблона.
  # @param output_path [String] Полный путь для сохранения результата.
  # @return [String] Путь к сформированному документу.
  def generate_combined_document(template_path, config, output_path)
    renderer = create_renderer(template_path, @issues, config)
    result = renderer.render

    # Сохраняем предупреждения renderer для возврата в браузер.
    @render_warnings.concat(renderer.warnings || [])

    if result.is_a?(String) && File.file?(result)
      # Копируем готовый файл в папку операции под UUID-именем.
      FileUtils.cp(result, output_path)
    else
      # Некоторые renderer возвращают объект документа, который требуется записать на диск.
      result.write(output_path)
    end

    output_path
  end

  # Генерирует отдельный документ для каждой задачи и объединяет документы в ZIP-архив.
  # Промежуточные файлы сохраняются под пользовательскими именами до закрытия ZIP.
  # @param template_path [String] Путь к временному файлу шаблона.
  # @param config [Hash] Конфигурация, полученная из парсера шаблона.
  # @param archive_path [String] Полный путь к итоговому ZIP-файлу.
  # @param export_dir [String] Папка текущей операции для промежуточных файлов.
  # @return [String] Путь к созданному ZIP-архиву.
  def generate_single_mode_archive(template_path, config, archive_path, export_dir)

    ext = File.extname(template_path).downcase
    intermediate_paths = []

    Zip::File.open(archive_path, Zip::File::CREATE) do |zipfile|
      @issues.each do |issue|
        renderer = create_renderer(template_path, [issue], config)
        result = renderer.render

        # Накапливаем предупреждения по каждой задаче, вошедшей в архив.
        @render_warnings.concat(renderer.warnings || [])

        # Используем пользовательское имя файла с добавлением ID задачи.
        filename_in_zip = "#{@file_name}#{issue.id}#{ext}"
        file_path = File.join(export_dir, filename_in_zip)

        if result.is_a?(String) && File.file?(result)
          # Копируем результат renderer в рабочую папку под требуемым именем.
          FileUtils.cp(result, file_path)
        else
          # Сохраняем объект документа под пользовательским именем.
          result.write(file_path)
        end

        intermediate_paths << file_path

        # Добавляем документ в архив под тем же именем, которое задано для записи.
        zipfile.add(filename_in_zip, file_path)
      end
    end

    # Удаляем промежуточные документы только после закрытия ZIP-файла.
    intermediate_paths.each do |file_path|
      FileUtils.rm_f(file_path)
    end

    archive_path
  end

  # Создаёт renderer в соответствии с расширением шаблона.
  #
  # @param template_path [String] Путь к шаблону.
  # @param issues [Array<Issue>] Задачи для обработки.
  # @param config [Hash] Конфигурация шаблона.
  # @return [Object] Экземпляр WordRenderer или ExcelRenderer.
  def create_renderer(template_path, issues, config)
    ext = File.extname(template_path).downcase

    renderer_class =
      case ext
      when '.docx'
        DocumentGenerator::WordRenderer
      when '.xlsx'
        DocumentGenerator::ExcelRenderer
      else
        raise DocumentGenerator::TemplateError,
              I18n.t('document_generator.error_invalid_format')
      end

    if ext == '.docx'
      renderer_class.new(
        template_path,
        issues,
        config,
        @error_behavior,
        @description_format
      )
    else
      renderer_class.new(
        template_path,
        issues,
        config,
        @error_behavior
      )
    end
  end

  def mime_type_for(ext)
    case ext.downcase
    when '.docx' then 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    when '.xlsx' then 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    else 'application/octet-stream'
    end
  end

  # Проверяет, что расширение файла шаблона поддерживается (.docx или .xlsx)
  # @param filename [String] Имя загруженного файла
  # @return [Boolean] true, если расширение допустимо
  def valid_template_extension?(filename)
    ext = File.extname(filename).downcase
    %w[.docx .xlsx].include?(ext)
  end

  # Возвращает корневую папку, в которой хранятся результаты генерации.
  # @return [String] Абсолютный путь к tmp/document_generator.
  def export_root_dir
    Rails.root.join('tmp', 'document_generator').to_s
  end

  # Внутренний метод контроллера.
  # Удаляет каталоги результатов, чей mtime старше заданного срока хранения.
  # Срок хранения задаётся в часах в настройках плагина.
  # @return [void]
  def cleanup_expired_directories
    # Каталог, в котором хранятся рабочие папки генератора документов.
    storage_root = Rails.root.join('tmp', 'document_generator')

    # Если каталог ещё не создан, очищать нечего.
    return unless Dir.exist?(storage_root)

    # Получаем срок хранения из настроек плагина; по умолчанию — 24 часа.
    retention_hours = Setting.plugin_redmine_document_generator
                             .fetch('retention_hours', '24')
                             .to_i

    # Защищаемся от некорректного или отрицательного значения настройки.
    retention_hours = 24 if retention_hours <= 0

    # Вычисляем предельное время: каталоги старше этой отметки подлежат удалению.
    expiration_time = Time.current - retention_hours.hours

    # Обрабатываем только непосредственные подкаталоги хранилища.
    Dir.children(storage_root).each do |entry|
      directory_path = storage_root.join(entry)

      # Не удаляем файлы и символические ссылки.
      next unless File.directory?(directory_path) && !File.symlink?(directory_path)

      # mtime каталога используется как единственный критерий срока хранения.
      next unless File.mtime(directory_path) < expiration_time

      # Удаляем весь каталог вместе с его содержимым.
      FileUtils.remove_entry(directory_path)
    rescue StandardError => e
      # Ошибка удаления одного каталога не должна останавливать очистку остальных.
      Rails.logger.error(
        "[DocumentGenerator] Failed to remove expired directory " \
        "#{directory_path}: #{e.message}"
      )
    end
  end

  # Внутренний метод контроллера.
  # Обновляет mtime рабочего каталога после завершения формирования результата.
  # @param working_dir [String] Путь к рабочему каталогу текущей операции.
  # @return [void]
  def touch_working_directory(working_dir)
    # Обновляем mtime каталога, чтобы срок хранения отсчитывался от завершения генерации.
    FileUtils.touch(working_dir)
  end

end
# v2610090944