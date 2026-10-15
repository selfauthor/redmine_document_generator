# frozen_string_literal: true

class DocumentGeneratorController < ApplicationController
  include QueriesHelper
  
  # Находим проект и проверяем права пользователя перед выполнением действий контроллера.
  before_action :find_project
  before_action :authorize_document_generator

  # Проверяем доступность библиотек только перед запуском генерации.
  before_action :check_gems_loaded, only: [:export]


  # GET /projects/:project_id/document_generator/dialog
  # Проверяет JSON-снимок и подготавливает данные для диалога выгрузки.
  # @return [void]
  def dialog
    # Удаляем просроченные результаты и снимки по существующему сроку хранения.
    cleanup_expired_directories
    DocumentGenerator::IssueListSnapshot.cleanup_expired!

    # Сохраняем UUID снимка для передачи в форму экспорта.
    @issue_list_token = params[:issue_list_token].to_s

    # Читаем снимок при открытии окна, а не во время запуска экспорта.
    issue_snapshot = DocumentGenerator::IssueListSnapshot.read(
      uuid: @issue_list_token,
      project_id: @project.id,
      user_id: User.current.id
    )

    if issue_snapshot
      # Используем число ID из снимка, а не повторно вычисленный счётчик IssueQuery.
      @record_count = issue_snapshot['issue_ids'].length
      @snapshot_error = nil
    else
      # Показываем ошибку сразу, если снимок отсутствует или недействителен.
      @record_count = 0
      @snapshot_error = I18n.t('document_generator.error_issue_list_expired')
    end

    # Получаем доступные результаты предыдущих выгрузок текущего проекта.
    @available_files = available_export_files

    respond_to do |format|
      format.js
    end
  end


  # POST /projects/:project_id/document_generator/export
  # Генерирует документ или ZIP-архив и возвращает ссылку на защищённое скачивание.
  # @return [void]
  def export
    # Сбрасываем кэш разрешения полей перед каждой генерацией.
    DocumentGenerator::FieldResolver.reset_cache!

    @template_file = params[:template_file]
    @export_mode = params[:export_mode]
    @file_name = params[:file_name].to_s.strip
    @error_behavior = params[:error_behavior]
    @render_warnings = []

    temp_template_dir = nil
    export_dir = nil
    progress_tracker = nil
    @progress_tracker = nil
    heartbeat_running = false
    heartbeat_thread = nil

    # Проверяем имя файла до запуска генерации.
    if @file_name.blank? || @file_name =~ %r{[\\/:*?"<>|]}
      return render_export_error(
        I18n.t('document_generator.error_invalid_filename')
      )
    end

    # Проверяем наличие шаблона и допустимость его расширения.
    unless @template_file &&
           valid_template_extension?(@template_file.original_filename)
      return render_export_error(
        I18n.t('document_generator.error_invalid_format')
      )
    end

    # Принимаем только корректный UUID операции, переданный браузером.
    export_uuid = params[:progress_uuid].to_s

    unless export_uuid.match?(
      /\A[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}\z/i
    )
      return render_export_error(
        I18n.t('document_generator.error_invalid_progress_id')
      )
    end

    # Для Excel всегда используется обычная текстовая подстановка.
    ext = File.extname(@template_file.original_filename).downcase
    @description_format =
      ext == '.docx' && params[:description_format].to_s == 'redmine' ?
        'redmine' : 'raw'

    # Файл состояния хранится отдельно от каталога результата.
    progress_tracker = DocumentGenerator::ProgressTracker.new(
      progress_state_path(export_uuid),
      project_id: @project.id,
      user_id: User.current.id
    )
    @progress_tracker = progress_tracker

    # Heartbeat обновляет состояние во время длительных операций,
    # в которых нет доступных контрольных точек.
    heartbeat_running = true

    heartbeat_thread = Thread.new do
      while heartbeat_running
        sleep 1
        progress_tracker.tick if heartbeat_running
      end
    rescue StandardError => e
      Rails.logger.warn(
        "[DocumentGenerator] Progress heartbeat stopped: #{e.message}"
      )
    end

    begin
      # Первый этап: получение выборки задач.
      progress_tracker.set_phase(
        stage: 'preparing',
        min: 0,
        max: @export_mode == 'single' ? 5 : 10,
        message: I18n.t('document_generator.progress_wait')
      )

      # Повторно читаем JSON-снимок непосредственно перед подготовкой документа.
      issue_snapshot = DocumentGenerator::IssueListSnapshot.read(
        uuid: params[:issue_list_token],
        project_id: @project.id,
        user_id: User.current.id
      )

      # Если снимок удалён, повреждён или просрочен, прекращаем экспорт.
      unless issue_snapshot
        raise DocumentGenerator::RenderError,
              I18n.t('document_generator.error_issue_list_expired')
      end

      # Загружаем существующие и доступные пользователю задачи
      # в порядке, сохранённом в JSON-снимке.
      @issues = fetch_issues_from_snapshot(issue_snapshot['issue_ids'])

      # Прерываем генерацию, если доступных задач не осталось.
      if @issues.empty?
        raise DocumentGenerator::RenderError,
              I18n.t('document_generator.error_no_records')
      end

      # Не пересчитываем фильтры, если снимок отсутствует или недействителен.
      unless issue_snapshot
        raise DocumentGenerator::RenderError,
              I18n.t('document_generator.error_issue_list_expired')
      end

      # Загружаем актуальные задачи в порядке ID, сохранённом в снимке.
      # Удалённые и ставшие недоступными пользователю задачи будут пропущены.
      @issues = fetch_issues_from_snapshot(issue_snapshot['issue_ids'])

      # Прерываем генерацию, если доступных задач не осталось.
      if @issues.empty?
        raise DocumentGenerator::RenderError,
              I18n.t('document_generator.error_no_records')
      end

      # Создаём каталог промежуточных и итоговых файлов операции.
      export_dir = File.join(export_root_dir, export_uuid)
      FileUtils.mkdir_p(export_dir)

      # Сохраняем шаблон и разбираем его конфигурацию.
      template_result = save_uploaded_template(@template_file)
      temp_template_path = template_result[:path]
      temp_template_dir = template_result[:dir]

      parser = DocumentGenerator::TemplateParser.new(temp_template_path)
      config = parser.parse

      if @export_mode == 'single'
        # При отдельных документах подготовка контекста и рендеринг
        # чередуются для каждой задачи.
        progress_tracker.set_phase(
          stage: 'preparing',
          min: 5,
          max: 95,
          message: I18n.t('document_generator.progress_wait'),
          percent: 5
        )

        output_path = File.join(export_dir, "#{export_uuid}.zip")

        generate_single_mode_archive(
          temp_template_path,
          config,
          output_path,
          export_dir
        )

        download_filename = "#{@file_name}.zip"
        content_type = 'application/zip'
      else
        # Для единого документа подготовка контекста занимает
        # диапазон 10–50%, а формирование файла — диапазон 50–95%.
        progress_tracker.set_phase(
          stage: 'preparing',
          min: 10,
          max: 90,
          message: I18n.t('document_generator.progress_wait'),
          percent: 10
        )

        @current_progress_callback = lambda do |event, fraction|
          case event
          when :context_progress
            progress_tracker.update_fraction(fraction)

          when :context_complete
            progress_tracker.set_phase(
              stage: 'rendering',
              min: 90,
              max: 95,
              message: I18n.t('document_generator.progress_wait'),
              percent: 50
            )

          when :render_complete
            progress_tracker.set_phase(
              stage: 'rendering',
              min: 90,
              max: 95,
              message: I18n.t('document_generator.progress_wait'),
              percent: 95
            )
          end
        end

        ext = File.extname(temp_template_path).downcase
        output_path = File.join(export_dir, "#{export_uuid}#{ext}")

        generate_combined_document(
          temp_template_path,
          config,
          output_path
        )

        download_filename = "#{@file_name}#{ext}"
        content_type = mime_type_for(ext)
      end

      # Сохраняем метаданные для проверки доступа и выдачи файла.
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

      # Сначала фиксируем успешное завершение, затем возвращаем результат браузеру.
      progress_tracker.finish(
        status: 'success',
        message: I18n.t('document_generator.progress_complete')
      )

      render json: {
        type: 'success',
        download_url: download_url,
        filename: download_filename,
        warnings: @render_warnings
      }, status: :ok

    rescue DocumentGenerator::TemplateError,
           DocumentGenerator::RenderError => e
      # Удаляем незавершённый результат, сохраняя состояние ошибки для браузера.
      FileUtils.rm_rf(export_dir) if export_dir && File.directory?(export_dir)

      progress_tracker&.finish(status: 'error', message: e.message)
      render_export_error(e.message)

    rescue StandardError => e
      # Технические подробности сохраняем в журнале Redmine.
      Rails.logger.error(
        "[DocumentGenerator] Export failed: #{e.message}\n#{e.backtrace&.join("\n")}"
      )

      FileUtils.rm_rf(export_dir) if export_dir && File.directory?(export_dir)

      error_message = I18n.t(
        'document_generator.error_render_failed',
        message: e.message.to_s.truncate(200)
      )

      progress_tracker&.finish(
        status: 'error',
        message: error_message
      )

      render_export_error(error_message)

    ensure
      # Останавливаем heartbeat при любом результате операции.
      heartbeat_running = false
      heartbeat_thread&.join(1.5)

      # Удаляем временный каталог шаблона независимо от результата.
      if temp_template_dir && File.directory?(temp_template_dir)
        FileUtils.rm_rf(temp_template_dir)
      end
    end
  end

  # GET /projects/:project_id/document_generator/progress/:uuid
  # Возвращает состояние операции только запустившему её пользователю.
  # @return [void]
  def progress
    # Проверяем UUID до формирования пути, исключая произвольные пути к файлам.
    uuid = params[:uuid].to_s

    unless uuid.match?(
      /\A[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}\z/i
    )
      return render_404
    end

    state_path = progress_state_path(uuid)
    return render_404 unless File.file?(state_path)

    # Разделяемая блокировка ожидает окончания записи состояния.
    state = DocumentGenerator::ProgressTracker.read(state_path)

    return render_404 unless state['uuid'] == uuid
    return render_404 unless state['project_id'].to_i == @project.id
    return render_404 unless state['user_id'].to_i == User.current.id

    render json: state, status: :ok

  rescue JSON::ParserError, Errno::ENOENT
    # При неполном состоянии браузер повторит запрос через секунду.
    render json: { status: 'pending' }, status: :accepted
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

  # Загружает доступные пользователю задачи в порядке, сохранённом в снимке.
  #
  # @param issue_ids [Array<Integer>] Упорядоченные ID задач из JSON-снимка.
  # @return [Array<Issue>] Доступные задачи с загруженными связанными объектами.
  def fetch_issues_from_snapshot(issue_ids)
    # Сохраняем исходный порядок и исключаем случайные дубликаты.
    ordered_ids = issue_ids.map(&:to_i).uniq
    return [] if ordered_ids.empty?

    # Загружаем только задачи, доступные текущему пользователю.
    # Связи нужны для последующей подготовки контекста генератора.
    issues = Issue.visible(User.current)
                  .where(id: ordered_ids)
                  .includes(
                    :status,
                    :priority,
                    :author,
                    :assigned_to,
                    :project,
                    :tracker,
                    :category,
                    :fixed_version,
                    :parent,
                    custom_values: :custom_field,
                    watchers: :user
                  )
                  .to_a

    # Восстанавливаем порядок снимка, поскольку SQL-запрос where(id: ...)
    # сам по себе не гарантирует порядок элементов в результате.
    issues_by_id = issues.index_by(&:id)
    ordered_issues = ordered_ids.filter_map do |issue_id|
      issues_by_id[issue_id]
    end

    # Загружаем связи задач, используемые стандартными механизмами Redmine.
    Issue.load_relations(ordered_issues) if ordered_issues.any?

    ordered_issues
  end

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
      @issues.each_with_index do |issue, issue_index|
        
        # Выделяем для каждого документа отдельный диапазон прогресса.
        # Первая половина диапазона отведена подготовке контекста, вторая — рендерингу.
        segment_min = 5 + (90.0 * issue_index / @issues.length)
        segment_mid = segment_min + (45.0 / @issues.length)
        segment_max = segment_min + (90.0 / @issues.length)

        # Начинаем диапазон подготовки контекста текущего документа.
        @progress_tracker&.set_phase(
          stage: 'preparing',
          min: segment_min.floor,
          max: segment_mid.floor,
          message: I18n.t('document_generator.progress_wait'),
          force: false
        )

        @current_progress_callback = lambda do |event, fraction|
          case event
          when :context_progress
            @progress_tracker&.update_fraction(fraction)

          when :context_complete
            @progress_tracker&.set_phase(
              stage: 'rendering',
              min: segment_mid.floor,
              max: segment_max.floor,
              message: I18n.t('document_generator.progress_wait'),
              percent: segment_mid.floor,
              force: false
            )

          when :render_complete
            @progress_tracker&.set_phase(
              stage: 'rendering',
              min: segment_mid.floor,
              max: segment_max.floor,
              message: I18n.t('document_generator.progress_wait'),
              percent: segment_max.floor,
              force: false
            )
          end
        end

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
        @description_format,
        @current_progress_callback
      )
    else
      renderer_class.new(
        template_path,
        issues,
        config,
        @error_behavior,
        @current_progress_callback
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

  # Возвращает путь к файлу состояния конкретной операции.
  # @param uuid [String] UUID операции генерации.
  # @return [String] Абсолютный путь к JSON-файлу состояния.
  def progress_state_path(uuid)
    File.join(export_root_dir, ".progress-#{uuid}.json")
  end

  # Возвращает корневую папку, в которой хранятся результаты генерации.
  # @return [String] Абсолютный путь к tmp/document_generator.
  def export_root_dir
    Rails.root.join('tmp', 'document_generator').to_s
  end

  # Возвращает список существующих результатов выгрузки текущего проекта.
  # Учитываются только корректные метаданные, допустимые UUID и реально существующие файлы.
  # @return [Array<Hash>] Список результатов с именем файла, временем создания,
  #   пользователем и URL защищённого скачивания.
  def available_export_files
    # Каталог может отсутствовать, если генератор ещё ни разу не запускался.
    storage_root = export_root_dir
    return [] unless Dir.exist?(storage_root)

    files = []

    # Просматриваем только непосредственные подкаталоги хранилища.
    Dir.children(storage_root).each do |entry|
      # Имя каталога должно быть UUID, чтобы исключить обработку произвольных путей.
      next unless entry.match?(
        /\A[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}\z/i
      )

      export_dir = File.join(storage_root, entry)

      # Не обрабатываем файлы и символические ссылки вместо каталогов.
      next unless File.directory?(export_dir) && !File.symlink?(export_dir)

      metadata_path = File.join(export_dir, 'metadata.json')

      # Пропускаем каталоги без метаданных.
      next unless File.file?(metadata_path)

      begin
        # Загружаем метаданные завершённой выгрузки.
        metadata = JSON.parse(File.read(metadata_path, mode: 'r:UTF-8'))

        # Метаданные должны содержать объект с ожидаемым UUID и ID проекта.
        next unless metadata.is_a?(Hash)
        next unless metadata['uuid'] == entry
        next unless metadata['project_id'].to_i == @project.id

        # Проверяем имя файла и разрешённое расширение.
        filename = metadata['filename'].to_s
        extension = File.extname(filename).downcase

        next if filename.blank? || filename =~ %r{[\\/:*?"<>|]}
        next unless %w[.docx .xlsx .zip].include?(extension)

        # Физическое имя результата формируется по UUID и расширению.
        file_path = File.join(export_dir, "#{entry}#{extension}")

        # Не показываем ссылки на отсутствующие файлы.
        next unless File.file?(file_path)

        # Преобразуем время из metadata.json в объект времени Rails.
        created_at = Time.zone.parse(metadata['created_at'].to_s)
        next unless created_at

        # Получаем автора выгрузки; запись пользователя могла быть удалена.
        user = User.find_by(id: metadata['user_id'])
        next unless user

        files << {
          filename: filename,
          created_at: created_at,
          user: user,
          url: download_document_generator_export_path(
            project_id: @project.identifier,
            uuid: entry
          )
        }
      rescue JSON::ParserError, ArgumentError, TypeError, Errno::ENOENT => e
        # Повреждение одной папки не должно мешать отображению остальных результатов.
        Rails.logger.warn(
          "[DocumentGenerator] Skipping invalid export metadata " \
          "#{metadata_path}: #{e.message}"
        )
        next
      end
    end

    # Сначала показываем последние созданные документы.
    files.sort_by { |file| file[:created_at] }.reverse
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

    # Удаляем устаревшие файлы состояния по тому же сроку хранения, что и результаты.
    Dir.children(storage_root).each do |entry|
      state_path = storage_root.join(entry)

      next unless entry.match?(/\A\.progress-[0-9a-f-]+\.json\z/i)
      next unless File.file?(state_path) && !File.symlink?(state_path)
      next unless File.mtime(state_path) < expiration_time

      FileUtils.rm_f(state_path)
    rescue StandardError => e
      Rails.logger.warn(
        "[DocumentGenerator] Failed to remove expired progress state " \
        "#{state_path}: #{e.message}"
      )
    end

    # Обрабатываем только каталоги результатов с именами в формате UUID.
    # Каталог issue_lists очищается отдельным механизмом JSON-снимков.
    Dir.children(storage_root).each do |entry|
      next unless entry.match?(
        /\A[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}\z/i
      )

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
# v2610141450