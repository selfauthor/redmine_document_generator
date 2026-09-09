# frozen_string_literal: true

class DocumentGeneratorController < ApplicationController
  before_action :find_project
  before_action :authorize_document_generator
  before_action :check_gems_loaded, only: [:export]

  # GET /projects/:project_id/document_generator/dialog
  # Открывает модальное окно с подсчётом записей по текущему фильтру
  def dialog
    @query = IssueQuery.new(name: '_document_generator_temp', project: @project)
    @query.build_from_params(params)
    @record_count = @query.issue_count

    respond_to do |format|
      format.js
    end
  end

  # POST /projects/:project_id/document_generator/export
  # Основной метод генерации и скачивания документа
  def export
    @template_file = params[:template_file]
    @export_mode = params[:export_mode]
    @file_name = params[:file_name]
    @error_behavior = params[:error_behavior]

    # Валидация имени файла на недопустимые символы
    if @file_name =~ /[/\\:*?"<>|]/
      flash[:error] = I18n.t('document_generator.error_invalid_filename')
      redirect_back(fallback_location: project_issues_path(@project)) and return
    end

    # Получение выборки записей с учётом фильтра и прав доступа
    data_provider = DocumentGenerator::DataProvider.new(@project, User.current, params)
    @issues = data_provider.fetch_issues

    if @issues.empty?
      flash[:error] = I18n.t('document_generator.error_no_records')
      redirect_back(fallback_location: project_issues_path(@project)) and return
    end

    # Сохранение загруженного файла шаблона во временную директорию
    temp_template_path = save_uploaded_template(@template_file)

    begin
      # Парсинг шаблона: определение типа, блоков, полей, извлечение текста
      parser = DocumentGenerator::TemplateParser.new(temp_template_path)
      config = parser.parse

      if @export_mode == 'single'
        # Режим "Один файл на запись": генерируем ZIP-архив
        archive_path = generate_single_mode_archive(temp_template_path, config)
        send_file archive_path,
                  filename: "#{@file_name}.zip",
                  type: 'application/zip',
                  disposition: 'attachment'
      else
        # Режим "Единый документ": один файл
        output_path = generate_combined_document(temp_template_path, config)
        ext = File.extname(temp_template_path)
        send_file output_path,
                  filename: "#{@file_name}#{ext}",
                  type: mime_type_for(ext),
                  disposition: 'attachment'
      end

    rescue StandardError => e
      Rails.logger.error("[DocumentGenerator] Export failed: #{e.message}\n#{e.backtrace.join("\n")}")
      flash[:error] = I18n.t('document_generator.error_invalid_template', message: e.message)
      redirect_back(fallback_location: project_issues_path(@project))
    ensure
      # Гарантированная очистка временного файла шаблона
      FileUtils.rm_f(temp_template_path) if temp_template_path && File.exist?(temp_template_path)
    end
  end

  private

  def find_project
    @project = Project.find(params[:project_id])
  rescue ActiveRecord::RecordNotFound
    render_404
  end

  def authorize_document_generator
    return if @project.module_enabled?(:document_generator) &&
              User.current.allowed_to?(:use_document_generator, @project)
    deny_access
  end

  def check_gems_loaded
    return if defined?(DOCUMENT_GENERATOR_GEMS_LOADED) && DOCUMENT_GENERATOR_GEMS_LOADED
    flash[:error] = I18n.t('document_generator.error_gems_not_loaded')
    redirect_to project_issues_path(@project)
  end

  # Сохраняет загруженный файл во временную директорию
  # @param uploaded_file [ActionDispatch::Http::UploadedFile] Загруженный файл
  # @return [String] Путь к временному файлу
  def save_uploaded_template(uploaded_file)
    temp_dir = Dir.mktmpdir
    temp_path = File.join(temp_dir, uploaded_file.original_filename)
    File.open(temp_path, 'wb') { |f| f.write(uploaded_file.read) }
    temp_path
  end

  # Генерирует единый документ (режим combined)
  # @param template_path [String] Путь к шаблону
  # @param config [Hash] Конфигурация из TemplateParser
  # @return [String] Путь к сгенерированному файлу
  def generate_combined_document(template_path, config)
    renderer = create_renderer(template_path, @issues, config)
    result = renderer.render

    # WordRenderer возвращает объект Sablon, ExcelRenderer — путь к файлу.
    # Унифицируем: если результат не строка (путь), записываем его в файл.
    if result.is_a?(String) && File.exist?(result)
      result
    else
      output_path = File.join(Dir.mktmpdir, "output#{File.extname(template_path)}")
      result.write(output_path)
      output_path
    end
  end

  # Генерирует ZIP-архив с отдельным файлом для каждой записи (режим single)
  # @param template_path [String] Путь к шаблону
  # @param config [Hash] Конфигурация из TemplateParser
  # @return [String] Путь к ZIP-архиву
  def generate_single_mode_archive(template_path, config)
    require 'zip'

    archive_path = File.join(Dir.mktmpdir, "#{@file_name}.zip")
    ext = File.extname(template_path)

    Zip::File.open(archive_path, Zip::File::CREATE) do |zipfile|
      @issues.each do |issue|
        # Для каждой записи создаём отдельный рендерер с массивом из одной задачи
        renderer = create_renderer(template_path, [issue], config)
        result = renderer.render

        # Унифицируем результат (объект Sablon или путь к файлу)
        if result.is_a?(String) && File.exist?(result)
          file_path = result
        else
          file_path = File.join(Dir.mktmpdir, "temp#{ext}")
          result.write(file_path)
        end

        filename_in_zip = "#{@file_name}_#{issue.id}#{ext}"
        zipfile.add(filename_in_zip, file_path)
        FileUtils.rm_f(file_path)
      end
    end

    archive_path
  end

  # Создаёт экземпляр рендерера в зависимости от расширения шаблона
  # @param template_path [String] Путь к шаблону
  # @param issues [Array<Issue>] Массив задач
  # @param config [Hash] Конфигурация
  # @return [WordRenderer, ExcelRenderer] Экземпляр рендерера
  def create_renderer(template_path, issues, config)
    ext = File.extname(template_path).downcase
    renderer_class = case ext
                     when '.docx' then DocumentGenerator::WordRenderer
                     when '.xlsx' then DocumentGenerator::ExcelRenderer
                     else
                       raise I18n.t('document_generator.error_invalid_format')
                     end

    renderer_class.new(template_path, issues, config, @error_behavior)
  end

  # Возвращает MIME-тип для расширения файла
  # @param ext [String] Расширение файла (например, '.docx')
  # @return [String] MIME-тип
  def mime_type_for(ext)
    case ext.downcase
    when '.docx' then 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    when '.xlsx' then 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    else 'application/octet-stream'
    end
  end
end