# frozen_string_literal: true

class DocumentGeneratorController < ApplicationController
  # Перед каждым действием находим проект по ID из параметров
  before_action :find_project
  
  # Проверяем право пользователя на использование модуля в данном проекте
  before_action :authorize_document_generator
  
  # Проверяем загрузку необходимых gem-ов только для действия export
  before_action :check_gems_loaded, only: [:export]

  # GET /projects/:project_id/document_generator/dialog
  # Отображает модальное окно с параметрами выгрузки и подсчитывает количество записей
  def dialog
    @query = IssueQuery.new(name: '_temp', project: @project)
    @query.build_from_params(params)
    @record_count = @query.issue_count

    respond_to do |format|
      format.js
    end
  end

  # POST /projects/:project_id/document_generator/export
  # Основной метод обработки запроса на генерацию документа
  def export
    @template_file = params[:template_file]
    @export_mode = params[:export_mode]
    @file_name = params[:file_name]
    @error_behavior = params[:error_behavior] || 'abort'

    # Валидация имени файла на наличие недопустимых символов файловой системы
    if @file_name =~ /[\/:*?"<>|]/
      flash[:error] = I18n.t('document_generator.error_invalid_filename')
      redirect_back(fallback_location: project_issues_path(@project)) and return
    end

    # Проверка расширения загруженного файла
    ext = File.extname(@template_file.original_filename).downcase
    unless ['.docx', '.doc'].include?(ext)
      flash[:error] = I18n.t('document_generator.error_invalid_format')
      redirect_back(fallback_location: project_issues_path(@project)) and return
    end

    # Получение данных из БД с учетом прав доступа и фильтров
    data_provider = DocumentGenerator::DataProvider.new(@project, User.current, params)
    issues = data_provider.fetch_issues

    # Если после применения прав доступа записей не осталось, прерываем выполнение
    if issues.empty?
      flash[:error] = I18n.t('document_generator.error_no_records')
      redirect_back(fallback_location: project_issues_path(@project)) and return
    end

    # Сохраняем загруженный шаблон во временный файл для обработки
    temp_template = Tempfile.new(['template', ext])
    temp_template.binmode
    temp_template.write(@template_file.read)
    temp_template.close

    begin
      # Парсим структуру шаблона
      parser = DocumentGenerator::TemplateParser.new(temp_template.path)
      parser_config = parser.parse

      # Пока реализован только рендеринг Word (Этап 4)
      if ext == '.docx'
        renderer = DocumentGenerator::WordRenderer.new(temp_template.path, issues, parser_config, @error_behavior)
        doc_content = renderer.render
        
        # Отправляем сгенерированный файл пользователю
        send_data doc_content, 
                  filename: "#{@file_name}.docx", 
                  type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                  disposition: 'attachment'
      else
        # Заглушка для Этапа 5 (Excel) с использованием локализованного сообщения
        flash[:error] = I18n.t('document_generator.xlsx_not_implemented')
        redirect_back(fallback_location: project_issues_path(@project))
      end
    rescue => e
      # Логирование ошибки для разработчика
      Rails.logger.error "[DocumentGenerator] Export error: #{e.message}\n#{e.backtrace.join("\n")}"
      
      # Обработка ошибки в зависимости от выбранной пользователем стратегии
      if @error_behavior == 'abort'
        flash[:error] = I18n.t('document_generator.error_invalid_template', message: e.message)
        redirect_back(fallback_location: project_issues_path(@project))
      else
        flash[:error] = I18n.t('document_generator.error_processing_skipped', message: e.message)
        redirect_back(fallback_location: project_issues_path(@project))
      end
    ensure
      # Гарантированное удаление временного файла шаблона
      temp_template.unlink
    end
  end

  private

  # Находит проект по параметру project_id или возвращает ошибку 404
  def find_project
    @project = Project.find(params[:project_id])
  rescue ActiveRecord::RecordNotFound
    render_404
  end

  # Проверяет, включен ли модуль в проекте и есть ли у пользователя соответствующее право
  def authorize_document_generator
    return if @project.module_enabled?(:document_generator) &&
              User.current.allowed_to?(:use_document_generator, @project)
    deny_access
  end

  # Проверяет, были ли успешно загружены необходимые gem-библиотеки при старте плагина
  def check_gems_loaded
    return if defined?(DOCUMENT_GENERATOR_GEMS_LOADED) && DOCUMENT_GENERATOR_GEMS_LOADED
    flash[:error] = I18n.t(:error_gems_not_loaded)
    redirect_to project_issues_path(@project)
  end
end