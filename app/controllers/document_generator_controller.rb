# frozen_string_literal: true

# v2609111239
class DocumentGeneratorController < ApplicationController
  include QueriesHelper
  
  before_action :find_project
  before_action :authorize_document_generator
  before_action :check_gems_loaded, only: [:export]

  def dialog
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
  # Основной метод генерации и скачивания документа
  def export
    @template_file = params[:template_file]
    @export_mode = params[:export_mode]
    @file_name = params[:file_name]
    @error_behavior = params[:error_behavior]

    # Валидация имени файла на недопустимые символы
    # Используем %r{}, чтобы избежать конфликта с символом / внутри регулярного выражения
    if @file_name =~ %r{[/:*?"<>|]}
      flash[:error] = I18n.t('document_generator.error_invalid_filename')
      redirect_back(fallback_location: project_issues_path(@project)) and return
    end

    # Валидация формата файла шаблона (только .docx и .xlsx)
    unless @template_file && valid_template_extension?(@template_file.original_filename)
      flash[:error] = I18n.t('document_generator.error_invalid_format')
      redirect_back(fallback_location: project_issues_path(@project)) and return
    end

    # Получение выборки записей с учётом фильтра и прав доступа
    data_provider = DocumentGenerator::DataProvider.new(@project, User.current, params)
    @issues = data_provider.fetch_issues

    if @issues.empty?
      flash[:error] = I18n.t('document_generator.error_no_records')
      redirect_back(fallback_location: project_issues_path(@project)) and return
    end

    temp_template_path = save_uploaded_template(@template_file)

    begin
      parser = DocumentGenerator::TemplateParser.new(temp_template_path)
      config = parser.parse

      if @export_mode == 'single'
        archive_path = generate_single_mode_archive(temp_template_path, config)
        send_file archive_path,
                  filename: "#{@file_name}.zip",
                  type: 'application/zip',
                  disposition: 'attachment'
      else
        output_path = generate_combined_document(temp_template_path, config)
        ext = File.extname(temp_template_path)
        send_file output_path,
                  filename: "#{@file_name}#{ext}",
                  type: mime_type_for(ext),
                  disposition: 'attachment'
      end

    rescue DocumentGenerator::TemplateError, DocumentGenerator::RenderError => e
      Rails.logger.error("[DocumentGenerator] Export failed: #{e.message}")
      flash[:error] = e.message
      redirect_back(fallback_location: project_issues_path(@project))
    rescue StandardError => e
      # Логируем полную ошибку с backtrace в журнал
      Rails.logger.error("[DocumentGenerator] Unexpected error: #{e.message}\n#{e.backtrace&.join("\n")}")
      # В flash записываем только короткое сообщение, чтобы избежать CookieOverflow
      short_msg = e.message.to_s.truncate(200)
      flash[:error] = I18n.t('document_generator.error_render_failed', message: short_msg)
      redirect_back(fallback_location: project_issues_path(@project))
    ensure
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

  def save_uploaded_template(uploaded_file)
    temp_dir = Dir.mktmpdir
    temp_path = File.join(temp_dir, uploaded_file.original_filename)
    File.open(temp_path, 'wb') { |f| f.write(uploaded_file.read) }
    temp_path
  end

  def generate_combined_document(template_path, config)
    renderer = create_renderer(template_path, @issues, config)
    result = renderer.render

    if result.is_a?(String) && File.exist?(result)
      result
    else
      output_path = File.join(Dir.mktmpdir, "output#{File.extname(template_path)}")
      result.write(output_path)
      output_path
    end
  end

  def generate_single_mode_archive(template_path, config)
    require 'zip'

    archive_path = File.join(Dir.mktmpdir, "#{@file_name}.zip")
    ext = File.extname(template_path)

    Zip::File.open(archive_path, Zip::File::CREATE) do |zipfile|
      @issues.each do |issue|
        renderer = create_renderer(template_path, [issue], config)
        result = renderer.render

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

  def create_renderer(template_path, issues, config)
    ext = File.extname(template_path).downcase
    renderer_class = case ext
                     when '.docx' then DocumentGenerator::WordRenderer
                     when '.xlsx' then DocumentGenerator::ExcelRenderer
                     else
                       raise DocumentGenerator::TemplateError, I18n.t('document_generator.error_invalid_format')
                     end

    renderer_class.new(template_path, issues, config, @error_behavior)
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
end