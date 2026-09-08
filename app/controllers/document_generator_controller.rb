# frozen_string_literal: true

class DocumentGeneratorController < ApplicationController
  before_action :find_project
  before_action :authorize_document_generator
  before_action :check_gems_loaded, only: [:export]

  # GET /projects/:project_id/document_generator/dialog
  def dialog
    # Восстанавливаем запрос из параметров (включая условия текущей фильтрации)
    @query = IssueQuery.new(name: '_temp', project: @project)
    @query.build_from_params(params)
    
    # Подсчитываем количество записей с учётом прав текущего пользователя
    @record_count = @query.issue_count
    
    respond_to do |format|
      format.js
    end
  end

  # POST /projects/:project_id/document_generator/export
  def export
    # Приём параметров (Этап 2)
    @template_file = params[:template_file]
    @export_mode = params[:export_mode]
    @file_name = params[:file_name]
    @error_behavior = params[:error_behavior]
    
    # Валидация имени файла на стороне сервера
    if @file_name =~ /[\/\\:*?"<>|]/
      flash[:error] = l('document_generator.error_invalid_filename')
      redirect_back(fallback_location: project_issues_path(@project)) and return
    end

    # Временно возвращаем информацию о принятых параметрах для отладки
    # На Этапе 3 здесь будет реализована реальная генерация документов
    render plain: "Параметры успешно получены:\n" \
                 "Режим выгрузки: #{@export_mode}\n" \
                 "Имя файла: #{@file_name}\n" \
                 "Поведение при ошибках: #{@error_behavior}\n" \
                 "Файл шаблона: #{@template_file&.original_filename}\n" \
                 "Количество записей: #{@query&.issue_count || 'N/A'}",
           content_type: 'text/plain'
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
    return if DOCUMENT_GENERATOR_GEMS_LOADED

    flash[:error] = l(:error_gems_not_loaded)
    redirect_to project_issues_path(@project)
  end
end