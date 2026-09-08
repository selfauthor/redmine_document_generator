# frozen_string_literal: true

class DocumentGeneratorController < ApplicationController
  before_action :find_project
  before_action :authorize_document_generator
  before_action :check_gems_loaded, only: [:export]

  # GET /projects/:project_id/document_generator/dialog
  def dialog
    respond_to do |format|
      format.html { render plain: 'Document Generator dialog will be available in Stage 2.' }
      format.js   { render plain: '// JS response will be implemented in Stage 2.' }
    end
  end

  # POST /projects/:project_id/document_generator/export
  def export
    render plain: 'Export functionality will be implemented in later stages.'
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

  # Проверка, что все необходимые gem-ы установлены
  def check_gems_loaded
    return if DOCUMENT_GENERATOR_GEMS_LOADED

    flash[:error] = l(:error_gems_not_loaded)
    redirect_to project_issues_path(@project)
  end
end