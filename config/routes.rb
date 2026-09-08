# frozen_string_literal: true

RedmineApp::Application.routes.draw do
  scope 'projects/:project_id/document_generator' do
    get  'dialog', to: 'document_generator#dialog', as: 'document_generator_dialog'
    post 'export', to: 'document_generator#export', as: 'document_generator_export'
  end
end