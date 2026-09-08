# frozen_string_literal: true

module DocumentGenerator
  class Hooks < Redmine::Hook::ViewListener
    # Хук срабатывает внизу страницы списка задач.
    # Рендерит partial со ссылкой "Генератор документов".
    render_on :view_issues_index_bottom, partial: 'document_generator/export_link'
  end
end