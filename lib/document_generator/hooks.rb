# frozen_string_literal: true

module DocumentGenerator
  class Hooks < Redmine::Hook::ViewListener

    # Создаёт снимок полного списка задач и передаёт его UUID в ссылку генератора.
    #
    # @param context [Hash] Контекст хука Redmine со списком задач, проектом,
    #   объектом запроса и средством рендеринга представлений.
    # @return [String] HTML ссылки генератора либо пустая строка при ошибке.
    def view_issues_index_bottom(context = {})
      # Получаем объекты, переданные Redmine в контексте хука.
      hook_caller = context[:hook_caller]
      project = context[:project]
      query = context[:query]

      # Без средства рендеринга, проекта или корректного запроса продолжать нельзя.
      return '' unless hook_caller && project && query&.valid?

      # Не создаём снимки для пользователей без доступа к генератору.
      return '' unless project.module_enabled?(:document_generator)
      return '' unless User.current.allowed_to?(:use_document_generator, project)

      # Получаем полный список ID в порядке, рассчитанном IssueQuery.
      # В отличие от context[:issues], результат не ограничен текущей страницей.
      issue_ids = query.issue_ids

      # Создаём снимок для текущего проекта и текущего пользователя.
      issue_list_token = DocumentGenerator::IssueListSnapshot.create(
        issue_ids: issue_ids,
        project_id: project.id,
        user_id: User.current.id
      )

      # Передаём UUID в partial, который формирует ссылку открытия диалога.
      hook_caller.render(
        partial: 'document_generator/export_link',
        locals: { issue_list_token: issue_list_token }
      )
    rescue StandardError => e
      # Ошибка создания снимка не должна мешать отображению списка задач Redmine.
      Rails.logger.error(
        "[DocumentGenerator] Failed to create issue-list snapshot: " \
        "#{e.class}: #{e.message}"
      )

      ''
    end

  end
end
# v2610141222