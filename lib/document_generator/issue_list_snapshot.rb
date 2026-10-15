# frozen_string_literal: true

require 'json'
require 'fileutils'
require 'securerandom'

module DocumentGenerator
  # Хранит снимки упорядоченных списков задач для последующей выгрузки документов.
  class IssueListSnapshot
    # Возвращает каталог, в котором хранятся JSON-снимки списков задач.
    #
    # @return [Pathname] Путь к каталогу снимков.
    def self.storage_dir
      Rails.root.join('tmp', 'document_generator', 'issue_lists')
    end

    # Создаёт JSON-снимок списка задач и возвращает его UUID.
    #
    # @param issue_ids [Array<Integer>] Упорядоченные идентификаторы задач.
    # @param project_id [Integer] Идентификатор проекта.
    # @param user_id [Integer] Идентификатор пользователя, создавшего снимок.
    # @return [String] UUID созданного снимка.
    def self.create(issue_ids:, project_id:, user_id:)
      # Удаляем просроченные снимки до создания нового.
      cleanup_expired!

      # Создаём каталог снимков, если он ещё не существует.
      FileUtils.mkdir_p(storage_dir)

      # Генерируем уникальное имя, не раскрывающее данные проекта и пользователя.
      uuid = SecureRandom.uuid

      # Сохраняем только метаданные и упорядоченный список идентификаторов задач.
      data = {
        version: 1,
        project_id: project_id.to_i,
        user_id: user_id.to_i,
        created_at: Time.current.utc.iso8601,
        issue_ids: issue_ids.map(&:to_i)
      }

      # Записываем JSON в UTF-8.
      File.write(
        storage_dir.join("#{uuid}.json"),
        JSON.pretty_generate(data),
        mode: 'w:UTF-8'
      )

      uuid
    end

    # Читает снимок и проверяет его принадлежность проекту и пользователю.
    #
    # @param uuid [String] UUID снимка из запроса.
    # @param project_id [Integer] Идентификатор текущего проекта.
    # @param user_id [Integer] Идентификатор текущего пользователя.
    # @return [Hash, nil] Данные снимка либо nil, если он отсутствует или недействителен.
    def self.read(uuid:, project_id:, user_id:)
      # Принимаем только UUID ожидаемого формата, исключая произвольные пути.
      return nil unless uuid.to_s.match?(
        /\A[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}\z/i
      )

      # Формируем путь исключительно внутри каталога снимков.
      file_path = storage_dir.join("#{uuid}.json")

      # Не читаем отсутствующие файлы и символические ссылки.
      return nil unless File.file?(file_path) && !File.symlink?(file_path)

      # Загружаем и проверяем структуру JSON.
      data = JSON.parse(File.read(file_path, mode: 'r:UTF-8'))
      return nil unless data.is_a?(Hash)
      return nil unless data['version'].to_i == 1
      return nil unless data['project_id'].to_i == project_id.to_i
      return nil unless data['user_id'].to_i == user_id.to_i
      return nil unless data['issue_ids'].is_a?(Array)

      # Проверяем, что все сохранённые идентификаторы являются положительными числами.
      return nil unless data['issue_ids'].all? do |id|
        id.to_s.match?(/\A[1-9]\d*\z/)
      end

      data
    rescue JSON::ParserError, Errno::ENOENT, Errno::EACCES, TypeError
      # Повреждённый или недоступный снимок считается недействительным.
      nil
    end

    # Удаляет JSON-снимки, срок хранения которых истёк.
    # Возраст определяется исключительно по времени изменения файла.
    #
    # @return [void]
    def self.cleanup_expired!
      # Если каталог ещё не создан, очищать нечего.
      return unless Dir.exist?(storage_dir)

      # Используем существующую настройку срока хранения плагина.
      retention_hours = Setting.plugin_redmine_document_generator
                               .fetch('retention_hours', '24')
                               .to_i

      # При некорректном значении сохраняем срок хранения по умолчанию.
      retention_hours = 24 if retention_hours <= 0

      # Вычисляем границу срока хранения.
      expiration_time = Time.current - retention_hours.hours

      # Обрабатываем только JSON-файлы непосредственно в каталоге снимков.
      Dir.children(storage_dir).each do |entry|
        next unless entry.match?(
          /\A[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}\.json\z/i
        )

        file_path = storage_dir.join(entry)

        # Не удаляем каталоги и символические ссылки.
        next unless File.file?(file_path) && !File.symlink?(file_path)

        # В качестве единственного критерия возраста используем mtime файла.
        next unless File.mtime(file_path) < expiration_time

        # Удаляем только просроченный снимок.
        File.delete(file_path)
      rescue StandardError => e
        # Ошибка удаления одного файла не должна останавливать очистку остальных.
        Rails.logger.error(
          "[DocumentGenerator] Failed to remove expired issue-list snapshot " \
          "#{file_path}: #{e.message}"
        )
      end
    end
  end
end
# v2610141220