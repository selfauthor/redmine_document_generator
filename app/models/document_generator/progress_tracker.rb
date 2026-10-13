# frozen_string_literal: true

require 'json'
require 'fileutils'

module DocumentGenerator
  # Хранит состояние генерации в отдельном файле и синхронизирует доступ через flock.
  class ProgressTracker
    # Инициализирует состояние операции.
    # @param path [String] Путь к JSON-файлу состояния.
    # @param project_id [Integer] ID проекта Redmine.
    # @param user_id [Integer] ID пользователя, запустившего генерацию.
    def initialize(path, project_id:, user_id:)
      @path = path
      FileUtils.mkdir_p(File.dirname(@path))
      @project_id = project_id
      @user_id = user_id
      @mutex = Mutex.new
      @last_write_at = 0.0
      @percent = 0
      @phase_min = 0
      @phase_max = 95
      @stage = 'preparing'
      @message = I18n.t('document_generator.progress_wait')
      @status = 'running'
      write_state(force: true)
    end

    # Переключает этап и задаёт диапазон процентов для его продвижения.
    # @param stage [String] Ключ текущего этапа.
    # @param min [Integer] Нижняя граница процентов.
    # @param max [Integer] Верхняя граница процентов.
    # @param message [String] Локализованное описание состояния.
    # @param percent [Integer, nil] Необязательное начальное значение.
    # @param force [Boolean] Разрешает немедленную запись состояния.
    def set_phase(stage:, min:, max:, message:, percent: nil, force: true)
      @mutex.synchronize do
        @stage = stage
        @phase_min = min.to_i
        @phase_max = max.to_i
        @message = message

        if percent
          @percent = [[percent.to_i, @phase_min].max, @phase_max].min
        else
          @percent = [[@percent, @phase_min].max, @phase_max].min
        end

        write_state(force: force)
      end
    end

    # Устанавливает процент по фактической доле обработанных записей.
    # @param fraction [Float] Доля завершения этапа от 0.0 до 1.0.
    def update_fraction(fraction)
      @mutex.synchronize do
        value = fraction.to_f.clamp(0.0, 1.0)
        calculated = @phase_min + ((@phase_max - @phase_min) * value).floor

        @percent = [[calculated, @percent].max, @phase_max].min
        write_state
      end
    end

    # Продвигает индикатор во время длительной операции.
    # Процент не пересекает границу этапа и не достигает 100 до завершения.
    def tick
      @mutex.synchronize do
        @percent = [@percent + 1, @phase_max].min if @status == 'running'
        write_state
      end
    end

    # Завершает операцию и сохраняет итоговое состояние.
    # @param status [String] Итоговый статус: success или error.
    # @param message [String, nil] Сообщение об ошибке.
    def finish(status:, message: nil)
      @mutex.synchronize do
        @status = status
        @message = message if message.present?
        @percent = 100 if status == 'success'
        write_state(force: true)
      end
    end

    # Читает состояние с разделяемой блокировкой.
    # Если писатель удерживает эксклюзивную блокировку, чтение ожидает.
    # @param path [String] Путь к JSON-файлу состояния.
    # @return [Hash] Состояние операции.
    def self.read(path)
      File.open(path, 'r') do |file|
        file.flock(File::LOCK_SH)
        JSON.parse(file.read)
      ensure
        file.flock(File::LOCK_UN) rescue nil
      end
    end

    private

    # Записывает полное состояние под эксклюзивной блокировкой.
    # Обычные записи ограничены интервалом, важные переходы выполняются немедленно.
    # @param force [Boolean] Если true, запись выполняется без ожидания интервала.
    def write_state(force: false)
      now = Process.clock_gettime(Process::CLOCK_MONOTONIC)
      return if !force && now - @last_write_at < 0.8

      state = {
        uuid: File.basename(@path, '.json').sub(/\A\.progress-/, ''),
        project_id: @project_id,
        user_id: @user_id,
        status: @status,
        percent: @percent,
        stage: @stage,
        message: @message,
        updated_at: Time.current.iso8601
      }

      # Файл открывается без усечения. Содержимое изменяется только после получения блокировки.
      File.open(@path, File::RDWR | File::CREAT, 0o600) do |file|
        file.flock(File::LOCK_EX)
        file.rewind
        file.truncate(0)
        file.write(JSON.generate(state))
        file.flush
        file.flock(File::LOCK_UN)
      end

      @last_write_at = now
    end
  end
end
# v2610131526
