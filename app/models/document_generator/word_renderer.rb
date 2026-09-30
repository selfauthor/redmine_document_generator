# frozen_string_literal: true

module DocumentGenerator
  # ===========================================================================
  # РЕНДЕРИНГ WORD-ДОКУМЕНТОВ
  # ===========================================================================
  # Класс отвечает за генерацию Word-документов (.docx) из шаблонов.
  # Обрабатывает циклы, условия, подстановку полей с сохранением форматирования.
  class WordRenderer
    # Возвращает предупреждения, накопленные при формировании документа.
    # @return [Array<String>] Список предупреждений для пользователя.
    attr_reader :warnings

    # Инициализация рендерера
    #
    # @param template_path [String] Путь к файлу шаблона
    # @param issues [Array<Issue>] Массив задач для выгрузки
    # @param parser_config [Hash] Конфигурация парсера (настройки блоков)
    # @param error_behavior [String] Поведение при ошибках: 'abort', 'skip_field', 'skip_record'
    def initialize(template_path, issues, parser_config, error_behavior)
      @template_path = template_path
      @issues = issues
      @parser_config = parser_config
      @error_behavior = error_behavior

      # Создаём массив для предупреждений текущей выгрузки.
      @warnings = []
    end

    # Генерация документа
    #
    # @return [String] Путь к сгенерированному файлу
    # @raise [RenderError] если произошла ошибка при рендеринге
    def render
      # Строим контекст данных
      context = ContextBuilder.new(@issues, @parser_config, @error_behavior).build

      # Используем тот же массив предупреждений, который создал ContextBuilder.
      # Благодаря общей ссылке предупреждения из обработки полей будут доступны renderer.
      @warnings = context['__warnings'] || @warnings
      # Путь для выходного файла
      output_path = "#{@template_path}.output.docx"
      # Файлы для обработки
      xml_targets = [
        'word/document.xml',
        'word/header*.xml',
        'word/footer*.xml'
      ]
      # Обрабатываем архив
      TemplateProcessor.process_archive(@template_path, output_path, xml_targets) do |doc, entry_name|
        process_word_xml(doc, context, entry_name)
      end
      output_path
    rescue StandardError => e
      Rails.logger.error "[DocumentGenerator] Word render failed: #{e.message}\n#{e.backtrace&.join("\n")}"
      error_msg = I18n.t('document_generator.error_word_render_failed', message: e.message)
      handle_error(error_msg)
    end

    private

    # Обработка XML документа Word
    #
    # @param doc [Nokogiri::XML::Document] XML-документ
    # @param context [Hash] Контекст данных
    # @param entry_name [String] Имя файла в архиве
    def process_word_xml(doc, context, entry_name)
      ns = { 'w' => 'http://schemas.openxmlformats.org/wordprocessingml/2006/main' }
      records = context['records'] || []
      # Если есть блок BEGIN_ROW/END_ROW и есть записи - обрабатываем циклы
      if @parser_config[:blocks][:row] && records.present?
        process_row_blocks(doc, context, records, ns)
      end
      # Контекст для рендеринга (общий + первая запись)
      render_context = context.merge(context['records'].first || {})
      # Обрабатываем все абзацы и строки таблиц
      doc.xpath('//w:p | //w:tr', ns).each do |block_node|
        next if block_node.name == 'tr' && @parser_config[:blocks][:row] &&
                block_node.xpath('.//w:t', ns).map(&:text).join.include?('BEGIN_ROW')
        
        # 1. Обрабатываем коллекции (подзадачи, наблюдатели, связи)
        TemplateProcessor.process_collection_blocks(block_node, render_context, ns, @error_behavior)
        # 2. Обрабатываем условия
        TemplateProcessor.process_conditionals_in_block(block_node, render_context, ns, @error_behavior)
        # 3. Подставляем значения
        TemplateProcessor.substitute_in_block(block_node, render_context, ns, @error_behavior)
      end
    end

    # Обработка блоков с циклами (BEGIN_ROW/END_ROW)
    #
    # @param doc [Nokogiri::XML::Document] XML-документ
    # @param context [Hash] Общий контекст данных
    # @param records [Array<Hash>] Массив записей для цикла
    # @param ns [Hash] Пространства имен XML
    def process_row_blocks(doc, context, records, ns)
      # Находим все узлы (строки таблиц или абзацы)
      all_nodes = doc.xpath('//w:tr | //w:p[not(ancestor::w:tr)]', ns)
      start_node = nil
      end_node = nil
      template_nodes = []
      in_block = false
      
      # Ищем блок BEGIN_ROW/END_ROW
      all_nodes.each do |node|
        text = node.xpath('.//w:t', ns).map(&:text).join
        if !in_block && text.include?('<%BEGIN_ROW%>')
          # Начало блока
          in_block = true
          start_node = node
          template_nodes << node
          if text.include?('<%END_ROW%>')
            # BEGIN_ROW и END_ROW в одном узле
            end_node = node
            in_block = false
            break
          end
        elsif in_block
          template_nodes << node
          if text.include?('<%END_ROW%>')
            end_node = node
            in_block = false
            break
          end
        end
      end
      
      # Проверяем, что блок найден корректно
      if start_node && end_node
        if start_node.parent != end_node.parent
          error_msg = I18n.t('document_generator.error_row_block_mismatch')
          raise DocumentGenerator::TemplateError, error_msg
        end
        
        parent = start_node.parent
        
        # Удаляем шаблонные узлы
        template_nodes.each(&:remove)
        
        # НЕ вызываем clean_node_text здесь! Маркеры ELSE и END нужны для process_conditionals_in_block
        
        # Для каждой записи создаем клон
        records.each_with_index do |record, record_idx|
          begin
            merged_context = context.merge(record)

            # Клонируем исходные XML-узлы для текущей записи.
            clones = template_nodes.map(&:dup)

            # Разворачиваем вложенные коллекции в контексте этой задачи.
            clones = TemplateProcessor.process_collection_blocks(
              clones, merged_context, ns, @error_behavior
            )

            # Обрабатываем условия и поля в каждом клоне.
            clones.each do |clone|
              TemplateProcessor.process_conditionals_in_block(
                clone, merged_context, ns, @error_behavior
              )

              TemplateProcessor.substitute_in_block(
                clone, merged_context, ns, @error_behavior
              )

              parent.add_child(clone)
            end
          rescue DocumentGenerator::SkipRecordError => e
            # Записываем информацию о пропущенной записи на английском языке.
            issue = record['__issue']

            Rails.logger.warn(
              "[DocumentGenerator] Issue ##{issue&.id || 'unknown'} was skipped because a required template field was missing."
            )

            # Переходим к следующей задаче, не прерывая формирование документа.
            next
          end
        end
      else
        error_msg = I18n.t('document_generator.error_missing_end_row')
        raise DocumentGenerator::TemplateError, error_msg
      end
    end

    # Обработка ошибок
    #
    # @param message [String] Сообщение об ошибке
    # @raise [RenderError]
    def handle_error(message)
      case @error_behavior
      when 'skip_field'
        Rails.logger.warn "[DocumentGenerator] Render warning (skip_field mode): #{message}"
        raise DocumentGenerator::RenderError, message
      when 'skip_record'
        Rails.logger.warn "[DocumentGenerator] Render warning (skip_record mode): #{message}"
        raise DocumentGenerator::RenderError, message
      else
        raise DocumentGenerator::RenderError, message
      end
    end
  end
end
# v2609301230