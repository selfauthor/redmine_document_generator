# frozen_string_literal: true

module DocumentGenerator
  class WordRenderer
    def initialize(template_path, issues, parser_config, error_behavior)
      @template_path = template_path
      @issues = issues
      @parser_config = parser_config
      @error_behavior = error_behavior
    end

    def render
      context = ContextBuilder.new(@issues, @parser_config, @error_behavior).build
      output_path = "#{@template_path}.output.docx"
      xml_targets = [
        'word/document.xml',
        'word/header*.xml',
        'word/footer*.xml'
      ]
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

    def process_word_xml(doc, context, entry_name)
      ns = { 'w' => 'http://schemas.openxmlformats.org/wordprocessingml/2006/main' }
      records = context['records'] || []

      # 1. Обработка циклов (клонирование строк)
      if @parser_config[:blocks][:row] && records.present?
        all_nodes = doc.xpath('//w:tr | //w:p[not(ancestor::w:tr)]', ns)
        start_node = nil
        end_node = nil
        template_nodes = []
        in_block = false

        all_nodes.each do |node|
          text = node.xpath('.//w:t', ns).map(&:text).join
          if !in_block && text.include?('<%BEGIN_ROW%>')
            in_block = true
            start_node = node
            template_nodes << node
            if text.include?('<%END_ROW%>')
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

        if start_node && end_node
          if start_node.parent != end_node.parent
            error_msg = I18n.t('document_generator.error_row_block_mismatch')
            raise DocumentGenerator::TemplateError, error_msg
          end
          parent = start_node.parent
          template_nodes.each(&:remove)
          template_nodes.each do |node|
            clean_node_text(node, ns)
          end
          records.each do |record|
            merged_context = context.merge(record)
            template_nodes.each do |template_node|
              clone = template_node.dup
              # СПЕЦИФИКА WORD: обработка XML-узлов с сохранением форматирования
              process_word_block(clone, merged_context, ns)
              parent.add_child(clone)
            end
          end
        else
          error_msg = I18n.t('document_generator.error_missing_end_row')
          raise DocumentGenerator::TemplateError, error_msg
        end
      end

      # 2. Глобальная подстановка для ВСЕГО документа
      render_context = context.merge(context['records'].first || {})
      doc.xpath('//w:p | //w:tr', ns).each do |block_node|
        next if block_node.name == 'tr' && @parser_config[:blocks][:row] && 
                block_node.xpath('.//w:t', ns).map(&:text).join.include?('BEGIN_ROW')
        # СПЕЦИФИКА WORD: обработка XML-узлов
        process_word_block(block_node, render_context, ns)
      end
    end

    # Обрабатывает блок Word (абзац или строку таблицы):
    # 1. Сначала применяет условия (общая логика из TemplateProcessor)
    # 2. Затем подставляет значения с сохранением форматирования
    #
    # @param block_node [Nokogiri::XML::Node] Узел w:p или w:tr
    # @param context [Hash] Данные для подстановки
    # @param ns [Hash] Пространства имен XML
    def process_word_block(block_node, context, ns)
      runs = block_node.xpath('.//w:r', ns)
      return if runs.empty?

      # 1. Сначала обрабатываем условия (ОБЩАЯ ЛОГИКА)
      # Собираем весь текст, обрабатываем условия, затем разбираем обратно по узлам
      full_text = runs.map { |r| r.xpath('.//w:t', ns).map(&:text).join }.join
      processed_text = TemplateProcessor.resolve_conditionals(full_text, context)
      
      # Применяем обработанный текст обратно к узлам
      if full_text != processed_text
        apply_text_to_runs(runs, processed_text, ns)
      end

      # 2. Подставляем значения (ОБЩАЯ ЛОГИКА + СПЕЦИФИКА WORD)
      join_and_substitute_block(block_node, context, ns)
    end

    # Применяет текст обратно к XML-узлам w:r, сохраняя их структуру
    def apply_text_to_runs(runs, new_text, ns)
      current_pos = 0
      runs.each do |run|
        t_nodes = run.xpath('.//w:t', ns)
        next if t_nodes.empty?

        run_length = t_nodes.map(&:text).join.length
        if current_pos < new_text.length
          chunk = new_text[current_pos, run_length] || ""
          t_nodes.first.content = chunk
          t_nodes[1..-1].each { |t| t.content = "" } if t_nodes.size > 1
        end
        current_pos += run_length
      end
    end

    # СПЕЦИФИКА WORD: склеивает текстовые узлы внутри блока и подставляет значения,
    # сохраняя форматирование первого w:r (шрифт, цвет, жирность и т.д.)
    #
    # @param block_node [Nokogiri::XML::Node] Узел w:p или w:tr
    # @param context [Hash] Данные для подстановки
    # @param ns [Hash] Пространства имен XML
    def join_and_substitute_block(block_node, context, ns)
      text_nodes = block_node.xpath('.//w:t', ns)
      return if text_nodes.empty?

      # 1. Собираем весь текст блока
      original_text = text_nodes.map(&:text).join
      
      # 2. Очищаем от управляющих маркеров (ОБЩАЯ ЛОГИКА)
      cleaned_text = TemplateProcessor.clean_control_markers(original_text)
      
      # 3. Подставляем значения (ОБЩАЯ ЛОГИКА)
      substituted_text = TemplateProcessor.substitute_markers(cleaned_text, context)

      # 4. Если текст изменился, перезаписываем с сохранением форматирования
      if original_text != substituted_text
        runs = block_node.xpath('.//w:r', ns)
        if runs.any?
          first_run = runs.first
          text_node = first_run.at_xpath('.//w:t', ns)
          unless text_node
            text_node = Nokogiri::XML::Node.new('w:t', block_node.document)
            text_node['xml:space'] = 'preserve'
            first_run.add_child(text_node)
          end
          text_node.content = substituted_text
          # Очищаем остальные узлы, чтобы не было дублирования
          text_nodes.each { |t| t.content = '' unless t == text_node }
        end
      end
    end

    # СПЕЦИФИКА WORD: очищает управляющие маркеры из узлов
    def clean_node_text(node, ns)
      node.xpath('.//w:t', ns).each do |text_node|
        text = text_node.text
        cleaned = TemplateProcessor.clean_control_markers(text)
        text_node.content = cleaned
      end
    end

    def handle_error(message)
      case @error_behavior
      when 'skip_field'
        Rails.logger.warn "[DocumentGenerator] Render warning (skip_field mode): #{message}"
        raise DocumentGenerator::TemplateError, message
      when 'skip_record'
        Rails.logger.warn "[DocumentGenerator] Render warning (skip_record mode): #{message}"
        raise DocumentGenerator::TemplateError, message
      else
        raise DocumentGenerator::TemplateError, message
      end
    end
  end
end
# v2609151130