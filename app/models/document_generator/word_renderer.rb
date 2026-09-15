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
              
              # ВАЖНО: применяем склейку и подстановку к каждому клону
              join_and_substitute_block(clone, merged_context, ns)
              
              parent.add_child(clone)
            end
          end
        else
          error_msg = I18n.t('document_generator.error_missing_end_row')
          raise DocumentGenerator::TemplateError, error_msg
        end
      end

      # 2. Глобальная подстановка для ВСЕГО документа
      # Используем тот же надежный метод склейки для абзацев и строк таблиц,
      # чтобы гарантированно найти и заменить маркеры вроде <%ProjectName%>, 
      # даже если Word разбил их на части.
      render_context = context.merge(context['records'].first || {})
      
      doc.xpath('//w:p | //w:tr', ns).each do |block_node|
        # Пропускаем блоки, которые уже были обработаны как часть цикла строк
        next if block_node.name == 'tr' && @parser_config[:blocks][:row] && 
                block_node.xpath('.//w:t', ns).map(&:text).join.include?('BEGIN_ROW')

        join_and_substitute_block(block_node, render_context, ns)
      end
    end

    # Склеивает все текстовые узлы (<w:t>) внутри блока (<w:p> или <w:tr>) в одну строку,
    # выполняет подстановку маркеров, а затем помещает результат в первый текстовый узел,
    # очищая остальные. Это решает проблему разбитых маркеров Word-ом.
    #
    # @param block_node [Nokogiri::XML::Node] Узел абзаца или строки таблицы
    # @param context [Hash] Данные для подстановки
    # @param ns [Hash] Пространство имен XML
    def join_and_substitute_block(block_node, context, ns)
      text_nodes = block_node.xpath('.//w:t', ns)
      return if text_nodes.empty?

      # 1. Собираем весь текст блока в одну строку
      original_text = text_nodes.map(&:text).join

      # 2. Очищаем от управляющих маркеров, которые не должны попадать в итоговый текст
      cleaned_text = original_text.gsub(/<%\s*(BEGIN_ROW|END_ROW|BEGIN_SUBTASKS|END_SUBTASKS|BEGIN_WATCHERS|END_WATCHERS|BEGIN_RELATIONS|END_RELATIONS|BEGIN_GROUP_HEADER|END_GROUP_HEADER|BEGIN_GROUP_HEADER_2|END_GROUP_HEADER_2|BEGIN_GROUP_FOOTER|END_GROUP_FOOTER|BEGIN_GROUP_FOOTER_2|END_GROUP_FOOTER_2|BEGIN_TOTAL|END_TOTAL|GROUP_BY|GROUP_BY_2)\s*%>/i, '')

      # 3. Выполняем подстановку данных
      substituted_text = TemplateProcessor.substitute_markers(cleaned_text, context)

      # 4. Если текст изменился, перезаписываем структуру блока
      if original_text != substituted_text
        runs = block_node.xpath('.//w:r', ns)
        if runs.any?
          first_run = runs.first
          
          # Находим или создаем узел w:t в первом прогоне (w:r)
          text_node = first_run.at_xpath('.//w:t', ns)
          unless text_node
            text_node = Nokogiri::XML::Node.new('w:t', block_node.document)
            first_run.add_child(text_node)
          end
          
          # Сохраняем атрибут пробелов, чтобы Word не схлопывал их
          text_node['xml:space'] = 'preserve'
          text_node.content = substituted_text

          # Очищаем все остальные текстовые узлы в этом блоке, чтобы избежать дублирования текста
          text_nodes.each do |t_node|
            t_node.content = '' unless t_node == text_node
          end
        end
      end
    end

    def clean_node_text(node, ns)
      node.xpath('.//w:t', ns).each do |text_node|
        text = text_node.text
        cleaned = text.gsub(/<%\s*(BEGIN_ROW|END_ROW|BEGIN_SUBTASKS|END_SUBTASKS|BEGIN_WATCHERS|END_WATCHERS|BEGIN_RELATIONS|END_RELATIONS|BEGIN_GROUP_HEADER|END_GROUP_HEADER|BEGIN_GROUP_HEADER_2|END_GROUP_HEADER_2|BEGIN_GROUP_FOOTER|END_GROUP_FOOTER|BEGIN_GROUP_FOOTER_2|END_GROUP_FOOTER_2|BEGIN_TOTAL|END_TOTAL|GROUP_BY|GROUP_BY_2)\s*%>/i, '')
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
  # v2609150946
end