# frozen_string_literal: true

require 'fileutils'

# v2609141507
module DocumentGenerator
  # WordRenderer is responsible for generating Word documents (.docx).
  # It uses TemplateProcessor for archive extraction and saving,
  # but manages row cloning and text replacement in <w:t> nodes specifically for Word XML.
  class WordRenderer
    # @param template_path [String] Path to the temporary template file
    # @param issues [ActiveRecord::Relation] The collection of records to export
    # @param parser_config [Hash] Configuration from TemplateParser
    # @param error_behavior [String] Error handling strategy ('abort', 'skip_field', 'skip_record')
    def initialize(template_path, issues, parser_config, error_behavior)
      @template_path = template_path
      @issues = issues
      @parser_config = parser_config
      @error_behavior = error_behavior
    end

    # Main document generation method.
    #
    # @return [String] Path to the generated .docx file
    def render
      context = ContextBuilder.new(@issues, @parser_config, @error_behavior).build
      output_path = "#{@template_path}.output.docx"

      # XML targets inside .docx that may contain data
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

    # Processes a specific Word XML file.
    # Finds row blocks (BEGIN_ROW/END_ROW) for cloning,
    # then replaces markers in all text nodes.
    #
    # @param doc [Nokogiri::XML::Document] The Word XML document
    # @param context [Hash] Data for substitution
    # @param entry_name [String] Name of the file being processed inside the archive
    def process_word_xml(doc, context, entry_name)
      ns = { 'w' => 'http://schemas.openxmlformats.org/wordprocessingml/2006/main' }

      # ШАГ 1: Принудительно склеиваем разбитые Word'ом текстовые узлы внутри абзаца,
      # если в этом абзаце обнаружены маркеры <% ... %>
      doc.xpath('//w:p', ns).each do |para|
        text_nodes = para.xpath('.//w:t', ns).to_a
        next if text_nodes.empty?

        full_text = text_nodes.map(&:text).join

        if full_text.include?('<%') && full_text.include?('%>')
          # Записываем полный, склеенный текст в самый первый узел <w:t>
          first_node = text_nodes.first
          first_node.content = full_text

          # Удаляем все последующие узлы <w:t> и их родительские <w:r>,
          # чтобы избежать дублирования текста в итоговом документе
          text_nodes[1..-1].each { |node| node.parent.remove }
        end
      end

      # ШАГ 2: Дальнейшая обработка (теперь маркеры гарантированно целые)
      if @parser_config[:blocks][:row] && context['records'].present?
        nodes_to_process = doc.xpath('//w:p | //w:tr', ns)
        
        nodes_to_process.each do |node|
          text = node.xpath('.//w:t', ns).map(&:text).join
          next unless text.include?('<%BEGIN_ROW%>')

          clean_node_text(node, ns)

          context['records'].each_with_index do |record, _index|
            clone = node.dup
            
            clone.xpath('.//w:t', ns).each do |text_node|
              original_text = text_node.text
              full_context = context.merge(record)
              text_node.content = TemplateProcessor.substitute_markers(original_text, full_context)
            end
            
            node.add_next_sibling(clone)
          end

          node.remove
        end
      else
        first_record = context['records'].first || {}
        render_context = context.merge(first_record)
        
        doc.xpath('//w:t', ns).each do |text_node|
          original_text = text_node.text
          text_node.content = TemplateProcessor.substitute_markers(original_text, render_context)
        end
      end
    end

    # Removes control markers (e.g., <%BEGIN_ROW%>) from text nodes,
    # leaving only the visible text.
    #
    # @param node [Nokogiri::XML::Node] The <w:p> or <w:tr> node
    # @param ns [Hash] XML namespace
    def clean_node_text(node, ns)
      node.xpath('.//w:t', ns).each do |text_node|
        text = text_node.text
        cleaned = text.gsub(/<%\s*(BEGIN_ROW|END_ROW|BEGIN_SUBTASKS|END_SUBTASKS|BEGIN_WATCHERS|END_WATCHERS|BEGIN_RELATIONS|END_RELATIONS|BEGIN_GROUP_HEADER|END_GROUP_HEADER|BEGIN_TOTAL|END_TOTAL)\s*%>/i, '')
        text_node.content = cleaned
      end
    end

    # Universal error handler.
    #
    # @param message [String] The error message (already localized)
    def handle_error(message)
      case @error_behavior
      when 'abort'
        raise DocumentGenerator::RenderError, message
      when 'skip_field', 'skip_record'
        Rails.logger.error "[DocumentGenerator] Render error (behavior: #{@error_behavior}): #{message}"
        raise DocumentGenerator::RenderError, message
      end
    end
  end
end