# frozen_string_literal: true

require 'sablon'
require 'zip'
require 'fileutils'
require 'nokogiri'

# v260914858

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
      inject_html_descriptions(context)
      processed_template_path = preprocess_template
      output_path = "#{processed_template_path}.output.docx"
      begin
        template = Sablon.template(processed_template_path)
        template.render_to_file(output_path, context)
        output_path
      rescue StandardError => e
        error_msg = I18n.t('document_generator.error_word_render_failed', message: e.message)
        handle_error(error_msg)
      ensure
        FileUtils.rm_f(processed_template_path) if processed_template_path && File.exist?(processed_template_path)
      end
    end

    private

    def inject_html_descriptions(obj)
      case obj
      when Hash
        obj.each do |key, value|
          if key.to_s.downcase == 'description' && value.is_a?(String)
            obj[key] = value.present? ? Sablon::Content::Html.new(value) : ''
          else
            inject_html_descriptions(value)
          end
        end
      when Array
        obj.each { |item| inject_html_descriptions(item) }
      end
    end

    def preprocess_template
      temp_path = "#{@template_path}.preprocessed.docx"
      temp_new_path = "#{temp_path}.new"
      FileUtils.cp(@template_path, temp_path)

      xml_entries_to_modify = []
      Zip::File.open(temp_path) do |zip|
        xml_entries_to_modify = (zip.glob('word/document.xml') +
                                 zip.glob('word/header*.xml') +
                                 zip.glob('word/footer*.xml')).map(&:name)
      end

      Zip::File.open(temp_path) do |input_zip|
        Zip::OutputStream.open(temp_new_path) do |output_stream|
          input_zip.each do |entry|
            output_stream.put_next_entry(entry.name, nil, nil, entry.compression_method)
            content = entry.get_input_stream.read
            if xml_entries_to_modify.include?(entry.name)
              content = process_xml_content(content)
            end
            output_stream.write(content)
          end
        end
      end

      FileUtils.mv(temp_new_path, temp_path)
      temp_path
    end

    # ============================================================
    # ИСПРАВЛЕНИЕ 1: Склеивание разбитых XML-узлов
    # ============================================================
    # Вместо обработки каждого <w:t> отдельно, мы:
    # 1. Для каждого <w:p> склеиваем текст из всех <w:r>/<w:t>
    # 2. Ищем маркеры в склеенном тексте
    # 3. Перестраиваем <w:r> элементы, помещая маркеры в отдельные runs
    # ============================================================
    def process_xml_content(xml_content)
      doc = Nokogiri::XML(xml_content)
      ns = { 'w' => 'http://schemas.openxmlformats.org/wordprocessingml/2006/main' }

      paragraphs = doc.xpath('//w:p', ns)
      paragraphs.each do |para|
        merge_and_process_paragraph(para, ns)
      end

      # ИСПРАВЛЕНИЕ 2: Автоматическое оборачивание в «records»
      auto_wrap_in_records(doc, ns)

      doc.to_xml
    end

    def merge_and_process_paragraph(para, ns)
      runs = para.xpath('.//w:r', ns).to_a
      return if runs.empty?

      # Склеиваем текст из всех runs в параграфе
      full_text = runs.map { |r| r.at_xpath('w:t', ns)&.text || '' }.join
      return unless full_text.include?('<%')

      # Находим все маркеры в полном (склеенном) тексте
      parts = []
      last_pos = 0
      full_text.scan(/<%\s*([^%]+?)\s*%>/) do |match|
        marker_start = $~.begin(0)
        marker_end = $~.end(0)
        marker_content = match[0].strip

        parts << { type: :text, content: full_text[last_pos...marker_start] } if marker_start > last_pos
        parts << { type: :marker, content: transform_marker(marker_content) }
        last_pos = marker_end
      end
      parts << { type: :text, content: full_text[last_pos..-1] } if last_pos < full_text.length

      return if parts.empty?
      return if parts.all? { |p| p[:type] == :text }

      # Берём форматирование из первого run с непустым rPr
      rpr_xml = ''
      runs.each do |r|
        rpr = r.at_xpath('w:rPr', ns)
        if rpr
          rpr_xml = rpr.to_xml
          break
        end
      end

      # Создаём новые runs
      new_runs_xml = parts.reject { |p| p[:content].empty? }.map do |part|
        content = escape_xml(part[:content])
        "<w:r>#{rpr_xml}<w:t xml:space=\"preserve\">#{content}</w:t></w:r>"
      end.join

      # Вставляем новые runs перед первым старым run и удаляем старые
      runs.first.before(new_runs_xml)
      runs.each(&:remove)
    end

    # ============================================================
    # ИСПРАВЛЕНИЕ 2: Если в шаблоне нет «records», но есть маркеры
    # полей — автоматически оборачиваем их в блок итерации
    # ============================================================
    def auto_wrap_in_records(doc, ns)
      xml_str = doc.to_xml
      return if xml_str.include?('«records»')

      # Метаданные и служебные маркеры, которые НЕ являются полями записей
      meta_markers = %w[
        end records groups totals else
        ExportDate ExportUser ProjectName QueryName FilterDescription
        total_count row_number row_number_in_group row_number_in_group_2
        GroupValue GroupValue2 count
      ]

      # Ищем параграфы, содержащие маркеры полей (не метаданные)
      paragraphs = doc.xpath('//w:p', ns).to_a
      first_field_para = nil
      last_field_para = nil

      paragraphs.each do |para|
        text = para.xpath('.//w:t', ns).map(&:text).join
        # Проверяем, есть ли в параграфе маркер, который не является метаданным
        has_field = text.scan(/«([^»]+)»/).flatten.any? do |m|
          next false if meta_markers.include?(m)
          next false if m.start_with?('if ')
          next false if m.start_with?('agg_')
          true
        end

        if has_field
          first_field_para ||= para
          last_field_para = para
        end
      end

      return unless first_field_para

      # Вставляем «records» перед первым параграфом с полем
      records_para = '<w:p><w:r><w:t xml:space="preserve">«records»</w:t></w:r></w:p>'
      first_field_para.before(records_para)

      # Вставляем «end» после последнего параграфа с полем
      end_para = '<w:p><w:r><w:t xml:space="preserve">«end»</w:t></w:r></w:p>'
      last_field_para.after(end_para)
    end

    def escape_xml(text)
      text.gsub('&', '&amp;')
          .gsub('<', '&lt;')
          .gsub('>', '&gt;')
          .gsub('"', '&quot;')
          .gsub("'", '&apos;')
    end

    def transform_marker(inner)
      # Блоки итерации и условий — конвертируем в синтаксис Sablon
      case inner
      when /^BEGIN_ROW$/i
        '{% records %}'
      when /^END_ROW$/i
        '{% end %}'
      when /^BEGIN_SUBTASKS$/i
        '{% subtasks %}'
      when /^END_SUBTASKS$/i
        '{% end %}'
      when /^BEGIN_WATCHERS$/i
        '{% watchers %}'
      when /^END_WATCHERS$/i
        '{% end %}'
      when /^BEGIN_RELATIONS$/i
        '{% relations %}'
      when /^END_RELATIONS$/i
        '{% end %}'
      when /^BEGIN_GROUP_HEADER$/i
        '{% groups %}'
      when /^END_GROUP_HEADER$/i
        '{% end %}'
      when /^BEGIN_GROUP_FOOTER_2$/i, /^BEGIN_GROUP_FOOTER$/i
        '' # Эти блоки обрабатываются на уровне ContextBuilder, не нужны в шаблоне
      when /^END_GROUP_FOOTER_2$/i, /^END_GROUP_FOOTER$/i
        '{% end %}'
      when /^BEGIN_TOTAL$/i
        '{% totals %}'
      when /^END_TOTAL$/i
        '{% end %}'
      when /^IF\((.+)\)$/i
        condition = $1.strip
        # Для простоты, преобразуем IF(field) в {% if field %}
        # Более сложные условия (==, !=) потребуют парсинга, но для начала этого хватит
        '{% if ' + condition + ' %}'
      when /^ELSE$/i
        '{% else %}'
      when /^END$/i
        '{% end %}'
      else
        # Это поле данных. Оставляем его как есть, без кавычек!
        # Именно так Sablon и ожидает видеть имя переменной.
        inner
      end
    end

    def handle_error(message)
      case @error_behavior
      when 'abort'
        raise DocumentGenerator::RenderError, message
      when 'skip_field', 'skip_record'
        Rails.logger.error "[DocumentGenerator] Critical rendering error (behavior: #{@error_behavior}): #{message}"
        raise DocumentGenerator::RenderError, message
      end
    end
  end
end