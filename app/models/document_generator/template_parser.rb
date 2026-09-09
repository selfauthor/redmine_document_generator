# frozen_string_literal: true

require 'zip'
require 'nokogiri'

module DocumentGenerator
  class TemplateParser
    attr_reader :file_path, :file_type

    def initialize(file_path)
      @file_path = file_path
      @file_type = detect_file_type
    end

    def detect_file_type
      ext = File.extname(@file_path).downcase
      case ext
      when '.docx' then :docx
      when '.xlsx' then :xlsx
      else raise "Unsupported file type: #{ext}"
      end
    end

    def parse
      text = extract_text
      config[:template_text] = text
      config = {
        group_by: nil,
        group_by_2: nil,
        blocks: {
          group_header: false, group_header_2: false, row: false,
          group_footer_2: false, group_footer: false, total: false
        },
        fields: [],
        has_subtasks: false,
        has_watchers: false,
        has_relations: false
      }

      # Поиск директив группировки
      config[:group_by] = text[/\<%\s*GROUP_BY\s*:\s*([^%]+)%\>/i, 1]&.strip
      config[:group_by_2] = text[/\<%\s*GROUP_BY_2\s*:\s*([^%]+)%\>/i, 1]&.strip

      # Поиск блоков
      config[:blocks][:group_header] = text.include?('<%BEGIN_GROUP_HEADER%>')
      config[:blocks][:group_header_2] = text.include?('<%BEGIN_GROUP_HEADER_2%>')
      config[:blocks][:row] = text.include?('<%BEGIN_ROW%>')
      config[:blocks][:group_footer_2] = text.include?('<%BEGIN_GROUP_FOOTER_2%>')
      config[:blocks][:group_footer] = text.include?('<%BEGIN_GROUP_FOOTER%>')
      config[:blocks][:total] = text.include?('<%BEGIN_TOTAL%>')

      # Поиск циклов
      config[:has_subtasks] = text.include?('<%BEGIN_SUBTASKS%>')
      config[:has_watchers] = text.include?('<%BEGIN_WATCHERS%>')
      config[:has_relations] = text.include?('<%BEGIN_RELATIONS%>')

      # Извлечение полей для валидации
      text.scan(/<%\s*([A-Za-z0-9_а-яА-ЯёЁ\s\.\:]+)\s*%>/) do |match|
        field_name = match[0].strip
        next if field_name.match?(/^(BEGIN_|END_|IF|ELSE|GROUP_BY|GROUP_BY_2|row_number|GroupValue|count|sum|avg|min|max|concat|total_|ExportDate|ExportUser|ProjectName|QueryName|FilterDescription|now|date|upper|lower|capitalize|truncate|strip_html|nl2br|replace|number|default|length)/i)
        
        config[:fields] << field_name unless config[:fields].include?(field_name)
      end

      config
    end

    private

    def extract_text
      text = ''
      Zip::File.open(@file_path) do |zip_file|
        if @file_type == :docx
          extract_docx_parts(zip_file).each { |entry| text += ' ' + extract_xml_text(entry, '//w:t', 'w' => 'http://schemas.openxmlformats.org/wordprocessingml/2006/main') }
        else
          extract_xlsx_parts(zip_file).each { |entry| text += ' ' + extract_xml_text(entry, '//si//t', 'xmlns' => 'http://schemas.openxmlformats.org/spreadsheetml/2006/main') }
          # Дополнительно проверяем ячейки напрямую
          zip_file.glob('xl/worksheets/sheet*.xml').each { |entry| text += ' ' + extract_xml_text(entry, '//c/v', 'xmlns' => 'http://schemas.openxmlformats.org/spreadsheetml/2006/main') }
        end
      end
      text
    end

    def extract_docx_parts(zip_file)
      parts = []
      parts << zip_file.find_entry('word/document.xml')
      parts.concat(zip_file.glob('word/header*.xml'))
      parts.concat(zip_file.glob('word/footer*.xml'))
      parts.compact
    end

    def extract_xlsx_parts(zip_file)
      [zip_file.find_entry('xl/sharedStrings.xml')].compact
    end

    def extract_xml_text(entry, xpath, ns)
      xml = Nokogiri::XML(entry.get_input_stream.read)
      xml.xpath(xpath, ns).map(&:text).join(' ')
    end
  end
end