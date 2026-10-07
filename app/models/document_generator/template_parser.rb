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

    # Разбирает шаблон и формирует конфигурацию его структуры.
    #
    # Определяет:
    # - параметры группировки;
    # - наличие управляющих строк групп;
    # - наличие блока строк записей;
    # - наличие итогового блока;
    # - наличие циклов subtasks/watchers/relations;
    # - используемые поля шаблона.
    #
    # @return [Hash] Конфигурация разобранного шаблона.
    def parse
      text = extract_text

      config = {
        template_text: text,
        group_by: nil,
        group_by_2: nil,
        blocks: {
          group_header: false,
          group_header_2: false,
          row: false,
          group_footer_2: false,
          group_footer: false,
          end_groups: false,
          total: false
        },
        fields: [],
        has_subtasks: false,
        has_watchers: false,
        has_relations: false
      }

      # GROUP_BY и GROUP_BY_2 задают поля, по которым формируется
      # первый и второй уровни вложенной группировки соответственно.
      config[:group_by] = text[
        /<%\s*GROUP_BY\s*:\s*([^%]+?)\s*%>/i,
        1
      ]&.strip

      config[:group_by_2] = text[
        /<%\s*GROUP_BY_2\s*:\s*([^%]+?)\s*%>/i,
        1
      ]&.strip

      # Управляющие строки группировки являются самостоятельными маркерами.
      # Они не имеют BEGIN_*/END_* и удаляются из итогового документа.
      config[:blocks][:group_header] = text.match?(
        /<%\s*GROUP_HEADER\s*%>/i
      )

      config[:blocks][:group_header_2] = text.match?(
        /<%\s*GROUP_HEADER_2\s*%>/i
      )

      config[:blocks][:row] = text.match?(
        /<%\s*BEGIN_ROW\s*%>/i
      )

      config[:blocks][:group_footer_2] = text.match?(
        /<%\s*GROUP_FOOTER_2\s*%>/i
      )

      config[:blocks][:group_footer] = text.match?(
        /<%\s*GROUP_FOOTER\s*%>/i
      )

      # END_GROUPS является единственным маркером, закрывающим всю
      # повторяемую область группировки.
      config[:blocks][:end_groups] = text.match?(
        /<%\s*END_GROUPS\s*%>/i
      )

      config[:blocks][:total] = text.match?(
        /<%\s*BEGIN_TOTAL\s*%>/i
      )

      # Поиск циклов коллекций.
      config[:has_subtasks] = text.match?(
        /<%\s*BEGIN_SUBTASKS\s*%>/i
      )

      config[:has_watchers] = text.match?(
        /<%\s*BEGIN_WATCHERS\s*%>/i
      )

      config[:has_relations] = text.match?(
        /<%\s*BEGIN_RELATIONS\s*%>/i
      )

      # При использовании группировки обязательны GROUP_FOOTER и
      # END_GROUPS. Первый закрывает внешний подвал группы, второй
      # закрывает всю повторяемую область группировки.
      if config[:group_by] || config[:group_by_2]
        unless config[:blocks][:group_footer]
          raise DocumentGenerator::TemplateError,
                I18n.t('document_generator.error_missing_group_footer')
        end

        end_groups_count = text.scan(/<%\s*END_GROUPS\s*%>/i).length

        if end_groups_count.zero?
          raise DocumentGenerator::TemplateError,
                I18n.t('document_generator.error_missing_end_groups')
        elsif end_groups_count > 1
          raise DocumentGenerator::TemplateError,
                I18n.t('document_generator.error_duplicate_end_groups')
        end
      elsif text.match?(/<%\s*END_GROUPS\s*%>/i)
        raise DocumentGenerator::TemplateError,
              I18n.t('document_generator.error_unexpected_end_groups')
      end

      # Второй уровень группировки не должен существовать без первого.
      if config[:group_by_2] && !config[:group_by]
        raise DocumentGenerator::TemplateError,
              I18n.t('document_generator.error_group_by_2_without_group_by')
      end

      # Для группировки требуется заголовок первого уровня.
      if (config[:group_by] || config[:group_by_2]) &&
         !config[:blocks][:group_header]
        raise DocumentGenerator::TemplateError,
              I18n.t('document_generator.error_missing_group_header')
      end

      # Извлекаем обычные поля и поля, передаваемые в функции.
      config[:fields] = extract_template_fields(text)

      config
    end

    private

    # Извлекает из шаблона имена полей, используемых обычными маркерами
    # и аргументами функций.
    #
    # В пользовательском синтаксисе отсутствует префикс CF:.
    # Поэтому стандартные и пользовательские поля обрабатываются одинаково.
    #
    # @param text [String] Полный текст шаблона.
    # @return [Array<String>] Уникальный список используемых полей.
    def extract_template_fields(text)
      fields = []

      # Обычные маркеры вида <%Поле%>.
      text.scan(/<%\s*([^%]+?)\s*%>/) do |match|
        expression = match[0].strip

        # Управляющие конструкции не являются полями.
        next if expression.match?(
          /\A(?:
            BEGIN_[A-Z0-9_]+ |
            END_[A-Z0-9_]+ |
            GROUP_BY(?:_2)?\s*: |
            GROUP_HEADER(?:_2)? |
            GROUP_FOOTER(?:_2)? |
            IF\b |
            ELSE\b |
            row_number\b |
            GroupValue\b |
            GroupValue2\b |
            count\b |
            total_count\b |
            ExportDate\b |
            ExportUser\b |
            ProjectName\b |
            QueryName\b |
            FilterDescription\b
          )/ix
        )

        # Функциональные выражения обрабатываются отдельно ниже.
        next if expression.match?(/\A[a-z_][a-z0-9_]*\s*\(/i)

        fields << expression unless fields.include?(expression)
      end

      # Извлекаем аргументы функций форматирования:
      # upper(Поле), lower(Поле), default(Поле, "значение") и т. д.
      text.scan(
        /<%\s*
          (?:upper|lower|capitalize|truncate|strip_html|nl2br|
             number|default|length|concat|replace|date)
          \s*\(([^%]+)\)
        \s*%>/ix
      ) do |match|
        extract_function_argument_fields(match[0]).each do |field|
          fields << field unless fields.include?(field)
        end
      end

      # Извлекаем поля агрегатных функций.
      #
      # count не содержит имени поля, поэтому отдельно не добавляется.
      text.scan(
        /<%\s*
          (?:total_)?
          (?:sum|avg|min|max)
          \s*\(\s*([^()]+?)\s*\)
        \s*%>/ix
      ) do |match|
        field = match[0].strip
        fields << field unless fields.include?(field)
      end

      fields
    end

    # Извлекает имена полей из аргументов функции.
    #
    # @param arguments [String] Строка аргументов функции.
    # @return [Array<String>] Список аргументов, которые являются именами полей.
    def extract_function_argument_fields(arguments)
      fields = []

      # Разделяем аргументы только по запятым верхнего уровня.
      # Запятые внутри строковых литералов не считаются разделителями.
      parts = []
      current = +''
      quote = nil
      escaped = false

      arguments.each_char do |char|
        if escaped
          current << char
          escaped = false
          next
        end

        if quote
          current << char

          if char == '\\'
            escaped = true
          elsif char == quote
            quote = nil
          end

          next
        end

        if char == "'" || char == '"'
          quote = char
          current << char
        elsif char == ','
          parts << current.strip
          current = +''
        else
          current << char
        end
      end

      parts << current.strip unless current.strip.empty?

      parts.each do |part|
        # Строковые литералы не являются полями.
        next if part.match?(/\A(['"]).*\1\z/m)

        # Числовые аргументы не являются полями.
        next if part.match?(/\A-?\d+(?:\.\d+)?\z/)

        # Имя поля может содержать Unicode-буквы, цифры, пробелы,
        # точки и символ подчёркивания.
        if part.match?(/\A[\p{L}_][\p{L}\p{N}_. ]*\z/u)
          fields << part
        end
      end

      fields
    end

    # Извлекает весь текст из XML-файлов шаблона, необходимый для анализа
    # структуры и поиска управляющих маркеров.
    #
    # Для DOCX читаются текстовые узлы w:t.
    # Для XLSX сначала читаются строки sharedStrings.xml, где обычно находятся
    # текстовые значения ячеек и управляющие маркеры шаблона.
    #
    # @return [String] Объединённый текст шаблона.
    def extract_text
      text = ''

      Zip::File.open(@file_path) do |zip_file|
        if @file_type == :docx
          extract_docx_parts(zip_file).each do |entry|
            text += ' ' + extract_xml_text(
              entry,
              '//w:t',
              'w' => 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
            )
          end
        else
          # sharedStrings.xml использует namespace Excel SpreadsheetML.
          # Поэтому префикс namespace должен присутствовать непосредственно
          # в XPath, иначе Nokogiri не найдёт элементы si и t.
          extract_xlsx_parts(zip_file).each do |entry|
            text += ' ' + extract_xml_text(
              entry,
              '//xmlns:si//xmlns:t',
              'xmlns' => 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
            )
          end

          # Дополнительно проверяем значения непосредственно в worksheet XML.
          # Это необходимо для XLSX-файлов, в которых часть строк хранится
          # непосредственно в XML листа, а не в sharedStrings.xml.
          zip_file.glob('xl/worksheets/sheet*.xml').each do |entry|
            text += ' ' + extract_xml_text(
              entry,
              '//xmlns:c//xmlns:v',
              'xmlns' => 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
            )
          end
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
# v2610071347