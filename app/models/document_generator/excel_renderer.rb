# frozen_string_literal: true

module DocumentGenerator
  # ============================================================================
  # РЕНДЕРИНГ EXCEL-ДОКУМЕНТОВ
  # ============================================================================
  # Класс отвечает за генерацию Excel-документов (.xlsx) из шаблонов.
  # Обрабатывает группировку, агрегатные функции, циклы по строкам.
  
  class ExcelRenderer
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
    end

    # Генерация документа
    #
    # @return [String] Путь к сгенерированному файлу
    # @raise [RenderError] если произошла ошибка при рендеринге
    def render
      # Строим контекст данных
      context = ContextBuilder.new(@issues, @parser_config, @error_behavior).build
      
      # Путь для выходного файла
      output_path = "#{@template_path}.output.xlsx"
      
      # Файлы для обработки
      xml_targets = ['xl/worksheets/sheet*.xml']
      
      # Обрабатываем архив
      TemplateProcessor.process_archive(@template_path, output_path, xml_targets) do |doc, entry_name|
        process_excel_xml(doc, context, entry_name)
      end
      
      output_path
    rescue StandardError => e
      Rails.logger.error "[DocumentGenerator] Excel render failed: #{e.message}\n#{e.backtrace&.join("\n")}"
      error_msg = I18n.t('document_generator.error_excel_render_failed', message: e.message)
      handle_error(error_msg)
    end

    private

    # Обработка XML документа Excel
    #
    # @param doc [Nokogiri::XML::Document] XML-документ
    # @param context [Hash] Контекст данных
    # @param entry_name [String] Имя файла в архиве
    def process_excel_xml(doc, context, entry_name)
      ns = { 'xmlns' => 'http://schemas.openxmlformats.org/spreadsheetml/2006/main' }
      records = context['records'] || []
      
      # Находим все строки
      rows = doc.xpath('//xmlns:row', ns)
      
      # Определяем роли строк по первой ячейке
      row_roles = {}
      rows.each_with_index do |row, idx|
        first_cell = row.xpath('xmlns:c[1]/xmlns:v | xmlns:c[1]/xmlns:t', ns).first
        if first_cell
          text = first_cell.text.strip
          row_roles[idx] = text if ['GROUP_HEADER', 'GROUP_HEADER_2', 'ROW', 'GROUP_FOOTER_2', 'GROUP_FOOTER', 'TOTAL'].include?(text)
        end
      end
      
      # Если есть ROW - обрабатываем циклы
      if row_roles.values.include?('ROW') && records.present?
        process_excel_rows(doc, rows, context, records, row_roles, ns)
      end
      
      # Обрабатываем остальные строки (условия и подстановка)
      rows.each do |row|
        # Пропускаем служебные строки
        next if row_roles.key?(rows.index(row))
        
        # 1. Обрабатываем коллекции
        TemplateProcessor.process_collection_blocks(row, context, ns, @error_behavior)
        # 2. Обрабатываем условия
        TemplateProcessor.process_conditionals_in_block(row, context, ns, @error_behavior)
        # 3. Подставляем значения
        TemplateProcessor.substitute_in_block(row, context, ns, @error_behavior)
      end
    end

    # Обработка строк Excel с циклами
    #
    # @param doc [Nokogiri::XML::Document] XML-документ
    # @param rows [Nokogiri::XML::NodeSet] Набор строк
    # @param context [Hash] Общий контекст данных
    # @param records [Array<Hash>] Массив записей для цикла
    # @param row_roles [Hash] Хэш ролей строк (индекс => роль)
    # @param ns [Hash] Пространства имен XML
    def process_excel_rows(doc, rows, context, records, row_roles, ns)
      # Находим индексы строк для клонирования
      row_indices = row_roles.select { |_, role| role == 'ROW' }.keys
      
      return if row_indices.empty?
      
      # Для каждой записи клонируем строки
      records.each_with_index do |record, record_idx|
        merged_context = context.merge(record)
        
        row_indices.each do |row_idx|
          original_row = rows[row_idx]
          clone = original_row.dup
          
           # 1. Обрабатываем коллекции внутри клонированной строки
          TemplateProcessor.process_collection_blocks(clone, merged_context, ns, @error_behavior)
          # 2. Обрабатываем условия
          TemplateProcessor.process_conditionals_in_block(clone, merged_context, ns, @error_behavior)
          # 3. Подставляем значения
          TemplateProcessor.substitute_in_block(clone, merged_context, ns, @error_behavior)

          # Вставляем клон после оригинала
          original_row.add_next_sibling(clone)
        end
      end
      
      # Удаляем оригинальные строки с маркерами
      row_indices.sort.reverse.each do |idx|
        rows[idx].remove
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
# v2609231538