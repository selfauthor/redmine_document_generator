# frozen_string_literal: true
require 'rubyXL'
require 'fileutils'

module DocumentGenerator
  # ExcelRenderer отвечает за генерацию документов Excel (.xlsx).
  # Использует библиотеку rubyXL для сохранения форматирования,
  # но полагается на TemplateProcessor для подстановки маркеров и условий.
  class ExcelRenderer
    # @param template_path [String] Путь к временному файлу шаблона
    # @param issues [ActiveRecord::Relation] Коллекция записей для экспорта
    # @param parser_config [Hash] Конфигурация из TemplateParser
    # @param error_behavior [String] Стратегия обработки ошибок
    def initialize(template_path, issues, parser_config, error_behavior)
      @template_path = template_path
      @issues = issues
      @parser_config = parser_config
      @error_behavior = error_behavior
    end

    # Основной метод генерации документа.
    #
    # @return [String] Путь к сгенерированному файлу .xlsx
    def render
      context = ContextBuilder.new(@issues, @parser_config, @error_behavior).build
      output_path = "#{@template_path}.output.xlsx"
      
      begin
        workbook = RubyXL::Parser.parse(@template_path)
        workbook.worksheets.each do |worksheet|
          process_excel_worksheet(worksheet, context)
        end
        workbook.write(output_path)
        output_path
      rescue StandardError => e
        Rails.logger.error "[DocumentGenerator] Excel render failed: #{e.message}\n#{e.backtrace&.join("\n")}"
        error_msg = I18n.t('document_generator.error_excel_render_failed', message: e.message)
        handle_error(error_msg)
      end
    end

    private

    # Обрабатывает рабочий лист Excel:
    # 1. Находит строки с маркером ROW
    # 2. Клонирует их для каждой записи
    # 3. Применяет условия и подставляет значения
    #
    # @param worksheet [RubyXL::Worksheet] Рабочий лист Excel
    # @param context [Hash] Данные для подстановки
    def process_excel_worksheet(worksheet, context)
      return unless worksheet.sheet_data
      
      rows_to_process = []
      worksheet.sheet_data.rows.each_with_index do |row, row_idx|
        next unless row
        row_text = row.cells.map { |c| c&.value.to_s }.join
        if row_text.include?('ROW') || row_text.include?('<%BEGIN_ROW%>')
          rows_to_process << row_idx
        end
      end

      # Обрабатываем строки в обратном порядке для безопасного клонирования
      rows_to_process.reverse_each do |row_idx|
        template_row = worksheet.sheet_data.rows[row_idx]
        next unless template_row
        
        records = context['records'] || [context]
        clean_row_markers(template_row)
        
        records.each do |record|
          new_row_cells = template_row.cells.map do |cell|
            next nil unless cell
            new_cell = cell.dup
            if new_cell.value.is_a?(String)
              # СПЕЦИФИКА EXCEL: применяем общую логику шаблонов
              cell_value = new_cell.value
              cell_value = TemplateProcessor.resolve_conditionals(cell_value, record)
              cell_value = TemplateProcessor.clean_control_markers(cell_value)
              cell_value = TemplateProcessor.substitute_markers(cell_value, record)
              new_cell.value = cell_value
            end
            new_cell
          end
          worksheet.sheet_data.add_row(new_row_cells, row_idx + 1)
        end
        
        worksheet.delete_row(row_idx)
      end

      # Если циклов не было, просто заменяем маркеры глобально
      if rows_to_process.empty?
        render_context = context['records'].first || context
        worksheet.sheet_data.rows.each do |row|
          next unless row
          row.cells.each do |cell|
            next unless cell && cell.value.is_a?(String)
            # СПЕЦИФИКА EXCEL: применяем общую логику шаблонов
            cell_value = cell.value
            cell_value = TemplateProcessor.resolve_conditionals(cell_value, render_context)
            cell_value = TemplateProcessor.clean_control_markers(cell_value)
            cell_value = TemplateProcessor.substitute_markers(cell_value, render_context)
            cell.value = cell_value
          end
        end
      end
      
      worksheet.sheet_data.rows.compact!
    end

    # СПЕЦИФИКА EXCEL: удаляет управляющие маркеры из ячеек строки
    #
    # @param row [RubyXL::Row] Строка Excel
    def clean_row_markers(row)
      row.cells.each do |cell|
        next unless cell && cell.value.is_a?(String)
        cell.value = TemplateProcessor.clean_control_markers(cell.value).strip
        cell.value = nil if cell.value.empty?
      end
    end

    # Универсальный обработчик ошибок
    #
    # @param message [String] Сообщение об ошибке (уже локализованное)
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
# v2609151130