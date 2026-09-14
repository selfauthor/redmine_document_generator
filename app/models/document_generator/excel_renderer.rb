# frozen_string_literal: true

require 'rubyXL'
require 'fileutils'

# v2609141507
module DocumentGenerator
  # ExcelRenderer is responsible for generating Excel documents (.xlsx).
  # It uses the rubyXL library to preserve formatting,
  # but relies on TemplateProcessor for marker substitution in cell values.
  class ExcelRenderer
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
    # @return [String] Path to the generated .xlsx file
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

    # Processes a specific Excel worksheet.
    # Finds rows with the ROW marker, clones them for each record,
    # and replaces markers in the cell values.
    #
    # @param worksheet [RubyXL::Worksheet] The Excel worksheet to process
    # @param context [Hash] Data for substitution
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

      # Process rows in reverse order to safely insert clones without breaking indices
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
              new_cell.value = TemplateProcessor.substitute_markers(new_cell.value, record)
            end
            new_cell
          end

          worksheet.sheet_data.add_row(new_row_cells, row_idx + 1)
        end

        worksheet.delete_row(row_idx)
      end

      # If no loops existed, just replace markers globally (for headers/footers)
      if rows_to_process.empty?
        render_context = context['records'].first || context
        worksheet.sheet_data.rows.each do |row|
          next unless row
          row.cells.each do |cell|
            next unless cell && cell.value.is_a?(String)
            cell.value = TemplateProcessor.substitute_markers(cell.value, render_context)
          end
        end
      end
      
      worksheet.sheet_data.rows.compact!
    end

    # Removes control markers from cell values in a row.
    #
    # @param row [RubyXL::Row] The Excel row
    def clean_row_markers(row)
      row.cells.each do |cell|
        next unless cell && cell.value.is_a?(String)
        
        cell.value = cell.value.gsub(/<%\s*(BEGIN_ROW|END_ROW|GROUP_BY|GROUP_BY_2)\s*%>/i, '').strip
        cell.value = nil if cell.value.empty?
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