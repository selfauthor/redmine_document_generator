# frozen_string_literal: true

require 'rubyXL'
require 'fileutils'

module DocumentGenerator
  # Класс отвечает за генерацию документа Excel (.xlsx) на основе шаблона и данных.
  class ExcelRenderer
    # @param template_path [String] Путь к временному файлу шаблона
    # @param issues [ActiveRecord::Relation] Выборка задач для выгрузки
    # @param parser_config [Hash] Конфигурация, полученная от TemplateParser
    # @param error_behavior [String] Стратегия обработки ошибок
    def initialize(template_path, issues, parser_config, error_behavior)
      @template_path = template_path
      @issues = issues
      @parser_config = parser_config
      @error_behavior = error_behavior
      @workbook = nil
      @worksheet = nil
      @template_rows = {}
    end

    # Основной метод генерации документа
    # @return [String] Путь к сгенерированному временному файлу .xlsx
    def render
      @workbook = RubyXL::Parser.parse(@template_path)
      @worksheet = @workbook[0]

      identify_template_rows
      context = ContextBuilder.new(@issues, @parser_config, @error_behavior).build

      if @parser_config[:group_by]
        render_grouped_mode(context)
      else
        render_flat_mode(context)
      end

      output_path = generate_output_path
      @workbook.write(output_path)
      output_path
    end

    private

    def identify_template_rows
      roles = %w[GROUP_HEADER GROUP_HEADER_2 ROW GROUP_FOOTER_2 GROUP_FOOTER TOTAL]
      
      @worksheet.sheet_data.rows.each_with_index do |row, index|
        next unless row
        
        first_cell = row.cells.first
        next unless first_cell && first_cell.value.is_a?(String)
        
        cell_value = first_cell.value.strip.upcase
        if roles.include?(cell_value)
          @template_rows[cell_value.downcase.to_sym] = index
          first_cell.change_contents('')
        end
      end
    end

    def render_flat_mode(context)
      row_idx = @template_rows[:row]
      handle_error('Row template missing') unless row_idx

      new_rows = []
      (0...row_idx).each { |i| new_rows << @worksheet.sheet_data.rows[i] }

      context['records'].each_with_index do |record, index|
        new_row = clone_row(@worksheet.sheet_data.rows[row_idx], index)
        replace_fields_in_row(new_row, record, index + 1)
        new_rows << new_row
      end

      last_idx = @template_rows.values.compact.max || row_idx
      ((last_idx + 1)...@worksheet.sheet_data.rows.size).each do |i|
        new_rows << @worksheet.sheet_data.rows[i]
      end

      @worksheet.sheet_data.rows = new_rows
    end

    def render_grouped_mode(context)
      new_rows = []
      first_idx = @template_rows.values.compact.min || 0
      (0...first_idx).each { |i| new_rows << @worksheet.sheet_data.rows[i] }

      global_row_num = 1
      context['groups'].each do |group|
        if @template_rows[:group_header]
          header_row = clone_row(@worksheet.sheet_data.rows[@template_rows[:group_header]], new_rows.size)
          replace_fields_in_row(header_row, group, global_row_num)
          replace_aggregates_in_row(header_row, group, 'group_')
          new_rows << header_row
        end

        group['records'].each do |record|
          row_template = @worksheet.sheet_data.rows[@template_rows[:row]]
          if row_template
            new_row = clone_row(row_template, new_rows.size)
            replace_fields_in_row(new_row, record, global_row_num)
            new_rows << new_row
          end
          global_row_num += 1
        end

        if @template_rows[:group_footer]
          footer_row = clone_row(@worksheet.sheet_data.rows[@template_rows[:group_footer]], new_rows.size)
          replace_aggregates_in_row(footer_row, group, 'group_')
          new_rows << footer_row
        end
      end

      if @template_rows[:total]
        total_row = clone_row(@worksheet.sheet_data.rows[@template_rows[:total]], new_rows.size)
        replace_aggregates_in_row(total_row, context['totals'].first, 'total_')
        new_rows << total_row
      end

      last_idx = @template_rows.values.compact.max || 0
      ((last_idx + 1)...@worksheet.sheet_data.rows.size).each do |i|
        new_rows << @worksheet.sheet_data.rows[i]
      end

      @worksheet.sheet_data.rows = new_rows
    end

    def clone_row(template_row, target_index)
      return nil unless template_row
      
      new_row = RubyXL::Row.new(worksheet: @worksheet, row_index: target_index)
      template_row.cells.each do |cell|
        next unless cell
        
        new_cell = RubyXL::Cell.new(
          worksheet: @worksheet,
          row_index: target_index,
          column_index: cell.column_index,
          value: cell.value,
          style_index: cell.style_index,
          type: cell.type
        )
        new_row.add_cell(new_cell)
      end
      new_row
    end

    def replace_fields_in_row(row, record, global_row_num)
      row.cells.each do |cell|
        next unless cell && cell.value.is_a?(String)
        
        new_value = cell.value.dup
        new_value = new_value.gsub(/<%\s*row_number\s*%>/i, global_row_num.to_s)
        new_value = new_value.gsub(/<%\s*row_number_in_group\s*%>/i, record['row_number_in_group'].to_s) if record['row_number_in_group']
        new_value = new_value.gsub(/<%\s*GroupValue\s*%>/i, record['GroupValue'].to_s) if record['GroupValue']
        new_value = new_value.gsub(/<%\s*GroupValue2\s*%>/i, record['GroupValue2'].to_s) if record['GroupValue2']
        new_value = new_value.gsub(/<%\s*count\s*%>/i, record['count'].to_s) if record['count']

        new_value.gsub!(/<%\s*([^%]+?)\s*%>/) do |match|
          field_name = $1.strip
          if record.key?(field_name)
            ContextBuilder.format_value(record[field_name])
          else
            handle_error("Unknown field: #{field_name}")
            match
          end
        end
        
        cell.change_contents(new_value) if new_value != cell.value
      end
    end

    def replace_aggregates_in_row(row, record, prefix)
      row.cells.each do |cell|
        next unless cell && cell.value.is_a?(String)
        
        new_value = cell.value.dup
        
        new_value.gsub!(/<%\s*(total_)?(count|sum|avg|min|max|concat)\s*\(\s*([^%]+?)\s*\)\s*%>/i) do |match|
          is_total = $1.present?
          func = $2.downcase
          field = $3.strip
          
          key = "#{is_total ? 'total_' : 'group_'}agg_#{func}_#{field}"
          ContextBuilder.format_value(record[key])
        end
        
        new_value.gsub!(/<%\s*total_count\s*%>/i, record['total_count'].to_s) if record['total_count']

        cell.change_contents(new_value) if new_value != cell.value
      end
    end

    def handle_error(message)
      case @error_behavior
      when 'abort'
        raise I18n.t('document_generator.error_invalid_template', message: message)
      when 'skip_field', 'skip_record'
        ''
      else
        ''
      end
    end

    def generate_output_path
      ext = File.extname(@template_path)
      dir = Dir.mktmpdir
      File.join(dir, "output#{ext}")
    end
  end
end