# frozen_string_literal: true

require 'sablon'
require 'zip'
require 'fileutils'
require 'nokogiri'

module DocumentGenerator
  # Класс отвечает за генерацию документа Word (.docx) на основе шаблона и данных.
  class WordRenderer
    # @param template_path [String] Путь к временному файлу шаблона
    # @param issues [ActiveRecord::Relation] Выборка записей для выгрузки
    # @param parser_config [Hash] Конфигурация, полученная от TemplateParser
    # @param error_behavior [String] Стратегия обработки ошибок
    def initialize(template_path, issues, parser_config, error_behavior)
      @template_path = template_path
      @issues = issues
      @parser_config = parser_config
      @error_behavior = error_behavior
    end

    # Основной метод генерации документа
    # @return [String] Бинарные данные сгенерированного документа .docx
    def render
      context = ContextBuilder.new(@issues, @parser_config, @error_behavior).build
      inject_html_descriptions(context)
      processed_template_path = preprocess_template

      begin
        template = Sablon.template(processed_template_path)
        template.render(context)
      rescue Sablon::Error, StandardError => e
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
      FileUtils.cp(@template_path, temp_path)

      Zip::File.open(temp_path) do |zip|
        xml_entries = zip.glob('word/document.xml') +
                      zip.glob('word/header*.xml') +
                      zip.glob('word/footer*.xml')

        xml_entries.each do |entry|
          xml_content = entry.get_input_stream.read

          xml_content = xml_content.gsub(/<%\s*([^%]+?)\s*%>/) do |match|
            inner = $1.strip

            if inner.match?(/^(date|now|upper|lower|default|strip_html)\s*\(/i)
              func = inner[/^([a-z_]+)/i, 1].downcase
              args = inner[/\((.*)\)$/, 1]
              "«#{func}_#{args}»"
            elsif inner.match?(/^(sum|avg|min|max|count|concat)\s*\(/i)
              func = inner[/^([a-z_]+)/i, 1].downcase
              field = inner[/\((.*?)\)$/, 1].strip
              "«agg_#{func}_#{field}»"
            elsif inner.match?(/^IF\((.*?)\)$/i)
              "«if #{$1.strip}»"
            elsif inner.match?(/^ELSE$/i)
              "«else»"
            elsif inner.match?(/^END$/i)
              "«end»"
            elsif inner.match?(/^BEGIN_ROW$/i)
              "«tr records»"
            elsif inner.match?(/^END_ROW$/i)
              "«end»"
            elsif inner.match?(/^BEGIN_SUBTASKS$/i)
              "«tr subtasks»"
            elsif inner.match?(/^END_SUBTASKS$/i)
              "«end»"
            elsif inner.match?(/^BEGIN_WATCHERS$/i)
              "«tr watchers»"
            elsif inner.match?(/^END_WATCHERS$/i)
              "«end»"
            elsif inner.match?(/^BEGIN_RELATIONS$/i)
              "«tr relations»"
            elsif inner.match?(/^END_RELATIONS$/i)
              "«end»"
            elsif inner.match?(/^BEGIN_GROUP_HEADER$/i)
              "«tr groups»"
            elsif inner.match?(/^END_GROUP_HEADER$/i)
              ""
            elsif inner.match?(/^BEGIN_GROUP_FOOTER$/i)
              ""
            elsif inner.match?(/^END_GROUP_FOOTER$/i)
              "«end»"
            elsif inner.match?(/^BEGIN_TOTAL$/i)
              "«tr totals»"
            elsif inner.match?(/^END_TOTAL$/i)
              "«end»"
            elsif inner.match?(/^(Subtask|Relation|Watcher)\.(.+)$/i)
              "«#{$2.strip}»"
            else
              "«#{inner}»"
            end
          end

          entry.get_output_stream.write(xml_content)
        end
      end

      temp_path
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