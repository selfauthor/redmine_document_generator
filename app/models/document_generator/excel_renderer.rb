# frozen_string_literal: true

module DocumentGenerator
  # ============================================================================
  # РЕНДЕРИНГ EXCEL-ДОКУМЕНТОВ
  # ============================================================================
  # Класс отвечает за генерацию Excel-документов (.xlsx) из шаблонов.
  # Обрабатывает группировку, агрегатные функции, циклы по строкам.
  
  class ExcelRenderer
    # Возвращает предупреждения, накопленные при формировании документа.
    # @return [Array<String>] Список предупреждений для пользователя.
    attr_reader :warnings

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

      # Создаём массив для предупреждений текущей выгрузки.
      @warnings = []
    end

    # Генерирует Excel-документ из шаблона.
    #
    # @return [String] Путь к сгенерированному XLSX-файлу.
    # @raise [RenderError] Если во время рендеринга произошла ошибка.
    def render
      # Формируем общий контекст данных для всех записей выгрузки.
      context = ContextBuilder.new(@issues, @parser_config, @error_behavior).build

      # Используем предупреждения, сформированные ContextBuilder.
      @warnings = context['__warnings'] || @warnings

      # Формируем путь к результирующему XLSX-файлу.
      output_path = "#{@template_path}.output.xlsx"

      # Определяем XML-файлы листов Excel, которые требуется обработать.
      xml_targets = ['xl/worksheets/sheet*.xml']

      # Читаем sharedStrings.xml один раз до обработки листов.
      # Это необходимо, поскольку Excel часто хранит текстовые значения
      # ячеек не непосредственно в worksheet XML, а по индексу общей строки.
      shared_strings = []

      Zip::File.open(@template_path) do |zip_file|
        shared_strings_entry = zip_file.find_entry('xl/sharedStrings.xml')

        if shared_strings_entry
          shared_strings_doc = Nokogiri::XML(
            zip_file.read('xl/sharedStrings.xml')
          )

          # Извлекаем текст каждого элемента <si>.
          # Если строка состоит из нескольких <t>, объединяем их в одно значение.
          shared_strings = shared_strings_doc.xpath(
            "//*[local-name()='si']"
          ).map do |string_item|
            string_item.xpath(".//*[local-name()='t']").map(&:text).join
          end
        end
      end

      # Обрабатываем XML-файлы листов внутри XLSX-архива.
      TemplateProcessor.process_archive(
        @template_path,
        output_path,
        xml_targets
      ) do |doc, entry_name|
        process_excel_xml(
          doc,
          context,
          entry_name,
          shared_strings
        )
      end

      output_path
    rescue StandardError => e
      # Записываем техническую информацию в журнал для диагностики ошибки.
      Rails.logger.error(
        "[DocumentGenerator] Excel render failed: #{e.message}\n#{e.backtrace&.join("\n")}"
      )

      # Формируем локализованное сообщение для пользователя.
      error_msg = I18n.t(
        'document_generator.error_excel_render_failed',
        message: e.message
      )

      handle_error(error_msg)
    end

    private

    # Обрабатывает XML отдельного листа Excel.
    #
    # @param doc [Nokogiri::XML::Document] XML-документ листа Excel.
    # @param context [Hash] Общий контекст данных выгрузки.
    # @param entry_name [String] Имя XML-файла листа внутри XLSX-архива.
    # @param shared_strings [Array<String>] Значения из xl/sharedStrings.xml.
    # @return [void]
    def process_excel_xml(doc, context, entry_name, shared_strings)
      ns = {
        'xmlns' => 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
      }

      records = context['records'] || []

      # Преобразуем ячейки типа shared string в inlineStr.
      # После этого все текстовые маркеры Excel доступны непосредственно
      # внутри XML листа и могут обрабатываться общей логикой TemplateProcessor.
      convert_shared_strings(doc, shared_strings, ns)

      # Обрабатываем основной цикл записей.
      #
      # BEGIN_ROW и END_ROW являются отдельными строками Excel.
      # Все строки между ними образуют шаблон одного элемента цикла.
      process_excel_row_blocks(
        doc,
        context,
        records,
        ns
      )

      # После разворачивания BEGIN_ROW/END_ROW обрабатываем оставшиеся строки.
      #
      # Здесь могут находиться:
      # - обычные поля;
      # - IF/ELSE/END;
      # - одиночные коллекции;
      # - статический текст.
      doc.xpath('//xmlns:row', ns).each do |row|
        # Разворачиваем коллекции, если их управляющие команды находятся
        # в пределах текущего XML-узла.
        TemplateProcessor.process_collection_blocks(
          row,
          context.merge(records.first || {}),
          ns,
          @error_behavior
        )

        # Обрабатываем условия IF/ELSE/END внутри ячеек.
        TemplateProcessor.process_conditionals_in_block(
          row,
          context.merge(records.first || {}),
          ns,
          @error_behavior
        )

        # Подставляем обычные поля шаблона.
        TemplateProcessor.substitute_in_block(
          row,
          context.merge(records.first || {}),
          ns,
          @error_behavior
        )
      end

      # После удаления шаблонных строк и вставки их копий Excel должен получить
      # последовательную нумерацию строк и адресов ячеек.
      reindex_excel_rows(doc, ns)
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

    # Преобразует ячейки Excel, использующие sharedStrings.xml,
    # в inlineStr с непосредственным текстом.
    #
    # @param doc [Nokogiri::XML::Document] XML-документ листа Excel.
    # @param shared_strings [Array<String>] Значения общей таблицы строк Excel.
    # @param ns [Hash] Пространства имён XML.
    # @return [void]
    def convert_shared_strings(doc, shared_strings, ns)
      # Находим только ячейки, которые используют shared string.
      doc.xpath('//xmlns:c[@t="s"]', ns).each do |cell|
        value_node = cell.at_xpath('xmlns:v', ns)

        # Ячейка без индекса общей строки не может быть преобразована.
        next unless value_node

        # Безопасно преобразуем значение <v> в индекс sharedStrings.xml.
        string_index = begin
          Integer(value_node.text, 10)
        rescue ArgumentError, TypeError
          nil
        end

        next if string_index.nil?
        next if string_index.negative?
        next if string_index >= shared_strings.length

        # Создаём inline string <is>.
        inline_string = Nokogiri::XML::Node.new(
          'is',
          doc
        )
        inline_string.namespace = cell.namespace

        # Создаём непосредственный текстовый узел <t>.
        text_node = Nokogiri::XML::Node.new(
          't',
          doc
        )
        text_node.namespace = cell.namespace
        text_node.content = shared_strings[string_index]

        # Сохраняем начальные и конечные пробелы текста.
        if text_node.content.match?(/\A\s|\s\z/)
          text_node['xml:space'] = 'preserve'
        end

        inline_string.add_child(text_node)

        # Удаляем числовой индекс shared string.
        value_node.remove

        # Меняем тип ячейки на inlineStr.
        cell['t'] = 'inlineStr'

        # Добавляем непосредственное текстовое содержимое.
        cell.add_child(inline_string)
      end
    end

    # Разворачивает блоки BEGIN_ROW/END_ROW на листе Excel.
    #
    # @param doc [Nokogiri::XML::Document] XML-документ листа Excel.
    # @param context [Hash] Общий контекст данных.
    # @param records [Array<Hash>] Записи, по которым выполняется цикл.
    # @param ns [Hash] Пространства имён XML.
    # @return [void]
    def process_excel_row_blocks(doc, context, records, ns)
      loop do
        # Получаем актуальный список строк после каждой операции вставки/удаления.
        rows = doc.xpath('//xmlns:row', ns).to_a

        # Ищем строку, содержащую начало основного цикла.
        begin_index = rows.index do |row|
          row_text = excel_row_text(row, ns)

          row_text.match?(
            /<%\s*BEGIN_ROW\s*%>/i
          )
        end

        # Больше циклов в листе нет.
        break unless begin_index

        # Ищем END_ROW после найденного BEGIN_ROW.
        end_index = nil

        ((begin_index + 1)...rows.length).each do |index|
          row_text = excel_row_text(rows[index], ns)

          if row_text.match?(
            /<%\s*END_ROW\s*%>/i
          )
            end_index = index
            break
          end
        end

        # BEGIN_ROW без END_ROW является ошибкой структуры шаблона.
        unless end_index
          error_msg = I18n.t(
            'document_generator.error_missing_end_row'
          )

          raise DocumentGenerator::TemplateError, error_msg
        end

        # Внутри одного основного ROW-блока другой BEGIN_ROW не допускается.
        # Вложенные циклы предназначены для SUBTASKS/WATCHERS/RELATIONS.
        nested_begin_index = nil

        ((begin_index + 1)...end_index).each do |index|
          row_text = excel_row_text(rows[index], ns)

          if row_text.match?(
            /<%\s*BEGIN_ROW\s*%>/i
          )
            nested_begin_index = index
            break
          end
        end

        if nested_begin_index
          raise DocumentGenerator::TemplateError,
                I18n.t('document_generator.error_row_block_mismatch')
        end

        # Строки между BEGIN_ROW и END_ROW являются шаблоном одной записи.
        template_rows = rows[
          (begin_index + 1)...end_index
        ].to_a

        # Сохраняем строку BEGIN_ROW как точку вставки.
        begin_row = rows[begin_index]

        # Формируем все строки, которые должны заменить исходный блок.
        expanded_rows = []

        records.each_with_index do |record, record_index|
          begin
            # Объединяем общий контекст с данными текущей задачи.
            merged_context = context.merge(record)

            # Создаём независимые XML-копии всех строк тела цикла.
            clones = template_rows.map(&:dup)

            # Обрабатываем вложенные коллекционные циклы:
            # SUBTASKS, WATCHERS и RELATIONS.
            clones = TemplateProcessor.process_collection_blocks(
              clones,
              merged_context,
              ns,
              @error_behavior
            )

            # Обрабатываем условия и обычные поля каждой строки.
            clones.each do |clone|
              TemplateProcessor.process_conditionals_in_block(
                clone,
                merged_context,
                ns,
                @error_behavior
              )

              TemplateProcessor.substitute_in_block(
                clone,
                merged_context,
                ns,
                @error_behavior
              )
            end

            # Добавляем готовые строки текущей записи в общий результат.
            expanded_rows.concat(clones)
          rescue DocumentGenerator::SkipRecordError
            # При skip_record пропускаем только текущую запись.
            issue = record['__issue']

            Rails.logger.warn(
              "[DocumentGenerator] Issue ##{issue&.id || 'unknown'} was skipped while processing Excel ROW block."
            )

            next
          end
        end

        # Вставляем готовые строки непосредственно перед BEGIN_ROW.
        #
        # Вставляем в обратном порядке, поскольку каждая новая строка
        # добавляется перед одной и той же исходной строкой BEGIN_ROW.
        expanded_rows.each do |clone|
          begin_row.add_previous_sibling(clone)
        end

        # Удаляем весь исходный блок:
        # BEGIN_ROW + тело шаблона + END_ROW.
        rows[
          begin_index..end_index
        ].each(&:remove)
      end
    end

    # Возвращает объединённый текст всех текстовых ячеек строки Excel.
    #
    # @param row [Nokogiri::XML::Node] XML-узел <row>.
    # @param ns [Hash] Пространства имён XML.
    # @return [String] Текст всех текстовых узлов строки.
    def excel_row_text(row, ns)
      row.xpath(
        ".//*[local-name()='t']"
      ).map(&:text).join
    end

    # Перенумеровывает строки и адреса ячеек после разворачивания циклов.
    #
    # @param doc [Nokogiri::XML::Document] XML-документ листа Excel.
    # @param ns [Hash] Пространства имён XML.
    # @return [void]
    def reindex_excel_rows(doc, ns)
      rows = doc.xpath(
        '//xmlns:row',
        ns
      )

      rows.each_with_index do |row, row_index|
        # Excel использует нумерацию строк начиная с 1.
        new_row_number = row_index + 1

        # Обновляем номер XML-строки.
        row['r'] = new_row_number.to_s

        # Обновляем адрес каждой ячейки этой строки.
        row.xpath(
          './xmlns:c'
        ).each do |cell|
          current_reference = cell['r'].to_s

          # Обычно адрес имеет вид A2, B2, AA15 и т.п.
          column_match = current_reference.match(
            /\A([A-Z]+)\d+\z/i
          )

          # Если адрес ячейки отсутствует или имеет нестандартный формат,
          # не пытаемся угадывать его автоматически.
          next unless column_match

          column_name = column_match[1].upcase

          cell['r'] = "#{column_name}#{new_row_number}"
        end
      end
    end

  end
end
# v2610051132