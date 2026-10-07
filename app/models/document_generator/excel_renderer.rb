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
    # @param shared_strings [Array<String>] Значения из sharedStrings.xml.
    # @return [void]
    def process_excel_xml(doc, context, entry_name, shared_strings)
      ns = {
        'xmlns' => 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
      }

      # Преобразуем shared strings в inlineStr, чтобы дальнейшая обработка
      # маркеров работала непосредственно с текстом ячеек.
      convert_shared_strings(doc, shared_strings, ns)

      process_excel_total_blocks(
        doc,
        context,
        ns
      )

      # При наличии групп используется отдельный механизм разворачивания
      # первого и второго уровней группировки.
      if context['groups']
        process_excel_grouped_blocks(
          doc,
          context,
          ns
        )
      else
        records = context['records'] || []

        # Обычный режим без группировки: BEGIN_ROW/END_ROW повторяется
        # для каждой основной записи.
        process_excel_row_blocks(
          doc,
          context,
          records,
          ns
        )

        # После разворачивания циклических блоков обрабатываем оставшиеся
        # статические строки, условия и обычные маркеры.
        doc.xpath('//xmlns:row', ns).each do |row|
          TemplateProcessor.process_collection_blocks(
            row,
            context.merge(records.first || {}),
            ns,
            @error_behavior
          )

          TemplateProcessor.process_conditionals_in_block(
            row,
            context.merge(records.first || {}),
            ns,
            @error_behavior
          )

          TemplateProcessor.substitute_in_block(
            row,
            context.merge(records.first || {}),
            ns,
            @error_behavior
          )
        end
      end

      # После удаления и вставки строк Excel должен получить непрерывную
      # нумерацию строк и корректные адреса ячеек.
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

    # Разворачивает Excel-шаблон с группировкой первого и второго уровней.
    #
    # @param doc [Nokogiri::XML::Document] XML-документ листа Excel.
    # @param context [Hash] Контекст с группами и агрегатами.
    # @param ns [Hash] Пространства имён XML.
    # @return [void]
    def process_excel_grouped_blocks(doc, context, ns)
      rows = doc.xpath('//xmlns:row', ns).to_a
      groups = context['groups'] || []

      # Находим обязательные управляющие маркеры всей области группировки.
      footer_index = rows.index do |row|
        excel_row_text(row, ns).match?(
          /\A\s*<%\s*GROUP_FOOTER\s*%>\s*\z/i
        )
      end

      end_groups_index = rows.index do |row|
        excel_row_text(row, ns).match?(
          /\A\s*<%\s*END_GROUPS\s*%>\s*\z/i
        )
      end

      unless footer_index
        raise DocumentGenerator::TemplateError,
              I18n.t('document_generator.error_missing_group_footer')
      end

      unless end_groups_index
        raise DocumentGenerator::TemplateError,
              I18n.t('document_generator.error_missing_end_groups')
      end

      # END_GROUPS должен находиться после GROUP_FOOTER, поскольку именно
      # он закрывает всю повторяемую область внешней группировки.
      if end_groups_index <= footer_index
        raise DocumentGenerator::TemplateError,
              I18n.t('document_generator.error_invalid_group_structure')
      end

      begin_index = rows.index do |row|
        excel_row_text(row, ns).match?(
          /\A\s*<%\s*BEGIN_ROW\s*%>\s*\z/i
        )
      end

      unless begin_index
        raise DocumentGenerator::TemplateError,
              I18n.t('document_generator.error_missing_end_row')
      end

      end_index = nil

      ((begin_index + 1)...rows.length).each do |index|
        if excel_row_text(rows[index], ns).match?(
          /\A\s*<%\s*END_ROW\s*%>\s*\z/i
        )
          end_index = index
          break
        end
      end

      unless end_index
        raise DocumentGenerator::TemplateError,
              I18n.t('document_generator.error_missing_end_row')
      end

      group_header_index = rows.index do |row|
        excel_row_text(row, ns).match?(
          /\A\s*<%\s*GROUP_HEADER\s*%>\s*\z/i
        )
      end

      unless group_header_index
        raise DocumentGenerator::TemplateError,
              I18n.t('document_generator.error_missing_group_header')
      end

      group_header_2_index = rows.index do |row|
        excel_row_text(row, ns).match?(
          /\A\s*<%\s*GROUP_HEADER_2\s*%>\s*\z/i
        )
      end

      group_footer_2_index = rows.index do |row|
        excel_row_text(row, ns).match?(
          /\A\s*<%\s*GROUP_FOOTER_2\s*%>\s*\z/i
        )
      end

      # Статический префикс находится до GROUP_HEADER.
      prefix_rows = rows[0...group_header_index].to_a.reject do |row|
        excel_group_control_row?(row, ns)
      end

      # Заголовок первого уровня заканчивается перед GROUP_HEADER_2 или
      # BEGIN_ROW. Поддерживаются оба допустимых варианта расположения.
      group_header_end = [
        group_header_2_index,
        begin_index
      ].compact.select { |index| index > group_header_index }.min

      group_header_end ||= begin_index

      group_header_rows = rows[
        (group_header_index + 1)...group_header_end
      ].to_a

      # Заголовок второго уровня заканчивается перед BEGIN_ROW.
      group_header_2_rows = []

      if group_header_2_index
        if group_header_2_index < begin_index
          group_header_2_rows = rows[
            (group_header_2_index + 1)...begin_index
          ].to_a
        else
          raise DocumentGenerator::TemplateError,
                I18n.t('document_generator.error_row_block_mismatch')
        end
      end

      row_template = rows[(begin_index + 1)...end_index].to_a

      # GROUP_FOOTER_2 начинает повторяемый подвал внутренней группы.
      # Всё до GROUP_FOOTER относится только к этому внутреннему подвалу.
      group_footer_2_rows = []

      if group_footer_2_index
        if group_footer_2_index < footer_index
          group_footer_2_rows = rows[
            (group_footer_2_index + 1)...footer_index
          ].to_a.reject do |row|
            excel_group_control_row?(row, ns)
          end
        else
          raise DocumentGenerator::TemplateError,
                I18n.t('document_generator.error_invalid_group_structure')
        end
      end

      # Содержимое между END_ROW и GROUP_FOOTER_2/GROUP_FOOTER сохраняется
      # как тело внешней группы. Оно выводится один раз для каждой внешней
      # группы после обработки всех её вложенных групп.
      outer_group_body_end = group_footer_2_index || footer_index

      group_body_rows = rows[
        (end_index + 1)...outer_group_body_end
      ].to_a.reject do |row|
        excel_group_control_row?(row, ns)
      end

      # GROUP_FOOTER начинает повторяемый подвал внешней группы.
      # Он заканчивается только на END_GROUPS.
      group_footer_rows = rows[
        (footer_index + 1)...end_groups_index
      ].to_a.reject do |row|
        excel_group_control_row?(row, ns)
      end

      # Всё после END_GROUPS является обычным продолжением документа и
      # поэтому добавляется только один раз, после разворачивания групп.
      suffix_rows = rows[
        (end_groups_index + 1)...rows.length
      ].to_a.reject do |row|
        excel_group_control_row?(row, ns)
      end

      expanded_rows = prefix_rows

      groups.each do |group|
        group_context = context.merge(
          'count' => group['count']
        )

        # Передаём агрегаты первого уровня.
        group.each do |key, value|
          if key.to_s.start_with?('group_agg_')
            group_context[key] = value
          end
        end

        # Заголовок первого уровня.
        expanded_rows.concat(
          render_excel_group_rows(
            group_header_rows,
            group_context,
            ns
          )
        )

        second_groups = group['groups_2'] || []

        if second_groups.empty?
          group['records'].to_a.each do |record|
            row_context = group_context.merge(record)

            expanded_rows.concat(
              render_excel_group_rows(
                row_template,
                row_context,
                ns
              )
            )
          end
        else
          second_groups.each do |group_2|
            group_2_context = group_context.merge(
              'count' => group_2['count']
            )

            # Передаём агрегаты второго уровня.
            group_2.each do |key, value|
              if key.to_s.start_with?('group_2_agg_')
                group_2_context[key] = value
              end
            end

            # Заголовок второго уровня.
            expanded_rows.concat(
              render_excel_group_rows(
                group_header_2_rows,
                group_2_context,
                ns
              )
            )

            # Основные записи второго уровня.
            group_2['records'].to_a.each do |record|
              row_context = group_2_context.merge(record)

              expanded_rows.concat(
                render_excel_group_rows(
                  row_template,
                  row_context,
                  ns
                )
              )
            end

            # Подвал второго уровня повторяется для каждой внутренней группы.
            unless group_footer_2_rows.empty?
              expanded_rows.concat(
                render_excel_group_rows(
                  group_footer_2_rows,
                  group_2_context,
                  ns
                )
              )
            end
          end
        end

        # Содержимое после END_ROW и до GROUP_FOOTER_2/GROUP_FOOTER
        # относится к внешней группе и повторяется один раз для неё.
        unless group_body_rows.empty?
          expanded_rows.concat(
            render_excel_group_rows(
              group_body_rows,
              group_context,
              ns
            )
          )
        end

        # Подвал первого уровня повторяется для каждой внешней группы.
        unless group_footer_rows.empty?
          expanded_rows.concat(
            render_excel_group_rows(
              group_footer_rows,
              group_context,
              ns
            )
          )
        end
      end

      rows.each(&:remove)

      root = doc.at_xpath('//xmlns:sheetData', ns)

      expanded_rows.each do |row|
        root.add_child(row)
      end

      suffix_rows.each do |row|
        root.add_child(row)
      end
    end

    # Клонирует и обрабатывает набор строк Excel с указанным контекстом.
    #
    # @param template_rows [Array<Nokogiri::XML::Node>] Шаблонные строки.
    # @param context [Hash] Контекст конкретной группы или записи.
    # @param ns [Hash] Пространства имен XML.
    # @return [Array<Nokogiri::XML::Node>] Обработанные копии строк.
    def render_excel_group_rows(template_rows, context, ns)
      clones = template_rows.map(&:dup)

      # Сначала разворачиваем вложенные коллекции текущего контекста.
      clones = TemplateProcessor.process_collection_blocks(
        clones,
        context,
        ns,
        @error_behavior
      )

      # Затем обрабатываем условия и обычные маркеры.
      clones.each do |clone|
        TemplateProcessor.process_conditionals_in_block(
          clone,
          context,
          ns,
          @error_behavior
        )

        TemplateProcessor.substitute_in_block(
          clone,
          context,
          ns,
          @error_behavior
        )
      end

      clones
    end

    # Проверяет, является ли строка Excel управляющей строкой группировки.
    #
    # @param row [Nokogiri::XML::Node] XML-узел строки Excel.
    # @param ns [Hash] Пространства имен XML.
    # @return [Boolean] true, если строка содержит только управляющую команду.
    def excel_group_control_row?(row, ns)
      text = excel_row_text(row, ns)

      text.match?(
        /\A\s*<%\s*(
          GROUP_BY(?:_2)?\s*:\s*[^%]+|
          GROUP_HEADER(?:_2)?|
          GROUP_FOOTER(?:_2)?|
          END_GROUPS
        )\s*%>\s*\z/ix
      )
    end

    # Разворачивает блок BEGIN_TOTAL/END_TOTAL один раз для всей выборки,
    # сохраняя его исходное положение относительно строк шаблона.
    #
    # @param doc [Nokogiri::XML::Document] XML-документ листа Excel.
    # @param context [Hash] Общий контекст выгрузки.
    # @param ns [Hash] Пространства имён XML.
    # @return [void]
    def process_excel_total_blocks(doc, context, ns)
      rows = doc.xpath('//xmlns:row', ns).to_a

      begin_index = rows.index do |row|
        excel_row_text(row, ns).match?(
          /\A\s*<%\s*BEGIN_TOTAL\s*%>\s*\z/i
        )
      end

      return unless begin_index

      end_index = nil

      ((begin_index + 1)...rows.length).each do |index|
        if excel_row_text(rows[index], ns).match?(
          /\A\s*<%\s*END_TOTAL\s*%>\s*\z/i
        )
          end_index = index
          break
        end
      end

      unless end_index
        raise DocumentGenerator::TemplateError,
              I18n.t('document_generator.error_missing_end_total')
      end

      total_context = context.merge(
        context['totals']&.first || {}
      )

      template_rows = rows[
        (begin_index + 1)...end_index
      ].to_a

      rendered_rows = render_excel_group_rows(
        template_rows,
        total_context,
        ns
      )

      begin_row = rows[begin_index]

      rendered_rows.reverse_each do |row|
        begin_row.add_previous_sibling(row)
      end

      rows[
        begin_index..end_index
      ].each(&:remove)
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
# v2610071414