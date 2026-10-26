# frozen_string_literal: true

module DocumentGenerator
  # ===========================================================================
  # РЕНДЕРИНГ WORD-ДОКУМЕНТОВ
  # ===========================================================================
  # Класс отвечает за генерацию Word-документов (.docx) из шаблонов.
  # Обрабатывает циклы, условия, подстановку полей с сохранением форматирования.
  class WordRenderer
    # Возвращает предупреждения, накопленные при формировании документа.
    # @return [Array<String>] Список предупреждений для пользователя.
    attr_reader :warnings

    # Инициализирует Word-рендерер.
    #
    # @param template_path [String] Путь к шаблону DOCX.
    # @param issues [Array<Issue>] Задачи для выгрузки.
    # @param parser_config [Hash] Конфигурация шаблона.
    # @param error_behavior [String] Поведение при ошибке.
    # @param description_format [String] Режим форматирования Description:
    #   raw или redmine.
    def initialize(
      template_path,
      issues,
      parser_config,
      error_behavior,
      description_format = 'raw',
      progress_callback = nil
    )
      @template_path = template_path
      @issues = issues
      @parser_config = parser_config
      @error_behavior = error_behavior

      # Нормализуем режим, чтобы некорректное значение никогда
      # не включало Redmine-форматирование случайно.
      @description_format =
        description_format.to_s == 'redmine' ? 'redmine' : 'raw'

      @parser_config = @parser_config.merge(
        description_format: @description_format
      )

      @warnings = []
      @progress_callback = progress_callback
    end

    # Генерация документа
    #
    # @return [String] Путь к сгенерированному файлу
    # @raise [RenderError] если произошла ошибка при рендеринге
    def render
      # Строим контекст данных
      context = ContextBuilder.new(
        @issues,
        @parser_config,
        @error_behavior,
        @progress_callback
      ).build

      # Используем тот же массив предупреждений, который создал ContextBuilder.
      # Благодаря общей ссылке предупреждения из обработки полей будут доступны renderer.
      @warnings = context['__warnings'] || @warnings
      # Путь для выходного файла
      output_path = "#{@template_path}.output.docx"
      # Файлы для обработки
      xml_targets = [
        'word/document.xml',
        'word/header*.xml',
        'word/footer*.xml'
      ]
      # Обрабатываем архив
      TemplateProcessor.process_archive(@template_path, output_path, xml_targets) do |doc, entry_name|
        process_word_xml(doc, context, entry_name)
      end

      # Добавляем определения стилей и нумерации, необходимые форматтеру описания.
      ensure_word_formatting_support(output_path)

      # Сообщаем о завершении формирования выходного документа.
      @progress_callback.call(:render_complete, 1.0) if @progress_callback

      output_path
    rescue StandardError => e
      Rails.logger.error "[DocumentGenerator] Word render failed: #{e.message}\n#{e.backtrace&.join("\n")}"
      error_msg = I18n.t('document_generator.error_word_render_failed', message: e.message)
      handle_error(error_msg)
    end

    private

    # Обработка XML документа Word
    #
    # @param doc [Nokogiri::XML::Document] XML-документ
    # @param context [Hash] Контекст данных
    # @param entry_name [String] Имя файла в архиве
    def process_word_xml(doc, context, entry_name)
      ns = { 'w' => 'http://schemas.openxmlformats.org/wordprocessingml/2006/main' }
      process_total_blocks(
        doc,
        context,
        ns
      )
      records = context['records'] || []

      # Обрабатываем циклы только в основном документе Word.
      # Заголовки и колонтитулы могут содержать поля, но не циклы задач.
      if entry_name == 'word/document.xml' &&
         @parser_config[:blocks][:row] &&
         records.present?
        process_row_blocks(doc, context, records, ns)
      end
      # Контекст для рендеринга (общий + первая запись)
      render_context = context.merge(context['records'].first || {})
      # Обрабатываем все абзацы и строки таблиц
      doc.xpath('//w:p | //w:tr', ns).each do |block_node|
        next if block_node.name == 'tr' && @parser_config[:blocks][:row] &&
                block_node.xpath('.//w:t', ns).map(&:text).join.include?('BEGIN_ROW')
        # 1. Обрабатываем коллекции (подзадачи, наблюдатели, связи)
        TemplateProcessor.process_collection_blocks(block_node, render_context, ns, @error_behavior)
        # 2. Обрабатываем условия
        TemplateProcessor.process_conditionals_in_block(block_node, render_context, ns, @error_behavior)
        # 3. Подставляем значения
        TemplateProcessor.substitute_in_block(block_node, render_context, ns, @error_behavior)
      end


      # ==========================================================================
      # ВОССТАНОВЛЕНИЕ СТРУКТУРЫ WORDPROCESSINGML
      # ==========================================================================
      # В теле Word-документа w:sectPr должен быть последним дочерним
      # элементом. Перемещаем существующие настройки секции в конец,
      # а не копируем их, чтобы не оставлять некорректный элемент в середине.
      body = doc.at_xpath('//w:body', ns)

      if body
        section_properties = body.xpath('./w:sectPr', ns).to_a

        # Сохраняем существующие настройки страницы, удаляя дубликаты.
        sect_pr = section_properties.first
        section_properties.drop(1).each(&:remove)

        if sect_pr
          # Перемещаем исходный элемент в конец тела документа.
          sect_pr.remove
          body.add_child(sect_pr)
        else
          # Создаём минимальный элемент секции, если шаблон его не содержит.
          sect_pr = Nokogiri::XML::Node.new(
            'sectPr',
            doc
          )
          sect_pr.namespace = body.namespace
          body.add_child(sect_pr)
        end

        # Исправляем некорректные прямые дочерние элементы w:r.
        # Word требует, чтобы текстовые прогоны находились внутри w:p.
        body.children.to_a.each do |node|
          next unless node.element? &&
                      node.name == 'r' &&
                      node.namespace&.href == ns['w']

          paragraph = Nokogiri::XML::Node.new(
            'p',
            doc
          )
          paragraph.namespace = body.namespace

          # Сохраняем содержимое исходного прогона без изменения форматирования.
          node.replace(paragraph)
          paragraph.add_child(node)
        end


        # Назначаем уникальные идентификаторы и имена закладкам,
        # которые могли дублироваться при копировании строк шаблона.
        bookmark_queues = Hash.new { |hash, key| hash[key] = [] }
        used_bookmark_names = {}
        next_bookmark_id = 0

        doc.xpath('//w:bookmarkStart | //w:bookmarkEnd', ns).each do |bookmark|
          if bookmark.name == 'bookmarkStart'
            original_id = bookmark['w:id']
            original_name = bookmark['w:name'].to_s

            # Подбираем уникальное имя, сохраняя исходное имя первой закладки.
            unique_name = original_name
            suffix = 2

            while used_bookmark_names[unique_name]
              suffix_text = "_DG#{suffix}"
              unique_name = "#{original_name[0, 40 - suffix_text.length]}#{suffix_text}"
              suffix += 1
            end

            used_bookmark_names[unique_name] = true

            # Связываем новую закладку с соответствующим закрывающим элементом.
            bookmark_queues[original_id] << next_bookmark_id
            bookmark['w:id'] = next_bookmark_id.to_s
            bookmark['w:name'] = unique_name
            next_bookmark_id += 1
          else
            # Закрывающий элемент получает ID соответствующей открывающей закладки.
            original_id = bookmark['w:id']
            matching_ids = bookmark_queues[original_id]

            # Для вложенных закладок закрываем последнюю открытую закладку
            # с таким исходным идентификатором.
            bookmark['w:id'] = matching_ids.pop.to_s if matching_ids.present?
          end
        end

        # Перемещение w:sectPr в конец выполняется после всех исправлений,
        # чтобы он гарантированно оставался последним элементом w:body.
        sect_pr.remove
        body.add_child(sect_pr)
      end

    end

    # Разворачивает блок BEGIN_TOTAL/END_TOTAL один раз для всей выборки.
    #
    # @param doc [Nokogiri::XML::Document] XML-документ Word.
    # @param context [Hash] Общий контекст выгрузки.
    # @param ns [Hash] Пространства имён Word.
    # @return [void]
    def process_total_blocks(doc, context, ns)
      all_nodes = doc.xpath(
        '//w:tr | //w:p[not(ancestor::w:tr)]',
        ns
      ).to_a

      begin_index = all_nodes.index do |node|
        node.xpath('.//w:t', ns).map(&:text).join.match?(
          /\A\s*<%\s*BEGIN_TOTAL\s*%>\s*\z/i
        )
      end

      return unless begin_index

      end_index = nil

      ((begin_index + 1)...all_nodes.length).each do |index|
        if all_nodes[index].xpath('.//w:t', ns).map(&:text).join.match?(
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

      template_nodes = all_nodes[
        (begin_index + 1)...end_index
      ].to_a

      parent = all_nodes[begin_index].parent

      template_nodes.reverse_each do |node|
        clone = node.dup

        TemplateProcessor.process_collection_blocks(
          clone,
          total_context,
          ns,
          @error_behavior
        )

        TemplateProcessor.process_conditionals_in_block(
          clone,
          total_context,
          ns,
          @error_behavior
        )

        TemplateProcessor.substitute_in_block(
          clone,
          total_context,
          ns,
          @error_behavior
        )

        all_nodes[begin_index].add_previous_sibling(clone)
      end

      all_nodes[
        begin_index..end_index
      ].each(&:remove)
    end

    # Обработка блоков с циклами (BEGIN_ROW/END_ROW)
    #
    # @param doc [Nokogiri::XML::Document] XML-документ
    # @param context [Hash] Общий контекст данных
    # @param records [Array<Hash>] Массив записей для цикла
    # @param ns [Hash] Пространства имен XML
    def process_row_blocks(doc, context, records, ns)
      # Используем цикл для обработки всех блоков BEGIN_ROW/END_ROW в документе,
      # аналогично тому, как это реализовано в Excel-рендерере.
      loop do
        # Находим все узлы (строки таблиц или абзацы)
        # Используем to_a, так как коллекция будет изменяться в процессе итерации
        all_nodes = doc.xpath('//w:tr | //w:p[not(ancestor::w:tr)]', ns).to_a
        
        start_node = nil
        end_node = nil
        template_nodes = []
        in_block = false
        
        # Ищем блок BEGIN_ROW/END_ROW
        all_nodes.each do |node|
          text = node.xpath('.//w:t', ns).map(&:text).join
          
          # Используем регулярное выражение с учетом возможных пробелов, 
          # как в Excel-рендерере, для большей надежности
          if !in_block && text.match?(/<%\s*BEGIN_ROW\s*%>/i)
            # Начало блока
            in_block = true
            start_node = node
            
            # Управляющий узел не добавляем в template_nodes, чтобы он не клонировался
            if text.match?(/<%\s*END_ROW\s*%>/i)
              # BEGIN_ROW и END_ROW находятся в одном узле
              end_node = node
              in_block = false
              break
            end
          elsif in_block
            if text.match?(/<%\s*END_ROW\s*%>/i)
              end_node = node
              in_block = false
              break
            else
              # В шаблон для клонирования добавляем только узлы, 
              # находящиеся строго между BEGIN_ROW и END_ROW
              template_nodes << node
            end
          end
        end
        
        # Если блок не найден, завершаем цикл обработки
        break unless start_node && end_node
        
        if start_node.parent != end_node.parent
          error_msg = I18n.t('document_generator.error_row_block_mismatch')
          raise DocumentGenerator::TemplateError, error_msg
        end
        
        parent = start_node.parent
        
        # Удаляем управляющие узлы (BEGIN_ROW и END_ROW) из документа, 
        # чтобы они не оставались в виде пустых строк после подстановки значений
        start_node.remove
        end_node.remove unless end_node == start_node
        
        # Удаляем шаблонные узлы из документа (они будут добавлены как клоны)
        template_nodes.each(&:remove)
        
        # НЕ вызываем clean_node_text здесь! Маркеры ELSE и END нужны для process_conditionals_in_block
        
        # Для каждой записи создаем клон только тела цикла (без управляющих маркеров)
        records.each_with_index do |record, record_idx|
          begin
            merged_context = context.merge(record)

            # Клонируем исходные XML-узлы для текущей записи.
            clones = template_nodes.map(&:dup)

            # Разворачиваем вложенные коллекции в контексте этой задачи.
            clones = TemplateProcessor.process_collection_blocks(
              clones, merged_context, ns, @error_behavior
            )

            # Обрабатываем каждый клон во временном XML-контейнере.
            # Это необходимо, поскольку форматирование описания может
            # заменять исходный абзац несколькими соседними Word-блоками.
            clones.each do |clone|
              # Создаём временный контейнер в том же XML-документе,
              # чтобы у корневого узла клона был родитель.
              temporary_parent = Nokogiri::XML::DocumentFragment.new(
                clone.document
              )
              temporary_parent.add_child(clone)

              # Обрабатываем условия и поля уже присоединённого клона.
              TemplateProcessor.process_conditionals_in_block(
                clone, merged_context, ns, @error_behavior
              )

              TemplateProcessor.substitute_in_block(
                clone, merged_context, ns, @error_behavior
              )

              # Переносим все получившиеся узлы в настоящий документ.
              # После форматирования их может быть несколько вместо одного.
              temporary_parent.children.to_a.each do |result_node|
                parent.add_child(result_node)
              end
            end
          rescue DocumentGenerator::SkipRecordError => e
            # Записываем информацию о пропущенной записи на английском языке.
            issue = record['__issue']

            Rails.logger.warn(
              "[DocumentGenerator] Issue ##{issue&.id || 'unknown'} was skipped because a required template field was missing."
            )

            # Переходим к следующей задаче, не прерывая формирование документа.
            next
          end
        end
      end
    end

    # Обеспечивает наличие стилей заголовков и определений нумерации в DOCX.
    #
    # @param output_path [String] Путь к сформированному DOCX.
    # @return [void]
    def ensure_word_formatting_support(output_path)
      require 'zip'

      # Открываем готовый DOCX и обновляем необходимые части пакета.
      Zip::File.open(output_path) do |zip_file|
        styles_xml = read_zip_entry(zip_file, 'word/styles.xml')
        numbering_xml = read_zip_entry(zip_file, 'word/numbering.xml')
        relationships_xml = read_zip_entry(zip_file, 'word/_rels/document.xml.rels')
        content_types_xml = read_zip_entry(zip_file, '[Content_Types].xml')

        # Стили заголовков должны существовать даже в минимальном шаблоне.
        styles_doc = Nokogiri::XML(styles_xml)
        ensure_heading_styles(styles_doc)

        # Создаём определения маркированного и нумерованного списков.
        numbering_doc =
          if numbering_xml
            Nokogiri::XML(numbering_xml)
          else
            Nokogiri::XML(
              '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'
            )
          end

        # Читаем основной документ, чтобы заменить временные numId
        # только у сформированных списков описания задачи.
        document_xml = read_zip_entry(zip_file, 'word/document.xml')
        document_doc = Nokogiri::XML(document_xml)

        ensure_list_numbering(numbering_doc, document_doc)

        # Регистрируем numbering.xml в отношениях документа и типах содержимого.
        relationships_doc = Nokogiri::XML(relationships_xml)
        content_types_doc = Nokogiri::XML(content_types_xml)

        ensure_numbering_relationship(relationships_doc)
        ensure_numbering_content_type(content_types_doc)

        # Записываем обновлённые XML-части обратно в DOCX.
        replace_zip_entry(zip_file, 'word/styles.xml', styles_doc.to_xml)
        replace_zip_entry(zip_file, 'word/numbering.xml', numbering_doc.to_xml)
        # Сохраняем основной документ с новыми идентификаторами списков.
        replace_zip_entry(zip_file, 'word/document.xml', document_doc.to_xml)
        replace_zip_entry(
          zip_file,
          'word/_rels/document.xml.rels',
          relationships_doc.to_xml
        )
        replace_zip_entry(
          zip_file,
          '[Content_Types].xml',
          content_types_doc.to_xml
        )
      end
    end

    # Читает содержимое XML-части ZIP-архива.
    #
    # @param zip_file [Zip::File] Открытый DOCX.
    # @param entry_name [String] Путь к части внутри архива.
    # @return [String, nil] Содержимое файла или nil, если он отсутствует.
    def read_zip_entry(zip_file, entry_name)
      entry = zip_file.find_entry(entry_name)
      entry ? zip_file.read(entry_name) : nil
    end

    # Заменяет XML-часть DOCX или добавляет её, если ранее её не было.
    #
    # @param zip_file [Zip::File] Открытый DOCX.
    # @param entry_name [String] Путь к части внутри архива.
    # @param content [String] Новое содержимое файла.
    # @return [void]
    def replace_zip_entry(zip_file, entry_name, content)
      existing_entry = zip_file.find_entry(entry_name)
      zip_file.remove(existing_entry) if existing_entry

      zip_file.get_output_stream(entry_name) do |stream|
        stream.write(content)
      end
    end

    # Добавляет стандартные стили заголовков, если шаблон их не содержит.
    #
    # @param styles_doc [Nokogiri::XML::Document] Документ word/styles.xml.
    # @return [void]
    def ensure_heading_styles(styles_doc)
      root = styles_doc.root
      namespace = root.namespace

      # Идентификатор a в некоторых минимальных шаблонах соответствует Normal.
      normal_style = root.at_xpath(
        './w:style[@w:styleId="a"] | ./w:style[@w:styleId="Normal"]',
        'w' => namespace.href
      )

      %w[1 2 3].each do |level|
        style_id = "Heading#{level}"

        # Не заменяем уже существующий стиль шаблона.
        next if root.at_xpath(
          "./w:style[@w:styleId='#{style_id}']",
          'w' => namespace.href
        )

        style = Nokogiri::XML::Node.new('style', styles_doc)
        style.namespace = namespace
        style['w:type'] = 'paragraph'
        style['w:styleId'] = style_id

        name = Nokogiri::XML::Node.new('name', styles_doc)
        name.namespace = namespace
        name['w:val'] = "heading #{level}"
        style.add_child(name)

        if normal_style
          based_on = Nokogiri::XML::Node.new('basedOn', styles_doc)
          based_on.namespace = namespace
          based_on['w:val'] = normal_style['w:styleId']
          style.add_child(based_on)
        end

        # Делаем стиль доступным в галерее стилей Word.
        qformat = Nokogiri::XML::Node.new('qFormat', styles_doc)
        qformat.namespace = namespace
        style.add_child(qformat)

        ppr = Nokogiri::XML::Node.new('pPr', styles_doc)
        ppr.namespace = namespace

        outline = Nokogiri::XML::Node.new('outlineLvl', styles_doc)
        outline.namespace = namespace
        outline['w:val'] = (level.to_i - 1).to_s
        ppr.add_child(outline)
        style.add_child(ppr)

        # Задаём базовое оформление, если шаблон не предоставил свой стиль.
        rpr = Nokogiri::XML::Node.new('rPr', styles_doc)
        rpr.namespace = namespace

        bold = Nokogiri::XML::Node.new('b', styles_doc)
        bold.namespace = namespace
        rpr.add_child(bold)

        size = Nokogiri::XML::Node.new('sz', styles_doc)
        size.namespace = namespace
        size['w:val'] = { '1' => '32', '2' => '28', '3' => '24' }.fetch(level)
        rpr.add_child(size)

        style.add_child(rpr)
        root.add_child(style)
      end
    end

    # Добавляет уникальные определения списков и заменяет временные numId
    # в основном документе, не изменяя существующие определения шаблона.
    #
    # @param numbering_doc [Nokogiri::XML::Document] Документ word/numbering.xml.
    # @param document_doc [Nokogiri::XML::Document] Основной документ word/document.xml.
    # @return [void]
    def ensure_list_numbering(numbering_doc, document_doc)
      root = numbering_doc.root
      namespace = root.namespace
      ns = { 'w' => namespace.href }

      # Находим временные numId списков, созданных DescriptionFormatter.
      # Идентификаторы начинаются с 2100000000; чётные соответствуют
      # маркированным спискам, нечётные — нумерованным.
      document_ns = {
        'w' => 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
      }

      temporary_ids = document_doc.xpath(
        '//w:pPr/w:numPr/w:numId',
        document_ns
      ).map do |num_id_node|
        num_id_node['w:val']
      end.select do |value|
        value.to_i >= 2_100_000_000
      end.uniq.to_h do |temporary_id|
        definition = if temporary_id.to_i.odd?
                       { format: 'decimal', text: '%1.' }
                     else
                       { format: 'bullet', text: '•' }
                     end

        [temporary_id, definition]
      end

      # Проверяем, что временные идентификаторы не заняты шаблоном.
      # При совпадении останавливаем обработку, чтобы не менять его списки.
      temporary_ids.each_key do |temporary_id|
        existing_num = root.at_xpath(
          "./w:num[@w:numId='#{temporary_id}']",
          ns
        )

        if existing_num
          raise DocumentGenerator::RenderError,
                'Temporary Word list identifier conflicts with the template.'
        end
      end

      # Находим свободные идентификаторы выше максимальных значений шаблона.
      used_abstract_ids = root.xpath('./w:abstractNum/@w:abstractNumId', ns)
                              .map { |attribute| attribute.value.to_i }
      used_num_ids = root.xpath('./w:num/@w:numId', ns)
                         .map { |attribute| attribute.value.to_i }

      next_abstract_id = (used_abstract_ids.max || 0) + 1
      next_num_id = (used_num_ids.max || 0) + 1

      # Подготавливаем определения списков и соответствие временных ID новым.
      id_mapping = {}

      temporary_ids.each do |temporary_id, definition|
        abstract_id = next_abstract_id
        num_id = next_num_id

        next_abstract_id += 1
        next_num_id += 1

        # Создаём абстрактное определение списка Word.
        abstract_num = Nokogiri::XML::Node.new('abstractNum', numbering_doc)
        abstract_num.namespace = namespace
        abstract_num['w:abstractNumId'] = abstract_id.to_s

        multi_level = Nokogiri::XML::Node.new('multiLevelType', numbering_doc)
        multi_level.namespace = namespace
        multi_level['w:val'] = 'singleLevel'
        abstract_num.add_child(multi_level)

        level = Nokogiri::XML::Node.new('lvl', numbering_doc)
        level.namespace = namespace
        level['w:ilvl'] = '0'

        start = Nokogiri::XML::Node.new('start', numbering_doc)
        start.namespace = namespace
        start['w:val'] = '1'
        level.add_child(start)

        num_format = Nokogiri::XML::Node.new('numFmt', numbering_doc)
        num_format.namespace = namespace
        num_format['w:val'] = definition[:format]
        level.add_child(num_format)

        level_text = Nokogiri::XML::Node.new('lvlText', numbering_doc)
        level_text.namespace = namespace
        level_text['w:val'] = definition[:text]
        level.add_child(level_text)

        justification = Nokogiri::XML::Node.new('lvlJc', numbering_doc)
        justification.namespace = namespace
        justification['w:val'] = 'left'
        level.add_child(justification)

        # Задаём отступ и позицию маркера, сохраняя отдельное форматирование абзаца.
        ppr = Nokogiri::XML::Node.new('pPr', numbering_doc)
        ppr.namespace = namespace

        tabs = Nokogiri::XML::Node.new('tabs', numbering_doc)
        tabs.namespace = namespace

        tab = Nokogiri::XML::Node.new('tab', numbering_doc)
        tab.namespace = namespace
        tab['w:val'] = 'num'
        tab['w:pos'] = '720'
        tabs.add_child(tab)
        ppr.add_child(tabs)

        indent = Nokogiri::XML::Node.new('ind', numbering_doc)
        indent.namespace = namespace
        indent['w:left'] = '720'
        indent['w:hanging'] = '360'
        ppr.add_child(indent)

        level.add_child(ppr)
        abstract_num.add_child(level)

        # По структуре Word все abstractNum должны идти до элементов num.
        first_num = root.at_xpath('./w:num', ns)
        if first_num
          first_num.add_previous_sibling(abstract_num)
        else
          root.add_child(abstract_num)
        end

        # Создаём экземпляр списка, связывающий numId с abstractNum.
        num = Nokogiri::XML::Node.new('num', numbering_doc)
        num.namespace = namespace
        num['w:numId'] = num_id.to_s

        abstract_num_id = Nokogiri::XML::Node.new('abstractNumId', numbering_doc)
        abstract_num_id.namespace = namespace
        abstract_num_id['w:val'] = abstract_id.to_s
        num.add_child(abstract_num_id)

        root.add_child(num)

        # Запоминаем новый numId для замены временного идентификатора в документе.
        id_mapping[temporary_id] = num_id.to_s
      end

      document_doc.xpath('//w:pPr/w:numPr/w:numId', document_ns).each do |num_id_node|
        replacement_id = id_mapping[num_id_node['w:val']]
        num_id_node['w:val'] = replacement_id if replacement_id
      end
    end

    # Регистрирует numbering.xml в отношениях основного документа Word.
    #
    # @param relationships_doc [Nokogiri::XML::Document] document.xml.rels.
    # @return [void]
    def ensure_numbering_relationship(relationships_doc)
      root = relationships_doc.root
      namespace = root.namespace

      # Проверяем, зарегистрирована ли уже связь с частью нумерации Word.
      exists = root.element_children.any? do |node|
        node['Type'].to_s.end_with?('/numbering')
      end

      return if exists

      relationship = Nokogiri::XML::Node.new('Relationship', relationships_doc)
      relationship.namespace = namespace
      relationship['Id'] = 'rIdDGNumbering'
      relationship['Type'] =
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering'
      relationship['Target'] = 'numbering.xml'

      root.add_child(relationship)
    end

    # Регистрирует numbering.xml в типах содержимого DOCX.
    #
    # @param content_types_doc [Nokogiri::XML::Document] [Content_Types].xml.
    # @return [void]
    def ensure_numbering_content_type(content_types_doc)
      root = content_types_doc.root
      namespace = root.namespace

      exists = root.element_children.any? do |node|
        node['PartName'] == '/word/numbering.xml'
      end
      return if exists

      override = Nokogiri::XML::Node.new('Override', content_types_doc)
      override.namespace = namespace
      override['PartName'] = '/word/numbering.xml'
      override['ContentType'] =
        'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml'

      root.add_child(override)
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
# v2610261408