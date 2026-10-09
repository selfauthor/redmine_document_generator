# frozen_string_literal: true

module DocumentGenerator
  # Преобразует значение поля «Описание» в WordprocessingML.
  #
  # Класс поддерживает два режима:
  # - raw — текст выводится как есть, без обработки разметки Redmine;
  # - redmine — описание сначала преобразуется штатным механизмом Redmine
  #   в HTML, после чего HTML преобразуется в элементы WordprocessingML.
  #
  # @param issue [Issue] Задача Redmine, из которой берётся описание.
  # @param document [Nokogiri::XML::Document] XML-документ Word.
  class DescriptionFormatter
    WORD_NAMESPACE = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    XHTML_NAMESPACE = 'http://www.w3.org/1999/xhtml'

    # Инициализирует форматтер описания.
    #
    # @param issue [Issue] Задача Redmine.
    # @param document [Nokogiri::XML::Document] XML-документ Word.
    def initialize(issue, document)
      @issue = issue
      @document = document
    end

    # Возвращает HTML, сформированный штатным форматтером Redmine.
    #
    # Используется ApplicationHelper#textilizable, поэтому выбранный
    # администратором Redmine формат разметки определяется через
    # Setting.text_formatting.
    #
    # @return [String] HTML-представление описания.
    def redmine_html
      ApplicationController.helpers.textilizable(
        @issue,
        :description,
        project: @issue.project,
        only_path: true
      ).to_s
    end

    # Создаёт набор Word-узлов для значения описания.
    #
    # @param marker_run [Nokogiri::XML::Node] Исходный w:r, в котором находился маркер.
    # @param paragraph [Nokogiri::XML::Node] Исходный w:p.
    # @param mode [String] Режим форматирования: raw или redmine.
    # @return [Array<Nokogiri::XML::Node>] Узлы, которыми следует заменить маркер.
    def build_nodes(marker_run, paragraph, mode)
      return build_raw_nodes(marker_run) if mode.to_s != 'redmine'

      build_redmine_nodes(marker_run, paragraph)
    end

    private

    # Формирует текстовое представление описания без обработки Redmine-разметки.
    #
    # Переводы строк преобразуются в w:br, а не в новые абзацы.
    # Форматирование берётся из исходного w:r, в котором находился маркер.
    #
    # @param marker_run [Nokogiri::XML::Node] Исходный w:r.
    # @return [Array<Nokogiri::XML::Node>] Один или несколько w:r.
    def build_raw_nodes(marker_run)
      run_properties = marker_run.at_xpath('w:rPr', word_namespaces)
      text = @issue.description.to_s

      parts = text.split(/\r\n|\r|\n/, -1)
      runs = []

      parts.each_with_index do |part, index|
        run = create_run(run_properties)

        append_text_to_run(run, part)

        runs << run

        next if index == parts.length - 1

        break_run = create_run(run_properties)
        break_run.add_child(
          Nokogiri::XML::Node.new('w:br', @document)
        )

        runs << break_run
      end

      # Пустое описание должно удалить маркер, но сохранить форматирование
      # самого места вставки.
      if text.empty?
        [create_run(run_properties)]
      else
        runs
      end
    end

    # Формирует описание с использованием HTML, полученного от Redmine.
    #
    # @param marker_run [Nokogiri::XML::Node] Исходный w:r с маркером.
    # @param paragraph [Nokogiri::XML::Node] Исходный w:p.
    # @return [Array<Nokogiri::XML::Node>] Word-узлы для вставки.
    def build_redmine_nodes(marker_run, paragraph)
      html = redmine_html

      fragment = Nokogiri::HTML.fragment(html)
      run_properties = marker_run.at_xpath('w:rPr', word_namespaces)

      # Если описание пустое, удаляем только содержимое маркера.
      return [create_run(run_properties)] if fragment.children.empty?

      # Если HTML содержит блочные элементы, они должны находиться
      # непосредственно внутри документа/таблицы, а не внутри w:p.
      if contains_block_elements?(fragment)
        return build_block_nodes(
          fragment,
          paragraph,
          run_properties
        )
      end

      build_inline_nodes(
        fragment,
        run_properties
      )
    end

    # Проверяет наличие блочных HTML-элементов.
    #
    # @param fragment [Nokogiri::HTML::DocumentFragment] HTML-фрагмент.
    # @return [Boolean] true, если присутствуют блочные элементы.
    def contains_block_elements?(fragment)
      fragment.xpath(
        './/p | .//div | .//h1 | .//h2 | .//h3 | .//h4 | .//h5 | .//h6 | .//ul | .//ol | .//table | .//pre | .//blockquote'
      ).any?
    end

    # Формирует Word-узлы для блочного HTML.
    #
    # Каждый HTML-абзац преобразуется в отдельный w:p.
    # Таблица преобразуется в настоящий w:tbl.
    #
    # @param fragment [Nokogiri::HTML::DocumentFragment] HTML-фрагмент.
    # @param source_paragraph [Nokogiri::XML::Node] Исходный w:p.
    # @param base_run_properties [Nokogiri::XML::Node, nil] Форматирование маркера.
    # @return [Array<Nokogiri::XML::Node>] Набор блоков Word.
    def build_block_nodes(fragment, source_paragraph, base_run_properties)
      nodes = []

      fragment.children.each do |child|
        next if child.text? && child.text.strip.empty?

        case child.name.downcase
        when 'table'
          nodes << build_table(child, base_run_properties)
        when 'ul'
          nodes.concat(
            build_list(
              child,
              false,
              source_paragraph,
              base_run_properties
            )
          )
        when 'ol'
          nodes.concat(
            build_list(
              child,
              true,
              source_paragraph,
              base_run_properties
            )
          )
        when 'p', 'div', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'pre', 'blockquote'
          nodes << build_paragraph(
            child,
            source_paragraph,
            base_run_properties
          )
        else
          nodes << build_paragraph(
            child,
            source_paragraph,
            base_run_properties
          )
        end
      end

      nodes
    end

    # Формирует inline-содержимое одного HTML-фрагмента.
    #
    # @param fragment [Nokogiri::HTML::DocumentFragment] HTML-фрагмент.
    # @param base_properties [Nokogiri::XML::Node, nil] Базовое форматирование.
    # @return [Array<Nokogiri::XML::Node>] Набор w:r и w:br.
    def build_inline_nodes(fragment, base_properties)
      nodes = []

      fragment.children.each do |child|
        nodes.concat(
          build_inline_node(
            child,
            base_properties
          )
        )
      end

      nodes
    end

    # Преобразует один HTML-узел во фрагмент Word.
    #
    # @param node [Nokogiri::XML::Node] HTML-узел.
    # @param properties [Nokogiri::XML::Node, nil] Текущее форматирование.
    # @return [Array<Nokogiri::XML::Node>] Word-узлы.
    def build_inline_node(node, properties)
      if node.text?
        return build_text_runs(node.text, properties)
      end

      properties_for_node = inline_properties(
        properties,
        node
      )

      case node.name.downcase
      when 'br'
        run = create_run(properties_for_node)
        run.add_child(
          Nokogiri::XML::Node.new('w:br', @document)
        )
        [run]
      when 'strong', 'b', 'em', 'i', 'u', 's', 'del', 'span', 'code'
        node.children.flat_map do |child|
          build_inline_node(
            child,
            properties_for_node
          )
        end
      else
        node.children.flat_map do |child|
          build_inline_node(
            child,
            properties_for_node
          )
        end
      end
    end

    # Формирует Word-прогоны из обычного текста.
    #
    # @param text [String] Текст.
    # @param properties [Nokogiri::XML::Node, nil] Форматирование текста.
    # @return [Array<Nokogiri::XML::Node>] Набор w:r.
    def build_text_runs(text, properties)
      parts = text.split(/\r\n|\r|\n/, -1)
      nodes = []

      parts.each_with_index do |part, index|
        run = create_run(properties)
        append_text_to_run(run, part)
        nodes << run

        next if index == parts.length - 1

        break_run = create_run(properties)
        break_run.add_child(
          Nokogiri::XML::Node.new('w:br', @document)
        )
        nodes << break_run
      end

      nodes
    end

    # Формирует один Word-абзац.
    #
    # @param html_node [Nokogiri::XML::Node] HTML-блочный узел.
    # @param source_paragraph [Nokogiri::XML::Node] Исходный абзац шаблона.
    # @param base_properties [Nokogiri::XML::Node, nil] Форматирование маркера.
    # @return [Nokogiri::XML::Node] w:p.
    def build_paragraph(html_node, source_paragraph, base_properties)
      paragraph = Nokogiri::XML::Node.new('w:p', @document)

      source_ppr = source_paragraph.at_xpath('w:pPr', word_namespaces)
      paragraph.add_child(source_ppr.dup) if source_ppr

      heading_level =
        if html_node.name.downcase.match?(/\Ah([1-6])\z/)
          Regexp.last_match(1).to_i
        end

      if heading_level
        ppr = paragraph.at_xpath('w:pPr', word_namespaces)

        unless ppr
          ppr = Nokogiri::XML::Node.new('w:pPr', @document)
          paragraph.add_child(ppr)
        end

        p_style = Nokogiri::XML::Node.new('w:pStyle', @document)
        p_style['w:val'] = "Heading#{heading_level}"
        ppr.add_child(p_style)
      end

      html_node.children.each do |child|
        build_inline_node(
          child,
          base_properties
        ).each do |word_node|
          paragraph.add_child(word_node)
        end
      end

      paragraph
    end

    # Формирует список Word.
    #
    # Для независимости от настроек numbering.xml используется визуальный
    # маркер списка в виде обычного текста, при этом каждый пункт остаётся
    # отдельным абзацем Word.
    #
    # @param list_node [Nokogiri::XML::Node] HTML ul/ol.
    # @param ordered [Boolean] true для нумерованного списка.
    # @param source_paragraph [Nokogiri::XML::Node] Исходный абзац шаблона.
    # @param base_properties [Nokogiri::XML::Node, nil] Базовое форматирование.
    # @return [Array<Nokogiri::XML::Node>] Абзацы списка.
    def build_list(list_node, ordered, source_paragraph, base_properties)
      result = []

      list_node.xpath('./li').each_with_index do |item, index|
        paragraph = Nokogiri::XML::Node.new('w:p', @document)

        source_ppr = source_paragraph.at_xpath('w:pPr', word_namespaces)
        paragraph.add_child(source_ppr.dup) if source_ppr

        marker_run = create_run(base_properties)

        marker_text = ordered ? "#{index + 1}. " : "• "
        append_text_to_run(marker_run, marker_text)
        paragraph.add_child(marker_run)

        item.children.each do |child|
          build_inline_node(
            child,
            base_properties
          ).each do |word_node|
            paragraph.add_child(word_node)
          end
        end

        result << paragraph
      end

      result
    end

    # Формирует нативную таблицу Word из HTML table.
    #
    # @param table_node [Nokogiri::XML::Node] HTML-таблица.
    # @param base_properties [Nokogiri::XML::Node, nil] Базовое форматирование.
    # @return [Nokogiri::XML::Node] w:tbl.
    def build_table(table_node, base_properties)
      table = Nokogiri::XML::Node.new('w:tbl', @document)

      tbl_pr = Nokogiri::XML::Node.new('w:tblPr', @document)

      tbl_width = Nokogiri::XML::Node.new('w:tblW', @document)
      tbl_width['w:w'] = '0'
      tbl_width['w:type'] = 'auto'
      tbl_pr.add_child(tbl_width)

      table.add_child(tbl_pr)

      rows = table_node.xpath('./thead/tr | ./tbody/tr | ./tfoot/tr | ./tr')

      rows.each do |html_row|
        word_row = Nokogiri::XML::Node.new('w:tr', @document)

        html_row.xpath('./th | ./td').each do |html_cell|
          word_cell = Nokogiri::XML::Node.new('w:tc', @document)

          cell_properties = Nokogiri::XML::Node.new('w:tcPr', @document)

          colspan = html_cell['colspan'].to_i
          if colspan > 1
            grid_span = Nokogiri::XML::Node.new('w:gridSpan', @document)
            grid_span['w:val'] = colspan.to_s
            cell_properties.add_child(grid_span)
          end

          word_cell.add_child(cell_properties)

          paragraph = Nokogiri::XML::Node.new('w:p', @document)

          html_cell.children.each do |child|
            build_inline_node(
              child,
              base_properties
            ).each do |word_node|
              paragraph.add_child(word_node)
            end
          end

          word_cell.add_child(paragraph)
          word_row.add_child(word_cell)
        end

        table.add_child(word_row)
      end

      table
    end

    # Создаёт Word run и копирует в него свойства исходного run.
    #
    # @param properties [Nokogiri::XML::Node, nil] Свойства форматирования.
    # @return [Nokogiri::XML::Node] Новый w:r.
    def create_run(properties)
      run = Nokogiri::XML::Node.new('w:r', @document)
      run.add_child(properties.dup) if properties
      run
    end

    # Добавляет текстовый узел в Word run.
    #
    # @param run [Nokogiri::XML::Node] Word run.
    # @param text [String] Текст.
    # @return [void]
    def append_text_to_run(run, text)
      text_node = Nokogiri::XML::Node.new('w:t', @document)
      text_node.content = text
      text_node['xml:space'] = 'preserve' if text.match?(/\A\s|\s\z/)
      run.add_child(text_node)
    end

    # Формирует свойства run с учётом HTML-тега.
    #
    # Свойства исходного места вставки остаются базовыми. HTML добавляет
    # только те свойства, которые явно присутствуют в Redmine-разметке.
    #
    # @param base_properties [Nokogiri::XML::Node, nil] Базовые свойства.
    # @param node [Nokogiri::XML::Node] HTML-узел.
    # @return [Nokogiri::XML::Node, nil] Новые свойства run.
    def inline_properties(base_properties, node)
      properties = base_properties&.dup

      unless properties
        properties = Nokogiri::XML::Node.new('w:rPr', @document)
      end

      case node.name.downcase
      when 'strong', 'b'
        ensure_toggle_property(properties, 'w:b')
      when 'em', 'i'
        ensure_toggle_property(properties, 'w:i')
      when 'u'
        ensure_underline_property(properties)
      when 's', 'del'
        ensure_toggle_property(properties, 'w:strike')
      end

      if node.name.downcase == 'span'
        apply_inline_style(
          properties,
          node['style'].to_s
        )
      end

      properties
    end

    # Добавляет переключаемое свойство Word.
    #
    # @param properties [Nokogiri::XML::Node] w:rPr.
    # @param name [String] Имя XML-элемента.
    # @return [void]
    def ensure_toggle_property(properties, name)
      return if properties.at_xpath(name, word_namespaces)

      properties.add_child(
        Nokogiri::XML::Node.new(name, @document)
      )
    end

    # Добавляет свойство подчёркивания.
    #
    # @param properties [Nokogiri::XML::Node] w:rPr.
    # @return [void]
    def ensure_underline_property(properties)
      return if properties.at_xpath('w:u', word_namespaces)

      underline = Nokogiri::XML::Node.new('w:u', @document)
      underline['w:val'] = 'single'
      properties.add_child(underline)
    end

    # Применяет основные CSS-свойства inline-элемента.
    #
    # @param properties [Nokogiri::XML::Node] w:rPr.
    # @param style [String] CSS-строка.
    # @return [void]
    def apply_inline_style(properties, style)
      declarations = style.split(';').filter_map do |declaration|
        name, value = declaration.split(':', 2)
        next unless name && value

        [name.strip.downcase, value.strip.downcase]
      end.to_h

      if declarations['font-weight'] == 'bold' ||
         declarations['font-weight'].to_i >= 600
        ensure_toggle_property(properties, 'w:b')
      end

      if declarations['font-style'] == 'italic'
        ensure_toggle_property(properties, 'w:i')
      end

      if declarations['text-decoration'].to_s.include?('underline')
        ensure_underline_property(properties)
      end

      if declarations['text-decoration'].to_s.include?('line-through')
        ensure_toggle_property(properties, 'w:strike')
      end
    end

    # Возвращает пространства имён Word.
    #
    # @return [Hash] Пространства имён XML.
    def word_namespaces
      { 'w' => WORD_NAMESPACE }
    end
  end
end
# v2610081717