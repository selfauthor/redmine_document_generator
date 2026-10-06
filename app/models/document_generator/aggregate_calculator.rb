# frozen_string_literal: true

module DocumentGenerator
  class AggregateCalculator
    def initialize(issues)
      @issues = issues
    end

    # Вычисляет агрегатное значение по указанному полю.
    #
    # @param func [String, Symbol] Тип агрегата: count, sum, avg, min или max.
    # @param field_name [String] Имя поля, по которому выполняется расчёт.
    # @param group_issues [Array<Issue>, nil] Необязательный набор задач текущей
    #   группы. Если не передан, используется весь набор задач.
    # @return [Integer, Float, nil] Результат агрегирования.
    def calculate(func, field_name, group_issues = nil)
      target_issues = group_issues || @issues

      case func.to_s.downcase
      when 'count'
        # count считает сами записи и поэтому не зависит от значения поля.
        return target_issues.size
      when 'sum', 'avg', 'min', 'max'
        # Для остальных агрегатов необходимо сначала получить значения поля.
        resolved = DocumentGenerator::FieldResolver.resolve(field_name)

        values = target_issues.filter_map do |issue|
          DocumentGenerator::FieldResolver.get_value(
            issue,
            resolved
          )
        end

        # Отсутствие значений не должно превращаться в искусственный ноль.
        return nil if values.empty?
      else
        # concat не является агрегатной функцией. Также здесь намеренно
        # не выполняется молчаливая обработка неизвестных функций.
        return nil
      end

      numeric_values = values.map { |value| value.to_f }

      case func.to_s.downcase
      when 'sum'
        numeric_values.sum
      when 'avg'
        numeric_values.sum / numeric_values.size.to_f
      when 'min'
        numeric_values.min
      when 'max'
        numeric_values.max
      end
    end

  end
end
# v2610061513