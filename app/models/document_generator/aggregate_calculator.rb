# frozen_string_literal: true

module DocumentGenerator
  class AggregateCalculator
    def initialize(issues)
      @issues = issues
    end

    def calculate(func, field_name, group_issues = nil)
      target_issues = group_issues || @issues
      return 0 if target_issues.empty?

      resolved = DocumentGenerator::FieldResolver.resolve(field_name)
      values = target_issues.map { |issue| DocumentGenerator::FieldResolver.get_value(issue, resolved) }.compact

      case func.to_s.downcase
      when 'count'
        target_issues.size
      when 'sum'
        values.map { |v| v.to_f }.sum
      when 'avg'
        values.map { |v| v.to_f }.sum / values.size.to_f
      when 'min'
        values.map { |v| v.to_f }.min
      when 'max'
        values.map { |v| v.to_f }.max
      when 'concat'
        values.join(', ')
      else
        nil
      end
    end
  end
end