# frozen_string_literal: true

require File.expand_path('../../../../../test/test_helper', __FILE__)

class DocumentGenerator::FieldResolverTest < ActiveSupport::TestCase
  def setup
    # Сбрасываем кэш перед каждым тестом для чистоты эксперимента
    DocumentGenerator::FieldResolver.reset_cache!
    @original_locale = I18n.locale
  end

  def teardown
    # Восстанавливаем исходную локаль после теста
    I18n.locale = @original_locale
    DocumentGenerator::FieldResolver.reset_cache!
  end

  test 'should resolve standard field names dynamically from I18n' do
    # Проверяем, что разрешается имя, полученное динамически из текущего языкового файла
    localized_subject = I18n.t('field_subject', default: 'Subject').downcase
    resolved = DocumentGenerator::FieldResolver.resolve(localized_subject)
    assert_equal 'subject', resolved[:key]
    assert_equal :standard, resolved[:type]

    localized_status = I18n.t('field_status', default: 'Status').downcase
    resolved = DocumentGenerator::FieldResolver.resolve(localized_status)
    assert_equal 'status', resolved[:key]
  end

  test 'should resolve standard field names using English keys directly' do
    # Проверяем, что латинские ключи также работают напрямую независимо от локали
    resolved = DocumentGenerator::FieldResolver.resolve('subject')
    assert_equal 'subject', resolved[:key]
    
    resolved = DocumentGenerator::FieldResolver.resolve('assigned_to')
    assert_equal 'assigned_to', resolved[:key]
  end

  test 'should resolve parent fields recursively' do
    # Проверяем рекурсивное разрешение для полей родительской задачи
    # Используем латинский ключ для соблюдения правила отсутствия русских литералов в исполнении
    resolved = DocumentGenerator::FieldResolver.resolve('Parent.subject')
    assert_equal :parent, resolved[:type]
    assert_equal 'subject', resolved[:key]
  end

  test 'should return unknown for unrecognized fields' do
    # Проверяем корректную обработку неизвестных полей
    resolved = DocumentGenerator::FieldResolver.resolve('NonExistentFieldXYZ')
    assert_equal :unknown, resolved[:type]
    assert_equal 'NonExistentFieldXYZ', resolved[:key]
  end

  test 'should handle custom fields by name' do
    # Создаём тестовое пользовательское поле
    cf = IssueCustomField.create!(name: 'Test Custom Field', field_format: 'string', is_for_all: true)
    DocumentGenerator::FieldResolver.reset_cache! # Обновляем кэш после создания поля

    # Проверяем разрешение по полному имени
    resolved = DocumentGenerator::FieldResolver.resolve('Test Custom Field')
    assert_equal :custom, resolved[:type]
    assert_equal "cf_#{cf.id}", resolved[:key]

    # Проверяем разрешение по имени без пробелов
    resolved_no_spaces = DocumentGenerator::FieldResolver.resolve('TestCustomField')
    assert_equal :custom, resolved_no_spaces[:type]
    assert_equal "cf_#{cf.id}", resolved_no_spaces[:key]
    
    # Проверяем явное указание через префикс CF:
    resolved_cf_prefix = DocumentGenerator::FieldResolver.resolve('CF:Test Custom Field')
    assert_equal :custom, resolved_cf_prefix[:type]
    assert_equal "cf_#{cf.id}", resolved_cf_prefix[:key]
  end
end