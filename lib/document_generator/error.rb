# frozen_string_literal: true

module DocumentGenerator
  # Базовый класс для всех специфичных ошибок плагина
  class Error < StandardError; end
  
  # Ошибка, возникающая при некорректном формате, повреждении или синтаксических ошибках в шаблоне
  class TemplateError < Error; end
  
  # Ошибка, возникающая в процессе непосредственной генерации (рендеринга) документа
  class RenderError < Error; end
  
  # Специальное исключение для пропуска проблемной записи без прерывания всей выгрузки
  class SkipRecordError < Error; end
end