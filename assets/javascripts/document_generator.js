var DG = DG || {};

// Проверка имени файла на недопустимые символы: / \ : * ? " < > |
DG.validateFilename = function(filename) {
  var invalidChars = /[\/:*?"<>|]/;
  return !invalidChars.test(filename) && filename.trim().length > 0;
};

// Проверка расширения файла шаблона (только .docx и .xlsx)
// Возвращает true, если файл не выбран (считаем валидным состоянием)
DG.validateTemplateFile = function(fileInput) {
  if (!fileInput || !fileInput.files || fileInput.files.length === 0) {
    return true; // Файл ещё не выбран — не ошибка
  }
  var filename = fileInput.files[0].name.toLowerCase();
  var validExtensions = ['.docx', '.xlsx'];
  return validExtensions.some(function(ext) {
    return filename.endsWith(ext);
  });
};

// Обновление информации в модальном окне
DG.updateInfo = function() {
  var recordCountEl = $('#dg-records-count');
  if (!recordCountEl.length) return; // Форма ещё не в DOM

  var recordCount = parseInt(recordCountEl.data('count'), 10) || 0;
  var exportMode = $('input[name="export_mode"]:checked').val();
  var fileName = $('#dg_file_name').val();
  var fileInput = document.getElementById('template_file');

  // Валидация имени файла
  var isFilenameValid = DG.validateFilename(fileName);
  
  // Валидация формата файла шаблона
  var isTemplateValid = DG.validateTemplateFile(fileInput);
  
  // Общая валидность (оба условия должны быть истинны)
  var isValid = isFilenameValid && isTemplateValid;
  
  // Обновление информации о количестве файлов
  var $filesCount = $('#dg-files-count');
  if (exportMode === 'single') {
    $filesCount.text($filesCount.data('text-multiple'));
  } else {
    $filesCount.text($filesCount.data('text-single'));
  }
  
  // Блокируем/разблокируем кнопку выгрузки
  $('#dg-submit-btn').prop('disabled', !isValid);
  
  // Показываем/скрываем ошибку имени файла
  if (fileName.length > 0 && !isFilenameValid) {
    $('#dg-filename-error').show();
  } else {
    $('#dg-filename-error').hide();
  }

  // Показываем/скрываем ошибку формата файла
  if (fileInput && fileInput.files.length > 0 && !isTemplateValid) {
    $('#dg-template-error').show();
  } else {
    $('#dg-template-error').hide();
  }
};

document.addEventListener('DOMContentLoaded', function() {
  // Перемещение ссылки в блок экспорта
  var linkContainer = document.getElementById('document-generator-export-link');
  if (linkContainer) {
    var otherFormats = document.querySelector('p.other-formats');
    if (otherFormats) {
      var span = document.createElement('span');
      span.className = 'dg-export-link';
      span.appendChild(linkContainer.querySelector('a'));
      otherFormats.appendChild(span);
      linkContainer.remove();
    }
  }

  // Делегирование событий для модального окна (создаётся динамически)
  $(document).on('input', '#dg_file_name', function() {
    DG.updateInfo();
  });

  $(document).on('change', 'input[name="export_mode"]', function() {
    DG.updateInfo();
  });

  $(document).on('change', '#template_file', function() {
    DG.updateInfo();
  });

  //v2609101339
});