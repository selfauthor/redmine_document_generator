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


// Показывает сообщение экспорта в стандартной области содержимого Redmine.
// type — тип сообщения: 'error' или 'warning'.
// message — текст сообщения либо массив текстовых сообщений.
DG.showExportMessage = function(type, message) {
  // Удаляем предыдущее сообщение экспорта, чтобы уведомления не накапливались.
  $('.dg-export-response-message').remove();

  var messages = Array.isArray(message) ? message : [message];
  var $container = $('<div>')
    .addClass('flash')
    .addClass(type === 'warning' ? 'warning' : 'error')
    .addClass('dg-export-response-message');

  // Используем text(), чтобы текст сообщения не интерпретировался как HTML.
  messages.forEach(function(item) {
    $('<div>').text(item).appendTo($container);
  });

  // Размещаем уведомление в начале рабочей области Redmine — перед заголовком страницы.
  var $content = $('#content');

  if ($content.length) {
    $content.prepend($container);
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
      // Перемещаем ссылку только при её наличии, чтобы не прерывать инициализацию страницы.
      var exportLink = linkContainer.querySelector('a');
      if (exportLink) {
        span.appendChild(exportLink);
        otherFormats.appendChild(span);
        linkContainer.remove();
      }
    }
  }


  // Отправляет форму экспорта и обрабатывает файл, ошибку и предупреждения через единый механизм.
  $(document).on('submit', '#document-generator-form', async function(event) {
    event.preventDefault();

    var form = this;
    var submitButton = form.querySelector('#dg-submit-btn');

    // Блокируем повторную отправку формы на время генерации.
    if (submitButton) {
      submitButton.disabled = true;
    }

    try {
      // Отправляем форму как multipart/form-data, включая загруженный шаблон.
      var response = await fetch(form.action, {
        method: 'POST',
        body: new FormData(form),
        credentials: 'same-origin',
        headers: {
          'Accept': 'application/json, application/octet-stream',
          'X-Requested-With': 'XMLHttpRequest'
        }
      });

      // При ошибке сервер возвращает JSON с полем message.
      if (!response.ok) {
        var errorResult;

        try {
          errorResult = await response.json();
        } catch (parseError) {
          errorResult = {};
        }

        var errorMessage = errorResult.message || errorResult.error || 'Document generation failed.';

        // Закрываем модальное окно и показываем ошибку на красном фоне.
        hideModal(form);
        DG.showExportMessage('error', errorMessage);
        return;
      }

      // Получаем файл из тела успешного ответа.
      var blob = await response.blob();

      // Получаем имя файла из заголовка Content-Disposition.
      var contentDisposition = response.headers.get('Content-Disposition') || '';
      var filename = 'document';

      // Сначала ищем UTF-8-имя: оно может идти после обычного filename
      // и содержит исходные символы, включая кириллицу.
      var utf8FilenameMatch = contentDisposition.match(
        /filename\*\s*=\s*(?:UTF-8'')?([^;]+)/i
      );

      if (utf8FilenameMatch) {
        filename = decodeURIComponent(
          utf8FilenameMatch[1].trim().replace(/^"|"$/g, '')
        );
      } else {
        // Используем обычное имя только при отсутствии UTF-8-варианта.
        var regularFilenameMatch = contentDisposition.match(
          /filename\s*=\s*"?([^";]+)"?/i
        );

        if (regularFilenameMatch) {
          filename = regularFilenameMatch[1].trim().replace(/^"|"$/g, '');
        }
      }

      // Создаём временную ссылку и запускаем скачивание сформированного документа.
      var downloadUrl = window.URL.createObjectURL(blob);
      var link = document.createElement('a');

      link.href = downloadUrl;
      link.download = filename;
      document.body.appendChild(link);
      link.click();
      link.remove();

      // Освобождаем временный URL после запуска скачивания.
      window.URL.revokeObjectURL(downloadUrl);

      // Читаем предупреждения из заголовка ответа.
      var warningsHeader = response.headers.get('X-DG-Warnings');

      if (warningsHeader) {
        try {
          // Заголовок URL-кодирован, поэтому сначала декодируем его, затем разбираем JSON.
          var warnings = JSON.parse(decodeURIComponent(warningsHeader.replace(/\+/g, ' ')));

          if (Array.isArray(warnings) && warnings.length > 0) {
            DG.showExportMessage('warning', warnings);
          }
        } catch (parseError) {
          console.error('Failed to parse document generator warnings:', parseError);
        }
      }

      // Закрываем модальное окно после успешной обработки ответа.
      hideModal(form);
    } catch (error) {
      // Показываем сетевые ошибки и ошибки обработки ответа тем же способом.
      hideModal(form);
      DG.showExportMessage('error', error.message || 'Document generation failed.');
    } finally {
      // Разблокируем кнопку, если форма и кнопка ещё существуют.
      if (submitButton && document.body.contains(submitButton)) {
        submitButton.disabled = false;
      }
    }
  });

  $(document).on('input', '#dg_file_name', function() {
    DG.updateInfo();
  });

  $(document).on('change', 'input[name="export_mode"]', function() {
    DG.updateInfo();
  });

  $(document).on('change', '#template_file', function() {
    DG.updateInfo();
  });

  //v2609301241
});