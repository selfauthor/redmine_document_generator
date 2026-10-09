var DG = DG || {};

// Проверка имени файла на недопустимые символы: / \ : * ? " < > |
DG.validateFilename = function(filename) {
  var invalidChars = /[\\/:*?"<>|]/;
  return !invalidChars.test(filename) && filename.trim().length > 0;
};

// Проверка расширения файла шаблона (только .docx и .xlsx).
// Возвращает true, если выбран допустимый файл.
DG.validateTemplateFile = function(fileInput) {
  if (!fileInput || !fileInput.files || fileInput.files.length === 0) {
    return false;
  }

  var filename = fileInput.files[0].name.toLowerCase();
  var validExtensions = ['.docx', '.xlsx'];

  return validExtensions.some(function(ext) {
    return filename.endsWith(ext);
  });
};

// Возвращает расширение выбранного шаблона.
// Если шаблон не выбран, возвращает пустую строку.
DG.getTemplateExtension = function(fileInput) {
  if (!fileInput || !fileInput.files || fileInput.files.length === 0) {
    return '';
  }

  var filename = fileInput.files[0].name.toLowerCase();
  var lastDot = filename.lastIndexOf('.');

  if (lastDot === -1) {
    return '';
  }

  return filename.substring(lastDot);
};

// Обновляет доступность вариантов форматирования поля «Описание».
// Для Excel разрешён только режим «Как есть», поскольку форматирование
// Redmine на данном этапе реализуется только для Word.
DG.updateDescriptionFormatAvailability = function() {
  var fileInput = document.getElementById('template_file');
  var extension = DG.getTemplateExtension(fileInput);

  var $raw = $('#dg-description-format-raw');
  var $redmine = $('#dg-description-format-redmine');
  var $redmineLabel = $('#dg-description-format-redmine-label');
  var $disabledMessage = $('#dg-description-format-disabled');

  // До выбора корректного шаблона оставляем оба варианта доступными.
  if (extension === '.docx') {
    $redmine.prop('disabled', false);
    $redmineLabel.removeClass('disabled');
    $disabledMessage.hide();

    return;
  }

  // Для Excel режим форматирования Redmine недоступен.
  if (extension === '.xlsx') {
    $raw.prop('checked', true);
    $redmine.prop('checked', false);
    $redmine.prop('disabled', true);
    $redmineLabel.addClass('disabled');
    $disabledMessage.show();

    return;
  }

  // Если файл ещё не выбран или имеет недопустимое расширение,
  // возвращаем элемент в исходное состояние.
  $redmine.prop('disabled', false);
  $redmineLabel.removeClass('disabled');
  $disabledMessage.hide();
};


// Обновление информации в модальном окне.
DG.updateInfo = function() {
  var recordCountEl = $('#dg-records-count');

  if (!recordCountEl.length) {
    return;
  }

  var exportMode = $('input[name="export_mode"]:checked').val();
  var fileName = $('#dg_file_name').val();
  var fileInput = document.getElementById('template_file');

  // Сначала синхронизируем доступность настройки форматирования
  // с выбранным расширением шаблона.
  DG.updateDescriptionFormatAvailability();

  // Валидация имени файла.
  var isFilenameValid = DG.validateFilename(fileName);

  // Валидация формата файла шаблона.
  var isTemplateValid = DG.validateTemplateFile(fileInput);

  // Общая валидность формы.
  var isValid = isFilenameValid && isTemplateValid;

  // Обновление информации о количестве формируемых файлов.
  var $filesCount = $('#dg-files-count');

  if (exportMode === 'single') {
    $filesCount.text($filesCount.data('text-multiple'));
  } else {
    $filesCount.text($filesCount.data('text-single'));
  }

  // Блокируем/разблокируем кнопку выгрузки.
  $('#dg-submit-btn').prop('disabled', !isValid);

  // Показываем/скрываем ошибку имени файла.
  if (fileName.length > 0 && !isFilenameValid) {
    $('#dg-filename-error').show();
  } else {
    $('#dg-filename-error').hide();
  }

  // Показываем/скрываем ошибку формата файла.
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

  // Размещаем уведомление в начале рабочей области Redmine.
  var $content = $('#content');

  if ($content.length) {
    $content.prepend($container);
  }
};


document.addEventListener('DOMContentLoaded', function() {
  // Перемещение ссылки в блок экспорта.
  var linkContainer = document.getElementById('document-generator-export-link');

  if (linkContainer) {
    var otherFormats = document.querySelector('p.other-formats');

    if (otherFormats) {
      var span = document.createElement('span');
      span.className = 'dg-export-link';

      // Перемещаем ссылку только при её наличии,
      // чтобы не прерывать инициализацию страницы.
      var exportLink = linkContainer.querySelector('a');

      if (exportLink) {
        span.appendChild(exportLink);
        otherFormats.appendChild(span);
        linkContainer.remove();
      }
    }
  }


  // Отправляет форму генерации и отображает ссылку на защищённое скачивание.
  $(document).on('submit', '#document-generator-form', async function(event) {
    event.preventDefault();

    var form = this;
    var submitButton = form.querySelector('#dg-submit-btn');
    var $downloadContainer = $('#dg-download-link-container');
    var $downloadLink = $('#dg-download-link');

    // Блокируем повторную отправку формы на время генерации.
    if (submitButton) {
      submitButton.disabled = true;
    }

    // Скрываем ссылку от предыдущего результата.
    $downloadContainer.hide();
    $downloadLink.attr('href', '#');

    try {
      // Отправляем форму как multipart/form-data,
      // включая загруженный шаблон и выбранный режим форматирования.
      var response = await fetch(form.action, {
        method: 'POST',
        body: new FormData(form),
        credentials: 'same-origin',
        headers: {
          'Accept': 'application/json',
          'X-Requested-With': 'XMLHttpRequest'
        }
      });

      var result;

      try {
        result = await response.json();
      } catch (parseError) {
        result = {};
      }

      // Обрабатываем ошибки генерации, не закрывая модальное окно.
      if (!response.ok || result.type !== 'success' || !result.download_url) {
        var errorMessage = result.message || result.error || 'Document generation failed.';
        DG.showExportMessage('error', errorMessage);
        return;
      }

      // Устанавливаем полученную от сервера ссылку для ручного запуска скачивания.
      $downloadLink.attr('href', result.download_url);
      $downloadLink.attr('download', result.filename || '');
      $downloadContainer.show();

      // Запускаем скачивание сразу после получения ссылки.
      // Окно остаётся открытым, чтобы пользователь мог воспользоваться ссылкой повторно.
      window.location.href = result.download_url;

      // Показываем предупреждения генератора, если они были возвращены сервером.
      if (Array.isArray(result.warnings) && result.warnings.length > 0) {
        DG.showExportMessage('warning', result.warnings);
      }
    } catch (error) {
      // При сетевой ошибке сохраняем окно открытым
      // и сообщаем пользователю об ошибке.
      DG.showExportMessage(
        'error',
        error.message || 'Document generation failed.'
      );
    } finally {
      // Разблокируем кнопку после завершения запроса,
      // если форма ещё существует.
      if (submitButton && document.body.contains(submitButton)) {
        submitButton.disabled = false;
      }
    }
  });


  // Проверяем форму при изменении имени выходного файла.
  $(document).on('input', '#dg_file_name', function() {
    DG.updateInfo();
  });


  // Пересчитываем состояние формы при изменении режима выгрузки.
  $(document).on('change', 'input[name="export_mode"]', function() {
    DG.updateInfo();
  });


  // При выборе другого шаблона одновременно меняем доступность
  // настройки форматирования «Описание».
  $(document).on('change', '#template_file', function() {
    DG.updateInfo();
  });

  //v2610081634
});