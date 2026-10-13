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

/**
 * Создаёт полноэкранный слой прогресса поверх модального окна.
 * form — форма генерации с локализованными data-атрибутами.
 */
DG.showProgressOverlay = function(form) {
  DG.hideProgressOverlay();

  var overlay = document.createElement('div');
  overlay.id = 'dg-progress-overlay';
  overlay.className = 'dg-progress-overlay';
  overlay.setAttribute('role', 'dialog');
  overlay.setAttribute('aria-modal', 'true');
  overlay.setAttribute('aria-live', 'polite');

  var panel = document.createElement('div');
  panel.className = 'dg-progress-panel';

  var title = document.createElement('h3');
  title.className = 'dg-progress-title';
  title.textContent = form.dataset.progressTitle;

  var stage = document.createElement('div');
  stage.id = 'dg-progress-stage';
  stage.className = 'dg-progress-stage';
  stage.textContent =
    form.dataset.progressStagePreparing + ': ' +
    form.dataset.progressPreparing;

  var track = document.createElement('div');
  track.className = 'dg-progress-track';
  track.setAttribute('role', 'progressbar');
  track.setAttribute('aria-valuemin', '0');
  track.setAttribute('aria-valuemax', '100');
  track.setAttribute('aria-valuenow', '0');

  var bar = document.createElement('div');
  bar.id = 'dg-progress-bar';
  bar.className = 'dg-progress-bar';
  track.appendChild(bar);

  var percent = document.createElement('div');
  percent.id = 'dg-progress-percent';
  percent.className = 'dg-progress-percent';
  percent.textContent = '0%';

  var message = document.createElement('p');
  message.id = 'dg-progress-message';
  message.className = 'dg-progress-message';
  message.textContent = form.dataset.progressWait;

  panel.appendChild(title);
  panel.appendChild(stage);
  panel.appendChild(track);
  panel.appendChild(percent);
  panel.appendChild(message);
  overlay.appendChild(panel);
  document.body.appendChild(overlay);

  return overlay;
};

/**
 * Обновляет индикатор по состоянию, полученному от контроллера.
 * state — объект состояния операции; form — форма с локализованными надписями.
 */
DG.updateProgressOverlay = function(state, form) {
  var percent = Math.max(0, Math.min(100, parseInt(state.percent, 10) || 0));
  var bar = document.getElementById('dg-progress-bar');
  var track = bar && bar.parentElement;
  var percentLabel = document.getElementById('dg-progress-percent');
  var stageLabel = document.getElementById('dg-progress-stage');
  var messageLabel = document.getElementById('dg-progress-message');

  if (bar) {
    bar.style.width = percent + '%';
  }

  if (track) {
    track.setAttribute('aria-valuenow', String(percent));
  }

  if (percentLabel) {
    percentLabel.textContent = percent + '%';
  }

  if (stageLabel) {
    stageLabel.textContent = state.stage === 'rendering'
      ? form.dataset.progressStageRendering + ': ' + form.dataset.progressRendering
      : form.dataset.progressStagePreparing + ': ' + form.dataset.progressPreparing;
  }

  if (messageLabel) {
    messageLabel.textContent = state.message || form.dataset.progressWait;
  }
};

/**
 * Удаляет слой прогресса после завершения операции.
 */
DG.hideProgressOverlay = function() {
  var overlay = document.getElementById('dg-progress-overlay');

  if (overlay) {
    overlay.remove();
  }
};

/**
 * Закрывает модальное окно, показывает ошибку в Redmine и прокручивает страницу вверх.
 * form — форма генерации; message — локализованный текст ошибки.
 */
DG.finishExportWithError = function(form, message) {
  DG.hideProgressOverlay();

  if (typeof hideModal === 'function') {
    hideModal(form.querySelector('#dg-submit-btn') || form);
  } else if ($('#ajax-modal').length && $('#ajax-modal').dialog) {
    $('#ajax-modal').dialog('close');
  }

  DG.showExportMessage('error', message);
  window.scrollTo({ top: 0, behavior: 'smooth' });
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



  // Отправляет форму, опрашивает серверный прогресс и обрабатывает результат.
  $(document).on('submit', '#document-generator-form', async function(event) {
    event.preventDefault();

    var form = this;
    var submitButton = form.querySelector('#dg-submit-btn');
    var $downloadContainer = $('#dg-download-link-container');
    var $downloadLink = $('#dg-download-link');

    // Генерируем UUID заранее, чтобы POST и GET использовали одну операцию.
    var progressUuid = (window.crypto && typeof window.crypto.randomUUID === 'function')
      ? window.crypto.randomUUID()
      : 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function(character) {
          var random = Math.random() * 16 | 0;
          var value = character === 'x' ? random : (random & 0x3 | 0x8);
          return value.toString(16);
        });

    var pollTimer = null;
    var pollBusy = false;

    // Останавливает опрос и удаляет индикатор.
    function stopProgressPolling() {
      if (pollTimer !== null) {
        window.clearInterval(pollTimer);
        pollTimer = null;
      }

      DG.hideProgressOverlay();
    }

    // Получает последнее состояние операции.
    // Повторная попытка выполняется по следующему таймеру при сетевой ошибке.
    async function pollProgress() {
      if (pollBusy) return;

      pollBusy = true;

      try {
        var progressUrl = new URL(form.action, window.location.href);

        // Сохраняем исходный адрес проекта и заменяем только конечный маршрут.
        progressUrl.pathname = progressUrl.pathname.replace(
          /\/export$/,
          '/progress/' + progressUuid
        );

        var progressResponse = await fetch(progressUrl.toString(), {
          method: 'GET',
          credentials: 'same-origin',
          headers: {
            'Accept': 'application/json',
            'X-Requested-With': 'XMLHttpRequest'
          }
        });

        // Файл состояния может ещё не существовать при самом первом запросе.
        if (!progressResponse.ok) return;

        var progressState = await progressResponse.json();

        if (progressState.status !== 'pending') {
          DG.updateProgressOverlay(progressState, form);
        }
      } catch (pollError) {
        // Ошибка отдельного GET-запроса не прерывает генерацию.
      } finally {
        pollBusy = false;
      }
    }

    // Блокируем повторную отправку и скрываем ссылку от предыдущей операции.
    if (submitButton) {
      submitButton.disabled = true;
    }

    $downloadContainer.hide();
    $downloadLink.attr('href', '#');

    DG.showProgressOverlay(form);

    // Первый опрос выполняется сразу, последующие — раз в секунду.
    pollProgress();
    pollTimer = window.setInterval(pollProgress, 1000);

    try {
      var formData = new FormData(form);
      formData.set('progress_uuid', progressUuid);

      var response = await fetch(form.action, {
        method: 'POST',
        body: formData,
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

      // При любой ошибке убираем индикатор и закрываем модальное окно.
      if (!response.ok || result.type !== 'success' || !result.download_url) {
        DG.finishExportWithError(
          form,
          result.message || result.error || form.dataset.progressError
        );
        return;
      }

      // Успешное завершение: закрываем опрос, но оставляем диалог открытым.
      stopProgressPolling();

      $downloadLink.attr('href', result.download_url);
      $downloadLink.attr('download', result.filename || '');
      $downloadContainer.show();

      // Сохраняем существующий механизм автоматического скачивания.
      window.location.href = result.download_url;

      // Показываем предупреждения, возвращённые сервером.
      if (Array.isArray(result.warnings) && result.warnings.length > 0) {
        DG.showExportMessage('warning', result.warnings);
      }
    } catch (error) {
      // Сетевая ошибка также приводит к закрытию окна и сообщению на странице.
      DG.finishExportWithError(form, form.dataset.progressError);
    } finally {
      // Дополнительная очистка гарантирует отсутствие индикатора после завершения.
      stopProgressPolling();

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

  //v2610131609
});