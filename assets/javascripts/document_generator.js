var DG = DG || {};

DG.validateFilename = function(filename) {
  // Проверка на недопустимые символы: / \ : * ? " < > |
  var invalidChars = /[\/\\:*?"<>|]/;
  return !invalidChars.test(filename) && filename.trim().length > 0;
};

DG.updateInfo = function() {
  var recordCountEl = $('#dg-records-count');
  if (!recordCountEl.length) return; // Если форма еще не в DOM, выходим

  var recordCount = parseInt(recordCountEl.data('count'), 10) || 0;
  var exportMode = $('input[name="export_mode"]:checked').val();
  var fileName = $('#dg_file_name').val();
  var isValid = DG.validateFilename(fileName);
  
  var $filesCount = $('#dg-files-count');
  if (exportMode === 'single') {
    $filesCount.text($filesCount.data('text-multiple'));
  } else {
    $filesCount.text($filesCount.data('text-single'));
  }
  
  // Блокируем/разблокируем кнопку выгрузки
  $('#dg-submit-btn').prop('disabled', !isValid);
  
  // Показываем/скрываем ошибку имени файла
  if (fileName.length > 0 && !isValid) {
    $('#dg-filename-error').show();
  } else {
    $('#dg-filename-error').hide();
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
});