// Перемещение ссылки "Генератор документов" в блок экспорта
// Срабатывает после полной загрузки DOM
document.addEventListener('DOMContentLoaded', function() {
  // Находим скрытый контейнер со ссылкой
  var linkContainer = document.getElementById('document-generator-export-link');
  if (linkContainer) {
    // Находим блок экспорта (Atom | CSV | PDF)
    var otherFormats = document.querySelector('p.other-formats');
    if (otherFormats) {
      // Создаём <span> для единообразия с другими ссылками
      var span = document.createElement('span');
      span.appendChild(linkContainer.querySelector('a'));
      
      // Добавляем разделитель и ссылку в конец блока
      //otherFormats.appendChild(document.createTextNode(' | '));
      otherFormats.appendChild(span);
      
      // Удаляем оригинальный скрытый контейнер
      linkContainer.remove();
    }
  }
});