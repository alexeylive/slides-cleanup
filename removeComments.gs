/**
 * Удаление всех комментариев из активной презентации Google Slides.
 * 
 * Комментарии в Google Slides хранятся на уровне Drive, а не Slides API,
 * поэтому используем Drive API (Advanced Service) для их удаления.
 * 
 * Перед использованием необходимо:
 * 1. В редакторе Apps Script: Resources → Advanced Google services → включить Drive API
 * 2. Или в новом редакторе: Services → Add a service → Drive API
 */

/**
 * Главная функция для удаления всех комментариев.
 * Запускается вручную или через меню.
 */
function removeAllComments() {
  const presentation = SlidesApp.getActivePresentation();
  const fileId = presentation.getId();
  
  const deletedCount = deleteAllCommentsFromFile(fileId);
  
  const message = deletedCount > 0
    ? `Удалено комментариев: ${deletedCount}`
    : 'Комментариев не найдено';
  
  SlidesApp.getUi().alert('Удаление комментариев', message, SlidesApp.getUi().ButtonSet.OK);
}

/**
 * Удаляет все комментарии из файла по его ID.
 * @param {string} fileId - ID файла Google Drive
 * @returns {number} - количество удалённых комментариев
 */
function deleteAllCommentsFromFile(fileId) {
  let deletedCount = 0;
  let pageToken = null;
  
  do {
    const response = Drive.Comments.list(fileId, {
      pageToken: pageToken,
      pageSize: 100,
      fields: 'comments(id),nextPageToken'
    });
    
    const comments = response.comments || [];
    
    for (const comment of comments) {
      Drive.Comments.remove(fileId, comment.id);
      deletedCount++;
    }
    
    pageToken = response.nextPageToken;
  } while (pageToken);
  
  return deletedCount;
}

/**
 * Удаляет заметки докладчика (Speaker Notes) со всех слайдов.
 */
function removeSpeakerNotes() {
  const presentation = SlidesApp.getActivePresentation();
  const slides = presentation.getSlides();
  
  let clearedCount = 0;
  
  for (const slide of slides) {
    const notesPage = slide.getNotesPage();
    const speakerNotesShape = notesPage.getSpeakerNotesShape();
    const textRange = speakerNotesShape.getText();
    
    if (textRange.asString().trim().length > 0) {
      textRange.clear();
      clearedCount++;
    }
  }
  
  const message = clearedCount > 0
    ? `Очищено заметок на слайдах: ${clearedCount}`
    : 'Заметки докладчика не найдены';
  
  SlidesApp.getUi().alert('Удаление заметок', message, SlidesApp.getUi().ButtonSet.OK);
}

/**
 * Удаляет все элементы, которые полностью находятся вне видимой области слайда.
 * Элементы, хотя бы частично пересекающиеся с видимой областью, сохраняются.
 */
function removeElementsOutsideSlide() {
  const presentation = SlidesApp.getActivePresentation();
  const slides = presentation.getSlides();
  
  // Размеры видимой области слайда (в points)
  const pageWidth = presentation.getPageWidth();
  const pageHeight = presentation.getPageHeight();
  
  let totalDeleted = 0;
  
  for (const slide of slides) {
    const deletedOnSlide = removeOffSlideElements(slide, pageWidth, pageHeight);
    totalDeleted += deletedOnSlide;
  }
  
  const message = totalDeleted > 0
    ? `Удалено элементов вне слайдов: ${totalDeleted}`
    : 'Элементов вне видимой области не найдено';
  
  SlidesApp.getUi().alert('Очистка элементов', message, SlidesApp.getUi().ButtonSet.OK);
}

/**
 * Удаляет элементы вне видимой области на конкретном слайде.
 * @param {SlidesApp.Slide} slide - слайд для обработки
 * @param {number} pageWidth - ширина слайда в points
 * @param {number} pageHeight - высота слайда в points
 * @returns {number} - количество удалённых элементов
 */
function removeOffSlideElements(slide, pageWidth, pageHeight) {
  const pageElements = slide.getPageElements();
  let deletedCount = 0;
  
  // Проходим в обратном порядке, чтобы удаление не сбивало индексы
  for (let i = pageElements.length - 1; i >= 0; i--) {
    const element = pageElements[i];
    
    if (isCompletelyOutside(element, pageWidth, pageHeight)) {
      element.remove();
      deletedCount++;
    }
  }
  
  return deletedCount;
}

/**
 * Проверяет, находится ли элемент полностью вне видимой области слайда.
 * @param {SlidesApp.PageElement} element - элемент для проверки
 * @param {number} pageWidth - ширина слайда
 * @param {number} pageHeight - высота слайда
 * @returns {boolean} - true, если элемент полностью вне слайда
 */
function isCompletelyOutside(element, pageWidth, pageHeight) {
  // Получаем позицию и размеры элемента
  const left = element.getLeft();
  const top = element.getTop();
  const width = element.getWidth();
  const height = element.getHeight();
  
  // Вычисляем границы bounding box элемента
  const elementRight = left + width;
  const elementBottom = top + height;
  
  // Границы видимой области слайда
  const slideLeft = 0;
  const slideTop = 0;
  const slideRight = pageWidth;
  const slideBottom = pageHeight;
  
  // Элемент полностью вне слайда, если нет пересечения прямоугольников.
  // Пересечение есть, когда:
  //   left < slideRight AND elementRight > slideLeft AND
  //   top < slideBottom AND elementBottom > slideTop
  // 
  // Нет пересечения (полностью вне), когда хотя бы одно условие нарушено:
  const noHorizontalOverlap = (elementRight <= slideLeft) || (left >= slideRight);
  const noVerticalOverlap = (elementBottom <= slideTop) || (top >= slideBottom);
  
  return noHorizontalOverlap || noVerticalOverlap;
}

/**
 * Создаёт пользовательское меню при открытии презентации.
 */
function onOpen() {
  SlidesApp.getUi()
    .createMenu('🧹 Очистка презы')
    .addItem('Удалить все комментарии', 'removeAllComments')
    .addItem('Удалить заметки докладчика', 'removeSpeakerNotes')
    .addItem('Удалить элементы вне слайда', 'removeElementsOutsideSlide')
    .addToUi();
}
