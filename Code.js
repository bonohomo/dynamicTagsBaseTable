/**
 * Process imported "template" sheet: read courses (col B) and proposed tags (col C),
 * then update sheet "Keys": add new tags, append new course URLs to existing tags,
 * dedupe & sort URLs in each cell, and color rows:
 *  - new tag -> A:D blue
 *  - existing tag that received new courses -> A:D green
 *
 * Подходит для шаблона вида:
 *  A1: "Название курса", B1: "Ссылка на курс, если есть", C1: "Предлагаемые теги через запятую"
 *
 * Настройки цветов можно менять в HEX ниже.
 */

/* ====== Настройки ====== */
const KEYS_SHEET_NAME = 'Keys';    // имя листа с тегами
const TEMPLATE_NAME_HINTS = ['шаблон', 'template', 'новых тег', 'new tags']; // ищем лист шаблона по названию (частичное совпадение)
const COLOR_NEW_COURSES = '#b7e1cd'; // светло-зелёный (измените по вкусу)
const COLOR_NEW_TAG = '#c9daf8';     // светло-синий (измените по вкусу)
const START_ROW_KEYS = 2; // с какой строки в Keys начинаются данные (A2:A)
const COL_TAG = 1;   // столбец A (тег)
const COL_COUNT = 2; // столбец B (количество / можно обновлять)
const COL_URLS = 3;  // столбец C (urls, comma-separated)
const LAST_COLOR_COL = 4; // подсвечиваем A:D -> 1..4
const DELETED_LABELS_SHEET_NAME = 'deleted_labels'; // архив удалённых тегов
const DELETED_COURSES_SHEET_NAME = 'deleted_courses'; // архив курсов, удалённых из тегов
/* ====================== */

/**
 * МЕНЮ СКАЧИВАНИЯ CSV
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📥 Скачать CSV")
    .addItem("Рубрики на сайт", "downloadLabels")
    .addItem("Рубрики для заполнения контентом", "downloadLabelsToFill")
    .addItem("Контент для рубрик", "downloadLabelsContent")
    .addItem("Контент для пар рубрик", "downloadPairsContent")
    .addSeparator()
    .addItem("Удалить теги или курс", "openDeleteKeysDialog")
    .addToUi();
}


// ===== КНОПКИ =====

function downloadLabels() {
  downloadSheetCSV("labels");
}

function downloadLabelsToFill() {
  downloadSheetCSV("labels_to_fill");
}

function downloadLabelsContent() {
  downloadSheetCSV("labels_content");
}

function downloadPairsContent() {
  downloadSheetCSV("pairs_content");
}


// ===== ОСНОВНАЯ ФУНКЦИЯ =====

function downloadSheetCSV(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    SpreadsheetApp.getUi().alert("Лист не найден: " + sheetName);
    return;
  }

  const values = sheet.getDataRange().getValues();

  const filtered = values.filter(row =>
    row.some(cell => cell !== "" && cell !== null)
  );

  const csv = filtered.map(row =>
    row.map(cell => {
      if (!cell) return "";
      return `"${cell.toString().replace(/"/g, '""')}"`;
    }).join(",")
  ).join("\n");

  const blob = Utilities.newBlob(csv, "text/csv", sheetName + ".csv");

  const file = DriveApp.createFile(blob);

  const url = file.getDownloadUrl();

  const html = HtmlService.createHtmlOutput(`
  <div style="font-family:sans-serif">
    <p>Подготовка файла... ⏳</p>
    <p id="fallback" style="display:none">
      Если ничего не произошло, нажмите:
      <br><br>
      <a href="${url}" target="_blank">Скачать CSV</a>
    </p>

    <script>
      const win = window.open("${url}", "_blank");

      if (!win) {
        document.getElementById("fallback").style.display = "block";
      } else {
        google.script.host.close();
      }
    </script>
  </div>
`);

  SpreadsheetApp.getUi().showModalDialog(html, "Скачивание...");
}

/**
 * ЗАПУСК ОБРАБОТКИ ШАБЛОНА
 */
function extractCourseKey(link) {
  if (!link) return '';

  const match = link.match(/\/education\/([^\/?#]+)/i);
  return match ? match[1].trim() : '';
}
function sortKeysSheet(sheet) {

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return; // если нет данных

  const range = sheet.getRange(2, 1, lastRow - 1, 4);

  range.sort([
    { column: 2, ascending: false }, // B — по убыванию
    { column: 1, ascending: true }   // A — по возрастанию
  ]);

}
function processTemplate() { 
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1) Найти лист Keys
  const keysSheet = ss.getSheetByName(KEYS_SHEET_NAME);
  if (!keysSheet) {
    SpreadsheetApp.getUi().alert('Не найден лист "' + KEYS_SHEET_NAME + '". Проверьте имя листа.');
    return;
  }

  // 2) Найти лист-шаблон (по названию или по заголовкам)
  const templateSheet = findTemplateSheet(ss);
  if (!templateSheet) {
    SpreadsheetApp.getUi().alert(
      'Не найден лист шаблона (не обнаружен лист с заголовками "Название курса / Ссылка на курс / Предлагаемые теги" или с именем, содержащим "' +
      TEMPLATE_NAME_HINTS.join(', ') +
      '").'
    );
    return;
  }

  // 3) Считать существующие теги и urls из Keys
  const keysData = readKeysSheet(keysSheet);

  // 4) Считать шаблон — mapping url -> [tags]
  const templateMap = readTemplateSheet(templateSheet);

  // 5) Применить изменения: собрать for each tag список url-ов, пометить новые теги/новые курсы
  const result = mergeTemplateIntoKeys(keysData, templateMap);

  // 6) Записать изменения в Keys, подсветить
  applyChangesToKeys(keysSheet, keysData, result);

  // 7) Вставка формулы
  keysSheet.getRange("Q2").setFormula('=GETBLUECELLS("A2:D1000")');

  // 8) Отсортировать Keys, чтобы в целевую таблицу ушёл уже финальный порядок
  sortKeysSheet(keysSheet);

  // 9) Синхронизировать Keys!A:A -> "Выбор тегов для курсов" / rubrics_data!A:A
  syncKeysToRubricsData();

  SpreadsheetApp.getUi().alert(
    'Обработка завершена:\n' +
    'Новых тегов: ' + result.newTags.length +
    '\nСтрок с добавленными курсами: ' + result.updatedExisting.length
  );
}

/* --------------------- Вспомогательные функции --------------------- */

function findTemplateSheet(ss) {
  // 1) Попытка: найти лист, имя которого содержит подсказку
  const sheets = ss.getSheets();
  for (let s of sheets) {
    const name = s.getName().toLowerCase();
    for (let hint of TEMPLATE_NAME_HINTS) {
      if (name.indexOf(hint.toLowerCase()) !== -1) return s;
    }
  }
  // 2) Попытка: найти лист, у которого в первой строке есть нужные заголовки (рус/англ)
  const headerCandidates = [
    ['название курса', 'ссылка на курс', 'предлагаемые теги'],
    ['course name', 'course link', 'proposed tags'],
    ['Название курса','Ссылка на курс','Предлагаемые теги']
  ];
  for (let s of sheets) {
    const row1 = s.getRange(1,1,1,6).getValues()[0].map(v => (v || '').toString().trim().toLowerCase());
    for (let cand of headerCandidates) {
      let ok = true;
      for (let i=0;i<cand.length;i++){
        if (row1[i].indexOf(cand[i].toLowerCase()) === -1) { ok = false; break; }
      }
      if (ok) return s;
    }
  }
  // 3) Ничего не нашлось
  return null;
}

function readKeysSheet(sheet) {
  const lastRow = Math.max(sheet.getLastRow(), START_ROW_KEYS);
  const readRange = sheet.getRange(START_ROW_KEYS, 1, lastRow - START_ROW_KEYS + 1, LAST_COLOR_COL);
  const values = readRange.getValues();
  const rows = [];
  const tagIndex = {}; // lowerTag -> index in rows
  for (let i=0;i<values.length;i++){
    const r = values[i];
    const tag = (r[COL_TAG-1] || '').toString().trim();
    const urlsCell = (r[COL_URLS-1] || '').toString().trim();
    const urls = parseUrlsList(urlsCell);
    const rowIndex = START_ROW_KEYS + i;
    if (tag) {
      rows.push({ tag: tag, urls: urls, rowIndex: rowIndex });
      tagIndex[tag.toLowerCase()] = rows.length - 1;
    } else {
      // сохранить пустые строки как placeholder (чтобы корректно добавлять новые строки в конце)
      rows.push({ tag: '', urls: urls, rowIndex: rowIndex });
    }
  }
  return { rows: rows, tagIndex: tagIndex, lastRow: lastRow };
}

function readTemplateSheet(sheet) {
  // шаблон: предполагаем, что ссылки в столбце B, теги в C, начиная с row 2
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return {};
  const values = sheet.getRange(2,1,lastRow-1,3).getValues(); // A:B:C (название, ссылка, теги)
  const map = {}; // tagLower -> Set(urls)
  for (let i=0;i<values.length;i++){
    const row = values[i];
    const url = extractCourseKey((row[1] || '').toString().trim()); // col B
    const tagsCell = (row[2] || '').toString().trim(); // col C
    if (!url || !tagsCell) continue;
    const tags = tagsCell.split(',').map(t => t.toString().trim()).filter(t => t);
    for (let t of tags) {
      const key = t.toLowerCase();
      if (!map[key]) map[key] = { tagName: t, urls: new Set() };
      map[key].urls.add(url);
    }
  }
  return map;
}

function syncKeysToRubricsData() {
  const sourceSs = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = sourceSs.getSheetByName(KEYS_SHEET_NAME);

  if (!sourceSheet) {
    throw new Error('Не найден лист "' + KEYS_SHEET_NAME + '" в текущей таблице.');
  }

  const targetSpreadsheetId = '1ruH1URqdUrJTRj8uIIcbqyaEmg6KEy_KpqkR7aDc6tw';
  const targetSs = SpreadsheetApp.openById(targetSpreadsheetId);
  const targetSheet = targetSs.getSheetByName('rubrics_data');

  if (!targetSheet) {
    throw new Error('Не найден лист "rubrics_data" в таблице "Выбор тегов для курсов".');
  }

  const lastRow = sourceSheet.getLastRow();

  // Очищаем старые данные в целевой таблице, но не трогаем заголовки
  const targetLastRow = targetSheet.getLastRow();
  if (targetLastRow >= 2) {
    targetSheet.getRange(2, 1, targetLastRow - 1, 2).clearContent(); // A2:B
  }

  // Если в Keys нет данных ниже заголовка — просто выходим
  if (lastRow < 2) {
    return;
  }

  // Берём данные из A и C, начиная со 2 строки
  const values = sourceSheet.getRange(2, 1, lastRow - 1, 3).getValues(); // A:C

  // Оставляем только строки, где заполнен тег в колонке A
  const rowsToWrite = values
    .map(row => {
      const tag = String(row[0] ?? '').trim();     // колонка A
      const colC = String(row[2] ?? '').trim();    // колонка C
      return [tag, colC];
    })
    .filter(row => row[0] !== '');

  if (rowsToWrite.length === 0) {
    return;
  }

  // Записываем в rubrics_data!A2:B
  targetSheet.getRange(2, 1, rowsToWrite.length, 2).setValues(rowsToWrite);
}

/**
 * Merge template map into keysData.
 * Returns info: newTags[], updatedExisting[] (indexes),
 * and the final updated urls for each entry in keysData (mutates keysData.rows).
 */
function mergeTemplateIntoKeys(keysData, templateMap) {
  const rows = keysData.rows;
  const tagIndex = keysData.tagIndex;
  const newTags = [];
  const updatedExisting = [];

  // For each tag in templateMap
  for (let tagLower in templateMap) {
    const tagEntry = templateMap[tagLower];
    const incomingUrls = Array.from(tagEntry.urls);
    if (tagIndex.hasOwnProperty(tagLower)) {
      // существует — добавить новые URL
      const idx = tagIndex[tagLower];
      const existing = rows[idx].urls.slice(); // array
      const set = new Set(existing.map(u => u.trim()));
      let added = false;
      for (let u of incomingUrls) {
        if (!set.has(u)) {
          set.add(u);
          added = true;
        }
      }
      if (added) {
        const merged = Array.from(set).map(s => s.toString().trim()).filter(s=>s).sort();
        rows[idx].urls = merged;
        updatedExisting.push(rows[idx].rowIndex);
      }
    } else {
      // нового тега нет — создать
      const merged = incomingUrls.map(s=>s.toString().trim()).filter(s=>s).sort();
      // добавляем новую строку placeholder (будем записывать ниже)
      const newRow = {
        tag: tagEntry.tagName, // сохраняем оригинальную форму тега из шаблона
        urls: merged,
        rowIndex: null // поставим при записи (после keysData.lastRow)
      };
      rows.push(newRow);
      // Обновим индекс: новое положение — в конце массива (после добавления)
      tagIndex[tagLower] = rows.length - 1;
      newTags.push(tagEntry.tagName);
    }
  }

  return { newTags: newTags, updatedExisting: updatedExisting };
}

function applyChangesToKeys(sheet, keysData, result) {
  const rows = keysData.rows;
  const updatedExistingSet = new Set(result.updatedExisting);

  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    if (!r.tag) continue;

    // === Новый тег ===
if (!r.rowIndex) {

  // находим последнюю строку с данными
  const lastDataRow = sheet.getLastRow();
  const newRowIndex = lastDataRow + 1;

  // записываем данные
  sheet.getRange(newRowIndex, 1).setValue(r.tag);                  // A
  sheet.getRange(newRowIndex, 2).setValue(r.urls.length);          // B
  sheet.getRange(newRowIndex, 3).setValue(r.urls.join(', '));      // C

  // подсветка синяя
  sheet.getRange(newRowIndex, 1, 1, 4).setBackground('#c9daf8');

}
    else {
      // Существующий тег
      sheet.getRange(r.rowIndex, 1).setValue(r.tag);
      sheet.getRange(r.rowIndex, 2).setValue(r.urls.length);
      sheet.getRange(r.rowIndex, 3).setValue(r.urls.join(', '));

      if (updatedExistingSet.has(r.rowIndex)) {
        sheet.getRange(r.rowIndex, 1, 1, 4).setBackground('#b7e1cd'); // зелёный
      }
    }
  }
}

/* --------------------- Утилиты --------------------- */

function parseUrlsList(cellText) {
  if (!cellText) return [];
  // split by comma (and semicolon?), trim, filter empty
  const parts = cellText.toString().split(/[;,]+/).map(s => s.trim()).filter(s => s);
  // normalize duplicates
  const set = new Set(parts);
  return Array.from(set).sort();
}


/* --------------------- УДАЛЕНИЕ ТЕГОВ И КУРСОВ --------------------- */

function openDeleteKeysDialog() {
  const template = HtmlService.createTemplateFromFile('delete');
  template.dialogData = getDeleteKeysDialogData();

  const html = template.evaluate()
    .setWidth(820)
    .setHeight(720);

  SpreadsheetApp.getUi().showModalDialog(html, 'Удаление тегов и курсов');
}

function getDeleteKeysDialogData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const keysSheet = ss.getSheetByName(KEYS_SHEET_NAME);

  if (!keysSheet) {
    throw new Error('Не найден лист "' + KEYS_SHEET_NAME + '".');
  }

  const keysData = readKeysSheet(keysSheet);
  const availableTags = [];
  const courseTagMap = {};

  keysData.rows.forEach(row => {
    if (!row.tag) return;

    availableTags.push(row.tag);

    row.urls.forEach(courseKey => {
      if (!courseTagMap[courseKey]) {
        courseTagMap[courseKey] = [];
      }
      courseTagMap[courseKey].push(row.tag);
    });
  });

  availableTags.sort((a, b) => a.localeCompare(b, 'ru'));
  Object.keys(courseTagMap).forEach(courseKey => {
    courseTagMap[courseKey].sort((a, b) => a.localeCompare(b, 'ru'));
  });

  return { availableTags: availableTags, courseTagMap: courseTagMap };
}

function deleteCourseFromKeys(courseKey) {
  const normalizedCourseKey = String(courseKey || '').trim();

  if (!normalizedCourseKey) {
    throw new Error('Не передан ключ курса для удаления.');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const keysSheet = ss.getSheetByName(KEYS_SHEET_NAME);

  if (!keysSheet) {
    throw new Error('Не найден лист "' + KEYS_SHEET_NAME + '".');
  }

  const keysData = readKeysSheet(keysSheet);
  const deletedCourseTags = [];
  let rowsUpdated = 0;

  keysData.rows.forEach(row => {
    if (!row.tag) return;

    const filteredUrls = row.urls.filter(url => url !== normalizedCourseKey);

    if (filteredUrls.length !== row.urls.length) {
      deletedCourseTags.push(row.tag);
      row.urls = filteredUrls;
      keysSheet.getRange(row.rowIndex, COL_COUNT).setValue(filteredUrls.length);
      keysSheet.getRange(row.rowIndex, COL_URLS).setValue(filteredUrls.join(', '));
      rowsUpdated += 1;
    }
  });

  if (rowsUpdated === 0) {
    throw new Error('Курс не найден в Keys!C2:C.');
  }

  archiveDeletedCourse_(ss, normalizedCourseKey, deletedCourseTags);
  finalizeKeysAfterDelete_(keysSheet);

  return {
    rowsUpdated: rowsUpdated,
    dialogData: getDeleteKeysDialogData()
  };
}

function deleteTagsFromKeys(tags) {
  const tagsToDelete = Array.isArray(tags) ? tags.map(tag => String(tag || '').trim()).filter(Boolean) : [];

  if (tagsToDelete.length === 0) {
    throw new Error('Не выбраны теги для удаления.');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const keysSheet = ss.getSheetByName(KEYS_SHEET_NAME);

  if (!keysSheet) {
    throw new Error('Не найден лист "' + KEYS_SHEET_NAME + '".');
  }

  const tagSet = new Set(tagsToDelete.map(tag => tag.toLowerCase()));
  const lastRow = keysSheet.getLastRow();
  const rowsToDelete = [];
  const rowsToArchive = [];

  for (let row = START_ROW_KEYS; row <= lastRow; row++) {
    const tag = String(keysSheet.getRange(row, COL_TAG).getValue() || '').trim();

    if (tag && tagSet.has(tag.toLowerCase())) {
      rowsToDelete.push(row);
      rowsToArchive.push(keysSheet.getRange(row, COL_TAG, 1, LAST_COLOR_COL).getValues()[0]);
    }
  }

  if (rowsToDelete.length === 0) {
    throw new Error('Выбранные теги не найдены в Keys!A2:A.');
  }

  archiveDeletedLabels_(ss, rowsToArchive);

  for (let i = rowsToDelete.length - 1; i >= 0; i--) {
    keysSheet.deleteRow(rowsToDelete[i]);
  }

  finalizeKeysAfterDelete_(keysSheet);

  return {
    tagsDeleted: rowsToDelete.length,
    dialogData: getDeleteKeysDialogData()
  };
}

function archiveDeletedCourse_(ss, courseKey, tags) {
  const normalizedCourseKey = String(courseKey || '').trim();
  const courseTags = Array.isArray(tags) ? tags.map(tag => String(tag || '').trim()).filter(Boolean) : [];

  if (!normalizedCourseKey || courseTags.length === 0) {
    return;
  }

  const deletedCoursesSheet = ss.getSheetByName(DELETED_COURSES_SHEET_NAME);

  if (!deletedCoursesSheet) {
    throw new Error('Не найден лист "' + DELETED_COURSES_SHEET_NAME + '".');
  }

  const archivedTags = Array.from(new Set(courseTags)).sort((a, b) => a.localeCompare(b, 'ru')).join(', ');
  const startRow = deletedCoursesSheet.getLastRow() + 1;

  deletedCoursesSheet
    .getRange(startRow, 1, 1, 2)
    .setValues([[normalizedCourseKey, archivedTags]]);
}

function archiveDeletedLabels_(ss, rows) {
  const rowsToArchive = Array.isArray(rows) ? rows.filter(row => Array.isArray(row) && row.length) : [];

  if (rowsToArchive.length === 0) {
    return;
  }

  const deletedLabelsSheet = ss.getSheetByName(DELETED_LABELS_SHEET_NAME);

  if (!deletedLabelsSheet) {
    throw new Error('Не найден лист "' + DELETED_LABELS_SHEET_NAME + '".');
  }

  const normalizedRows = rowsToArchive.map(row => row.slice(0, LAST_COLOR_COL));
  const startRow = deletedLabelsSheet.getLastRow() + 1;

  deletedLabelsSheet
    .getRange(startRow, COL_TAG, normalizedRows.length, LAST_COLOR_COL)
    .setValues(normalizedRows);
}

function finalizeKeysAfterDelete_(keysSheet) {
  const keysData = readKeysSheet(keysSheet);

  keysData.rows.forEach(row => {
    if (!row.tag) return;

    keysSheet.getRange(row.rowIndex, COL_COUNT).setValue(row.urls.length);
    keysSheet.getRange(row.rowIndex, COL_URLS).setValue(row.urls.join(', '));
  });

  sortKeysSheet(keysSheet);
  syncKeysToRubricsData();
  updateAll();
}

/* --------------------- ИМПОРТ ТАБЛИЦЫ С ТЕГАМИ --------------------- */

function openImportDialog() {
  const html = HtmlService.createHtmlOutputFromFile('import')
    .setWidth(420)
    .setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, 'Импорт XLSX в Tags');
}
function importXLSX(base64Data, fileName) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tagsSheet = ss.getSheetByName("Tags");

  if (!tagsSheet) throw new Error("Лист 'Tags' не найден");

  const blob = Utilities.newBlob(
    Utilities.base64Decode(base64Data),
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    fileName
  );

  const tempFile = DriveApp.createFile(blob);

  const converted = Drive.Files.copy(
    { mimeType: MimeType.GOOGLE_SHEETS },
    tempFile.getId()
  );

  const tempSpreadsheet = SpreadsheetApp.openById(converted.id);
  const firstSheet = tempSpreadsheet.getSheets()[0];

  const lastRow = firstSheet.getLastRow();
  if (lastRow < 2) return;

  const data = firstSheet.getRange(2, 1, lastRow - 1, 3).getValues();

  tagsSheet.getRange("A2:C").clearContent();
  tagsSheet.getRange(2, 1, data.length, 3).setValues(data);

  DriveApp.getFileById(tempFile.getId()).setTrashed(true);
  DriveApp.getFileById(converted.id).setTrashed(true);

  // запускаем обработку тегов
  processTemplate();
}

/* --------------------- НАЙТИ НОВЫЕ ТЕГИ --------------------- */

function GETBLUECELLS(rangeA1) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Keys");
  const range = sheet.getRange(rangeA1);
  
  const values = range.getValues();
  const backgrounds = range.getBackgrounds();
  
  const targetColor = "#c9daf8";
  const result = [];

  for (let i = 0; i < values.length; i++) {
    let row = [];
    let hasColor = false;

    // 👇 ИДЁМ ТОЛЬКО ПО ПЕРВЫМ 3 СТОЛБЦАМ (A,B,C)
    for (let j = 0; j < Math.min(3, values[i].length); j++) {
      
      if (backgrounds[i][j].toLowerCase() === targetColor) {
        row.push(values[i][j]);
        hasColor = true;
      } else {
        row.push("");
      }
    }

    if (hasColor) {
      result.push(row);
    }
  }

  return result;
}










/**
 * Копирует текущую таблицу в корень Google Drive с именем:
 * "Динамические Теги для сайта DD.MM.YYYY"
 * Если файлы с такой датой уже есть, добавляет " upd2", " upd3" и т.д.
 * После создания показывает модальное окно с ссылкой на новую копию.
 */
function copySpreadsheetWithDateAndShowDialog() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const fileId = ss.getId();
    const file = DriveApp.getFileById(fileId);

    // Получаем часовой пояс таблицы (надёжнее для отображения даты)
    const tz = (typeof ss.getSpreadsheetTimeZone === 'function') ? ss.getSpreadsheetTimeZone() : Session.getTimeZone();
    const today = Utilities.formatDate(new Date(), tz, "dd.MM.yyyy");

    const baseName = `Динамические Теги для сайта ${today}`;

    // Ищем все файлы, у которых в имени содержится baseName
    // (т.е. и exact match, и возможные " updN")
    const query = 'title contains "' + baseName.replace(/"/g, '\\"') + '"';
    const files = DriveApp.searchFiles(query);

    // Найдём максимально существующий суффикс updN (или отметим наличие exact match)
    let maxSuffix = 0;   // 0 = нет файла с baseName и без updN
    while (files.hasNext()) {
      const f = files.next();
      const name = f.getName();

      if (name === baseName) {
        // уже есть точный файл, считаем как суффикс 1 (следующий будет upd2)
        maxSuffix = Math.max(maxSuffix, 1);
      } else {
        // ищем шаблон "baseName updN"
        const re = new RegExp('^' + escapeRegExp(baseName) + ' upd(\\d+)$');
        const m = name.match(re);
        if (m && m[1]) {
          const num = parseInt(m[1], 10);
          if (!isNaN(num)) maxSuffix = Math.max(maxSuffix, num);
        }
      }
    }

    // выбираем имя-кандидат
    let candidateName = baseName;
    if (maxSuffix >= 1) candidateName = baseName + ' upd' + (maxSuffix + 1);

    // Копируем в корень (можешь заменить на DriveApp.getFolderById(ID))
    const destinationFolder = DriveApp.getRootFolder();
    const newFile = file.makeCopy(candidateName, destinationFolder);
const newUrl = newFile.getUrl();

// === Убираем заливку в копии ===
const newSpreadsheet = SpreadsheetApp.openById(newFile.getId());
const keysSheet = newSpreadsheet.getSheetByName("Keys");

if (keysSheet) {
  const lastRow = keysSheet.getLastRow();
  if (lastRow >= 2) {
    keysSheet.getRange(2, 1, lastRow - 1, 4).setBackground(null);
  }
}

// === Очищаем только заполненные ячейки в диапазоне Q2:S ===
if (keysSheet) {
  const lastRow = keysSheet.getLastRow();
  if (lastRow >= 2) {
    const range = keysSheet.getRange(2, 17, lastRow - 1, 3); // Q=17, R=18, S=19
    const values = range.getValues();

    for (let i = 0; i < values.length; i++) {
      for (let j = 0; j < values[i].length; j++) {
        if (values[i][j] !== "" && values[i][j] !== null) {
          range.getCell(i + 1, j + 1).clearContent();
        }
      }
    }
  }
}

// === Заменяем TRUE на FALSE в Changes!A2:A700 ===
const changesSheet = newSpreadsheet.getSheetByName("Changes");

if (changesSheet) {
  const range = changesSheet.getRange("A2:A700");
  const values = range.getValues();

  const updated = values.map(row => [
    row[0] === true ? false : row[0]
  ]);

  range.setValues(updated);
}

    // Показать модальное окно с ссылкой и кнопками
    const htmlContent =
      `<div style="font-family:Arial,Helvetica,sans-serif;padding:12px;">
         <h3 style="margin:0 0 8px 0;">Копия создана</h3>
         <p style="margin:0 0 8px 0;">Файл: <b>${escapeHtml(candidateName)}</b></p>
         <p style="margin:0 0 12px 0;">
           <a href="${newUrl}" target="_blank">${newUrl}</a>
         </p>
         <div style="display:flex;gap:8px;">
           <button onclick="window.open('${newUrl}','_blank')" style="padding:6px 10px;">Открыть в новой вкладке</button>
           <button onclick="google.script.host.close()" style="padding:6px 10px;">Закрыть</button>
         </div>
       </div>`;

    const html = HtmlService.createHtmlOutput(htmlContent).setWidth(420).setHeight(160);
    SpreadsheetApp.getUi().showModalDialog(html, 'Копирование завершено');

    // лог и возврат (на случай, если функция вызывается программно)
    Logger.log('Копия создана: ' + candidateName + ' — ' + newUrl);
    return { name: candidateName, url: newUrl };

  } catch (err) {
    const ui = SpreadsheetApp.getUi();
    ui.alert('Ошибка при копировании: ' + err.message);
    Logger.log(err);
    throw err;
  }
}

/** Экранирует строку для RegExp */
function escapeRegExp(string) {
  return String(string).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/** Экранирует HTML (имена файлов) */
function escapeHtml(text) {
  return String(text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}


/**
* Подгрузка таблицы в таблицу
 */

function showUploadDialog() {
  const html = HtmlService.createHtmlOutput(getUploadHtml())
    .setWidth(500)
    .setHeight(320);
  SpreadsheetApp.getUi().showModalDialog(html, "Импорт XLSX в Content");
}

function uploadAndImportXlsx(base64, fileName) {

  try {

    const bytes = Utilities.base64Decode(base64);
    const blob = Utilities.newBlob(
      bytes,
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      fileName
    );

    // ⭐ ВАЖНО — правильная конвертация XLSX → Google Sheets
    const resource = {
      name: fileName,
      mimeType: "application/vnd.google-apps.spreadsheet"
    };

    const file = Drive.Files.create(resource, blob, { convert: true });

    const tempSS = SpreadsheetApp.openById(file.id);
    const sourceSheet = tempSS.getSheets()[0];
    const data = sourceSheet.getDataRange().getValues();

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let content = ss.getSheetByName("Content");
    if (!content) content = ss.insertSheet("Content");

    const startRow = 1;
    const startCol = 3;

    // ⭐ ОЧИЩАЕМ ВСЁ начиная с колонки C
const lastRow = content.getMaxRows();
const lastCol = content.getMaxColumns();

content.getRange(1, 3, lastRow, lastCol - 2).clearContent();
    content.getRange(startRow, startCol, data.length, data[0].length).setValues(data);

    // удаляем временный файл
    DriveApp.getFileById(file.id).setTrashed(true);

    return "Импорт завершён ✔";

  } catch (err) {
    return "ОШИБКА: " + err.message;
  }
}


function getUploadHtml() {
  return `
  <html>
  <head>
  <style>
    body { font-family: Arial; padding: 20px; }
    #status { margin-top: 15px; color: #444; }
    .success { color: green; font-weight: bold; }
    .loading { color: #666; }
  </style>
  </head>
  <body style="font-family:Arial;padding:15px;">
    <h3>Загрузите XLSX-файл с рубриками или парами рубрик, чтобы получить файлы для загрузки на сайт</h3>
    <input type="file" id="fileInput" accept=".xlsx"/>
    <br><br>
    <button onclick="upload()">Импортировать</button>
    <div id="status"></div>

    <script>
function upload() {
  const file = document.getElementById('fileInput').files[0];
  if (!file) {
    alert("Выберите файл");
    return;
  }

  const status = document.getElementById("status");
  status.innerHTML = "<span class='loading'>⏳ Идёт копирование элементов таблицы</span>";

  const reader = new FileReader();

  reader.onload = function(e) {
    const base64 = e.target.result.split(',')[1];

    google.script.run
      .withSuccessHandler(msg => {
        status.innerHTML = "<span class='success'>✅ Копирование успешно завершено!</span>";
        setTimeout(() => google.script.host.close(), 2000);
      })
      .withFailureHandler(err => {
        alert("❌ Ошибка: " + err.message);
      })
      .uploadAndImportXlsx(base64, file.name);
  };

  reader.readAsDataURL(file);
}
</script>
  </body>
  </html>
  `;
}

/* --------------------- Функции заменяющие формулы --------------------- */

function SPLIT_SORT_UNIQUE(range) {
  if (!range) return [];

  // Превращаем диапазон в плоский массив
  const flat = range.flat().filter(v => v !== "" && v != null);

  // Объединяем без лимита TEXTJOIN
  const joined = flat.join(", ");

  // Делим по разделителю
  const split = joined.split(", ");

  // Убираем дубликаты
  const unique = [...new Set(split)];

  // Сортируем (как SORT)
  unique.sort((a, b) => a.localeCompare(b));

  // Возвращаем столбцом (аналог TRANSPOSE)
  return unique.map(v => [v]);
}

function COUNT_BY_DELIMITER(range) {
  if (!range) return [];

  return range.map(row => {
    return row.map(cell => {
      if (cell === "" || cell == null) return 0;

      // Приводим к строке
      const text = String(cell);

      // Разбиваем по ", "
      const parts = text.split(", ");

      // Количество элементов
      return parts.length;
    });
  });
}


/**
 * Исправленный объединённый скрипт.
 * Главное исправления:
 *  - надёжно парсим Tags!C (split по запятым/точкам с запятой + trim + unique)
 *  - перед массовой записью очищаем старые DataValidations, чтобы не получать Exception
 *  - аккуратно создаём и назначаем новые валидации
 *
 * Форматы:
 * Tags: A = tag, B = popularity (число, опц.), C = comma-separated courses
 *
 * DynamicTags:
 *   A = courses, B = tag_count, C... = tag1, tag2, ...
 *
 * DynamicPairs:
 *   A = tag1 (more popular), B = tag2, C = pop(tag1), D = pop(tag2), E = common_count, F = common_courses
 */


/* ------------------ Helpers ------------------ */

// Разбивает строку по запятой / точке с запятой в массив уникальных trimmed значений
function splitList(cellText) {
  if (cellText === undefined || cellText === null) return [];
  const s = cellText.toString();
  if (s.trim() === '') return [];
  // split by comma or semicolon
  const parts = s.split(/[;,]+/).map(x => x.trim()).filter(Boolean);
  // unique preserving order
  const seen = new Set();
  const out = [];
  parts.forEach(p => {
    if (!seen.has(p)) { seen.add(p); out.push(p); }
  });
  return out;
}

function ensureRows(sheet, neededRows) {
  const have = sheet.getMaxRows();
  if (have < neededRows) {
    sheet.insertRowsAfter(have, neededRows - have);
  }
}

function clearAllValidationsOnSheet(sheet) {
  // безопасно очищаем валидации по существующему диапазону
  const maxRows = sheet.getMaxRows();
  const maxCols = sheet.getMaxColumns();
  if (maxRows > 0 && maxCols > 0) {
    sheet.getRange(1, 1, maxRows, maxCols).clearDataValidations();
  }
}

/* ------------------ Build maps ------------------ */
/** Возвращает объект { tagPopularity, tagToCourses, courseToTags } */
function buildTagCourseMaps() {
  const ss = SpreadsheetApp.getActive();
  const src = ss.getSheetByName('Keys');
  if (!src) return null;

  const data = src.getDataRange().getValues();
  const tagPopularity = {};    // tag -> number
  const tagToCourses = {};     // tag -> Set(courses)
  const courseToTags = {};     // course -> Set(tags)

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const tag = (row[0] || '').toString().trim();
    if (!tag) continue;

    // popularity safe parse
    const rawPop = row[1];
    const popNum = (rawPop === '' || rawPop === null || rawPop === undefined) ? 0 : Number(rawPop);
    tagPopularity[tag] = isFinite(popNum) ? popNum : 0;

    const coursesCell = (row[2] || '').toString().trim();
    const courses = splitList(coursesCell);
    if (!tagToCourses[tag]) tagToCourses[tag] = new Set();
    courses.forEach(course => {
      tagToCourses[tag].add(course);
      if (!courseToTags[course]) courseToTags[course] = new Set();
      courseToTags[course].add(tag);
    });
  }

  return { tagPopularity, tagToCourses, courseToTags };
}


/* ------------------ DynamicTags (новая логика столбцов) ------------------ */
function updateDynamicTags() {
  const maps = buildTagCourseMaps();
  if (!maps) return;
  const { courseToTags, tagPopularity } = maps;

  const ss = SpreadsheetApp.getActive();
  const sheetName = 'DynamicTags';
  let dst = ss.getSheetByName(sheetName);
  if (!dst) dst = ss.insertSheet(sheetName);

  // Список курсов
  const courses = Object.keys(courseToTags).sort((a, b) => a.localeCompare(b, 'ru'));

  // максимальное число тегов у курсов
  let maxTags = 0;
  courses.forEach(c => {
    const ct = courseToTags[c] ? courseToTags[c].size : 0;
    if (ct > maxTags) maxTags = ct;
  });

  // Полностью очищаем лист (контент + формат/валидации) — удобнее начать с чистого листа
  dst.clear();
  clearAllValidationsOnSheet(dst);

  // Заголовки: courses, tag_count, tag1..tagN
  const headers = ['courses', 'tag_count'];
  for (let i = 1; i <= maxTags; i++) headers.push('tag' + i);
  dst.getRange(1, 1, 1, headers.length).setValues([headers]);
  dst.setFrozenRows(1);

  if (courses.length === 0) {
    return;
  }

  // Собираем output
  const out = courses.map(course => {
    const tagsArr = Array.from(courseToTags[course] || [])
      .sort((t1, t2) => {
        const p1 = tagPopularity[t1] || 0;
        const p2 = tagPopularity[t2] || 0;
        if (p2 !== p1) return p2 - p1;
        return t1.localeCompare(t2, 'ru');
      });

    const row = [course, tagsArr.length];
    for (let i = 0; i < maxTags; i++) row.push(tagsArr[i] || '');
    return row;
  });

  ensureRows(dst, out.length + 1);
  dst.getRange(2, 1, out.length, out[0].length).setValues(out);

  // Форматируем ширины
  dst.setColumnWidth(1, 420);
  dst.setColumnWidth(2, 110);
  for (let c = 3; c <= headers.length; c++) dst.setColumnWidth(c, 220);
}


/* ------------------ DynamicPairs (исправленная логика пересечений) ------------------ */
function updateDynamicPairs() {
  const maps = buildTagCourseMaps();
  if (!maps) return;
  const { tagToCourses, tagPopularity } = maps;

  const ss = SpreadsheetApp.getActive();
  const sheetName = 'DynamicPairs';
  let dst = ss.getSheetByName(sheetName);
  if (!dst) dst = ss.insertSheet(sheetName);

  // Список всех тегов (отсортирован для детерминизма)
  const tags = Object.keys(tagToCourses).sort((a, b) => a.localeCompare(b, 'ru'));

  // Очистка листа и удаление валидаций перед записью (ВАЖНО!!)
  dst.clear();
  clearAllValidationsOnSheet(dst);

  // Заголовки
  const headers = ['tag1', 'tag2', 'pop1', 'pop2', 'common_count', 'common_courses'];
  dst.getRange(1, 1, 1, headers.length).setValues([headers]);
  dst.setFrozenRows(1);

  const pairs = []; // соберём пары с пересечениями

  for (let i = 0; i < tags.length; i++) {
    for (let j = i + 1; j < tags.length; j++) {
      const a = tags[i];
      const b = tags[j];
      const setA = tagToCourses[a] || new Set();
      const setB = tagToCourses[b] || new Set();

      // Находим пересечение (используем меньший set для ускорения)
      const common = [];
      if (setA.size <= setB.size) {
        setA.forEach(x => { if (setB.has(x)) common.push(x); });
      } else {
        setB.forEach(x => { if (setA.has(x)) common.push(x); });
      }

      if (common.length === 0) continue; // пропускаем пары без пересечения (по вашей логике)

      // сортировка и уник
      const commonSorted = Array.from(new Set(common)).sort((x, y) => x.localeCompare(y, 'ru'));

      // popularity
      const popA = tagPopularity[a] || 0;
      const popB = tagPopularity[b] || 0;

      // решаем порядок: более популярный должен быть tag1
      let tag1 = a, tag2 = b, pop1 = popA, pop2 = popB;
      if (popB > popA || (popB === popA && b.localeCompare(a, 'ru') < 0)) {
        tag1 = b; tag2 = a; pop1 = popB; pop2 = popA;
      }

      pairs.push({
        tag1: tag1,
        tag2: tag2,
        pop1: pop1,
        pop2: pop2,
        commonCount: commonSorted.length,
        commonCourses: commonSorted
      });
    }
  }

  // Сортируем пары: по commonCount desc, затем по сумме популярностей desc, затем лексически
  pairs.sort((x, y) => {
    if (y.commonCount !== x.commonCount) return y.commonCount - x.commonCount;
    const sumX = (x.pop1 || 0) + (x.pop2 || 0);
    const sumY = (y.pop1 || 0) + (y.pop2 || 0);
    if (sumY !== sumX) return sumY - sumX;
    if (x.tag1 !== y.tag1) return x.tag1.localeCompare(y.tag1, 'ru');
    return x.tag2.localeCompare(y.tag2, 'ru');
  });

  if (pairs.length === 0) return;

  // Подготовим массив для записи
  const out = pairs.map(p => [p.tag1, p.tag2, p.pop1, p.pop2, p.commonCount, p.commonCourses.join(', ')]);

  ensureRows(dst, out.length + 1);
  dst.getRange(2, 1, out.length, out[0].length).setValues(out);

  // Подправим ширины
  dst.setColumnWidth(1, 220);
  dst.setColumnWidth(2, 220);
  dst.setColumnWidth(3, 100);
  dst.setColumnWidth(4, 100);
  dst.setColumnWidth(5, 120);
  dst.setColumnWidth(6, 900);
}

/* ------------------ Dispatcher ------------------ */
function onEditTrigger(e) {
  if (!e || !e.range) return;
  const sh = e.range.getSheet();
  if (!sh) return;
  if (sh.getName() !== 'Keys') return;

  try { updateDynamicTags(); } catch (err) { console.error('updateDynamicTags error: ' + err); }
  try { updateDynamicPairs(); } catch (err) { console.error('updateDynamicPairs error: ' + err); }
}

/* ------------------ Manual runner ------------------ */
function updateAll() {
  updateDynamicTags();
  updateDynamicPairs();
}
