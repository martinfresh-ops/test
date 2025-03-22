// Количество строк, занятых заголовками
var HEADER_ROWS = 2;

////////////////////////////
// Основной триггер редактирования
////////////////////////////

function onEditTrigger(e) {
  var range = e.range;
  var sheet = range.getSheet();
  var row = range.getRow();
  var col = range.getColumn();
  if (col === 16 && e.value && e.value.toString().toLowerCase() === "завершено") {
    sheet.getRange(row, 23).setValue(new Date());
  }
  
  // Пропускаем строки заголовков и пустые изменения
  if (row <= HEADER_ROWS || !e.value) return;

  // 1) При изменении приоритета (приоритет теперь в колонке M=13)
  if (col === 13) {
    handlePriorityChange(e, sheet, row);
  }
  // 2) При выборе подразделения (D=4)
  if (col === 4 && e.value !== "") {
    handleDepartmentSelection(sheet, row);
  }
  // 3) Форматирование текста для колонок [6,7,8,11,14,15,16,17,18,19]
  if (col === 7) {
    handleColumnGFormatting(e);
  } else if ([6,8,11,14,15,16,17,18,19].includes(col)) {
    handleTextFormatting(e);
  }
  // 4) Форматирование номера телефона (E=5)
  if (col === 5) {
    handlePhoneFormatting(e);
  }
  // 5) Обработка описания неисправности (K=11)
  if (col === 11) {
    handleIssueDescription(e, sheet, row);
  }
  // 6) Автоматическая смена статуса при выборе исполнителя (O=15, если редактируется вручную)
  if (col === 15) {
    handleExecutorSelection(sheet, row);
  }
  // 7) Исправление формата даты (J=10)
  if (col === 10) {
    handleDateFormatting(e);
  }
  // 8) Применяем стандартное форматирование к изменённой ячейке
  applyStandardFormatting(range);
}

////////////////////////////
// Функции обработки редактирования
////////////////////////////

// 1. Изменение приоритета (M=13)
function handlePriorityChange(e, sheet, row) {
  var priority = e.value;
  var requestDate = sheet.getRange(row, 2).getValue(); // B
  var requestTime = sheet.getRange(row, 3).getValue(); // C

  if (requestDate && requestTime && priority) {
    var requestDateTime = new Date(requestDate);
    if (requestTime instanceof Date) {
      requestDateTime.setHours(requestTime.getHours(), requestTime.getMinutes(), requestTime.getSeconds());
    }
    var deadline = calculateDeadline(requestDateTime, priority);
    sheet.getRange(row, 14).setValue(
      Utilities.formatDate(deadline, "GMT+10", "dd.MM.yyyy HH:mm")
    );
    // Устанавливаем статус «Не начато» в колонке P=16
    sheet.getRange(row, 16).setValue("Не начато");
    // Отправляем форму инженерам – передаём row для формирования callback_data
    sendFormToEngineers(sheet, row);
  }
}

// 2. Выбор подразделения (D=4)
function handleDepartmentSelection(sheet, row) {
  var numberCell = sheet.getRange(row, 1);
  var dateCell = sheet.getRange(row, 2);
  var timeCell = sheet.getRange(row, 3);

  if (numberCell.isBlank()) {
    var timestamp = new Date();
    var formattedNumber = Utilities.formatDate(timestamp, "GMT+10", "yyyyMMdd-HHmmss");
    var formattedDate = Utilities.formatDate(timestamp, "GMT+10", "dd.MM.yyyy");
    var formattedTime = Utilities.formatDate(timestamp, "GMT+10", "HH:mm:ss");

    numberCell.setValue(formattedNumber);
    dateCell.setValue(formattedDate);
    timeCell.setValue(formattedTime);
  }
}

// 3. Стандартное форматирование текста (с орфопроверкой)
function handleTextFormatting(e) {
  var range = e.range;
  var newText = e.value;
  var columnsToCheck = [11, 17, 18, 19];
  if (columnsToCheck.includes(e.range.getColumn())) {
    var correctedText = checkSpelling(newText);
    if (correctedText && correctedText !== newText) {
      newText = correctedText;
    }
  }
  newText = newText.toLowerCase().replace(/(^|[.!?]\s*)([а-яёa-z])/g, function(match, sep, char) {
    return sep + char.toUpperCase();
  });
  range.setValue(newText);
}

// 4. Форматирование номера телефона (E=5)
function handlePhoneFormatting(e) {
  var range = e.range;
  var phone = e.value.replace(/\D/g, "");
  if (phone.length === 11 && (phone.startsWith("8") || phone.startsWith("7"))) {
    phone = "+7(" + phone.slice(1, 4) + ") " + phone.slice(4, 7) + "-" + phone.slice(7, 9) + "-" + phone.slice(9, 11);
    range.setValue(phone);
  } else {
    SpreadsheetApp.getUi().alert("⚠️ Введите номер телефона в формате 89990809578 (11 цифр).");
  }
}

// 5. Обработка описания неисправности (K=11)
function handleIssueDescription(e, sheet, row) {
  var textK = e.value.toLowerCase();
  var statusCell = sheet.getRange(row, 12); // статус можно обновлять в данной колонке (при необходимости измените)
  var regexDiagnostic = /(полом|неисправ|сломал|треснул|вибраци|не работает|не включается|горит лампа|замыкан|стучит)/;
  var regexMontage = /(перемещ|монтаж|демонтаж|налад|перестанов|разбор|установ|перенести|занести|вынести)/;
  
  if (regexDiagnostic.test(textK)) {
    statusCell.setValue("Техническое диагностирование");
  } else if (regexMontage.test(textK)) {
    statusCell.setValue("Монтаж/демонтаж или наладка");
  }
  
  // Автозаполнение колонок H и I, если данных нет
  var hCell = sheet.getRange(row, 8);
  var iCell = sheet.getRange(row, 9);
  if (hCell.isBlank()) {
    hCell.setValue("Информация отсутствует");
  }
  if (iCell.isBlank()) {
    iCell.setValue("Информация отсутствует");
  }
}

// 6. Обработка выбора исполнителя (O=15, если редактируется вручную)
function handleExecutorSelection(sheet, row) {
  var statusCell = sheet.getRange(row, 16); // P=16
  var currentStatus = statusCell.getValue();
  if (!currentStatus || currentStatus === "Не начато") {
    statusCell.setValue("Выполняется");
  }
}

// 7. Исправление формата даты (J=10)
function handleDateFormatting(e) {
  var range = e.range;
  var rawInput = e.value.toString();
  var fixed = fixDateFormat(rawInput);
  if (fixed) {
    var parts = fixed.split(".");
    if (parts.length === 3) {
      var dd = parseInt(parts[0], 10);
      var mm = parseInt(parts[1], 10) - 1;
      var yyyy = parseInt(parts[2], 10);
      var parsedDate = new Date(yyyy, mm, dd);
      if (!isNaN(parsedDate.getTime())) {
        range.setValue(parsedDate);
        range.setNumberFormat("dd.MM.yyyy");
        return;
      }
    }
    range.setValue(fixed);
    range.setNumberFormat("@STRING@");
  }
}

// 8. Применение стандартного форматирования к ячейке
function applyStandardFormatting(range) {
  range.setFontFamily("Roboto");
  range.setFontSize(8);
  range.setFontColor("#000000");
  range.setHorizontalAlignment("left");
  range.setVerticalAlignment("middle");
  range.setWrap(true);
}

////////////////////////////
// Функции форматирования, проверки и расчёта
////////////////////////////

// Функция «умного» исправления формата даты
function fixDateFormat(input) {
  try {
    var str = input.toString()
      .replace(/[^0-9\/.,-]/g, '')
      .replace(/,/g, '.')
      .replace(/-/g, '.')
      .replace(/\//g, '.');
    var parts = str.split('.').filter(function(p){ return p.length > 0; });
    var dd = '01', mm = '01', yyyy = '';
    if (parts.length === 1 && parts[0].length === 4) {
      yyyy = parts[0];
    } else if (parts.length === 2) {
      if (parts[1].length === 4) {
        mm = parts[0];
        yyyy = parts[1];
      } else {
        mm = parts[0];
        yyyy = '20' + parts[1].padStart(2, '0');
      }
    } else if (parts.length === 3) {
      dd = parts[0];
      mm = parts[1];
      yyyy = parts[2];
      if (yyyy.length === 2) {
        yyyy = '20' + yyyy;
      }
    }
    dd = dd.padStart(2, '0').slice(-2);
    mm = mm.padStart(2, '0').slice(-2);
    yyyy = yyyy.padStart(4, '20');
    return dd + '.' + mm + '.' + yyyy;
  } catch(e) {
    return input.toString();
  }
}
// Функция проверки орфографии
function checkSpelling(text) {
  var url = 'https://speller.yandex.net/services/spellservice.json/checkText';
  var response = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/x-www-form-urlencoded',
    payload: { text: text, lang: 'ru', format: 'plain' },
    muteHttpExceptions: true
  });
  var corrections = JSON.parse(response.getContentText());
  if (corrections.length === 0) return text;
  corrections.forEach(function(correction) {
    text = text.replace(correction.word, correction.s[0]);
  });
  return text;
}
// Функция проверки рабочего времени инженеров
function isWorkingTime(date) {
  var day = date.getDay();
  var hours = date.getHours();
  var minutes = date.getMinutes();
  var timeInMinutes = hours * 60 + minutes;
  var workStart = 8 * 60 + 30;  // 08:30
  var workEnd = 17 * 60;        // 17:00
  var holidays = ['01.01','07.01','23.02','08.03','01.05','09.05','12.06','04.11'];
  var formattedDate = Utilities.formatDate(date, "GMT+10", "dd.MM");
  if (holidays.includes(formattedDate)) return false;
  return day >= 1 && day <= 5 && timeInMinutes >= workStart && timeInMinutes < workEnd;
}
// Функция расчёта дедлайна реакции
function calculateDeadline(startDate, priority) {
  var reactionMinutes;
  switch(priority) {
    case 'Низкий': reactionMinutes = 8 * 60; break;
    case 'Средний': reactionMinutes = 6 * 60; break;
    case 'Высокий': reactionMinutes = 2 * 60; break;
    case 'Экстренный': reactionMinutes = 15; break;
    default: reactionMinutes = 8 * 60; break;
  }
  var deadline = new Date(startDate);
  if (priority === 'Экстренный' && !isWorkingTime(startDate)) {
    deadline.setMinutes(deadline.getMinutes() + 15 + 60);
    return deadline;
  }
  while (reactionMinutes > 0) {
    if (isWorkingTime(deadline)) {
      deadline.setMinutes(deadline.getMinutes() + 1);
      reactionMinutes--;
    } else {
      deadline.setMinutes(deadline.getMinutes() + 1);
      while (!isWorkingTime(deadline)) {
        deadline.setMinutes(deadline.getMinutes() + 1);
      }
    }
  }
  return deadline;
}
// Функция проверки времени реакции и изменения цвета ячейки (тестовые пороги)
function checkResponseTime() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var timeZone = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  var lastRow = sheet.getLastRow();
  var now = new Date();
  for (var row = HEADER_ROWS + 1; row <= lastRow; row++) {
    var timeCell = sheet.getRange(row, 14);
    var statusCell = sheet.getRange(row, 16);
    var timeValue = timeCell.getValue();
    var statusValue = statusCell.getValue();
    if (!timeValue || timeValue.toString().trim() === "") continue;
    if (statusValue !== "Не начато") continue;
    try {
      var parsedTime = new Date(Utilities.formatDate(timeValue, timeZone, "yyyy-MM-dd HH:mm:ss"));
    } catch (err) { continue; }
    var diffMinutes = (now - parsedTime) / 60000;
    var newColor = null;
    if (diffMinutes >= 60) { newColor = "#FF0000"; }
    else if (diffMinutes >= 30) { newColor = "#FFD700"; }
    else if (diffMinutes >= 15) { newColor = "#32CD32"; }
    if (newColor) { timeCell.setBackground(newColor); }
  }
}

////////////////////////////
// Функции для формирования и отправки сообщений в Telegram
////////////////////////////

// Вспомогательные функции для форматирования дат
function formatDateField(date) {
  if (date instanceof Date) {
    return Utilities.formatDate(date, "GMT+10", "dd.MM.yyyy");
  }
  return date;
}
function formatTimeField(time) {
  if (time instanceof Date) {
    return Utilities.formatDate(time, "GMT+10", "HH:mm:ss");
  }
  return time;
}
function formatReactionTimeField(rt) {
  if (rt instanceof Date) {
    return Utilities.formatDate(rt, "GMT+10", "dd.MM.yyyy HH:mm");
  }
  return rt;
}
// Формирование сообщения заявки
function composeRequestMessage(data) {
  var lines = [];
  lines.push("==============================");
  lines.push("       *ЗАЯВКА № " + data.number + " ⚠*");
  lines.push("==============================");
  lines.push("");
  lines.push("🗓 *Дата (и время):* " + formatDateField(data.date) + " " + formatTimeField(data.time));
  lines.push("🏥 *Отделение:* " + data.department);
  lines.push("🔧 *Модель:* " + data.model);
  lines.push("🔢 *Заводской №:* " + data.serial);
  lines.push("📇 *Инв. №:* " + data.invNumber);
  lines.push("📍 *Место:* " + data.location);
  lines.push("💥 *Неисправность:* " + data.issueDesc);
  lines.push("☎️ *Телефон:* " + data.phone);
  lines.push("⚠️ *Приоритет:* " + data.priority);
  lines.push("🛠 *Вид ТО:* " + data.maintenance);
  lines.push("⏱ *Время реакции:* " + (data.reactionTime ? formatReactionTimeField(data.reactionTime) : "—"));
  if (data.executor) {
    lines.push("");
    lines.push("==============================");
    lines.push("*Назначен исполнитель:* " + data.executor);
  }
  if (data.status) {
    lines.push("*Статус заявки:* " + data.status);
  }
  return lines.join("\n");
}
// Отправка личного сообщения с кнопками для выбора исполнителя
function sendToTelegramWithButtons(message, row) {
  var token = PropertiesService.getScriptProperties().getProperty("TELEGRAM_BOT_TOKEN");
  var defaultChatId = PropertiesService.getScriptProperties().getProperty("TELEGRAM_CHAT_ID");
  var spesialChatId = PropertiesService.getScriptProperties().getProperty("TELEGRAM_GROUP_CHAT_ID");

  if (!token || !defaultChatId) {
    Logger.log("❌ Не найден TOKEN или CHAT_ID в Script Properties");
    return;
  }
  var keyboard = {
    inline_keyboard: [
      [
        { text: "Гиря А.Г", callback_data: "assign_engineer:" + row + ":Гиря А.Г" },
        { text: "Потеряйкин А.В.", callback_data: "assign_engineer:" + row + ":Потеряйкин А.В." },
        { text: "Демин Д.С.", callback_data: "assign_engineer:" + row + ":Демин Д.С." }
      ]
    ]
  };
  var payload = {
    chat_id: spesialChatId,
    text: message,
    parse_mode: "Markdown",
    reply_markup: JSON.stringify(keyboard)
  };
  var params = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  try {
    var response = UrlFetchApp.fetch("https://api.telegram.org/bot" + token + "/sendMessage", params);
    Logger.log("Telegram (with buttons) → " + response.getContentText());
  } catch (err) {
    Logger.log("❌ Ошибка при отправке (withButtons): " + err.message);
  }
}
// Отправка сообщения в группу с кнопками для обновления статуса
function sendToTelegramGroup(message, row) {
  var token = PropertiesService.getScriptProperties().getProperty("TELEGRAM_BOT_TOKEN");
  var groupChatId = PropertiesService.getScriptProperties().getProperty("TELEGRAM_GROUP_CHAT_ID");
  if (!token || !groupChatId) {
    Logger.log("❌ Не найден TOKEN или GROUP_CHAT_ID в Script Properties");
    return;
  }
  var keyboard = {
    inline_keyboard: [
      [
        { text: "Выполняется", callback_data: "update_status:" + row + ":Выполняется" },
        { text: "Завершено", callback_data: "update_status:" + row + ":Завершено" },
        { text: "Приостановлено", callback_data: "update_status:" + row + ":Приостановлено" }
      ]
    ]
  };
  var payload = {
    chat_id: groupChatId,
    text: message,
    parse_mode: "Markdown",
    reply_markup: JSON.stringify(keyboard)
  };
  var params = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  try {
    var response = UrlFetchApp.fetch("https://api.telegram.org/bot" + token + "/sendMessage", params);
    Logger.log("Telegram group message → " + response.getContentText());
  } catch (err) {
    Logger.log("❌ Ошибка при отправке в группу: " + err.message);
  }
}
// Считывание данных заявки из таблицы
function readRequestData(sheet, row) {
  var data = {};
  data.number       = sheet.getRange(row, 1).getValue();
  data.date         = sheet.getRange(row, 2).getValue();
  data.time         = sheet.getRange(row, 3).getValue();
  data.department   = sheet.getRange(row, 4).getValue();
  data.phone        = sheet.getRange(row, 5).getValue();
  data.location     = sheet.getRange(row, 6).getValue();
  data.model        = sheet.getRange(row, 7).getValue();
  data.serial       = sheet.getRange(row, 8).getValue();
  data.invNumber    = sheet.getRange(row, 9).getValue();
  data.releaseDate  = sheet.getRange(row,10).getValue();
  data.issueDesc    = sheet.getRange(row,11).getValue();
  data.maintenance  = sheet.getRange(row,12).getValue();
  data.priority     = sheet.getRange(row,13).getValue();
  data.reactionTime = sheet.getRange(row,14).getValue();
  data.executor     = sheet.getRange(row,15).getValue();
  data.status       = sheet.getRange(row,16).getValue();
  return data;
}
// Отправка заявки инженерам (личное сообщение)
function sendFormToEngineers(sheet, row) {
  var data = readRequestData(sheet, row);
  var message = composeRequestMessage(data);
  sendToTelegramWithButtons(message, row);
}
//Функции для форматирования колонки G
function handleColumnGFormatting(e) {
  var range = e.range;
  var originalText = e.value;
  Logger.log("Оригинальный текст в колонке G: " + originalText);
  var correctedText = checkSpelling(originalText);
  if (correctedText && correctedText !== originalText) {
    originalText = correctedText;
    Logger.log("После орфопроверки: " + originalText);
  }
  var formattedText = formatOnlyFirstLetter(originalText);
  Logger.log("После форматирования первой буквы: " + formattedText);
  range.setValue(formattedText);
}
function formatOnlyFirstLetter(text) {
  if (typeof text !== 'string' || text.length === 0) return text;
  return text.charAt(0).toUpperCase() + text.slice(1);
}

// Обработка callback запросов из Telegram (doPost)
// 1) assign_engineer: обновление исполнителя, редактирование личного сообщения и отправка сообщения в группу
// 2) update_status: обновление статуса заявки и редактирование сообщения в группе
function doPost(e) {
  var contents = JSON.parse(e.postData.contents);
  if (contents.callback_query) {
    var callbackQuery = contents.callback_query;
    var callbackData = callbackQuery.data;
    
    if (callbackData.indexOf("assign_engineer") === 0) {
      // Обработка назначения исполнителя
      var parts = callbackData.split(":");
      if (parts.length >= 3) {
        var row = parseInt(parts[1], 10);
        var engineerName = parts.slice(2).join(":");
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var sheet = ss.getActiveSheet();
        // Обновляем колонку O (исполнитель)
        sheet.getRange(row, 15).setValue(engineerName);
        var data = readRequestData(sheet, row);
        data.executor = engineerName;
        
        // Удаляем исходное личное сообщение после выбора исполнителя
        deleteTelegramMessage(callbackQuery.message.chat.id, callbackQuery.message.message_id);
        
        // Отправляем сообщение в группу с кнопками для обновления статуса
        var groupMessage = composeRequestMessage(data);
        sendToTelegramGroup(groupMessage, row);
        
        // Отвечаем на callback
        answerCallback(callbackQuery.id, "Исполнитель назначен: " + engineerName);
      }
    } else if (callbackData.indexOf("update_status") === 0) {
      // Обработка обновления статуса заявки
      var parts = callbackData.split(":");
      if (parts.length >= 3) {
        var row = parseInt(parts[1], 10);
        var newStatus = parts.slice(2).join(":");
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var sheet = ss.getActiveSheet();
        // Обновляем колонку P (статус заявки)
        sheet.getRange(row, 16).setValue(newStatus);
        var data = readRequestData(sheet, row);
        data.status = newStatus;
        var updatedGroupMessage = composeRequestMessage(data);
        // Редактируем сообщение в группе, сохраняя клавиатуру
        var originalKeyboard = callbackQuery.message.reply_markup;
        editTelegramMessage(callbackQuery.message.chat.id, callbackQuery.message.message_id, updatedGroupMessage, originalKeyboard);
        answerCallback(callbackQuery.id, "Статус заявки обновлён: " + newStatus);
      }
    }
  }
  return ContentService.createTextOutput("");
}

// Вспомогательная функция для редактирования сообщения в Telegram, сохраняя inline-клавиатуру
function editTelegramMessage(chatId, messageId, text, keyboard) {
  var token = PropertiesService.getScriptProperties().getProperty("TELEGRAM_BOT_TOKEN");
  var payload = {
    chat_id: chatId,
    message_id: messageId,
    text: text,
    parse_mode: "Markdown"
  };
  if (keyboard) {
    payload.reply_markup = typeof keyboard === 'string' ? keyboard : JSON.stringify(keyboard);
  }
  var params = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload)
  };
  UrlFetchApp.fetch("https://api.telegram.org/bot" + token + "/editMessageText", params);
}

// Вспомогательная функция для ответа на callback запросы
function answerCallback(callbackQueryId, text) {
  var token = PropertiesService.getScriptProperties().getProperty("TELEGRAM_BOT_TOKEN");
  var answerPayload = {
    callback_query_id: callbackQueryId,
    text: text
  };
  var answerParams = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(answerPayload)
  };
  UrlFetchApp.fetch("https://api.telegram.org/bot" + token + "/answerCallbackQuery", answerParams);
}

//Вспомогательная функция для удаления сообщения
function deleteTelegramMessage(chatId, messageId) {
  var token = PropertiesService.getScriptProperties().getProperty("TELEGRAM_BOT_TOKEN");
  var payload = {
    chat_id: chatId,
    message_id: messageId
  };
  var params = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload)
  };
  UrlFetchApp.fetch("https://api.telegram.org/bot" + token + "/deleteMessage", params);
}

// Функция проверки обязательных полей (столбцы D (4) до M (13))
function checkRequiredFields(sheet, row) {
  for (var col = 4; col <= 13; col++) {
    if (sheet.getRange(row, col).isBlank()) {
      return false;
    }
  }
  return true;
}

// Функция проверки обязательных полей (столбцы D (4) до M (13))
function checkRequiredFields(sheet, row) {
  for (var col = 4; col <= 13; col++) {
    if (sheet.getRange(row, col).isBlank()) {
      return false;
    }
  }
  return true;
}

// Обработка изменения приоритета (приоритет теперь в колонке M=13)
function handlePriorityChange(e, sheet, row) {
  // Проверяем, что все обязательные поля заполнены
  if (!checkRequiredFields(sheet, row)) {
    SpreadsheetApp.getUi().alert("Пожалуйста, заполните все поля от D до M перед изменением приоритета!");
    if (e.oldValue !== undefined) {
      sheet.getRange(row, 13).setValue(e.oldValue);
    } else {
      sheet.getRange(row, 13).clearContent();
    }
    return;
  }
  
  // Проверка флага отправки (например, столбец АА, номер 27)
  var flagCell = sheet.getRange(row, 27);
  if (!flagCell.isBlank()) {
    // Сообщение уже отправлено – выходим, чтобы не отправлять повторно
    return;
  }
  
  var priority = e.value;
  var requestDate = sheet.getRange(row, 2).getValue(); // B
  var requestTime = sheet.getRange(row, 3).getValue(); // C

  if (requestDate && requestTime && priority) {
    var requestDateTime = new Date(requestDate);
    if (requestTime instanceof Date) {
      requestDateTime.setHours(requestTime.getHours(), requestTime.getMinutes(), requestTime.getSeconds());
    }
    var deadline = calculateDeadline(requestDateTime, priority);
    sheet.getRange(row, 14).setValue(
      Utilities.formatDate(deadline, "GMT+10", "dd.MM.yyyy HH:mm")
    );
    // Устанавливаем статус «Не начато» в колонке P=16
    sheet.getRange(row, 16).setValue("Не начато");
    // Отправляем форму инженерам – передаём row для формирования callback_data
    sendFormToEngineers(sheet, row);
    
    // После успешной отправки ставим флаг (например, "отправлено")
    flagCell.setValue("отправлено");
  }
}

// Количество строк, занятых заголовками
var HEADER_ROWS = 2;

////////////////////////////
// Основной триггер редактирования
////////////////////////////

function onEditTrigger(e) {
  var range = e.range;
  var sheet = range.getSheet();
  var row = range.getRow();
  var col = range.getColumn();
  if (col === 16 && e.value && e.value.toString().toLowerCase() === "завершено") {
    sheet.getRange(row, 23).setValue(new Date());
  }
  
  // Пропускаем строки заголовков и пустые изменения
  if (row <= HEADER_ROWS || !e.value) return;

  // 1) При изменении приоритета (приоритет теперь в колонке M=13)
  if (col === 13) {
    handlePriorityChange(e, sheet, row);
  }
  // 2) При выборе подразделения (D=4)
  if (col === 4 && e.value !== "") {
    handleDepartmentSelection(sheet, row);
  }
  // 3) Форматирование текста для колонок [6,7,8,11,14,15,16,17,18,19]
  if (col === 7) {
    handleColumnGFormatting(e);
  } else if ([6,8,11,14,15,16,17,18,19].includes(col)) {
    handleTextFormatting(e);
  }
  // 4) Форматирование номера телефона (E=5)
  if (col === 5) {
    handlePhoneFormatting(e);
  }
  // 5) Обработка описания неисправности (K=11)
  if (col === 11) {
    handleIssueDescription(e, sheet, row);
  }
  // 6) Автоматическая смена статуса при выборе исполнителя (O=15, если редактируется вручную)
  if (col === 15) {
    handleExecutorSelection(sheet, row);
  }
  // 7) Исправление формата даты (J=10)
  if (col === 10) {
    handleDateFormatting(e);
  }
  // 8) Применяем стандартное форматирование к изменённой ячейке
  applyStandardFormatting(range);
}

////////////////////////////
// Функции обработки редактирования
////////////////////////////

// 1. Изменение приоритета (M=13)
function handlePriorityChange(e, sheet, row) {
  var priority = e.value;
  var requestDate = sheet.getRange(row, 2).getValue(); // B
  var requestTime = sheet.getRange(row, 3).getValue(); // C

  if (requestDate && requestTime && priority) {
    var requestDateTime = new Date(requestDate);
    if (requestTime instanceof Date) {
      requestDateTime.setHours(requestTime.getHours(), requestTime.getMinutes(), requestTime.getSeconds());
    }
    var deadline = calculateDeadline(requestDateTime, priority);
    sheet.getRange(row, 14).setValue(
      Utilities.formatDate(deadline, "GMT+10", "dd.MM.yyyy HH:mm")
    );
    // Устанавливаем статус «Не начато» в колонке P=16
    sheet.getRange(row, 16).setValue("Не начато");
    // Отправляем форму инженерам – передаём row для формирования callback_data
    sendFormToEngineers(sheet, row);
  }
}

// 2. Выбор подразделения (D=4)
function handleDepartmentSelection(sheet, row) {
  var numberCell = sheet.getRange(row, 1);
  var dateCell = sheet.getRange(row, 2);
  var timeCell = sheet.getRange(row, 3);

  if (numberCell.isBlank()) {
    var timestamp = new Date();
    var formattedNumber = Utilities.formatDate(timestamp, "GMT+10", "yyyyMMdd-HHmmss");
    var formattedDate = Utilities.formatDate(timestamp, "GMT+10", "dd.MM.yyyy");
    var formattedTime = Utilities.formatDate(timestamp, "GMT+10", "HH:mm:ss");

    numberCell.setValue(formattedNumber);
    dateCell.setValue(formattedDate);
    timeCell.setValue(formattedTime);
  }
}

// 3. Стандартное форматирование текста (с орфопроверкой)
function handleTextFormatting(e) {
  var range = e.range;
  var newText = e.value;
  var columnsToCheck = [11, 17, 18, 19];
  if (columnsToCheck.includes(e.range.getColumn())) {
    var correctedText = checkSpelling(newText);
    if (correctedText && correctedText !== newText) {
      newText = correctedText;
    }
  }
  newText = newText.toLowerCase().replace(/(^|[.!?]\s*)([а-яёa-z])/g, function(match, sep, char) {
    return sep + char.toUpperCase();
  });
  range.setValue(newText);
}

// 4. Форматирование номера телефона (E=5)
function handlePhoneFormatting(e) {
  var range = e.range;
  var phone = e.value.replace(/\D/g, "");
  if (phone.length === 11 && (phone.startsWith("8") || phone.startsWith("7"))) {
    phone = "+7(" + phone.slice(1, 4) + ") " + phone.slice(4, 7) + "-" + phone.slice(7, 9) + "-" + phone.slice(9, 11);
    range.setValue(phone);
  } else {
    SpreadsheetApp.getUi().alert("⚠️ Введите номер телефона в формате 89990809578 (11 цифр).");
  }
}

// 5. Обработка описания неисправности (K=11)
function handleIssueDescription(e, sheet, row) {
  var textK = e.value.toLowerCase();
  var statusCell = sheet.getRange(row, 12); // статус можно обновлять в данной колонке (при необходимости измените)
  var regexDiagnostic = /(полом|неисправ|сломал|треснул|вибраци|не работает|не включается|горит лампа|замыкан|стучит)/;
  var regexMontage = /(перемещ|монтаж|демонтаж|налад|перестанов|разбор|установ|перенести|занести|вынести)/;
  
  if (regexDiagnostic.test(textK)) {
    statusCell.setValue("Техническое диагностирование");
  } else if (regexMontage.test(textK)) {
    statusCell.setValue("Монтаж/демонтаж или наладка");
  }
  
  // Автозаполнение колонок H и I, если данных нет
  var hCell = sheet.getRange(row, 8);
  var iCell = sheet.getRange(row, 9);
  if (hCell.isBlank()) {
    hCell.setValue("Информация отсутствует");
  }
  if (iCell.isBlank()) {
    iCell.setValue("Информация отсутствует");
  }
}

// 6. Обработка выбора исполнителя (O=15, если редактируется вручную)
function handleExecutorSelection(sheet, row) {
  var statusCell = sheet.getRange(row, 16); // P=16
  var currentStatus = statusCell.getValue();
  if (!currentStatus || currentStatus === "Не начато") {
    statusCell.setValue("Выполняется");
  }
}

// 7. Исправление формата даты (J=10)
function handleDateFormatting(e) {
  var range = e.range;
  var rawInput = e.value.toString();
  var fixed = fixDateFormat(rawInput);
  if (fixed) {
    var parts = fixed.split(".");
    if (parts.length === 3) {
      var dd = parseInt(parts[0], 10);
      var mm = parseInt(parts[1], 10) - 1;
      var yyyy = parseInt(parts[2], 10);
      var parsedDate = new Date(yyyy, mm, dd);
      if (!isNaN(parsedDate.getTime())) {
        range.setValue(parsedDate);
        range.setNumberFormat("dd.MM.yyyy");
        return;
      }
    }
    range.setValue(fixed);
    range.setNumberFormat("@STRING@");
  }
}

// 8. Применение стандартного форматирования к ячейке
function applyStandardFormatting(range) {
  range.setFontFamily("Roboto");
  range.setFontSize(8);
  range.setFontColor("#000000");
  range.setHorizontalAlignment("left");
  range.setVerticalAlignment("middle");
  range.setWrap(true);
}

////////////////////////////
// Функции форматирования, проверки и расчёта
////////////////////////////

// Функция «умного» исправления формата даты
function fixDateFormat(input) {
  try {
    var str = input.toString()
      .replace(/[^0-9\/.,-]/g, '')
      .replace(/,/g, '.')
      .replace(/-/g, '.')
      .replace(/\//g, '.');
    var parts = str.split('.').filter(function(p){ return p.length > 0; });
    var dd = '01', mm = '01', yyyy = '';
    if (parts.length === 1 && parts[0].length === 4) {
      yyyy = parts[0];
    } else if (parts.length === 2) {
      if (parts[1].length === 4) {
        mm = parts[0];
        yyyy = parts[1];
      } else {
        mm = parts[0];
        yyyy = '20' + parts[1].padStart(2, '0');
      }
    } else if (parts.length === 3) {
      dd = parts[0];
      mm = parts[1];
      yyyy = parts[2];
      if (yyyy.length === 2) {
        yyyy = '20' + yyyy;
      }
    }
    dd = dd.padStart(2, '0').slice(-2);
    mm = mm.padStart(2, '0').slice(-2);
    yyyy = yyyy.padStart(4, '20');
    return dd + '.' + mm + '.' + yyyy;
  } catch(e) {
    return input.toString();
  }
}
// Функция проверки орфографии
function checkSpelling(text) {
  var url = 'https://speller.yandex.net/services/spellservice.json/checkText';
  var response = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/x-www-form-urlencoded',
    payload: { text: text, lang: 'ru', format: 'plain' },
    muteHttpExceptions: true
  });
  var corrections = JSON.parse(response.getContentText());
  if (corrections.length === 0) return text;
  corrections.forEach(function(correction) {
    text = text.replace(correction.word, correction.s[0]);
  });
  return text;
}
// Функция проверки рабочего времени инженеров
function isWorkingTime(date) {
  var day = date.getDay();
  var hours = date.getHours();
  var minutes = date.getMinutes();
  var timeInMinutes = hours * 60 + minutes;
  var workStart = 8 * 60 + 30;  // 08:30
  var workEnd = 17 * 60;        // 17:00
  var holidays = ['01.01','07.01','23.02','08.03','01.05','09.05','12.06','04.11'];
  var formattedDate = Utilities.formatDate(date, "GMT+10", "dd.MM");
  if (holidays.includes(formattedDate)) return false;
  return day >= 1 && day <= 5 && timeInMinutes >= workStart && timeInMinutes < workEnd;
}
// Функция расчёта дедлайна реакции
function calculateDeadline(startDate, priority) {
  var reactionMinutes;
  switch(priority) {
    case 'Низкий': reactionMinutes = 8 * 60; break;
    case 'Средний': reactionMinutes = 6 * 60; break;
    case 'Высокий': reactionMinutes = 2 * 60; break;
    case 'Экстренный': reactionMinutes = 15; break;
    default: reactionMinutes = 8 * 60; break;
  }
  var deadline = new Date(startDate);
  if (priority === 'Экстренный' && !isWorkingTime(startDate)) {
    deadline.setMinutes(deadline.getMinutes() + 15 + 60);
    return deadline;
  }
  while (reactionMinutes > 0) {
    if (isWorkingTime(deadline)) {
      deadline.setMinutes(deadline.getMinutes() + 1);
      reactionMinutes--;
    } else {
      deadline.setMinutes(deadline.getMinutes() + 1);
      while (!isWorkingTime(deadline)) {
        deadline.setMinutes(deadline.getMinutes() + 1);
      }
    }
  }
  return deadline;
}
// Функция проверки времени реакции и изменения цвета ячейки (тестовые пороги)
function checkResponseTime() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var timeZone = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  var lastRow = sheet.getLastRow();
  var now = new Date();
  for (var row = HEADER_ROWS + 1; row <= lastRow; row++) {
    var timeCell = sheet.getRange(row, 14);
    var statusCell = sheet.getRange(row, 16);
    var timeValue = timeCell.getValue();
    var statusValue = statusCell.getValue();
    if (!timeValue || timeValue.toString().trim() === "") continue;
    if (statusValue !== "Не начато") continue;
    try {
      var parsedTime = new Date(Utilities.formatDate(timeValue, timeZone, "yyyy-MM-dd HH:mm:ss"));
    } catch (err) { continue; }
    var diffMinutes = (now - parsedTime) / 60000;
    var newColor = null;
    if (diffMinutes >= 60) { newColor = "#FF0000"; }
    else if (diffMinutes >= 30) { newColor = "#FFD700"; }
    else if (diffMinutes >= 15) { newColor = "#32CD32"; }
    if (newColor) { timeCell.setBackground(newColor); }
  }
}

////////////////////////////
// Функции для формирования и отправки сообщений в Telegram
////////////////////////////

// Вспомогательные функции для форматирования дат
function formatDateField(date) {
  if (date instanceof Date) {
    return Utilities.formatDate(date, "GMT+10", "dd.MM.yyyy");
  }
  return date;
}
function formatTimeField(time) {
  if (time instanceof Date) {
    return Utilities.formatDate(time, "GMT+10", "HH:mm:ss");
  }
  return time;
}
function formatReactionTimeField(rt) {
  if (rt instanceof Date) {
    return Utilities.formatDate(rt, "GMT+10", "dd.MM.yyyy HH:mm");
  }
  return rt;
}
// Формирование сообщения заявки
function composeRequestMessage(data) {
  var lines = [];
  lines.push("==============================");
  lines.push("       *ЗАЯВКА № " + data.number + " ⚠*");
  lines.push("==============================");
  lines.push("");
  lines.push("🗓 *Дата (и время):* " + formatDateField(data.date) + " " + formatTimeField(data.time));
  lines.push("🏥 *Отделение:* " + data.department);
  lines.push("🔧 *Модель:* " + data.model);
  lines.push("🔢 *Заводской №:* " + data.serial);
  lines.push("📇 *Инв. №:* " + data.invNumber);
  lines.push("📍 *Место:* " + data.location);
  lines.push("💥 *Неисправность:* " + data.issueDesc);
  lines.push("☎️ *Телефон:* " + data.phone);
  lines.push("⚠️ *Приоритет:* " + data.priority);
  lines.push("🛠 *Вид ТО:* " + data.maintenance);
  lines.push("⏱ *Время реакции:* " + (data.reactionTime ? formatReactionTimeField(data.reactionTime) : "—"));
  if (data.executor) {
    lines.push("");
    lines.push("==============================");
    lines.push("*Назначен исполнитель:* " + data.executor);
  }
  if (data.status) {
    lines.push("*Статус заявки:* " + data.status);
  }
  return lines.join("\n");
}
// Отправка личного сообщения с кнопками для выбора исполнителя
function sendToTelegramWithButtons(message, row) {
  var token = PropertiesService.getScriptProperties().getProperty("TELEGRAM_BOT_TOKEN");
  var defaultChatId = PropertiesService.getScriptProperties().getProperty("TELEGRAM_CHAT_ID");
  var spesialChatId = PropertiesService.getScriptProperties().getProperty("TELEGRAM_GROUP_CHAT_ID");

  if (!token || !defaultChatId) {
    Logger.log("❌ Не найден TOKEN или CHAT_ID в Script Properties");
    return;
  }
  var keyboard = {
    inline_keyboard: [
      [
        { text: "Гиря А.Г", callback_data: "assign_engineer:" + row + ":Гиря А.Г" },
        { text: "Потеряйкин А.В.", callback_data: "assign_engineer:" + row + ":Потеряйкин А.В." },
        { text: "Демин Д.С.", callback_data: "assign_engineer:" + row + ":Демин Д.С." }
      ]
    ]
  };
  var payload = {
    chat_id: spesialChatId,
    text: message,
    parse_mode: "Markdown",
    reply_markup: JSON.stringify(keyboard)
  };
  var params = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  try {
    var response = UrlFetchApp.fetch("https://api.telegram.org/bot" + token + "/sendMessage", params);
    Logger.log("Telegram (with buttons) → " + response.getContentText());
  } catch (err) {
    Logger.log("❌ Ошибка при отправке (withButtons): " + err.message);
  }
}
// Отправка сообщения в группу с кнопками для обновления статуса
function sendToTelegramGroup(message, row) {
  var token = PropertiesService.getScriptProperties().getProperty("TELEGRAM_BOT_TOKEN");
  var groupChatId = PropertiesService.getScriptProperties().getProperty("TELEGRAM_GROUP_CHAT_ID");
  if (!token || !groupChatId) {
    Logger.log("❌ Не найден TOKEN или GROUP_CHAT_ID в Script Properties");
    return;
  }
  var keyboard = {
    inline_keyboard: [
      [
        { text: "Выполняется", callback_data: "update_status:" + row + ":Выполняется" },
        { text: "Завершено", callback_data: "update_status:" + row + ":Завершено" },
        { text: "Приостановлено", callback_data: "update_status:" + row + ":Приостановлено" }
      ]
    ]
  };
  var payload = {
    chat_id: groupChatId,
    text: message,
    parse_mode: "Markdown",
    reply_markup: JSON.stringify(keyboard)
  };
  var params = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  try {
    var response = UrlFetchApp.fetch("https://api.telegram.org/bot" + token + "/sendMessage", params);
    Logger.log("Telegram group message → " + response.getContentText());
  } catch (err) {
    Logger.log("❌ Ошибка при отправке в группу: " + err.message);
  }
}
// Считывание данных заявки из таблицы
function readRequestData(sheet, row) {
  var data = {};
  data.number       = sheet.getRange(row, 1).getValue();
  data.date         = sheet.getRange(row, 2).getValue();
  data.time         = sheet.getRange(row, 3).getValue();
  data.department   = sheet.getRange(row, 4).getValue();
  data.phone        = sheet.getRange(row, 5).getValue();
  data.location     = sheet.getRange(row, 6).getValue();
  data.model        = sheet.getRange(row, 7).getValue();
  data.serial       = sheet.getRange(row, 8).getValue();
  data.invNumber    = sheet.getRange(row, 9).getValue();
  data.releaseDate  = sheet.getRange(row,10).getValue();
  data.issueDesc    = sheet.getRange(row,11).getValue();
  data.maintenance  = sheet.getRange(row,12).getValue();
  data.priority     = sheet.getRange(row,13).getValue();
  data.reactionTime = sheet.getRange(row,14).getValue();
  data.executor     = sheet.getRange(row,15).getValue();
  data.status       = sheet.getRange(row,16).getValue();
  return data;
}
// Отправка заявки инженерам (личное сообщение)
function sendFormToEngineers(sheet, row) {
  var data = readRequestData(sheet, row);
  var message = composeRequestMessage(data);
  sendToTelegramWithButtons(message, row);
}
//Функции для форматирования колонки G
function handleColumnGFormatting(e) {
  var range = e.range;
  var originalText = e.value;
  Logger.log("Оригинальный текст в колонке G: " + originalText);
  var correctedText = checkSpelling(originalText);
  if (correctedText && correctedText !== originalText) {
    originalText = correctedText;
    Logger.log("После орфопроверки: " + originalText);
  }
  var formattedText = formatOnlyFirstLetter(originalText);
  Logger.log("После форматирования первой буквы: " + formattedText);
  range.setValue(formattedText);
}
function formatOnlyFirstLetter(text) {
  if (typeof text !== 'string' || text.length === 0) return text;
  return text.charAt(0).toUpperCase() + text.slice(1);
}

// Обработка callback запросов из Telegram (doPost)
// 1) assign_engineer: обновление исполнителя, редактирование личного сообщения и отправка сообщения в группу
// 2) update_status: обновление статуса заявки и редактирование сообщения в группе
function doPost(e) {
  var contents = JSON.parse(e.postData.contents);
  if (contents.callback_query) {
    var callbackQuery = contents.callback_query;
    var callbackData = callbackQuery.data;
    
    if (callbackData.indexOf("assign_engineer") === 0) {
      // Обработка назначения исполнителя
      var parts = callbackData.split(":");
      if (parts.length >= 3) {
        var row = parseInt(parts[1], 10);
        var engineerName = parts.slice(2).join(":");
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var sheet = ss.getActiveSheet();
        // Обновляем колонку O (исполнитель)
        sheet.getRange(row, 15).setValue(engineerName);
        var data = readRequestData(sheet, row);
        data.executor = engineerName;
        
        // Удаляем исходное личное сообщение после выбора исполнителя
        deleteTelegramMessage(callbackQuery.message.chat.id, callbackQuery.message.message_id);
        
        // Отправляем сообщение в группу с кнопками для обновления статуса
        var groupMessage = composeRequestMessage(data);
        sendToTelegramGroup(groupMessage, row);
        
        // Отвечаем на callback
        answerCallback(callbackQuery.id, "Исполнитель назначен: " + engineerName);
      }
    } else if (callbackData.indexOf("update_status") === 0) {
      // Обработка обновления статуса заявки
      var parts = callbackData.split(":");
      if (parts.length >= 3) {
        var row = parseInt(parts[1], 10);
        var newStatus = parts.slice(2).join(":");
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var sheet = ss.getActiveSheet();
        // Обновляем колонку P (статус заявки)
        sheet.getRange(row, 16).setValue(newStatus);
        var data = readRequestData(sheet, row);
        data.status = newStatus;
        var updatedGroupMessage = composeRequestMessage(data);
        // Редактируем сообщение в группе, сохраняя клавиатуру
        var originalKeyboard = callbackQuery.message.reply_markup;
        editTelegramMessage(callbackQuery.message.chat.id, callbackQuery.message.message_id, updatedGroupMessage, originalKeyboard);
        answerCallback(callbackQuery.id, "Статус заявки обновлён: " + newStatus);
      }
    }
  }
  return ContentService.createTextOutput("");
}

// Вспомогательная функция для редактирования сообщения в Telegram, сохраняя inline-клавиатуру
function editTelegramMessage(chatId, messageId, text, keyboard) {
  var token = PropertiesService.getScriptProperties().getProperty("TELEGRAM_BOT_TOKEN");
  var payload = {
    chat_id: chatId,
    message_id: messageId,
    text: text,
    parse_mode: "Markdown"
  };
  if (keyboard) {
    payload.reply_markup = typeof keyboard === 'string' ? keyboard : JSON.stringify(keyboard);
  }
  var params = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload)
  };
  UrlFetchApp.fetch("https://api.telegram.org/bot" + token + "/editMessageText", params);
}

// Вспомогательная функция для ответа на callback запросы
function answerCallback(callbackQueryId, text) {
  var token = PropertiesService.getScriptProperties().getProperty("TELEGRAM_BOT_TOKEN");
  var answerPayload = {
    callback_query_id: callbackQueryId,
    text: text
  };
  var answerParams = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(answerPayload)
  };
  UrlFetchApp.fetch("https://api.telegram.org/bot" + token + "/answerCallbackQuery", answerParams);
}

//Вспомогательная функция для удаления сообщения
function deleteTelegramMessage(chatId, messageId) {
  var token = PropertiesService.getScriptProperties().getProperty("TELEGRAM_BOT_TOKEN");
  var payload = {
    chat_id: chatId,
    message_id: messageId
  };
  var params = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload)
  };
  UrlFetchApp.fetch("https://api.telegram.org/bot" + token + "/deleteMessage", params);
}

// Функция проверки обязательных полей (столбцы D (4) до M (13))
function checkRequiredFields(sheet, row) {
  for (var col = 4; col <= 13; col++) {
    if (sheet.getRange(row, col).isBlank()) {
      return false;
    }
  }
  return true;
}

// Функция проверки обязательных полей (столбцы D (4) до M (13))
function checkRequiredFields(sheet, row) {
  for (var col = 4; col <= 13; col++) {
    if (sheet.getRange(row, col).isBlank()) {
      return false;
    }
  }
  return true;
}

// Обработка изменения приоритета (приоритет теперь в колонке M=13)
function handlePriorityChange(e, sheet, row) {
  // Проверяем, что все обязательные поля заполнены
  if (!checkRequiredFields(sheet, row)) {
    SpreadsheetApp.getUi().alert("Пожалуйста, заполните все поля от D до M перед изменением приоритета!");
    if (e.oldValue !== undefined) {
      sheet.getRange(row, 13).setValue(e.oldValue);
    } else {
      sheet.getRange(row, 13).clearContent();
    }
    return;
  }
  
  // Проверка флага отправки (например, столбец АА, номер 27)
  var flagCell = sheet.getRange(row, 27);
  if (!flagCell.isBlank()) {
    // Сообщение уже отправлено – выходим, чтобы не отправлять повторно
    return;
  }
  
  var priority = e.value;
  var requestDate = sheet.getRange(row, 2).getValue(); // B
  var requestTime = sheet.getRange(row, 3).getValue(); // C

  if (requestDate && requestTime && priority) {
    var requestDateTime = new Date(requestDate);
    if (requestTime instanceof Date) {
      requestDateTime.setHours(requestTime.getHours(), requestTime.getMinutes(), requestTime.getSeconds());
    }
    var deadline = calculateDeadline(requestDateTime, priority);
    sheet.getRange(row, 14).setValue(
      Utilities.formatDate(deadline, "GMT+10", "dd.MM.yyyy HH:mm")
    );
    // Устанавливаем статус «Не начато» в колонке P=16
    sheet.getRange(row, 16).setValue("Не начато");
    // Отправляем форму инженерам – передаём row для формирования callback_data
    sendFormToEngineers(sheet, row);
    
    // После успешной отправки ставим флаг (например, "отправлено")
    flagCell.setValue("отправлено");
  }
}
