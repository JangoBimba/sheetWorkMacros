function transferDataByDate() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const sourceSheet = ss.getSheetByName("План по дням");
    if (!sourceSheet) {
        throw new Error('Лист "План по дням" не найден. Проверьте имя листа.');
    }

    const targetSheet = ss.getSheetByName("Вечерний отчет");
    if (!targetSheet) {
        throw new Error('Лист "Вечерний отчет" не найден. Проверьте имя листа.');
    }

    const ui = SpreadsheetApp.getUi();
    const response = ui.prompt('Введите дату', 'Введите дату в формате ГГГГ-ММ-ДД:', ui.ButtonSet.OK_CANCEL);

    if (response.getSelectedButton() !== ui.Button.OK) {
        return;
    }

    const inputDate = response.getResponseText();
    const userDate = new Date(inputDate);
    if (isNaN(userDate.getTime())) {
        ui.alert('Ошибка', 'Неверный формат даты. Используйте формат ГГГГ-ММ-ДД.', ui.ButtonSet.OK);
        return;
    }
    userDate.setHours(0, 0, 0, 0);

    const rangesToClear = ["H2:H26", "I2:I26", "D2:D26", "G2:G26", "K2:K26"];
    for (let i = 0; i < rangesToClear.length; i++) {
        const range = targetSheet.getRange(rangesToClear[i]);
        if (range) {
            range.clearContent();
        }
    }

    const dateRow = sourceSheet.getRange(2, 3, 1, sourceSheet.getLastColumn() - 2).getValues()[0];
    const dataRange = sourceSheet.getRange(4, 3, 25, sourceSheet.getLastColumn() - 2).getValues();

    let hasEmptyCells = false;
    for (let col = 0; col < dateRow.length; col++) {
        const rowDate = new Date(dateRow[col]);
        rowDate.setHours(0, 0, 0, 0);

        if (rowDate.getTime() === userDate.getTime()) {
            for (let row = 0; row < dataRange.length; row++) {
                const value = dataRange[row][col];
                if (!value || value.toString().trim() === "") {
                    hasEmptyCells = true;
                    break;
                }
            }
        }
        if (hasEmptyCells) break;
    }

    if (hasEmptyCells) {
        ui.alert('Ошибка', 'В исходных данных найдены пустые ячейки. Заполните их и повторите попытку.', ui.ButtonSet.OK);
        return;
    }

    const dataToTransfer = [];

    for (let col = 0; col < dateRow.length; col++) {
        const rowDate = new Date(dateRow[col]);
        rowDate.setHours(0, 0, 0, 0);

        if (rowDate.getTime() === userDate.getTime()) {
            for (let row = 0; row < dataRange.length; row++) {
                const value = dataRange[row][col];
                if (value && value.toString().trim() !== "") {
                    dataToTransfer.push([value]);
                }
            }
        }
    }

    if (dataToTransfer.length > 0) {
        targetSheet.getRange(2, 9, dataToTransfer.length, 1).setValues(dataToTransfer);
    } else {
        ui.alert('Информация', 'Нет данных для переноса на указанную дату.', ui.ButtonSet.OK);
    }
}