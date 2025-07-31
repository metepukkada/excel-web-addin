function flattenObject(obj, parentKey = "", result = {}) {
    for (const [key, value] of Object.entries(obj)) {
        const newKey = parentKey ? `${parentKey}.${key}` : key;

        if (value && typeof value === "object" && !Array.isArray(value)) {
            flattenObject(value, newKey, result);
        } else {
            result[newKey] = value;
        }
    }
    return result;
}

function prepareData(dataArray) {
    const flatDataArray = dataArray.map(item => flattenObject(item));

    const allKeysSet = new Set();
    flatDataArray.forEach(item => {
        Object.keys(item).forEach(key => allKeysSet.add(key));
    });

    const allKeys = Array.from(allKeysSet);

    const values = flatDataArray.map(item =>
        allKeys.map(key => (item[key] !== undefined ? item[key] : ""))
    );

    return { headers: allKeys, values };
}

export async function writeJsonToExcel(data) {
    let dataArray;

    if (Array.isArray(data)) {
        if (data.length === 0) throw new Error("Veri boş ya da uygun formatta değil.");
        dataArray = data;
    } else if (data && typeof data === "object") {
        dataArray = [data];
    } else {
        throw new Error("Veri boş ya da uygun formatta değil.");
    }

    const { headers, values } = prepareData(dataArray);

    await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();

        const selection = context.workbook.getActiveCell();
        selection.load(["rowIndex", "columnIndex"]);
        await context.sync();

        const startRow = selection.rowIndex;
        const startCol = selection.columnIndex;

        sheet.getRangeByIndexes(startRow, startCol, 1, headers.length).values = [headers];
        sheet.getRangeByIndexes(startRow + 1, startCol, values.length, headers.length).values = values;

        for (let i = 0; i < headers.length; i++) {
            sheet.getRangeByIndexes(startRow, startCol + i, values.length + 1, 1).format.autofitColumns();
        }

        await context.sync();
    });
}


