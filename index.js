function handleSheet(sheet) {
    clearCols(sheet, 'delete');
    let handleArr = findAllColMarksFor(sheet, 'handle');
    handleArr.forEach((colMark, i) => {
        //处理一列数据
        handleCol(sheet, colMark);
        //移动
        clipCol(sheet, getColMark('C', i), colMark);
    });
    resetRange(sheet);
    return sheet;
}


//最早的"向上移"的需求
function moveValue(sheet) {
    let maxRow = getMaxRowNum(sheet);
    iterateAllCell(sheet, (key, c, r) => {
        if (sheet[key] && sheet[key].v === 0) {
            let currentR = r;
            let nextR = r + 1;
            while (nextR <= maxRow) {
                if (sheet[c + nextR]) {
                    sheet[c + currentR] = sheet[c + nextR];
                    currentR = nextR;
                }
                nextR++;
                if (currentR === maxRow) {
                    delete sheet[c + currentR]
                }
            }
        }
    });
}

function handleCol(sheet, colMark) {
    let keys = getCol(sheet, colMark).keys();
    let lastValueKey = keys[keys.length - 1];
    let lastValueRow = seperate(lastValueKey).r;
    //填0
    for (let i = 1; i < lastValueRow; i++) {
        if (!sheet[colMark + i]) {
            sheet[colMark + i] = {v: 0}
        }
    }
    //删除最后一个值
    delete sheet[lastValueKey];
}

function test() {
    // let resultWb = XLSX.utils.book_new();
    //
    // let resultSheet = {
    //     A1: {
    //         v: 1
    //     }
    // };
    // resultSheet['!ref'] = "A1:A1";
    //
    // XLSX.utils.book_append_sheet(resultWb, resultSheet, 'Sheet1');
    // XLSX.writeFile(resultWb, 'result.xlsx');


}
