//js通用方法和表格通用方法

//js通用
const CODE_A = 'A'.charCodeAt(0);
const CODE_Z = 'Z'.charCodeAt(0);

function isNum(string) {
    return !isNaN(Number.parseInt(string));
}

//getColMark('A',1) ===> B
//getColMark('Z',1) ===> A
function getColMark(colMark, num) {
    if (colMark.length === 2) {
        let chars = colMark.toArray();
    }
    return String.fromCharCode(colMark.charCodeAt(0) + num);
}

//列的序号转为字母
//只考虑到'ZZ'的情况
function num2c(num) {

    let z = Math.floor(num / 26);
    let y = num % 26;

    let letter = '';
    if (z !== 0) {
        letter += String.fromCharCode(CODE_A - 1 + z);
    }
    letter += String.fromCharCode(CODE_A - 1 + y);
    return letter;
}

//列的字母转为序号
function c2num(c) {
    let num = 0;
    if (c.length === 2) {
        num += (letterNum(c.charAt(0))) * 10;
    }
    num += letterNum(c.charAt(1));
    return num;
}

//得到是第几个字母
function letterNum(letter) {
    return letter.charCodeAt(0) - 'A'.charCodeAt(0) + 1;
}

//得到序号对应的字母
function numLetter(num) {

}

//表格通用

//得到最右下角的cell key
function getMaxKey(sheet) {
    if (sheet['!ref']) {
        return sheet['!ref'].split(':')[1];
    }
}

//设置...
function setMaxKey(sheet, maxKey) {
    sheet['!ref'] = `A1:${maxKey}`
}

function getMaxColMark(sheet) {
    return seperate(getMaxKey(sheet)).c;
}

function getMaxRowNum(sheet) {
    return seperate(getMaxKey(sheet)).r;
}

//分离数字和字母
function seperate(cellKey) {
    for (var i = 0; i < cellKey.length; i++) {
        if (isNum(cellKey[i])) {
            return {
                c: cellKey.slice(0, i),
                r: Number.parseInt(cellKey.slice(i))
            }

        }
    }
    return null;
}


//删除特定一列
function deleteCol(sheet, colMark) {

}

//删除整列
function deleteCols(sheet, colName) {

}

/**
 * 将一个cell的值(剪切)移动到另个一cell上
 * @param from : from cell key
 * @param to : to cell key
 */
function clipCell(sheet, to, from) {
    sheet[to] = sheet[from];
    delete sheet[from];
}

function copyCell(sheet, to, from) {
    sheet[to] = sheet[from];
}

function getCellValue(sheet, cellMark) {
    if (sheet.hasValue(cellMark)) {
        return sheet[cellMark].v;
    } else {
        return null;
    }
}


function deleteRow(sheet, rowMark) {

}

//清空某一整列的值
function clearCol(sheet, colMark) {
    iterateCol(sheet, colMark, (key, c, r) => {
        delete sheet[key];
    })
}

//清空所有标题为colName的列的值
function clearCols(sheet, colName) {
    let arr = findAllColMarksFor(sheet, colName);
    arr.forEach(v => {
        clearCol(sheet, v);
    })
}


//遍历整个sheet的所有cell
function iterateAllCell(sheet, func) {

    if (sheet && sheet.keys().length > 0) {
        sheet.keys().forEach(key => {
            let sep = seperate(key);
            if (sep && func) {
                func(key, sep.c, sep.r);
            }
        });
    }

}

//遍历一列的所有cell
function iterateCol(sheet, colMark, func) {
    iterateAllCell(sheet, (key, c, r) => {
        if (c === colMark && func) {
            func(key, c, r)
        }
    })
}

function iterateRow(sheet, rowNum, func) {
    iterateAllCell(sheet, (key, c, r) => {
        if (r === rowNum && func) {
            func(key, c, r);
        }
    })
}

function getCol(sheet, colMark) {
    let col = {};
    iterateCol(sheet, colMark, (key, c, r) => {
        col[key] = sheet[key];
    });
    return col;
}

function getRow(sheet, rowNum) {

}


//剪切移动一整列
function clipCol(sheet, targetColMark, sourceColMark) {

    clearCol(sheet, targetColMark);
    iterateCol(sheet, sourceColMark, (key, c, r) => {
        sheet[targetColMark + r] = sheet[sourceColMark + r];
        delete sheet[sourceColMark + r];
    });

}


function setMaxColMark(sheet, maxColMark) {
    let maxRow = getMaxRowNum(sheet);
    let maxKey = maxColMark + maxRow;
    setMaxKey(sheet, maxKey);
}

function changeColRange(sheet, changeNum) {
    setMaxColMark(sheet, getColMark(getMaxColMark(sheet), changeNum));
}

function findAllColMarksFor(sheet, colName) {
    let colMarks = [];
    iterateRow(sheet, 1, (key, c, r) => {
        if (sheet[key] && colName === sheet[key].v) {
            colMarks.push(c);
        }
    });
    return colMarks;
}

function resetRange(sheet) {
    let colNum = 0;
    let rowNum = 0;
    iterateRow(sheet, 1, () => {
        colNum++;
    });
    iterateCol(sheet, 'A', () => {
        rowNum++;
    });
    let maxColMark = getColMark('A', colNum - 1);
    setMaxKey(sheet, maxColMark + rowNum);
}

