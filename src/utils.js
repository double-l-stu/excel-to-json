/*
 * @Author: yunmin
 * @Email: 362279869@qq.com
 * @Date: 2020-02-11 16:00:54
 */

const Utils = module.exports;

// 字母A的ASCII值与1的偏移
const ascii_offset = 64;
const twenty_six_system = 26; 

function unfix_col(letter) {
    return letter.replace(/^\$([A-Z])/, "$1");
}

/**
 * getSheetRange  获取表单行列的起点终点范围
 *
 * @param {String} ref the range string
 * @return {Object} {start: {x: 0, y: 0}, end: {x: 0, y: 0}
 */
Utils.getSheetRange = function (ref) {
    let o = {
        start: {x: 0, y: 0},
        end: {x: 0, y: 0}
    };
    let idx, i = 0, cc = 0;
    let len = ref.length;
    for (idx = 0; i < len; ++i) {
        if ((cc = ref.charCodeAt(i) - 64) < 1 || cc > 26) {
            break;
        }
        idx = 26 * idx + cc;
    }
    o.start.x = --idx;

    for (idx = 0; i < len; ++i) {
        if ((cc = ref.charCodeAt(i) - 48) < 0 || cc > 9) {
            break;
        }
        idx = 10 * idx + cc;
    }
    o.start.y = --idx;

    if (i === len || ref.charCodeAt(++i) === 58) {
        o.end.x = o.start.x;
        o.end.y = o.start.y;
        return o;
    }

    for (idx = 0; i !== len; ++i) {
        if ((cc = ref.charCodeAt(i) - 64) < 1 || cc > 26) {
            break;
        }
        idx = 26 * idx + cc;
    }
    o.end.x = --idx;

    for (idx = 0; i !== len; ++i) {
        if ((cc = ref.charCodeAt(i) - 48) < 0 || cc > 9) {
            break;
        }
        idx = 10 * idx + cc;
    }
    o.end.y = --idx;
    return o;
};


/**
 * Converts column letters(A-Z) to column numbers(0-X)
 *
 * Instructions: A=>0, B=>1...
 *
 * @param {String} letter Column letter(A...Z...)
 * @return {Number} 0...n
 */
Utils.letter2Col = function (letter) {
    let c = unfix_col(letter), d = 0, i = 0;
    for (; i !== c.length; ++i) {
        d = 26 * d + c.charCodeAt(i) - 64;
    }
    return d - 1;
};

/**
 * col2Letter  列数转换为字母表现形式
 *
 * 说明: 0=>A, 1=>B...
 *
 * @param {Number} col 列数(0...n)
 * @return {String} A...Z
 */
Utils.col2Letter = function (col) {
    let letter = "";
    // excel index(1...n)/(A...Z...)
    for (++col; col; col = Math.floor((col - 1) / 26)) {
        letter = String.fromCharCode(((col - 1) % 26) + 65) + letter;
    }
    return letter;
};

/**
 * Converts coordinates to sheet keys(A1...Zn)
 *
 * @param {Number} x coord(0...n)
 * @param {Number} y coord(0...n)
 * @return {String} key(A1...Zn)
 */
Utils.coord2SheetKey = function (x, y) {
    let key = Utils.col2Letter(x);
    // excel index(1...n)/(A...Z...)
    key += ++y;
    return key;
};


/**
 * Converts ASCII code(letters) to the corresponding letters
 *
 * @param {Number} number ASCII code
 * @return {String} letter(A...Z)
 */
Utils.ascii2Letter = function (number) {
    // 转26进制
    let temp = '';
    number -= ascii_offset;
    while (number > twenty_six_system) {
        let miss = number % twenty_six_system;
        number = Math.floor(number / twenty_six_system);
        temp = String.fromCharCode(miss + ascii_offset) + temp;
    }
    if (number > 0) {
        temp += String.fromCharCode(number + ascii_offset);
    }
    return temp;
};

/**
 * Converts letters to the corresponding ASCII code
 *
 * @param {String} letter Letter(A...Z)
 * @return {Number} ASCII code
 */
Utils.letter2ASCII = function (letter) {
    // 转10进制
    let number = 0;
    let length = letter.length;
    for (let i = length - 1; i >= 0; i--) {
        let miss = letter.charCodeAt(i) - ascii_offset;
        number += miss * Math.pow(twenty_six_system, i);
    }
    return number + ascii_offset;
};


/**
 * objectSimplify 简化object mode的数据结构
 *
 * {'p1':{'id':'p1','name':'rose'},'p2':{'id':'p2','name':'jordan','sex':1},...}
 * be simplify to: {'p1':'rose','p2':{'id':'p2','name':'jordan','sex':1},...}
 *
 * @param {Object} json  object mode导出的json数据
 * @return {void}
 */
Utils.objectSimplify = function (json) {
    if (!json) {
        return;
    }
    let keys = Object.keys(json);
    if (keys.length <= 0) {
        return;
    }
    keys.forEach(function (k) {
        let item = json[k];
        let iks = Object.keys(item);
        if (iks.length === 2) {
            json[k] = item[iks[1]];
        }
    });
};

/**
 * arraySimplify 简化array mode的数据结构
 *
 * [{'id':false},{'id':100},{'id':'test'},{'id':'p1','name':'rose'},...]
 * be simplify to: [false,100,'test',{'id':'p1','name':'rose'},...]
 *
 * @param {Object} json  array mode导出的json数据
 * @return {void}
 */
Utils.arraySimplify = function (json) {
    if (!json) {
        return;
    }
    if (json.length <= 0) {
        return;
    }
    for (let i = 0; i < json.length; i++) {
        let item = json[i];
        let iks = Object.keys(item);
        if (iks.length === 1) {
            json[i] = item[iks[0]];
        }
    }
};
