/*
 * @Author: yunmin
 * @Email: 362279869@qq.com
 * @Date: 2022-10-08 23:01:07
 */

const {
    fileSystem
} = require('exp-utils');
const CONST = require('./const');
const Document = require('./entity/document');
const utils = require('./utils')

/**
 * Gets Unit(Cell) Info(VALUE/ARRAY/OBJECT)
 * 
 * see <template.xlsx> for more
 *
 * @param {Object} value the value be checked
 * @return {Object} unit info
 */
function getUnitInfo(value) {
    let info = {
        unit: CONST.UNIT.VALUE
    };
    if (typeof value !== 'string') {
        return info;
    }
    // [A:B]/{A:B} length>=5
    if (value.length < 5) {
        return info;
    }
    let c0 = value.charAt(0);
    let cn = value.charAt(value.length - 1);
    if (c0 === '[' && cn === ']') {
        let index = value.indexOf(':');
        if (index >= 0) {
            // [sheet:A]
            info.sheet = value.substring(1, index++);
            let col = value.substring(index, value.length - 1);
            // col number
            info.col = utils.letter2Col(col);
            info.unit = CONST.UNIT.ARRAY;
        }
    }
    if (c0 === '{' && cn === '}') {
        let index = value.indexOf(':');
        if (index >= 0) {
            // {sheet:1}
            info.sheet = value.substring(1, index++);
            let row = value.substring(index, value.length - 1);
            // row number
            info.row = parseInt(row);
            info.unit = CONST.UNIT.OBJECT;
        }
    }
    return info;
}

/**
 * All documents read
 */
class Entity {
    /**
     * Constructor
     *
     * @param {String} folder  full path
     */
    constructor(folder) {
        // documents
        this.documents = new Map();
        // load files
        let files = [];
        // check if is a excel file
        CONST.EXTENDS.forEach(function (exd) {
            let t = fileSystem.filesInDirSync(folder, exd, true);
            files = files.concat(t);
        });
        if (files.length <= 0) {
            console.log("no excel in: " + folder);
            return;
        }
        // create document
        let self = this;
        files.forEach(function (item) {
            // filter cache files
            if (item.name.charAt(0) !== '~') {
                let path = item.path + item.name;
                let doc = Document.create(path);
                // remove the extend(eg: xxx.xlsx 2 xxx)
                let i = item.name.lastIndexOf('.');
                let name = item.name.substring(0, i);
                self.documents.set(name, doc);
            }
        });
    };

    /**
     * Gets all documents name
     *
     * @return {Array} documents name
     */
    getDocumentsName() {
        let names = [];
        this.documents.forEach(function (value, key) {
            names.push(key);
        });
        return names;
    };

    /**
     * Checks if the document exists
     *
     * @param {String} n document name
     * @return {Boolean} true/false
     */
    hasDocument(n) {
        return this.documents.has(n);
    };

    /**
     * Gets a document by name
     *
     * @param {String} n document name
     * @return {Object} the document
     */
    getDocument(n) {
        return this.documents.get(n);
    };

    /**
     * Converts a document to JSON Array
     *
     * 连结结构 文档内根据各种单元连接成一个单一JSON对象
     *
     * 特殊说明 Sheet不能指向本表单
     *
     * 数组单元 ([Sheet表单名:A-Z列号])
     * 对象单元 ({Sheet表单名:1-n行号})
     *
     * Structure [{id:0,name:"king",skills:[1,5,6],describe:{chinese:"国王",english:"King"}},
     *           {id:1,name:"queen",skills:[2,3,4],describe:{chinese:"公主",english:"Queen"}}]
     *
     * @param {String} n document name
     * @return {Object} the json object/null
     */
    doc2Array(n) {
        if (this.hasDocument(n)) {
            let doc = this.getDocument(n);
            let json = doc.table2Json(0, true);
            // parse unit(cell) type([SheetName:A-Z],{SheetName:2-n})
            for (let i = 0; i < json.length; i++) {
                let keys = Object.keys(json[i]);
                for (let j = 0; j < keys.length; j++) {
                    let key = keys[j];
                    let info = getUnitInfo(json[i][key]);
                    if (info.sheet === undefined) {
                        continue;
                    }
                    if (!doc.hasTable(info.sheet)) {
                        console.warn('invalid sheet: ' + info.sheet);
                        continue;
                    }
                    if (info.unit === CONST.UNIT.ARRAY) {
                        let table = doc.getTable(info.sheet);
                        let index = 0, arr = [];
                        // get the whole col's value
                        while (table.hasField(info.col, index)) {
                            let v = table.getValue(info.col, index++);
                            arr.push(v);
                        }
                        json[i][key] = arr;
                    } else if (info.unit === CONST.UNIT.OBJECT) {
                        let table = doc.table2Json(info.sheet, false);
                        // (row = 0: keys),index offset = 1
                        json[i][key] = table[info.row - 2];
                    }
                }
            }
            return json;
        } else {
            console.warn(`excel(${n}) not found.`);
        }
        return [];
    };

    /**
     * Converts a document to JSON Object
     *
     * 连结结构 文档内根据各种单元连接成一个单一JSON对象
     *
     * 特殊说明 Sheet不能指向本表单,基础键值必须为字符串(基础键值将使用第一列值)
     *
     * 数组单元 ([Sheet表单名:A-Z列号])
     * 对象单元 ({Sheet表单名:1-n行号})
     *
     * Structure {p0:{id:p0,name:"king",skills:[1,5,6],describe:{chinese:"国王",english:"King"}},
     *           p1:{id:p1,name:"queen",skills:[2,3,4],describe:{chinese:"公主",english:"Queen"}}}
     *
     * @param {String} n document name
     * @return {Object} the json object
     */
    doc2Object(n) {
        let arr = this.doc2Array(n);
        // empty data
        if (arr.length <= 0) {
            console.warn("empty data: " + n);
            return {};
        }
        // get the basic key
        let keys = Object.keys(arr[0]);
        // not object struct
        if (keys.length < 2) {
            console.warn("not object struct: " + n);
            return {};
        }
        let rt = {nor: {}, rep: []}, bk = keys[0];
        //console.log("basic key: " + bk);
        arr.forEach(function (item) {
            let k = item[bk];
            if (!rt.nor.hasOwnProperty(k)) {
                rt.nor[k] = item;
            } else {
                rt.rep.push(item);
            }
        });
        // simplified data
        utils.objectSimplify(rt.nor);
        // repetitive ids
        if (rt.rep.length > 0) {
            console.warn("repetitive: " + rt.rep);
        }
        return rt.nor;
    };

}

/**
 * create  create a entity
 *
 * @param {String} folder folder path
 * @return {Object} the entity created
 */
exports.create = function (folder) {
    let entity = new Entity(folder);
    return entity;
};