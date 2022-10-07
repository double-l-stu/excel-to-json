/*
 * @Author: yunmin
 * @Email: 362279869@qq.com
 * @Date: 2022-10-07 10:21:57
 */

const xlsx = require('xlsx');
const Table = require('./table');

/**
 * A complete Excel document
 */
class Entity {
    /**
     * Constructor
     *
     * @param {Object} file  full filename
     */
    constructor(file) {
        // src data
        this.src = xlsx.readFile(file);
        // sheet table
        this.tables = new Map();

        let self = this;
        // Sheets: sheet content list
        let sheets = this.src.Sheets;
        // SheetNames: sheet name list
        let sheet_names = this.src.SheetNames;
        sheet_names.forEach(function (name) {
            // read sheet table data..
            let config = sheets[name];
            let table = Table.create(config);
            self.tables.set(name, table);
        });
    };

    /**
     * Checks if the table exists
     *
     * @param {String} n table name
     * @return {Boolean} true/false
     */
    hasTable(n) {
        return this.tables.has(n);
    };

    /**
     * Gets a table by name
     *
     * @param {String/Number} n table name/table index
     * @return {Object} Table/null
     */
    getTable(n) {
        if (typeof n === 'string') {
            return this.tables.get(n);
        } else {
            let names = this.src.SheetNames;
            if (n >= 0 && n < names.length) {
                return this.tables.get(names[n]);
            }
        }
        console.warn('index overflow: ' + n);
        return null;
    };

    /**
     * Gets all tables
     *
     * @return {Array} Tables
     */
    getAllTable() {
        let tables = [];
        this.tables.forEach(function (v) {
            tables.push(v);
        });
        return tables;
    };

    /**
     * Converts a table to json object.
     *
     * @param {String/Number} n table name/table index
     * @return {Object} JSON/null
     */
    table2Json(n) {
        let table = this.getTable(n);
        if (table) {
            return table.toJson();
        }
        return null;
    };

    /**
     * Converts excel to json object.
     *
     * @return {Object} {sheet1...sheetn}
     */
    toJson() {
        let json = {};
        this.tables.forEach(function (table, key) {
            json[key] = table.toJson();
        });
        return json;
    };
}

/**
 * create  create a entity
 *
 * @param {String} pathname  full file path
 * @return {Object} the entity created
 */
exports.create = function (pathname) {
    let entity = new Entity(pathname);
    return entity;
};