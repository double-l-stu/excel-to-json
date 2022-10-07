/*
 * @Author: yunmin
 * @Email: 362279869@qq.com
 * @Date: 2022-10-07 10:22:57
 */

const local = require('../utils');
const CONST = require('../const');
const xlsx = require('xlsx');

/**
 * A Sheet Table From Excel
 */
class Entity {
    /**
     * Constructor
     *
     * @param {Object} sheet  source date
     */
    constructor(sheet) {
        // src data
        this.src = sheet;
        // {"!reg": "A1:C6"},行数: 6行(1-6)，列数: 3列(A-C)
        let ref = sheet[CONST.TABLE.REF_NAME];
        if (!ref) {
            console.log('Empty Sheet.');
            return;
        }
        // Range{start: {x: 0, y: 0}, end: {x: 0, y: 0}
        this.range = local.getSheetRange(ref);
    };

    /**
     * Gets the end cell coordinates
     *
     * @return {Object} {x, y}
     */
    getEnd() {
        return this.range.end;
    };

    /**
     * Gets the start cell coordinates
     *
     * @return {Object} {x, y}
     */
    getStart() {
        return this.range.start;
    };

    /**
     * Gets the cell range coordinates
     *
     * @return {Object} {start: {x: 0, y: 0}, end: {x: 0, y: 0}
     */
    getRange() {
        return this.range;
    };

    /**
     * Gets the type of the value in the cell
     *
     * @param {Number} x the x coord(0...n)
     * @param {Number} y the y coord(0...n)
     * @return {String} the type(CONST.TYPE)
     */
    getType(x, y) {
        let key = local.coord2SheetKey(x, y);
        // console.log("key: " + key);
        return this.src[key].t;
    };

    /**
     * Gets the value in the cell
     *
     * @param {Number} x the x coord(0...n)
     * @param {Number} y the y coord(0...n)
     * @return {Number} the value
     */
    getValue(x, y) {
        let key = local.coord2SheetKey(x, y);
        // console.log("key: " + key);
        return this.src[key].v;
    };

    /**
     * Checks if there is a value in the cell
     *
     * @param {Number} x the x coord(0...n)
     * @param {Number} y the y coord(0...n)
     * @return {Boolean}
     */
    hasField(x, y) {
        let key = local.coord2SheetKey(x, y);
        // console.log("key: " + key);
        return this.src[key] !== undefined;
    };

    /**
     * Converts to json object.
     *
     * @return {Object} [{...}]
     */
    toJson() {
        return xlsx.utils.sheet_to_json(this.src);
    };
}

/**
 * create  create a entity
 *
 * @param {Object} opts custom config
 * @return {Object} the entity created
 */
exports.create = function (opts) {
    let entity = new Entity(opts);
    return entity;
};