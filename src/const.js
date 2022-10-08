/*
 * @Author: yunmin
 * @Email: 362279869@qq.com
 * @Date: 2022-10-07 10:20:57
 */

module.exports = {
    TYPE: {
        STRING: 's',        // String
        NUMBER: 'n',        // Numerical(Float/Int)
        BOOLEAN: 'b'        // Boolean
    },
    TABLE: {
        REF_NAME: "!ref",       // Col&Raw Range
        MERGES_NAME: "!merges"  // Cell Merges
    },
    UNIT: {
        VALUE: 0,  // string/number/float/bool..
        ARRAY: 1,  // [SheetName:A-Z(Column letter)]
        OBJECT: 2  // {SheetName:2-n(Line number)}
    },
    EXTENDS: [".xls", ".xlsx", ".xlsm", ".xlsb"] 
};
