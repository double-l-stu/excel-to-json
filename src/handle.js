/*
 * @Author: yunmin
 * @Email: 362279869@qq.com
 * @Date: 2022-10-08 21:29:07
 */

const path = require("path");
const {
    utils,
    fileSystem
} = require("exp-utils");
const CONST = require('./const');
const Portfolio = require('./portfolio');
const configPath = "./config/config.json";

let config = fileSystem.readAsJsonSync(configPath);

console.log(`Read Path: ${config.readPath}`);
console.log(`Write Path: ${config.writePath}\n`);

fileSystem.makeDirsSync(config.writePath);

exports.setReadPath = function (params, next) {
    if (params.length > 0) {
        config.readPath = params[0];
        fileSystem.writeSync(configPath, config);
        console.log("Update read path to: " + config.readPath);
    }
    else {
        console.warn("The right command: -r your-path");
    }
    utils.invoke(next);
};

exports.setWritePath = function (params, next) {
    if (params.length > 0) {
        config.writePath = params[0];
        fileSystem.writeSync(configPath, config);
        fileSystem.makeDirsSync(config.writePath);
        console.log("Update write path to: " + config.writePath);
    } else {
        console.warn("The right command: -w your-path");
    }
    utils.invoke(next);
};

exports.arrayConvert = function (params, next) {
    let portfolio = Portfolio.create(config.readPath);
    if (params.length <= 0) {
        params = portfolio.getDocumentsName();
    }
    params.forEach(function (name) {
        console.log('Converting document: ' + name);
        let filename = path.join(config.writePath, name) + '.json';
        fileSystem.writeSync(filename, portfolio.doc2Array(name));
        console.log('Exported document: ' + filename);
    });
    utils.invoke(next);
};

exports.objectConvert = function (params, next) {
    let portfolio = Portfolio.create(config.readPath);
    if (params.length <= 0) {
        params = portfolio.getDocumentsName();
    }
    params.forEach(function (name) {
        console.log('Converting document: ' + name);
        let filename = path.join(config.writePath, name) + '.json';
        fileSystem.writeSync(filename, portfolio.doc2Object(name));
        console.log('Exported document: ' + filename);
    });
    utils.invoke(next);
};

exports.listExcelFiles = function (params, next) {
    CONST.EXTENDS.forEach(function (exd) {
        let arr = fileSystem.filesInDirSync(config.readPath, exd, true);
        if (arr.length > 0) {
            arr.forEach((e) => {
                console.log(e.name);
            });
        }
    });
    utils.invoke(next);
};
