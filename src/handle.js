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
const Portfolio = require('./portfolio');
const configPath = "./config/config.json";

let config = fileSystem.readAsJsonSync(configPath);

console.log(`Read Path: ${config.readPath}`);
console.log(`Write Path: ${config.writePath}\n`);

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
        console.log("Update write path to: " + config.writePath);
    } else {
        console.warn("The right command: -w your-path");
    }
    utils.invoke(next);
};

exports.arrayConvert = function (params, next) {
    console.log(params);
    let portfolio = Portfolio.create(config.readPath);
    let names = portfolio.getDocumentsName();
    names.forEach(function (name) {
        console.log('Converting document: ' + name);
        fileSystem.makeDirsSync(config.writePath);
        let filename = path.join(config.writePath, name) + '.json';
        fileSystem.writeSync(filename, portfolio.doc2Array(name));
        console.log('Exported document: ' + filename);
    });
    utils.invoke(next);
};

exports.objectConvert = function (params, next) {
    console.log(params);
    let portfolio = Portfolio.create(config.readPath);
    let names = portfolio.getDocumentsName();
    names.forEach(function (name) {
        console.log('Converting document: ' + name);
        fileSystem.makeDirsSync(config.writePath);
        let filename = path.join(config.writePath, name) + '.json';
        fileSystem.writeSync(filename, portfolio.doc2Object(name));
        console.log('Exported document: ' + filename);
    });
    utils.invoke(next);
};
