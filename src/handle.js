/*
 * @Author: yunmin
 * @Email: 362279869@qq.com
 * @Date: 2022-10-08 21:29:07
 */

const {
    utils,
    fileSystem
} = require("exp-utils");
const localutils = require('./utils')
const Portfolio = require('./portfolio');
const configPath = "./config/config.json";

let config = fileSystem.readAsJsonSync(configPath);

console.log(`Read Path: ${config.readPath}`);
console.log(`Write Path: ${config.writePath}\n`);

exports.setReadPath = function (params, next) {
    if (params.length > 0) {
        config.readPath = params[0];
        fileSystem.writeSync(configPath, config);
    }
    else {
        console.log("The right command: -r your-path");
    }
    utils.invoke(next);
};

exports.setWritePath = function (params, next) {
    if (params.length > 0) {
        config.writePath = params[0];
        fileSystem.writeSync(configPath, config);
    } else {
        console.log("The right command: -w your-path");
    }
    utils.invoke(next);
};

exports.startConvert = function (params, next) {
    console.log(params);
    let portfolio = Portfolio.create(config.readPath);
    let names = portfolio.getDocumentsName();
    names.forEach(function (name) {
        console.log('Converting document: ' + name);
        let data = portfolio.doc2Object(name);
        // simplified data
        if (data) {
            localutils.objectSimplify(data.nor);
        }
        // save to file
        if (data) {
            fileSystem.writeSync(config.writePath + name + '.json', data.nor);
            if (data.rep.length > 0) {
                console.warn('repeating field:');
                console.warn(data.rep);
            }
            console.log('Exported document: ' + name);
        } else {
            console.warn('Invalid excel file: ' + name);
        }
    });
    utils.invoke(next);
};
