/*
 * @Author: yunmin
 * @Email: 362279869@qq.com
 * @Date: 2022-10-11 21:30:58
 */
const fs = require("fs");

function start() {
    if (fs.existsSync("./node_modules")) {
        return;
    }
    console.log("Start install package modules...");
    const cp = require("child_process");
    cp.execSync('npm install -d', {stdio: 'inherit'});    
};

start();
