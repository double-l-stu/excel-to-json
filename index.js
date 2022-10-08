/*
 * @Author: yunmin
 * @Email: 362279869@qq.com
 * @Date: 2022-10-07 15:18:57
 */

const {
    cmdSystem
} = require("exp-utils");
const handle = require("./src/handle");

cmdSystem.add("-r", handle.setReadPath, "Set read directory of excel files(eg: -r your-path).");
cmdSystem.add("-w", handle.setWritePath, "Set export directory of json files(eg: -w your-path).");
cmdSystem.add("-c", handle.startConvert, "Convert excel files to json files(eg: -c or -c file1 file2...).");

cmdSystem.start(true);
