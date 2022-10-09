/*
 * @Author: yunmin
 * @Email: 362279869@qq.com
 * @Date: 2022-10-07 15:18:57
 */

const {
    cmdSystem
} = require("exp-utils");
const handle = require("./src/handle");

cmdSystem.add("-a", handle.arrayConvert, "Convert excel to json array(eg: -c or -c file1 file2...)");
cmdSystem.add("-o", handle.objectConvert, "Convert excel to json object(eg: -c or -c file1 file2...)");
cmdSystem.add("-r", handle.setReadPath, "Set read directory of excel files(eg: -r your-path)");
cmdSystem.add("-w", handle.setWritePath, "Set export directory of json files(eg: -w your-path)");
cmdSystem.add("-l", handle.listExcelFiles, "List all excel files in the read folder");

cmdSystem.start(true);
