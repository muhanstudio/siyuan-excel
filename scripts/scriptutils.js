const util = require("util");
const os = require("os");
const path = require("path")
const fs = require("fs")
const copyFile = util.promisify(fs.copyFile);
const mkdir = util.promisify(fs.mkdir);
const exec = util.promisify(require('child_process').exec);

async function executeScript(script) {
    const {stdout, stderr} = await exec(script);
    if (stderr) {
        console.error("an error occured =>", stderr)
    }
    if(stdout){
        console.log(stdout)
    }
    return { stdout, stderr }
}

async function executeVueCli(cmd) {
    const serveAppScript = os.platform() === 'win32'
        ? `set NODE_OPTIONS=--openssl-legacy-provider && vue-cli-service ${cmd}`
        : `export NODE_OPTIONS=--openssl-legacy-provider && vue-cli-service ${cmd}`;
    await executeScript(serveAppScript)
}


async function copyDir(src, dest) {
    await mkdir(dest, {recursive: true});
    const entries = await fs.promises.readdir(src, {withFileTypes: true});

    for (const entry of entries) {
        const srcPath = path.join(src, entry.name);
        const destPath = path.join(dest, entry.name);

        if (entry.isDirectory()) {
            await copyDir(srcPath, destPath);
        } else {
            await copyFile(srcPath, destPath);
        }
    }
}

async function copyFileToDir(src, dest) {
    const fileName = path.basename(src);
    const destPath = path.join(dest, fileName);
    await copyFile(src, destPath);
}

function joinPath(path1, path2) {
    return path.join(path1, path2)
}

module.exports = {
    executeScript,
    executeVueCli,
    copyDir,
    copyFile,
    copyFileToDir,
    joinPath
}