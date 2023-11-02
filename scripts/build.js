const path = require('path');
const os = require('os');
const fs = require('fs');
const archiver = require('archiver');
const { executeScript } = require(path.join(__dirname, "scriptutils"))

// ==============================================

async function doBuild() {
    // 构建app
    const buildAppScript = os.platform() === 'win32'
        ? 'set NODE_OPTIONS=--openssl-legacy-provider && vue-cli-service build'
        : 'export NODE_OPTIONS=--openssl-legacy-provider && vue-cli-service build';
    const {stdout: appResult} = await executeScript(buildAppScript)
    console.log(appResult)

    // 插件，插件不用打包
}

async function doPackage() {
    // 创建输出流
    const output = fs.createWriteStream('package.zip');
    const archive = archiver('zip', {
        zlib: {level: 9} // 设置压缩级别
    });

    // 将输出流管道到文件
    archive.pipe(output);

    // 添加 dist 目录
    archive.directory('dist', 'dist');
    // 添加 plugin 目录
    archive.directory('plugin', '');
    // 添加单独的文件
    archive.file('package.json', {name: 'package.json'});
    archive.file('README.md', {name: 'README.md'});
    archive.file('README_zh_CN.md', {name: 'README_zh_CN.md'});
    archive.file('plugin.json', {name: 'plugin.json'});
    archive.file('preview.png', {name: 'preview.png'});
    archive.file('icon.png', {name: 'icon.png'});

    // 完成压缩
    await archive.finalize();
    console.log('打包完成: package.zip');

    // 复制一份到 siyuan-excel.zip
    setTimeout(()=>{
        fs.copyFileSync('package.zip', 'siyuan-excel.zip');
        console.log('复制完成: package.zip -> siyuan-excel.zip');
    }, 5000)

}

// ==============================================

(async () => {
    await doBuild()
    await doPackage()
})()
