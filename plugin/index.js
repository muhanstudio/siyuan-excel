const {
    Plugin,
    showMessage,
    openTab
} = require("siyuan");

class SiyuanExcel extends Plugin {
    // 挂在便于后面用
    fn = {
        // 插件相关能力
        openTab: openTab,
        showMessage: showMessage,

        // tab
        customTabObject: undefined,
        tabInstance: undefined,
    }

    onload() {
        const {initTopbar} = this.requireLib("topbar")
        initTopbar(this)
    }

    requireLib(libpath) {
        const dataDir = window.siyuan.config.system.dataDir
        const thisPluginDir = `${dataDir}/plugins/siyuan-excel/`
        const fullLibpath = `${thisPluginDir}${libpath}`
        console.log("fullLibpath =>", fullLibpath)
        return window.require(fullLibpath)
    }
}

module.exports = SiyuanExcel;