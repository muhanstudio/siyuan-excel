const initTopbar = (pluginInstance) => {
    const i18n = pluginInstance.i18n
    const topBarElement = pluginInstance.addTopBar({
        icon: "iconTable",
        title: i18n.siyuanExcel,
        position: "right",
        callback: () => {},
    });
    topBarElement.addEventListener("click", async () => {
        const { showTab } = pluginInstance.requireLib("tab")
        await showTab(pluginInstance, "index.html")
    });
    console.log("inited topbar from siyuan excel", pluginInstance)
}

module.exports = {
    initTopbar
}