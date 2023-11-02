const showTab = async (pluginInstance, pageUrl) => {
    const fn = pluginInstance.fn
    const i18n = pluginInstance.i18n
    const app = pluginInstance.app
    const { createLogger } = pluginInstance.requireLib("logger")
    const logger = createLogger("tab", "siyuan-excel", false)

    // 自定义tab
    fn.tabInstance = fn.openTab({
        app: app,
        custom: {
            id: "siyuan-excel-tab",
            icon: "iconAccount",
            title: i18n.siyuanExcel,
            fn: fn.customTabObject,
        },
    })
    if (fn.tabInstance instanceof Promise) {
        fn.tabInstance = await fn.tabInstance
    }
    // dev
    // const url = `http://localhost:8080/${pageUrl}`
    const url = `/plugins/siyuan-excel/dist/${pageUrl}`
    logger.info("will show webview page =>", url)

    // 参考 https://github.com/zuoez02/siyuan-plugin-webview-flomo/blob/main/index.js#L380C20-L382C29
    fn.tabInstance.panelElement.innerHTML = `
      <div style="display: flex" class="fn__flex-column fn__flex fn__flex-1 plugin-publisher__custom-tab">
          <iframe allowfullscreen allowpopups style="border: none" class="fn__flex-column fn__flex  fn__flex-1" src="${url}"></iframe>
      </div>`
}

module.exports = {
    showTab
}