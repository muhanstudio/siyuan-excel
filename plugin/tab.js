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
      <div class="fn__flex-column">
        <div class="protyle-breadcrumb">
          <span class="protyle-breadcrumb__space"></span>
          <button data-type="fullscreen" aria-label="${window.siyuan?.languages?.fullscreen}" class="block__icon fn__flex-center block__icon--show b3-tooltips b3-tooltips__sw">
            <svg style="" data-id="" class=""><use xlink:href="#iconFullscreen"></use></svg>
          </button>
        </div>
        <div class="protyle-preview">
          <iframe class="fn__flex fn__flex-1" style="border: none;" src="${url}" allowfullscreen allowpopups ></iframe>
        </div>
      </div>`

    const breadcrumbElement = fn.tabInstance.panelElement.querySelector(".protyle-breadcrumb");
    const tabBodyElement = breadcrumbElement?.parentElement;
    const fullscreenElement = breadcrumbElement?.querySelector("button.block__icon[data-type=fullscreen]");
    if (fullscreenElement) {
      fullscreenElement.onclick = () => {
        const flag_fullscreen = tabBodyElement.classList.contains("fullscreen");
        tabBodyElement.classList.toggle("fullscreen", !flag_fullscreen);
        fullscreenElement.classList.toggle("toolbar__item--active", !flag_fullscreen);
      }
    }
}

module.exports = {
    showTab
}
