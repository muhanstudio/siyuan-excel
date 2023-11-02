const {executeVueCli} = require("./scriptutils");

// ==============================================

async function doServe() {
    await executeVueCli("serve")
}

// ==============================================

(async () => {
    console.log("vue app is serving on http://localhost:8080 ...")
    await doServe()
})()