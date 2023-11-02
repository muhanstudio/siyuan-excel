const {executeVueCli} = require("./scriptutils");

// ==============================================

async function doLint() {
    await executeVueCli("lint")
}

// ==============================================

(async () => {
    await doLint()
})()