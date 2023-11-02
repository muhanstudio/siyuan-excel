/**
 * 一个简单轻量级的日志记录器
 *
 * @author terwer
 * @version 1.0.0
 * @since 1.0.0
 */

const createLogger = (name, customSign, isDev) => {
    const sign = customSign || "zhi";

    const formatDate = (date) => {
        const year = date.getFullYear();
        const month = String(date.getMonth() + 1).padStart(2, "0");
        const day = String(date.getDate()).padStart(2, "0");
        const hours = String(date.getHours()).padStart(2, "0");
        const minutes = String(date.getMinutes()).padStart(2, "0");
        const seconds = String(date.getSeconds()).padStart(2, "0");

        return `${year}-${month}-${day} ${hours}:${minutes}:${seconds}`;
    };

    const log = (msg, obj) => {
        const time = formatDate(new Date());
        const formattedMsg = typeof obj === "boolean" ? String(obj) : obj;

        if (formattedMsg) {
            console.log(`[${sign}] [${time}] [DEBUG] [${name}] ${msg}`, formattedMsg);
        } else {
            console.log(`[${sign}] [${time}] [DEBUG] [${name}] ${msg}`);
        }
    };

    const infoLog = (msg, obj) => {
        const time = formatDate(new Date());
        const formattedMsg = typeof obj === "boolean" ? String(obj) : obj;

        if (formattedMsg) {
            console.info(`[${sign}] [${time}] [INFO] [${name}] ${msg}`, formattedMsg);
        } else {
            console.info(`[${sign}] [${time}] [INFO] [${name}] ${msg}`);
        }
    };

    const warnLog = (msg, obj) => {
        const time = formatDate(new Date());
        const formattedMsg = typeof obj === "boolean" ? String(obj) : obj;

        if (formattedMsg) {
            console.warn(`[${sign}] [${time}] [WARN] [${name}] ${msg}`, formattedMsg);
        } else {
            console.warn(`[${sign}] [${time}] [WARN] [${name}] ${msg}`);
        }
    };

    const errorLog = (msg, obj) => {
        const time = formatDate(new Date());
        if (obj) {
            console.error(`[${sign}] [${time}] [ERROR] [${name}] ${msg.toString()}`, obj);
        } else {
            console.error(`[${sign}] [${time}] [ERROR] [${name}] ${msg.toString()}`);
        }
    };

    return {
        debug: (msg, obj) => {
            if (isDev) {
                if (obj) {
                    log(msg, obj);
                } else {
                    log(msg);
                }
            }
        },
        info: (msg, obj) => {
            if (obj) {
                infoLog(msg, obj);
            } else {
                infoLog(msg);
            }
        },
        warn: (msg, obj) => {
            if (obj) {
                warnLog(msg, obj);
            } else {
                warnLog(msg);
            }
        },
        error: (msg, obj) => {
            if (obj) {
                errorLog(msg, obj);
            } else {
                errorLog(msg);
            }
        },
    };
};

module.exports = {
    createLogger
}