const log4js = require('log4js')

log4js.configure({
    appenders: {
        default: {
            type: 'dateFile',
            filename: __dirname + '\\logs/default/log.log',
            compress: true,
            keepFileExt: true,
            numBackups: 30,
            layout: {
                type: 'pattern',
                pattern: '[%d{ISO8601_WITH_TZ_OFFSET}] [%p] %c - %m'
            }
        },
        output: {
            type: 'console',
            layout: {
                type: 'pattern',
                pattern: '%[[%d{ISO8601_WITH_TZ_OFFSET}] [%p] %c - %] %m'
            }
        }
    },
    categories: {
        default: {
            appenders: ['default', 'output'],
            level: 'info'
        },
    }
})
