const fs = require('fs')
const XLSX = require('xlsx')

const log4js = require('log4js')
const logger = log4js.getLogger('default')

const { permanent: { root, LastRead } } = require('./config.json')

const Second = 1000
const Minute = 60 * Second
const Hour = 60 * Minute
const Day = 24 * Hour

const newOrder = require('./newOrder')

function delay(milliseconds) {
    return new Promise(resolve => setTimeout(resolve, milliseconds))
}

function ExcelDateToJSDate(serial) {
    var utc_days = Math.floor(serial - 25569)
    var utc_value = utc_days * 86400
    var date_info = new Date(utc_value * 1000)

    var fractional_day = serial - Math.floor(serial) + 0.0000001

    var total_seconds = Math.floor(86400 * fractional_day)

    var seconds = total_seconds % 60

    total_seconds -= seconds

    var hours = Math.floor(total_seconds / (60 * 60))
    var minutes = Math.floor(total_seconds / 60) % 60

    return +new Date(date_info.getFullYear(), date_info.getMonth(), date_info.getDate(), hours, minutes, seconds)
}

async function permanent(controller) {
    while (true) {
        logger.info('Reading orders directory')
        let orders = []
        let newLR = +new Date()
        fs.readdirSync(root).map(costumer => {
            if (fs.lstatSync(root + '\\' + costumer).isDirectory())
                fs.readdirSync(root + '\\' + costumer).map(year => {
                    if (fs.lstatSync(root + '\\' + costumer + '\\' + year).isDirectory())
                        fs.readdirSync(root + '\\' + costumer + '\\' + year).map(order => {
                            orders.push({
                                dir: root + '\\' + costumer + '\\' + year + '\\' + order,
                                updatedAt: fs.lstatSync(root + '\\' + costumer + '\\' + year + '\\' + order).mtime
                            })
                        })
                    else
                        orders.push({
                            dir: root + '\\' + costumer + '\\' + year,
                            updatedAt: fs.lstatSync(root + '\\' + costumer + '\\' + year).mtime
                        })
                })
            else
                orders.push({
                    dir: root + '\\' + costumer,
                    updatedAt: fs.lstatSync(root + '\\' + costumer).mtime
                })
        })
        logger.info('Orders directory read')

        orders = orders.filter(({ updatedAt, dir }) => updatedAt >= LastRead && !dir.includes('~') && dir.includes('.xlsx'))
        orders = orders.filter(({ dir }) => {
            if (!fs.existsSync(dir))
                return
            let ws = XLSX.readFile(dir)

            if (!ws.Sheets['СЧЕТ'] || !ws.Sheets['ЗАКАЗ-НАРЯД'])
                return

            return ws.Sheets['ЗАКАЗ-НАРЯД']?.['M8']?.v && !`${ws.Sheets['СЧЕТ']['O9']?.v}`?.toLowerCase()?.match(/барнаул/) && !`${ws.Sheets['ЗАКАЗ-НАРЯД']['M9']?.v}`?.toLowerCase()?.match(/барнаул/) &&
                ws.Sheets['СЧЕТ']['N' + JSON.parse(Object.keys(ws.Sheets['СЧЕТ']).find(key => ws.Sheets['СЧЕТ'][key].v === 'Дата готовности заказа').match(/\d+/)[0])]?.v &&
                ExcelDateToJSDate(ws.Sheets['СЧЕТ']['N' + JSON.parse(Object.keys(ws.Sheets['СЧЕТ']).find(key => ws.Sheets['СЧЕТ'][key].v === 'Дата готовности заказа').match(/\d+/)[0])]?.v) + Day >= LastRead
        })
        logger.info('Read orders filtered')

        orders = (await Promise.all(orders.map(({ dir }) => dir).map(newOrder)).catch(logger.error.bind(logger)))?.filter(order => typeof order != 'string')
        logger.info('Read orders transformated')

        for (let order of orders)
            await controller.send(order).then(logger.info.bind(logger)).catch(logger.error.bind(logger))

        let config = JSON.parse(fs.readFileSync(__dirname + '/config.json'))
        config.permanent.LastRead = newLR
        fs.writeFileSync(__dirname + '/config.json', JSON.stringify(config))

        logger.info('New orders sent to service')

        await delay(5 * Minute)
    }
}

module.exports = permanent
