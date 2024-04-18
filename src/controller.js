const request = require('request')
const fs = require('fs')
const XLSX = require('xlsx')

const newOrder = require('./newOrder')

const log4js = require('log4js')
const logger = log4js.getLogger('default')

const { permanent, send: { root } } = require('./config.json')

const Second = 1000
const Minute = 60 * Second
const Hour = 60 * Minute
const Day = 24 * Hour

class Controller {
    #jwt
    #jwtP

    constructor() { }

    sign() {
        return new Promise((resolve, reject) => {
            request(`${root}/signCompany`, {
                method: 'POST',
                body: JSON.stringify({
                    id: '325fabe9-44c6-4051-b13e-c87d226d4441',
                    authorization: 'termitAPPMAIN4132'
                }),
                headers: {
                    'content-type': 'application/json'
                }
            }, (error, response, body) => {
                if (error || response.statusCode != 200)
                    return reject(error || response.statusCode)

                this.#jwt = JSON.parse(body).jwt
                return resolve('success')
            })
        })
    }

    signP() {
        return new Promise((resolve, reject) => {
            request(`${root}/sign`, {
                method: 'POST',
                body: JSON.stringify({
                    login: 'PaymentPointer',
                    password: 'paymentPointer1029384756'
                }),
                headers: {
                    'content-type': 'application/json'
                }
            }, (error, response, body) => {
                if (error || response.statusCode != 200)
                    return reject(error || response.statusCode)

                this.#jwtP = JSON.parse(body).jwt
                return resolve('success')
            })
        })
    }

    verify() {
        request(`${root}/verifyCompany`, {
            method: 'GET',
            headers: {
                jwt: this.#jwt
            }
        }, (error, response, body) => {
            if (error || response.statusCode != 200) {
                logger.error(error || response.statusCode)
                return setTimeout(this.verify.bind(this), 5 * Minute)
            }

            logger.info(body)

            if (body != 'verified')
                return this.sign().catch(logger.error.bind(logger)).then(() => setTimeout(this.verify.bind(this), 5 * Minute))

            setTimeout(this.verify.bind(this), 5 * Minute)
        })
    }

    send(order) {
        return new Promise((resolve, reject) => {
            // console.log(this.#jwt)
            request(`${root}/setOrder`, {
                method: 'POST',
                headers: {
                    jwt: this.#jwt,
                    'content-type': 'application/json'
                },
                body: JSON.stringify(order)
            }, (error, response, body) => {
                if (error || response.statusCode != 200)
                    return reject(error || response.statusCode)
                return resolve(JSON.parse(body))
            })
        })
    }

    lounch() {
        return new Promise((resolve, reject) => {
            request(`${root}/add`, {
                method: 'GET',
                headers: {
                    jwt: this.#jwt
                }
            }, (error, response, body) => {
                if (error || response.statusCode != 200)
                    return reject(error || response.statusCode)
                return resolve(JSON.parse(body))
            })
        })
    }

    updateAll() {
        request(`${root}/updateAll`, {
            method: 'GET',
            headers: {
                jwt: this.#jwt
            }
        }, (error, response, body) => {
            if (error || response.statusCode != 200)
                logger.error(error || response.statusCode)
            logger.info(body)

            setTimeout(this.updateAll.bind(this), 15 * Minute)
        })
    }

    getPathes(origin) {
        let orders = []
        fs.readdirSync(permanent.root).map(costumer => {
            if (fs.lstatSync(permanent.root + '\\' + costumer).isDirectory())
                fs.readdirSync(permanent.root + '\\' + costumer).filter(year => year >= new Date().getFullYear()).map(year => {
                    if (fs.lstatSync(permanent.root + '\\' + costumer + '\\' + year).isDirectory())
                        fs.readdirSync(permanent.root + '\\' + costumer + '\\' + year).map(order => {
                            orders.push({
                                dir: permanent.root + '\\' + costumer + '\\' + year + '\\' + order,
                                updatedAt: fs.lstatSync(permanent.root + '\\' + costumer + '\\' + year + '\\' + order).mtime
                            })
                        })
                    else
                        orders.push({
                            dir: permanent.root + '\\' + costumer + '\\' + year,
                            updatedAt: fs.lstatSync(permanent.root + '\\' + costumer + '\\' + year).mtime
                        })
                })
            else
                orders.push({
                    dir: permanent.root + '\\' + costumer,
                    updatedAt: fs.lstatSync(permanent.root + '\\' + costumer).mtime
                })
        })
        orders = orders.filter(({ dir }) => !dir.includes('~') && dir.includes('.xlsx') && origin.find(number => dir.includes(number)))
        orders = orders.map((order) => {
            if (!fs.existsSync(order.dir))
                return
            const ws = XLSX.readFile(order.dir)

            if (!ws.Sheets['ЗАКАЗ-НАРЯД'])
                return
            if (ws.Sheets['ЗАКАЗ-НАРЯД']?.['N7']?.v && typeof ws.Sheets['ЗАКАЗ-НАРЯД']?.['N7']?.v === 'string' && ws.Sheets['ЗАКАЗ-НАРЯД']?.['N7']?.v?.match(/\d+/))
                order.number = JSON.parse(ws.Sheets['ЗАКАЗ-НАРЯД']?.['N7']?.v?.match(/\d+/)?.[0])
            else if (ws.Sheets['ЗАКАЗ-НАРЯД']?.['M7']?.v && typeof ws.Sheets['ЗАКАЗ-НАРЯД']?.['M7']?.v === 'string' && ws.Sheets['ЗАКАЗ-НАРЯД']?.['M7']?.v?.match(/\d+/))
                order.number = JSON.parse(ws.Sheets['ЗАКАЗ-НАРЯД']?.['M7']?.v?.match(/\d+/)?.[0])
            if (!order.number)
                return

            return order
        }).filter(order => order && origin.includes(order.number))
        orders.sort((a, b) => +b.updatedAt - +a.updatedAt)
        orders = orders.filter(({ number }, index, array) => array.findIndex(order => order.number === number) === index)

        return orders.map(({ dir }) => dir)
    }

    setOrderTask(result = null) {
        request(`${root}/setOrderTask`, {
            method: 'POST',
            headers: {
                jwt: this.#jwt,
                'content-type': 'application/json'
            },
            body: JSON.stringify({ result })
        }, async (error, response, body) => {
            if (error || response.statusCode != 200)
                return this.setOrderTask(error || response.statusCode)

            body = JSON.parse(body)
            if (body.path[0] === '"')
                body.path = body.path.slice(1,)
            if (body.path[body.path.length - 1] === '"')
                body.path = body.path.slice(0, -1)
            if (!body.path.replaceAll(/[\d+ {0,},{0,}]/g, '').length)
                body.orders = body.path.split(/ {0,}, {0,}/).map(number => +number)
            else if (!body.path.match(/D:\\/) && !body.path.match(/\\\\/))
                body.path = permanent.root + '\\' + body.path
            console.log(body.path)

            if (body.orders) {
                body.pathes = this.getPathes(body.orders)

                body.orders = (await Promise.all(body.pathes.map(newOrder)).catch(logger.error.bind(logger)))?.filter(order => typeof order != 'string')

                for (let order of body.orders) {
                    console.log('Sended order', order.number)
                    let result = await this.send(order).catch(logger.error.bind(logger))

                    logger.info(result)
                }

                return this.setOrderTask.bind(this)('Orders created successfully')
            }

            if (!fs.existsSync(body.path))
                return this.setOrderTask('Не могу найти такой файл, пожалуйста проверьте правильность его введения')
            let ws = XLSX.readFile(body.path)

            if (!ws.Sheets['СЧЕТ'] || !ws.Sheets['ЗАКАЗ-НАРЯД'])
                return this.setOrderTask('Не могу прочитать файл, пожалуйста проверьте правильность его заполнения и повторите попытку')

            if (!ws.Sheets['ЗАКАЗ-НАРЯД']?.['N8']?.v && !ws.Sheets['ЗАКАЗ-НАРЯД']?.['M8']?.v || !ws.Sheets['СЧЕТ']['O' + JSON.parse(Object.keys(ws.Sheets['СЧЕТ']).find(key => ws.Sheets['СЧЕТ'][key].v === 'Дата готовности заказа').match(/\d+/)[0])]?.v && !ws.Sheets['СЧЕТ']['N' + JSON.parse(Object.keys(ws.Sheets['СЧЕТ']).find(key => ws.Sheets['СЧЕТ'][key].v === 'Дата готовности заказа').match(/\d+/)[0])]?.v)
                return this.setOrderTask('В заказ наряде нет номера или даты готовности, пожалуйста поправьте заказ наряд и повторите попытку')

            newOrder(body.path).then(order => this.send(order).then(this.setOrderTask.bind(this)).catch(this.setOrderTask.bind(this)))
        })
    }
}

module.exports = Controller
