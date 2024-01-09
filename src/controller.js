const request = require('request')
const fs = require('fs')
const XLSX = require('xlsx')

const newOrder = require('./newOrder')

const log4js = require('log4js')
const logger = log4js.getLogger('default')

const { send: { root } } = require('./config.json')

const Second = 1000
const Minute = 60 * Second
const Hour = 60 * Minute
const Day = 24 * Hour

class Controller {
    #jwt

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

    send(order) {
        return new Promise((resolve, reject) => {
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

    setOrderTask(result = null) {
        request(`${root}/setOrderTask`, {
            method: 'POST',
            headers: {
                jwt: this.#jwt,
                'content-type': 'application/json'
            },
            body: JSON.stringify({ result })
        }, (error, response, body) => {
            if (error || response.statusCode != 200)
                return this.setOrderTask(error || response.statusCode)

            body = JSON.parse(body)
            if (body.path[0] === '"')
                body.path = body.path.slice(1,)
            if (body.path[body.path.length - 1] === '"')
                body.path = body.path.slice(0, -1)

            if (!fs.existsSync(body.path))
                return this.setOrderTask('Не могу найти такой файл, пожалуйста проверьте правильность его введения')
            let ws = XLSX.readFile(body.path)

            if (!ws.Sheets['СЧЕТ'] || !ws.Sheets['ЗАКАЗ-НАРЯД'])
                return this.setOrderTask('Не могу прочитать файл, пожалуйста проверьте правильность его заполнения и повторите попытку')

            if (!ws.Sheets['ЗАКАЗ-НАРЯД']?.['N8']?.v && !!ws.Sheets['ЗАКАЗ-НАРЯД']?.['M8']?.v || !ws.Sheets['СЧЕТ']['O' + JSON.parse(Object.keys(ws.Sheets['СЧЕТ']).find(key => ws.Sheets['СЧЕТ'][key].v === 'Дата готовности заказа').match(/\d+/)[0])]?.v && !ws.Sheets['СЧЕТ']['N' + JSON.parse(Object.keys(ws.Sheets['СЧЕТ']).find(key => ws.Sheets['СЧЕТ'][key].v === 'Дата готовности заказа').match(/\d+/)[0])]?.v)
                return this.setOrderTask('В заказ наряде нет номера или даты готовности, пожалуйста поправьте заказ наряд и повторите попытку')

            newOrder(body.path).then(order => this.send(order).then(this.setOrderTask.bind(this)).catch(this.setOrderTask.bind(this)))
        })
    }
}

module.exports = Controller
