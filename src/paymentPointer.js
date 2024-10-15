const log4js = require('log4js')
const logger = log4js.getLogger('default')

const Second = 1000
const Minute = 60 * Second
const Hour = 60 * Minute
const Day = 24 * Hour

const newOrder = require('./newOrder')

function delay(milliseconds) {
    return new Promise(resolve => setTimeout(resolve, milliseconds))
}


async function paymentPointer(controller) {
    
}

module.exports = paymentPointer
