const Controller = require('./controller')
const permanent = require('./permanent')
require('./loggerConfigurer')

process.env["NODE_TLS_REJECT_UNAUTHORIZED"] = 0

const log4js = require('log4js')
// const test = require('./test')
const logger = log4js.getLogger('default')

process.on('uncaughtException', error => {
    debugger
})

const controller = new Controller()

controller.sign().then(() => {
    logger.info('Controller signed successfully')
    controller.verify()
    // controller.lounch().then(result => {
        // logger.info(result)
        // logger.info('Controller document loaded successfully')
        controller.setOrderTask()
        // controller.updateAll()
        logger.info('Permanent reading orders directory started')
        // permanent(controller)
        // test(controller)
    // })
}).catch(logger.error.bind(logger))
