const Controller = require('./controller')
const permanent = require('./permanent')
require('./loggerConfigurer')

const log4js = require('log4js')
const logger = log4js.getLogger('default')

process.on('uncaughtException', error => {
    debugger
})

const controller = new Controller()

controller.sign().then(() => {
    logger.info('Controller signed successfully')
    controller.lounch().then(result => {
        logger.info(result)
        logger.info('Controller document loaded successfully')
        controller.setOrderTask()
        // controller.updateAll()
        logger.info('Permanent reading orders directory started')
        permanent(controller)
    })
}).catch(logger.error.bind(logger))
