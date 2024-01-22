const prices = require('./prices.json')

const Type = [/прямой/, /фрез/, /f/, /фас кл/, /фасад клиента/, /тумба/]
const eFs = ['f-01', 'f-02', 'f-03', 'f-04', 'f-06', 'f-07', 'f-08', 'f-09', 'f-10', 'f-11', 'f-15', 'f-16', 'f-18', 'f-19', 'F20', 'f-21']
const sFs = ['f-05', 'f-12', 'f-13', 'f-14', 'f-17']

function getPublisherSalary(publisher) {
    for (let stage of publisher.stages) {
        let price = 0
        if (stage.stage === 'Шлифовка к изолятору') {
            if (sFs.find(F => publisher.description.toLowerCase().includes(F)))
                price += prices['фрезированные'][stage.stage]['сложная'] * publisher.square
            else
                price += prices['фрезированные'][stage.stage]['простая'] * publisher.square
        } else if (stage.stage === 'Нанесение изолятора') {
            price += prices['фрезированные'][stage.stage] * publisher.square
        } else if (stage.stage === 'Шлифовка к грунту' || stage.stage === 'Нанесение грунта') {
            if (publisher.H)
                price += prices['ручки'][stage.stage] * publisher.amount
            if (publisher.description.toLowerCase().match(Type[1]))
                if (sFs.find(F => publisher.description.toLowerCase().includes(F)))
                    price += prices['фрезированные'][stage.stage]['сложная'] * publisher.square
                else
                    price += prices['фрезированные'][stage.stage]['простая'] * publisher.square
            else {
                price += prices['прямые'][stage.stage] * publisher.square
                if (publisher.thickness === 22 && stage.stage === 'Нанесение грунта')
                    price += prices['прямые']['Нанесение изолятора'] * publisher.square
            }
            if (publisher.sides === 2 && publisher.colourType === 'глянец')
                price += prices['обратки'][stage.stage]
        } else if (stage.stage === 'Шлифовка к покраске') {
            if (publisher.H)
                price += prices['ручки'][stage.stage] * publisher.amount
            if (publisher.description.toLowerCase().match(Type[1]))
                if (sFs.find(F => publisher.description.toLowerCase().includes(F)))
                    price += prices['фрезированные'][stage.stage]['сложная'] * publisher.square
                else
                    price += prices['фрезированные'][stage.stage]['простая'] * publisher.square
            else
                price += prices['прямые'][stage.stage] * publisher.square
            if (publisher.sides === 2)
                price += prices['обратки'][stage.stage]
        } else if (stage.stage === 'Покраска') {
            if (publisher.H)
                price += prices['ручки'][stage.stage] * publisher.amount
            if (!stage.index)
                price += prices['прямые'][stage.stage] * publisher.square
            if (stage.index)
                price += prices['обратки'][stage.stage] * publisher.square
            if (publisher.colourType.match(/лак/))
                price += prices['лак']['односторонний'] * publisher.square
            if (publisher.colourType.match(/лак2/))
                price += prices['лак']['двусторонний'] * publisher.square
        } else if (stage.stage === 'Аппликация'/*  && (publisher.sides === 2 || publisher.T || publisher.description.toLowerCase().match(/лдсп/)) */) {
            price += prices['упаковка']['Аппликация'] * publisher.square
        } else if (stage.stage === 'Полировка') {
            if (publisher.colourType.match(/глянец2/))
                price += prices['глянец']['двусторонний'] * publisher.square
            else if (publisher.colourType.match(/глянец/))
                price += prices['глянец']['односторонний'] * publisher.square
        } else if (stage.stage === 'Упаковка') {
            if (publisher.sides === 1 && !publisher.colourType.match(/глянец/))
                price += prices['упаковка']['чистка'] * publisher.square
            price += prices['упаковка'][stage.stage] * publisher.square
        }

        stage.price = price
    }
}

module.exports = getPublisherSalary
