const prices = require('./prices.json')

const Type = [/прямой/, /фрез/, /f/, /фас кл/, /фасад клиента/, /тумба/]
const eFs = ['f-01', 'f-02', 'f-03', 'f-04', 'f-06', 'f-07', 'f-08', 'f-09', 'f-10', 'f-11', 'f-15', 'f-16', 'f-18', 'f-19', 'F20', 'f-21']
const sFs = ['f-05', 'f-12', 'f-13', 'f-14', 'f-17']

function getPublisherSalary(publisher) {
    for (let stage of publisher.stages) {
        let price = 0
        if (stage.stage === 'Шлифовка к изолятору') {
            if (publisher.description.toLowerCase().match(/карниз/))
                price += Math.max(publisher.height, publisher.width) / 1000 * publisher.amount * prices['Карнизы МДФ'][stage.stage]
            else {
                if (stage.H)
                    price += prices['ручки']['Шлифовка к грунту'] * publisher.amount
                else if (publisher.H && stage.stage === 'Шлифовка к изолятору' && publisher.stages.find(({ H }) => H).stage === 'Шлифовка к изолятору') {
                    if (sFs.find(F => publisher.description.toLowerCase().includes(F)))
                        price -= prices['фрезированные'][stage.stage]['сложная'] * publisher.square
                    else
                        price -= prices['фрезированные'][stage.stage]['простая'] * publisher.square
                }
                if (sFs.find(F => publisher.description.toLowerCase().includes(F)))
                    price += prices['фрезированные'][stage.stage]['сложная'] * publisher.square
                else
                    price += prices['фрезированные'][stage.stage]['простая'] * publisher.square
            }
        } else if (stage.stage === 'Нанесение изолятора') {
            if (publisher.description.toLowerCase().match(/карниз/))
                price += Math.max(publisher.height, publisher.width) / 1000 * publisher.amount * prices['Карнизы МДФ'][stage.stage]
            else
                price += prices['фрезированные'][stage.stage] * publisher.square
        } else if (stage.stage === 'Шлифовка к грунту' || stage.stage === 'Нанесение грунта') {
            if (publisher.description.toLowerCase().match(/карниз/))
                price += Math.max(publisher.height, publisher.width) / 1000 * publisher.amount * prices['Карнизы МДФ'][stage.stage]
            else {
                if (stage.H)
                    price += prices['ручки'][stage.stage] * publisher.amount
                else if (publisher.H && stage.stage === 'Шлифовка к грунту' && publisher.stages.find(({ H }) => H).stage === 'Шлифовка к грунту') {
                    if ((publisher.description.toLowerCase().match(Type[2]) || publisher.description.toLowerCase().match(Type[1])) && !publisher.description.toLowerCase().match(Type[0]))
                        if (sFs.find(F => publisher.description.toLowerCase().includes(F)))
                            price -= prices['фрезированные'][stage.stage]['сложная'] * publisher.square * (publisher.T || publisher.description.toLowerCase().includes('в сборе') ? 1.5 : 1)
                        else
                            price -= prices['фрезированные'][stage.stage]['простая'] * publisher.square * (publisher.T || publisher.description.toLowerCase().includes('в сборе') ? 1.5 : 1)
                    else {
                        if ((publisher.description.toLowerCase().match(Type[2]) || publisher.description.toLowerCase().match(Type[1])) && publisher.description.toLowerCase().match(Type[0]))
                            price -= 50 * publisher.square
                        price -= prices['прямые'][stage.stage] * publisher.square * (publisher.T || publisher.description.toLowerCase().includes('в сборе') ? 1.5 : 1)
                        if (publisher.thickness === 22 && stage.stage === 'Нанесение грунта')
                            price -= prices['прямые']['Нанесение изолятора'] * publisher.square * (publisher.T || publisher.description.toLowerCase().includes('в сборе') ? 1.5 : 1)
                    }
                    if (publisher.sides > 1 && publisher.colourType?.match(/глянец/) && !publisher.T)
                        price -= prices['обратки'][stage.stage] * publisher.square * (publisher.T || publisher.description.toLowerCase().includes('в сборе') ? 1.5 : 1) * (publisher.sides - 1)
                }

                if (publisher.stages.find(({ stage, index }) => stage === 'Шлифовка к грунту' && index === 2) && stage.stage === 'Шлифовка к грунту') {
                    if (stage.index === 2 && (publisher.description.toLowerCase().match(Type[2]) || publisher.description.toLowerCase().match(Type[1])) && !publisher.description.toLowerCase().match(Type[0]))
                        if (sFs.find(F => publisher.description.toLowerCase().includes(F)))
                            price += prices['фрезированные'][stage.stage]['сложная'] * publisher.square * (publisher.T || publisher.description.toLowerCase().includes('в сборе') ? 1.5 : 1)
                        else
                            price += prices['фрезированные'][stage.stage]['простая'] * publisher.square * (publisher.T || publisher.description.toLowerCase().includes('в сборе') ? 1.5 : 1)
                } else {
                    if ((publisher.description.toLowerCase().match(Type[2]) || publisher.description.toLowerCase().match(Type[1])) && !publisher.description.toLowerCase().match(Type[0]))
                        if (sFs.find(F => publisher.description.toLowerCase().includes(F)))
                            price += prices['фрезированные'][stage.stage]['сложная'] * publisher.square * (publisher.T || publisher.description.toLowerCase().includes('в сборе') ? 1.5 : 1)
                        else
                            price += prices['фрезированные'][stage.stage]['простая'] * publisher.square * (publisher.T || publisher.description.toLowerCase().includes('в сборе') ? 1.5 : 1)
                    else {
                        if ((publisher.description.toLowerCase().match(Type[2]) || publisher.description.toLowerCase().match(Type[1])) && publisher.description.toLowerCase().match(Type[0]))
                            price += 50 * publisher.square
                        price += prices['прямые'][stage.stage] * publisher.square * (publisher.T || publisher.description.toLowerCase().includes('в сборе') ? 1.5 : 1)
                        if (publisher.thickness === 22 && stage.stage === 'Нанесение грунта')
                            price += prices['прямые']['Нанесение изолятора'] * publisher.square * (publisher.T || publisher.description.toLowerCase().includes('в сборе') ? 1.5 : 1)
                    }
                }
                if (publisher.sides > 1 && publisher.colourType?.match(/глянец2/) && !publisher.T && !publisher.description.toLowerCase().includes('в сборе'))
                    price += prices['обратки'][stage.stage] * publisher.square * (publisher.description.toLowerCase().includes('в сборе') ? 1.5 : 1) * (publisher.sides - 1)
            }
        } else if (stage.stage === 'Шлифовка к покраске') {
            if (publisher.description.toLowerCase().match(/карниз/))
                price += Math.max(publisher.height, publisher.width) / 1000 * publisher.amount * prices['Карнизы МДФ'][stage.stage]
            else {
                if (publisher.H)
                    price += prices['ручки'][stage.stage] * publisher.amount
                if (publisher.stages.find(({ stage, index }) => stage === 'Шлифовка к покраске' && index === 2)) {
                    if (stage.index === 2 && (publisher.description.toLowerCase().match(Type[2]) || publisher.description.toLowerCase().match(Type[1])) && !publisher.description.toLowerCase().match(Type[0]))
                        if (sFs.find(F => publisher.description.toLowerCase().includes(F)))
                            price += prices['фрезированные'][stage.stage]['сложная'] * publisher.square * (publisher.T || publisher.description.toLowerCase().includes('в сборе') ? 1.5 : 1)
                        else
                            price += prices['фрезированные'][stage.stage]['простая'] * publisher.square * (publisher.T || publisher.description.toLowerCase().includes('в сборе') ? 1.5 : 1)
                } else {
                    if ((publisher.description.toLowerCase().match(Type[2]) || publisher.description.toLowerCase().match(Type[1])) && !publisher.description.toLowerCase().match(Type[0]))
                        if (sFs.find(F => publisher.description.toLowerCase().includes(F)))
                            price += prices['фрезированные'][stage.stage]['сложная'] * publisher.square * (publisher.T || publisher.description.toLowerCase().includes('в сборе') ? 1.5 : 1)
                        else
                            price += prices['фрезированные'][stage.stage]['простая'] * publisher.square * (publisher.T || publisher.description.toLowerCase().includes('в сборе') ? 1.5 : 1)
                    else {
                        if ((publisher.description.toLowerCase().match(Type[2]) || publisher.description.toLowerCase().match(Type[1])) && publisher.description.toLowerCase().match(Type[0]))
                            price += 50 * publisher.square
                        price += prices['прямые'][stage.stage] * publisher.square * (publisher.T || publisher.description.toLowerCase().includes('в сборе') ? 1.5 : 1)
                        if (publisher.thickness === 22 && stage.stage === 'Нанесение грунта')
                            price += prices['прямые']['Нанесение изолятора'] * publisher.square * (publisher.T || publisher.description.toLowerCase().includes('в сборе') ? 1.5 : 1)
                    }
                }
                if (publisher.sides > 1 && !publisher.T && !publisher.description.toLowerCase().includes('в сборе'))
                    price += prices['обратки'][stage.stage] * publisher.square * (publisher.description.toLowerCase().includes('в сборе') ? 1.5 : 1) * (publisher.sides - 1)
            }
        } else if (stage.stage === 'Покраска') {
            if (publisher.description.toLowerCase().match(/профиль|метал/))
                price += Math.max(publisher.height, publisher.width) / 1000 * publisher.amount * prices['Алюминиевый профиль']
            else if (publisher.description.toLowerCase().match(/карниз|цоколь/) && publisher.description.toLowerCase().match(/пластик/))
                price += Math.max(publisher.height, publisher.width) / 1000 * publisher.amount * prices['Пластиковый карниз']
            else if (publisher.description.toLowerCase().match(/карниз/))
                price += Math.max(publisher.height, publisher.width) / 1000 * publisher.amount * prices['Карнизы МДФ']['Покраска']
            else {
                if (publisher.H)
                    price += prices['ручки'][stage.stage] * publisher.amount
                if (!stage.index)
                    price += prices['прямые'][stage.stage] * publisher.square * (/* !publisher.T && */ publisher.description.toLowerCase().includes('в сборе') ? 1.5 : 1)
                if (stage.index)
                    price += prices['обратки'][stage.stage] * publisher.square * (/* !publisher.T &&  */publisher.description.toLowerCase().includes('в сборе') ? 1.5 : 1) * (publisher.sides - 1)
                if (publisher.colourType?.match(/лак/) && !stage.index)
                    price += prices['лак']['односторонний'] * publisher.square * (/* !publisher.T &&  */publisher.description.toLowerCase().includes('в сборе') ? 1.5 : 1)
                if (publisher.colourType?.match(/лак2/) && stage.index === 1)
                    price += prices['лак']['двусторонний'] * publisher.square * (/* !publisher.T && */ publisher.description.toLowerCase().includes('в сборе') ? 1.5 : 1)
            }
        } else if (stage.stage === 'Аппликация' && (publisher.sides === 2 || publisher.T || publisher.description.toLowerCase().includes('в сборе') || publisher.description.toLowerCase().match(/лдсп/))) {
            price += prices['упаковка']['Аппликация'] * publisher.square
        } else if (stage.stage === 'Полировка') {
            if (publisher.colourType?.match(/глянец2/))
                price += prices['глянец']['двусторонний'] * publisher.square
            else if (publisher.colourType?.match(/глянец/))
                price += prices['глянец']['односторонний'] * publisher.square
        } else if (stage.stage === 'Упаковка') {
            if ((publisher.sides === 1 || publisher.sides === 1.5) && !publisher.colourType?.match(/глянец/))
                price += prices['упаковка']['чистка'] * publisher.square
            price += prices['упаковка'][stage.stage] * publisher.square
        }

        stage.price = price
    }
}

module.exports = getPublisherSalary
