const prices = require('./prices.json')

const Type = [/прямой/, /фрез/, /f/, /фас кл/, /фасад клиента/, /тумба/]
const eFs = ['f-01', 'f-03', 'f-04', 'f-08', 'f-10', 'f-11', 'f-15', 'f-20', 'f-21']
const sFs = ['f-02', 'f-05', 'f-06', 'f-07', 'f-09', 'f-12', 'f-13', 'f-14', 'f-16', 'f-17', 'f-18', 'f-19']

function getPublisherSalary(publisher) {
    for (let stage of publisher.stages) {
        let price = 0
        if (stage.stage === 'Подготовка к изолятору') {
            if (publisher.description.toLowerCase().match(/карниз/))
                price += Math.max(publisher.height, publisher.width) / 1000 * publisher.amount * prices['Карнизы МДФ'][stage.stage] * (publisher.description.toLowerCase().includes('в сборе') ? 1.5 : 1)
            else {
                if (stage.H)
                    price += prices['ручки']['Подготовка к грунту'] * publisher.amount
                else if (publisher.H && stage.stage === 'Подготовка к изолятору' && publisher.stages.find(({ H }) => H).stage === 'Подготовка к изолятору') {
                    if (publisher.category.toLowerCase().includes('f') || sFs.find(F => publisher.description.toLowerCase().includes(F)))
                        price -= prices['фрезированные'][stage.stage]['сложная'] * publisher.square
                    else
                        price -= prices['фрезированные'][stage.stage]['простая'] * publisher.square
                }
                if (publisher.category.toLowerCase().includes('f') || sFs.find(F => publisher.description.toLowerCase().includes(F)))
                    price += prices['фрезированные'][stage.stage]['сложная'] * publisher.square
                else
                    price += prices['фрезированные'][stage.stage]['простая'] * publisher.square
            }
        } else if (stage.stage === 'Нанесение изолятора') {
            if (publisher.description.toLowerCase().match(/карниз/))
                price += Math.max(publisher.height, publisher.width) / 1000 * publisher.amount * prices['Карнизы МДФ'][stage.stage] * (publisher.description.toLowerCase().includes('в сборе') ? 1.5 : 1)
            else
                price += prices['фрезированные'][stage.stage] * publisher.square
        } else if (stage.stage === 'Подготовка к грунту' || stage.stage === 'Нанесение грунта') {
            if (publisher.description.toLowerCase().match(/карниз/))
                price += Math.max(publisher.height, publisher.width) / 1000 * publisher.amount * prices['Карнизы МДФ'][stage.stage] * (publisher.description.toLowerCase().includes('в сборе') ? 1.5 : 1)
            else {
                if (stage.H)
                    price += prices['ручки'][stage.stage] * publisher.amount
                else if (publisher.H && stage.stage === 'Подготовка к грунту' && publisher.stages.find(({ H }) => H).stage === 'Подготовка к грунту') {
                    if ((publisher.category.toLowerCase().includes('f') || publisher.description.toLowerCase().match(Type[2]) || publisher.description.toLowerCase().match(Type[1])) && !publisher.description.toLowerCase().match(Type[0]))
                        if (publisher.category.toLowerCase().includes('f') || sFs.find(F => publisher.description.toLowerCase().includes(F)))
                            price -= prices['фрезированные'][stage.stage]['сложная'] * publisher.square * (publisher.T || publisher.description.toLowerCase().includes('в сборе') ? 1.5 : 1)
                        else
                            price -= prices['фрезированные'][stage.stage]['простая'] * publisher.square * (publisher.T || publisher.description.toLowerCase().includes('в сборе') ? 1.5 : 1)
                    else {
                        if ((publisher.category.toLowerCase().includes('f') || publisher.description.toLowerCase().match(Type[2]) || publisher.description.toLowerCase().match(Type[1])) && publisher.description.toLowerCase().match(Type[0]))
                            price -= 50 * publisher.square
                        price -= prices['прямые'][stage.stage] * publisher.square * (publisher.T || publisher.description.toLowerCase().includes('в сборе') ? 1.5 : 1)
                        if (publisher.thickness === 22 && stage.stage === 'Нанесение грунта')
                            price -= prices['прямые']['Нанесение изолятора'] * publisher.square * (publisher.T || publisher.description.toLowerCase().includes('в сборе') ? 1.5 : 1)
                    }
                    if (publisher.sides > 1 && publisher.colourType && publisher.colourType?.match(/глянец/) && !publisher.T)
                        price -= prices['обратки'][stage.stage] * publisher.square * (publisher.T || publisher.description.toLowerCase().includes('в сборе') ? 1.5 : 1) * (publisher.sides - 1)
                }

                if (publisher.stages.find(({ stage, index }) => stage === 'Подготовка к грунту' && index === 2) && stage.stage === 'Подготовка к грунту') {
                    if (stage.index === 2 && (publisher.description.toLowerCase().match(Type[2]) || publisher.description.toLowerCase().match(Type[1])) && !publisher.description.toLowerCase().match(Type[0]))
                        if (publisher.category.toLowerCase().includes('f') || sFs.find(F => publisher.description.toLowerCase().includes(F)))
                            price += prices['фрезированные'][stage.stage]['сложная'] * publisher.square * (publisher.T || publisher.description.toLowerCase().includes('в сборе') ? 1.5 : 1)
                        else
                            price += prices['фрезированные'][stage.stage]['простая'] * publisher.square * (publisher.T || publisher.description.toLowerCase().includes('в сборе') ? 1.5 : 1)
                } else {
                    if ((publisher.description.toLowerCase().match(Type[2]) || publisher.description.toLowerCase().match(Type[1])) && !publisher.description.toLowerCase().match(Type[0]))
                        if (publisher.category.toLowerCase().includes('f') || sFs.find(F => publisher.description.toLowerCase().includes(F)))
                            price += prices['фрезированные'][stage.stage]['сложная'] * publisher.square * (publisher.T || publisher.description.toLowerCase().includes('в сборе') ? 1.5 : 1)
                        else
                            price += prices['фрезированные'][stage.stage]['простая'] * publisher.square * (publisher.T || publisher.description.toLowerCase().includes('в сборе') ? 1.5 : 1)
                    else {
                        if ((publisher.category.toLowerCase().includes('f') || publisher.description.toLowerCase().match(Type[2]) || publisher.description.toLowerCase().match(Type[1])) && publisher.description.toLowerCase().match(Type[0]))
                            price += 50 * publisher.square
                        price += prices['прямые'][stage.stage] * publisher.square * (publisher.T || publisher.description.toLowerCase().includes('в сборе') ? 1.5 : 1)
                        if (publisher.thickness === 22 && stage.stage === 'Нанесение грунта')
                            price += prices['прямые']['Нанесение изолятора'] * publisher.square * (publisher.T || publisher.description.toLowerCase().includes('в сборе') ? 1.5 : 1)
                    }
                }
                if (publisher.sides > 1 && publisher.colourType && (publisher.colourType?.match(/глянец2/) || stage.stage === 'Подготовка к грунту') && !publisher.T && !publisher.description.toLowerCase().includes('в сборе'))
                    price += prices['обратки'][stage.stage] * publisher.square * (publisher.description.toLowerCase().includes('в сборе') ? 1.5 : 1) * (publisher.sides - 1)
            }
        } else if (stage.stage === 'Подготовка к покраске') {
            if (publisher.description.toLowerCase().match(/карниз/))
                price += Math.max(publisher.height, publisher.width) / 1000 * publisher.amount * prices['Карнизы МДФ'][stage.stage] * (publisher.description.toLowerCase().includes('в сборе') ? 1.5 : 1)
            else {
                if (publisher.H)
                    price += prices['ручки'][stage.stage] * publisher.amount
                if (publisher.stages.find(({ stage, index }) => stage === 'Подготовка к покраске' && index === 2)) {
                    if (stage.index === 2 && (publisher.category.toLowerCase().includes('f') || publisher.description.toLowerCase().match(Type[2]) || publisher.description.toLowerCase().match(Type[1])) && !publisher.description.toLowerCase().match(Type[0]))
                        if (publisher.category.toLowerCase().includes('f') || sFs.find(F => publisher.description.toLowerCase().includes(F)))
                            price += prices['фрезированные'][stage.stage]['сложная'] * publisher.square * (publisher.T || publisher.description.toLowerCase().includes('в сборе') ? 1.5 : 1)
                        else
                            price += prices['фрезированные'][stage.stage]['простая'] * publisher.square * (publisher.T || publisher.description.toLowerCase().includes('в сборе') ? 1.5 : 1)
                } else {
                    if ((publisher.category.toLowerCase().includes('f') || publisher.description.toLowerCase().match(Type[2]) || publisher.description.toLowerCase().match(Type[1])) && !publisher.description.toLowerCase().match(Type[0]))
                        if (publisher.category.toLowerCase().includes('f') || sFs.find(F => publisher.description.toLowerCase().includes(F)))
                            price += prices['фрезированные'][stage.stage]['сложная'] * publisher.square * (publisher.T || publisher.description.toLowerCase().includes('в сборе') ? 1.5 : 1)
                        else
                            price += prices['фрезированные'][stage.stage]['простая'] * publisher.square * (publisher.T || publisher.description.toLowerCase().includes('в сборе') ? 1.5 : 1)
                    else {
                        if ((publisher.category.toLowerCase().includes('f') || publisher.description.toLowerCase().match(Type[2]) || publisher.description.toLowerCase().match(Type[1])) && publisher.description.toLowerCase().match(Type[0]))
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
            if (publisher.description.toLowerCase().match(/профиль|метал/)) {
                price += Math.max(publisher.height, publisher.width) / 1000 * publisher.amount * prices['Алюминиевый профиль']["Покраска"] * publisher.sides
                if (publisher.colourType && publisher.colourType?.match(/лак/))
                    price += Math.max(publisher.height, publisher.width) / 1000 * publisher.amount * prices['Алюминиевый профиль']['Лак']
                if (publisher.colourType && publisher.colourType?.match(/лак2/))
                    price += Math.max(publisher.height, publisher.width) / 1000 * publisher.amount * prices['Алюминиевый профиль']['Лак']
            }
            else if (publisher.description.toLowerCase().match(/карниз|цоколь/) && publisher.description.toLowerCase().match(/пластик/)) {
                price += Math.max(publisher.height, publisher.width) / 1000 * publisher.amount * prices['Пластиковый карниз']["Покраска"] * publisher.sides
                if (publisher.colourType && publisher.colourType?.match(/лак/))
                    price += Math.max(publisher.height, publisher.width) / 1000 * publisher.amount * prices['Пластиковый карниз']['Лак']
                if (publisher.colourType && publisher.colourType?.match(/лак2/))
                    price += Math.max(publisher.height, publisher.width) / 1000 * publisher.amount * prices['Пластиковый карниз']['Лак']
            }
            else if (publisher.description.toLowerCase().match(/карниз/)) {
                price += Math.max(publisher.height, publisher.width) / 1000 * publisher.amount * prices['Карнизы МДФ']['Покраска'] * publisher.sides
                if (publisher.colourType && publisher.colourType?.match(/лак/))
                    price += Math.max(publisher.height, publisher.width) / 1000 * publisher.amount * prices['Карнизы МДФ']['Лак']
                if (publisher.colourType && publisher.colourType?.match(/лак2/))
                    price += Math.max(publisher.height, publisher.width) / 1000 * publisher.amount * prices['Карнизы МДФ']['Лак']
            }
            else {
                if (publisher.H)
                    price += prices['ручки'][stage.stage] * publisher.amount
                if (!stage.index)
                    price += prices['прямые'][stage.stage] * publisher.square * (/* !publisher.T && */ publisher.description.toLowerCase().includes('в сборе') ? 1.5 : 1)
                if (stage.index)
                    price += prices['обратки'][stage.stage] * publisher.square * (/* !publisher.T &&  */publisher.description.toLowerCase().includes('в сборе') ? 1.5 : 1) * (publisher.sides - 1)
                if (publisher.colourType && publisher.colourType?.match(/лак/) && !stage.index)
                    price += prices['лак']['односторонний'] * publisher.square * (/* !publisher.T &&  */publisher.description.toLowerCase().includes('в сборе') ? 1.5 : 1)
                if (publisher.colourType && publisher.colourType?.match(/лак2/) && stage.index === 1)
                    price += prices['лак']['двусторонний'] * publisher.square * (/* !publisher.T && */ publisher.description.toLowerCase().includes('в сборе') ? 1.5 : 1)
            }
        } else if (stage.stage === 'Аппликация' && (publisher.sides === 2 || publisher.T || publisher.description.toLowerCase().includes('в сборе') || publisher.description.toLowerCase().match(/лдсп/))) {
            price += prices['упаковка']['Аппликация'] * publisher.square
        } else if (stage.stage === 'Полировка') {
            if (publisher.colourType && publisher.colourType?.match(/глянец2/))
                price += prices['глянец']['двусторонний'] * publisher.square
            else if (publisher.colourType && publisher.colourType?.match(/глянец/))
                price += prices['глянец']['односторонний'] * publisher.square
        } else if (stage.stage === 'Упаковка') {
            if ((publisher.sides === 1 || publisher.sides === 1.5) && publisher.colourType && publisher.colourType?.match(/глянец/))
                price += prices['упаковка']['чистка'] * publisher.square
            price += prices['упаковка'][stage.stage] * publisher.square
        }

        stage.price = price
    }
}

module.exports = getPublisherSalary
