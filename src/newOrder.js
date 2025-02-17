const XLSX = require('xlsx')

const Сarpentry = [/р/, /п/, /с/, /б/]
const Type = [/прямой/, /фрез/, /f/, /фас кл/, /фасад клиента/, /тумба/, /в сборе/]
const eFs = ['f-01', 'f-02', 'f-03', 'f-04', 'f-06', 'f-07', 'f-08', 'f-09', 'f-10', 'f-11', 'f-15', 'f-16', 'f-18', 'f-19', 'F20', 'f-21']
const sFs = ['f-05', 'f-12', 'f-13', 'f-14', 'f-17']

const Second = 1000
const Minute = 60 * Second
const Hour = 60 * Minute
const Day = 24 * Hour

const prices = require('./prices.json')
const getPublisherSalary = require('./getPublisherSalary')

function getTimeZoneOffset(date, timeZone) {
    let iso = date.toLocaleString('en-CA', { timeZone, hour12: false }).replace(', ', 'T')
    iso += '.' + date.getMilliseconds().toString().padStart(3, '0')
    const lie = new Date(iso + 'Z')
    return -(lie - date) / 60 / 1000
}

function ExcelDateToJSDate(serial) {
    var utc_days = Math.floor(serial - 25569)
    var utc_value = utc_days * 86400
    var date_info = new Date(+new Date(utc_value * 1000) + (getTimeZoneOffset(new Date(), 'Europe/Moscow') - getTimeZoneOffset(new Date())) * Minute)

    var fractional_day = serial - Math.floor(serial) + 0.0000001

    var total_seconds = Math.floor(86400 * fractional_day)

    var seconds = total_seconds % 60

    total_seconds -= seconds

    var hours = Math.floor(total_seconds / (60 * 60))
    var minutes = Math.floor(total_seconds / 60) % 60

    return +new Date(+new Date(date_info.getFullYear(), date_info.getMonth(), date_info.getDate(), hours, minutes, seconds) + (getTimeZoneOffset(new Date(), 'Europe/Moscow') - getTimeZoneOffset(new Date())) * Minute)
}

function round(num, round = 0, func = Math.round) {
    return func(num * 10 ** round) / 10 ** round
}

function newOrder(file) {
    return new Promise(resolve => {
        let workbook
        if (typeof file === 'string') {
            console.log('Читаю файл', file)
            workbook = XLSX.readFile(file)
        } else
            workbook = XLSX.read(file)
        let ws = workbook.Sheets['ЗАКАЗ-НАРЯД']

        let length = JSON.parse(ws['!ref'].match(/\:\w\d+/)[0].match(/\d+/)[0]) - JSON.parse(ws['!ref'].match(/\w\d+\:/)[0].match(/\d+/)[0]) + 1
        let height = ws['!ref'].match(/\:\w\d+/)[0].match(/\w/)[0].charCodeAt(0) - ws['!ref'].match(/\w\d+\:/)[0].match(/\w/)[0].charCodeAt(0) - 1

        let rows = [...new Array(length)].map(_ => {
            return [...new Array(height)].map(_ => null)
        })
        Object.keys(ws).filter(id => id.match(/\w\d+/)).forEach(key => {
            rows[JSON.parse(key.match(/\w\d+/)[0].match(/\d+/)[0]) - JSON.parse(ws['!ref'].match(/\w\d+\:/)[0].match(/\d+/)[0])][key.match(/\w\d+/)[0].match(/\w/)[0].charCodeAt(0) - ws['!ref'].match(/\w\d+\:/)[0].match(/\w/)[0].charCodeAt(0)] = ws[key].v
        })
        rows = rows.filter(row => row[0])
            .filter(row => typeof row[2] === 'number' && row[2] && row[2] != 23 || row[2] === 'Ширина')

        rows = rows.map(row => {
            if (row[rows[0].findIndex(row => row === 'Примечание')] === 0)
                row[rows[0].findIndex(row => row === 'Примечание')] = ''
            return row
        })

        let publishers = rows.slice(1, rows.length).map(row => {
            // if (JSON.parse(ws['N7']?.v ? ws['N7']?.v?.match(/\d+/)[0] : ws['M7']?.v?.match(/\d+/)[0]) == 28)
            //     console.log(row[rows[0].findIndex(row => row === '№')], row[rows[0].findIndex(row => row === 'Примечание')])
            let square = rows.filter(row => row[rows[0].findIndex(row => row === 'Примечание')]?.toLowerCase().match(Type[2]) || row[rows[0].findIndex(row => row === 'Примечание')]?.toLowerCase().match(Type[1]))
                .map(row => row[rows[0].findIndex(row => row === 'Площадь')]).reduce((a, b) => a + b, 0) * 20 / 7 + rows.filter(row => !row[rows[0].findIndex(row => row === 'Примечание')]?.toLowerCase().match(Type[2]) && !row[rows[0].findIndex(row => row === 'Примечание')]?.toLowerCase().match(Type[1]))
                    .map(row => row[rows[0].findIndex(row => row === 'Площадь')]).splice(1,).reduce((a, b) => a + b, 0)
            let stages = [
                {
                    term: 3,
                    stage: 'Распил'
                },
                {
                    term: Math.ceil(square / 20),
                    stage: 'Подготовка к грунту'
                },
                {
                    term: 1,
                    stage: 'Нанесение грунта'
                },
                {
                    term: Math.ceil(square / 20),
                    stage: 'Подготовка к покраске'
                },
                {
                    term: 1,
                    stage: 'Покраска',
                    index: 0
                },
                {
                    term: 2,
                    stage: 'Упаковка'
                },
                {
                    term: 1,
                    stage: 'Оплата'
                },
                {
                    term: 1,
                    stage: 'Отгрузка'
                }
            ]
            if ((row[rows[0].findIndex(row => row === 'Примечание')]?.toLowerCase().match(Type[2]) || row[rows[0].findIndex(row => row === 'Примечание')]?.toLowerCase().match(Type[1])) && !row[rows[0].findIndex(row => row === 'Примечание')]?.toLowerCase().match(Type[0])) {
                let square = rows.filter(row => row[rows[0].findIndex(row => row === 'Примечание')]?.toLowerCase().match(Type[2]) || row[rows[0].findIndex(row => row === 'Примечание')]?.toLowerCase().match(Type[1]))
                    .map(row => row[rows[0].findIndex(row => row === 'Площадь')]).reduce((a, b) => a + b, 0)
                stages[stages.findIndex(stage => stage.stage === 'Распил')].term += 2
                stages.splice(stages.findIndex(stage => stage.stage === 'Распил') + 1, 0,
                    {
                        term: Math.ceil(square / 7),
                        stage: 'Подготовка к изолятору'
                    },
                    {
                        term: 1,
                        stage: 'Нанесение изолятора'
                    }
                )
                stages.splice(stages.findIndex(stage => stage.stage === 'Подготовка к грунту') + 1, 0, {
                    term: 1,
                    stage: 'Подготовка к грунту',
                    index: 2
                })
                stages.splice(stages.findIndex(stage => stage.stage === 'Подготовка к покраске') + 1, 0, {
                    term: 1,
                    stage: 'Подготовка к покраске',
                    index: 2
                })
            }
            if (row[rows[0].findIndex(row => row === 'Примечание')]?.toLowerCase().match(/шпон/)) {
                stages.splice(stages.findIndex(stage => stage.stage === 'Подготовка к грунту'), 1)
                stages.splice(stages.findIndex(stage => stage.stage === 'Нанесение грунта'), 1)
            } else if (row[rows[0].findIndex(row => row === 'Примечание')]?.toLowerCase().match(/профиль|метал|пластик/)) {
                stages.splice(stages.findIndex(stage => stage.stage === 'Подготовка к грунту'), 1)
                stages.splice(stages.findIndex(stage => stage.stage === 'Нанесение грунта'), 1)
                stages.splice(stages.findIndex(stage => stage.stage === 'Подготовка к покраске'), 1)
            } else if (row[rows[0].findIndex(row => row === 'Примечание')]?.toLowerCase().match(/карниз/)) {
                stages.splice(stages.findIndex(stage => stage.stage === 'Распил') + 1, 0,
                    {
                        term: 1,
                        stage: 'Подготовка к изолятору'
                    },
                    {
                        term: 1,
                        stage: 'Нанесение изолятора'
                    }
                )
            }
            if (row[rows[0].findIndex(row => row === 'Примечание')]?.toLowerCase().match(/лдсп/)) {
                stages.splice(stages.findIndex(stage => stage.stage === 'Подготовка к грунту'), 1)
                stages.splice(stages.findIndex(stage => stage.stage === 'Нанесение грунта'), 1)
                stages.splice(stages.findIndex(stage => stage.stage === 'Покраска'), 0, {
                    term: 1,
                    stage: 'Аппликация'
                })
            }
            if (row[rows[0].findIndex(row => row === 'Примечание')]?.toLowerCase().match(Type[5]))
                stages.splice(stages.findIndex(stage => stage.stage === 'Покраска'), 0, {
                    term: 1,
                    stage: 'Аппликация'
                })

            let P = row[rows[0].findIndex(row => row === 'Категория')]?.toString().toLowerCase().match(Сarpentry[1]) ? 1 : 0
            let H = row[rows[0].findIndex(row => row === 'Категория')]?.toString().toLowerCase().match(Сarpentry[0]) ? 1 : 0
            let G = row[rows[0].findIndex(row => row === 'Категория')]?.toString().toLowerCase().match(Сarpentry[2]) ? 1 : 0
            let T = (/* row[rows[0].findIndex(row => row === 'Примечание')]?.toLowerCase().match(Type[6]) ||  */row[rows[0].findIndex(row => row === 'Примечание')]?.toLowerCase().match(Type[5])) && (typeof row[rows[0].findIndex(row => row === 'Кол-во сторон')] === 'number' ? round(row[rows[0].findIndex(row => row === 'Кол-во сторон')], 0) : row[rows[0].findIndex(row => row === 'Кол-во сторон')] === '-' ? 0 : 2) === 2 ? 1 : 0
            let Paint = row[rows[0].findIndex(row => row === 'Категория')]?.toString().toLowerCase().match(Сarpentry[3]) ? 1 : 0
            if (Paint)
                stages.splice(stages.findIndex(stage => stage.stage === 'Покраска'), 1)

            if (H) {
                if (row[rows[0].findIndex(row => row === 'Примечание')]?.toLowerCase().match(/шпон/)) {
                } else if ((row[rows[0].findIndex(row => row === 'Примечание')]?.toLowerCase().match(Type[2]) || row[rows[0].findIndex(row => row === 'Примечание')]?.toLowerCase().match(Type[1])) && !row[rows[0].findIndex(row => row === 'Примечание')]?.toLowerCase().match(Type[0]))
                    stages.splice(stages.findIndex(stage => stage.stage === 'Подготовка к изолятору') + 1, 0, {
                        term: 1,
                        stage: 'Подготовка к изолятору',
                        index: 1,
                        H
                    })
                else
                    stages.splice(stages.findIndex(stage => stage.stage === 'Подготовка к грунту') + 1, 0, {
                        term: 1,
                        stage: 'Подготовка к грунту',
                        index: 1,
                        H
                    })
                if (stages.find(({ stage }) => stage === 'Нанесение грунта'))
                    stages.find(({ stage }) => stage === 'Нанесение грунта').H = 1
            }

            if (P)
                stages.splice(stages.findIndex(stage => stage.stage === 'Распил') + 1, 0, {
                    term: 1,
                    stage: 'Присадка'
                })
            if (G)
                stages.splice(stages.findIndex(stage => stage.stage === 'Распил') + 1, 0, {
                    term: 1,
                    stage: 'Склейка'
                })

            if (row[rows[0].findIndex(row => row === 'МДФ')] === '-' || row[rows[0].findIndex(row => row === 'Примечание')]?.toLowerCase().match(Type[3]) || row[rows[0].findIndex(row => row === 'Примечание')]?.toLowerCase().match(Type[4]))
                stages.splice(stages.findIndex(stage => stage.stage === 'Распил'), 1)

            if (row[rows[0].findIndex(row => row === 'Цвет')] === '-')
                stages.splice(stages.findIndex(stage => stage.stage === 'Покраска'), 1)
            else if (row[rows[0].findIndex(row => row === 'Тип краски')]?.toString().toLowerCase().includes('глянец') && !Paint) {
                let squere = rows.filter(row => row[rows[0].findIndex(row => row === 'Тип краски')].toLowerCase().includes('глянец'))
                    .map(row => row[rows[0].findIndex(row => row === 'Площадь')]).reduce((a, b) => a + b, 0)
                stages[stages.findIndex(stage => stage.stage === 'Покраска')]
                stages.splice(stages.findIndex(stage => stage.stage === 'Покраска') + 1, 0,
                    {
                        term: Math.ceil(square / 4),
                        stage: 'Полировка'
                    }
                )
                if (row[rows[0].findIndex(row => row === 'Кол-во сторон')] === 2 && !row[rows[0].findIndex(row => row === 'Тип краски')].toLowerCase().match(/глянец2/))
                    stages[stages.findIndex(stage => stage.stage === 'Нанесение грунта')].term++
                // if (squere > 10 || row[rows[0].findIndex(row => row === 'Кол-во сторон')] === 2 || row[rows[0].findIndex(row => row === 'Кол-во сторон')] === 'все' || row[rows[0].findIndex(row => row === 'Тип краски')].toLowerCase().match(/лак/))
                //     stages[stages.findIndex(stage => stage.stage === 'Полировка')].term += 2
                if (row[rows[0].findIndex(row => row === 'Кол-во сторон')] === 'все')
                    row[rows[0].findIndex(row => row === 'Кол-во сторон')] = 2
                let paintstages = row[rows[0].findIndex(row => row === 'Кол-во сторон')] - 1 + ((row[rows[0].findIndex(row => row === 'Цвет')].toString() || undefined)?.toLowerCase().split(',').length || 1) - 1
                for (let i = 0; i < Math.ceil(paintstages); i++)
                    stages.splice(stages.findIndex(stage => stage.stage === 'Покраска'), 0, {
                        term: 1,
                        stage: 'Покраска',
                        index: Math.ceil(paintstages) - i
                    }, {
                        term: 1,
                        stage: 'Аппликация',
                        index: Math.ceil(paintstages) - i - 1
                    })
                if (squere > 10 && stages.filter(({ stage }) => stage === 'Покраска').length === 1)
                    stages[stages.findIndex(stage => stage.stage === 'Покраска')].term++
            } else if (!Paint) {
                let squere = rows.filter(row => row[rows[0].findIndex(row => row === 'Тип краски')]?.toString().toLowerCase().includes('мат'))
                    .map(row => row[rows[0].findIndex(row => row === 'Площадь')]).reduce((a, b) => a + b, 0)
                let paintstages = row[rows[0].findIndex(row => row === 'Кол-во сторон')] - T - 1 + ((row[rows[0].findIndex(row => row === 'Цвет')].toString() || undefined)?.toLowerCase().split(',').length || 1) - 1
                for (let i = 0; i < Math.ceil(paintstages); i++)
                    stages.splice(stages.findIndex(stage => stage.stage === 'Покраска'), 0, {
                        term: 1,
                        stage: 'Покраска',
                        index: Math.ceil(paintstages) - i
                    }, {
                        term: 1,
                        stage: 'Аппликация',
                        index: Math.ceil(paintstages) - i - 1
                    })
                if (squere > 10 && stages.filter(({ stage }) => stage === 'Покраска').length === 1)
                    stages[stages.findIndex(stage => stage.stage === 'Покраска')].term++
            }

            return {
                number: row[rows[0].findIndex(row => row === '№')],
                height: row[rows[0].findIndex(row => row === 'Высота')],
                amount: row[rows[0].findIndex(row => row === 'Кол-во')],
                width: row[rows[0].findIndex(row => row === 'Ширина')],
                square: row[rows[0].findIndex(row => row === 'Площадь')],
                square2: ((typeof row[rows[0].findIndex(row => row === 'Кол-во сторон')] === 'number' ? typeof row[rows[0].findIndex(row => row === 'Кол-во сторон')] : 1) - 1) * row[rows[0].findIndex(row => row === 'Площадь')],
                thickness: typeof row[rows[0].findIndex(row => row === 'МДФ')] === 'number' ? row[rows[0].findIndex(row => row === 'МДФ')] : 0,
                description: row[rows[0].findIndex(row => row === 'Примечание')],
                T,
                H,
                colour: row[rows[0].findIndex(row => row === 'Цвет')],
                category: row[rows[0].findIndex(row => row === 'Категория')] || '',
                colourType: row[rows[0].findIndex(row => row === 'Тип краски')],
                sides: typeof row[rows[0].findIndex(row => row === 'Кол-во сторон')] === 'number' ? row[rows[0].findIndex(row => row === 'Кол-во сторон')] : 2,
                radius: row[rows[0].findIndex(row => row === 'Радиус')] === 'мин' ? 1 : typeof row[rows[0].findIndex(row => row === 'МДФ')] === 'number' ? row[rows[0].findIndex(row => row === 'МДФ')] : 0,
                stages
            }
        }).map(publisher => {
            getPublisherSalary(publisher)
            return publisher
        })
        // console.log(publishers.map(({ stages, square }) => stages.filter(({ stage }) => stage.includes('Подготовка')).map(({ stage, price }) => `${stage} (${square}) = ${price}`).join(', ')).join('\r\n'))
        if (publishers.includes(undefined))
            return
        // if (publishers.find(({ stages }) => stages.find(({ P }) => P)) && publishers.find(({ stages }) => stages.find(({ G }) => G)))
        //     publishers = publishers.map(publisher => {
        //         publisher.stages = publisher.stages.map(stage => {
        //             if (stage.P && !stage.G || stage.G && !stage.P)
        //                 stage.term += 1
        //             return stage
        //         })
        //         return publisher
        //     })

        // if (!ws['N8'])
        //     return resolve('Нет даты поступления в работу, пожалуйста добавьте ее прежде чем загружать заказ в систему')
        // if (!workbook.Sheets['СЧЕТ']['O' + JSON.parse(Object.keys(workbook.Sheets['СЧЕТ']).find(key => workbook.Sheets['СЧЕТ'][key]?.v === 'Дата готовности заказа')?.match(/\d+/)?.[0])])
        //     return resolve('Нет даты окончания работы, пожалуйста добавьте ее прежде чем загружать заказ в систему')

        let created = ws['N8']?.v ? ExcelDateToJSDate(ws['N8']?.v) : undefined
        if (!created)
            created = ws['M8']?.v ? ExcelDateToJSDate(ws['M8']?.v) : undefined
        let lastDate = created + 31 * Day
        // let lastDate = workbook.Sheets['СЧЕТ']['O' + JSON.parse(Object.keys(workbook.Sheets['СЧЕТ']).find(key => workbook.Sheets['СЧЕТ'][key].v === 'Дата готовности заказа').match(/\d+/)[0])]?.v ? ExcelDateToJSDate(workbook.Sheets['СЧЕТ']['O' + JSON.parse(Object.keys(workbook.Sheets['СЧЕТ']).find(key => workbook.Sheets['СЧЕТ'][key].v === 'Дата готовности заказа').match(/\d+/)[0])]?.v) : undefined
        // if (!lastDate)
        //     lastDate = workbook.Sheets['СЧЕТ']['N' + JSON.parse(Object.keys(workbook.Sheets['СЧЕТ']).find(key => workbook.Sheets['СЧЕТ'][key].v === 'Дата готовности заказа').match(/\d+/)[0])]?.v ? ExcelDateToJSDate(workbook.Sheets['СЧЕТ']['N' + JSON.parse(Object.keys(workbook.Sheets['СЧЕТ']).find(key => workbook.Sheets['СЧЕТ'][key].v === 'Дата готовности заказа').match(/\d+/)[0])]?.v) : undefined

        let stages = [
            {
                term: 0,
                stage: 'Распил'
            },
            {
                term: 0,
                stage: 'Присадка'
            },
            {
                term: 0,
                stage: 'Склейка'
            },
            {
                term: 0,
                stage: 'Подготовка к изолятору'
            },
            {
                term: 0,
                stage: 'Нанесение изолятора'
            },
            {
                term: 0,
                stage: 'Подготовка к грунту'
            },
            {
                term: 0,
                stage: 'Нанесение грунта'
            },
            {
                term: 0,
                stage: 'Подготовка к покраске'
            },
            {
                term: 0,
                stage: 'Покраска',
                weekend: true,
                index: 0
            },
            {
                term: 0,
                stage: 'Полировка'
            },
            {
                term: 0,
                stage: 'Упаковка'
            },
            {
                term: 0,
                stage: 'Оплата'
            },
            {
                term: 0,
                stage: 'Отгрузка'
            }
        ]
        if (publishers.find(({ stages }) => stages.filter(({ stage }) => stage === 'Подготовка к изолятору').length === 2))
            stages.splice(stages.findIndex(stage => stage.stage === 'Подготовка к изолятору') + 1, 0, {
                term: 0,
                stage: 'Подготовка к изолятору',
                H: true,
                index: 1
            })
        else if (publishers.find(({ stages }) => stages.find(({ stage, index }) => stage === 'Подготовка к грунту' && index === 1)))
            stages.splice(stages.findIndex(stage => stage.stage === 'Подготовка к грунту') + 1, 0, {
                term: 0,
                stage: 'Подготовка к грунту',
                H: true,
                index: 1
            })
        if (publishers.find(({ stages }) => stages.find(({ stage, index }) => stage === 'Подготовка к грунту' && index === 2)))
            stages.splice(stages.findIndex(stage => stage.stage === 'Подготовка к грунту') + 1, 0, {
                term: 0,
                stage: 'Подготовка к грунту',
                index: 2
            })
        if (publishers.find(({ stages }) => stages.find(({ stage, index }) => stage === 'Подготовка к покраске' && index === 2)))
            stages.splice(stages.findIndex(stage => stage.stage === 'Подготовка к покраске') + 1, 0, {
                term: 0,
                stage: 'Подготовка к покраске',
                index: 2
            })
        if (publishers.find(({ stages }) => stages.find(({ stage, index }) => stage === 'Аппликация' && index === undefined)))
            stages.splice(stages.findIndex(stage => stage.stage === 'Покраска'), 0, {
                term: 1,
                stage: 'Аппликация',
                weekend: true
            })

        let paintstages = Math.max(...publishers.map(({ stages }) => Math.max(...stages.map(({ index, stage }) => stage === 'Покраска' ? (index || 0) : 0))))
        for (let i = 0; i < paintstages; i++)
            stages.splice(stages.findIndex(stage => stage.stage === 'Покраска'), 0, {
                term: 1,
                stage: 'Покраска',
                index: paintstages - i,
                weekend: true
            }, {
                term: 1,
                stage: 'Аппликация',
                index: paintstages - i - 1,
                weekend: true
            })

        stages = stages.filter(stage => publishers.find(({ stages }) => stages.find(item => item.stage === stage.stage)))
        stages = stages.map(stage => {
            let value = round(publishers.filter(({ stages }) => stages.find(item => item.stage === stage.stage && item.index === stage.index)).map(({ square }) => square).reduce((a, b) => a + b, 0), 2)
            let price = publishers.map(({ stages }) => stages.find(item => item.stage === stage.stage && item.index === stage.index)).filter(stage => stage).map(({ price }) => price).reduce((a, b) => a + b, 0)
            if (price != 0 /* && publishers.filter(({ H }) => H).length === publishers.length || publishers.filter(({ H }) => !H).length === publishers.length */)
                if (value <= 0.5) {
                    if (stage.stage.includes('Подготовка') || stage.stage === 'Нанесение грунта')
                        if (publishers.filter(({ description, category }) => (category.toLowerCase().includes('f') || description?.toLowerCase().match(Type[2]) || description?.toLowerCase().match(Type[1])) && !description?.toLowerCase().match(Type[0])).length === publishers.length) {
                            if (publishers.filter(({ description, category }) => category.toLowerCase().includes('f') || sFs.find(sF => description.toLowerCase().includes(sF))).length === publishers.length)
                                price = 0.5 * prices['фрезированные'][stage.stage]['сложная']
                            else if (publishers.filter(({ description, category }) => !category.toLowerCase().includes('f') && !sFs.find(sF => description.toLowerCase().includes(sF))).length === publishers.length)
                                price = 0.5 * prices['фрезированные'][stage.stage]['простая']
                        } else if (publishers.filter(({ description }) => (description?.toLowerCase().match(Type[2]) || description?.toLowerCase().match(Type[1])) && description?.toLowerCase().match(Type[0])).length === publishers.length)
                            price = 0.5 * (prices['прямые'][stage.stage] + 50)
                        else if (publishers.filter(({ description }) => !(description?.toLowerCase().match(Type[2]) || description?.toLowerCase().match(Type[1])) && description?.toLowerCase().match(Type[0])).length === publishers.length)
                            price = 0.5 * prices['прямые'][stage.stage]

                    if (stage.stage === 'Нанесение изолятора')
                        if (publishers.filter(({ description }) => (description?.toLowerCase().match(Type[2]) || description?.toLowerCase().match(Type[1])) && !description?.toLowerCase().match(Type[0])).length === publishers.length)
                            if (publishers.filter(({ description }) => sFs.find(sF => description.toLowerCase().includes(sF))).length === publishers.length)
                                price = 0.5 * prices['фрезированные'][stage.stage]
                            else if (publishers.filter(({ description }) => !sFs.find(sF => description.toLowerCase().includes(sF))).length === publishers.length)
                                price = 0.5 * prices['фрезированные'][stage.stage]

                    if (stage.stage === 'Покраска' && !publishers.find(({ description }) => description.toLowerCase().match(/профиль|метал/) || description.toLowerCase().match(/карниз|цоколь/) || description.toLowerCase().match(/пластик/) || description.toLowerCase().match(/карниз/)))
                        if (publishers.filter(({ sides }) => sides === 2).length === publishers.length || publishers.filter(({ sides }) => sides === 1.5).length === publishers.length || publishers.filter(({ sides }) => sides === 1).length === publishers.length)
                            if (publishers.filter(({ colourType }) => colourType && typeof colourType === 'string' && colourType.toLowerCase().includes('лак2')).length === publishers.length)
                                if (!stage.index)
                                    price = 0.5 * (prices['прямые'][stage.stage] + prices['лак']['односторонний'])
                                else
                                    price = 0.5 * (prices['обратки'][stage.stage] + prices['лак']['двусторонний'])
                            else if (publishers.filter(({ colourType }) => colourType && typeof colourType === 'string' && colourType.toLowerCase().includes('лак') && !colourType.toLowerCase().includes('лак2')).length === publishers.length)
                                if (!stage.index)
                                    price = 0.5 * (prices['прямые'][stage.stage] + prices['лак']['односторонний'])
                                else
                                    price = 0.5 * prices['обратки'][stage.stage]
                            else if (publishers.filter(({ colourType }) => colourType && typeof colourType === 'string' && !colourType.toLowerCase().includes('лак')).length === publishers.length)
                                if (!stage.index)
                                    price = 0.5 * prices['прямые'][stage.stage]
                                else
                                    price = 0.5 * prices['обратки'][stage.stage]

                    if (stage.stage === 'Полировка')
                        if (publishers.filter(({ colourType }) => colourType && typeof colourType === 'string' && colourType.toLowerCase().includes('глянец2')).length === publishers.length)
                            price = 0.5 * prices['глянец']['двусторонний']
                        else if (publishers.filter(({ colourType }) => colourType && typeof colourType === 'string' && colourType.toLowerCase().includes('глянец') && !colourType.toLowerCase().includes('глянец2')).length === publishers.length)
                            price = 0.5 * prices['глянец']['односторонний']

                    if (stage.stage === 'Подготовка к грунту' && stage.H || stage.stage != 'Подготовка к грунту' && price != publishers.map(({ stages }) => stages.find(item => item.stage === stage.stage && item.index === stage.index)).filter(stage => stage).map(({ price }) => price).reduce((a, b) => a + b, 0))
                        price += publishers.filter(({ H }) => H).map(({ amount }) => amount).reduce((a, b) => a + b, 0) * (prices['ручки'][stage.stage] || 0)
                } else if (value <= 1) {
                    if (stage.stage.includes('Подготовка') || stage.stage === 'Нанесение грунта')
                        if (publishers.filter(({ description, category }) => (category.toLowerCase().includes('f') || description?.toLowerCase().match(Type[2]) || description?.toLowerCase().match(Type[1])) && !description?.toLowerCase().match(Type[0])).length === publishers.length) {
                            if (publishers.filter(({ description, category }) => category.toLowerCase().includes('f') || sFs.find(sF => description.toLowerCase().includes(sF))).length === publishers.length)
                                price = prices['фрезированные'][stage.stage]['сложная']
                            else if (publishers.filter(({ description, category }) => !category.toLowerCase().includes('f') && !sFs.find(sF => description.toLowerCase().includes(sF))).length === publishers.length)
                                price = prices['фрезированные'][stage.stage]['простая']
                        } else if (publishers.filter(({ description }) => (description?.toLowerCase().match(Type[2]) || description?.toLowerCase().match(Type[1])) && description?.toLowerCase().match(Type[0])).length === publishers.length)
                            price = (prices['прямые'][stage.stage] + 50)
                        else if (publishers.filter(({ description }) => !(description?.toLowerCase().match(Type[2]) || description?.toLowerCase().match(Type[1])) && description?.toLowerCase().match(Type[0])).length === publishers.length)
                            price = prices['прямые'][stage.stage]

                    if (stage.stage === 'Нанесение изолятора')
                        if (publishers.filter(({ description }) => (description?.toLowerCase().match(Type[2]) || description?.toLowerCase().match(Type[1])) && !description?.toLowerCase().match(Type[0])).length === publishers.length)
                            if (publishers.filter(({ description }) => sFs.find(sF => description.toLowerCase().includes(sF))).length === publishers.length)
                                price = prices['фрезированные'][stage.stage]
                            else if (publishers.filter(({ description }) => !sFs.find(sF => description.toLowerCase().includes(sF))).length === publishers.length)
                                price = prices['фрезированные'][stage.stage]

                    if (stage.stage === 'Покраска' && !publishers.find(({ description }) => description.toLowerCase().match(/профиль|метал/) || description.toLowerCase().match(/карниз|цоколь/) || description.toLowerCase().match(/пластик/) || description.toLowerCase().match(/карниз/)))
                        if (publishers.filter(({ sides }) => sides === 2).length === publishers.length || publishers.filter(({ sides }) => sides === 1.5).length === publishers.length || publishers.filter(({ sides }) => sides === 1).length === publishers.length)
                            if (publishers.filter(({ colourType }) => colourType && typeof colourType === 'string' && colourType.toLowerCase().includes('лак2')).length === publishers.length)
                                if (!stage.index)
                                    price = (prices['прямые'][stage.stage] + prices['лак']['односторонний'])
                                else
                                    price = (prices['обратки'][stage.stage] + prices['лак']['двусторонний'])
                            else if (publishers.filter(({ colourType }) => colourType && typeof colourType === 'string' && colourType.toLowerCase().includes('лак') && !colourType.toLowerCase().includes('лак2')).length === publishers.length)
                                if (!stage.index)
                                    price = (prices['прямые'][stage.stage] + prices['лак']['односторонний'])
                                else
                                    price = prices['обратки'][stage.stage]
                            else if (publishers.filter(({ colourType }) => colourType && typeof colourType === 'string' && !colourType.toLowerCase().includes('лак')).length === publishers.length)
                                if (!stage.index)
                                    price = prices['прямые'][stage.stage]
                                else
                                    price = prices['обратки'][stage.stage]

                    if (stage.stage === 'Полировка')
                        if (publishers.filter(({ colourType }) => colourType && typeof colourType === 'string' && colourType.toLowerCase().includes('глянец2')).length === publishers.length)
                            price = prices['глянец']['двусторонний']
                        else if (publishers.filter(({ colourType }) => colourType && typeof colourType === 'string' && colourType.toLowerCase().includes('глянец') && !colourType.toLowerCase().includes('глянец2')).length === publishers.length)
                            price = prices['глянец']['односторонний']

                    if (stage.stage === 'Подготовка к грунту' && stage.H || stage.stage != 'Подготовка к грунту' && price != publishers.map(({ stages }) => stages.find(item => item.stage === stage.stage && item.index === stage.index)).filter(stage => stage).map(({ price }) => price).reduce((a, b) => a + b, 0))
                        price += publishers.filter(({ H }) => H).map(({ amount }) => amount).reduce((a, b) => a + b, 0) * (prices['ручки'][stage.stage] || 0)
                }

            let addition = []
            if (publishers.find(({ stages }) => stages.find(({ stage }) => stage === 'Подготовка к изолятору')) && stage.stage.match(/Подготовка|Распил/))
                addition.push(`Ф: ${round(publishers.filter(({ stages }) => stages.find(({ stage }) => stage === 'Подготовка к изолятору')).map(({ square }) => square).reduce((a, b) => a + b, 0), 2)}`)
            if (publishers.find(({ colourType, stages }) => typeof colourType === 'string' && colourType?.toLowerCase().match(/лак/) && (stage.stage === 'Полировка' ? stages.find(({ stage }) => stage === 'Полировка') : true)) && stage.stage.match(/Покраска|Полировка/))
                addition.push(`Л: ${round(publishers.filter(({ colourType, stages }) => typeof colourType === 'string' && colourType?.toLowerCase().match(/лак/) && (stage.stage === 'Полировка' ? stages.find(({ stage }) => stage === 'Полировка') : true)).map(({ square }) => square).reduce((a, b) => a + b, 0), 2)}`)
            if (publishers.find(({ colourType, stages }) => typeof colourType === 'string' && colourType?.toLowerCase().match(/лак2/) && (stage.stage === 'Полировка' ? stages.find(({ stage }) => stage === 'Полировка') : true)) && stage.stage.match(/Покраска|Полировка/))
                addition.push(`Л2: ${round(publishers.filter(({ colourType, stages }) => typeof colourType === 'string' && colourType?.toLowerCase().match(/лак2/) && (stage.stage === 'Полировка' ? stages.find(({ stage }) => stage === 'Полировка') : true)).map(({ square }) => square).reduce((a, b) => a + b, 0), 2)}`)
            addition = addition.join('\r\n')

            let factvalue = 0
            for (let publisher of publishers.filter(({ stages }) => stages.find(item => item.stage === stage.stage && item.index === stage.index)))
                if (stage.stage === 'Покраска')
                    if (stages.find(({ stage, factvalue }) => stage === 'Покраска' && factvalue))
                        factvalue += publisher.square * (typeof publisher.colourType === 'string' && publisher.colourType?.toLowerCase().match(/лак2/) ? 2 : 1)
                    else
                        factvalue += publisher.square * (typeof publisher.colourType === 'string' && publisher.colourType?.toLowerCase().match(/лак/) ? 2 : 1)
                else if (stage.stage.match('Подготовка'))
                    if (stage.stage === 'Подготовка к грунту' || stage.stage === 'Подготовка к покраске')
                        if (publisher.stages.find(({ stage, index }) => (stage === 'Подготовка к грунту' || stage.stage === 'Подготовка к покраске') && index === 2)) {
                            if (stage.index === 2)
                                factvalue += publisher.square * 20 / 7
                        } else
                            factvalue += publisher.square * (publisher.stages.find(({ stage }) => stage === 'Подготовка к изолятору') ? 20 / 7 : 1)
                    else
                        factvalue += publisher.square * (publisher.stages.find(({ stage }) => stage === 'Подготовка к изолятору') ? 20 / 7 : 1)
                else
                    factvalue += publisher.square
            factvalue = round(factvalue, 2)

            return {
                stage: stage.stage,
                term: Math.max(...publishers.map(({ stages }) => stages.find(item => item.stage === stage.stage)?.term || 0)),
                value,
                factvalue,
                index: stage.index,
                price: round(price, 0),
                weekend: stage.weekend,
                addition,
                H: stage.H && publishers.find(publisher => publisher.stages.find(item => item.stage === stage.stage && item.H && item.index === stage.index)) ? true : false
            }
        })
        if (stages.find(({ stage, H }) => stage === 'Подготовка к грунту' && H))
            stages.find(({ stage, H }) => stage === 'Подготовка к грунту' && H).term = 1
        if (stages.find(({ stage, index }) => stage === 'Подготовка к грунту' && index === 2)) {
            let value = stages.find(({ stage, index }) => stage === 'Подготовка к грунту' && !index).factvalue
            stages.find(({ stage, index }) => stage === 'Подготовка к грунту' && index === 2).term = 0
            if (value === 0) {
                let stage = stages.find(({ stage, index }) => stage === 'Подготовка к грунту' && !index)
                stages.splice(stages.findIndex(({ stage, index }) => stage === 'Подготовка к грунту' && !index), 1)
                stages.find(({ stage, index }) => stage === 'Подготовка к грунту' && index === 2).term = stage.term
                stages.find(({ stage, index }) => stage === 'Подготовка к грунту' && index === 2).index = stage.index
            } else {
                stages.find(({ stage, index }) => stage === 'Подготовка к грунту' && !index).value = value
                stages.find(({ stage, index }) => stage === 'Подготовка к грунту' && !index).addition = stages.find(({ stage, index }) => stage === 'Подготовка к грунту' && !index).addition.replace(/Ф: \d+\.?\d{0,}/, '')
            }
        }
        if (stages.find(({ stage, index }) => stage === 'Подготовка к покраске' && index === 2)) {
            let value = stages.find(({ stage, index }) => stage === 'Подготовка к покраске' && !index).factvalue
            stages.find(({ stage, index }) => stage === 'Подготовка к покраске' && index === 2).term = 0
            if (value === 0) {
                let stage = stages.find(({ stage, index }) => stage === 'Подготовка к покраске' && !index)
                stages.splice(stages.findIndex(({ stage, index }) => stage === 'Подготовка к покраске' && !index), 1)
                stages.find(({ stage, index }) => stage === 'Подготовка к покраске' && index === 2).term = stage.term
                stages.find(({ stage, index }) => stage === 'Подготовка к покраске' && index === 2).index = stage.index
            } else {
                stages.find(({ stage, index }) => stage === 'Подготовка к покраске' && !index).value = value
                stages.find(({ stage, index }) => stage === 'Подготовка к покраске' && !index).addition = stages.find(({ stage, index }) => stage === 'Подготовка к покраске' && !index).addition.replace(/Ф: \d+\.?\d{0,}/, '')
            }
        }

        if (publishers.find(({ T }) => T) && stages.find(stage => stage.stage === 'Склейка')) {
            stages.splice(stages.findIndex(stage => stage.stage === 'Склейка'), 0,
                {
                    stage: 'Подготовка к грунту',
                    term: Math.ceil(publishers.filter(({ T }) => T).map(({ square }) => square).reduce((a, b) => a + b, 0) / 20),
                    value: round(publishers.filter(({ T }) => T).map(({ square }) => square).reduce((a, b) => a + b, 0), 2),
                    factvalue: round(publishers.filter(({ T }) => T).map(({ square }) => square).reduce((a, b) => a + b, 0), 2),
                    // price: round(publishers.filter(({ T }) => T).map(({ square }) => square).reduce((a, b) => a + b, 0) * prices['обратки']['Подготовка к грунту'] * 1.5, 2),
                    weekend: false,
                    addition: `Т: ${round(publishers.filter(({ T }) => T).map(({ square }) => square).reduce((a, b) => a + b, 0), 2)}`,
                },
                {
                    stage: 'Нанесение грунта',
                    term: Math.ceil(publishers.filter(({ T }) => T).map(({ square }) => square).reduce((a, b) => a + b, 0) / 20),
                    value: round(publishers.filter(({ T }) => T).map(({ square }) => square).reduce((a, b) => a + b, 0), 2),
                    factvalue: round(publishers.filter(({ T }) => T).map(({ square }) => square).reduce((a, b) => a + b, 0), 2),
                    // price: round(publishers.filter(({ T }) => T).map(({ square }) => square).reduce((a, b) => a + b, 0) * prices['обратки']['Нанесение грунта'] * 1.5, 2),
                    weekend: false,
                    addition: `Т: ${round(publishers.filter(({ T }) => T).map(({ square }) => square).reduce((a, b) => a + b, 0), 2)}`,
                },
                {
                    stage: 'Подготовка к покраске',
                    term: Math.ceil(publishers.filter(({ T }) => T).map(({ square }) => square).reduce((a, b) => a + b, 0) / 20),
                    value: round(publishers.filter(({ T }) => T).map(({ square }) => square).reduce((a, b) => a + b, 0), 2),
                    factvalue: round(publishers.filter(({ T }) => T).map(({ square }) => square).reduce((a, b) => a + b, 0), 2),
                    price: round(publishers.filter(({ T }) => T).map(({ square }) => square).reduce((a, b) => a + b, 0) * prices['обратки']['Подготовка к покраске'], 0),
                    weekend: false,
                    addition: `Т: ${round(publishers.filter(({ T }) => T).map(({ square }) => square).reduce((a, b) => a + b, 0), 2)}`,
                }
            )
            if (publishers.find(({ T, stages }) => T && stages.find(({ stage }) => stage === 'Покраска')))
                stages.splice(stages.findIndex(stage => stage.stage === 'Склейка'), 0,
                    {
                        stage: 'Покраска',
                        term: Math.ceil(publishers.filter(({ T, stages }) => T && stages.find(({ stage }) => stage === 'Покраска')).map(({ colourType, square }) => square * (typeof colourType === 'string' && colourType?.toLowerCase().match(/лак2/) ? 2 : 1)).reduce((a, b) => a + b, 0) / 10),
                        value: round(publishers.filter(({ T, stages }) => T && stages.find(({ stage }) => stage === 'Покраска')).map(({ square }) => square).reduce((a, b) => a + b, 0), 2),
                        factvalue: round(publishers.filter(({ T, stages }) => T && stages.find(({ stage }) => stage === 'Покраска')).map(({ colourType, square }) => square * (typeof colourType === 'string' && colourType?.toLowerCase().match(/лак2/) ? 2 : 1)).reduce((a, b) => a + b, 0), 2),
                        price: round(publishers.filter(({ T, stages }) => T && stages.find(({ stage }) => stage === 'Покраска')).map(({ square }) => square).reduce((a, b) => a + b, 0) * prices['обратки']['Покраска'], 0),
                        weekend: false,
                        addition: `Т: ${round(publishers.filter(({ T, stages }) => T && stages.find(({ stage }) => stage === 'Покраска')).map(({ square }) => square).reduce((a, b) => a + b, 0), 2)}`,
                    })
        }

        let last = created
        while ([0, 6].includes(new Date(last).getDay()))
            last += Day

        for (let i = 0; i < stages.length; i++) {
            if (stages[i].H || stages[i].A) {
                stages[i].delay = 0
                let add = last
                while ([0, 6].includes(new Date(add).getDay()))
                    add += Day
                stages.splice(i, 0, {
                    days: `${add}`,
                    term: 1,
                    weekend: false,
                    stage: 'Фрезеровка ручек',
                    value: stages[i].value,
                    factvalue: stages[i].factvalue,
                    addition: ''
                })
                delete stages[i + 1].H
                delete stages[i + 1].A

                continue
            }
            if (stages[i].stage === 'Аппликация') {
                let days = []
                let appl = last + Day
                if (!stages[i].weekend)
                    while ([0, 6].includes(new Date(appl).getDay()))
                        appl += Day

                days.push(`${appl}`)
                stages[i].days = days.join(',')

                continue
            }
            if (stages[i].stage === 'Покраска' && i - 1 >= 0 && stages[i - 1].stage === 'Аппликация')
                stages[i].delay = 0
            if (stages[i].stage === 'Отгрузка') {
                stages[i].days = `${lastDate}`
                stages[i].shiftable = false

                continue
            }
            if (stages[i].stage === 'Оплата') {
                stages[i].delay = 0
                stages[i].days = `${stages.find(({ stage }) => stage === 'Упаковка').days.split(',').reverse()[0]}`
                continue
            }
            if (stages[i].stage === 'Полировка') {
                stages[i].delay = 2
            }
            if ((stages[i].stage === 'Подготовка к грунту' || stages[i].stage === 'Подготовка к покраске') && stages[i].index === 2) {
                stages[i].delay = 0
                let days = []
                let appl = last
                if (!stages[i].weekend)
                    while ([0, 6].includes(new Date(appl).getDay()))
                        appl += Day

                days.push(`${appl}`)
                stages[i].days = days.join(',')

                continue
            }

            let days = []
            let term = stages[i].term
            if (!term)
                days.push(`${last}`)
            if (stages[i].stage === 'Полировка')
                last += Day
            while (term) {
                last += Day
                if (!stages[i].weekend)
                    while ([0, 6].includes(new Date(last).getDay()))
                        last += Day

                days.push(`${last}`)
                term--
            }
            stages[i].days = days.join(',')
            if (stages[i].stage === 'Распил') {
                stages[i].days = stages[i].days.split(',').slice(-1).join(',')
                stages[i].term = 1
            }

            continue
        }

        let order = {
            number: JSON.parse(ws['N7']?.v ? ws['N7']?.v?.match(/\d+/)[0] : ws['M7']?.v?.match(/\d+/)[0]),
            created: /* +new Date(+new Date(2024, 0, 1) + 4 * Hour),// */created,
            lastDate: /* +new Date(+new Date(2024, 0, 31) + 4 * Hour),// */lastDate,
            stages,
            snapshot: typeof XLSX.utils.sheet_to_html(ws) === 'string' && XLSX.utils.sheet_to_html(ws).match(/<table>.+<\/table>/) && XLSX.utils.sheet_to_html(ws).match(/<table>.+<\/table>/)[0].length <= 32768 ? XLSX.utils.sheet_to_html(ws).match(/<table>.+<\/table>/)[0] : null,
            costumer: 'Null'//workbook.Sheets['СЧЕТ']['P6']?.v || workbook.Sheets['СЧЕТ']['O6']?.v,
        }
        console.log('Файл', file, 'прочитан успешно')

        return resolve(order)
    })
}

module.exports = newOrder
