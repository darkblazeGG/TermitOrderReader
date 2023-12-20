const XLSX = require('xlsx')

const Сarpentry = [/р/, /п/, /с/, /б/]
const Type = [/прямой/, /фрез/, /f/, /фас кл/, /фасад клиента/]

const Second = 1000
const Minute = 60 * Second
const Hour = 60 * Minute
const Day = 24 * Hour

const prices = require('./prices.json')

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
            console.log('reading file', file)
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
            let square = rows.filter(row => row[rows[0].findIndex(row => row === 'Примечание')]?.toLowerCase().match(Type[2]) || row[rows[0].findIndex(row => row === 'Примечание')]?.toLowerCase().match(Type[1]))
                .map(row => row[rows[0].findIndex(row => row === 'Площадь')]).reduce((a, b) => a + b, 0) * 20 / 7 + rows.filter(row => !row[rows[0].findIndex(row => row === 'Примечание')].toLowerCase().match(Type[2]) && !row[rows[0].findIndex(row => row === 'Примечание')].toLowerCase().match(Type[1]))
                    .map(row => row[rows[0].findIndex(row => row === 'Площадь')]).splice(1,).reduce((a, b) => a + b, 0)
            let stages = [
                {
                    term: 3,
                    stage: 'Распил'
                },
                {
                    term: Math.ceil(square / 20),
                    stage: 'Шлифовка к грунту'
                },
                {
                    term: 1,
                    stage: 'Нанесение грунта'
                },
                {
                    term: Math.ceil(square / 20),
                    stage: 'Шлифовка к покраске'
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

            if (row[rows[0].findIndex(row => row === 'МДФ')] === '-' || row[rows[0].findIndex(row => row === 'Примечание')]?.toLowerCase().match(Type[3]) || row[rows[0].findIndex(row => row === 'Примечание')]?.toLowerCase().match(Type[4]))
                stages.splice(stages.findIndex(stage => stage.stage === 'Распил'), 1)
            else if (row[rows[0].findIndex(row => row === 'Примечание')]?.toLowerCase().match(Type[2]) || row[rows[0].findIndex(row => row === 'Примечание')]?.toLowerCase().match(Type[1])) {
                let square = rows.filter(row => row[rows[0].findIndex(row => row === 'Примечание')]?.toLowerCase().match(Type[2]) || row[rows[0].findIndex(row => row === 'Примечание')]?.toLowerCase().match(Type[1]))
                    .map(row => row[rows[0].findIndex(row => row === 'Площадь')]).reduce((a, b) => a + b, 0)
                stages[stages.findIndex(stage => stage.stage === 'Распил')].term += 2
                stages.splice(stages.findIndex(stage => stage.stage === 'Распил') + 1, 0,
                    {
                        term: Math.ceil(square / 7),
                        stage: 'Шлифовка к изолятору'
                    },
                    {
                        term: 1,
                        stage: 'Нанесение изолятора'
                    }
                )
            }
            if (row[rows[0].findIndex(row => row === 'Примечание')]?.toLowerCase().match(/лдсп/)) {
                stages.splice(stages.findIndex(stage => stage.stage === 'Шлифовка к грунту'), 1)
                stages.splice(stages.findIndex(stage => stage.stage === 'Нанесение грунта'), 1)
                stages.splice(stages.findIndex(stage => stage.stage === 'Покраска'), 0, {
                    term: 1,
                    stage: 'Аппликация'
                })
            }

            let P = row[rows[0].findIndex(row => row === 'Категория')]?.toString().toLowerCase().match(Сarpentry[1]) ? 1 : 0
            let H = row[rows[0].findIndex(row => row === 'Категория')]?.toString().toLowerCase().match(Сarpentry[0]) ? 1 : 0
            let G = row[rows[0].findIndex(row => row === 'Категория')]?.toString().toLowerCase().match(Сarpentry[2]) ? 1 : 0
            let Paint = row[rows[0].findIndex(row => row === 'Категория')]?.toString().toLowerCase().match(Сarpentry[3]) ? 1 : 0
            if (Paint)
                stages.splice(stages.findIndex(stage => stage.stage === 'Покраска'), 1)

            if (H) {
                if (row[rows[0].findIndex(row => row === 'Примечание')]?.toLowerCase().match(Type[2]) || row[rows[0].findIndex(row => row === 'Примечание')]?.toLowerCase().match(Type[1])) {
                    stages[stages.findIndex(stage => stage.stage === 'Шлифовка к изолятору')].H = true
                    stages[stages.findIndex(stage => stage.stage === 'Шлифовка к изолятору')].term++
                } else {
                    stages[stages.findIndex(stage => stage.stage === 'Шлифовка к грунту')].H = true
                    stages[stages.findIndex(stage => stage.stage === 'Шлифовка к грунту')].term++
                }
            }

            if (P || G)
                stages.splice(stages.findIndex(stage => stage.stage === 'Распил') + 1, 0, {
                    term: P + G,
                    stage: 'Столярка',
                    P,
                    G
                })

            if (row[rows[0].findIndex(row => row === 'Цвет')] === '-')
                stages.splice(stages.findIndex(stage => stage.stage === 'Покраска'), 1)
            else if (row[rows[0].findIndex(row => row === 'Тип краски')]?.toString().toLowerCase().includes('глянец') && !Paint) {
                let squere = rows.filter(row => row[rows[0].findIndex(row => row === 'Тип краски')].toLowerCase().includes('глянец'))
                    .map(row => row[rows[0].findIndex(row => row === 'Площадь')]).reduce((a, b) => a + b, 0)
                stages[stages.findIndex(stage => stage.stage === 'Покраска')]
                stages.splice(stages.findIndex(stage => stage.stage === 'Покраска') + 1, 0,
                    {
                        term: 3,
                        stage: 'Полировка'
                    }
                )
                if (row[rows[0].findIndex(row => row === 'Кол-во сторон')] === 2 && !row[rows[0].findIndex(row => row === 'Тип краски')].toLowerCase().match(/глянец2/))
                    stages[stages.findIndex(stage => stage.stage === 'Нанесение грунта')].term++
                if (squere > 10 || row[rows[0].findIndex(row => row === 'Кол-во сторон')] === 2 || row[rows[0].findIndex(row => row === 'Кол-во сторон')] === 'все' || row[rows[0].findIndex(row => row === 'Тип краски')].toLowerCase().match(/лак/))
                    stages[stages.findIndex(stage => stage.stage === 'Полировка')].term += 2
                if (row[rows[0].findIndex(row => row === 'Кол-во сторон')] === 'все')
                    row[rows[0].findIndex(row => row === 'Кол-во сторон')] = 2
                let paintstages = row[rows[0].findIndex(row => row === 'Кол-во сторон')] - 1 + ((row[rows[0].findIndex(row => row === 'Цвет')].toString() || undefined)?.toLowerCase().split(',').length || 1) - 1
                for (let i = 0; i < paintstages; i++)
                    stages.splice(stages.findIndex(stage => stage.stage === 'Покраска') + 1, 0, {
                        term: 1,
                        stage: 'Аппликация',
                        index: paintstages - i - 1
                    }, {
                        term: 1,
                        stage: 'Покраска',
                        index: paintstages - i
                    })
                if (squere > 10 && stages.filter(({ stage }) => stage === 'Покраска').length === 1)
                    stages[stages.findIndex(stage => stage.stage === 'Покраска')].term++
            } else if (!Paint) {
                let squere = rows.filter(row => row[rows[0].findIndex(row => row === 'Тип краски')]?.toString().toLowerCase().includes('мат'))
                    .map(row => row[rows[0].findIndex(row => row === 'Площадь')]).reduce((a, b) => a + b, 0)
                let paintstages = row[rows[0].findIndex(row => row === 'Кол-во сторон')] - 1 + ((row[rows[0].findIndex(row => row === 'Цвет')].toString() || undefined)?.toLowerCase().split(',').length || 1) - 1
                for (let i = 0; i < paintstages; i++)
                    stages.splice(stages.findIndex(stage => stage.stage === 'Покраска') + 1, 0, {
                        term: 1,
                        stage: 'Аппликация',
                        index: paintstages - i - 1
                    }, {
                        term: 1,
                        stage: 'Покраска',
                        index: paintstages - i
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
                colour: row[rows[0].findIndex(row => row === 'Цвет')],
                colourType: row[rows[0].findIndex(row => row === 'Тип краски')],
                sides: typeof row[rows[0].findIndex(row => row === 'Кол-во сторон')] === 'number' ? round(row[rows[0].findIndex(row => row === 'Кол-во сторон')], 0) : 2,
                radius: row[rows[0].findIndex(row => row === 'Радиус')] === 'мин' ? 1 : typeof row[rows[0].findIndex(row => row === 'МДФ')] === 'number' ? row[rows[0].findIndex(row => row === 'МДФ')] : 0,
                stages
            }
        })
        if (publishers.includes(undefined))
            return
        if (publishers.find(({ stages }) => stages.find(({ P }) => P)) && publishers.find(({ stages }) => stages.find(({ G }) => G)))
            publishers = publishers.map(publisher => {
                publisher.stages = publisher.stages.map(stage => {
                    if (stage.P && !stage.G || stage.G && !stage.P)
                        stage.term += 1
                    return stage
                })
                return publisher
            })

        if (!ws['N8'])
            return resolve('Нет даты поступления в работу, пожалуйста добавьте ее прежде чем загружать заказ в систему')
        if (!workbook.Sheets['СЧЕТ']['O' + JSON.parse(Object.keys(workbook.Sheets['СЧЕТ']).find(key => workbook.Sheets['СЧЕТ'][key]?.v === 'Дата готовности заказа')?.match(/\d+/)?.[0])])
            return resolve('Нет даты окончания работы, пожалуйста добавьте ее прежде чем загружать заказ в систему')

        let created = ws['N8']?.v ? ExcelDateToJSDate(ws['N8']?.v) : undefined
        let lastDate = workbook.Sheets['СЧЕТ']['O' + JSON.parse(Object.keys(workbook.Sheets['СЧЕТ']).find(key => workbook.Sheets['СЧЕТ'][key].v === 'Дата готовности заказа').match(/\d+/)[0])]?.v ? ExcelDateToJSDate(workbook.Sheets['СЧЕТ']['O' + JSON.parse(Object.keys(workbook.Sheets['СЧЕТ']).find(key => workbook.Sheets['СЧЕТ'][key].v === 'Дата готовности заказа').match(/\d+/)[0])]?.v) : undefined

        let stages = [
            {
                term: 0,
                stage: 'Распил'
            },
            {
                term: 0,
                stage: 'Столярка'
            },
            {
                term: 0,
                stage: 'Шлифовка к изолятору'
            },
            {
                term: 0,
                stage: 'Нанесение изолятора'
            },
            {
                term: 0,
                stage: 'Шлифовка к грунту'
            },
            {
                term: 0,
                stage: 'Нанесение грунта'
            },
            {
                term: 0,
                stage: 'Шлифовка к покраске'
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
        if (publishers.find(({ stages }) => stages.find(({ stage, index }) => stage === 'Аппликация' && index === undefined)))
            stages.splice(stages.findIndex(stage => stage.stage === 'Покраска'), 0, {
                term: 1,
                stage: 'Аппликация',
                weekend: true
            })

        let paintstages = Math.max(...publishers.map(({ stages }) => Math.max(...stages.map(({ index }) => index || 0))))
        for (let i = 0; i < paintstages; i++)
            stages.splice(stages.findIndex(stage => stage.stage === 'Покраска') + 1, 0, {
                term: 1,
                stage: 'Аппликация',
                index: paintstages - i - 1,
                weekend: true
            }, {
                term: 1,
                stage: 'Покраска',
                index: paintstages - i,
                weekend: true
            })

        stages = stages.filter(stage => publishers.find(({ stages }) => stages.find(item => item.stage === stage.stage)))
        stages = stages.map(stage => {
            // let price = 0
            // if (stage.stage === 'Шлифовка к изолятору') {
            //     if (publishers.map(({ square }) => square).reduce((a, b) => a + b, 0) <= 0.5) {
            //         if (publishers.filter(({ stages }) => stages.find(item => item.stage === 'Нанесение изолятора')).length === publishers.length)
            //             price += prices['фрезированные']['Шлифовка к изолятору'] * 0.5
            //         else
            //             price += prices['фрезированные']['Шлифовка к изолятору'] * publishers.filter(({ stages }) => stages.find(item => item.stage === 'Нанесение изолятора')).map(({ square }) => square).reduce((a, b) => a + b, 0)
            //     } else if (publishers.map(({ square }) => square).reduce((a, b) => a + b, 0) < 1) {
            //         if (publishers.filter(({ stages }) => stages.find(item => item.stage === 'Нанесение изолятора')).length === publishers.length)
            //             price += prices['фрезированные']['Шлифовка к изолятору']
            //         else
            //             price += prices['фрезированные']['Шлифовка к изолятору'] * publishers.filter(({ stages }) => stages.find(item => item.stage === 'Нанесение изолятора')).map(({ square }) => square).reduce((a, b) => a + b, 0)
            //     } else
            //         price += prices['фрезированные']['Шлифовка к изолятору'] * publishers.filter(({ stages }) => stages.find(item => item.stage === 'Нанесение изолятора')).map(({ square }) => square).reduce((a, b) => a + b, 0)

            // } else if (stage.stage === 'Нанесение изолятора') {
            //     if (publishers.map(({ square }) => square).reduce((a, b) => a + b, 0) <= 0.5) {
            //         if (publishers.filter(({ stages }) => stages.find(item => item.stage === 'Нанесение изолятора')).length === publishers.length)
            //             price += prices['фрезированные']['Нанесение изолятора'] * 0.5
            //         else
            //             price += prices['фрезированные']['Нанесение изолятора'] * publishers.filter(({ stages }) => stages.find(item => item.stage === 'Нанесение изолятора')).map(({ square }) => square).reduce((a, b) => a + b, 0)
            //     } else if (publishers.map(({ square }) => square).reduce((a, b) => a + b, 0) < 1) {
            //         if (publishers.filter(({ stages }) => stages.find(item => item.stage === 'Нанесение изолятора')).length === publishers.length)
            //             price += prices['фрезированные']['Нанесение изолятора']
            //         else
            //             price += prices['фрезированные']['Нанесение изолятора'] * publishers.filter(({ stages }) => stages.find(item => item.stage === 'Нанесение изолятора')).map(({ square }) => square).reduce((a, b) => a + b, 0)
            //     } else
            //         price += prices['фрезированные']['Нанесение изолятора'] * publishers.filter(({ stages }) => stages.find(item => item.stage === 'Нанесение изолятора')).map(({ square }) => square).reduce((a, b) => a + b, 0)
            // } else if (stage.stage === 'Шлифовка к грунту') {
            //     if (publishers.filter(({ stages }) => stages.find(item => item.stage === 'Нанесение изолятора')).length === publishers.length) {
            //         if (publishers.map(({ square }) => square).reduce((a, b) => a + b, 0) <= 0.5)
            //             price += prices['фрезированные']['Шлифовка к грунту'] * 0.5 + prices['обратки']['Шлифовка к грунту'] * 0.5
            //         else if (publishers.map(({ square }) => square).reduce((a, b) => a + b, 0) < 1)
            //             price += prices['фрезированные']['Шлифовка к грунту'] + prices['обратки']['Шлифовка к грунту']
            //         else {
            //             price += prices['обратки']['Шлифовка к грунту'] * publishers.filter(({ stages, sides, colourType }) => typeof colourType === 'string' && colourType?.toLowerCase().match(/глянец2/) && sides === 2 && stages.find(item => item.stage === 'Полировка')).map(({ square }) => square).reduce((a, b) => a + b, 0)
            //             price += prices['прямые']['Шлифовка к грунту'] * publishers.filter(({ stages }) => !stages.find(item => item.stage === 'Нанесение изолятора')).map(({ square }) => square).reduce((a, b) => a + b, 0)
            //             price += prices['фрезированные']['Шлифовка к грунту'] * publishers.filter(({ stages }) => stages.find(item => item.stage === 'Нанесение изолятора')).map(({ square }) => square).reduce((a, b) => a + b, 0)
            //         }
            //     } else if (publishers.filter(({ stages }) => !stages.find(item => item.stage === 'Нанесение изолятора')).length === publishers.length) {
            //         if (publishers.map(({ square }) => square).reduce((a, b) => a + b, 0) <= 0.5)
            //             price += prices['прямые']['Шлифовка к грунту'] * 0.5 + prices['обратки']['Шлифовка к грунту'] * 0.5
            //         else if (publishers.map(({ square }) => square).reduce((a, b) => a + b, 0) < 1)
            //             price += prices['прямые']['Шлифовка к грунту'] + prices['обратки']['Шлифовка к грунту']
            //         else {
            //             price += prices['обратки']['Шлифовка к грунту'] * publishers.filter(({ stages, sides, colourType }) => typeof colourType === 'string' && colourType?.toLowerCase().match(/глянец2/) && sides === 2 && stages.find(item => item.stage === 'Полировка')).map(({ square }) => square).reduce((a, b) => a + b, 0)
            //             price += prices['прямые']['Шлифовка к грунту'] * publishers.filter(({ stages }) => !stages.find(item => item.stage === 'Нанесение изолятора')).map(({ square }) => square).reduce((a, b) => a + b, 0)
            //             price += prices['фрезированные']['Шлифовка к грунту'] * publishers.filter(({ stages }) => stages.find(item => item.stage === 'Нанесение изолятора')).map(({ square }) => square).reduce((a, b) => a + b, 0)
            //         }
            //     } else {
            //         price += prices['обратки']['Шлифовка к грунту'] * publishers.filter(({ stages, sides, colourType }) => typeof colourType === 'string' && colourType?.toLowerCase().match(/глянец2/) && sides === 2 && stages.find(item => item.stage === 'Полировка')).map(({ square }) => square).reduce((a, b) => a + b, 0)
            //         price += prices['прямые']['Шлифовка к грунту'] * publishers.filter(({ stages }) => !stages.find(item => item.stage === 'Нанесение изолятора')).map(({ square }) => square).reduce((a, b) => a + b, 0)
            //         price += prices['фрезированные']['Шлифовка к грунту'] * publishers.filter(({ stages }) => stages.find(item => item.stage === 'Нанесение изолятора')).map(({ square }) => square).reduce((a, b) => a + b, 0)
            //     }
            // } else if (stage.stage === 'Нанесение грунта') {
            //     if (publishers.filter(({ stages }) => stages.find(item => item.stage === 'Нанесение изолятора')).length === publishers.length && (publishers.filter(({ thickness }) => thickness === 22).length === publishers.length || publishers.filter(({ thickness }) => thickness != 22).length === publishers.length)) {
            //         if (publishers.map(({ square }) => square).reduce((a, b) => a + b, 0) <= 0.5)
            //             price += prices['фрезированные']['Нанесение грунта'] * 0.5 + prices['обратки']['Нанесение грунта'] * 0.5
            //         else if (publishers.map(({ square }) => square).reduce((a, b) => a + b, 0) < 1)
            //             price += prices['фрезированные']['Нанесение грунта'] + prices['обратки']['Нанесение грунта']
            //         else {
            //             price += prices['обратки']['Нанесение грунта'] * publishers.filter(({ stages, sides, colourType }) => typeof colourType === 'string' && colourType?.toLowerCase().match(/глянец2/) && sides && stages.find(item => item.stage === 'Полировка')).map(({ square }) => square).reduce((a, b) => a + b, 0)
            //             price += prices['прямые']['Нанесение изолятора'] * publishers.filter(({ stages, thickness }) => (!stages.find(item => item.stage === 'Нанесение изолятора') && thickness === 22 || stages.find(item => item.stage === 'Полировка'))).map(({ square }) => square).reduce((a, b) => a + b, 0)
            //             price += prices['прямые']['Нанесение грунта'] * publishers.filter(({ stages }) => !stages.find(item => item.stage === 'Нанесение изолятора')).map(({ square }) => square).reduce((a, b) => a + b, 0)
            //             price += prices['фрезированные']['Нанесение грунта'] * publishers.filter(({ stages }) => stages.find(item => item.stage === 'Нанесение изолятора')).map(({ square }) => square).reduce((a, b) => a + b, 0)
            //         }
            //     } else if (publishers.filter(({ stages }) => !stages.find(item => item.stage === 'Нанесение изолятора')).length === publishers.length && (publishers.filter(({ thickness }) => thickness === 22).length === publishers.length || publishers.filter(({ thickness }) => thickness != 22).length === publishers.length)) {
            //         if (publishers.map(({ square }) => square).reduce((a, b) => a + b, 0) <= 0.5) {
            //             price += prices['прямые']['Нанесение грунта'] * 0.5 + prices['обратки']['Нанесение грунта'] * 0.5
            //             if (publishers.filter(({ thickness }) => thickness === 22).length === publishers.length)
            //                 price += prices['прямые']['Нанесение изолятора'] * 0.5
            //         } else if (publishers.map(({ square }) => square).reduce((a, b) => a + b, 0) < 1) {
            //             price += prices['прямые']['Нанесение грунта'] + prices['обратки']['Нанесение грунта']
            //             if (publishers.filter(({ thickness }) => thickness === 22).length === publishers.length)
            //                 price += prices['прямые']['Нанесение изолятора']
            //         } else {
            //             price += prices['обратки']['Нанесение грунта'] * publishers.filter(({ stages, sides, colourType }) => typeof colourType === 'string' && colourType?.toLowerCase().match(/глянец2/) && sides && stages.find(item => item.stage === 'Полировка')).map(({ square }) => square).reduce((a, b) => a + b, 0)
            //             price += prices['прямые']['Нанесение изолятора'] * publishers.filter(({ stages, thickness }) => (!stages.find(item => item.stage === 'Нанесение изолятора') && thickness === 22 || stages.find(item => item.stage === 'Полировка'))).map(({ square }) => square).reduce((a, b) => a + b, 0)
            //             price += prices['прямые']['Нанесение грунта'] * publishers.filter(({ stages }) => !stages.find(item => item.stage === 'Нанесение изолятора')).map(({ square }) => square).reduce((a, b) => a + b, 0)
            //             price += prices['фрезированные']['Нанесение грунта'] * publishers.filter(({ stages }) => stages.find(item => item.stage === 'Нанесение изолятора')).map(({ square }) => square).reduce((a, b) => a + b, 0)
            //         }
            //     } else {
            //         price += prices['обратки']['Нанесение грунта'] * publishers.filter(({ stages, sides, colourType }) => typeof colourType === 'string' && colourType?.toLowerCase().match(/глянец2/) && sides && stages.find(item => item.stage === 'Полировка')).map(({ square }) => square).reduce((a, b) => a + b, 0)
            //         price += prices['прямые']['Нанесение изолятора'] * publishers.filter(({ stages, thickness }) => (!stages.find(item => item.stage === 'Нанесение изолятора') && thickness === 22 || stages.find(item => item.stage === 'Полировка'))).map(({ square }) => square).reduce((a, b) => a + b, 0)
            //         price += prices['прямые']['Нанесение грунта'] * publishers.filter(({ stages }) => !stages.find(item => item.stage === 'Нанесение изолятора')).map(({ square }) => square).reduce((a, b) => a + b, 0)
            //         price += prices['фрезированные']['Нанесение грунта'] * publishers.filter(({ stages }) => stages.find(item => item.stage === 'Нанесение изолятора')).map(({ square }) => square).reduce((a, b) => a + b, 0)
            //     }
            // } else if (stage.stage === 'Шлифовка к покраске') {
            //     if (publishers.filter(({ stages }) => stages.find(item => item.stage === 'Нанесение изолятора')).length === publishers.length) {
            //         if (publishers.map(({ square }) => square).reduce((a, b) => a + b, 0) <= 0.5)
            //             price += prices['фрезированные']['Шлифовка к покраске'] * 0.5 + prices['обратки']['Шлифовка к покраске'] * 0.5
            //         else if (publishers.map(({ square }) => square).reduce((a, b) => a + b, 0) < 1)
            //             price += prices['фрезированные']['Шлифовка к покраске'] + prices['обратки']['Шлифовка к покраске']
            //         else {
            //             price += prices['обратки']['Шлифовка к покраске'] * publishers.filter(({ sides }) => sides === 2).map(({ square }) => square).reduce((a, b) => a + b, 0)
            //             price += prices['прямые']['Шлифовка к покраске'] * publishers.filter(({ stages }) => !stages.find(item => item.stage === 'Нанесение изолятора')).map(({ square }) => square).reduce((a, b) => a + b, 0)
            //             price += prices['фрезированные']['Шлифовка к покраске'] * publishers.filter(({ stages }) => stages.find(item => item.stage === 'Нанесение изолятора')).map(({ square }) => square).reduce((a, b) => a + b, 0)
            //         }
            //     } else if (publishers.filter(({ stages }) => !stages.find(item => item.stage === 'Нанесение изолятора')).length === publishers.length) {
            //         if (publishers.map(({ square }) => square).reduce((a, b) => a + b, 0) <= 0.5)
            //             price += prices['прямые']['Шлифовка к покраске'] * 0.5 + prices['обратки']['Шлифовка к покраске'] * 0.5
            //         else if (publishers.map(({ square }) => square).reduce((a, b) => a + b, 0) < 1)
            //             price += prices['прямые']['Шлифовка к покраске'] + prices['обратки']['Шлифовка к покраске']
            //         else {
            //             price += prices['обратки']['Шлифовка к покраске'] * publishers.filter(({ sides }) => sides === 2).map(({ square }) => square).reduce((a, b) => a + b, 0)
            //             price += prices['прямые']['Шлифовка к покраске'] * publishers.filter(({ stages }) => !stages.find(item => item.stage === 'Нанесение изолятора')).map(({ square }) => square).reduce((a, b) => a + b, 0)
            //             price += prices['фрезированные']['Шлифовка к покраске'] * publishers.filter(({ stages }) => stages.find(item => item.stage === 'Нанесение изолятора')).map(({ square }) => square).reduce((a, b) => a + b, 0)
            //         }
            //     } else {
            //         price += prices['обратки']['Шлифовка к покраске'] * publishers.filter(({ sides }) => sides === 2).map(({ square }) => square).reduce((a, b) => a + b, 0)
            //         price += prices['прямые']['Шлифовка к покраске'] * publishers.filter(({ stages }) => !stages.find(item => item.stage === 'Нанесение изолятора')).map(({ square }) => square).reduce((a, b) => a + b, 0)
            //         price += prices['фрезированные']['Шлифовка к покраске'] * publishers.filter(({ stages }) => stages.find(item => item.stage === 'Нанесение изолятора')).map(({ square }) => square).reduce((a, b) => a + b, 0)
            //     }
            // } else if (stage.stage === 'Аппликация') {
            //     price += prices['Упаковка']['Аппликация'] * publishers.filter(({ stages }) => stages.filter(({ stage }) => stage === 'Аппликация').map(({ index }) => index).includes(stage.index)).map(({ square }) => square).reduce((a, b) => a + b, 0)
            // } else if (stage.stage === 'Покраска') {
            //     if ((publishers.filter(({ sides }) => sides === 2).length === publishers.length || publishers.filter(({ sides }) => sides === 1).length === publishers.length) && (publishers.filter(({ colourType }) => typeof colourType === 'string' && colourType?.toLowerCase().match(/лак/)).length === 0 || publishers.filter(({ colourType }) => typeof colourType === 'string' && colourType?.toLowerCase().match(/лак/)).length === publishers.length || publishers.filter(({ colourType }) => typeof colourType === 'string' && colourType?.toLowerCase().match(/лак2/)).length === publishers.length)) {
            //         if (publishers.map(({ square }) => square).reduce((a, b) => a + b, 0) <= 0.5) {
            //             price += prices['обратки']['Покраска'] * 0.5
            //             price += prices['прямые']['Покраска'] * 0.5
            //             price += prices['фрезированные']['Покраска'] * 0.5
            //             price += prices['лак']['односторонний'] * 0.5
            //             price += prices['лак']['двусторонний'] * 0.5
            //         } else if (publishers.map(({ square }) => square).reduce((a, b) => a + b, 0) < 1) {
            //             price += prices['обратки']['Покраска']
            //             price += prices['прямые']['Покраска']
            //             price += prices['фрезированные']['Покраска']
            //             price += prices['лак']['односторонний']
            //             price += prices['лак']['двусторонний']
            //         } else {
            //             price += prices['обратки']['Покраска'] * publishers.filter(({ sides }) => sides === 2).map(({ square }) => square).reduce((a, b) => a + b, 0)
            //             price += prices['прямые']['Покраска'] * publishers.filter(({ stages }) => !stages.find(item => item.stage === 'Нанесение изолятора')).map(({ square }) => square).reduce((a, b) => a + b, 0)
            //             price += prices['фрезированные']['Покраска'] * publishers.filter(({ stages }) => stages.find(item => item.stage === 'Нанесение изолятора')).map(({ square }) => square).reduce((a, b) => a + b, 0)
            //             price += prices['лак']['односторонний'] * publishers.filter(({ colourType }) => typeof colourType === 'string' && colourType?.toLowerCase().match(/лак/)).map(({ square }) => square).reduce((a, b) => a + b, 0)
            //             price += prices['лак']['двусторонний'] * publishers.filter(({ colourType }) => typeof colourType === 'string' && colourType?.toLowerCase().match(/лак2/)).map(({ square }) => square).reduce((a, b) => a + b, 0)
            //         }
            //     } else {
            //         price += prices['обратки']['Покраска'] * publishers.filter(({ sides }) => sides === 2).map(({ square }) => square).reduce((a, b) => a + b, 0)
            //         price += prices['прямые']['Покраска'] * publishers.filter(({ stages }) => !stages.find(item => item.stage === 'Нанесение изолятора')).map(({ square }) => square).reduce((a, b) => a + b, 0)
            //         price += prices['фрезированные']['Покраска'] * publishers.filter(({ stages }) => stages.find(item => item.stage === 'Нанесение изолятора')).map(({ square }) => square).reduce((a, b) => a + b, 0)
            //         price += prices['лак']['односторонний'] * publishers.filter(({ colourType }) => typeof colourType === 'string' && colourType?.toLowerCase().match(/лак/)).map(({ square }) => square).reduce((a, b) => a + b, 0)
            //         price += prices['лак']['двусторонний'] * publishers.filter(({ colourType }) => typeof colourType === 'string' && colourType?.toLowerCase().match(/лак2/)).map(({ square }) => square).reduce((a, b) => a + b, 0)
            //     }
            // } else if (stage.stage === 'Полировка') {
            //     if (publishers.filter(({ stages, sides }) => sides === 1 && stages.find(item => item.stage === 'Полировка')).length === publishers.filter(({ stages }) => stages.find(item => item.stage === 'Полировка')).length) {
            //         if (publishers.map(({ square }) => square).reduce((a, b) => a + b, 0) <= 0.5)
            //             price += prices['глянец']['односторонний'] * 0.5
            //         else if (publishers.map(({ square }) => square).reduce((a, b) => a + b, 0) < 1)
            //             price += prices['глянец']['односторонний']
            //         else {
            //             price += prices['глянец']['односторонний'] * publishers.filter(({ stages, sides }) => sides === 1 && stages.find(item => item.stage === 'Полировка')).map(({ square }) => square).reduce((a, b) => a + b, 0)
            //             price += prices['глянец']['двусторонний'] * publishers.filter(({ stages, sides }) => sides === 2 && stages.find(item => item.stage === 'Полировка')).map(({ square }) => square).reduce((a, b) => a + b, 0)
            //         }
            //     } else if (publishers.filter(({ stages, sides }) => sides === 2 && stages.find(item => item.stage === 'Полировка')).length === publishers.filter(({ stages }) => stages.find(item => item.stage === 'Полировка')).length) {
            //         if (publishers.map(({ square }) => square).reduce((a, b) => a + b, 0) <= 0.5)
            //             price += prices['глянец']['двусторонний'] * 0.5
            //         else if (publishers.map(({ square }) => square).reduce((a, b) => a + b, 0) < 1)
            //             price += prices['глянец']['двусторонний']
            //         else {
            //             price += prices['глянец']['односторонний'] * publishers.filter(({ stages, sides }) => sides === 1 && stages.find(item => item.stage === 'Полировка')).map(({ square }) => square).reduce((a, b) => a + b, 0)
            //             price += prices['глянец']['двусторонний'] * publishers.filter(({ stages, sides }) => sides === 2 && stages.find(item => item.stage === 'Полировка')).map(({ square }) => square).reduce((a, b) => a + b, 0)
            //         }
            //     } else {
            //         price += prices['глянец']['односторонний'] * publishers.filter(({ stages, sides }) => sides === 1 && stages.find(item => item.stage === 'Полировка')).map(({ square }) => square).reduce((a, b) => a + b, 0)
            //         price += prices['глянец']['двусторонний'] * publishers.filter(({ stages, sides }) => sides === 2 && stages.find(item => item.stage === 'Полировка')).map(({ square }) => square).reduce((a, b) => a + b, 0)
            //     }
            // } else if (stage.stage === 'Упаковка') {
            //     if (publishers.filter(({ stages, sides }) => sides === 1 && !stages.find(item => item.stage === 'Полировка')).length === publishers.length) {
            //         if (publishers.map(({ square }) => square).reduce((a, b) => a + b, 0) <= 0.5)
            //             price += prices['Упаковка']['Упаковка'] * 0.5 + prices['Упаковка']['чистка'] * 0.5
            //         else if (publishers.map(({ square }) => square).reduce((a, b) => a + b, 0) < 1)
            //             price += prices['Упаковка']['Упаковка'] + prices['Упаковка']['чистка']
            //         else {
            //             price += prices['Упаковка']['Упаковка'] * publishers.map(({ square }) => square).reduce((a, b) => a + b, 0)
            //             price += prices['Упаковка']['чистка'] * publishers.filter(({ stages, sides }) => sides === 1 && !stages.find(item => item.stage === 'Полировка')).map(({ square }) => square).reduce((a, b) => a + b, 0)
            //         }
            //     } else if (publishers.filter(({ stages, sides }) => sides === 2 || stages.find(item => item.stage === 'Полировка')).length === publishers.length) {
            //         if (publishers.map(({ square }) => square).reduce((a, b) => a + b, 0) <= 0.5)
            //             price += prices['Упаковка']['Упаковка'] * 0.5
            //         else if (publishers.map(({ square }) => square).reduce((a, b) => a + b, 0) < 1)
            //             price += prices['Упаковка']['Упаковка']
            //         else
            //             price += prices['Упаковка']['Упаковка'] * publishers.map(({ square }) => square).reduce((a, b) => a + b, 0)
            //     } else {
            //         price += prices['Упаковка']['чистка'] * publishers.filter(({ stages, sides }) => sides === 1 && !stages.find(item => item.stage === 'Полировка')).map(({ square }) => square).reduce((a, b) => a + b, 0)
            //         // price += prices['Упаковка']['Аппликация'] * publishers.filter(({ sides }) => sides === 2).map(({ square }) => square).reduce((a, b) => a + b, 0)
            //         price += prices['Упаковка']['Упаковка'] * publishers.map(({ square }) => square).reduce((a, b) => a + b, 0)
            //     }
            // }

            // if (publishers.find(({ stages }) => stages.find(item => item.stage === 'Шлифовка к изолятору' && item.H))) {
            //     if (stage.stage === 'Шлифовка к изолятору') {
            //         price += prices['ручки']['Шлифовка к изолятору'] * publishers.filter(({ stages }) => stages.find(item => item.stage === 'Шлифовка к изолятору' && item.H)).map(({ amount }) => amount).reduce((a, b) => a + b, 0)
            //         price += prices['ручки']['Нанесение изолятора'] * publishers.filter(({ stages }) => stages.find(item => item.stage === 'Шлифовка к изолятору' && item.H)).map(({ amount }) => amount).reduce((a, b) => a + b, 0)
            //     } else if (stage.stage === 'Шлифовка к грунту') {
            //         price += prices['ручки']['Шлифовка к грунту'] * publishers.filter(({ stages }) => stages.find(item => item.stage === 'Шлифовка к изолятору' && item.H)).map(({ amount }) => amount).reduce((a, b) => a + b, 0)
            //     } else if (stage.stage === 'Нанесение грунта') {
            //         price += prices['ручки']['Нанесение грунта'] * publishers.filter(({ stages }) => stages.find(item => item.stage === 'Шлифовка к изолятору' && item.H)).map(({ amount }) => amount).reduce((a, b) => a + b, 0)
            //     } else if (stage.stage === 'Шлифовка к покраске') {
            //         price += prices['ручки']['Шлифовка к покраске'] * publishers.filter(({ stages }) => stages.find(item => item.stage === 'Шлифовка к изолятору' && item.H)).map(({ amount }) => amount).reduce((a, b) => a + b, 0)
            //     } else if (stage.stage === 'Покраска') {
            //         price += prices['ручки']['Покраска'] * publishers.filter(({ stages }) => stages.find(item => item.stage === 'Шлифовка к изолятору' && item.H)).map(({ amount }) => amount).reduce((a, b) => a + b, 0)
            //     }
            // } else if (publishers.find(({ stages }) => stages.find(item => item.stage === 'Шлифовка к грунту' && item.H))) {
            //     if (stage.stage === 'Шлифовка к грунту') {
            //         price += prices['ручки']['Шлифовка к изолятору'] * publishers.filter(({ stages }) => stages.find(item => item.stage === 'Шлифовка к грунту' && item.H)).map(({ amount }) => amount).reduce((a, b) => a + b, 0)
            //         price += prices['ручки']['Нанесение изолятора'] * publishers.filter(({ stages }) => stages.find(item => item.stage === 'Шлифовка к грунту' && item.H)).map(({ amount }) => amount).reduce((a, b) => a + b, 0)
            //         price += prices['ручки']['Шлифовка к грунту'] * publishers.filter(({ stages }) => stages.find(item => item.stage === 'Шлифовка к грунту' && item.H)).map(({ amount }) => amount).reduce((a, b) => a + b, 0)
            //     } else if (stage.stage === 'Нанесение грунта') {
            //         price += prices['ручки']['Нанесение грунта'] * publishers.filter(({ stages }) => stages.find(item => item.stage === 'Шлифовка к грунту' && item.H)).map(({ amount }) => amount).reduce((a, b) => a + b, 0)
            //     } else if (stage.stage === 'Шлифовка к покраске') {
            //         price += prices['ручки']['Шлифовка к покраске'] * publishers.filter(({ stages }) => stages.find(item => item.stage === 'Шлифовка к грунту' && item.H)).map(({ amount }) => amount).reduce((a, b) => a + b, 0)
            //     } else if (stage.stage === 'Покраска') {
            //         price += prices['ручки']['Покраска'] * publishers.filter(({ stages }) => stages.find(item => item.stage === 'Шлифовка к грунту' && item.H)).map(({ amount }) => amount).reduce((a, b) => a + b, 0)
            //     }
            // }

            let addition = []
            if (publishers.find(({ stages }) => stages.find(({ stage }) => stage === 'Шлифовка к изолятору')) && stage.stage.match(/Шлифовка|Распил/))
                addition.push(`Ф: ${round(publishers.filter(({ stages }) => stages.find(({ stage }) => stage === 'Шлифовка к изолятору')).map(({ square }) => square).reduce((a, b) => a + b, 0), 2)}`)
            if (publishers.find(({ colourType, stages }) => typeof colourType === 'string' && colourType?.toLowerCase().match(/лак/) && (stage.stage === 'Полировка' ? stages.find(({ stage }) => stage === 'Полировка') : true)) && stage.stage.match(/Покраска|Полировка/))
                addition.push(`Л: ${round(publishers.filter(({ colourType, stages }) => typeof colourType === 'string' && colourType?.toLowerCase().match(/лак/) && (stage.stage === 'Полировка' ? stages.find(({ stage }) => stage === 'Полировка') : true)).map(({ square }) => square).reduce((a, b) => a + b, 0), 2)}`)
            if (publishers.find(({ colourType, stages }) => typeof colourType === 'string' && colourType?.toLowerCase().match(/лак2/) && (stage.stage === 'Полировка' ? stages.find(({ stage }) => stage === 'Полировка') : true)) && stage.stage.match(/Покраска|Полировка/))
                addition.push(`Л2: ${round(publishers.filter(({ colourType, stages }) => typeof colourType === 'string' && colourType?.toLowerCase().match(/лак2/) && (stage.stage === 'Полировка' ? stages.find(({ stage }) => stage === 'Полировка') : true)).map(({ square }) => square).reduce((a, b) => a + b, 0), 2)}`)
            addition = addition.join('\r\n')

            return {
                stage: stage.stage,
                term: Math.max(...publishers.map(({ stages }) => stages.find(item => item.stage === stage.stage)?.term || 0)),
                value: round(publishers.filter(({ stages }) => stages.find(item => item.stage === stage.stage && item.index === stage.index)).map(({ square, colourType }) => square * (stage.stage === 'Покраска' && typeof colourType === 'string' && colourType?.toLowerCase().match(/лак/) ? 2 : 1)).reduce((a, b) => a + b, 0), 2),
                weekend: stage.weekend,
                addition,
                H: publishers.find(publisher => publisher.stages.find(item => item.stage === stage.stage && item.H)) ? true : false
            }
        })
        if (stages.find(stage => stage.H && stage.stage === 'Шлифовка к изолятору') && stages.find(stage => stage.H && stage.stage === 'Шлифовка к грунту')) {
            stages.find(stage => stage.H && stage.stage === 'Шлифовка к изолятору').A = true
            stages.find(stage => stage.H && stage.stage === 'Шлифовка к грунту').term -= 1
            stages.find(stage => stage.H && stage.stage === 'Шлифовка к грунту').H = false
        }

        let last = created
        while ([0, 6].includes(new Date(last).getDay()))
            last += Day

        for (let i = 0; i < stages.length; i++) {
            if (stages[i].H || stages[i].A) {
                stages[i].delay = 0
                let add = last + Day
                while ([0, 6].includes(new Date(add).getDay()))
                    add += Day
                stages.splice(i, 0, {
                    days: `${add}`,
                    term: 1,
                    weekend: false,
                    stage: 'Столярка',
                    value: stages[i].value,
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
                continue
            }

            let days = []
            let term = stages[i].term
            if (!term)
                days.push(`${last}`)
            while (term) {
                last += Day
                if (stages[i].stage === 'Полировка')
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
            number: JSON.parse(ws['N7'].v?.match(/\d+/)[0]),
            created: created ? new Date(created) : undefined,
            lastDate: lastDate ? new Date(lastDate) : undefined,
            stages,
            costumer: workbook.Sheets['СЧЕТ']['P6'].v,
        }

        return resolve(order)
    })
}

module.exports = newOrder
