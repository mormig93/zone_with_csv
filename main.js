const csv = require('csv-parser')
const fs = require('fs')
const axios = require("axios").default;

//for telegram bot api
const tg_token = '5212294496:AAGcQF613aFdOXt-RJWi42ijsxSXBPtahdM'
const tg_chatId = '@synctest'
const post_type = ['sendMessage', 'sendPhoto']
const url = `https://api.telegram.org/bot${tg_token}`
const options = {
    method: 'GET',
    headers: {
        'User-Agent': 'Telegram Bot SDK - (https://github.com/irazasyed/telegram-bot-sdk)',
        'Content-Type': 'application/json'
    },
    data: {
        chat_id: tg_chatId,
        parse_mode: 'HTML',
    }
}

//time per each post
const time_per_post = 5 * 1000

//for store csv data
const data_from_csv = [];


fs.createReadStream('./data/items.csv')
    .pipe(csv({ headers: true }))
    .on('data', (data) => data_from_csv.push(data))
    .on('end', () => {
        let time = 1
        const processID = setInterval(function () {
            if (data_from_csv[time] === undefined) {
                clearInterval(processID)
            } else {
                shareToSocial(data_from_csv[time])
                time++
            }
        }, time_per_post);
    });

function shareToSocial(info_from_csv) {
    let pre_title = (info_from_csv._1).toUpperCase()
    let img = info_from_csv._2
    let link = info_from_csv._7
    let pre_promocode = info_from_csv._10
    let pre_discount = info_from_csv._4

    let limit_char = 30
    let title = pre_title.slice(0, limit_char) + (pre_title.length > limit_char ? "..." : "")
    let promocode = `\n<b>PROMO CODE</b>: <code>${pre_promocode}</code>\n`
    let discount = pre_discount + '% ON'

    if (link) {
        let message = `
        📢 SAVE OFF ${pre_discount ? discount : 'ON'} ${title}${pre_promocode ? promocode : '\n'}👇👇👇\n<a href="${link}">Clic Here</a>
        `

        img ? sendPost(post_type[1], message, img) : sendPost(post_type[0], message, img)
    } else {
        sendPost(post_type[0], 'link no exist', img)
    }
}

function sendPost(type_post, text, img) {
    let option = options

    if (type_post === 'sendMessage') {
        option.url = url + '/' + type_post
        option.data.text = text
    }
    if (type_post === 'sendPhoto') {
        option.url = url + '/' + type_post
        option.data.photo = img
        option.data.caption = text
    }

    axios.request(option)
        .then(function (response) {
            console.log('true');
            //console.log(response.data.ok);
        }).catch(function (error) {
            console.log('error');
        });
}









