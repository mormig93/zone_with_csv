const axios = require("axios").default;
const fs = require('fs');
const path = require('path');
const csv = require ('csv-parser');

//for telegram bot api
const tg_token = '5212294496:AAGcQF613aFdOXt-RJWi42ijsxSXBPtahdM'
const tg_chatId = '@synctest'
const post_type = ['sendMessage', 'sendPhoto']
const url = `https://api.telegram.org/bot${tg_token}/`
const options = {
    method: 'GET',
    headers: {
        'Accept': 'application/json',
        'Content-Type': 'application/json',
        'User-Agent': 'Telegram Bot SDK - (https://github.com/irazasyed/telegram-bot-sdk)'
    },
    data: {
        chat_id: tg_chatId,
        parse_mode: 'HTML',
    }
}

//time per each post
const time_per_post = 1000

//for store csv data
const data_from_csv = [];


fs.createReadStream(path.resolve(__dirname, 'data', 'items.csv'))
    .pipe(csv({ headers: true, separator: ';' }))
    .on('error', (error) => console.log(error))
    .on('data', (row) => data_from_csv.push(row))
    .on('end', () => {
        let position = 1
        const processID = setInterval(function () {
            if (data_from_csv[position] === undefined) {
                clearInterval(processID)
            } else {
                shareToSocial(data_from_csv[position])
                position++
            }
        }, time_per_post);
    });

function shareToSocial(data_from_csv) {
    let pre_title = (data_from_csv._1).toUpperCase()
    let img = data_from_csv._2
    let link = data_from_csv._7
    let pre_promocode = data_from_csv._10
    let pre_discount = data_from_csv._4

    //formating title, title only use limit character for to post
    let limit_char = 30
    let title = pre_title.slice(0, limit_char) + (pre_title.length > limit_char ? "..." : "")
    let promocode = `\n<b>PROMO CODE</b>: <code>${pre_promocode}</code>\n`
    let discount = pre_discount + '% ON'

    if (link) {
        //message formating for post
        let message = `
        📢 SAVE OFF${pre_discount ? discount : 'ON'} ${title}${pre_promocode ? promocode : '\n'}👇👇👇\n<a href="${link}">Clic Here</a>
        `

        img ? sendPost(post_type[1], message, img) : sendPost(post_type[0], message, img)
    }/* else {
        sendPost(post_type[0], 'link no exist', img)
    }*/
}

function sendPost(type_post, text, img) {
    let option = options

    if (type_post === post_type[0]) {
        option.url = url + type_post
        option.data.text = text
    }
    if (type_post === post_type[1]) {
        option.url = url + type_post
        option.data.photo = img
        option.data.caption = text
    }

    axios(option)
        .then(function (response) {
            console.log('true');
        }).catch(function (error) {
            console.log('error');
        });
}









