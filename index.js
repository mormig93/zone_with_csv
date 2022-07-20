const CsvReader = require('promised-csv');
var reader = new CsvReader();
const dataCsv = [];

function readCSV() {
    return new Promise((resolve, reject) => {
        reader.on('row', data => {
            dataCsv.push(data);
        });

        reader.on('done', () => {
            resolve(dataCsv);
        });

        reader.on('error', err => reject(err));

        reader.read('./new.csv');
    });
}

async function main() {
    await readCSV()
    const data = []
    const body = []
    const head = dataCsv[0]

    console.log('columnas: ' + head.length)
    console.log('items: ' + (dataCsv.length - 1))

    for (let i = 0; i < dataCsv.length - 1; i++) {
        body.push(dataCsv[i + 1])
    }

    for (let i = 0; i < body.length; i++) {
        let list = {}

        //first struct object data
        /* for (let j = 0; j < tamColumns; j++) {
            let json = `{ "${head[j]}":"${body[i][j]}" }`
            Object.assign(list, JSON.parse(json))
        } */
        //second struct object data
        for (let j = 0; j < head.length; j++) {
            let json = {}
            json[head[j]] = body[i][j]
            Object.assign(list, json)
        }
        data.push(list)
    }
    console.log(data)
}

main()



