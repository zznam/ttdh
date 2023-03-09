import csv from 'csv-parser';
import { createReadStream } from 'fs';

import { createObjectCsvWriter as createCsvWriter } from 'csv-writer';
const csvWriter = createCsvWriter({
    path: 'out.csv',
    header: [{
            id: 'Name',
            title: 'Name'
        },
        {
            id: 'Phone',
            title: 'Phone'
        },
        {
            id: 'Position',
            title: 'Position'
        },
    ]
});
let list1 = [];
let list2 = [];
createReadStream('data.csv')
    .pipe(csv())
    .on('data', (row) => {
        list1.push(row);
    })
    .on('end', () => {
        console.log('CSV file successfully processed');
        console.log(list1.length);

        createReadStream('danh_sach_tnv.csv')
            .pipe(csv())
            .on('data', (row) => {
                list2.push(row);
            })
            .on('end', () => {
                console.log('CSV file successfully processed');
                console.log(list2.length);

                for (let i = 0; i < list1.length; i++) {
                    const e1 = list1[i];
                    let isFound = false;
                    for (let i = 0; i < list2.length; i++) {
                        const e2 = list2[i];
                        if (e2.Phone.indexOf(e1.Phone) >= 0) {
                            e1.Position = e2.Position;
                            isFound = true;
                        }
                    }
                    if (!isFound) {
                        e1.Position = "Chưa xác định";
                    }
                }
                csvWriter
                    .writeRecords(list1)
                    .then(() => console.log('The CSV file was written successfully'));

            });

    });