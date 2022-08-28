const fetch = require('node-fetch')
const excel = require('excel4node')
const wb = new excel.Workbook()
const workSheet = wb.addWorksheet('Countries list')


async function func() {
    const response = await fetch('https://restcountries.com/v3.1/all?fields=name,capital,area,currencies')
    const resJSON = await response.json();
    
    const dataStyle = wb.createStyle({
        font: {
            size: 12
        },
        alignment: {
            // horizontal: 'center',
            vertical: 'center',
            wrapText: true,
        },
        numberFormat: '#,##0.00'
    })

    //Filling the data rows
    for (let index = 0; index < resJSON.length; index++) {
        const country = resJSON[index];

        ////Name column
        workSheet.cell(index+3, 1)
        .string(country['name']['common'])
        .style(dataStyle)

        ////Capital column
        if (country['capital'][0] == undefined) {
            workSheet.cell(index+3, 2)
            .string('-')
            .style(dataStyle)
        } else {
            workSheet.cell(index+3, 2)
            .string(country['capital'][0])
            .style(dataStyle)
        }

        ////Area column
        workSheet.cell(index+3, 3)
        .number(country['area'])
        .style(dataStyle)


        ////Currencies column
        var keys = Object.keys(country['currencies'])
        var currString = ""
        if (keys[0] == undefined) {
            currString = '-'
        } else {
            if (keys.length > 1) {
                for (let index = 0; index < keys.length; index++) {
                    const key = keys[index];
                    if (index == 0) {
                        currString = key + ","
                    } else {
                        if (index == keys.length-1) {
                            currString = currString + key
                        } else {
                            currString = currString + key + ","
                        }
                    }
                }
            } else{
                currString = keys[0]
            }
        }
        workSheet.cell(index+3, 4)
        .string(currString)
        .style(dataStyle)
    }

    //Setting up the worksheet structure

    ////Setting Title
    const titleStyle = wb.createStyle({
        font: {
            bold: true,
            color: '#4F4F4F',
            size: 16
        },
        alignment: {
            horizontal: 'center'
        }
    })

    workSheet.cell(1,1,1,4,true)
    .string('Countries List')
    .style(titleStyle)

    ////Setting Headers
    const sheetHeaders = ['Name','Capital','Area','Currencies']
    const headerStyle = wb.createStyle({
        font: {
            bold: true,
            color: '#808080',
            size: 12
        }
    })
    let headersIndex = 1
    sheetHeaders.forEach(header => {
        workSheet.cell(2, headersIndex)
        .string(header)
        .style(headerStyle)
        headersIndex++
    })
    
    ////Adjusting column sizes
    workSheet.column(1).setWidth(40)
    workSheet.column(2).setWidth(30)
    workSheet.column(3).setWidth(20)
    workSheet.column(4).setWidth(15)
    
    //Creating the xlsx file with the data
    wb.write('list.xlsx')    
}

func()