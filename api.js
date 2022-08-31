const Api = require('./api.js')

const fetch = require('node-fetch')
const xl = require('excel4node');

const wb = new xl.Workbook();
const ws = wb.addWorksheet('Countries List');

const archiveName = 'teste.xlsx'



const url = 'https://restcountries.com/v3.1/all'

fetch(url)
  .then(res => res.json())
  .then(data => apiCountriesExcel(data))
  .catch(err => console.error("erro:", err))


function apiCountriesExcel(countriesIndex) {
  const titleWorksheet = ["Contries List"]
  const headingColumnNames = [
    "Nome",
    "Capital",
    "Area",
    "Currencies"
  ];


  let data = []

  for (i = 0; i < countriesIndex.length; i++) {
    const valueName = String(countriesIndex[i].name.common)
    const valueCapital = String(countriesIndex[i].capital)
    const valueArea = String(countriesIndex[i].area)
    const valueCurrencies = String(JSON.stringify(countriesIndex[i].currencies)).substring(2, 5)

    if (valueCurrencies == 'undefined') {
      data.push({
        name: valueName,
        capital: valueCapital,
        area: valueArea,
        currencies: "-"
      })
    }else if (valueName == 'undefined') {
      data.push({
        name: "-",
        capital: valueCapital,
        area: valueArea,
        currencies: valueCurrencies
      })
    }else if(valueCapital == 'undefined'){
      data.push({
        name: valueName,
        capital: "-",
        area: valueArea,
        currencies: valueCurrencies
      })
    }else if(valueArea == 'undefined'){
      data.push({
        name: valueName,
        capital: valueCapital,
        area: "-",
        currencies: valueCurrencies
      })
    } else {
      data.push({
        name: valueName,
        capital: valueCapital,
        area: valueArea,
        currencies: valueCurrencies
      })
    }

  }

  const titleStyles = wb.createStyle({
    font: {
      bold: true,
      color: '4F4F4F',
      size: 16,
    },
    alignment: {
      wrapText: true,
      horizontal: 'center',
    },
  });


  const ColumnTitlesStyles = wb.createStyle({
    font: {
      bold: true,
      color: '808080',
      size: 12,
    },
    alignment: {
      wrapText: true,
      horizontal: 'center',
    },
  });




  titleWorksheet.forEach((title) => {
    ws.cell(1, 1, 1, 4, true)
      .string(title)
      .style(titleStyles)
  })

  let headingColumnIndex = 1;

  headingColumnNames.forEach((heading) => {
    ws.cell(2, headingColumnIndex++)
      .string(heading)
      .style(ColumnTitlesStyles)

  });

  let rowIndex = 3;

  data.forEach(record => {

    let columnIndex = 1;
    Object.keys(record).forEach(columnName => {
      ws.cell(rowIndex, columnIndex++).string(record[columnName])
    });

    rowIndex++;
  });


  wb.write(archiveName, () =>{
    
  })

}