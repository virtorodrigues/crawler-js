const puppeteer = require('puppeteer');
const xl = require('excel4node');
const wb = new xl.Workbook();
const ws = wb.addWorksheet('nome planilha');
const reader = require('xlsx');
const file = reader.readFile('./sites.xlsx');

(async () => {

  //--------------GETTING SITES FROM EXCEL------------------------------------------------------------------------------------------

  const sheets = file.SheetNames;
  let listSiteXlsx = [];

  for (let i = 0; i < sheets.length; i++) {
    const temp = reader.utils.sheet_to_json(file.Sheets[file.SheetNames[i]]);
    temp.forEach((res) => {
      listSiteXlsx.push(res);
    })
  }

  const sites = listSiteXlsx.map(d => d.sites); //sites is first row of excel

  //--------------GETTING ELEMENTS FROM SITE------------------------------------------------------------------------------------------

  const browser = await puppeteer.launch();
  const page = await browser.newPage();

  let elements = [];

  for (currentURL of sites) {
    await page.goto(currentURL);
    elements.push(await page.evaluate(() => ({
      title: document.querySelector('#DynamicHeading_productTitle')?.innerText || '',
      description: document.querySelector('#product-description')?.innerText || '',
      realeaseData: document.querySelector('#releaseDate-toggle-target span')?.innerText || '',
      price: document.querySelector('.price-disclaimer span') ? document.querySelector('.price-disclaimer span')?.innerText : document.querySelector('.pi-price-text span')?.innerText || ''
    })));
  }
  await browser.close();

  //---------------FORMATTING ELEMENTS-----------------------------------------------------------------------------------------

  elements.map(element => {
    element.title = element.title.replace('&amp;', '&');
    element.price = element.price.replace('R$', '').replace(',', '.');
    element.realeaseData = element.realeaseData.split('/')[2];
    element.description = element.description.replace('&amp;', '&');
    return element;
  });

  //---------------WRITING ELEMENTS IN EXCEL-----------------------------------------------------------------------------------------

  const columnsName = [
    'title',
    'price',
    'description',
    'realeaseData',
  ];

  let headIndex = 1;

  columnsName.forEach(column => {
    ws.cell(1, headIndex++).string(column);
  });

  let rowIndex = 2;

  elements.map(record => {
    let columnIndex = 1;

    Object.keys(record).forEach(columnName => {
      ws.cell(rowIndex, columnIndex++).string(record[columnName]);
    });
    rowIndex++;
  });

  wb.write('arquivo.xlsx');
})();
