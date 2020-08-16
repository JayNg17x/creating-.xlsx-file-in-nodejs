'use strict';

const data = [{
  firstName: 'John',
  lastName: 'Bailey',
  purchasePrice: 1000,
  paymentsMade: 100
}, {
  firstName: 'Leonard',
  lastName: 'Clark',
  purchasePrice: 1000,
  paymentsMade: 150
}, {
  firstName: 'Phil',
  lastName: 'Knox',
  purchasePrice: 1000,
  paymentsMade: 200
}, {
  firstName: 'Sonia',
  lastName: 'Glover',
  purchasePrice: 1000,
  paymentsMade: 250
}, {
  firstName: 'Adam',
  lastName: 'Mackay',
  purchasePrice: 1000,
  paymentsMade: 350
}, {
  firstName: 'Lisa',
  lastName: 'Ogden',
  purchasePrice: 1000,
  paymentsMade: 400
}, {
  firstName: 'Elizabeth',
  lastName: 'Murray',
  purchasePrice: 1000,
  paymentsMade: 500
}, {
  firstName: 'Caroline',
  lastName: 'Jackson',
  purchasePrice: 1000,
  paymentsMade: 350
}, {
  firstName: 'Kylie',
  lastName: 'James',
  purchasePrice: 1000,
  paymentsMade: 900
}, {
  firstName: 'Harry',
  lastName: 'Peake',
  purchasePrice: 1000,
  paymentsMade: 1000
}];


const Excel = require('exceljs');
let workbook = new Excel.Workbook();
let worksheet = workbook.addWorksheet('Sample');

// worksheet columns
worksheet.columns = [
	{header: 'First Name', key: 'firstName'},
	{header: 'Last Name', key: 'lastName'},
	{header: 'Purchase Price', key: 'purchasePrice'},
	{header: 'Payments Made', key: 'paymentsMade'},
	{header: 'Amount Remaining', key: 'amountRemaining'},
	{header: '% Remaining', key: 'percentRemaining'}
];

// formating the header
worksheet.columns.forEach(column => {
	column.width = column.header.length < 12 ? 12 : column.header.length;
});

// make the header bold 
worksheet.getRow(1).font = { bold:true };

// insert data into excel
data.forEach((elem, index) => {
	// row 1 is the header
	const rowIndex = index + 2;

	worksheet.addRow({
		...elem,
		amountRemaining: {
			formula: `=C${rowIndex} - D${rowIndex}`
		},
		percentRemaining: {
			formula: `=E${rowIndex} - C${rowIndex}`
		}
	});
});


const totalNumberOfRows = worksheet.rowCount;

// Add the total Rows
worksheet.addRow([
  '',
  'Total',
  {
    formula: `=sum(C2:C${totalNumberOfRows})`
  },
  {
    formula: `=sum(D2:D${totalNumberOfRows})`
  },
  {
    formula: `=sum(E2:E${totalNumberOfRows})`
  },
  {
    formula: `=E${totalNumberOfRows + 1}/C${totalNumberOfRows + 1}`
  }
]);

// FORMATTING DATA

// set the way columns C-F are formatted
const figureColumns = [3,4,5,6];
figureColumns.forEach(index => {
	worksheet.getColumn(index).numFmt = '$0.00';
	worksheet.getColumn(index).alignment = { horizontal: 'center' }; 
});

// column F need to be formatted as a percentage
worksheet.getColumn(6).numFmt = '0.00$';


// FORMATTING BORDERS

// loop through all of the rows and set the outline style
worksheet.eachRow({ includeEmpty: false }, (row, rowNum) => {
	worksheet.getCell(`A${rowNum}`).border = {
		top: { style: 'thin' },
		left: { style: 'thin' },
		bottom: { style: 'none' },
		right: { style: 'none' }
	};

	const insideColumns = ['B', 'C', 'D', 'E'];
	insideColumns.forEach(v => {
		worksheet.getCell(`${v}${rowNum}`).border = {
			top: { style: 'thin' },
			left: { style: 'none' },
			bottom: { style: 'thin' },
			right: { style: 'thin' }
		};
	});

	worksheet.getCell(`F${rowNum}`).border = {
		top: { style: 'thin' },
		left: { style: 'none' },
		bottom: { style: 'thin' },
		right: { style: 'thin' }
	};
});


const totalCell = worksheet.getCell(`B${worksheet.rowCount}`);
totalCell.font = {bold: true};
totalCell.alignment = {horizontal: 'center'};

// Create a freeze pane, which means we'll always see the header as we scroll around.
worksheet.views = [
  { state: 'frozen', xSplit: 0, ySplit: 1, activeCell: 'B2' }
];


// SAVING THE EXCEL FILE 

// keep in mind that reading and writing is promise based 
workbook.xlsx.writeFile('Sample.xlsx');



