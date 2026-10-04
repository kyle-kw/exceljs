import {Workbook, TableStyleProperties} from '../../..';

const workbook = new Workbook();
const worksheet = workbook.addWorksheet('Table styles');

worksheet.addTable({
  name: 'UnstyledTable',
  ref: 'A1',
  style: {theme: null},
  columns: [{name: 'Amount'}],
  rows: [[70.1]],
});

const defaultStyle: TableStyleProperties = {};
const namedStyle: TableStyleProperties = {theme: 'TableStyleMedium2'};
const noTheme: TableStyleProperties = {theme: null};

// @ts-expect-error Only built-in theme names or null are supported.
const invalidStyle: TableStyleProperties = {theme: 'NotATableStyle'};
