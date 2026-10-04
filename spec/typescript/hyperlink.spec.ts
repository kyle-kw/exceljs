import {expect} from 'chai';
import {Workbook, CellHyperlinkValue} from '../..';

describe('hyperlink types', () => {
  it('can round-trip rich text and plain text hyperlinks', async () => {
    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet('Hyperlinks');
    const richTextLink: CellHyperlinkValue = {
      text: {
        richText: [
          {text: 'Excel', font: {bold: true}},
          {text: 'JS'},
        ],
      },
      hyperlink: 'https://github.com/exceljs/exceljs',
    };
    const plainTextLink: CellHyperlinkValue = {
      text: 'ExcelJS',
      hyperlink: 'https://github.com/exceljs/exceljs',
    };

    worksheet.getCell('A1').value = richTextLink;
    worksheet.getCell('B1').value = plainTextLink;

    const reloaded = new Workbook();
    await reloaded.xlsx.load(await workbook.xlsx.writeBuffer());
    const reloadedSheet = reloaded.getWorksheet('Hyperlinks');
    const richTextValue = reloadedSheet.getCell('A1').value as CellHyperlinkValue;
    const plainTextValue = reloadedSheet.getCell('B1').value as CellHyperlinkValue;

    expect(richTextValue).to.deep.equal(richTextLink);
    expect(plainTextValue).to.deep.equal(plainTextLink);
    if (typeof richTextValue.text !== 'string') {
      expect(richTextValue.text.richText[0].font.bold).to.equal(true);
    }
  });
});
