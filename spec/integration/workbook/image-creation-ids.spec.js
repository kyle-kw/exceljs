const fs = require('fs');
const JSZip = require('jszip');

const ExcelJS = verquire('exceljs');

describe('Workbook', () => {
  describe('Image creation IDs', () => {
    it('keeps picture IDs unique across worksheets and after a roundtrip', async () => {
      const imageData = fs.readFileSync(`${__dirname}/../data/image.png`);
      const bluePixel = Buffer.from(
        'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGNgYPj/HwADAgH/5ncLrgAAAABJRU5ErkJggg==',
        'base64'
      );
      const workbook = new ExcelJS.Workbook();
      const first = workbook.addWorksheet('Images');
      workbook.addWorksheet('No images');
      const second = workbook.addWorksheet('More images');
      const imageId = workbook.addImage({buffer: imageData, extension: 'png'});
      const blueId = workbook.addImage({buffer: bluePixel, extension: 'png'});
      const hyperlinks = {
        hyperlink: 'https://example.com/image',
        tooltip: 'Image hyperlink',
      };

      first.addImage(imageId, {
        tl: {col: 0, row: 0},
        ext: {width: 100, height: 80},
        editAs: 'absolute',
        hyperlinks,
      });
      first.addImage(blueId, {
        tl: {col: 0, row: 10},
        ext: {width: 100, height: 100},
      });
      first.addImage(imageId, 'C3:E6');
      second.addImage(imageId, {
        tl: {col: 1, row: 4},
        ext: {width: 50, height: 60},
        editAs: 'oneCell',
      });

      async function checkPackage(buffer) {
        const zip = await JSZip.loadAsync(buffer);
        const drawings = zip.file(/^xl\/drawings\/drawing\d+\.xml$/);
        expect(drawings).to.have.length(2);

        const drawingXml = await Promise.all(drawings.map(drawing => drawing.async('string')));
        const creationIds = [];
        drawingXml.forEach(xml => {
          const pictureCount = (xml.match(/<xdr:pic>/g) || []).length;
          const ids = [];
          xml.replace(/<a16:creationId\b[^>]*\bid="([^"]+)"/g, (match, id) => {
            ids.push(id);
            return match;
          });
          expect(ids).to.have.length(pictureCount);
          creationIds.push(...ids);
        });
        expect(creationIds).to.have.length(4);
        expect(new Set(creationIds).size).to.equal(4);
        creationIds.forEach(id => {
          expect(id).to.match(/^\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}$/);
        });

        const media = zip.file(/^xl\/media\/image\d+\.png$/);
        expect(media).to.have.length(2);
        const mediaData = await Promise.all(media.map(image => image.async('nodebuffer')));
        expect(mediaData).to.have.deep.members([imageData, bluePixel]);
      }

      function checkImages(wb) {
        expect(wb.getWorksheet('No images').getImages()).to.have.length(0);
        const firstImages = wb.getWorksheet('Images').getImages();
        const secondImages = wb.getWorksheet('More images').getImages();
        expect(firstImages).to.have.length(3);
        expect(secondImages).to.have.length(1);

        const oneCell = firstImages[0].range;
        expect(oneCell.tl.col).to.equal(0);
        expect(oneCell.tl.row).to.equal(0);
        expect(oneCell.ext).to.deep.equal({width: 100, height: 80});
        expect(oneCell.editAs).to.equal('absolute');
        expect(oneCell.hyperlinks).to.deep.equal(hyperlinks);

        const blue = firstImages[1].range;
        expect(blue.tl.col).to.equal(0);
        expect(blue.tl.row).to.equal(10);
        expect(blue.ext).to.deep.equal({width: 100, height: 100});
        expect(wb.getImage(firstImages[1].imageId).buffer).to.deep.equal(bluePixel);

        const twoCell = firstImages[2].range;
        expect(twoCell.tl.col).to.equal(2);
        expect(twoCell.tl.row).to.equal(2);
        expect(twoCell.br.col).to.equal(5);
        expect(twoCell.br.row).to.equal(6);
        expect(twoCell.editAs).to.equal('oneCell');

        const otherSheet = secondImages[0].range;
        expect(otherSheet.tl.col).to.equal(1);
        expect(otherSheet.tl.row).to.equal(4);
        expect(otherSheet.ext).to.deep.equal({width: 50, height: 60});
        expect(otherSheet.editAs).to.equal('oneCell');

        [firstImages[0], firstImages[2], secondImages[0]].forEach(image => {
          expect(wb.getImage(image.imageId).buffer).to.deep.equal(imageData);
        });
      }

      const buffer = await workbook.xlsx.writeBuffer();
      await checkPackage(buffer);
      const restored = new ExcelJS.Workbook();
      await restored.xlsx.load(buffer);
      checkImages(restored);

      const rewritten = await restored.xlsx.writeBuffer();
      await checkPackage(rewritten);
      const reread = new ExcelJS.Workbook();
      await reread.xlsx.load(rewritten);
      checkImages(reread);
    });
  });
});
