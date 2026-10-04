const fs = require('fs');

const testXformHelper = require('../test-xform-helper');

const DrawingXform = verquire('xlsx/xform/drawing/drawing-xform');
const XmlStream = verquire('utils/xml-stream');

const options = {
  rels: {
    rId1: {Target: '../media/image1.jpg'},
    rId2: {Target: '../media/image2.jpg'},
  },
  mediaIndex: {image1: 0, image2: 1},
  media: [{}, {}],
};

const expectations = [
  {
    title: 'Drawing 1',
    create() {
      return new DrawingXform({tag: 'xdr:from'});
    },
    initialModel: require('./data/drawing.1.0'),
    preparedModel: require('./data/drawing.1.1'),
    xml: fs.readFileSync(`${__dirname}/data/drawing.1.2.xml`).toString(),
    parsedModel: require('./data/drawing.1.3'),
    reconciledModel: require('./data/drawing.1.4'),
    tests: ['prepare', 'parse', 'reconcile'],
    options,
  },
];

describe('DrawingXform', () => {
  testXformHelper(expectations);

  [false, true].forEach(composite => {
    it(`renders unique picture creation IDs${composite ? ' in a composite' : ''}`, () => {
      const xform = new DrawingXform();
      const xmlStream = new XmlStream();
      const expectation = expectations[0];

      if (composite) {
        xmlStream.openNode('compy');
        xmlStream.leafNode('pre');
      }
      xform.render(xmlStream, expectation.preparedModel);
      if (composite) {
        xmlStream.leafNode('post');
        xmlStream.closeNode();
      }

      const {xml} = xmlStream;
      const creationIds = [];
      const normalize = value => value.replace(/(<a16:creationId\b[^>]*\bid=")[^"]+/g, '$1creation-id');
      xml.replace(/<a16:creationId\b[^>]*\bid="([^"]+)"/g, (match, id) => {
        creationIds.push(id);
        return match;
      });
      expect(creationIds).to.have.length(2);
      expect(new Set(creationIds).size).to.equal(2);
      creationIds.forEach(id => {
        expect(id).to.match(/^\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}$/);
      });

      const expectedXml = composite ? `<compy><pre/>${expectation.xml}<post/></compy>` : expectation.xml;
      expect(normalize(xml)).xml.to.equal(normalize(expectedXml));
    });
  });
});
