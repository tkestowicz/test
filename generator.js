var JSZip = require('jszip');
var Docxtemplater = require('docxtemplater');
var expressions = require('angular-expressions');
var fs = require('fs');
var path = require('path');

var angularParser = function (tag) {
    expr = expressions.compile(tag);
    return { get: expr };
}

var DocxGen = {

    constructor() {
        this.doc = new Docxtemplater();
    },

    generate(filename, data) {

        var content = fs
            .readFileSync(path.resolve(__dirname, filename), 'binary');

        var zip = new JSZip(content);

        this.doc.loadZip(zip);

        this.doc.setOptions({ parser: angularParser });

        this.doc.setData(data);

        this.doc.render();
    },

    saveAs(filename) {
        var buf = this.doc.getZip()
            .generate({ type: 'nodebuffer' });

        // buf is a nodejs buffer, you can either write it to a file or do anything else with it.
        fs.writeFileSync(path.resolve(__dirname, filename), buf);
    }
}

module.exports = DocxGen;