import * as fs from "fs";
import * as mammoth from 'mammoth';
import { Document, Packer, Paragraph, Table, TableRow, TableCell, TextRun } from "docx";

const doc = new Document({
    sections: [{
        children: [
            new Paragraph({
                children: [new TextRun('A vida é bastante bele você souber ser neutro a tudo que lhe ocorre.')]
            }),
            new Paragraph({
                children: [new TextRun('As coisas da vida são belas')]
            })
        ],
    }],
});

// Used to export the file into a .docx file
Packer.toBuffer(doc).then((buffer) => {
    fs.writeFileSync("MainDocx.docx", buffer);
});

const inputDocx = 'MainDocx.docx';
const outputHtml = 'Maindox.html';

fs.promises.readFile(inputDocx, 'binary')
  .then(data => mammoth.convertToHtml({ buffer: Buffer.from(data, 'binary') }))
  .then(result => fs.promises.writeFile(outputHtml, result.value))
  .then(() => {
    console.log(`O arquivo HTML foi gerado em ${outputHtml}`);
  })
  .catch(error => {
    console.error(`Erro na conversão: ${error}`);
  });
