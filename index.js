import * as fs from "fs";
import * as mammoth from 'mammoth';
import { Document, Packer, Paragraph, Table, TableRow, TableCell } from "docx";


const doc = new Document({
    sections: [
        {
            properties: {},
            children: [
                new Table({
                    rows: [
                        new TableRow({
                            children: [
                                new TableCell({
                                    children: [new Paragraph("Kayo Leanndro")],
                                }),
                                new TableCell({
                                    children: [new Paragraph("19 Anos")],
                                }),
                            ],
                        }),
                        new TableRow({
                            children: [
                                new TableCell({
                                    children: [new Paragraph("Leticia")],
                                }),
                                new TableCell({
                                    children: [new Paragraph("99 Anos")],
                                }),
                            ],
                        }),
                        new TableRow({
                            children: [
                                new TableCell({
                                    children: [new Paragraph("Sarah")]
                                }),
                                new TableCell({
                                    children: [new Paragraph("18 Anos")]
                                }),
                            ],
                        }),
                    ],
                }),
            ],
        },
    ],
});


// Used to export the file into a .docx file
Packer.toBuffer(doc).then((buffer) => {
    fs.writeFileSync("My Document.docx", buffer);
});

const inputDocx = 'My Document.docx';
const outputHtml = 'output.html';

fs.promises.readFile(inputDocx, 'binary')
  .then(data => mammoth.convertToHtml({ buffer: Buffer.from(data, 'binary') }))
  .then(result => fs.promises.writeFile(outputHtml, result.value))
  .then(() => {
    console.log(`O arquivo HTML foi gerado em ${outputHtml}`);
  })
  .catch(error => {
    console.error(`Erro na conversão: ${error}`);
  });
