const axios = require('axios')
const docx = require("docx")
const moment = require("moment")
const fs = require('fs');
const { createEpiTable, createClauseWithParagraphs, countDays } = require('./contractUtils');

const caminhoParaJson = 'db.json'

fs.readFile(caminhoParaJson, 'utf8', async (err, data) => {
    if(err){
        console.error('Erro ao ler arquivo Json' + err);
        return
    }
    let task = JSON.parse(data);

    const { Document, Packer } = docx;
    const doc = new Document({
        creator: "Usuário criador",
        description: `Contrato`,
        title: 'Contrato',
        sections: [],
    })
    await createEpiTable(doc, task[0].task_executors[1]);

    Packer.toBuffer(doc).then((buffer) => {
        fs.writeFileSync("DocumentoEPI.docx", buffer);
    });
})