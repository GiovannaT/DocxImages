const docx = require("docx")
const fs = require('fs');

const { createPageFooter, createPageHeader, createPersonalDataTable, createExamsParagraphs, createEnvironmentsParagraphs } = require('./contractUtils');

const caminhoParaJson = 'dbAjudanteObra.json';
const contractType = 'Admission';

fs.readFile(caminhoParaJson, 'utf8', async (err, data) => {
    if(err){
        console.error('Erro ao ler arquivo Json' + err);
        return
    }

    let user = JSON.parse(data);
    const company = 'HABIT'

    const { Document, Packer, Paragraph, TextRun, Header, Footer, AlignmentType, convertInchesToTwip} = docx;

    const doc = new Document({
        creator: "Usuário criador",
        description: `Contrato`,
        title: 'Contrato',
        sections: [{
            headers: {
                default: new Header({ 
                    children: [
                        createPageHeader(company)
                    ],
                }),
            },
            footers: {
                default: new Footer({
                    children: [
                       ...createPageFooter(company)
                    ],
                }),
            },
            properties: {
                page:{
                    margin: {
                        top: 800,     
                        right: 800,   
                        bottom: 800,  
                        left: 800    
                    },
                },
              },
            children: [
                new Paragraph({
                    alignment: AlignmentType.CENTER,
                    indent: {
                        left: convertInchesToTwip(0.5),
                        right: convertInchesToTwip(0.5),
                    },
                    spacing: {
                        line: 276,
                        after: 200,
                    },
                    children: [
                        new TextRun({
                            text: 'GUIA DE ENCAMINHAMENTO PARA EXAME',
                            size: 22,
                            bold: true,
                            font: "Calibri",
                        }),
                    ],
                }),
            ]
        }],
    })

    await createPersonalDataTable(company, user, doc, contractType);
    createEnvironmentsParagraphs(user);

    Packer.toBuffer(doc).then((buffer) => {
        fs.writeFileSync("AjudanteObra.docx", buffer);
    });
})