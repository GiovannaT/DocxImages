const docx = require("docx")
const fs = require('fs');

const { createXPContract, createClauseHeader, createSignSection, createRenew, createAttributionsSection, createNewSection, createCompanyTextRun, createEmployeeTextRun, createPageFooter, createPageHeader } = require('./contractUtils');

const caminhoParaJson = 'dbUser.json'

fs.readFile(caminhoParaJson, 'utf8', async (err, data) => {
    if(err){
        console.error('Erro ao ler arquivo Json' + err);
        return
    }

    let user = JSON.parse(data);
    const company = 'HABIT'

    const { Document, Packer, Paragraph, TextRun, Header, Footer, AlignmentType, BorderStyle, convertInchesToTwip} = docx;

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
                    borders:{
                        pageBorderBottom: {
                            style: BorderStyle.SINGLE,
                            size: 1*8,
                            color: '000000',
                        },
                        pageBorderLeft: {
                            style: BorderStyle.SINGLE,
                            size: 1*8,
                            color: '000000',
                        },
                        pageBorderRight: {
                            style: BorderStyle.SINGLE,
                            size: 1*8,
                            color: '000000',
                        },
                        pageBorderTop: {
                        style: BorderStyle.SINGLE,
                        size: 1*8,
                        color: '000000',
                        },
                        
                        pageBorders: {
                        display: "allPages", 
                        offsetFrom: "text", 
                        zOrder: "front"
                        }
                    }
                },
              },
            children: [
                createXPContract(),
                createClauseHeader('CLÁUSULA I - DAS PARTES'),
                new Paragraph({
                    alignment: AlignmentType.JUSTIFIED,
                    indent: {
                        left: convertInchesToTwip(0.5),
                        right: convertInchesToTwip(0.5),
                    },
                    spacing: {
                        line: 276,
                    },
                    children: [
                        new TextRun({
                            text: '1.1 ',
                            size: 24,
                            bold: true,
                            font: "Calibri Light",
                        }),
                        ...createCompanyTextRun('AVANCI'),
                    ],
                }),
                new Paragraph({
                    alignment: AlignmentType.JUSTIFIED,
                    margin: {
                        top: 400,     
                        right: 400,   
                        bottom: 400,  
                        left: 400    
                    },
                    indent: {
                        left: convertInchesToTwip(0.5),
                        right: convertInchesToTwip(0.5),
                    },
                    spacing: {
                        line: 276,
                        before: 200,
                    },
                    children: [
                        new TextRun({
                            text: `1.2 `,
                            bold: true,
                            size: 24,
                            font: "Calibri Light",
                        }),
                        ...createEmployeeTextRun(user),
                    ],
                }),
                createClauseHeader('CLÁUSULA II - DAS FUNÇÕES'),
                new Paragraph({
                    alignment: AlignmentType.JUSTIFIED,
                    indent: {
                        left: convertInchesToTwip(0.5),
                        right: convertInchesToTwip(0.5),
                    },
                    spacing: {
                        line: 276,
                        before: 200,
                    },
                    children: [
                        new TextRun({
                            text: '2.1 ',
                            size: 24,
                            bold: true,
                            font: "Calibri Light",
                        }),
                        new TextRun({
                            size: 24,
                            text: 'Fica o(a) ',
                            font: "Calibri Light",
                        }),
                        new TextRun({
                            text: 'EMPREGADO(A) ',
                            size: 24,
                            bold: true,
                            font: "Calibri Light",
                        }),
                        new TextRun({
                            text: 'admitido no quadro de funcionários do(a) ',
                            size: 24,
                            font: "Calibri Light",
                        }),
                        new TextRun({
                            text: 'EMPREGADOR(A) ',
                            size: 24,
                            bold: true,
                            font: "Calibri Light",
                        }),
                        new TextRun({
                            text: `para exercer o cargo de ${user.skill.name}, CBO ${user.skill.cbo}.`,
                            size: 24,
                            font: "Calibri Light",
                        }),
                    ],
                }),
                new Paragraph({
                    alignment: AlignmentType.JUSTIFIED,
                    indent: {
                        left: convertInchesToTwip(0.5),
                        right: convertInchesToTwip(0.5),
                    },
                    spacing: {
                        line: 276,
                        before: 200,
                    },
                    children: [
                        new TextRun({
                            text: '2.2 SUMÁRIO DO CARGO ',
                            bold: true,
                            size: 24,
                            font: "Calibri Light",
                        }),
                        new TextRun({
                            text: `${user.skill.summary}`,
                            size: 24,
                            font: "Calibri Light",
                        }),
                       
                    ],
                }),
                new Paragraph({
                    alignment: AlignmentType.JUSTIFIED,
                    indent: {
                        left: convertInchesToTwip(0.5),
                        right: convertInchesToTwip(0.5),
                    },
                    spacing: {
                        line: 276,
                        before: 200,
                    },
                    children: [
                        new TextRun({
                            text: '2.3 ATRIBUIÇÕES ',
                            size: 24,
                            bold: true,
                            font: "Calibri Light",
                        }),                       
                    ],
                }),
            ]
        }],
    })
    
    createAttributionsSection(user.skill.assignments, doc),
    await createNewSection(user, doc);
    createSignSection(user, doc, company);
    createRenew(user, doc),
    createSignSection(user, doc, company);
    
    Packer.toBuffer(doc).then((buffer) => {
        fs.writeFileSync("ContratoExperiência.docx", buffer);
    });
})