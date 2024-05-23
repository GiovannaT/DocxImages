const docx = require("docx")
const fs = require('fs');

const { createXPContract, createClauseHeader, createSignSection, createRenew, createAttributionsSection, createNewSection, createCompanyTextRun, createEmployeeTextRun } = require('./contractUtils');

const caminhoParaJson = 'dbUser.json'

fs.readFile(caminhoParaJson, 'utf8', async (err, data) => {
    if(err){
        console.error('Erro ao ler arquivo Json' + err);
        return
    }

    let user = JSON.parse(data);

    // const getCompany = async () => {
    //     try {
    //         const response = await axios.get(`https://3337-avanciconstru-apiserver-0ae2jz7xl1m.ws-us110.gitpod.io/company?id=${company[0].company_id}`, {
    //             headers: {
    //             'Authorization': `Bearer eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpYXQiOjE3MTUyNjA5OTMsImV4cCI6MTc2NzEwMDk5Mywic3ViIjoiZjM1ZDg2M2QtMmI4My00MGM4LWI4ZDUtM2ExNzU5YTU2NTc2In0.EepQV0WVRYNQLQp2sPDhJU_Cm34c2rDFPuq0I_fqpDI`
    //         }});
    //         return response.data[0];

    //     } catch (error) {
    //         console.error('erro getcompany')
    //         return null;
    //     }
    // }

    const { Document, Packer, Paragraph, TextRun, ImageRun, Header, Footer, AlignmentType, BorderStyle, convertInchesToTwip} = docx;

    const doc = new Document({
        creator: "Usuário criador",
        description: `Contrato`,
        title: 'Contrato',
        sections: [{
            headers: {
                default: new Header({ 
                    children: [
                        new Paragraph({
                            alignment: AlignmentType.CENTER,    
                            children: [
                                new ImageRun({
                                    type: 'gif',
                                    data: fs.readFileSync("./public/hbt.png"),
                                    transformation: {
                                        width: 100,
                                        height: 100,
                                    },
                                })
                            ]
                        })
                    ],
                }),
            },
            footers: {
                default: new Footer({
                    children: [
                        new Paragraph({
                            alignment: AlignmentType.CENTER,    
                            children: [
                                new TextRun({
                                    text: "HABIT CONSTRUÇÕES E SERVIÇOS EIRELLI",
                                    bold: true,
                                    size: 10,
                                })
                            ],
                        }),
                        new Paragraph({
                            alignment: AlignmentType.CENTER,    
                            children: [
                                new TextRun({
                                    text: "CNPJ: 28.697.931/0001-43 - Rua Projetada F, nº 76 - Ribeirão do Lipa - Cuiabá - Mato Grosso",
                                    size: 10,
                                })
                            ],
                        }),
                        new Paragraph({
                            alignment: AlignmentType.CENTER,    
                            children: [
                                new TextRun({
                                    text: "CEP: 78048-163 - Telefone para contato: 65981088373 - recepcaohbt@gmail.com",
                                    size: 10,
                                })
                            ],
                        }),
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
                            font: "Arial",
                        }),
                        ...createCompanyTextRun(),
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
                            font: "Arial",
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
                            font: "Arial",
                        }),
                        new TextRun({
                            size: 24,
                            text: 'Fica o(a) ',
                            font: "Arial",
                        }),
                        new TextRun({
                            text: 'EMPREGADO(A) ',
                            size: 24,
                            bold: true,
                            font: "Arial",
                        }),
                        new TextRun({
                            text: 'admitido no quadro de funcionários do(a) ',
                            size: 24,
                            font: "Arial",
                        }),
                        new TextRun({
                            text: 'EMPREGADOR(A) ',
                            size: 24,
                            bold: true,
                            font: "Arial",
                        }),
                        new TextRun({
                            text: `para exercer o cargo de ${user.skill.name}, CBO ${user.skill.cbo}.`,
                            size: 24,
                            font: "Arial",
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
                            font: "Arial",
                        }),
                        new TextRun({
                            text: `${user.skill.summary}`,
                            size: 24,
                            font: "Arial",
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
                            font: "Arial",
                        }),                       
                    ],
                }),
            ]
        }],
    })
    
    createAttributionsSection(user.skill.assignments, doc),
    await createNewSection(user, doc);
    createSignSection(user, doc);
    createRenew(user, doc),
    createSignSection(user, doc);
    
    Packer.toBuffer(doc).then((buffer) => {
        fs.writeFileSync("ContratoExperiência.docx", buffer);
    });
})