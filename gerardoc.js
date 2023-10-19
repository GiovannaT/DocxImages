const axios = require('axios')
const docx = require("docx")
const fs = require('fs');

const caminhoParaJson = 'db.json'

fs.readFile(caminhoParaJson, 'utf8', async (err, data) => {
    if(err){
        console.error('Erro ao ler arquivo Json' + err);
        return
    }

    let lista = JSON.parse(data);

    const importFile = async (path) => {
        try {
            const response = await axios.get('https://avanci-bucket.s3.amazonaws.com/' + path, {responseType: 'arraybuffer'});
            return response.data;

        } catch (error) {
            console.log('erro ao buscar imagem')
            return null;
        }
    }

    const { Document, Packer, Paragraph, Footer, TextRun, HeadingLevel, PageBreak, ImageRun, AlignmentType, PageNumber, Table, TableRow, TableCell, VerticalAlign, WidthType} = docx;
    const doc = new Document({
        creator: "Usuário criador",
        description: `Relatório fotográfico`,
        title: 'Relatório Fotográfico',
        sections: [{
            footers: {
                default: new Footer({
                    children: [new Paragraph({
                        alignment: AlignmentType.END,
                        children: [
                            new TextRun("Página "),
                            new TextRun({
                                children: [PageNumber.CURRENT]
                            })
                        ]
                    })]
                })
            },
            children: [
                new Paragraph({
                    children: [
                        new TextRun({
                            text: "HABIT",
                            color: "#DD8400",
                        })
                    ],
                    heading: HeadingLevel.HEADING_1,
                }),
                new Paragraph({
                    border:{
                        top: {
                            color: "#0317fc",
                            space: 1,
                            style: 'single',
                            size: 555,
                        }
                    },
                    text: "Relatório Fotográfico",
                    heading: HeadingLevel.TITLE,
                    bold: true,
                }),
                new Paragraph({
                    children: [
                        new TextRun("CNPJ: 0000000"),
                    ],
                    alignment: AlignmentType.CENTER,
                    spacing: {
                        before: 200, 
                        after: 200
                    },
                }),
                new Paragraph({
                    alignment: AlignmentType.CENTER,
                    spacing: {
                        before: 200, 
                        after: 200
                    },
                    children: [
                        new TextRun("Endereço: aosjaosjas"),
                    ],
                }),
                new Paragraph({
                    alignment: AlignmentType.CENTER,
                    spacing: {
                        before: 200, 
                        after: 200
                    },
                    children: [
                        new TextRun("Email: 00000000"),
                    ],
                }),
                new Paragraph({
                    spacing: {
                        before: 1500, 
                    },
                    children: [
                        new TextRun("Secretaria de Estado de Saúde"),
                    ],
                }),
                new Paragraph({
                    children: [
                        new TextRun("Contrato: 202/109"),
                    ],
                }),
                new Paragraph({

                    children: [
                        new TextRun("Assunto: taltaltal"),
                        
                        new PageBreak(),
                    ],
                }),
            ]
        }]
    })
    
    lista.forEach(async object => {
        const files = object.files.slice(0,6);
        const imageRuns = [];
        
        for (const file of files) {
            const imageRun = new ImageRun({
                data: importFile(file.path),
                transformation: {
                    width: 200,
                    height: 200,
                },
            });
            imageRuns.push(imageRun);
        }

        const tableRows = [
            new TableRow({
                children: [
                    new TableCell({
                        borders:{
                            top: {
                                color: "000000",
                            },
                            bottom: {
                                color: "000000",
                            },
                            left: {
                                color: "000000",
                            },
                            right: {
                                color: "000000",
                            },
                        },
                        verticalAlign: VerticalAlign.CENTER,

                        children: [new Paragraph({ children: [imageRuns[0]] })],
                    }),
                    new TableCell({
                        borders:{
                            top: {
                                color: "000000",
                            },
                            bottom: {
                                color: "000000",
                            },
                            left: {
                                color: "000000",
                            },
                            right: {
                                color: "000000",
                            },
                        },
                        verticalAlign: VerticalAlign.CENTER,

                        children: [new Paragraph({ children: [imageRuns[1]] })],
                    }),
                ],
            }),
            new TableRow({
                height: {
                    value: 5000,
                },
                children: [
                    new TableCell({
                        borders:{
                            top: {
                                color: "000000",
                            },
                            bottom: {
                                color: "000000",
                            },
                            left: {
                                color: "000000",
                            },
                            right: {
                                color: "000000",
                            },
                        },
                        verticalAlign: VerticalAlign.CENTER,

                        children: [new Paragraph({ children: [imageRuns[2]] })],
                    }),
                    new TableCell({
                        borders:{
                            top: {
                                color: "000000",
                            },
                            bottom: {
                                color: "000000",
                            },
                            left: {
                                color: "000000",
                            },
                            right: {
                                color: "000000",
                            },
                        },
                        verticalAlign: VerticalAlign.CENTER,

                        children: [new Paragraph({ children: [imageRuns[3]] })],
                    }),
                ],
            }),
            new TableRow({
                height: {
                    value: 2000,
                },
                children: [
                    new TableCell({
                        borders:{
                            top: {
                                color: "000000",
                            },
                            bottom: {
                                color: "000000",
                            },
                            left: {
                                color: "000000",
                            },
                            right: {
                                color: "000000",
                            },
                        },
                        verticalAlign: VerticalAlign.CENTER,

                        children: [new Paragraph({ children: [imageRuns[4]] })],
                    }),
                    new TableCell({
                        borders:{
                            top: {
                                color: "000000",
                            },
                            bottom: {
                                color: "000000",
                            },
                            left: {
                                color: "000000",
                            },
                            right: {
                                color: "000000",
                            },
                        },
                        verticalAlign: VerticalAlign.CENTER,

                        children: [new Paragraph({ children: [imageRuns[5]] })],
                    }),
                ],
            }),
        ];
    
        const section = {
            children: [
                new Paragraph({
                    border:{
                        top: {
                            color: "#5763e6",
                            space: 1,
                            style: 'single',
                            size: 555,
                        }
                    },
                    text: `${object.place.name}`,
                    heading: HeadingLevel.HEADING_1,
                }),
                new Paragraph({
                    text: `Sala: ${object.department.name}`
                }),
                new Paragraph({
                    spacing: {after: 200},
                    text: `Serviço: ${object.service_description}`
                }),     
                new Table({
                    columnWidths: [8500, 8500],
                    width: {
                        size: 8640,
                        type: WidthType.DXA,
                    },
                    rows: tableRows,
                })

            ],
        }
        doc.addSection(section)
    })

    Packer.toBuffer(doc).then((buffer) => {
        fs.writeFileSync("Relatório Fotográfico.docx", buffer);
    });
})