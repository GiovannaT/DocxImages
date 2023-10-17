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

    const { Document, Packer, Paragraph, Footer, TextRun, HeadingLevel, PageBreak, ImageRun, AlignmentType, PageNumber, ExternalHyperlink} = docx;
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
                    text: "HABIT",
                    heading: HeadingLevel.HEADING_1,
                }),
                new Paragraph({
                    border:{
                        top: {
                            color: "#ffffff",
                            space: 1,
                            style: 'single',
                            size: 6,
                        }
                    },
                    text: "Relatório Fotográfico",
                    heading: HeadingLevel.TITLE,
                    bold: true,
                }),
                new Paragraph({
                    text: `CNPJ: 0000000`,
                    heading: HeadingLevel.HEADING_5,
                }),
                new Paragraph({
                    children: [
                        new TextRun("Endereço: aosjaosjas"),
                    ],
                }),
                new Paragraph({
                    children: [
                        new TextRun("Email: 00000000"),
                    ],
                }),
                new Paragraph({
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
        const files = object.files.slice(0,4);
        const imageRuns = [];
        files.forEach(file => {
            const imageRun = new ImageRun({                          
                data: importFile(file.path),                
                transformation: {
                    width: 250,
                    height: 250,
                },
            })
            imageRuns.push(imageRun)
        })
        const section = {
            children: [
                new Paragraph({
                    text: `${object.place.name}`,
                    heading: HeadingLevel.HEADING_1,
                }),
                new Paragraph({
                    text: `Sala: ${object.department.name}`
                }),
                new Paragraph({
                    text: `Serviço: ${object.service_description}`
                }),     
                new Paragraph({
                    children: imageRuns
                })          
            ],
        }
        doc.addSection(section)
    })

    Packer.toBuffer(doc).then((buffer) => {
        fs.writeFileSync("Relatório Fotográfico.docx", buffer);
    });
})