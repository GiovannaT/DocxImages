const axios = require('axios')
const moment = require("moment")
const { convertInchesToTwip, Paragraph, Table, TextRun, TableRow, TableCell, ShadingType, AlignmentType, WidthType, SectionType, BorderStyle} = require("docx");

function countDays(initialDate, finalDate){
    var startDate = new Date(initialDate);
    var endDate = new Date(finalDate);
    var timeDifference = endDate.getTime() - startDate.getTime();

    return daysDifference = Math.ceil(timeDifference / (1000 * 60 * 60 * 24));
}

function createClauseWithParagraphs(number, objective){
    return (
        new Paragraph({
            children: [
                new TextRun({text: `CLÁUSULA ${number}ª – ${objective} –`, bold: true}),
            ],
            spacing: {
                before: 200,
            },
        })   
    )
}

function createEpiProductsTable(productRows){
    if(productRows.length > 0){
        return (
            new Table({
                columnWidths: [3505, 5505],
                width: {
                    size: 9000,
                    type: WidthType.DXA,
                },
                rows: [
                    new TableRow({
                        cantSplit: true,
                        tableHeader: true,
                        children: [
                            new TableCell({
                                tableHeader: true,
                                height: {value: 50, type: WidthType},
                                borders: {
                                    bottom: {
                                        color: "ffffff",
                                    },
                                },
                                shading: {
                                    fill: "D9D9D9",
                                    type: ShadingType.CLEAR,
                                    color: "auto",
                                },
                                children: [
                                    new Paragraph({
                                    alignment: AlignmentType.CENTER,
                                    text: `QUANT.`,
                                    bold: true,
                                    }),
                                ],
                            }),
                            new TableCell({
                                tableHeader: true,
                                height: {value: 50, type: WidthType},
                                shading: {
                                    fill: "D9D9D9",
                                    type: ShadingType.CLEAR,
                                    color: "auto",
                                },
                                borders: {
                                    bottom: {
                                        color: "ffffff",
                                    },
                                },
                                children: [new Paragraph({
                                    alignment: AlignmentType.CENTER,
                                    text: `TIPO DE EPI`,
                                    bold: true,
                                }),],
                            }),
                            new TableCell({
                                tableHeader: true,
                                height: {value: 50, type: WidthType},
                                shading: {
                                    fill: "D9D9D9",
                                    type: ShadingType.CLEAR,
                                    color: "auto",
                                },
                                borders: {
                                    bottom: {
                                        color: "ffffff",
                                    },
                                },
                                children: [new Paragraph({
                                    alignment: AlignmentType.CENTER,
                                    text: `Nº C.A.`,
                                    bold: true,
                                }),],
                            }),
                            new TableCell({
                                tableHeader: true,
                                height: {value: 50, type: WidthType},
                                shading: {
                                    fill: "D9D9D9",
                                    type: ShadingType.CLEAR,
                                    color: "auto",
                                },
                                borders: {
                                    bottom: {
                                        color: "ffffff",
                                    },
                                },
                                children: [new Paragraph({
                                    alignment: AlignmentType.CENTER,
                                    text: `Data entrega`,
                                    bold: true,
                                }),],
                            }),
                            new TableCell({
                                tableHeader: true,
                                height: {value: 50, type: WidthType},
                                shading: {
                                    fill: "D9D9D9",
                                    type: ShadingType.CLEAR,
                                    color: "auto",
                                },
                                borders: {
                                    bottom: {
                                        color: "ffffff",
                                    },
                                },
                                children: [new Paragraph({
                                    alignment: AlignmentType.CENTER,
                                    text: `Data devolução`,
                                    bold: true,
                                }),],
                            }),
                            new TableCell({
                                tableHeader: true,
                                height: {value: 50, type: WidthType},
                                shading: {
                                    fill: "D9D9D9",
                                    type: ShadingType.CLEAR,
                                    color: "auto",
                                },
                                children: [new Paragraph({
                                    alignment: AlignmentType.CENTER,
                                    text: `Motivo`,
                                    bold: true,
                                }),],
                            }),
                            new TableCell({
                                tableHeader: true,
                                height: {value: 50, type: WidthType},
                                shading: {
                                    fill: "D9D9D9",
                                    type: ShadingType.CLEAR,
                                    color: "auto",
                                },
                                columnSpan: 2,
                                borders: {
                                    bottom: {
                                        color: "ffffff",
                                    },
                                },
                                children: [new Paragraph({
                                    alignment: AlignmentType.CENTER,
                                    text: `Nº C.A. Novo Equipamento`,
                                    bold: true,
                                }),],
                            }),
                        ],
                    }),
                    new TableRow({
                        children: [
                            new TableCell({
                                height: {value: 50, type: WidthType},
                                shading: {
                                    fill: "D9D9D9",
                                    type: ShadingType.CLEAR,
                                    color: "auto",
                                },
                                borders: {
                                    top: {
                                        color: "ffffff",
                                    },
                                    bottom: {
                                        color: "ffffff",
                                    },
                                },
                                children: [],
                            }),
                            new TableCell({
                                height: {value: 50, type: WidthType},
                                shading: {
                                    fill: "D9D9D9",
                                    type: ShadingType.CLEAR,
                                    color: "auto",
                                },
                                borders: {
                                    top: {
                                        color: "ffffff",
                                    },
                                    bottom: {
                                        color: "ffffff",
                                    },
                                },
                                children: [],
                            }),
                            new TableCell({
                                height: {value: 50, type: WidthType},
                                shading: {
                                    fill: "D9D9D9",
                                    type: ShadingType.CLEAR,
                                    color: "auto",
                                },
                                borders: {
                                    top: {
                                        color: "ffffff",
                                    },
                                    bottom: {
                                        color: "ffffff",
                                    },
                                },
                                children: [],
                            }),
                            new TableCell({
                                height: {value: 50, type: WidthType},
                                shading: {
                                    fill: "D9D9D9",
                                    type: ShadingType.CLEAR,
                                    color: "auto",
                                },
                                borders: {
                                    top: {
                                        color: "ffffff",
                                    },
                                    bottom: {
                                        color: "ffffff",
                                    },
                                },
                                children: [],
                            }),
                            new TableCell({
                                height: {value: 50, type: WidthType},
                                shading: {
                                    fill: "D9D9D9",
                                    type: ShadingType.CLEAR,
                                    color: "auto",
                                },
                                borders: {
                                    top: {
                                        color: "ffffff",
                                    },
                                    bottom: {
                                        color: "ffffff",
                                    },
                                },
                                children: [],
                            }),
                            new TableCell({
                                height: {value: 50, type: WidthType},
                                shading: {
                                    fill: "D9D9D9",
                                    type: ShadingType.CLEAR,
                                    color: "auto",
                                },
                                children: [new Paragraph({
                                    alignment: AlignmentType.CENTER,
                                    text: `Substituição`,
                                    bold: true,
                                }),],
                            }),
                            new TableCell({
                                height: {value: 50, type: WidthType},
                                shading: {
                                    fill: "D9D9D9",
                                    type: ShadingType.CLEAR,
                                    color: "auto",
                                },
                                borders: {
                                    top: {
                                        color: "ffffff",
                                    },
                                    bottom: {
                                        color: "ffffff",
                                    },
                                },
                                children: [],
                            }),
                        ]
                    }),
                    ...productRows,
                    new TableRow({
                        children: [
                            new TableCell({
                                height: {value: 50, type: WidthType},
                                children: [],
                            }),
                            new TableCell({
                                height: {value: 50, type: WidthType},
                                children: [new Paragraph({
                                    alignment: AlignmentType.CENTER,
                                    text: `O: Outro(descrever motivo)`,
                                    bold: true,
                                }),],
                            }),
                            new TableCell({
                                height: {value: 50, type: WidthType},
                                children: [
                                    new Paragraph({
                                    alignment: AlignmentType.CENTER,
                                    text: `I: Inadequado`,
                                    bold: true,
                                    }),
                                ],
                            }),
                            new TableCell({
                                height: {value: 50, type: WidthType},
                                children: [
                                    new Paragraph({
                                    alignment: AlignmentType.CENTER,
                                    text: `T: Tempo de Uso`,
                                    bold: true,
                                    }),
                                ],
                            }),
                            new TableCell({
                                height: {value: 50, type: WidthType},
                                children: [
                                    new Paragraph({
                                    alignment: AlignmentType.CENTER,
                                    text: `D: Defeito`,
                                    bold: true,
                                    }),
                                ],
                            }),
                            new TableCell({
                                height: {value: 50, type: WidthType},
                                children: [],
                            }),
                            new TableCell({
                                height: {value: 50, type: WidthType},
                                children: [],
                            }),
                        ]
                    }),
                ],
            })
        )
    }
}

async function getSkillProducts(id){
    try {
        const response = await axios.get(`https://3337-avanciconstru-apiserver-i1jsgtd4yo0.ws-us110.gitpod.io/skillproducts?skill_id=${id}`, {
            headers: {
            'Authorization': `Bearer eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpYXQiOjE3MTU5NjYwMDgsImV4cCI6MTc2NzgwNjAwOCwic3ViIjoiZjM1ZDg2M2QtMmI4My00MGM4LWI4ZDUtM2ExNzU5YTU2NTc2In0.eK88F638zrMQ7x6WDgrX6eAOBE4M4jklTlTZVTuOlOk`
        }});
        return response.data;
    } catch (error) {
        console.error('erro no get skill product')
        return null;
    }
}

async function createEpiTable(doc, executor){
    if(executor.executor){
        var tableRows = [];
        var productRows = [];
        try{
            const skillProducts = await getSkillProducts(executor.executor.skill_id);
            if(skillProducts && skillProducts.length > 0){
                skillProducts.forEach(product => {
                    productRows.push(
                        new TableRow({
                            children: [
                                new TableCell({
                                    height: {value: 50, type: WidthType},
                                    children: [
                                        new Paragraph({
                                            alignment: AlignmentType.CENTER,
                                            text: `1`,
                                            bold: true,
                                            }),
                                    ],
                                }),
                                new TableCell({
                                    height: {value: 50, type: WidthType},
                                    children: [
                                        new Paragraph({
                                            alignment: AlignmentType.CENTER,
                                            text: `${product?.product?.description}.`,
                                            bold: true,
                                            }),
                                    ],
                                }),
                                new TableCell({
                                    height: {value: 50, type: WidthType},
                                    children: [],
                                }),
                                new TableCell({
                                    height: {value: 50, type: WidthType},
                                    children: [
                                        new Paragraph({
                                            alignment: AlignmentType.CENTER,
                                            text: `${moment().format('DD/MM/YYYY')}.`,
                                            bold: true,
                                        }),
                                    ],
                                }),
                                new TableCell({
                                    height: {value: 50, type: WidthType},
                                    children: [],
                                }),
                                new TableCell({
                                    height: {value: 50, type: WidthType},
                                    children: [],
                                }),
                                new TableCell({
                                    height: {value: 50, type: WidthType},
                                    children: [],
                                }),
                            ]
                        }),
                    )
                })
            }    
        } catch (error){
            console.log(error);
        }
        tableRows = [
            new TableRow({
                tableHeader: true,
                cantSplit:true,
                children: [
                    new TableCell({
                        height: {value: 50, type: WidthType},
                        width: {
                            size: 8000,
                            type: WidthType.DXA,
                        },
                        shading: {
                            fill: "D9D9D9",
                            type: ShadingType.CLEAR,
                            color: "auto",
                        },
                        columnSpan: 2,
                        children: [new Paragraph({
                            alignment: AlignmentType.CENTER,
                            // pageBreakBefore: true,
                            text: `FICHA DE CONTROLE DE EPI'S`,
                            bold: true,
                        }),],
                    }),
                ],
            }),
            new TableRow({
                tableHeader: true,
                height: {value: 50, type: WidthType},
                width: {
                    size: 8000,
                    type: WidthType.DXA,
                },
                children: [
                    new TableCell({
                        height: {value: 50, type: WidthType},
                        width: {
                            size: 8000,
                            type: WidthType.DXA,
                        },
                        children: [
                            new Paragraph(`NOME: ${executor.executor.name}`)
                        ],
                    }),
                    new TableCell({
                        height: {value: 50, type: WidthType},
                        width: {
                            size: 8000,
                            type: WidthType.DXA,
                        },
                        children: [
                            new Paragraph(`FUNÇÃO: ${executor.executor.skill.name}`)
                        ],
                    }),
                ],
            }),
            new TableRow({
                tableHeader: true,
                height: {value: 50, type: WidthType},
                width: {
                    size: 8000,
                    type: WidthType.DXA,
                },
                children: [
                    new TableCell({
                        height: {value: 50, type: WidthType},
                        width: {
                            size: 8000,
                            type: WidthType.DXA,
                        },
                        shading: {
                            fill: "D9D9D9",
                            type: ShadingType.CLEAR,
                            color: "auto",
                        },
                        columnSpan: 2,
                        children: [new Paragraph({
                            alignment: AlignmentType.CENTER,
                            text: `TERMO DE RESPONSABILIDADE`,
                            bold: true,
                        }),],
                    }),
                ],
            }),
            new TableRow({
                height: {value: 50, type: WidthType},
                width: {
                    size: 8000,
                    type: WidthType.DXA,
                },
                children: [
                    new TableCell({
                        borders: {
                            bottom: {
                                color: "ffffff",
                            },
                        },
                        height: {value: 50, type: WidthType},
                        width: {
                            size: 8000,
                            type: WidthType.DXA,
                        },
                        columnSpan: 2,
                        children: [
                            new Paragraph({
                            text: `Recebi de HABIT CONSTRUÇÕES E SERVIÇOS EIRELI , para meu uso obrigatório os EPI's (Equipamentos de proteção Individual) e EPC's (Equipamentos de2 proteção coletiva) constantes nesta ficha, o qual obrigo-me a utiliza-los corretamente durante o tempo que permanecerem ao meu dispor, observando as medidas gerais de disciplina e uso que integram a NR-06 - Equipamento de Proteção Individual - EPI's - da portaria n.º 3.214 de 08/06/1978.`,
                            }),
                            new Paragraph(
                                {
                                text: "1. Usar o EPI Indicado apenas para as finalidades a que se destina.",
                                indent: {
                                    left: 300,
                                },
                            }),
                            new Paragraph({
                                text: "2. Somente iniciar o serviço se estiver usando os EPI's indicados na tarefa a realizar.",
                                indent: {
                                    left: 300,
                                },
                            }),
                            new Paragraph({
                                text: "3. Responsabilizar-se pela guarda e conservação dos EPI's.",
                                indent: {
                                    left: 300,
                                },
                            }),
                            new Paragraph({
                                text: "4. Comunicar qualquer dano ou extravio no EPI, para aquisição de outro.",
                                indent: {
                                    left: 300,
                                },
                            }),
                            new Paragraph({
                                text: "5. Responder perante a empresa pelo custo integral ao preço de mercado do dia, quando:",
                                indent: {
                                    left: 300,
                                },
                            }),
                            new Paragraph({
                                text: "a)  Alegar Perda ou Extravio",
                                indent: {
                                    left: 500,
                                },
                            }),
                            new Paragraph({
                                text: "b)  Alterar seu padrão",
                                indent: {
                                    left: 500,
                                },
                            }),
                            new Paragraph({
                                text: "c)  Inutilizá-lo por procedimento inadequado",
                                indent: {
                                    left: 500,
                                },
                            }),
                            new Paragraph({
                                text: "d)  Desligar-se da empresa sem devolver o EPI.",
                                indent: {
                                    left: 500,
                                },
                            }),
                            new Paragraph({
                                text: "6. A recusa em não usar os EPI's, gerará punição em lei (CLT art. 482)",
                                indent: {
                                    left: 300,
                                },
                            }),
                        ],
                    }),
                ],
            }),
            new TableRow({
                children:[
                    new TableCell({
                        columnSpan: 2,
                        borders: {
                            top: {
                                color: "ffffff",
                            },
                            bottom: {
                                color: "ffffff",
                            },
                        },
                        children:[
                            new Paragraph({
                            spacing:{
                                before: 200,
                                after: 200,
                            },
                            children: [
                                new TextRun({
                                    text: "Declaro haver recebido treinamento sobre o uso dos EPI's e estar de pleno acordo com as normas dos equipamentos de proteção individual, acima estipulado.",
                                    bold: true,
                                    })
                            ]
                            }),
                        ]
                    }
                    ),
                ]
            }),
            new TableRow({
                children:[
                    new TableCell({
                        borders: {
                            top: {
                                color: "ffffff",
                            },
                            bottom: {
                                color: "ffffff",
                            },
                        },
                        columnSpan: 2,
                        children:[
                            new Table({
                                columnWidths: [3505, 5505],
                                width: {
                                    size: 9000,
                                    type: WidthType.DXA,
                                },
                                rows: [
                                    new TableRow({
                                        cantSplit: true,
                                        tableHeader: true,
                                        children: [
                                            new TableCell({
                                                tableHeader: true,
                                                height: {value: 50, type: WidthType},
                                                borders: {
                                                    bottom: {
                                                        color: "ffffff",
                                                    },
                                                },
                                                shading: {
                                                    fill: "D9D9D9",
                                                    type: ShadingType.CLEAR,
                                                    color: "auto",
                                                },
                                                children: [
                                                    new Paragraph({
                                                    alignment: AlignmentType.CENTER,
                                                    text: `QUANT.`,
                                                    bold: true,
                                                    }),
                                                ],
                                            }),
                                            new TableCell({
                                                tableHeader: true,
                                                height: {value: 50, type: WidthType},
                                                shading: {
                                                    fill: "D9D9D9",
                                                    type: ShadingType.CLEAR,
                                                    color: "auto",
                                                },
                                                borders: {
                                                    bottom: {
                                                        color: "ffffff",
                                                    },
                                                },
                                                children: [new Paragraph({
                                                    alignment: AlignmentType.CENTER,
                                                    text: `TIPO DE EPI`,
                                                    bold: true,
                                                }),],
                                            }),
                                            new TableCell({
                                                tableHeader: true,
                                                height: {value: 50, type: WidthType},
                                                shading: {
                                                    fill: "D9D9D9",
                                                    type: ShadingType.CLEAR,
                                                    color: "auto",
                                                },
                                                borders: {
                                                    bottom: {
                                                        color: "ffffff",
                                                    },
                                                },
                                                children: [new Paragraph({
                                                    alignment: AlignmentType.CENTER,
                                                    text: `Nº C.A.`,
                                                    bold: true,
                                                }),],
                                            }),
                                            new TableCell({
                                                tableHeader: true,
                                                height: {value: 50, type: WidthType},
                                                shading: {
                                                    fill: "D9D9D9",
                                                    type: ShadingType.CLEAR,
                                                    color: "auto",
                                                },
                                                borders: {
                                                    bottom: {
                                                        color: "ffffff",
                                                    },
                                                },
                                                children: [new Paragraph({
                                                    alignment: AlignmentType.CENTER,
                                                    text: `Data entrega`,
                                                    bold: true,
                                                }),],
                                            }),
                                            new TableCell({
                                                tableHeader: true,
                                                height: {value: 50, type: WidthType},
                                                shading: {
                                                    fill: "D9D9D9",
                                                    type: ShadingType.CLEAR,
                                                    color: "auto",
                                                },
                                                borders: {
                                                    bottom: {
                                                        color: "ffffff",
                                                    },
                                                },
                                                children: [new Paragraph({
                                                    alignment: AlignmentType.CENTER,
                                                    text: `Data devolução`,
                                                    bold: true,
                                                }),],
                                            }),
                                            new TableCell({
                                                tableHeader: true,
                                                height: {value: 50, type: WidthType},
                                                shading: {
                                                    fill: "D9D9D9",
                                                    type: ShadingType.CLEAR,
                                                    color: "auto",
                                                },
                                                children: [new Paragraph({
                                                    alignment: AlignmentType.CENTER,
                                                    text: `Motivo`,
                                                    bold: true,
                                                }),],
                                            }),
                                            new TableCell({
                                                tableHeader: true,
                                                height: {value: 50, type: WidthType},
                                                shading: {
                                                    fill: "D9D9D9",
                                                    type: ShadingType.CLEAR,
                                                    color: "auto",
                                                },
                                                columnSpan: 2,
                                                borders: {
                                                    bottom: {
                                                        color: "ffffff",
                                                    },
                                                },
                                                children: [new Paragraph({
                                                    alignment: AlignmentType.CENTER,
                                                    text: `Nº C.A. Novo Equipamento`,
                                                    bold: true,
                                                }),],
                                            }),
                                        ],
                                    }),
                                    new TableRow({
                                        children: [
                                            new TableCell({
                                                height: {value: 50, type: WidthType},
                                                shading: {
                                                    fill: "D9D9D9",
                                                    type: ShadingType.CLEAR,
                                                    color: "auto",
                                                },
                                                borders: {
                                                    top: {
                                                        color: "ffffff",
                                                    },
                                                    bottom: {
                                                        color: "ffffff",
                                                    },
                                                },
                                                children: [],
                                            }),
                                            new TableCell({
                                                height: {value: 50, type: WidthType},
                                                shading: {
                                                    fill: "D9D9D9",
                                                    type: ShadingType.CLEAR,
                                                    color: "auto",
                                                },
                                                borders: {
                                                    top: {
                                                        color: "ffffff",
                                                    },
                                                    bottom: {
                                                        color: "ffffff",
                                                    },
                                                },
                                                children: [],
                                            }),
                                            new TableCell({
                                                height: {value: 50, type: WidthType},
                                                shading: {
                                                    fill: "D9D9D9",
                                                    type: ShadingType.CLEAR,
                                                    color: "auto",
                                                },
                                                borders: {
                                                    top: {
                                                        color: "ffffff",
                                                    },
                                                    bottom: {
                                                        color: "ffffff",
                                                    },
                                                },
                                                children: [],
                                            }),
                                            new TableCell({
                                                height: {value: 50, type: WidthType},
                                                shading: {
                                                    fill: "D9D9D9",
                                                    type: ShadingType.CLEAR,
                                                    color: "auto",
                                                },
                                                borders: {
                                                    top: {
                                                        color: "ffffff",
                                                    },
                                                    bottom: {
                                                        color: "ffffff",
                                                    },
                                                },
                                                children: [],
                                            }),
                                            new TableCell({
                                                height: {value: 50, type: WidthType},
                                                shading: {
                                                    fill: "D9D9D9",
                                                    type: ShadingType.CLEAR,
                                                    color: "auto",
                                                },
                                                borders: {
                                                    top: {
                                                        color: "ffffff",
                                                    },
                                                    bottom: {
                                                        color: "ffffff",
                                                    },
                                                },
                                                children: [],
                                            }),
                                            new TableCell({
                                                height: {value: 50, type: WidthType},
                                                shading: {
                                                    fill: "D9D9D9",
                                                    type: ShadingType.CLEAR,
                                                    color: "auto",
                                                },
                                                children: [new Paragraph({
                                                    alignment: AlignmentType.CENTER,
                                                    text: `Substituição`,
                                                    bold: true,
                                                }),],
                                            }),
                                            new TableCell({
                                                height: {value: 50, type: WidthType},
                                                shading: {
                                                    fill: "D9D9D9",
                                                    type: ShadingType.CLEAR,
                                                    color: "auto",
                                                },
                                                borders: {
                                                    top: {
                                                        color: "ffffff",
                                                    },
                                                    bottom: {
                                                        color: "ffffff",
                                                    },
                                                },
                                                children: [],
                                            }),
                                        ]
                                    }),
                                    ...productRows,
                                    new TableRow({
                                        children: [
                                            new TableCell({
                                                height: {value: 50, type: WidthType},
                                                children: [],
                                            }),
                                            new TableCell({
                                                height: {value: 50, type: WidthType},
                                                children: [new Paragraph({
                                                    alignment: AlignmentType.CENTER,
                                                    text: `O: Outro(descrever motivo)`,
                                                    bold: true,
                                                }),],
                                            }),
                                            new TableCell({
                                                height: {value: 50, type: WidthType},
                                                children: [
                                                    new Paragraph({
                                                    alignment: AlignmentType.CENTER,
                                                    text: `I: Inadequado`,
                                                    bold: true,
                                                    }),
                                                ],
                                            }),
                                            new TableCell({
                                                height: {value: 50, type: WidthType},
                                                children: [
                                                    new Paragraph({
                                                    alignment: AlignmentType.CENTER,
                                                    text: `T: Tempo de Uso`,
                                                    bold: true,
                                                    }),
                                                ],
                                            }),
                                            new TableCell({
                                                height: {value: 50, type: WidthType},
                                                children: [
                                                    new Paragraph({
                                                    alignment: AlignmentType.CENTER,
                                                    text: `D: Defeito`,
                                                    bold: true,
                                                    }),
                                                ],
                                            }),
                                            new TableCell({
                                                height: {value: 50, type: WidthType},
                                                children: [],
                                            }),
                                            new TableCell({
                                                height: {value: 50, type: WidthType},
                                                children: [],
                                            }),
                                        ]
                                    }),
                                ],
                            })
                        ]
                    }),
                ]
            }),
            new TableRow({
                children:[
                    new TableCell(
                        {
                        borders: {
                            top: {
                                color: "ffffff",
                            },
                        },
                        columnSpan: 2,
                        children: [
                            new Paragraph(
                                {
                                    spacing:{
                                        before: 200,
                                        after: 200,
                                    },
                                    text: `Colaborador declara que possui todos os EPIs em ótimo estado para uso de acordo com sua função, _______`},
                            ),
                        ]
                        }
                    )
                ]
            })
        ]
        const section = {
            children: [
                new Table({
                    columnWidths: [8500, 8500],
                    width: {
                        size: 9000,
                        type: WidthType.DXA,
                    },
                    rows: tableRows,
                })
            ],
        }
        doc.addSection(section);
    }
}

function createXPContract(){
    return (
        new Paragraph({
            children: [
                new TextRun({
                    text: "CONTRATO DE EXPERIÊNCIA",
                    bold: true,
                    size: 36,
                    font: "Arial",
                })
            ],
            alignment: AlignmentType.CENTER,
        })
    )
}

function createClauseHeader(text){
    return (
        new Paragraph({
            children: [
                new TextRun({
                    text: text,
                    bold: true,
                    size: 30,
                    font: "Arial",
                })
            ],
            indent: {
                left: convertInchesToTwip(0.5),
                right: convertInchesToTwip(0.5),
            },
            spacing: {
                line: 276,
            },
            spacing: {
                before: 300,
            },
        })
    )
}

function createSignSection(doc){
    const section = {
        properties: {
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
            type: SectionType.CONTINUOUS,
            column: {
                space: 708,
                count: 2,
            },
        },
        children: [
            new Paragraph({
                spacing: {
                    before: 400,
                },
                alignment: AlignmentType.CENTER,
                children: [
                    new TextRun({
                        size: 24,
                        text: `EMPREGADOR(A):`,
                        bold: true,
                        font: "Arial",
                    }),
                ],
            }),
            new Paragraph({
                border:{
                    top: {
                        color: "#000000",
                        space: 1,
                        style: 'single',
                        size: 1,
                    }
                },
                children: [
                    new TextRun({text:`HABIT CONSTRUÇÕES E SERVIÇOS LTDA`, 
                    bold: true,
                    font: "Arial",}),
                ],
                alignment: AlignmentType.CENTER,
                spacing: {
                    before: 400,
                },
            }),
            new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [
                    new TextRun({
                        size: 16,
                        text: `CNPJ nº: 28.697.934/0001-43`,
                        font: "Arial",
                    }),
                ],
            }),
            new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: {
                    before: 500,
                },
                children: [
                    new TextRun({
                        size: 24,
                        text: `EMPREGADO(A):`,
                        bold: true,
                        font: "Arial",
                    }),
                ],
            }),
            new Paragraph({
                border:{
                    top: {
                        color: "#000000",
                        space: 1,
                        style: 'single',
                        size: 1,
                    }
                },
                children: [
                    new TextRun({text:`NOME DO FULANO`, 
                    bold: true,
                    font: "Arial",}),
                ],
                alignment: AlignmentType.CENTER,
                spacing: {
                    before: 500,
                },
            }),
            new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [
                    new TextRun({
                        size: 16,
                        text: `CPF cpf do fulano`,
                        font: "Arial",
                    }),
                ],
            }),
            new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: {
                    // before: 12000,
                    before: 300,
                },
                children: [
                    new TextRun({
                        size: 24,
                        text: `TESTEMUNHAS:`,
                        bold: true,
                        font: "Arial",
                    }),
                ],
            }),
            new Paragraph({
                border:{
                    top: {
                        color: "#000000",
                        space: 1,
                        style: 'single',
                        size: 1,
                    }
                },
                children: [
                    new TextRun({text:`LARISSA K. RAMOS MAGALHAES`, 
                    bold: true,
                    font: "Arial",}),
                ],
                alignment: AlignmentType.CENTER,
                spacing: {
                    before: 400,
                },
            }),
            new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [
                    new TextRun({
                        size: 16,
                        text: `CPF: 049.391.961-94`,
                        font: "Arial",
                    }),
                ],
            }),
            new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [
                    new TextRun({
                        size: 16,
                        text: `SETOR DE RH`,
                        font: "Arial",
                    }),
                ],
            }),
            new Paragraph({
                border:{
                    top: {
                        color: "#000000",
                        space: 1,
                        style: 'single',
                        size: 1,
                    }
                },
                children: [
                    new TextRun({text:`NOME DO GERENTE`, 
                    bold: true,
                    font: "Arial",}),
                ],
                alignment: AlignmentType.CENTER,
                spacing: {
                    before: 500,
                },
            }),
            new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [
                    new TextRun({
                        size: 16,
                        text: `CPF cpf do gerente`,
                        font: "Arial",
                    }),
                ],
            }),
            new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [
                    new TextRun({
                        size: 16,
                        text: `ENGENHEIRO GERENTE`,
                        font: "Arial",
                    }),
                ],
            }),
            new Paragraph({
                border:{
                    top: {
                        color: "#000000",
                        space: 1,
                        style: 'single',
                        size: 1,
                    }
                },
                children: [
                    new TextRun({text:`NOME DO GESTOR`, 
                    bold: true,
                    font: "Arial",}),
                ],
                alignment: AlignmentType.CENTER,
                spacing: {
                    before: 500,
                },
            }),
            new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [
                    new TextRun({
                        size: 16,
                        text: `CPF cpf do GESTOR`,
                        font: "Arial",
                    }),
                ],
            }),
            new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [
                    new TextRun({
                        size: 16,
                        text: `GESTOR RESPONSÁVEL`,
                        font: "Arial",
                    }),
                ],
            }),
        ],
    }
    doc.addSection(section);
}

function createRenew(doc){
    const section = {
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
            new Paragraph({
                indent: {
                    left: convertInchesToTwip(0.5),
                    right: convertInchesToTwip(0.5),
                },
                spacing: {
                    line: 276,
                },
                children: [
                    new TextRun({
                        text: "TERMO DE RENOVAÇÃO",
                        bold: true,
                        size: 36,
                        font: "Arial",
                    })
                ],
                alignment: AlignmentType.CENTER,
            }),
            new Paragraph({
                indent: {
                    left: convertInchesToTwip(0.5),
                    right: convertInchesToTwip(0.5),
                },
                spacing: {
                    line: 276,
                    before: 100,
                },
                children: [
                    new TextRun({
                        text: 'Pelo presente TERMO ADITIVO ao contrato de trabalho entabulado entre: ',
                        size: 24,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: `EMPREGADOR(A): NOME DA EMPRESA, `,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        text: `inscrito no CNPJ nº: CNPJ DA EMPRESA, email EMAIL DA EMPRESA, telefone para contato TELEFONE DA EMPRESA, com sede em ENDEREÇO DA EMPRESA`,
                        size: 24,
                        font: "Arial",
                    }),
                ],
            }),
            new Paragraph({
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
                        text: `EMPREGADO(A): NOME COMPLETO EMPREGADO`,
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: `CPF, RG, data de emissão EMISSAO, data de nascimento NASCIMENTO, telefone para contato TELEFONE, estado civil, declaração étnica, residente em ENDEREÇO, `,
                        font: "Arial",
                    }),
                ],
            }),
            new Paragraph({
                indent: {
                    left: convertInchesToTwip(0.5),
                    right: convertInchesToTwip(0.5),
                },
                spacing: {
                    line: 276,
                    before: 300,
                },
                children: [
                    new TextRun({
                        size: 24,
                        text: `Resolvem as partes pela prorrogação do CONTRATO DE EXPERIÊNCIA pelo período de 45 dias`,
                        font: "Arial",
                    }),
                ],
            }),
            new Paragraph({
                indent: {
                    left: convertInchesToTwip(0.5),
                    right: convertInchesToTwip(0.5),
                },
                spacing: {
                    line: 276,
                    before: 300,
                },
                children: [
                    new TextRun({
                        size: 24,
                        text: `As demais cláusulas permanecem inalteradas.`,
                        font: "Arial",
                    }),
                ],
            }),
            new Paragraph({
                indent: {
                    left: convertInchesToTwip(0.5),
                    right: convertInchesToTwip(0.5),
                },
                spacing: {
                    line: 276,
                    before: 300,
                },
                children: [
                    new TextRun({
                        size: 24,
                        text: `E por estarem em pleno acordo, as partes contratantes assinam o presente TERMO DE RENOVAÇÃO DO CONTRATO DE TRABALHO, em duas vias, ficando a primeira em poder do `,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: `EMPREGADOR(A), `,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: `e a segunda com o(a) `,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: `EMPREGADO(A). `,
                        bold: true,
                        font: "Arial",
                    }),
                ],
            }),
            new Paragraph({
                indent: {
                    left: convertInchesToTwip(0.5),
                    right: convertInchesToTwip(0.5),
                },
                spacing: {
                    line: 276,
                    before: 300,
                },
                children: [
                    new TextRun({
                        size: 24,
                        text: `{{NOME DA CIDADE}}, {{DIA DO MÊS}} de {{MÊS DO ANO}} de {{ANO}}.`,
                        font: "Arial",
                    }),
                ],
            }),
        ],
    }
    doc.addSection(section);
}

module.exports = {createClauseWithParagraphs, createEpiProductsTable, getSkillProducts, createEpiTable, countDays, createXPContract, createClauseHeader, createSignSection, createRenew}

