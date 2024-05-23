const axios = require('axios')
const moment = require("moment")
const extenso = require('extenso')

const { convertInchesToTwip, Paragraph, Table, TextRun, TableRow, TableCell, ShadingType, AlignmentType, WidthType, SectionType, BorderStyle, Column} = require("docx");

moment.locale('pt'); 
moment.updateLocale('pt', {
    months : [
        "janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho",
        "agosto", "setembro", "outubro", "novembro", "dezembro"
    ]
});

function capitalizeString(string) {
    const slice = string.toLowerCase().trim().split(' ');
    for (let i = 0; i < slice.length; i++) {
        slice[i] = slice[i].charAt(0).toUpperCase() + slice[i].slice(1);
    }
    return slice.join(' ');
}

function countDays(initialDate, finalDate){
    var startDate = new Date(initialDate);
    var endDate = new Date(finalDate);
    var timeDifference = endDate.getTime() - startDate.getTime();

    return daysDifference = Math.ceil(timeDifference / (1000 * 60 * 60 * 24));
}

function createClauseWithParagraphs(number, objective){
    return (
        new Paragraph({
                    alignment: AlignmentType.JUSTIFIED,

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

async function getShiftTurn(id){
    try {
        const response = await axios.get(`https://3337-avanciconstru-apiserver-i1jsgtd4yo0.ws-us114.gitpod.io/timeclock/turno?id=${id}`, {
            headers: {
            'Authorization': `Bearer eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpYXQiOjE3MTYzODAwMTAsImV4cCI6MTc2ODIyMDAxMCwic3ViIjoiZjM1ZDg2M2QtMmI4My00MGM4LWI4ZDUtM2ExNzU5YTU2NTc2In0.JCxRXmEFbmjV-G9SZ9sdmMYXeFeDlZdyGY11b03ERWA`
        }});
        console.log(response.data)
        return response.data;
    } catch (error) {
        console.error('erro no get shift')
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

function createCompanyTextRun(){
    return [
        new TextRun({
            text: 'EMPREGADOR(A): HABIT CONSTRUÇÕES E SERVIÇOS LTDA, ',
            size: 24,
            bold: true,
            font: "Arial",
        }),
        new TextRun({
            text: 'inscrito no CNPJ nº: 28.697.934/0001- 43, e-mail recepcaohbt@gmail.com, telefone para contato 6598108-8373, com sede em rua Projetada F, Nº 76, Ribeirão do Lipa, Cuiabá – MT, CEP: 78048-163.',
            size: 24,
            font: "Arial",
        })
    ]
}

function createEmployeeTextRun(user){
    return [
        new TextRun({
            text: `EMPREGADO(A): ${user.name} `,
            size: 24,
            bold: true,
            font: "Arial",
        }),
        new TextRun({
            size: 24,
            text: `CPF: ${user.cpf}, RG: ${user.rg}, data de emissão ${moment(user.meta.issuedOn).format('DD/MM/YYYY')}, data de nascimento: ${moment(user.birthday).format('DD/MM/YYYY')}, telefone para contato ${user.phone}, ${user.meta.maritalStatus}, residente em ${capitalizeString(user.address.address)}, nº ${user.address.number}, complemento: ${user.address.complement ? capitalizeString(user.address.complement) : 'Não informado'}, ${capitalizeString(user.address.neighborhood)}, ${user.address.city.name}, ${user.address.state.name}, CEP: ${user.address.postal_code}.`,
            font: "Arial",
        }),
    ]
}

async function createShiftText(user){
    const turnos = await getShiftTurn(user.job_interviews[0].turno);
    if(turnos && turnos.length > 0){
        return new Paragraph({
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
                    size: 24,
                    text: `${turnos[0].name}`,
                    font: "Arial",
                }),
            ],
        })    
    }    
}

function createTimeStamp(user){
    return [
        new Paragraph({
            alignment: AlignmentType.JUSTIFIED,
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
                    text: `${user.address.city.name}, ${moment().format('DD [de] MMMM [de] YYYY.')}`,
                    font: "Arial",
                }),
            ],
        }),
    ]
}

function createSignSection(user, doc){
    const sectionCompany = {
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
        },
        children: [
            new Table({
                columnWidths: [4500, 4500],
                    width: {
                        size: 8800,
                        type: WidthType.DXA,
                    },
                rows: [
                    new TableRow({
                        children:[
                            new TableCell({
                                borders: {
                                    top: {
                                        style: BorderStyle.SINGLE,
                                        size: 30,
                                        color: "ffffff",
                                    },
                                    bottom: {
                                        style: BorderStyle.SINGLE,
                                        size: 30,
                                        color: "ffffff",
                                    },
                                    left: {
                                        style: BorderStyle.SINGLE,
                                        size: 30,
                                        color: "ffffff",
                                    },
                                    right: {
                                        style: BorderStyle.SINGLE,
                                        size: 30,
                                        color: "ffffff",
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
                                            new TextRun({text:`${user.name}`, 
                                            bold: true,
                                            font: "Arial",}),
                                        ],
                                        alignment: AlignmentType.CENTER,
                                        spacing: {
                                            before: 500,
                                        },
                                    }),
                                    new Paragraph({
                                        spacing: {
                                            after: 1200,
                                        },
                                        alignment: AlignmentType.CENTER,
                                        children: [
                                            new TextRun({
                                                size: 16,
                                                text: `CPF: ${user.cpf}`,
                                                font: "Arial",
                                            }),
                                        ],
                                    }),
                                ]
                            }),
                            new TableCell({
                                borders: {
                                    top: {
                                        style: BorderStyle.SINGLE,
                                        size: 30,
                                        color: "ffffff",
                                    },
                                    bottom: {
                                        style: BorderStyle.SINGLE,
                                        size: 30,
                                        color: "ffffff",
                                    },
                                    left: {
                                        style: BorderStyle.SINGLE,
                                        size: 30,
                                        color: "ffffff",
                                    },
                                    right: {
                                        style: BorderStyle.SINGLE,
                                        size: 30,
                                        color: "ffffff",
                                    },
                                },
                                children:[
                                    new Paragraph({
                                        alignment: AlignmentType.CENTER,
                                        spacing: {
                                            before: 500,
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
                                                text: `CPF DO GERENTE`,
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
                                                text: `CPF DO GESTOR`,
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
                                ]
                            })
                        ]
                    })
                ]
            }),
        ],
    }
    doc.addSection(sectionCompany);
}

function createRenew(user, doc){
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
                alignment: AlignmentType.JUSTIFIED,
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
                    ...createCompanyTextRun()
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
                    ...createEmployeeTextRun(user),
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
                alignment: AlignmentType.JUSTIFIED,
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
                alignment: AlignmentType.JUSTIFIED,
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
            ...createTimeStamp(user),
        ],
    }
    doc.addSection(section);
}

function createAttributionsSection(attributions, doc) {
    const assignments = [];
    attributions.forEach(att => {
        assignments.push(
            new Paragraph({
                alignment: AlignmentType.JUSTIFIED,
                children: [
                    new TextRun({
                        text: `- ${att}`,
                        size: 24,
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
            })
        )
    })
    const section = {
        properties: {
            type: SectionType.CONTINUOUS,
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
        children: assignments,
    }
    doc.addSection(section);
}

async function createNewSection(user, doc){
    const section = {
        properties: {
            type: SectionType.CONTINUOUS,
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
                        text: '2.4 ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        text: 'Além disso, se obriga a realizar o que vier a ser objeto de  cartas, avisos ou ordens, dentro da natureza do seu cargo e também o que  dispensa especificações por estar naturalmente compreendido, subentendido ou  relacionado ao seu cargo, não constituindo a indicação supra ou a de adendos,  qualquer limitação ou restrição, considerando-se falta grave a recusa por parte  do(a) ',
                        size: 24,
                        font: "Arial",
                    }),
                    new TextRun({
                        text: 'EMPREGADO(A) ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        text: 'em executar qualquer um dos serviços referidos, mesmo que  anteriormente não os tenha feito, mas que se entendam atinentes à função para  a qual fica contratado.',
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
                        text: '2.5 ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        text: 'A circunstância, porém, de ser a função especificada não importa na intransferibilidade do(a) ',
                        size: 24,
                        font: "Arial",
                    }),
                    new TextRun({
                        text: 'EMPREGADO(A) ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'para outro serviço no qual demonstre melhor capacidade de adaptação desde que compatível com sua condição pessoal.  ',
                        font: "Arial",
                    }),
                ],
            }),
            createClauseHeader('CLÁUSULA III - DA REMUNERAÇÃO E BENEFÍCIOS'),
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
                        text: '3.1 ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'O(a) ',
                        font: "Arial",
                    }),
                    new TextRun({
                        text: 'EMPREGADO(A) ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        text: 'receberá pelos serviços prestados do(a) ',
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
                        text: `a quantia de R$ ${user.skill.salary_base} (${extenso(user.skill.salary_base.toString().replace('.', ','), { mode: 'currency' })}), mensais no quinto dia útil do mês subsequente.`,
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
                        text: 'Além do salário base, o(a) ',
                        size: 24,
                        font: "Arial",
                    }),
                    new TextRun({
                        text: 'EMPREGADO(A) ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        text: 'terá direito aos seguintes benefícios:',
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
                        text: 'A. ',
                        size: 24,
                        font: "Arial",
                    }),
                    new TextRun({
                        text: 'ALIMENTAÇÃO NO LOCAL: O(A) ',
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
                        text: 'fornecerá ao ',
                        size: 24,
                        font: "Arial",
                    }),
                    new TextRun({
                        text: 'EMPREGADO(A) ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        text: 'almoço e café da manhã durante os dias de trabalho.',
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
                        text: 'B. ',
                        size: 24,
                        font: "Arial",
                    }),
                    new TextRun({
                        text: 'PRÊMIO POR ASSIDUIDADE: O(A) ',
                        size: 24,
                        font: "Arial",
                    }),
                    new TextRun({
                        text: 'EMPREGADO(A) ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        text: 'terá direito a um prêmio caso não tenha faltas injustificadas durante o mês, conforme função e localidade',
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
                        text: 'C. ',
                        size: 24,
                        font: "Arial",
                    }),
                    new TextRun({
                        text: 'AUXÍLIO TRANSPORTE: O(A) ',
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
                        text: 'concederá auxílio transporte referente aos dias úteis em que o ',
                        size: 24,
                        font: "Arial",
                    }),
                    new TextRun({
                        text: 'EMPREGADO(A) ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        text: 'comparecer ao trabalho.',
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
                        text: 'D. ',
                        size: 24,
                        font: "Arial",
                    }),
                    new TextRun({
                        text: 'PREMIAÇÃO POR TAREFAS EXCEPCIONAIS: O(A) ',
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
                        text: 'poderá conceder premiações adicionais ao ',
                        size: 24,
                        font: "Arial",
                    }),
                    new TextRun({
                        text: 'EMPREGADO(A) ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        text: 'pela realização de tarefas excepcionais, conforme critérios a serem estabelecidos pelo(a) ',
                        size: 24,
                        font: "Arial",
                    }),
                    new TextRun({
                        text: 'EMPREGADOR(A) ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                ],
            }),
            createClauseHeader('CLÁUSULA IV - DO HORÁRIO DE TRABALHO'),
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
                        text: '4.1 ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'A jornada será de 44 horas semanais distribuídas da seguinte maneira:',
                        font: "Arial",
                    }),
                    
                ],
            }),
            //JORNADA DE TRABALHO
            await createShiftText(user),
            new Paragraph({
                    alignment: AlignmentType.JUSTIFIED,

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
                        text: '4.2 ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'O(a) ',
                        font: "Arial",
                    }),
                    new TextRun({
                        text: 'EMPREGADO(A) ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'se compromete a trabalhar em regime de compensação e de prorrogação de horas, inclusive em período noturno, sempre que as necessidades assim o exigirem, observadas as formalidades legais.',
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
                        text: '4.3 ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'Neste ato, o(a) ',
                        font: "Arial",
                    }),
                    new TextRun({
                        text: 'EMPREGADO(A) ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'está ciente de que apenas poderá trabalhar fora do horário de trabalho se previamente autorizado por seu gestor ou com fulcro no art. 59 §5º da CLT, referente ao banco de horas, a reger-se sob com as seguintes regras:',
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
                    before: 300,
                },
                children: [
                    new TextRun({
                        text: 'A. ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'A jornada extra será limitada a 2 (duas) horas por dia;',
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
                },
                children: [
                    new TextRun({
                        text: 'B. ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'A compensação será realizada na proporção 1x1 (uma hora extra trabalhada corresponderá uma hora de folga).',
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
                },
                children: [
                    new TextRun({
                        text: 'C. ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'O prazo para compensação será de 06 meses, contados a partir da inclusão da hora realizada ao Banco de Horas. Caso haja disposição contrária em instrumento coletivo da categoria, esta prevalecerá.',
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
                },
                children: [
                    new TextRun({
                        text: 'D. ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'A data para desfrutar da compensação está sujeita à aprovação por parte do(a)',
                        font: "Arial",
                    }),
                    new TextRun({
                        text: 'EMPREGADOR(A) ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'e deve ser solicitada pelo(a) ',
                        font: "Arial",
                    }),
                    new TextRun({
                        text: 'EMPREGADO(A) ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'com antecedência mínima de 15 dias, afim de não prejudicar a operacionalização das atividades da empresa.',
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
                },
                children: [
                    new TextRun({
                        text: '4.4 ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'O(A)',
                        font: "Arial",
                    }),
                    new TextRun({
                        text: 'EMPREGADO(A) ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'aceita que poderá laborar em sábados, domingos e feriados, conforme escala divulgada e respeitando a legislação a respeito do tema.',
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
                },
                children: [
                    new TextRun({
                        text: '4.5 ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'Não será tolerado descumprimentos nas políticas de horários, não podendo atrasar mais que 10 minutos no total por dia e atrasos injustificados. Em caso de faltas, levar o atestado em até 48 horas imediatamente subsequentes. Faltas sem atestados podem ser descontadas na folha de pagamento e poderá ser aplicada advertência.',
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
                },
                children: [
                    new TextRun({
                        text: '4.6 ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'Ao funcionário é vedado realizar jornada extraordinária sem autorização da direção.',
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
                },
                children: [
                    new TextRun({
                        text: '4.7 ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'É vedado ao funcionário, em qualquer hipótese, realizar repouso intrajornada inferior a 1 (uma) hora.',
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
                },
                children: [
                    new TextRun({
                        text: '4.8 ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'As folgas poderão ser participadas pelo funcionário, entretanto, somente serão concedidas por expressa previsão da direção.',
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
                },
                children: [
                    new TextRun({
                        text: '4.9 ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'As ditas folgas somente serão concedidas caso o colaborador tenha saldo positivo de horas ou caso se trate de folga compensatória de feriados, sendo vedado banco de horas negativo.',
                        font: "Arial",
                    }),
                ],
            }),
            createClauseHeader('CLÁUSULA V - DOS DESCONTOS'),
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
                        text: '5.1 ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'O(A) ',
                        font: "Arial",
                    }),
                    new TextRun({
                        text: 'EMPREGADO(A) ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'autoriza o(a) ',
                        font: "Arial",
                    }),
                    new TextRun({
                        text: 'EMPREGADOR(A)',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'a descontar de sua folha de pagamento a contribuição sindical/confederativa/assistencial de sua categoria econômica ou profissional.',
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
                        text: '5.2 ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'Caso se oponha ao desconto, o funcionário deve realizar sua manifestação nos moldes do instrumento coletivo, sendo que o desconto somente será suspenso após demonstrado que houve o cumprimento dos requisitos trazidos pela CCT (Convenção Coletiva de Trabalho).',
                        font: "Arial",
                    }),
                ],
            }),
            createClauseHeader('CLÁUSULA VI - DO LOCAL DE TRABALHO'),
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
                        text: '6.1 ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'O(A) ',
                        font: "Arial",
                    }),
                    new TextRun({
                        text: 'EMPREGADO(A) ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'desempenhará suas funções, já estabelecidas no presente contrato, ao(à) ',
                        font: "Arial",
                    }),
                    new TextRun({
                        text: 'EMPREGADOR(A)',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: `no seguinte endereço: ${user.address.address}, nº ${user.address.number}, complemento: ${user.address.complement ? user.address.complement : 'Não informado'}, ${user.address.neighborhood}, ${user.address.city.name}, ${user.address.state.name}, CEP: ${user.address.postal_code}.`,
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
                        text: '6.2 ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'O(A) ',
                        font: "Arial",
                    }),
                    new TextRun({
                        text: 'EMPREGADOR(A) ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'poderá transferir o(a) ',
                        font: "Arial",
                    }),
                    new TextRun({
                        text: 'EMPREGADO(A) ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'entre, filiais, unidades do mesmo grupo econômico ou caso haja mudança de endereço da empresa, sem pagamento de qualquer adicional, desde que a mudança não importe a alteração do domicílio.',
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
                        text: '6.3 ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'Caso o(a) ',
                        font: "Arial",
                    }),
                    new TextRun({
                        text: 'EMPREGADO(A) ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'seja promovido a cargo de confiança, a vedação presente na cláusula anterior não se aplica, podendo o mesmo ser transferido, independente da alteração de domicílio, tendo em vista a natureza do cargo.',
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
                        text: '6.4 ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'Durante a vigência deste contrato, o(a) ',
                        font: "Arial",
                    }),
                    new TextRun({
                        text: 'EMPREGADO(A) ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'poderá ser transferido de forma provisória ou definitiva, para exercer sua função em localidade diversa daquela acima indicada, havendo mudança de domicílio, desde que haja a sua anuência ou que sejam verificadas as hipóteses legais tal como previsto no artigo 469 da Consolidação das Leis do Trabalho.',
                        font: "Arial",
                    }),
                ],
            }),
            createClauseHeader('CLÁUSULA VII - DAS OBRIGAÇÕES DAS PARTES'),
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
                        text: '7.1 ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'SÃO OBRIGAÇÕES DO ',
                        font: "Arial",
                    }),
                    new TextRun({
                        text: 'EMPREGADOR(A): ',
                        size: 24,
                        bold: true,
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
                        text: 'A. ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'O(A) EMPREGADOR(A) deverá pagar ao(a) EMPREGADO(A) os valores previstos na Cláusula Terceira, dentro do prazo e da forma previamente indicada, a título salarial;',
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
                        text: 'B. ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'O(A) ',
                        font: "Arial",
                    }),
                    new TextRun({
                        text: 'EMPREGADOR(A) ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'deverá fornecer todas as condições para que o(a) ',
                        font: "Arial",
                    }),
                    new TextRun({
                        text: 'EMPREGADO(A) ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'labore em ambiente de trabalho seguro, com boas condições sanitárias e com infraestrutura adequada à execução das atividades pelo(a) ',
                        font: "Arial",
                    }),
                    new TextRun({
                        text: 'EMPREGADO(A).',
                        size: 24,
                        bold: true,
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
                        text: 'C. ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'O(A) ',
                        font: "Arial",
                    }),
                    new TextRun({
                        text: 'EMPREGADOR(A), ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'no ato de celebração deste contrato, deverá cientificar o(a) ',
                        font: "Arial",
                    }),
                    new TextRun({
                        text: 'EMPREGADO(A) ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'de todas as regras de conduta estabelecidas e políticas internas, devendo entregar uma cópia do regulamento interno;',
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
                    before: 300,
                },
                children: [
                    new TextRun({
                        text: '7.1 ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'SÃO OBRIGAÇÕES DO ',
                        font: "Arial",
                    }),
                    new TextRun({
                        text: 'EMPREGADO(A): ',
                        size: 24,
                        bold: true,
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
                        text: 'A. ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'O(A) ',
                        font: "Arial",
                    }),
                    new TextRun({
                        text: 'EMPREGADO(A) ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'se compromete a executar as funções objeto do presente contrato, conforme as exigências, diretrizes e padrões exigidos pelo(a) ',
                        font: "Arial",
                    }),
                    new TextRun({
                        text: 'EMPREGADOR(A) ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'bem como realizá-las com empenho para o melhor desenvolvimento do trabalho, preservando a qualidade e os prazos pactuados;',
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
                        text: 'B. ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'O(A) ',
                        font: "Arial",
                    }),
                    new TextRun({
                        text: 'EMPREGADO(A) ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'se compromete a prestar ao(à) ',
                        font: "Arial",
                    }),
                    new TextRun({
                        text: 'EMPREGADOR(A) ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'as informações necessárias sobre o andamento das atividades desenvolvidas;',
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
                        text: 'C. ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'O(A) ',
                        font: "Arial",
                    }),
                    new TextRun({
                        text: 'EMPREGADO(A) ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'deverá manter durante toda vigência deste contrato, comportamento compatível com as normas de disciplina, da ética profissional e de segurança estabelecidas pela legislação brasileira e pelas normas internas do(a) ',
                        font: "Arial",
                    }),
                    new TextRun({
                        text: 'EMPREGADOR(A) ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'declarando estar ciente dos seus termos e condições;',
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
                        text: 'D. ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'O(A) ',
                        font: "Arial",
                    }),
                    new TextRun({
                        text: 'EMPREGADO(A) ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'se compromete a utilizar adequadamente os equipamentos, uniformes e materiais fornecidos pelo(a) ',
                        font: "Arial",
                    }),
                    new TextRun({
                        text: 'EMPREGADOR(A) ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'os quais devem ser utilizados apenas para os fins profissionais contratados, podendo o(a) ',
                        font: "Arial",
                    }),
                    new TextRun({
                        text: 'EMPREGADOR(A) ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'realizar vistorias periódicas nos equipamentos por ele fornecido, desde a verificação de e-mails corporativos até a delimitação do recebimento e envio de arquivos;',
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
                        text: 'E. ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'O(A) ',
                        font: "Arial",
                    }),
                    new TextRun({
                        text: 'EMPREGADO(A) ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'assume estar ciente de que todos os códigos e senhas, chaves e contas de monitoramento fornecidos pelo(a) ',
                        font: "Arial",
                    }),
                    new TextRun({
                        text: 'EMPREGADOR(A) ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'para utilização dos equipamentos são estritamente confidenciais, devendo ele tomar todas as cautelas na sua guarda.',
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
                        text: 'F. ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'Obriga-se o(a) ',
                        font: "Arial",
                    }),
                    new TextRun({
                        text: 'EMPREGADO(A) ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'além de executar com dedicação e lealdade o seu serviço, a cumprir o Regulamento Interno do(a) ',
                        font: "Arial",
                    }),
                    new TextRun({
                        text: 'EMPREGADOR(A), ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'as instruções de sua administração e as ordens de seus chefes e superiores hierárquicos, relativas às peculiaridades dos serviços que lhe forem confiados.',
                        font: "Arial",
                    }),
                ],
            }),
            createClauseHeader('CLÁUSULA VIII - DO PRAZO'),
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
                        text: '8.1 ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'O presente contrato entrará em vigor a partir da sua assinatura, na condição de CONTRATO DE EXPERIÊNCIA, e terá vigência de 45 (quarenta e cinco dias), renováveis por igual período automaticamente.',
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
                        text: '8.2 ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'Ao final do prazo de vigência previsto na cláusula anterior, o contrato poderá ser rescindido; caso contrário, será tacitamente convertido em contrato de trabalho por tempo indeterminado, sendo mantidas todas as demais cláusulas e obrigações aqui estabelecidas.',
                        font: "Arial",
                    }),
                ],
            }),
            createClauseHeader('CLÁUSULA IX - DA CONFIDENCIALIDADE'),
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
                        text: '9.1 ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'O(A) ',
                        font: "Arial",
                    }),
                    new TextRun({
                        text: 'EMPREGADO(A) ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'deverá manter em sigilo, durante a vigência do presente termo e mesmo após sua extinção, sobre qualquer informação relativa aos negócios, políticas, segredos institucionais, organização, criação, lista de clientes, quadro de funcionários, faturamento, metas e comissões, bem como outras informações relativas ao(à) ',
                        font: "Arial",
                    }),
                    new TextRun({
                        text: 'EMPREGADOR(A), ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'dados de seus clientes, fornecedores, representantes, demais ',
                        font: "Arial",
                    }),
                    new TextRun({
                        text: 'EMPREGADO(A)S ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'ou classificadas como confidenciais.',
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
                        text: '9.2 ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'Para fins do presente contrato, entende-se por informação confidencial:',
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
                },
                children: [
                    new TextRun({
                        text: 'A. ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'Qualquer informação relacionada ao negócio e operações do(a) ',
                        font: "Arial",
                    }),
                    new TextRun({
                        text: 'EMPREGADOR(A) ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'que não sejam públicas',
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
                        text: 'B. ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'Informações contidas em pesquisas, faturamento, metas, comissões, planos de negócio, venda ou marketing, informações financeiras, custos, dados de precificação, parceiros de negócios, informações de fornecedores e clientes, propriedade intelectual, especificações, expertises, técnicas, invenções e todos os métodos, conceitos ou ideias relacionadas ao negócio do ',
                        font: "Arial",
                    }),
                    new TextRun({
                        text: 'EMPREGADOR(A) ',
                        size: 24,
                        bold: true,
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
                        text: 'C. ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'É vedado ao(a)  ',
                        font: "Arial",
                    }),
                    new TextRun({
                        text: 'EMPREGADO(A) ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'repassar a terceiros, sejam particulares ou pessoas jurídicas, quaisquer destas informações, exceto quando expressamente autorizado pelo(a) ',
                        font: "Arial",
                    }),
                    new TextRun({
                        text: 'EMPREGADOR(A) ',
                        size: 24,
                        bold: true,
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
                        text: 'D. ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'A confidencialidade dessas informações independe de aviso prévio do(a) ',
                        font: "Arial",
                    }),
                    new TextRun({
                        text: 'EMPREGADOR(A) ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'devendo o(a)',
                        font: "Arial",
                    }),
                    new TextRun({
                        text: 'EMPREGADO(A) ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'considerar toda e qualquer informação relacionada ao negócio do(a) ',
                        font: "Arial",
                    }),
                    new TextRun({
                        text: 'EMPREGADOR(A) ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'e dos serviços prestado em sede dele como confidencial.',
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
                        text: 'E. ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'Ressalta-se que o dever de confidencialidade permanece mesmo após o término deste contrato de trabalho.',
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
                        text: 'F. ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'A violação da obrigação de confidencialidade pode causar a rescisão imediata deste contrato por justa causa, conforme o artigo 482, alínea g da CLT.',
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
                        text: 'G. ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'Em caso de violação desta cláusula o(a) ',
                        font: "Arial",
                    }),
                    new TextRun({
                        text: 'EMPREGADO(A), ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'poderá ser responsabilizado pelo pagamento das quantias equivalentes ao dano causado e estará sujeito ao pagamento de multa no valor de 5.000,00 (cinco mil reais), a ser devidamente atualizada e corrigidas no momento de sua aplicação e, ainda, estará sujeito a eventuais penalidades civis e criminais eventualmente aplicáveis. ',
                        font: "Arial",
                    }),
                ],
            }),
            createClauseHeader('CLÁUSULA X - DOS DIREITOS AUTORAIS E DA PROPRIEDADE INTELECTUAL'),
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
                        text: '10.1  ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'O(A) ',
                        font: "Arial",
                    }),
                    new TextRun({
                        text: 'EMPREGADO(A), ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'declara estar ciente de que todo e qualquer direito advindo ou relacionado ao trabalho por ele(a) desempenhado, direta ou indiretamente, com os serviços prestados em decorrência do presente contrato, pertencerão exclusivamente ao(à) ',
                        font: "Arial",
                    }),
                    new TextRun({
                        text: 'EMPREGADOR(A), ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'nos termos da legislação vigente.',
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
                        text: '10.2  ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'Nesse ponto, também é objeto do presente contrato a cessão e transferência em favor do(a) ',
                        font: "Arial",
                    }),
                    new TextRun({
                        text: 'EMPREGADOR(A), ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'expressamente, na integralidade, a título universal e gratuito, em caráter irretratável e irrevogável, para fins de utilização a qualquer tempo, para fins de utilização econômica ou não, no Brasil e/ou no Exterior, de todos os direitos patrimoniais de autoria sobre documentos de modo geral referente às Obras que já tenham sido ou ainda sejam criadas pelo(a) ',
                        font: "Arial",
                    }),
                    new TextRun({
                        text: 'EMPREGADO(A) ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'no âmbito da relação de trabalho com o(a) ',
                        font: "Arial",
                    }),
                    new TextRun({
                        text: 'EMPREGADOR(A) ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'abrangendo tal cessão a criação, aperfeiçoamento, redação, revisão, edição, tradução, adaptação e toda e qualquer atividade que enseje proteção de direito de autor com relação às referidas Obras, que decorra, direta ou indiretamente, das atividades exercidas pelo(a) ',
                        font: "Arial",
                    }),
                    new TextRun({
                        text: 'EMPREGADO(A) ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'em razão da relação mantida com ',
                        font: "Arial",
                    }),
                    new TextRun({
                        text: 'EMPREGADOR(A).',
                        size: 24,
                        bold: true,
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
                        text: '10.3  ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'O disposto na Cláusula acima tem validade por todo o tempo em que a obra estiver protegida por direitos autorais.',
                        font: "Arial",
                    }),
                ],
            }),
            createClauseHeader('CLÁUSULA XI – DO REGISTRO EM CARTEIRA DE TRABALHO'),
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
                        text: '11.1  ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'O(A) ',
                        font: "Arial",
                    }),
                    new TextRun({
                        text: 'EMPREGADO(A), ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'declara estar cientes e de comum acordo com o registro digital da carteira de trabalho, mesmo que tenha sido fornecida uma (CTPS) física para recolhimento de dados no ato da sua contratação',
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
                        text: '11.2 ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'Qualquer anotação, alteração ou modificação referente ao vínculo empregatício será realizada na versão digital da carteira de trabalho do(a) ',
                        font: "Arial",
                    }),
                    new TextRun({
                        text: 'EMPREGADO(A), ',
                        size: 24,
                        bold: true,
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
                        text: '11.3 ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'O(A) ',
                        font: "Arial",
                    }),
                    new TextRun({
                        text: 'EMPREGADO(A), ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'assume responsabilidade de manter-se informado(a) sobre as atualizações realizadas na carteira de trabalho digital, garantindo o acompanhamento de todas as anotações pertinentes ao seu histórico profissional.',
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
                        text: '11.4 ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'O(A) ',
                        font: "Arial",
                    }),
                    new TextRun({
                        text: 'EMPREGADO(A), ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'está ciente que as implicações e os benefícios decorrentes do registro digital da minha carteira de trabalho, e concorda em utilizar este meio como forma de registro e documentação do seu histórico laboral.',
                        font: "Arial",
                    }),
                ],
            }), 
            createClauseHeader('CLÁUSULA XII - DAS DISPOSIÇÕES GERAIS'),
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
                        text: '12.1  ',
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
                        text: 'EMPREGADO(A), ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'comunicado e ciente que durante a permanência no local de trabalho está sendo monitorado por câmeras de segurança que possuem gravações de áudio e vídeo.',
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
                    before: 300,
                },
                children: [
                    new TextRun({
                        text: '12.2  ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'O(A) ',
                        font: "Arial",
                    }),
                    new TextRun({
                        text: 'EMPREGADOR(A), ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'poderá fiscalizar e ter acesso a qualquer informação constante nos softwares utilizados no ambiente laborativo, inclusive se utilizando de programas de monitoramento.',
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
                    before: 300,
                },
                children: [
                    new TextRun({
                        text: '12.3  ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'Conforme dispõe a CLT em seu art.2º, é o(a) ',
                        font: "Arial",
                    }),
                    new TextRun({
                        text: 'EMPREGADOR(A), ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'quem dirige a prestação pessoal de serviço. Assim, fica estipulado que é expressamente proibida a utilização de celular pessoal ou de outro aparelho eletrônico que se assemelhe (tablet, smartwatch, etc.) durante a jornada de trabalho para tratar de qualquer assunto de interesse pessoal do funcionário.',
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
                    before: 300,
                },
                children: [
                    new TextRun({
                        text: '12.4  ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'É vedado ao(a) ',
                        font: "Arial",
                    }),
                    new TextRun({
                        text: 'EMPREGADO(A), ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'receber qualquer tipo de contraprestação, comissões, glosas, gorjetas de fornecedores, clientes ou terceiros, sem o consentimento expresso do(a) ',
                        font: "Arial",
                    }),
                    new TextRun({
                        text: 'EMPREGADOR(A), ',
                        size: 24,
                        bold: true,
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
                    before: 300,
                },
                children: [
                    new TextRun({
                        text: '12.5  ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'O(A) ',
                        font: "Arial",
                    }),
                    new TextRun({
                        text: 'EMPREGADO(A), ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'autoriza o uso de sua imagem e voz, em todo e qualquer material entre fotos, documentos e outros meios de comunicação, para campanhas promocionais e institucionais. A presente autorização é concedida a título gratuito, abrangendo o uso da imagem acima mencionada em todo território nacional e no exterior, sob qualquer forma e meios, ou sejam, em destaque: (I) out-door; (II) bus-door; folhetos em geral (encartes, mala direta, catálogo, etc.); (III) folder de apresentação; (IV) anúncios em revistas e jornais em geral; (V) home page; (VI) cartazes; (VII) back-light; (VIII) mídia eletrônica (painéis, vídeo-tapes, televisão, cinema, programa para rádio, entre outros); (IX) redes e mídias sociais. ',
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
                    before: 300,
                },
                children: [
                    new TextRun({
                        text: '12.6  ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'Fica proibido tratar de trabalho após o expediente em qualquer meio telemático.',
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
                    before: 300,
                },
                children: [
                    new TextRun({
                        text: '12.7  ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'E por estarem em pleno acordo, as partes contratantes assinam o presente Contrato de Experiência em duas vias, ficando a primeira em poder do(a) ',
                        font: "Arial",
                    }),
                    new TextRun({
                        text: 'EMPREGADOR(A), ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'e a segunda com o(a) ',
                        font: "Arial",
                    }),
                    new TextRun({
                        text: 'EMPREGADO(A), ',
                        size: 24,
                        bold: true,
                        font: "Arial",
                    }),
                ],
            }),
            ...createTimeStamp(user),
        ],
    }
    doc.addSection(section);
}

module.exports = {createCompanyTextRun, createEmployeeTextRun, createAttributionsSection, createClauseWithParagraphs, createEpiProductsTable, getSkillProducts, createEpiTable, countDays, createXPContract, createClauseHeader, createSignSection, createRenew, createNewSection}