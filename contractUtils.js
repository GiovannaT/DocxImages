const axios = require('axios')
const moment = require("moment")
const extenso = require('extenso')
const fs = require('fs');


const { convertInchesToTwip, Paragraph, Table, TextRun, TableRow, TableCell, ShadingType, AlignmentType, WidthType, SectionType, BorderStyle, ImageRun} = require("docx");

moment.locale('pt'); 
moment.updateLocale('pt', {
    months : [
        "janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho",
        "agosto", "setembro", "outubro", "novembro", "dezembro"
    ]
});

const leadership = true;

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
        const response = await axios.get(`https://3337-avanciconstru-apiserver-liom39gdw1s.ws-us114.gitpod.io/skillproducts?skill_id=${id}`, {
            headers: {
            'Authorization': `Bearer eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpYXQiOjE3MTgwNTA5NjksImV4cCI6MTc2OTg5MDk2OSwic3ViIjoiZjM1ZDg2M2QtMmI4My00MGM4LWI4ZDUtM2ExNzU5YTU2NTc2In0.RVOvNkHAqsTmp_iRfvu28sbeMbwfauCA4E7_7fWpp78`
        }});
        return response.data;
    } catch (error) {
        console.error('erro no get skill product')
        return null;
    }
}

async function getShiftTurn(id){
    try {
        const response = await axios.get(`https://3337-avanciconstru-apiserver-liom39gdw1s.ws-us114.gitpod.io/timeclock/turno?id=${id}`, {
            headers: {
            'Authorization': `Bearer eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpYXQiOjE3MTgwNTA5NjksImV4cCI6MTc2OTg5MDk2OSwic3ViIjoiZjM1ZDg2M2QtMmI4My00MGM4LWI4ZDUtM2ExNzU5YTU2NTc2In0.RVOvNkHAqsTmp_iRfvu28sbeMbwfauCA4E7_7fWpp78`
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
                            text: `Recebi de ${company === 'HABIT' ? "HABIT CONSTRUÇÕES E SERVIÇOS EIRELLI" : "AVANCI CONSTRUÇÃO E COMÉRCIO DE IMPORTAÇÃO E EXPORTAÇÃO EIRELI"}, para meu uso obrigatório os EPI's (Equipamentos de proteção Individual) e EPC's (Equipamentos de2 proteção coletiva) constantes nesta ficha, o qual obrigo-me a utiliza-los corretamente durante o tempo que permanecerem ao meu dispor, observando as medidas gerais de disciplina e uso que integram a NR-06 - Equipamento de Proteção Individual - EPI's - da portaria n.º 3.214 de 08/06/1978.`,
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
                    font: "Calibri",
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
                    size: 28,
                    font: "Calibri",
                })
            ],
            indent: {
                left: convertInchesToTwip(0.5),
                right: convertInchesToTwip(0.5),
            },
            spacing: {
                before: 300,
            },
        })
    )
}

function createPageHeader(company){
    return (
        new Paragraph({
            alignment: AlignmentType.CENTER,    
            children: [
                new ImageRun({
                    type: 'png',
                    data: fs.readFileSync(`./public/${company === 'HABIT' ? 'hbt.png' : 'avanci.png'}`),
                    transformation: {
                        width: 100,
                        height: 50,
                    },
                })
            ]
        })
    )
}

function createPageFooter(company){
    return [
        new Paragraph({
            alignment: AlignmentType.CENTER,    
            children: [
                new TextRun({
                    text: company === 'HABIT' ? "HABIT CONSTRUÇÕES E SERVIÇOS EIRELLI" : "AVANCI CONSTRUÇÃO E COMÉRCIO DE IMPORTAÇÃO E EXPORTAÇÃO EIRELI",
                    bold: true,
                    size: 10,
                })
            ],
        }),
        new Paragraph({
            alignment: AlignmentType.CENTER,    
            children: [
                new TextRun({
                    text:  company === 'HABIT' ? "CNPJ: 28.697.931/0001-43 - Rua Projetada F, nº 76 - Ribeirão do Lipa - Cuiabá - Mato Grosso" : "CNPJ: 32.953.515/0001-00 - Rodovia Emanuel Pinheiro - Vila Formosa, KM 04, nº 130",
                    size: 10,
                })
            ],
        }),
        new Paragraph({
            alignment: AlignmentType.CENTER,    
            children: [
                new TextRun({
                    text: company === 'HABIT' ? "CEP: 78048-163 - Telefone para contato: 65981088373 - recepcaohbt@gmail.com" : "Cuiabá - Mato Grosso - CEP: 78055-799 - Telefone para contato: 6598108-8373, recepcaohbt@gmail.com",
                    size: 10,
                })
            ],
        })
    ]
}

function createCompanyTextRun(company){
    return [
        new TextRun({
            text: `EMPREGADOR(A): ${company === 'HABIT' ? "HABIT CONSTRUÇÕES E SERVIÇOS LTDA," : "AVANCI CONSTRUÇÃO E COMÉRCIO DE IMPORTAÇÃO E EXPORTAÇÃO EIRELI"} `,
            size: 24,
            bold: true,
            font: "Calibri Light",
        }),
        new TextRun({
            text: `${company === 'HABIT' 
                ? 'inscrito no CNPJ nº: 28.697.934/0001- 43, e-mail recepcaohbt@gmail.com, telefone para contato 6598108-8373, com sede em rua Projetada F, Nº 76, Ribeirão do Lipa, Cuiabá – MT, CEP: 78048-163.' 
                : 'inscrito no CNPJ nº: 32.953.515/0001-00, e-mail recepcaohbt@gmail.com, telefone para contato 6598108-8373, com sede em Rodovia Emanuel Pinheiro - Vila Formosa, KM 04, nº 130, Cuiabá – MT, CEP: 78055-799.'}`,
            size: 24,
            font: "Calibri Light",
        })
    ]
}

function createEmployeeTextRun(user){
    return [
        new TextRun({
            text: `EMPREGADO(A): ${user.name} `,
            size: 24,
            bold: true,
            font: "Calibri Light",
        }),
        new TextRun({
            size: 24,
            text: `CPF: ${user.cpf}, RG: ${user.rg}, data de emissão ${moment(user.meta.issuedOn).format('DD/MM/YYYY')}, data de nascimento: ${moment(user.birthday).format('DD/MM/YYYY')}, telefone para contato ${user.phone}, ${user.meta.maritalStatus}, residente em ${capitalizeString(user.address.address)}, nº ${user.address.number}, complemento: ${user.address.complement ? capitalizeString(user.address.complement) : 'Não informado'}, ${capitalizeString(user.address.neighborhood)}, ${user.address.city.name}, ${user.address.state.name}, CEP: ${user.address.postal_code}.`,
            font: "Calibri Light",
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
                    font: "Calibri Light",
                }),
            ],
        })    
    }    
}

async function createShiftContent(user, leadership){
    try {
        if(leadership){
            return [ 
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
                            font: "Calibri Light",
                        }),
                        new TextRun({
                            size: 24,
                            text: 'Em se tratando de cargo de confiança, valendo-se do que dispõe a CLT no que tange à jornada de trabalho, não haverá controle de jornada, não fazendo jus o colaborador ao percebimento, dentre outras, das seguintes verbas/direitos:',
                            font: "Calibri Light",
                        }),
                    ]
                }),
                new Paragraph({
                    alignment: AlignmentType.JUSTIFIED,
                    indent: {
                        left: convertInchesToTwip(1.5),
                    },
                    spacing: {
                        line: 276,
                        before: 200,
                    },
                    children: [
                        new TextRun({
                            text: 'I. ',
                            size: 24,
                            bold: true,
                            font: "Calibri Light",
                        }),
                        new TextRun({
                            size: 24,
                            text: 'adicional de horas extras',
                            font: "Calibri Light",
                        }),
                    ],
                }),
                new Paragraph({
                    alignment: AlignmentType.JUSTIFIED,
                    indent: {
                        left: convertInchesToTwip(1.5),
                    },
                    spacing: {
                        line: 276,
                        before: 200,
                    },
                    children: [
                        new TextRun({
                            text: 'II. ',
                            size: 24,
                            bold: true,
                            font: "Calibri Light",
                        }),
                        new TextRun({
                            size: 24,
                            text: 'adicional noturno',
                            font: "Calibri Light",
                        }),
                    ],
                }),
                new Paragraph({
                    alignment: AlignmentType.JUSTIFIED,
                    indent: {
                        left: convertInchesToTwip(1.5),
                    },
                    spacing: {
                        line: 276,
                        before: 200,
                    },
                    children: [
                        new TextRun({
                            text: 'III. ',
                            size: 24,
                            bold: true,
                            font: "Calibri Light",
                        }),
                        new TextRun({
                            size: 24,
                            text: 'intervalo interjornada',
                            font: "Calibri Light",
                        }),
                    ],
                }),
                new Paragraph({
                    alignment: AlignmentType.JUSTIFIED,
                    indent: {
                        left: convertInchesToTwip(1.5),
                    },
                    spacing: {
                        line: 276,
                        before: 200,
                    },
                    children: [
                        new TextRun({
                            text: 'IV. ',
                            size: 24,
                            bold: true,
                            font: "Calibri Light",
                        }),
                        new TextRun({
                            size: 24,
                            text: 'intervalo intrajornada',
                            font: "Calibri Light",
                        }),
                    ],
                }),
            ]
        } else {
            return [
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
                            font: "Calibri Light",
                        }),
                        new TextRun({
                            size: 24,
                            text: 'A jornada será de 44 horas semanais distribuídas da seguinte maneira:',
                            font: "Calibri Light",
                        }),
                        
                    ],
                }),
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
                            font: "Calibri Light",
                        }),
                        new TextRun({
                            size: 24,
                            text: 'O(a) ',
                            font: "Calibri Light",
                        }),
                        new TextRun({
                            text: 'EMPREGADO(A) ',
                            size: 24,
                            bold: true,
                            font: "Calibri Light",
                        }),
                        new TextRun({
                            size: 24,
                            text: 'se compromete a trabalhar em regime de compensação e de prorrogação de horas, inclusive em período noturno, sempre que as necessidades assim o exigirem, observadas as formalidades legais.',
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
                            text: '4.3 ',
                            size: 24,
                            bold: true,
                            font: "Calibri Light",
                        }),
                        new TextRun({
                            size: 24,
                            text: 'Neste ato, o(a) ',
                            font: "Calibri Light",
                        }),
                        new TextRun({
                            text: 'EMPREGADO(A) ',
                            size: 24,
                            bold: true,
                            font: "Calibri Light",
                        }),
                        new TextRun({
                            size: 24,
                            text: 'está ciente de que apenas poderá trabalhar fora do horário de trabalho se previamente autorizado por seu gestor ou com fulcro no art. 59 §5º da CLT, referente ao banco de horas, a reger-se sob com as seguintes regras:',
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
                        before: 300,
                    },
                    children: [
                        new TextRun({
                            text: 'A. ',
                            size: 24,
                            bold: true,
                            font: "Calibri Light",
                        }),
                        new TextRun({
                            size: 24,
                            text: 'A jornada extra será limitada a 2 (duas) horas por dia;',
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
                            text: 'B. ',
                            size: 24,
                            bold: true,
                            font: "Calibri Light",
                        }),
                        new TextRun({
                            size: 24,
                            text: 'A compensação será realizada na proporção 1x1 (uma hora extra trabalhada corresponderá uma hora de folga).',
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
                            text: 'C. ',
                            size: 24,
                            bold: true,
                            font: "Calibri Light",
                        }),
                        new TextRun({
                            size: 24,
                            text: 'O prazo para compensação será de 06 meses, contados a partir da inclusão da hora realizada ao Banco de Horas. Caso haja disposição contrária em instrumento coletivo da categoria, esta prevalecerá.',
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
                            text: 'D. ',
                            size: 24,
                            bold: true,
                            font: "Calibri Light",
                        }),
                        new TextRun({
                            size: 24,
                            text: 'A data para desfrutar da compensação está sujeita à aprovação por parte do(a)',
                            font: "Calibri Light",
                        }),
                        new TextRun({
                            text: 'EMPREGADOR(A) ',
                            size: 24,
                            bold: true,
                            font: "Calibri Light",
                        }),
                        new TextRun({
                            size: 24,
                            text: 'e deve ser solicitada pelo(a) ',
                            font: "Calibri Light",
                        }),
                        new TextRun({
                            text: 'EMPREGADO(A) ',
                            size: 24,
                            bold: true,
                            font: "Calibri Light",
                        }),
                        new TextRun({
                            size: 24,
                            text: 'com antecedência mínima de 15 dias, afim de não prejudicar a operacionalização das atividades da empresa.',
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
                            text: '4.4 ',
                            size: 24,
                            bold: true,
                            font: "Calibri Light",
                        }),
                        new TextRun({
                            size: 24,
                            text: 'O(A)',
                            font: "Calibri Light",
                        }),
                        new TextRun({
                            text: 'EMPREGADO(A) ',
                            size: 24,
                            bold: true,
                            font: "Calibri Light",
                        }),
                        new TextRun({
                            size: 24,
                            text: 'aceita que poderá laborar em sábados, domingos e feriados, conforme escala divulgada e respeitando a legislação a respeito do tema.',
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
                            text: '4.5 ',
                            size: 24,
                            bold: true,
                            font: "Calibri Light",
                        }),
                        new TextRun({
                            size: 24,
                            text: 'Não será tolerado descumprimentos nas políticas de horários, não podendo atrasar mais que 10 minutos no total por dia e atrasos injustificados. Em caso de faltas, levar o atestado em até 48 horas imediatamente subsequentes. Faltas sem atestados podem ser descontadas na folha de pagamento e poderá ser aplicada advertência.',
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
                            text: '4.6 ',
                            size: 24,
                            bold: true,
                            font: "Calibri Light",
                        }),
                        new TextRun({
                            size: 24,
                            text: 'Ao funcionário é vedado realizar jornada extraordinária sem autorização da direção.',
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
                            text: '4.7 ',
                            size: 24,
                            bold: true,
                            font: "Calibri Light",
                        }),
                        new TextRun({
                            size: 24,
                            text: 'É vedado ao funcionário, em qualquer hipótese, realizar repouso intrajornada inferior a 1 (uma) hora.',
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
                            text: '4.8 ',
                            size: 24,
                            bold: true,
                            font: "Calibri Light",
                        }),
                        new TextRun({
                            size: 24,
                            text: 'As folgas poderão ser participadas pelo funcionário, entretanto, somente serão concedidas por expressa previsão da direção.',
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
                        before: 200,
                        line: 276,
                    },
                    children: [
                        new TextRun({
                            text: '4.9 ',
                            size: 24,
                            bold: true,
                            font: "Calibri Light",
                        }),
                        new TextRun({
                            size: 24,
                            text: 'As ditas folgas somente serão concedidas caso o colaborador tenha saldo positivo de horas ou caso se trate de folga compensatória de feriados, sendo vedado banco de horas negativo.',
                            font: "Calibri Light",
                        }),
                    ],
                }),
            ]
        }
    }catch(error){
        console.log('erro no create shiftcontetn')
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
                    font: "Calibri Light",
                }),
            ],
        }),
    ]
}

function createMotoristaLaborPlaceClause(user){
    if(user.skill.name === 'Motorista'){
        return [
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
                    text: '6.5 ',
                    size: 24,
                    bold: true,
                    font: "Calibri Light",
                }),
                new TextRun({
                    size: 24,
                    text: 'A atividade do colaborador poderá se operar em viagens, hipótese em que através de diárias de viagem lhe serão indenizadas as despesas com hospedagem e alimentação, observado os limites impostos pela empresa em suas comunicações, normas e procedimentos internos.',
                    font: "Calibri Light",
                }),
            ],
        }),
        ]
    }
    return []
}

function createPaymentContent(user, leadership){
    if(leadership){
        return [ 
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
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'O(a) ',
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'EMPREGADO(A) ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'receberá pelos serviços prestados do(a) ',
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
                        text: `a quantia de R$ ${user.skill.salary_base} (${extenso(user.skill.salary_base.toString().replace('.', ','), { mode: 'currency' })}), mensais no quinto dia útil do mês subsequente.`,
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
                        text: 'Além do salário base, o(a) ',
                        size: 24,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'EMPREGADO(A) ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'terá direito aos seguintes benefícios:',
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
                        text: 'A. ',
                        size: 24,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'ALIMENTAÇÃO NO LOCAL: O(A) ',
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
                        text: 'fornecerá ao ',
                        size: 24,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'EMPREGADO(A) ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'almoço e café da manhã durante os dias de trabalho.',
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
                        text: 'D. ',
                        size: 24,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'PREMIAÇÃO POR TAREFAS EXCEPCIONAIS: O(A) ',
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
                        text: 'poderá conceder premiações adicionais ao ',
                        size: 24,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'EMPREGADO(A) ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'pela realização de tarefas excepcionais, conforme critérios a serem estabelecidos pelo(a) ',
                        size: 24,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'EMPREGADOR(A) ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                ],
            })
        ]
    } else {
        return [
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
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'O(a) ',
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'EMPREGADO(A) ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'receberá pelos serviços prestados do(a) ',
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
                        text: `a quantia de R$ ${user.skill.salary_base} (${extenso(user.skill.salary_base.toString().replace('.', ','), { mode: 'currency' })}), mensais no quinto dia útil do mês subsequente.`,
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
                        text: 'Além do salário base, o(a) ',
                        size: 24,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'EMPREGADO(A) ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'terá direito aos seguintes benefícios:',
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
                        text: 'A. ',
                        size: 24,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'ALIMENTAÇÃO NO LOCAL: O(A) ',
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
                        text: 'fornecerá ao ',
                        size: 24,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'EMPREGADO(A) ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'almoço e café da manhã durante os dias de trabalho.',
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
                        text: 'B. ',
                        size: 24,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'PRÊMIO POR ASSIDUIDADE: O(A) ',
                        size: 24,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'EMPREGADO(A) ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'terá direito a um prêmio caso não tenha faltas injustificadas durante o mês, conforme função e localidade',
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
                        text: 'C. ',
                        size: 24,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'AUXÍLIO TRANSPORTE: O(A) ',
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
                        text: 'concederá auxílio transporte referente aos dias úteis em que o ',
                        size: 24,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'EMPREGADO(A) ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'comparecer ao trabalho.',
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
                        text: 'D. ',
                        size: 24,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'PREMIAÇÃO POR TAREFAS EXCEPCIONAIS: O(A) ',
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
                        text: 'poderá conceder premiações adicionais ao ',
                        size: 24,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'EMPREGADO(A) ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'pela realização de tarefas excepcionais, conforme critérios a serem estabelecidos pelo(a) ',
                        size: 24,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'EMPREGADOR(A) ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                ],
            })
        ]
    }
}

function createClause12Paragraph(user){
    if(user.skill.name === 'Motorista'){
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
                        text: '12.6  ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'A perda da habilitação para o cargo é motivo para justa causa.',
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
                    before: 300,
                },
                children: [
                    new TextRun({
                        text: '12.7  ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'A utilização do veículo da empresa em qualquer tipo de atividade de cunho pessoal é expressamente proibida.',
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
                    before: 300,
                },
                children: [
                    new TextRun({
                        text: '12.8  ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'Fica proibido tratar de trabalho após o expediente em qualquer meio telemático.',
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
                    before: 300,
                },
                children: [
                    new TextRun({
                        text: '12.9  ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'E por estarem em pleno acordo, as partes contratantes assinam o presente Contrato de Experiência em duas vias, ficando a primeira em poder do(a) ',
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'EMPREGADOR(A), ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'e a segunda com o(a) ',
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'EMPREGADO(A), ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                ],
            })
        ]
    } else {
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
                        text: '12.6  ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'Fica proibido tratar de trabalho após o expediente em qualquer meio telemático.',
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
                    before: 300,
                },
                children: [
                    new TextRun({
                        text: '12.7  ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'E por estarem em pleno acordo, as partes contratantes assinam o presente Contrato de Experiência em duas vias, ficando a primeira em poder do(a) ',
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'EMPREGADOR(A), ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'e a segunda com o(a) ',
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'EMPREGADO(A), ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                ],
            })
        ]
    }
}

function createLaborPlace(user, leadership){
    if(leadership){
        return [
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
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'O(A) ',
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'EMPREGADO(A) ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'desempenhará suas funções, já estabelecidas no presente contrato, ao(à) ',
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'EMPREGADOR(A)',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: `no seguinte endereço: ${capitalizeString(user.address.address)}, nº ${user.address.number}, complemento: ${user.address.complement ? capitalizeString(user.address.complement) : 'Não informado'}, ${capitalizeString(user.address.neighborhood)}, ${user.address.city.name}, ${user.address.state.name}, CEP: ${user.address.postal_code}.`,
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
                        text: '6.2 ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'O(A) ',
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'EMPREGADO(A) ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'poderá ser transferido(a),independente da alteração de domicílio, tendo em vista a natureza do cargo; ',
                        font: "Calibri Light",
                    }),
                ],
            })
        ]
    } else {
        return [
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
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'O(A) ',
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'EMPREGADO(A) ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'desempenhará suas funções, já estabelecidas no presente contrato, ao(à) ',
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'EMPREGADOR(A)',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: `no seguinte endereço: ${capitalizeString(user.address.address)}, nº ${user.address.number}, complemento: ${user.address.complement ? capitalizeString(user.address.complement) : 'Não informado'}, ${capitalizeString(user.address.neighborhood)}, ${user.address.city.name}, ${user.address.state.name}, CEP: ${user.address.postal_code}.`,
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
                        text: '6.2 ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'O(A) ',
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'EMPREGADOR(A) ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'poderá transferir o(a) ',
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'EMPREGADO(A) ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'entre, filiais, unidades do mesmo grupo econômico ou caso haja mudança de endereço da empresa, sem pagamento de qualquer adicional, desde que a mudança não importe a alteração do domicílio.',
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
                        text: '6.3 ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'Caso o(a) ',
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'EMPREGADO(A) ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'seja promovido a cargo de confiança, a vedação presente na cláusula anterior não se aplica, podendo o mesmo ser transferido, independente da alteração de domicílio, tendo em vista a natureza do cargo.',
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
                        text: '6.4 ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'Durante a vigência deste contrato, o(a) ',
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'EMPREGADO(A) ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'poderá ser transferido de forma provisória ou definitiva, para exercer sua função em localidade diversa daquela acima indicada, havendo mudança de domicílio, desde que haja a sua anuência ou que sejam verificadas as hipóteses legais tal como previsto no artigo 469 da Consolidação das Leis do Trabalho.',
                        font: "Calibri Light",
                    }),
                ],
            }),
            ...createMotoristaLaborPlaceClause(user),
        ]
    }
}

function createSignSection(user, doc, company){
    const sectionCompany = {
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
            type: SectionType.CONTINUOUS,
        },
        children: [
            new Table({
                alignment: AlignmentType.CENTER,
                columnWidths: [4500, 4500],
                rows: [
                    new TableRow({
                        children:[
                            new TableCell({
                                width: {
                                    size: 4500,
                                    type: WidthType.DXA,
                                },
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
                                margins:{
                                  left: 200,
                                  right: 200,  
                                },
                                children: [
                                    new Paragraph({
                                        spacing: {
                                            before: 400,
                                        },
                                        children: [
                                            new TextRun({
                                                size: 24,
                                                text: `EMPREGADOR(A):`,
                                                bold: true,
                                                font: "Calibri Light",
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
                                            new TextRun({text:`${company === 'HABIT' ? "HABIT CONSTRUÇÕES E SERVIÇOS EIRELLI" : "AVANCI CONSTRUÇÃO E COMÉRCIO DE IMPORTAÇÃO E EXPORTAÇÃO EIRELI"}`, 
                                            bold: true,
                                            size: 20,
                                            font: "Calibri Light",}),
                                        ],
                                        spacing: {
                                            before: 400,
                                        },
                                    }),
                                    new Paragraph({
                                        children: [
                                            new TextRun({
                                                size: 20,
                                                text: `CNPJ nº: 28.697.934/0001-43`,
                                                font: "Calibri Light",
                                            }),
                                        ],
                                    }),
                                    new Paragraph({
                                        spacing: {
                                            before: 500,
                                        },
                                        children: [
                                            new TextRun({
                                                size: 24,
                                                text: `EMPREGADO(A):`,
                                                bold: true,
                                                font: "Calibri Light",
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
                                            size: 20,
                                            font: "Calibri Light",}),
                                        ],
                                        spacing: {
                                            before: 500,
                                        },
                                    }),
                                    new Paragraph({
                                        children: [
                                            new TextRun({
                                                size: 20,
                                                text: `CPF: ${user.cpf}`,
                                                font: "Calibri Light",
                                            }),
                                        ],
                                    }),
                                ]
                            }),
                            new TableCell({
                                width: {
                                    size: 4500,
                                    type: WidthType.DXA,
                                },
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
                                margins:{
                                    left: 200,
                                    right: 200,  
                                },
                                children:[
                                    new Paragraph({
                                        spacing: {
                                            before: 200,
                                        },
                                        children: [
                                            new TextRun({
                                                size: 24,
                                                text: `TESTEMUNHAS:`,
                                                bold: true,
                                                font: "Calibri Light",
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
                                            size: 20,
                                            font: "Calibri Light",}),
                                        ],
                                        spacing: {
                                            before: 800,
                                        },
                                    }),
                                    new Paragraph({
                                        children: [
                                            new TextRun({
                                                size: 20,
                                                text: `CPF: 049.391.961-94`,
                                                font: "Calibri Light",
                                            }),
                                        ],
                                    }),
                                    new Paragraph({
                                        children: [
                                            new TextRun({
                                                size: 20,
                                                text: `SETOR DE RH`,
                                                font: "Calibri Light",
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
                                            size: 20,
                                            font: "Calibri Light",}),
                                        ],
                                        spacing: {
                                            before: 800,
                                        },
                                    }),
                                    new Paragraph({
                                        children: [
                                            new TextRun({
                                                size: 20,
                                                text: `CPF DO GERENTE`,
                                                font: "Calibri Light",
                                            }),
                                        ],
                                    }),
                                    new Paragraph({
                                        children: [
                                            new TextRun({
                                                size: 20,
                                                text: `ENGENHEIRO GERENTE`,
                                                font: "Calibri Light",
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
                                            font: "Calibri Light",}),
                                        ],
                                        spacing: {
                                            before: 800,
                                        },
                                    }),
                                    new Paragraph({
                                        children: [
                                            new TextRun({
                                                size: 20,
                                                text: `CPF DO GESTOR`,
                                                font: "Calibri Light",
                                            }),
                                        ],
                                    }),
                                    new Paragraph({
                                        children: [
                                            new TextRun({
                                                size: 20,
                                                text: `GESTOR RESPONSÁVEL`,
                                                font: "Calibri Light",
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

function createRenew(user, doc, company){
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
                        font: "Calibri",
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
                        font: "Calibri Light",
                    }),
                    ...createCompanyTextRun(company)
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
                    before: 300,
                },
                children: [
                    new TextRun({
                        size: 24,
                        text: `As demais cláusulas permanecem inalteradas.`,
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
                    before: 300,
                },
                children: [
                    new TextRun({
                        size: 24,
                        text: `E por estarem em pleno acordo, as partes contratantes assinam o presente TERMO DE RENOVAÇÃO DO CONTRATO DE TRABALHO, em duas vias, ficando a primeira em poder do `,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: `EMPREGADOR(A), `,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: `e a segunda com o(a) `,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: `EMPREGADO(A). `,
                        bold: true,
                        font: "Calibri Light",
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
                        font: "Calibri Light",
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
            column: {
                space: 790,
                count: 3,
            },
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

function createLastClausesConditional(){
    if(leadership){
        return [
            createClauseHeader('CLÁUSULA XII - DA VEDAÇÃO AO RECRUTAMENTO'),
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
                        text: '12.0 ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'O(A) ',
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'EMPREGADO(A), ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'está vedado de recrutar qualquer empregado do EMPREGADOR, mesmo após o término da vigência deste contrato.',
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
                        text: '12.1 ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'A vedação ao recrutamento perdurará pelo prazo de 60 (sessenta) dias contado da data de extinção deste contrato. Após esse período a presente cláusula perde sua vigência.',
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
                        text: '12.2 ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'O descumprimento desta cláusula poderá gerar a rescisão contratual, devendo o(a) ',
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'EMPREGADO(A), ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'pagar a multa no valor de R$5.000,00 (cinco mil reais), a ser devidamente atualizada e corrigida no momento de sua aplicação e, ainda, estará sujeito a eventuais penalidades civis e criminais eventualmente aplicáveis.',
                        font: "Calibri Light",
                    }),
                ],
            }),
            createClauseHeader('CLÁUSULA XIII - DA EXCLUSIVIDADE DO VÍNCULO EMPREGATÍCIO'),
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
                        text: '13.0 ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'Durante a vigência do presente instrumento, o(a) ',
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'EMPREGADO(A), ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'se compromete a manter a exclusividade do vínculo empregatício com o(a) ',
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'EMPREGADOR(A), ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'sendo-lhe vedada prestar serviços ou constituir quaisquer outros contratos de natureza trabalhista, com particulares ou com pessoas jurídicas.',
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
                        text: '13.1 ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'É vedado ao ',
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'EMPREGADO(A), ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'constituir empresa que preste os mesmos serviços e na mesma área geográfica onde são prestados os serviços comercializados pelo ',
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'EMPREGADOR(A). ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'Essa obrigação perdurará pelo prazo de vigência do contrato de trabalho se estendendo a 06 meses após a extinção desse.',
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
                        text: '13.2 ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'O descumprimento desta cláusula poderá gerar a rescisão contratual, devendo o(a) ',
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'EMPREGADO(A), ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'pagar multa no valor de R$10.000,00 (dez mil reais), a ser devidamente atualizada e corrigida no momento de sua aplicação, e, ainda, estará sujeito a eventuais penalidades civis e criminais eventualmente aplicáveis.',
                        font: "Calibri Light",
                    })
                ],
            }),
            createClauseHeader('CLÁUSULA XIV - DAS DISPOSIÇÕES GERAIS'),
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
                        text: '14.1 ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'Qualquer ato que configure concorrência ao empregador é considerado falta grave, sendo ao empregado vedado trabalhar em qualquer empresa da concorrência ou até mesmo por conta própria nas mesmas atividades do empregador durante a duração do contrato de trabalho.',
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
                        text: '14.2 ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'Ainda que não realizada no mesmo ramo da empresa, atividades paralelas que causem prejuízo ao cumprimento da função também poderão gerar, conforme o caso, demissão por justo motivo.',
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
                        text: '14.3 ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'Também incorre em justa causa aquele colaborador que se utiliza dos benefícios e prerrogativas dos funcionários para adquirir produtos/insumos ou afins e repassá-los a terceiros, sem o consentimento expresso da empregadora, independentemente de ter auferido ganhos para si ou para outrem.',
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
                        text: '14.4 ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'É vedado ao empregado(a), receber qualquer tipo de contraprestação, comissões, glosas, gorjetas de fornecedores, clientes ou terceiros, sem o consentimento expresso do empregador.',
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
                        text: '14.5 ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'Conforme dispõe a CLT em seu art. 2º, é o empregador quem dirige a prestação pessoal de serviço. Assim, fica estipulado que é expressamente proibida a utilização de celular ou de outro aparelho eletrônico que se assemelhe (tablet, smartwatch, etc) durante a jornada de trabalho para assuntos de cunho pessoal.',
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
                        text: '14.6 ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'Considerando que a EMPREGADORA disponibiliza para o EMPREGADO(A) os equipamentos de tecnologia da informação, e que estes são concedidos exclusivamente para atividades laborativas, poderá a EMPREGADORA fiscalizar e ter acesso a qualquer informação constante no referido equipamento, desde que armazenado no mesmo e independentemente da sua origem. ',
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
                        text: '14.7 ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'O(A) EMPREGADO(A) está ciente da realização de filmagem e captação de áudio no seu ambiente de trabalho, durante a sua jornada laboral, estando ciente que a respectiva filmagem visa trazer uma maior segurança em seu ambiente de trabalho.',
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
                        text: '14.7.1 ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'A EMPREGADORA não pagará nenhum valor pecuniário ao EMPREGADO pelo uso da sua imagem, sendo essa dentro da legislação civil, respeitados os locais íntimos.',
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
                        text: '14.8 ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'O empregado(a) autoriza o uso de sua imagem e voz, em todo e qualquer material entre fotos, documentos e outros meios de comunicação, para campanhas promocionais e institucionais. A presente autorização é concedida a título gratuito, abrangendo o uso da imagem acima mencionada em todo território nacional e no exterior, sob qualquer forma e meios, ou sejam, em destaque: (I) out-door; (II) bus-door; folhetos em geral (encartes, mala direta, catálogo, etc.); (III) folder de apresentação; (IV) anúncios em revistas e jornais em geral; (V) home page; (VI) cartazes; (VII) back-light; (VIII) mídia eletrônica (painéis, vídeo-tapes, televisão, cinema, programa para rádio, entre outros); (IX) redes e mídias sociais.',
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
                        size: 24,
                        text: 'E por estarem em pleno acordo, as partes contratantes assinam o presente Contrato de Experiência em duas vias, ficando a primeira em poder do EMPREGADOR, e a segunda com o(a) ',
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'EMPREGADO(A)',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                ],
            }),
        ]
    } else {
        return [
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
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'Fica o(a) ',
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'EMPREGADO(A), ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'comunicado e ciente que durante a permanência no local de trabalho está sendo monitorado por câmeras de segurança que possuem gravações de áudio e vídeo.',
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
                    before: 300,
                },
                children: [
                    new TextRun({
                        text: '12.2  ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'O(A) ',
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'EMPREGADOR(A), ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'poderá fiscalizar e ter acesso a qualquer informação constante nos softwares utilizados no ambiente laborativo, inclusive se utilizando de programas de monitoramento.',
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
                    before: 300,
                },
                children: [
                    new TextRun({
                        text: '12.3  ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'Conforme dispõe a CLT em seu art.2º, é o(a) ',
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'EMPREGADOR(A), ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'quem dirige a prestação pessoal de serviço. Assim, fica estipulado que é expressamente proibida a utilização de celular pessoal ou de outro aparelho eletrônico que se assemelhe (tablet, smartwatch, etc.) durante a jornada de trabalho para tratar de qualquer assunto de interesse pessoal do funcionário.',
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
                    before: 300,
                },
                children: [
                    new TextRun({
                        text: '12.4  ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'É vedado ao(a) ',
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'EMPREGADO(A), ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'receber qualquer tipo de contraprestação, comissões, glosas, gorjetas de fornecedores, clientes ou terceiros, sem o consentimento expresso do(a) ',
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'EMPREGADOR(A), ',
                        size: 24,
                        bold: true,
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
                    before: 300,
                },
                children: [
                    new TextRun({
                        text: '12.5  ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'O(A) ',
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'EMPREGADO(A), ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'autoriza o uso de sua imagem e voz, em todo e qualquer material entre fotos, documentos e outros meios de comunicação, para campanhas promocionais e institucionais. A presente autorização é concedida a título gratuito, abrangendo o uso da imagem acima mencionada em todo território nacional e no exterior, sob qualquer forma e meios, ou sejam, em destaque: (I) out-door; (II) bus-door; folhetos em geral (encartes, mala direta, catálogo, etc.); (III) folder de apresentação; (IV) anúncios em revistas e jornais em geral; (V) home page; (VI) cartazes; (VII) back-light; (VIII) mídia eletrônica (painéis, vídeo-tapes, televisão, cinema, programa para rádio, entre outros); (IX) redes e mídias sociais. ',
                        font: "Calibri Light",
                    }),
                ],
            }),
            ...createClause12Paragraph(user),
        ]
    }
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
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'Além disso, se obriga a realizar o que vier a ser objeto de  cartas, avisos ou ordens, dentro da natureza do seu cargo e também o que  dispensa especificações por estar naturalmente compreendido, subentendido ou  relacionado ao seu cargo, não constituindo a indicação supra ou a de adendos,  qualquer limitação ou restrição, considerando-se falta grave a recusa por parte  do(a) ',
                        size: 24,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'EMPREGADO(A) ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'em executar qualquer um dos serviços referidos, mesmo que  anteriormente não os tenha feito, mas que se entendam atinentes à função para  a qual fica contratado.',
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
                        text: '2.5 ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'A circunstância, porém, de ser a função especificada não importa na intransferibilidade do(a) ',
                        size: 24,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'EMPREGADO(A) ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'para outro serviço no qual demonstre melhor capacidade de adaptação desde que compatível com sua condição pessoal.  ',
                        font: "Calibri Light",
                    }),
                ],
            }),
            createClauseHeader('CLÁUSULA III - DA REMUNERAÇÃO E BENEFÍCIOS'),
            ...createPaymentContent(user, leadership),
            createClauseHeader('CLÁUSULA IV - DO HORÁRIO DE TRABALHO'),
            createShiftContent(user, leadership).then(console.log('concluiu shiftcontent')),  
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
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'O(A) ',
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'EMPREGADO(A) ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'autoriza o(a) ',
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'EMPREGADOR(A)',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'a descontar de sua folha de pagamento a contribuição sindical/confederativa/assistencial de sua categoria econômica ou profissional.',
                        font: "Calibri Light",
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
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'Caso se oponha ao desconto, o funcionário deve realizar sua manifestação nos moldes do instrumento coletivo, sendo que o desconto somente será suspenso após demonstrado que houve o cumprimento dos requisitos trazidos pela CCT (Convenção Coletiva de Trabalho).',
                        font: "Calibri Light",
                    }),
                ],
            }),
            createClauseHeader('CLÁUSULA VI - DO LOCAL DE TRABALHO'),
            ...createLaborPlace(user, leadership),
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
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'SÃO OBRIGAÇÕES DO ',
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'EMPREGADOR(A): ',
                        size: 24,
                        bold: true,
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
                        text: 'A. ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'O(A) EMPREGADOR(A) deverá pagar ao(a) EMPREGADO(A) os valores previstos na Cláusula Terceira, dentro do prazo e da forma previamente indicada, a título salarial Light;',
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
                        text: 'B. ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'O(A) ',
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'EMPREGADOR(A) ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'deverá fornecer todas as condições para que o(a) ',
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'EMPREGADO(A) ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'labore em ambiente de trabalho seguro, com boas condições sanitárias e com infraestrutura adequada à execução das atividades pelo(a) ',
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'EMPREGADO(A).',
                        size: 24,
                        bold: true,
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
                        text: 'C. ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'O(A) ',
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'EMPREGADOR(A), ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'no ato de celebração deste contrato, deverá cientificar o(a) ',
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'EMPREGADO(A) ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'de todas as regras de conduta estabelecidas e políticas internas, devendo entregar uma cópia do regulamento interno;',
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
                    before: 300,
                },
                children: [
                    new TextRun({
                        text: '7.2 ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'SÃO OBRIGAÇÕES DO ',
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'EMPREGADO(A): ',
                        size: 24,
                        bold: true,
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
                        text: 'A. ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'O(A) ',
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'EMPREGADO(A) ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'se compromete a executar as funções objeto do presente contrato, conforme as exigências, diretrizes e padrões exigidos pelo(a) ',
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'EMPREGADOR(A) ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'bem como realizá-las com empenho para o melhor desenvolvimento do trabalho, preservando a qualidade e os prazos pactuados;',
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
                        text: 'B. ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'O(A) ',
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'EMPREGADO(A) ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'se compromete a prestar ao(à) ',
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'EMPREGADOR(A) ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'as informações necessárias sobre o andamento das atividades desenvolvidas;',
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
                        text: 'C. ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'O(A) ',
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'EMPREGADO(A) ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'deverá manter durante toda vigência deste contrato, comportamento compatível com as normas de disciplina, da ética profissional e de segurança estabelecidas pela legislação brasileira e pelas normas internas do(a) ',
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'EMPREGADOR(A) ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'declarando estar ciente dos seus termos e condições;',
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
                        text: 'D. ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'O(A) ',
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'EMPREGADO(A) ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'se compromete a utilizar adequadamente os equipamentos, uniformes e materiais fornecidos pelo(a) ',
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'EMPREGADOR(A) ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'os quais devem ser utilizados apenas para os fins profissionais contratados, podendo o(a) ',
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'EMPREGADOR(A) ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'realizar vistorias periódicas nos equipamentos por ele fornecido, desde a verificação de e-mails corporativos até a delimitação do recebimento e envio de arquivos;',
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
                        text: 'E. ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'O(A) ',
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'EMPREGADO(A) ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'assume estar ciente de que todos os códigos e senhas, chaves e contas de monitoramento fornecidos pelo(a) ',
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'EMPREGADOR(A) ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'para utilização dos equipamentos são estritamente confidenciais, devendo ele tomar todas as cautelas na sua guarda.',
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
                        text: 'F. ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'Obriga-se o(a) ',
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'EMPREGADO(A) ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'além de executar com dedicação e lealdade o seu serviço, a cumprir o Regulamento Interno do(a) ',
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'EMPREGADOR(A), ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'as instruções de sua administração e as ordens de seus chefes e superiores hierárquicos, relativas às peculiaridades dos serviços que lhe forem confiados.',
                        font: "Calibri Light",
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
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'O presente contrato entrará em vigor a partir da sua assinatura, na condição de CONTRATO DE EXPERIÊNCIA, e terá vigência de 45 (quarenta e cinco dias), renováveis por igual período automaticamente.',
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
                        text: '8.2 ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'Ao final do prazo de vigência previsto na cláusula anterior, o contrato poderá ser rescindido; caso contrário, será tacitamente convertido em contrato de trabalho por tempo indeterminado, sendo mantidas todas as demais cláusulas e obrigações aqui estabelecidas.',
                        font: "Calibri Light",
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
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'O(A) ',
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'EMPREGADO(A) ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'deverá manter em sigilo, durante a vigência do presente termo e mesmo após sua extinção, sobre qualquer informação relativa aos negócios, políticas, segredos institucionais, organização, criação, lista de clientes, quadro de funcionários, faturamento, metas e comissões, bem como outras informações relativas ao(à) ',
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'EMPREGADOR(A), ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'dados de seus clientes, fornecedores, representantes, demais ',
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'EMPREGADO(A)S ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'ou classificadas como confidenciais.',
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
                        text: '9.2 ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'Para fins do presente contrato, entende-se por informação confidencial:',
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
                },
                children: [
                    new TextRun({
                        text: 'A. ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'Qualquer informação relacionada ao negócio e operações do(a) ',
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'EMPREGADOR(A) ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'que não sejam públicas',
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
                        text: 'B. ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'Informações contidas em pesquisas, faturamento, metas, comissões, planos de negócio, venda ou marketing, informações financeiras, custos, dados de precificação, parceiros de negócios, informações de fornecedores e clientes, propriedade intelectual, especificações, expertises, técnicas, invenções e todos os métodos, conceitos ou ideias relacionadas ao negócio do ',
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'EMPREGADOR(A) ',
                        size: 24,
                        bold: true,
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
                        text: 'C. ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'É vedado ao(a)  ',
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'EMPREGADO(A) ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'repassar a terceiros, sejam particulares ou pessoas jurídicas, quaisquer destas informações, exceto quando expressamente autorizado pelo(a) ',
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'EMPREGADOR(A) ',
                        size: 24,
                        bold: true,
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
                        text: 'D. ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'A confidencialidade dessas informações independe de aviso prévio do(a) ',
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'EMPREGADOR(A) ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'devendo o(a)',
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'EMPREGADO(A) ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'considerar toda e qualquer informação relacionada ao negócio do(a) ',
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'EMPREGADOR(A) ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'e dos serviços prestado em sede dele como confidencial.',
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
                        text: 'E. ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'Ressalta-se que o dever de confidencialidade permanece mesmo após o término deste contrato de trabalho.',
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
                        text: 'F. ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'A violação da obrigação de confidencialidade pode causar a rescisão imediata deste contrato por justa causa, conforme o artigo 482, alínea g da CLT.',
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
                        text: 'G. ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'Em caso de violação desta cláusula o(a) ',
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'EMPREGADO(A), ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'poderá ser responsabilizado pelo pagamento das quantias equivalentes ao dano causado e estará sujeito ao pagamento de multa no valor de 5.000,00 (cinco mil reais), a ser devidamente atualizada e corrigidas no momento de sua aplicação e, ainda, estará sujeito a eventuais penalidades civis e criminais eventualmente aplicáveis. ',
                        font: "Calibri Light",
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
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'O(A) ',
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'EMPREGADO(A), ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'declara estar ciente de que todo e qualquer direito advindo ou relacionado ao trabalho por ele(a) desempenhado, direta ou indiretamente, com os serviços prestados em decorrência do presente contrato, pertencerão exclusivamente ao(à) ',
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'EMPREGADOR(A), ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'nos termos da legislação vigente.',
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
                        text: '10.2  ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'Nesse ponto, também é objeto do presente contrato a cessão e transferência em favor do(a) ',
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'EMPREGADOR(A), ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'expressamente, na integralidade, a título universal e gratuito, em caráter irretratável e irrevogável, para fins de utilização a qualquer tempo, para fins de utilização econômica ou não, no Brasil e/ou no Exterior, de todos os direitos patrimoniais de autoria sobre documentos de modo geral referente às Obras que já tenham sido ou ainda sejam criadas pelo(a) ',
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'EMPREGADO(A) ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'no âmbito da relação de trabalho com o(a) ',
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'EMPREGADOR(A) ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'abrangendo tal cessão a criação, aperfeiçoamento, redação, revisão, edição, tradução, adaptação e toda e qualquer atividade que enseje proteção de direito de autor com relação às referidas Obras, que decorra, direta ou indiretamente, das atividades exercidas pelo(a) ',
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'EMPREGADO(A) ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'em razão da relação mantida com ',
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'EMPREGADOR(A).',
                        size: 24,
                        bold: true,
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
                        text: '10.3  ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'O disposto na Cláusula acima tem validade por todo o tempo em que a obra estiver protegida por direitos autorais.',
                        font: "Calibri Light",
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
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'O(A) ',
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'EMPREGADO(A), ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'declara estar cientes e de comum acordo com o registro digital da carteira de trabalho, mesmo que tenha sido fornecida uma (CTPS) física para recolhimento de dados no ato da sua contratação',
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
                        text: '11.2 ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'Qualquer anotação, alteração ou modificação referente ao vínculo empregatício será realizada na versão digital da carteira de trabalho do(a) ',
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'EMPREGADO(A), ',
                        size: 24,
                        bold: true,
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
                        text: '11.3 ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'O(A) ',
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'EMPREGADO(A), ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'assume responsabilidade de manter-se informado(a) sobre as atualizações realizadas na carteira de trabalho digital, garantindo o acompanhamento de todas as anotações pertinentes ao seu histórico profissional.',
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
                        text: '11.4 ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'O(A) ',
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        text: 'EMPREGADO(A), ',
                        size: 24,
                        bold: true,
                        font: "Calibri Light",
                    }),
                    new TextRun({
                        size: 24,
                        text: 'está ciente que as implicações e os benefícios decorrentes do registro digital da minha carteira de trabalho, e concorda em utilizar este meio como forma de registro e documentação do seu histórico laboral.',
                        font: "Calibri Light",
                    }),
                ],
            }), 
            ...createLastClausesConditional(),
            ...createTimeStamp(user),
        ],
    }
    doc.addSection(section);
}

// ARQUIVO DE AJUDANTE DE OBRA
async function getSkillExams(id){
    try {
        const response = await axios.get(`https://3337-avanciconstru-apiserver-liom39gdw1s.ws-us114.gitpod.io/skillclinic?skill_id=${id}`, {
            headers: {
            'Authorization': `Bearer eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpYXQiOjE3MTgwNTA5NjksImV4cCI6MTc2OTg5MDk2OSwic3ViIjoiZjM1ZDg2M2QtMmI4My00MGM4LWI4ZDUtM2ExNzU5YTU2NTc2In0.RVOvNkHAqsTmp_iRfvu28sbeMbwfauCA4E7_7fWpp78`
        }});
        return response.data;
    } catch (error) {
        console.error('erro no get skill exam')
        console.log(error.data);
        return null;
    }
}

async function createExamsParagraphs(user){
    const response = await getSkillExams(user.skill.id);
    console.log('response', response);
    var paragraphs = [];
    response.forEach((exam, index)=> {
        paragraphs.push( 
            new Paragraph({
            children: [
                new TextRun({
                    text: `${index + 1}. ${exam.clinical_examination.name.toUpperCase()}`,
                    size: 22,
                    font: "Calibri",
                }),
            ],
        }))
    })
    console.log('paragraphs',paragraphs);
    return paragraphs;
}

function createEnvironmentsParagraphs(user){
    var paragraphs = [];
    user.skill.environments.forEach((env)=> {
        console.log(env);
        paragraphs.push( 
            new Paragraph({
            children: [
                new TextRun({
                    text: `- ${env}`,
                    size: 22,
                    font: "Calibri",
                }),
            ],
        }))
    })
    console.log('paragraphs',paragraphs);
    return paragraphs;
}

async function createPersonalDataTable(company, user, doc, contractType){
    const examParagraphs = await createExamsParagraphs(user);
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
        children:[
            new Table({
                columnWidths: [5000, 5000],
                width: {
                    size: 10000,
                    type: WidthType.DXA,
                },
                rows: [
                    new TableRow({
                        cantSplit: true,
                        tableHeader: true,
                        children: [
                            new TableCell({
                                margins: {
                                    top: convertInchesToTwip(0.1),
                                    bottom: convertInchesToTwip(0.1),
                                    left: convertInchesToTwip(0.1),
                                    right: convertInchesToTwip(0.1),
                                },
                                tableHeader: true,
                                height: {value: 50, type: WidthType},
                                columnSpan: 2,
                                children: [
                                    new Paragraph({
                                        children: [
                                            new TextRun({
                                                text: "RAZÃO SOCIAL: ",
                                                bold: true,
                                                size: 22,
                                                font: "Calibri",
                                            }),
                                            new TextRun({
                                                text: `${company==='HABIT' ? 'HABIT CONSTRUÇÕES E SERVIÇOS EIRELLI' : 'AVANCI'}`,
                                                size: 22,
                                                font: "Calibri",
                                            })
                                        ],
                                    }),
                                    new Paragraph({
                                        children: [
                                            new TextRun({
                                                text: "CNPJ: ",
                                                bold: true,
                                                size: 22,
                                                font: "Calibri",
                                            }),
                                            new TextRun({
                                                text: `${company==='HABIT' ? '28.697.934/0001-43' : 'AVANCI'}`,
                                                size: 22,
                                                font: "Calibri",
                                            })
                                        ],
                                    }),
                                ],
                            }),
                        ],
                    }),
                    new TableRow({
                        cantSplit: true,
                        tableHeader: true,
                        children: [
                            new TableCell({
                                margins:{
                                    left: convertInchesToTwip(0.2),
                                    right: convertInchesToTwip(0.2),
                                },
                                tableHeader: true,
                                height: {value: 50, type: WidthType},
                                borders: {
                                    bottom: {
                                        color: "ffffff",
                                    },
                                },
                                children: [
                                    new Paragraph({
                                        alignment: AlignmentType.CENTER,
                                        children: [
                                            new TextRun({
                                                text: "Responsável pelo encaminhamento: LARISSA K. RAMOS MAGALHÃES ",
                                                bold: true,
                                                size: 22,
                                                font: "Calibri",
                                            }),
                                        ],
                                    })
                                ],
                            }),
                            new TableCell({
                                margins:{
                                    left: convertInchesToTwip(0.2),
                                    right: convertInchesToTwip(0.2),
                                },
                                tableHeader: true,
                                height: {value: 50, type: WidthType},
                                borders: {
                                    bottom: {
                                        color: "ffffff",
                                    },
                                },
                                children: [
                                    new Paragraph({
                                        alignment: AlignmentType.CENTER,
                                        children: [
                                            new TextRun({
                                                text: `Data: ${moment().format('DD/MM/YYYY')}`,
                                                bold: true,
                                                size: 22,
                                                font: "Calibri",
                                            }),
                                        ],
                                    })
                                ],
                            }),
                        ],
                    }),
                    new TableRow({
                        cantSplit: true,
                        tableHeader: true,
                        children: [
                            new TableCell({
                                margins:{
                                    left: convertInchesToTwip(0.2),
                                    right: convertInchesToTwip(0.2),
                                },
                                tableHeader: true,
                                height: {value: 50, type: WidthType},
                                borders: {
                                    bottom: {
                                        color: "ffffff",
                                    },
                                },
                                columnSpan: 2,
                                children: [
                                    new Paragraph({
                                        alignment: AlignmentType.CENTER,
                                        children: [
                                            new TextRun({
                                                text: `ATENÇÃO: Para exames de sangue estar em jejum`,
                                                size: 22,
                                                font: "Calibri",
                                            }),
                                        ],
                                    }),
                                    new Paragraph({
                                        alignment: AlignmentType.CENTER,
                                        children: [
                                            new TextRun({
                                                text: `LEVAR DOCUMENTO COM FOTO`,
                                                size: 22,
                                                font: "Calibri",
                                            }),
                                        ],
                                    })
                                ],
                            }),
                        ],
                    }),
                    new TableRow({
                        cantSplit: true,
                        tableHeader: true,
                        children: [
                            new TableCell({
                                margins:{
                                    left: convertInchesToTwip(0.2),
                                    right: convertInchesToTwip(0.2),
                                },
                                tableHeader: true,
                                height: {value: 50, type: WidthType},
                                columnSpan: 2,
                                children: [
                                    new Paragraph({
                                        alignment: AlignmentType.CENTER,
                                        children: [
                                            new TextRun({
                                                text: `DADOS PESSOAIS`,
                                                bold: true,
                                                size: 22,
                                                font: "Calibri",
                                            }),
                                        ],
                                    }),
                                ],
                            }),
                        ],
                    }),
                ],
            }),
            new Paragraph({text: ''}),
            new Table({
                columnWidths: [5000, 5000],
                width: {
                    size: 10000,
                    type: WidthType.DXA,
                },
                rows: [
                    new TableRow({
                        cantSplit: true,
                        tableHeader: true,
                        children: [
                            new TableCell({
                                margins:{
                                    left: convertInchesToTwip(0.2),
                                    right: convertInchesToTwip(0.2),
                                },
                                tableHeader: true,
                                height: {value: 50, type: WidthType},
                                borders: {
                                    bottom: {
                                        color: "ffffff",
                                    },
                                },
                                columnSpan: 4,
                                children: [
                                    new Paragraph({
                                        children: [
                                            new TextRun({
                                                text: "NOME COMPLETO: ",
                                                bold: true,
                                                size: 22,
                                                font: "Calibri",
                                            }),
                                            new TextRun({
                                                text: `${user.name}`,
                                                size: 22,
                                                font: "Calibri",
                                            })
                                        ],
                                    }),
                                ],
                            }),
                            new TableCell({
                                margins:{
                                    left: convertInchesToTwip(0.2),
                                    right: convertInchesToTwip(0.2),
                                },
                                tableHeader: true,
                                height: {
                                    value: 50, 
                                    type: WidthType
                                },
                                borders: {
                                    bottom: {
                                        color: "ffffff",
                                    },
                                },
                                columnSpan: 2,
                                children: [
                                    new Paragraph({
                                        children: [
                                            new TextRun({
                                                text: "RG: ",
                                                bold: true,
                                                size: 22,
                                                font: "Calibri",
                                            }),
                                            new TextRun({
                                                text: `${user.rg ? user.rg : ''}`,
                                                size: 22,
                                                font: "Calibri",
                                            })
                                        ],
                                    }),
                                    new Paragraph({
                                        children: [
                                            new TextRun({
                                                text: "CPF: ",
                                                bold: true,
                                                size: 22,
                                                font: "Calibri",
                                            }),
                                            new TextRun({
                                                text: `${user.cpf}`,
                                                size: 22,
                                                font: "Calibri",
                                            })
                                        ],
                                    }),
                                    new Paragraph({
                                        children: [
                                            new TextRun({
                                                text: "DATA DE NASCIMENTO ",
                                                bold: true,
                                                size: 22,
                                                font: "Calibri",
                                            }),
                                            new TextRun({
                                                text: `${moment(user.birthday).format('DD/MM/YYYY')}`,
                                                size: 22,
                                                font: "Calibri",
                                            })
                                        ],
                                    }),
                                ],
                            }),
                        ],
                    }),
                    new TableRow({
                        cantSplit: true,
                        tableHeader: true,
                        children: [
                            new TableCell({
                                margins:{
                                    left: convertInchesToTwip(0.2),
                                    right: convertInchesToTwip(0.2),
                                },
                                tableHeader: true,
                                height: {value: 50, type: WidthType},
                                borders: {
                                    bottom: {
                                        color: "ffffff",
                                    },
                                },
                                columnSpan: 6,
                                children: [
                                    new Paragraph({
                                        children: [
                                            new TextRun({
                                                text: "Cargo: ",
                                                bold: true,
                                                size: 22,
                                                font: "Calibri",
                                            }),
                                            new TextRun({
                                                text: `${user.skill.name} - CBO: ${user.skill.cbo}`,
                                                size: 22,
                                                font: "Calibri",
                                            })
                                        ],
                                    }),
                                ],
                            }),
                        ],
                    }),
                    new TableRow({
                        cantSplit: true,
                        tableHeader: true,
                        children: [
                            new TableCell({
                                margins:{
                                    left: convertInchesToTwip(0.2),
                                    right: convertInchesToTwip(0.2),
                                },
                                tableHeader: true,
                                height: {value: 50, type: WidthType},
                                borders: {
                                    bottom: {
                                        color: "ffffff",
                                    },
                                },
                                columnSpan: 6,
                                children: [
                                    new Paragraph({
                                        children: [
                                            new TextRun({
                                                text: "Ambientes: ",
                                                bold: true,
                                                size: 22,
                                                font: "Calibri",
                                            }),
                                            ...createEnvironmentsParagraphs(user),
                                        ],
                                    }),
                                ],
                            }),
                        ],
                    }),
                    new TableRow({
                        cantSplit: true,
                        tableHeader: true,
                        children: [
                            new TableCell({
                                margins:{
                                    left: convertInchesToTwip(0.2),
                                    right: convertInchesToTwip(0.2),
                                },
                                tableHeader: true,
                                height: {value: 50, type: WidthType},
                                borders: {
                                    bottom: {
                                        color: "ffffff",
                                    },
                                },
                                children: [
                                    new Paragraph({
                                        alignment: AlignmentType.CENTER,
                                        children: [
                                            new TextRun({
                                                text: "Admissional: ",
                                                size: 22,
                                                font: "Calibri",
                                            }),
                                        ],
                                    }),
                                ],
                            }),
                            new TableCell({
                                margins:{
                                    left: convertInchesToTwip(0.2),
                                    right: convertInchesToTwip(0.2),
                                },
                                tableHeader: true,
                                height: {value: 50, type: WidthType},
                                borders: {
                                    bottom: {
                                        color: "ffffff",
                                    },
                                },
                                children: [],
                            }),
                            new TableCell({
                                margins:{
                                    left: convertInchesToTwip(0.2),
                                    right: convertInchesToTwip(0.2),
                                },
                                tableHeader: true,
                                height: {value: 50, type: WidthType},
                                borders: {
                                    bottom: {
                                        color: "ffffff",
                                    },
                                },
                                children: [
                                    new Paragraph({
                                        alignment: AlignmentType.CENTER,
                                        children: [
                                            new TextRun({
                                                text: "Mudança de riscos: ",
                                                size: 22,
                                                font: "Calibri",
                                            }),
                                        ],
                                    }),
                                ],
                            }),
                            new TableCell({
                                margins:{
                                    left: convertInchesToTwip(0.2),
                                    right: convertInchesToTwip(0.2),
                                },
                                tableHeader: true,
                                height: {value: 50, type: WidthType},
                                borders: {
                                    bottom: {
                                        color: "ffffff",
                                    },
                                },
                                children: [],
                            }),
                            new TableCell({
                                margins:{
                                    left: convertInchesToTwip(0.2),
                                    right: convertInchesToTwip(0.2),
                                },
                                tableHeader: true,
                                height: {value: 50, type: WidthType},
                                borders: {
                                    bottom: {
                                        color: "ffffff",
                                    },
                                },
                                children: [
                                    new Paragraph({
                                        alignment: AlignmentType.CENTER,
                                        children: [
                                            new TextRun({
                                                text: "Periódico",
                                                size: 22,
                                                font: "Calibri",
                                            }),
                                        ],
                                    }),
                                ],
                            }),
                            new TableCell({
                                margins:{
                                    left: convertInchesToTwip(0.2),
                                    right: convertInchesToTwip(0.2),
                                },
                                tableHeader: true,
                                height: {value: 50, type: WidthType},
                                borders: {
                                    bottom: {
                                        color: "ffffff",
                                    },
                                },
                                children: [
                                    new Paragraph({
                                        alignment: AlignmentType.CENTER,
                                        children: [
                                            new TextRun({
                                                text: "Retorno ao trabalho",
                                                size: 22,
                                                font: "Calibri",
                                            }),
                                        ],
                                    }),
                                ],
                            }),
                        ],
                    }),
                    new TableRow({
                        cantSplit: true,
                        tableHeader: true,
                        children: [
                            new TableCell({
                                margins:{
                                    left: convertInchesToTwip(0.2),
                                    right: convertInchesToTwip(0.2),
                                },
                                tableHeader: true,
                                height: {value: 50, type: WidthType},
                                children: [
                                    new Paragraph({
                                        alignment: AlignmentType.CENTER,
                                        children: [
                                            new TextRun({
                                                text: `${contractType === 'Admission' ? 'X' : ''}`,
                                                size: 22,
                                                font: "Calibri",
                                            }),
                                        ],
                                    }),
                                ],
                            }),
                            new TableCell({
                                margins:{
                                    left: convertInchesToTwip(0.2),
                                    right: convertInchesToTwip(0.2),
                                },
                                tableHeader: true,
                                height: {value: 50, type: WidthType},
                                children: [],
                            }),
                            new TableCell({
                                margins:{
                                    left: convertInchesToTwip(0.2),
                                    right: convertInchesToTwip(0.2),
                                },
                                tableHeader: true,
                                height: {value: 50, type: WidthType},
                                children: [
                                    new Paragraph({
                                        alignment: AlignmentType.CENTER,
                                        children: [
                                            new TextRun({
                                                text: `${contractType === 'RiskChange' ? 'X' : ''}`,
                                                size: 22,
                                                font: "Calibri",
                                            }),
                                        ],
                                    }),
                                ],
                            }),
                            new TableCell({
                                margins:{
                                    left: convertInchesToTwip(0.2),
                                    right: convertInchesToTwip(0.2),
                                },
                                tableHeader: true,
                                height: {value: 50, type: WidthType},
                                children: [],
                            }),
                            new TableCell({
                                tableHeader: true,
                                height: {value: 50, type: WidthType},
                                children: [
                                    new Paragraph({
                                        alignment: AlignmentType.CENTER,
                                        children: [
                                            new TextRun({
                                                text: `${contractType === 'Periodical' ? 'X' : ''}`,
                                                size: 22,
                                                font: "Calibri",
                                            }),
                                        ],
                                    }),
                                ],
                            }),
                            new TableCell({
                                tableHeader: true,
                                height: {value: 50, type: WidthType},
                                children: [
                                    new Paragraph({
                                        alignment: AlignmentType.CENTER,
                                        children: [
                                            new TextRun({
                                                text: `${contractType === 'Return' ? 'X' : ''}`,
                                                size: 22,
                                                font: "Calibri",
                                            }),
                                        ],
                                    }),
                                ],
                            }),
                        ],
                    }),
                    new TableRow({
                        cantSplit: true,
                        tableHeader: true,
                        children: [
                            new TableCell({
                                margins:{
                                    left: convertInchesToTwip(0.2),
                                    right: convertInchesToTwip(0.2),
                                },
                                tableHeader: true,
                                height: {value: 50, type: WidthType},
                                columnSpan: 6,
                                children: [
                                    new Paragraph({
                                        alignment: AlignmentType.CENTER,
                                        children: [
                                            new TextRun({
                                                text: "EXAMES COMPLEMENTARES A REALIZAR:",
                                                bold: true,
                                                size: 22,
                                                font: "Calibri",
                                            }),
                                        ],
                                    }),
                                ],
                            }),
                        ],
                    }),
                    new TableRow({
                        cantSplit: true,
                        tableHeader: true,
                        children: [
                            new TableCell({
                                margins:{
                                    left: convertInchesToTwip(0.2),
                                    right: convertInchesToTwip(0.2),
                                },
                                tableHeader: true,
                                height: {value: 50, type: WidthType},
                                columnSpan: 4,
                                children: examParagraphs,
                            }),
                            new TableCell({
                                margins:{
                                    left: convertInchesToTwip(0.2),
                                    right: convertInchesToTwip(0.2),
                                },
                                tableHeader: true,
                                height: {
                                    value: 50, 
                                    type: WidthType
                                },
                                columnSpan: 2,
                                children: [],
                            }),
                        ],
                    }),
                ],
            })
        ]
    }
    doc.addSection(section);
}

module.exports = {createCompanyTextRun, createEmployeeTextRun, createAttributionsSection, createClauseWithParagraphs, createEpiProductsTable, getSkillProducts, createEpiTable, countDays, createXPContract, createClauseHeader, createSignSection, createRenew, createNewSection, createPageHeader, createPageFooter, createPaymentContent, createPersonalDataTable, createExamsParagraphs, createEnvironmentsParagraphs}