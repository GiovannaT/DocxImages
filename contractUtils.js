const axios = require('axios')
const moment = require("moment")
const { Paragraph, Table, TextRun, TableRow, TableCell, ShadingType, AlignmentType, WidthType} = require("docx");

function createClauseWithParagraphs(number, objective) {
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

async function getSkillProducts (id) {
    try {
        const response = await axios.get(`https://3337-avanciconstru-apiserver-0ae2jz7xl1m.ws-us110.gitpod.io/skillproducts?skill_id=${id}`, {
            headers: {
            'Authorization': `Bearer eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpYXQiOjE3MTUyNjA5OTMsImV4cCI6MTc2NzEwMDk5Mywic3ViIjoiZjM1ZDg2M2QtMmI4My00MGM4LWI4ZDUtM2ExNzU5YTU2NTc2In0.EepQV0WVRYNQLQp2sPDhJU_Cm34c2rDFPuq0I_fqpDI`
        }});
        return response.data;
    } catch (error) {
        // console.error('erro no get skill product')
        // console.log(error)
        return null;
    }
}

async function createEpiTable(doc, executor){
    if(executor.executor){
        var tableRows = [];
        var productRows = [];
        // try{
        //     const skillProducts = await getSkillProducts(executor.executor.skill_id);
        //     if(skillProducts && skillProducts.length > 0){
        //         skillProducts.forEach(product => {
        //             productRows.push(
        //                 new TableRow({
        //                     children: [
        //                         new TableCell({
        //                             height: {value: 50, type: WidthType},
        //                             children: [
        //                                 new Paragraph({
        //                                     alignment: AlignmentType.CENTER,
        //                                     text: `1`,
        //                                     bold: true,
        //                                     }),
        //                             ],
        //                         }),
        //                         new TableCell({
        //                             height: {value: 50, type: WidthType},
        //                             children: [
        //                                 new Paragraph({
        //                                     alignment: AlignmentType.CENTER,
        //                                     text: `${product?.product?.description}.`,
        //                                     bold: true,
        //                                     }),
        //                             ],
        //                         }),
        //                         new TableCell({
        //                             height: {value: 50, type: WidthType},
        //                             children: [],
        //                         }),
        //                         new TableCell({
        //                             height: {value: 50, type: WidthType},
        //                             children: [
        //                                 new Paragraph({
        //                                     alignment: AlignmentType.CENTER,
        //                                     text: `${moment().format('DD/MM/YYYY')}.`,
        //                                     bold: true,
        //                                 }),
        //                             ],
        //                         }),
        //                         new TableCell({
        //                             height: {value: 50, type: WidthType},
        //                             children: [],
        //                         }),
        //                         new TableCell({
        //                             height: {value: 50, type: WidthType},
        //                             children: [],
        //                         }),
        //                         new TableCell({
        //                             height: {value: 50, type: WidthType},
        //                             children: [],
        //                         }),
        //                     ]
        //                 }),
        //             )
        //         })
        //     }    
        // } catch (error){
        //     console.log(error);
        // }
        console.log(...productRows);
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
                            pageBreakBefore: true,
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
            // new TableRow({
            //     children:[
            //         new TableCell({
            //             borders: {
            //                 top: {
            //                     color: "ffffff",
            //                 },
            //                 bottom: {
            //                     color: "ffffff",
            //                 },
            //             },
            //             columnSpan: 2,
            //             children:[
            //                 new Table({
            //                     columnWidths: [3505, 5505],
            //                     width: {
            //                         size: 9000,
            //                         type: WidthType.DXA,
            //                     },
            //                     rows: [
            //                         new TableRow({
            //                             cantSplit: true,
            //                             tableHeader: true,
            //                             children: [
            //                                 new TableCell({
            //                                     tableHeader: true,
            //                                     height: {value: 50, type: WidthType},
            //                                     borders: {
            //                                         bottom: {
            //                                             color: "ffffff",
            //                                         },
            //                                     },
            //                                     shading: {
            //                                         fill: "D9D9D9",
            //                                         type: ShadingType.CLEAR,
            //                                         color: "auto",
            //                                     },
            //                                     children: [
            //                                         new Paragraph({
            //                                         alignment: AlignmentType.CENTER,
            //                                         text: `QUANT.`,
            //                                         bold: true,
            //                                         }),
            //                                     ],
            //                                 }),
            //                                 new TableCell({
            //                                     tableHeader: true,
            //                                     height: {value: 50, type: WidthType},
            //                                     shading: {
            //                                         fill: "D9D9D9",
            //                                         type: ShadingType.CLEAR,
            //                                         color: "auto",
            //                                     },
            //                                     borders: {
            //                                         bottom: {
            //                                             color: "ffffff",
            //                                         },
            //                                     },
            //                                     children: [new Paragraph({
            //                                         alignment: AlignmentType.CENTER,
            //                                         text: `TIPO DE EPI`,
            //                                         bold: true,
            //                                     }),],
            //                                 }),
            //                                 new TableCell({
            //                                     tableHeader: true,
            //                                     height: {value: 50, type: WidthType},
            //                                     shading: {
            //                                         fill: "D9D9D9",
            //                                         type: ShadingType.CLEAR,
            //                                         color: "auto",
            //                                     },
            //                                     borders: {
            //                                         bottom: {
            //                                             color: "ffffff",
            //                                         },
            //                                     },
            //                                     children: [new Paragraph({
            //                                         alignment: AlignmentType.CENTER,
            //                                         text: `Nº C.A.`,
            //                                         bold: true,
            //                                     }),],
            //                                 }),
            //                                 new TableCell({
            //                                     tableHeader: true,
            //                                     height: {value: 50, type: WidthType},
            //                                     shading: {
            //                                         fill: "D9D9D9",
            //                                         type: ShadingType.CLEAR,
            //                                         color: "auto",
            //                                     },
            //                                     borders: {
            //                                         bottom: {
            //                                             color: "ffffff",
            //                                         },
            //                                     },
            //                                     children: [new Paragraph({
            //                                         alignment: AlignmentType.CENTER,
            //                                         text: `Data entrega`,
            //                                         bold: true,
            //                                     }),],
            //                                 }),
            //                                 new TableCell({
            //                                     tableHeader: true,
            //                                     height: {value: 50, type: WidthType},
            //                                     shading: {
            //                                         fill: "D9D9D9",
            //                                         type: ShadingType.CLEAR,
            //                                         color: "auto",
            //                                     },
            //                                     borders: {
            //                                         bottom: {
            //                                             color: "ffffff",
            //                                         },
            //                                     },
            //                                     children: [new Paragraph({
            //                                         alignment: AlignmentType.CENTER,
            //                                         text: `Data devolução`,
            //                                         bold: true,
            //                                     }),],
            //                                 }),
            //                                 new TableCell({
            //                                     tableHeader: true,
            //                                     height: {value: 50, type: WidthType},
            //                                     shading: {
            //                                         fill: "D9D9D9",
            //                                         type: ShadingType.CLEAR,
            //                                         color: "auto",
            //                                     },
            //                                     children: [new Paragraph({
            //                                         alignment: AlignmentType.CENTER,
            //                                         text: `Motivo`,
            //                                         bold: true,
            //                                     }),],
            //                                 }),
            //                                 new TableCell({
            //                                     tableHeader: true,
            //                                     height: {value: 50, type: WidthType},
            //                                     shading: {
            //                                         fill: "D9D9D9",
            //                                         type: ShadingType.CLEAR,
            //                                         color: "auto",
            //                                     },
            //                                     columnSpan: 2,
            //                                     borders: {
            //                                         bottom: {
            //                                             color: "ffffff",
            //                                         },
            //                                     },
            //                                     children: [new Paragraph({
            //                                         alignment: AlignmentType.CENTER,
            //                                         text: `Nº C.A. Novo Equipamento`,
            //                                         bold: true,
            //                                     }),],
            //                                 }),
            //                             ],
            //                         }),
            //                         new TableRow({
            //                             children: [
            //                                 new TableCell({
            //                                     height: {value: 50, type: WidthType},
            //                                     shading: {
            //                                         fill: "D9D9D9",
            //                                         type: ShadingType.CLEAR,
            //                                         color: "auto",
            //                                     },
            //                                     borders: {
            //                                         top: {
            //                                             color: "ffffff",
            //                                         },
            //                                         bottom: {
            //                                             color: "ffffff",
            //                                         },
            //                                     },
            //                                     children: [],
            //                                 }),
            //                                 new TableCell({
            //                                     height: {value: 50, type: WidthType},
            //                                     shading: {
            //                                         fill: "D9D9D9",
            //                                         type: ShadingType.CLEAR,
            //                                         color: "auto",
            //                                     },
            //                                     borders: {
            //                                         top: {
            //                                             color: "ffffff",
            //                                         },
            //                                         bottom: {
            //                                             color: "ffffff",
            //                                         },
            //                                     },
            //                                     children: [],
            //                                 }),
            //                                 new TableCell({
            //                                     height: {value: 50, type: WidthType},
            //                                     shading: {
            //                                         fill: "D9D9D9",
            //                                         type: ShadingType.CLEAR,
            //                                         color: "auto",
            //                                     },
            //                                     borders: {
            //                                         top: {
            //                                             color: "ffffff",
            //                                         },
            //                                         bottom: {
            //                                             color: "ffffff",
            //                                         },
            //                                     },
            //                                     children: [],
            //                                 }),
            //                                 new TableCell({
            //                                     height: {value: 50, type: WidthType},
            //                                     shading: {
            //                                         fill: "D9D9D9",
            //                                         type: ShadingType.CLEAR,
            //                                         color: "auto",
            //                                     },
            //                                     borders: {
            //                                         top: {
            //                                             color: "ffffff",
            //                                         },
            //                                         bottom: {
            //                                             color: "ffffff",
            //                                         },
            //                                     },
            //                                     children: [],
            //                                 }),
            //                                 new TableCell({
            //                                     height: {value: 50, type: WidthType},
            //                                     shading: {
            //                                         fill: "D9D9D9",
            //                                         type: ShadingType.CLEAR,
            //                                         color: "auto",
            //                                     },
            //                                     borders: {
            //                                         top: {
            //                                             color: "ffffff",
            //                                         },
            //                                         bottom: {
            //                                             color: "ffffff",
            //                                         },
            //                                     },
            //                                     children: [],
            //                                 }),
            //                                 new TableCell({
            //                                     height: {value: 50, type: WidthType},
            //                                     shading: {
            //                                         fill: "D9D9D9",
            //                                         type: ShadingType.CLEAR,
            //                                         color: "auto",
            //                                     },
            //                                     children: [new Paragraph({
            //                                         alignment: AlignmentType.CENTER,
            //                                         text: `Substituição`,
            //                                         bold: true,
            //                                     }),],
            //                                 }),
            //                                 new TableCell({
            //                                     height: {value: 50, type: WidthType},
            //                                     shading: {
            //                                         fill: "D9D9D9",
            //                                         type: ShadingType.CLEAR,
            //                                         color: "auto",
            //                                     },
            //                                     borders: {
            //                                         top: {
            //                                             color: "ffffff",
            //                                         },
            //                                         bottom: {
            //                                             color: "ffffff",
            //                                         },
            //                                     },
            //                                     children: [],
            //                                 }),
            //                             ]
            //                         }),
            //                         ...productRows,
            //                         new TableRow({
            //                             children: [
            //                                 new TableCell({
            //                                     height: {value: 50, type: WidthType},
            //                                     children: [],
            //                                 }),
            //                                 new TableCell({
            //                                     height: {value: 50, type: WidthType},
            //                                     children: [new Paragraph({
            //                                         alignment: AlignmentType.CENTER,
            //                                         text: `O: Outro(descrever motivo)`,
            //                                         bold: true,
            //                                     }),],
            //                                 }),
            //                                 new TableCell({
            //                                     height: {value: 50, type: WidthType},
            //                                     children: [
            //                                         new Paragraph({
            //                                         alignment: AlignmentType.CENTER,
            //                                         text: `I: Inadequado`,
            //                                         bold: true,
            //                                         }),
            //                                     ],
            //                                 }),
            //                                 new TableCell({
            //                                     height: {value: 50, type: WidthType},
            //                                     children: [
            //                                         new Paragraph({
            //                                         alignment: AlignmentType.CENTER,
            //                                         text: `T: Tempo de Uso`,
            //                                         bold: true,
            //                                         }),
            //                                     ],
            //                                 }),
            //                                 new TableCell({
            //                                     height: {value: 50, type: WidthType},
            //                                     children: [
            //                                         new Paragraph({
            //                                         alignment: AlignmentType.CENTER,
            //                                         text: `D: Defeito`,
            //                                         bold: true,
            //                                         }),
            //                                     ],
            //                                 }),
            //                                 new TableCell({
            //                                     height: {value: 50, type: WidthType},
            //                                     children: [],
            //                                 }),
            //                                 new TableCell({
            //                                     height: {value: 50, type: WidthType},
            //                                     children: [],
            //                                 }),
            //                             ]
            //                         }),
            //                     ],
            //                 })
            //             ]
            //         }),
            //     ]
            // }),
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
        // console.log(tableRows);

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

module.exports = {createClauseWithParagraphs, createEpiProductsTable, getSkillProducts, createEpiTable}

