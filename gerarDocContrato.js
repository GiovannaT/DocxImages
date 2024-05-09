const axios = require('axios')
const docx = require("docx")
const moment = require("moment")
const fs = require('fs');
const { createEpiTable, createClauseWithParagraphs } = require('./contractUtils');

const caminhoParaJson = 'db.json'

fs.readFile(caminhoParaJson, 'utf8', async (err, data) => {
    if(err){
        console.error('Erro ao ler arquivo Json' + err);
        return
    }

    let task = JSON.parse(data);
    var tableRows = [];

    const countDays = () => {
        var startDate = new Date(task[0].initial_date);
        var endDate = new Date(task[0].final_date);

        var timeDifference = endDate.getTime() - startDate.getTime();
        return daysDifference = Math.ceil(timeDifference / (1000 * 60 * 60 * 24));
    }

    const company = task[0].task_executors.filter((executor) => executor.company_id !== null);
    const leader = task[0].task_executors.filter((executor) => executor.leader === true);

    const getCompany = async () => {
        try {
            const response = await axios.get(`https://3337-avanciconstru-apiserver-0ae2jz7xl1m.ws-us110.gitpod.io/company?id=${company[0].company_id}`, {
                headers: {
                'Authorization': `Bearer eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpYXQiOjE3MTUyNjA5OTMsImV4cCI6MTc2NzEwMDk5Mywic3ViIjoiZjM1ZDg2M2QtMmI4My00MGM4LWI4ZDUtM2ExNzU5YTU2NTc2In0.EepQV0WVRYNQLQp2sPDhJU_Cm34c2rDFPuq0I_fqpDI`
            }});
            // console.log(response.data);
            return response.data[0];

        } catch (error) {
            console.error('erro getcompany')
            // console.error(error.response)
            return null;
        }
    }
    const companyData = await getCompany();

    const { Document, Packer, Paragraph, Table, TextRun, TableRow, TableCell, ShadingType, HeadingLevel, AlignmentType, LevelFormat, WidthType} = docx;
    const doc = new Document({
        creator: "Usuário criador",
        description: `Contrato`,
        title: 'Contrato',
        numbering: {
            config: [
                {
                    reference: "my-numbering",
                    levels: [
                        {
                            level: 0,
                            format: LevelFormat.DECIMAL,
                            text: "%1.",
                            alignment: AlignmentType.START,
                        },
                        {
                            level: 1,
                            format: LevelFormat.DECIMAL,
                            text: "%2.",
                            alignment: AlignmentType.START,
                        },
                        {
                            level: 2,
                            format: LevelFormat.DECIMAL,
                            text: "%3.",
                            alignment: AlignmentType.START,
                        },
                        {
                            level: 3,
                            format: LevelFormat.DECIMAL,
                            text: "%4.",
                            alignment: AlignmentType.START,
                        },
                        {
                            level: 4,
                            format: LevelFormat.DECIMAL,
                            text: "%5.",
                            alignment: AlignmentType.START,
                        },
                        {
                            level: 5,
                            format: LevelFormat.LOWER_LETTER,
                            text: "a)",
                            alignment: AlignmentType.START,
                        },
                        {
                            level: 6,
                            format: LevelFormat.LOWER_LETTER,
                            text: "b)",
                            alignment: AlignmentType.START,
                        },
                        {
                            level: 7,
                            format: LevelFormat.LOWER_LETTER,
                            text: "c)",
                            alignment: AlignmentType.START,
                        },
                        {
                            level: 8,
                            format: LevelFormat.LOWER_LETTER,
                            text: "%d)",
                            alignment: AlignmentType.START,
                        },
                    ],
                },
            ],
        },
        sections: [{
            children: [
                new Paragraph({
                    children: [
                        new TextRun({
                            text: "CONTRATO DE PRESTAÇÃO DE SERVIÇOS",
                            bold: true,
                        })
                    ],
                    alignment: AlignmentType.CENTER,
                }),
                new Paragraph({
                    children:[
                        new TextRun({
                            text: `${leader[0].executor.name} - CONTRATO - ${task[0].place.name}`,
                            bold: true,
                        }),
                    ],
                    alignment: AlignmentType.CENTER,
                }
                ),
                new Paragraph({
                    children: [
                        new TextRun({text: "CLÁUSULA 1ª – DAS PARTES –", bold: true}),
                    ],
                    spacing: {
                        before: 200,
                    },
                }),
                new Paragraph({
                    children: [
                        new TextRun("A pessoa jurídica: "),
                        new TextRun({text: "HABIT CONSTRUÇÕES E SERVIÇOS LTDA, ", bold: true,}),
                        new TextRun("inscrito no CNPJ nº: 28.697.934/0001- 43, e-mail recepcaohbt@gmail.com, com sede em Rua Projetada F, Nº 76, Ribeirão do Lipa, Cuiabá – MT, CEP: 78048- 163 e telefone para contato 6598108-8373, doravante denominado "),
                        new TextRun({text: "CONTRATANTE.", bold: true,}),
                    ],
                }),
                new Paragraph({
                    spacing: {
                        before: 200,
                        after: 200
                    },
                    children: [
                        new TextRun(`Pessoa jurídica: `),
                        new TextRun({
                            text: `${company[0].company.name}, CNPJ: ${company[0].company.cnpj},`,
                            bold:true,
                        }),
                        new TextRun(` telefone para contato ${company[0].company.phone}, neste ato representada por ${leader[0].executor.name}, CPF: ${leader[0].executor.cpf}, data de nascimento: ${moment(leader[0].executor.birthday).format('DD/MM/YYYY')}, denominada `),
                        new TextRun({
                            text: "CONTRATADA.",
                            bold: true,
                        })
                    ],
                }),
                new Paragraph({
                    spacing: {
                        before: 200,
                        after: 200
                    },
                    children: [
                        new TextRun(`As partes acima identificadas têm, entre si, justo e acertado, o presente contrato de prestação de serviços, ficando desde já aceito, pelas cláusulas abaixo descritas.`),
                    ],
                }),
                new Paragraph({
                    children: [
                        new TextRun({text: "CLÁUSULA 2ª – DO OBJETO –", bold: true}),
                    ],
                    spacing: {
                        before: 200,
                    },
                }),
                new Paragraph({
                    children: [
                        new TextRun(`Por meio deste contrato, a CONTRATADA se compromete a prestar ao CONTRATANTE os seguintes serviços:
                        `),
                    ],
                }),
                new Paragraph({
                    children: task[0].task_service_orders.map(service => {
                        return new TextRun({text: `${service.service_order.service.description}; `, bold: true})
                    })
                }),
                new Paragraph({
                    spacing: {
                        before: 200,
                    },
                    children: [
                        new TextRun(`§ 1º. A `),
                        new TextRun({
                            text: "CONTRATADA ",
                            bold: true,
                        }),
                        new TextRun(`prestará os serviços descritos nesta cláusula sem qualquer exclusividade, podendo desempenhar atividades para terceiros, desde que não haja conflito de interesses com o pactuado no presente contrato.`)
                    ],
                }),
                new Paragraph({
                    spacing: {
                        before: 200,
                        after: 200
                    },
                    children: [
                        new TextRun(`§ 2º. Os serviços descritos acima serão prestados com total autonomia, sem pessoalidade e sem qualquer subordinação ao `),
                        new TextRun({
                            text: "CONTRATANTE.",
                            bold: true,
                        })
                    ],
                }),
                createClauseWithParagraphs(3, 'DO PRAZO'),
                new Paragraph({
                    children: [
                        new TextRun({
                            text: "Os serviços ora contratados serão prestados pelo prazo de"
                        }),
                        new TextRun({text:` ${countDays()} dias `, bold: true}),
                        new TextRun({
                            text: "com início em "
                        }),
                        new TextRun({
                        text: `${moment(task[0].initial_date).format("DD/MM/YYYY")}`,
                            bold: true,
                        }),
                        new TextRun(" e será finalizado na data "),
                        new TextRun({
                            text:`${moment(task[0].final_date).format("DD/MM/YYYY")}`,
                            bold: true,
                        })
                    ],
                }),
                new Paragraph({
                    children: [
                        new TextRun(`O serviço prestado deverá ocorrer da seguinte forma:`),
                    ],
                    spacing: {
                        before: 200,
                    },
                }),
                new Paragraph({
                    spacing: {
                        before: 200,
                    },
                    children: [
                        new TextRun(`§ 1º. Ao final do prazo acima referido, poderá o contrato ser renovado em mútuo acordo, devendo constar em termo aditivo o novo prazo acordado.
                        `),
                    ],
                }),
                new Paragraph({
                    spacing: {
                        before: 200,
                        after: 200
                    },
                    children: [
                        new TextRun(`§ 2º. Caso não ocorra a renovação pelas partes no final do prazo acima referido, este contrato será automaticamente rescindido sem que haja a necessidade de aviso prévio.
                        `),
                    ],
                }),
                createClauseWithParagraphs(4, 'DA RETRIBUIÇÃO'),
                new Paragraph({
                    children: [
                        new TextRun(`Pela prestação dos serviços o `),
                        new TextRun({
                            text: "CONTRATANTE ",
                            bold: true,
                        }),
                        new TextRun("pagará à "),
                        new TextRun({
                            text:"CONTRATADA ",
                            bold: true,
                        }),
                        new TextRun("conforme o valor do orçamento, que será pago à vista assim que finalizado e aprovados os serviços, e "),
                        new TextRun({
                            text:"não será incluso o fornecimento de refeição.",
                            bold: true,
                        })
                    ],
                }),
                new Paragraph({
                    children: [
                        new TextRun(`§ 1º. Deverá o pagamento acordado neste instrumento ser efetuado por meio de transferência bancária em nome do contratado cito a conta: `),
                        new TextRun({
                            text: 'Chave Pix: ',
                        }),
                        // new TextRun({
                        //     text: `${companyData.bank_account.pix_type} `,
                        //     bold: true,
                        // }),
                        // new TextRun({
                        //     text: `${companyData.bank_account.pix}, `,
                        //     bold: true,
                        // }),

                        
                        // new TextRun({
                        //     text: 'Responsável: ',
                        // }),
                        // new TextRun({
                        //     text: `${companyData.bank_account.account_owner} `,
                        //     bold: true,
                        // }),
                    ],
                    spacing: {
                        before: 200,
                    },
                }),
                createClauseWithParagraphs(5, 'DAS OBRIGAÇÕES DA CONTRATADA'),
                new Paragraph({
                    children: [
                        new TextRun(`Sem prejuízo de outras disposições deste contrato, constituem obrigações da `),
                        new TextRun({
                            text: "CONTRATADA:",
                            bold:true,
                        })
                    ],
                }),
                new Paragraph({
                    indent: {
                        left: 300,
                    },
                    children: [
                        new TextRun(`A. Prestar os serviços contratados na forma e modo ajustados, dentro das normas, de segurança do trabalho, e NRs aplicáveis, com profissionalismo, retorno dos serviços realizados, por meio de relatórios fotográficos, filmagens, dando plena e total garantia dos mesmos; nos termos da Lei do Código do Consumidor.`),
                    ],
                    spacing: {
                        before: 200,
                    },
                }),
                new Paragraph({
                    indent: {
                        left: 300,
                    },
                    children: [
                        new TextRun(`B.	Executar os serviços contratados utilizando a melhor didática e aplicabilidade, visando sempre atingir o melhor resultado, sob sua exclusiva responsabilidade, sendo-lhe vedada a transferência dos mesmos a terceiros, sem prévia e expressa concordância do CONTRATANTE;
                        `),
                    ],
                    spacing: {
                        before: 200,
                    },
                }),
                new Paragraph({
                    indent: {
                        left: 300,
                    },
                    children: [
                        new TextRun(`C.	Ser responsável pelos atos praticados por seus responsáveis, bem como pelos danos que os mesmos venham a causar para o CONTRATANTE, desde que comprovados, em decorrência da prestação dos serviços prestados neste contrato;
                        `),
                    ],
                    spacing: {
                        before: 200,
                    },
                }),
                new Paragraph({
                    indent: {
                        left: 300,
                    },
                    children: [
                        new TextRun(`D.	Cumprir todas as determinações impostas pelas autoridades públicas competentes;
                        `),
                    ],
                    spacing: {
                        before: 200,
                    },
                }),
                new Paragraph({
                    indent: {
                        left: 300,
                    },
                    children: [
                        new TextRun(`E.	Arcar devidamente, nos termos da legislação trabalhista, com a remuneração e demais verbas laborais, tais como INSS, devidas a seus subordinados, inclusive encargos fiscais e previdenciários referentes às relações de trabalho;
                        `),
                    ],
                    spacing: {
                        before: 200,
                    },
                }),
                new Paragraph({
                    indent: {
                        left: 300,
                    },
                    children: [
                        new TextRun(`F.	Arcar com todas as despesas de natureza tributária decorrentes dos serviços especificados neste contrato;
                        `),
                    ],
                    spacing: {
                        before: 200,
                    },
                }),
                new Paragraph({
                    indent: {
                        left: 300,
                    },
                    children: [
                        new TextRun(`G.	Fornecer os equipamentos necessários à prestação dos serviços, ou em comum acordo, usar as ferramentas do CONTRATANTE, sob sua responsabilidade de uso e devolução no mesmo estado.`),
                    ],
                    spacing: {
                        before: 200,
                    },
                }),
                createClauseWithParagraphs(6, 'DAS OBRIGAÇÕES DO CONTRATANTE'),
                new Paragraph({
                    children: [
                        new TextRun(`Sem prejuízo de outras disposições deste contrato, constituem obrigações do `),
                        new TextRun({
                            text: 'CONTRATANTE:',
                            bold:true,
                        })
                    ],
                }),
                new Paragraph({
                    indent: {
                        left: 300,
                    },
                    children: [
                        new TextRun(`A.	Fornecer à CONTRATADA todas as informações necessárias à realização do serviço, devendo especificar os detalhes necessários à perfeita execução do mesmo, e a forma de como ele deve ser entregue;
                        `),
                    ],
                    spacing: {
                        before: 200,
                    },
                }),
                new Paragraph({
                    indent: {
                        left: 300,
                    },
                    children: [
                        new TextRun(`B.	Efetuar o pagamento, nas datas e nos termos definidos neste contrato.
                        `),
                    ],
                    spacing: {
                        before: 200,
                    },
                }),
                createClauseWithParagraphs(7, 'DA RESCISÃO'),
                new Paragraph({
                    children: [
                        new TextRun(`O presente instrumento poderá ser rescindido caso qualquer uma das partes descumpra o disposto neste contrato.
                        `),
                    ],
                }),
                new Paragraph({
                    children: [
                        new TextRun(`§ 1º. Na hipótese do CONTRATANTE solicitar a rescisão antecipada deste contrato sem justa causa, será obrigado a pagar à CONTRATADA por inteiro qualquer retribuição vencida e não paga, e um terço do que ela receberia até o final do contrato (caso a rescisão contratual seja sem justa causa - motivação).
                        `),
                    ],
                    spacing: {
                        before: 200,
                    },
                }),
                new Paragraph({
                    children: [
                        new TextRun(`§ 2º. Na hipótese da CONTRATADA solicitar a rescisão antecipada deste contrato sem justa causa, esta terá direito à retribuição vencida, mas responderá por perdas e danos que cause ao CONTRATANTE.
                        `),
                    ],
                    spacing: {
                        before: 200,
                    },
                }),
                new Paragraph({
                    children: [
                        new TextRun(`§ 3º. A rescisão com justa causa por parte do CONTRATANTE obriga a devolução pela CONTRATADA de quaisquer valores já pagos referentes a serviços não desenvolvidos.
                        `),
                    ],
                    spacing: {
                        before: 200,
                    },
                }),
                createClauseWithParagraphs(8, 'DA MULTA'),
                new Paragraph({
                    children: [
                        new TextRun(`§ 1º. A CONTRATADA, no caso de atraso na entrega dos serviços superior a 5 dias, está sujeita a MULTA de 10% com base no valor do contrato.
                        `),
                    ],
                }),
                new Paragraph({
                    children: [
                        new TextRun(`§ 2º. O CONTRATANTE não vindo a efetuar o pagamento, fica obrigado a pagar multa de 1 % (um por cento) sobre o valor devido, bem como juros de mora de 3% (três por cento) ao mês, mais correção monetária apurada conforme variação do IGP-M no período.
                        `),
                    ],
                    spacing: {
                        before: 200,
                    },
                }),
                createClauseWithParagraphs(9, 'DA EXTINÇÃO DO CONTRATO e TIPO DO CONTRATO'),
                new Paragraph({
                    children: [
                        new TextRun(`O presente contrato extingue-se sem que assista às partes direito a qualquer tipo de indenização, ressarcimento ou multa, por mais especial que seja, nas seguintes hipóteses:
                        `),
                    ],
                }),
                new Paragraph({
                    indent: {
                        left: 300,
                    },
                    children: [
                        new TextRun(`A.	Por insolvência, impetração ou solicitação de concordata, ou falência, de qualquer uma das partes;
                        `),
                    ],
                    spacing: {
                        before: 200,
                    },
                }),
                new Paragraph({
                    indent: {
                        left: 300,
                    },
                    children: [
                        new TextRun(`B.	Por qualquer impossibilidade da continuação do contrato, motivada por força maior ou caso fortuito.
                        `),
                    ],
                    spacing: {
                        before: 200,
                    },
                }),
                new Paragraph({
                    indent: {
                        left: 300,
                    },
                    children: [
                        new TextRun(`C.	Este contrato é particular, não tendo o CONTRATANTE, qualquer vínculo empregatício com a
                        CONTRATADA, sendo o CONTRATADA, remunerado por diária, conforme acordado entre as partes.

                        `),
                    ],
                    spacing: {
                        before: 200,
                    },
                }),
                createClauseWithParagraphs(10,'DAS CONDIÇÕES GERAIS'),
                new Paragraph({
                    children: [
                        new TextRun(`Salvo expressa autorização do CONTRATANTE, não poderá a CONTRATADA transferir ou subcontratar os serviços previstos neste instrumento, sob o risco de ocorrer a rescisão imediata.
                        `),
                    ],
                }),
                new Paragraph({
                    children: [
                        new TextRun(`§ 1º. Qualquer condescendência entre as partes quanto ao cumprimento de qualquer cláusula do presente contrato, constituirá mera tolerância e não importará em alteração ou modificação das cláusulas contratuais.
                        `),
                    ],
                    spacing: {
                        before: 200,
                    },
                }),
                new Paragraph({
                    children: [
                        new TextRun(`§ 2º. Qualquer serviço adicional, desde que acordado entre as partes, será objeto de TERMO ADITIVO ao instrumento original.
                        `),
                    ],
                    spacing: {
                        before: 200,
                    },
                }),
                createClauseWithParagraphs(11, 'DO FORO'),
                new Paragraph({
                    children: [
                        new TextRun(`Fica desde já eleito o foro da comarca de Cuiabá-MT para serem resolvidas eventuais pendências decorrentes deste contrato.
                        `),
                    ],
                }),
                new Paragraph({
                    children: [
                        new TextRun(`Por estarem assim certos e ajustados, firmam os signatários este instrumento em 02 (duas) vias de igual forma.
                        `),
                    ],
                    spacing: {
                        before: 200,
                    },
                }),
                new Paragraph({
                    children: [
                        new TextRun(`CUIABÁ-MT, ${new Date().getDate()}/${new Date().getMonth()+1}/${new Date().getFullYear()}`),
                    ],
                    spacing: {
                        before: 200,
                    },
                }),
                new Paragraph({
                    children: [
                        new TextRun({text:`CONTRATANTE:`, bold: true}),
                    ],
                    alignment: AlignmentType.CENTER,
                    spacing: {
                        before: 300,
                    },
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
                        new TextRun({text:`HABIT CONSTRUÇÕES E SERVIÇOS LTDA`, bold: true}),
                    ],
                    alignment: AlignmentType.CENTER,
                    spacing: {
                        before: 400,
                    },
                }),
                new Paragraph({
                    children: [
                        new TextRun(`CNPJ: 28.697.934/0001- 43`),
                    ],
                    alignment: AlignmentType.CENTER,

                }),
                new Paragraph({
                    children: [
                        new TextRun({text:`CONTRATADO(A):`, bold: true}),
                    ],
                    alignment: AlignmentType.CENTER,
                    spacing: {
                        before: 300,
                    },
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
                        new TextRun({text:`${leader[0].executor.name}`, bold:true,}),
                    ],
                    alignment: AlignmentType.CENTER,
                    spacing: {
                        before: 400,
                    },
                }),
                new Paragraph({
                    children: [
                        new TextRun({text:`CPF: ${leader[0].executor.cpf}`}),
                    ],
                    alignment: AlignmentType.CENTER,

                }),
                new Paragraph({
                    children: [
                        new TextRun({text:`TESTEMUNHAS:`, bold: true}),
                    ],
                    alignment: AlignmentType.CENTER,
                    spacing: {
                        before: 300,
                    },
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
                        new TextRun({text:`Testemunha 1`, bold:true}),
                    ],
                    alignment: AlignmentType.CENTER,
                    spacing: {
                        before: 400,
                    },
                }),
                new Paragraph({
                    children: [
                        new TextRun({text:`CPF: DA TESTEMUNHA`, bold:true}),
                    ],
                    alignment: AlignmentType.CENTER,

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
                        new TextRun({text:`Testemunha 1`, bold:true}),
                    ],
                    alignment: AlignmentType.CENTER,
                    spacing: {
                        before: 400,
                    },
                }),
                new Paragraph({
                    children: [
                        new TextRun({text:`CPF: DA TESTEMUNHA`, bold:true}),
                    ],
                    alignment: AlignmentType.CENTER,

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
                        new TextRun({text:`Testemunha 1`, bold:true}),
                    ],
                    alignment: AlignmentType.CENTER,
                    spacing: {
                        before: 400,
                    },
                }),
                new Paragraph({
                    children: [
                        new TextRun({text:`CPF: DA TESTEMUNHA`, bold:true}),
                    ],
                    alignment: AlignmentType.CENTER,

                }),
                new Paragraph({

                }),
            ]
        }],
    })

    for (const executor of task[0].task_executors) {
        createEpiTable(doc, executor);
    }

    
    let totalValueExecutor = 0;
    serviceTableRows = []

    task[0].task_service_orders.forEach(async object => {
        totalValueExecutor += object.value_executor;
        serviceTableRows.push(
            new TableRow({
                cantSplit: true,
                tableHeader: true,
                children: [
                    new TableCell({
                        children: [
                            new Paragraph({
                            text: `${object.service_order.service.description}`,
                            })
                        ],
                    }),
                    new TableCell({
                        height: {value: 50, type: WidthType},
                        width: {
                            size: 1000,
                            type: WidthType.DXA,
                        },
                        children: [new Paragraph(`${object.service_order.service.unity}`)],
                    }),
                    new TableCell({
                        height: {value: 50, type: WidthType},
                        width: {
                            size: 1000,
                            type: WidthType.DXA,
                        },
                        children: [new Paragraph(`${object.amount}`)],
                    }),
                    new TableCell({
                        height: {value: 50, type: WidthType},
                        width: {
                            size: 1000,
                            type: WidthType.DXA,
                        },
                        children: [new Paragraph(`R$ ${object.value_executor}`)],
                    }),
                ],
            }),
        )
    })

    const newSection = () => {
        const section = {
            children: [
                new Paragraph({
                    pageBreakBefore: true,
                }),
                new Table({
                    width: {
                        size: 9000,
                        type: WidthType.DXA,
                    },
                    columnWidths: [8500, 8500],
                    rows: [
                        new TableRow({
                            cantSplit: true,
                            tableHeader: true,
                            children: [
                                new TableCell({
                                    columnSpan: 4,
                                    children: [
                                        new Paragraph({
                                        alignment: AlignmentType.CENTER,
                                        text: `Serviços`,
                                        heading: HeadingLevel.HEADING_3,
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
                                    children: [
                                        new Paragraph({
                                        alignment: AlignmentType.CENTER,
                                        text: `Descrição`,
                                        heading: HeadingLevel.HEADING_3,
                                        })
                                    ],
                                }),
                                new TableCell({
                                    width: {
                                        size: 1000,
                                        type: WidthType.DXA,
                                    },
                                    children: [
                                        new Paragraph({
                                        alignment: AlignmentType.CENTER,
                                        text: `Unidade`,
                                        heading: HeadingLevel.HEADING_3,
                                        })
                                    ],
                                }),
                                new TableCell({
                                    width: {
                                        size: 1000,
                                        type: WidthType.DXA,
                                    },
                                    children: [
                                        new Paragraph({
                                        alignment: AlignmentType.CENTER,
                                        text: `Quantidade`,
                                        heading: HeadingLevel.HEADING_3,
                                        })
                                    ],
                                }),
                                new TableCell({
                                    width: {
                                        size: 1000,
                                        type: WidthType.DXA,
                                    },
                                    children: [
                                        new Paragraph({
                                        alignment: AlignmentType.CENTER,
                                        text: `Valor`,
                                        heading: HeadingLevel.HEADING_3,
                                        })
                                    ],
                                }),
                            ],
                        }),
                        ...serviceTableRows,
                        new TableRow({
                            cantSplit: true,
                            tableHeader: true,
                            children: [
                                new TableCell({
                                    children: [],
                                }),
                                new TableCell({
                                    children: [],
                                }),
                                new TableCell({
                                    alignment: AlignmentType.RIGHT,
                                    children: [
                                        new Paragraph({
                                        text: `Valor total: `,
                                        })
                                    ],
                                }),
                                new TableCell({
                                    height: {value: 50, type: WidthType},
                                    width: {
                                        size: 2000,
                                        type: WidthType.DXA,
                                    },
                                    children: [new Paragraph(`R$ ${totalValueExecutor}`)],
                                }),
                            ],
                        }),
                    ]
                }),
            ],
        }
        doc.addSection(section);
    }
    newSection();


    Packer.toBuffer(doc).then((buffer) => {
        fs.writeFileSync("NewDocContract.docx", buffer);
    });
})