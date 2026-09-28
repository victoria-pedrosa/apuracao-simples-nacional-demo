/**
 * SIMPLES NACIONAL — Cálculo de Alíquotas Efetivas DAS e ISS
 * Baseado na LC 123/2006 (com redação dada pela LC 155/2016)
 *
 * Fórmula oficial: Alíquota Efetiva = ((RBT12 × Aliq. Nominal) − Parcela a Deduzir) / RBT12
 *
 * Estrutura esperada da aba "EMPRESAS CLIENTES":
 *   A: CÓDIGO | B: CNPJ | C: EMPRESA | D: ANEXO | E: RBT12 | F: ISS | G: DAS
 */
function gerarAliquotasClientes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetClientes = ss.getSheetByName("EMPRESAS CLIENTES");

  const lastRow = sheetClientes.getLastRow();
  if (lastRow < 2) return;

  const data = sheetClientes.getRange(2, 1, lastRow - 1, 7).getValues();

  // ============================================================
  // 1) TABELAS OFICIAIS DO SIMPLES NACIONAL (LC 155/2016)
  //    nom  = alíquota nominal da faixa
  //    ded  = parcela a deduzir
  //    iss  = % de repartição do ISS dentro do DAS (anexos III, IV, V)
  // ============================================================
  const TABELAS = {
    "Anexo I": [
      { max: 180000,    nom: 0.040, ded: 0,      iss: 0      },
      { max: 360000,    nom: 0.073, ded: 5940,   iss: 0      },
      { max: 720000,    nom: 0.095, ded: 13860,  iss: 0      },
      { max: 1800000,   nom: 0.107, ded: 22500,  iss: 0      },
      { max: 3600000,   nom: 0.143, ded: 87300,  iss: 0      },
      { max: 4800000,   nom: 0.190, ded: 378000, iss: 0      }
    ],
    "Anexo II": [
      { max: 180000,    nom: 0.045, ded: 0,      iss: 0      },
      { max: 360000,    nom: 0.078, ded: 5940,   iss: 0      },
      { max: 720000,    nom: 0.100, ded: 13860,  iss: 0      },
      { max: 1800000,   nom: 0.112, ded: 22500,  iss: 0      },
      { max: 3600000,   nom: 0.147, ded: 85500,  iss: 0      },
      { max: 4800000,   nom: 0.300, ded: 720000, iss: 0      }
    ],
    "Anexo III": [
      { max: 180000,    nom: 0.060, ded: 0,      iss: 0.335  },
      { max: 360000,    nom: 0.112, ded: 9360,   iss: 0.320  },
      { max: 720000,    nom: 0.135, ded: 17640,  iss: 0.325  },
      { max: 1800000,   nom: 0.160, ded: 35640,  iss: 0.325  },
      { max: 3600000,   nom: 0.210, ded: 125640, iss: 0.335  },
      { max: 4800000,   nom: 0.330, ded: 648000, iss: 0      } // 6ª faixa: ISS sai do DAS
    ],
    "Anexo IV": [
      { max: 180000,    nom: 0.045, ded: 0,      iss: 0.445  },
      { max: 360000,    nom: 0.090, ded: 8100,   iss: 0.400  },
      { max: 720000,    nom: 0.102, ded: 12420,  iss: 0.400  },
      { max: 1800000,   nom: 0.140, ded: 39780,  iss: 0.400  },
      { max: 3600000,   nom: 0.220, ded: 183780, iss: 0.434  }, // CORRIGIDO: era 125340
      { max: 4800000,   nom: 0.330, ded: 828000, iss: 0      }
    ],
    "Anexo V": [
      { max: 180000,    nom: 0.155, ded: 0,      iss: 0.140  },
      { max: 360000,    nom: 0.180, ded: 4500,   iss: 0.170  },
      { max: 720000,    nom: 0.195, ded: 9900,   iss: 0.190  },
      { max: 1800000,   nom: 0.205, ded: 17100,  iss: 0.210  },
      { max: 3600000,   nom: 0.230, ded: 62100,  iss: 0.235  },
      { max: 4800000,   nom: 0.305, ded: 540000, iss: 0      }
    ]
  };

  // Anexos que NÃO POSSUEM ISS (comércio e indústria)
  const SEM_ISS = { "Anexo I": true, "Anexo II": true };

  // Teto legal do ISS no Simples Nacional
  const TETO_ISS = 0.05;

  const resultadosISS = [];
  const resultadosDAS = [];

  data.forEach(function (linha) {
    const anexoNome = String(linha[3]).trim();

    // Sanitiza RBT12 — aceita "R$ 1.234,56", número puro, etc.
    let rbt12 = linha[4];
    if (typeof rbt12 === "string") {
      rbt12 = parseFloat(
        rbt12.replace(/R\$\s?/g, "").replace(/\./g, "").replace(",", ".").trim()
      );
    }
    if (isNaN(rbt12) || rbt12 === "" || rbt12 === null) rbt12 = 0;

    // Sem anexo válido ou RBT12 zerado → deixa em branco
    if (!TABELAS[anexoNome] || rbt12 <= 0) {
      resultadosISS.push([""]);
      resultadosDAS.push([""]);
      return;
    }

    // Localiza a faixa
    const faixa = TABELAS[anexoNome].find(function (f) { return rbt12 <= f.max; });

    // RBT12 acima de R$ 4.800.000 → empresa sai do Simples
    if (!faixa) {
      resultadosISS.push(["EXCEDEU LIMITE"]);
      resultadosDAS.push(["EXCEDEU LIMITE"]);
      return;
    }

    // ============================================================
    // 2) CÁLCULO DA ALÍQUOTA EFETIVA DAS (fórmula oficial)
    //    Para 1ª faixa (ded=0), o resultado é a própria alíquota nominal.
    // ============================================================
    const aliqDAS = ((rbt12 * faixa.nom) - faixa.ded) / rbt12;

    // ============================================================
    // 3) CÁLCULO DO ISS DESTACADO NA NOTA
    // ============================================================
    let aliqISS;
    if (SEM_ISS[anexoNome]) {
      aliqISS = "NÃO POSSUI";
    } else if (faixa.iss === 0) {
      // 6ª faixa: ISS deixa de ser destacado (recolhido fora do DAS)
      aliqISS = "NÃO DESTACADO";
    } else {
      aliqISS = aliqDAS * faixa.iss;
      // Trava legal: ISS no Simples não pode ultrapassar 5%
      if (aliqISS > TETO_ISS) aliqISS = TETO_ISS;
    }

    resultadosISS.push([aliqISS]);
    resultadosDAS.push([aliqDAS]);
  });

  // ============================================================
  // 4) GRAVAÇÃO — formato percentual nativo (não usar "@" texto!)
  // ============================================================
  const colISS = sheetClientes.getRange(2, 6, resultadosISS.length, 1);
  const colDAS = sheetClientes.getRange(2, 7, resultadosDAS.length, 1);

  colISS.setValues(resultadosISS);
  colDAS.setValues(resultadosDAS);

  // Aplica formato % apenas onde o valor é numérico (preserva textos como "NÃO POSSUI")
  for (let i = 0; i < resultadosISS.length; i++) {
    const cellISS = sheetClientes.getRange(i + 2, 6);
    const cellDAS = sheetClientes.getRange(i + 2, 7);

    if (typeof resultadosISS[i][0] === "number") {
      cellISS.setNumberFormat("0.0000%");
    } else {
      cellISS.setNumberFormat("@");
    }
    if (typeof resultadosDAS[i][0] === "number") {
      cellDAS.setNumberFormat("0.0000%");
    } else {
      cellDAS.setNumberFormat("@");
    }
  }

  SpreadsheetApp.getUi().alert(
    "Cálculo concluído!\n\n" +
    resultadosDAS.length + " empresas processadas.\n" +
    "Fórmula aplicada: ((RBT12 × Nominal) − Dedução) / RBT12\n" +
    "Teto de ISS de 5% aplicado conforme legislação."
  );
}
