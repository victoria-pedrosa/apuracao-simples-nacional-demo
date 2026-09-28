/**
 * =============================================================================
 * ROTINA DE ANTECIPAÇÃO DO DAS 
 * 1. Importar Antecipação (Tabelão + CSV Exemplo) - COM PRIORIDADES
 * 2. Importar Faturamento (Domínio) -> Lê linha "Total" e arrasta o Código
 * 3. Conferir Faturamento (FAT x DOMINIO)
 * 4. Varredura do Drive (DAS e Extrato)
 * =============================================================================
 */

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('🚀 Rotina DAS')
    .addItem('📥 1. Importar Arquivo e Gerar Rotina', 'iniciarRotina')
    .addItem('💰 2. Importar Faturamento (Domínio)', 'iniciarImportacaoFaturamento')
    .addItem('📊 3. Conferência: Faturamento x Domínio', 'conferirFaturamento')
    .addSeparator()
    .addItem('🔍 4. Varredura: Checar DAS/Extrato no Drive', 'checarArquivosDrive')
    .addToUi();
}

function buscarIndice(headers, nomeBusca) {
  let nomeLimpo = nomeBusca.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim();
  return headers.findIndex(h => h.toString().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim() === nomeLimpo);
}

function parseDinheiro(val) {
  if (val === "" || val == null) return 0;
  if (typeof val === 'string') {
    let str = val.toUpperCase().trim();
    if (str === "SF" || str === "SEM FATURAMENTO" || str === "SEM MOVIMENTO") return "SF";
    let numStr = str.replace(/[^\d,.-]/g, '').replace(/\./g, '').replace(',', '.');
    let parsed = parseFloat(numStr);
    return isNaN(parsed) ? 0 : parsed;
  }
  return parseFloat(val) || 0;
}

// =============================================================================
// FUNÇÃO 1: JANELA PARA PEDIR COMPETÊNCIA E ARQUIVO (ANTECIPAÇÃO)
// =============================================================================
function iniciarRotina() {
  const ui = SpreadsheetApp.getUi();
  
  const prompt = ui.prompt(
    '📅 Competência', 
    'Digite a competência para nomear a aba (ex: 06/2026):', 
    ui.ButtonSet.OK_CANCEL
  );
  
  if (prompt.getSelectedButton() !== ui.Button.OK) return; 
  const comp = prompt.getResponseText().trim();
  if (!comp) {
    ui.alert("O nome da aba não pode ficar vazio.");
    return;
  }

  const htmlForm = `
  <!DOCTYPE html>
  <html>
    <head>
      <style>
        body { font-family: Arial, sans-serif; padding: 10px; color: #333; }
        input[type="file"] { margin: 15px 0; font-size: 14px; }
        button { background-color: #274e13; color: white; padding: 10px 15px; border: none; cursor: pointer; border-radius: 4px; font-weight: bold; width: 100%; }
        button:hover { background-color: #1a350c; }
        #status { color: #d93025; font-size: 13px; margin-top: 15px; font-weight: bold; text-align: center; }
      </style>
    </head>
    <body>
      <p>Anexe a planilha CSV de <b>ANTECIPAÇÃO</b> para a competência <b>${comp}</b>:</p>
      <input type="file" id="arquivoCsv" accept=".csv" />
      <button onclick="enviar()">🚀 Gerar Tabela</button>
      <div id="status"></div>
      
      <script>
        function enviar() {
          const fileInput = document.getElementById('arquivoCsv');
          if (fileInput.files.length === 0) {
            alert('Por favor, selecione o arquivo CSV!');
            return;
          }
          document.getElementById('status').innerText = 'Lendo arquivo e cruzando dados... aguarde!';
          document.querySelector('button').disabled = true;
          
          const reader = new FileReader();
          reader.onload = function(e) {
            const conteudo = e.target.result;
            google.script.run
              .withSuccessHandler(function() {
                google.script.host.close();
              })
              .withFailureHandler(function(err) {
                alert('Erro: ' + err.message);
                google.script.host.close();
              })
              .processarDadosRecebidos(conteudo, '${comp}');
          };
          reader.readAsText(fileInput.files[0], 'UTF-8');
        }
      </script>
    </body>
  </html>
  `;
  
  const htmlOutput = HtmlService.createHtmlOutput(htmlForm).setWidth(350).setHeight(220);
  ui.showModalDialog(htmlOutput, '📥 Importar Antecipação');
}

// =============================================================================
// FUNÇÃO 1.1: PROCESSAR ARQUIVO DE ANTECIPAÇÃO E GERAR ABA
// =============================================================================
function processarDadosRecebidos(csvText, nomeAbaDestino) {
  const ssDestino = SpreadsheetApp.getActiveSpreadsheet();
  
  // Detector Automático de Delimitador do CSV
  let dadosAnt = Utilities.parseCsv(csvText, ';'); 
  if (dadosAnt.length > 0 && dadosAnt[0].length <= 1) {
    dadosAnt = Utilities.parseCsv(csvText, ','); 
  }

  const cabecalhoAnt = dadosAnt[0];
  const idxCnpjAnt = buscarIndice(cabecalhoAnt, "CNPJ");
  const idxStatusAnt = buscarIndice(cabecalhoAnt, "Status");
  const idxNotasFora = buscarIndice(cabecalhoAnt, "Notas Fora Comp.");
  
  const mapaAntecipacao = {};
  for (let i = 1; i < dadosAnt.length; i++) {
    let cnpj = String(dadosAnt[i][idxCnpjAnt] || "").replace(/[^\d]/g, ''); 
    if (cnpj) {
      mapaAntecipacao[cnpj] = {
        status: idxStatusAnt !== -1 ? String(dadosAnt[i][idxStatusAnt]).trim() : "",
        notasFora: idxNotasFora !== -1 ? dadosAnt[i][idxNotasFora] : 0
      };
    }
  }

  const urlTabelao = "https://docs.google.com/spreadsheets/d/ID_EXEMPLO/edit";
  const ssTabelao = SpreadsheetApp.openByUrl(urlTabelao);
  const abaTabelao = ssTabelao.getSheetByName("empresas");
  const dadosTabelao = abaTabelao.getDataRange().getValues();
  const cabecalhoTab = dadosTabelao[0];
  
  const cols = {
    cod: buscarIndice(cabecalhoTab, "Código Dominio"),
    empresa: buscarIndice(cabecalhoTab, "Empresa"),
    cnpj: buscarIndice(cabecalhoTab, "CNPJ"),
    matriz: buscarIndice(cabecalhoTab, "Matriz / Filial"),
    fiscal: buscarIndice(cabecalhoTab, "Fiscal Responsavel"),
    prolabore: buscarIndice(cabecalhoTab, "Tira Pro-labore?"),
    folha: buscarIndice(cabecalhoTab, "Tem Folha de Pagamento?"),
    cidade: buscarIndice(cabecalhoTab, "Cidade"),
    abertura: buscarIndice(cabecalhoTab, "DATA DE ABERTURA"),
    escr: buscarIndice(cabecalhoTab, "DATA EXEMPLO")
  };
  
  cols.particularidades = buscarIndice(cabecalhoTab, "Particularidades");
  if (cols.particularidades === -1) cols.particularidades = buscarIndice(cabecalhoTab, "Particularidade");
  
  const idxRegimeCaixa = 24; 
  
  const cabecalhoFinal = [
    "Código Dominio", "Empresa", "CNPJ", "Matriz / Filial", "Fiscal Responsavel", 
    "Tira Pro-labore?", "Tem Folha de Pagamento?", "Cidade", "DATA DE ABERTURA", 
    "DATA EXEMPLO", "ANTECIPA?", "FAT", "DOMINIO", "IMPORTAÇÃO", "ACESSO SN", 
    "SISTEMA", "DAS", "EXTRATO"
  ];
  
  const linhasFiltradas = [];
  
  for (let i = 1; i < dadosTabelao.length; i++) {
    let linhaTab = dadosTabelao[i];
    let fiscal = String(linhaTab[cols.fiscal] || "").toLowerCase();
    let regime = String(linhaTab[idxRegimeCaixa] || "").toLowerCase();
    
    let ehFiscal4ou5 = fiscal.includes("fiscal 4") || fiscal.includes("fiscal 5");
    let naoEhCaixa = !regime.includes("caixa");
    
    if (ehFiscal4ou5 && naoEhCaixa) {
      let cnpjLimpo = String(linhaTab[cols.cnpj] || "").replace(/[^\d]/g, '');
      
      let dadosAntCnpj = mapaAntecipacao[cnpjLimpo] || { status: "", notasFora: 0 };
      let statusCsv = dadosAntCnpj.status;
      
      if (statusCsv.toUpperCase() === "NAO ANTECIPA" || statusCsv.toUpperCase() === "NÃO ANTECIPA") {
        statusCsv = "NÃO ANTECIPAR";
      }
      
      let cidadeTab = String(cols.cidade !== -1 ? linhaTab[cols.cidade] : "");
      let cidadeLimpa = cidadeTab.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim();
      
      if (cidadeLimpa === "salvador") {
        let numNotasFora = parseInt(dadosAntCnpj.notasFora);
        if (isNaN(numNotasFora)) numNotasFora = 0;
        
        if (statusCsv !== "" && statusCsv.toUpperCase() !== "SEM NOTAS" && numNotasFora > 0) {
          statusCsv = "EMITIU FORA DA COMPETÊNCIA - NÃO ANTECIPAR";
        }
      }
      
      let finalStatus = statusCsv; 
      let particularidade = "";
      
      if (cols.particularidades !== -1) {
        particularidade = String(linhaTab[cols.particularidades] || "").toUpperCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim();
      }
      
      if (particularidade.includes("DAS PRIORITARIO")) {
        if (statusCsv !== "") {
          finalStatus = "DAS PRIORITÁRIO - " + statusCsv;
        } else {
          finalStatus = "DAS PRIORITÁRIO";
        }
      }
      
      linhasFiltradas.push([
        linhaTab[cols.cod], linhaTab[cols.empresa], linhaTab[cols.cnpj], 
        linhaTab[cols.matriz], linhaTab[cols.fiscal], 
        cols.prolabore !== -1 ? linhaTab[cols.prolabore] : "-", 
        cols.folha !== -1 ? linhaTab[cols.folha] : "-", 
        cols.cidade !== -1 ? linhaTab[cols.cidade] : "-", 
        linhaTab[cols.abertura], linhaTab[cols.escr], 
        finalStatus, 
        "", "", "", "", "", "", ""
      ]);
    }
  }
  
  let abaDestino = ssDestino.getSheetByName(nomeAbaDestino);
  if (!abaDestino) {
    abaDestino = ssDestino.insertSheet(nomeAbaDestino);
  } else {
    abaDestino.clear(); 
  }
  
  abaDestino.getRange(1, 1, 1, cabecalhoFinal.length).setValues([cabecalhoFinal])
            .setFontWeight("bold").setBackground("#274e13").setFontColor("#ffffff");
            
  if (linhasFiltradas.length > 0) {
    abaDestino.getRange(2, 1, linhasFiltradas.length, cabecalhoFinal.length).setValues(linhasFiltradas);
    
    const totalLinhasTabela = linhasFiltradas.length + 1;
    for (let r = 3; r <= totalLinhasTabela; r += 2) {
      abaDestino.getRange(r, 1, 1, cabecalhoFinal.length).setBackground("#efefef");
    }
  }
  
  ssDestino.toast("Aba criada! Utilize o botão 2 para subir o Faturamento.", "✅ Concluído!", 10);
}

// =============================================================================
// FUNÇÃO 2: JANELA PARA IMPORTAR FATURAMENTO (DOMÍNIO)
// =============================================================================
function iniciarImportacaoFaturamento() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSheet();
  const nomeAba = sheet.getName();
  
  if (nomeAba === "empresas" || nomeAba === "BASE_ANTECIPACAO") {
    ui.alert("⚠️ Vá para a aba da competência (ex: 06/2026) antes de importar o faturamento.");
    return;
  }

  const htmlForm = `
  <!DOCTYPE html>
  <html>
    <head>
      <style>
        body { font-family: Arial, sans-serif; padding: 10px; color: #333; }
        input[type="file"] { margin: 15px 0; font-size: 14px; }
        button { background-color: #274e13; color: white; padding: 10px 15px; border: none; cursor: pointer; border-radius: 4px; font-weight: bold; width: 100%; }
        button:hover { background-color: #1a350c; }
        #status { color: #d93025; font-size: 13px; margin-top: 15px; font-weight: bold; text-align: center; }
      </style>
    </head>
    <body>
      <p>Anexe o <b>CSV</b> do Faturamento da Domínio para a aba <b>${nomeAba}</b>:</p>
      <p style="font-size: 12px; color: #555;">(Lembre-se de salvar o arquivo .xls como .CSV no Excel antes de subir aqui)</p>
      <input type="file" id="arquivoCsv" accept=".csv" />
      <button onclick="enviar()">💰 Importar Faturamento</button>
      <div id="status"></div>
      
      <script>
        function enviar() {
          const fileInput = document.getElementById('arquivoCsv');
          if (fileInput.files.length === 0) {
            alert('Por favor, selecione o arquivo CSV!');
            return;
          }
          document.getElementById('status').innerText = 'Atualizando faturamentos... aguarde!';
          document.querySelector('button').disabled = true;
          
          const reader = new FileReader();
          reader.onload = function(e) {
            const conteudo = e.target.result;
            google.script.run
              .withSuccessHandler(function() {
                google.script.host.close();
              })
              .withFailureHandler(function(err) {
                alert('Erro: ' + err.message);
                google.script.host.close();
              })
              .processarFaturamentoRecebido(conteudo, '${nomeAba}');
          };
          reader.readAsText(fileInput.files[0], 'ISO-8859-1'); 
        }
      </script>
    </body>
  </html>
  `;
  
  const htmlOutput = HtmlService.createHtmlOutput(htmlForm).setWidth(350).setHeight(250);
  ui.showModalDialog(htmlOutput, '💰 Importar Faturamento (Domínio)');
}

// =============================================================================
// FUNÇÃO 2.1: PROCESSAR ARQUIVO DA DOMÍNIO (LÓGICA DO ARRASTO DE CÓDIGO)
// =============================================================================
function processarFaturamentoRecebido(csvText, nomeAba) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(nomeAba);
  
  // Detector Automático de Delimitador do CSV
  let dadosFat = Utilities.parseCsv(csvText, ';'); 
  if (dadosFat.length > 0 && dadosFat[0].length <= 1) {
    dadosFat = Utilities.parseCsv(csvText, ','); 
  }
  
  const mapaFaturamento = {};
  let ultimoCodigoMemorizado = "";
  
  for (let i = 0; i < dadosFat.length; i++) {
    let linha = dadosFat[i];
    if (linha.length < 5) continue; 
    
    // O código de domínio vem na coluna B do CSV (índice 1)
    let codDominio = String(linha[1]).trim(); 
    if (codDominio !== "" && !isNaN(codDominio)) {
      ultimoCodigoMemorizado = codDominio;
    }
    
    let isLinhaTotal = false;
    for (let col = 0; col < linha.length - 1; col++) {
      if (String(linha[col]).trim().toUpperCase() === "TOTAL") {
        isLinhaTotal = true;
        break;
      }
    }
    
    if (isLinhaTotal && ultimoCodigoMemorizado !== "") {
      // Pega a última coluna preenchida
      let totalStr = String(linha[linha.length - 1]).trim();
      if (totalStr !== "") {
        mapaFaturamento[ultimoCodigoMemorizado] = totalStr;
        ultimoCodigoMemorizado = ""; 
      }
    }
  }
  
  const dadosSheet = sheet.getDataRange().getValues();
  const cabecalhoSheet = dadosSheet[0];
  const idxCodSheet = buscarIndice(cabecalhoSheet, "Código Dominio");
  const idxDomSheet = buscarIndice(cabecalhoSheet, "DOMINIO"); 
  
  if (idxCodSheet === -1 || idxDomSheet === -1) {
    throw new Error("Não foi possível encontrar as colunas 'Código Dominio' ou 'DOMINIO' na aba atual.");
  }
  
  const atualizacoes = [];
  let qtdAtualizados = 0;
  
  for (let i = 1; i < dadosSheet.length; i++) {
    let codSheet = String(dadosSheet[i][idxCodSheet]).trim();
    
    if (mapaFaturamento[codSheet] !== undefined) {
      atualizacoes.push([mapaFaturamento[codSheet]]);
      qtdAtualizados++;
    } else {
      atualizacoes.push([dadosSheet[i][idxDomSheet]]);
    }
  }
  
  sheet.getRange(2, idxDomSheet + 1, atualizacoes.length, 1).setValues(atualizacoes);
  ss.toast(`Faturamento importado para ${qtdAtualizados} empresas!`, "✅ Sucesso", 10);
}

// =============================================================================
// FUNÇÃO 3: CONFERÊNCIA FATURAMENTO (FAT x DOMÍNIO)
// =============================================================================
function conferirFaturamento() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const nomeAba = sheet.getName();

  if (nomeAba === "empresas" || nomeAba === "BASE_ANTECIPACAO") {
    ui.alert("⚠️ Vá para a aba da competência (ex: 06/2026) antes de rodar a conferência.");
    return;
  }

  const dadosSheet = sheet.getDataRange().getValues();
  if (dadosSheet.length < 2) return;

  const cabecalho = dadosSheet[0];
  const idxFat = buscarIndice(cabecalho, "FAT");
  const idxDom = buscarIndice(cabecalho, "DOMINIO");
  const idxImportacao = buscarIndice(cabecalho, "IMPORTAÇÃO");

  if (idxFat === -1 || idxDom === -1 || idxImportacao === -1) {
    ui.alert("❌ Erro", "As colunas 'FAT', 'DOMINIO' ou 'IMPORTAÇÃO' não foram encontradas.", ui.ButtonSet.OK);
    return;
  }

  const atualizacoesImportacao = [];
  let qtdProcessados = 0;

  for (let i = 1; i < dadosSheet.length; i++) {
    let strFat = String(dadosSheet[i][idxFat]).trim().toUpperCase();
    let strDom = String(dadosSheet[i][idxDom]).trim().toUpperCase();
    
    let valFat = parseDinheiro(strFat);
    let valDom = parseDinheiro(strDom);
    
    let statusFinal = "";

    if (valFat === "SF") {
      if (valDom === 0) {
        statusFinal = "SF";
      } else {
        statusFinal = "DIVERGENTE";
      }
    } else if (valDom === 0 && valFat > 0) {
      statusFinal = "PENDENTE";
    } else if (Math.abs(valFat - valDom) > 0.01) {
      statusFinal = "DIVERGENTE";
    } else if (valFat === valDom) {
      statusFinal = "OK";
    }

    atualizacoesImportacao.push([statusFinal]);
    qtdProcessados++;
  }

  sheet.getRange(2, idxImportacao + 1, atualizacoesImportacao.length, 1).setValues(atualizacoesImportacao);
  ui.alert("📊 Conferência Concluída!", "O cruzamento entre FAT e DOMÍNIO foi realizado para " + qtdProcessados + " empresas.", ui.ButtonSet.OK);
}

// =============================================================================
// FUNÇÃO 4: VARREDURA DA PASTA DO DRIVE
// =============================================================================
function checarArquivosDrive() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const nomeAba = sheet.getName();
  
  if (nomeAba === "empresas") {
    ui.alert("⚠️ Aba Incorreta", "Por favor, execute este botão estando na aba da competência (ex: 06/2026).", ui.ButtonSet.OK);
    return;
  }

  const dados = sheet.getDataRange().getValues();
  if (dados.length < 2) return;

  const cabecalho = dados[0];
  const idxCod = buscarIndice(cabecalho, "Código Dominio");
  const idxDas = buscarIndice(cabecalho, "DAS");
  const idxExtrato = buscarIndice(cabecalho, "EXTRATO");

  if (idxCod === -1 || idxDas === -1 || idxExtrato === -1) {
    ui.alert("❌ Erro", "As colunas 'Código Dominio', 'DAS' ou 'EXTRATO' não foram localizadas nesta aba.", ui.ButtonSet.OK);
    return;
  }

  const pastaId = "ID_EXEMPLO";
  let pasta;
  try {
    pasta = DriveApp.getFolderById(pastaId);
  } catch (e) {
    ui.alert("❌ Erro de Acesso", "Não foi possível acessar a pasta. Verifique permissões.", ui.ButtonSet.OK);
    return;
  }

  ss.toast("Lendo arquivos do Drive...", "Aguarde");

  const arquivos = pasta.getFiles();
  const listaNomes = [];
  while (arquivos.hasNext()) {
    listaNomes.push(arquivos.next().getName());
  }

  const atualizacoesDas = [];
  const atualizacoesExtrato = [];
  let processados = 0;

  for (let i = 1; i < dados.length; i++) {
    let cod = String(dados[i][idxCod]).trim();
    
    if (!cod || cod === "") {
      atualizacoesDas.push([""]);
      atualizacoesExtrato.push([""]);
      continue;
    }

    let padraoDas = "GuiaDAS_" + cod + "_";
    let padraoExtrato = "PGDASD-EXTRATO_" + cod + "_";
    
    let achouDAS = listaNomes.some(nome => nome.includes(padraoDas));
    let achouExtrato = listaNomes.some(nome => nome.includes(padraoExtrato));

    atualizacoesDas.push([achouDAS ? "ENCONTRADO" : "PENDENTE"]);
    atualizacoesExtrato.push([achouExtrato ? "ENCONTRADO" : "PENDENTE"]);
    processados++;
  }

  sheet.getRange(2, idxDas + 1, atualizacoesDas.length, 1).setValues(atualizacoesDas);
  sheet.getRange(2, idxExtrato + 1, atualizacoesExtrato.length, 1).setValues(atualizacoesExtrato);

  ui.alert("✅ Varredura Concluída!", processados + " empresas foram verificadas no Drive.", ui.ButtonSet.OK);
}