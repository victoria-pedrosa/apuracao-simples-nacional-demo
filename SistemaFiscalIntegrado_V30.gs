/**
 * =============================================================================
 * SISTEMA FISCAL INTEGRADO CONSOLIDADO - V30 (INTEGRAÇÃO ACESSÓRIAS)
 * Contém: Importação Seletiva, Extração RBT12, Alíquotas, PDFs e Acessórias
 * Modificação: Nova coluna "ENVIO" e motor de importação de CSV/Planilha.
 * =============================================================================
 */

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('⚙️ Sistema Fiscal Integrado')
      .addSubMenu(ui.createMenu('📥 Importação de Dados')
          .addItem('🔍 Filtrar e Importar Empresas', 'importarEmpresasFiscal'))
      .addSeparator()
      .addSubMenu(ui.createMenu('📊 RBT12 (Aba Ativa)')
          .addItem('▶️ Extrair Dados (Extratos PDF)', 'extrairSelecionados')
          .addItem('🕵️ Detetive: Checar Pastas e Arquivos', 'detetivePastas'))
      .addSeparator()
      .addSubMenu(ui.createMenu('🧮 Alíquotas')
          .addItem('📈 Calcular Alíquotas Clientes (Simples Nacional)', 'gerarAliquotasClientes'))
      .addSeparator()
      .addSubMenu(ui.createMenu('📄 Geração de PDFs')
          .addItem('📋 Gerar PDFs - Apenas ISS', 'gerarPDF_ISS_Final')
          .addItem('💰 Gerar PDFs - Conjunto (ISS/ICMS e DAS)', 'gerarPDF_ISS_DAS_Final'))
      .addSeparator()
      .addSubMenu(ui.createMenu('📦 Integração Acessórias')
          .addItem('✔️ Marcar Envios do Acessórias', 'marcarAcessorias'))
      .addToUi();
}

// =============================================================================
// MOTOR DE SELEÇÃO INTELIGENTE
// =============================================================================
function obterLinhasParaProcessar(sheet, headers) {
  var linhas = [];
  var allData = sheet.getDataRange().getValues();
  
  var idxSelecionar = buscarIndiceColuna(headers, "SELECIONAR");
  if (idxSelecionar !== -1) {
    for (var i = 1; i < allData.length; i++) {
      if (allData[i][idxSelecionar] === true) {
        linhas.push(i + 1);
      }
    }
  }
  
  if (linhas.length === 0) {
    var rangeList = sheet.getActiveRangeList();
    if (rangeList) {
      var ranges = rangeList.getRanges();
      for (var i = 0; i < ranges.length; i++) {
        var start = ranges[i].getRow();
        var numRows = ranges[i].getNumRows();
        for (var j = 0; j < numRows; j++) {
          var row = start + j;
          if (row > 1 && linhas.indexOf(row) === -1) { 
            linhas.push(row);
          }
        }
      }
    }
  }
  return linhas.sort(function(a, b) { return a - b; }); 
}

function buscarIndiceColuna(headers, nomeDesejado) {
  function normalizar(txt) {
    if (!txt) return "";
    return String(txt).toLowerCase()
      .normalize("NFD").replace(/[\u0300-\u036f]/g, "") 
      .replace(/[^a-z0-9]/g, ""); 
  }
  let nNome = normalizar(nomeDesejado);
  let idx = headers.findIndex(h => normalizar(h) === nNome);
  if (idx !== -1) return idx;
  return headers.findIndex(h => {
    let hNorm = normalizar(h);
    if (!hNorm) return false;
    return hNorm.includes(nNome) || nNome.includes(hNorm);
  });
}

function garantirColuna(sheet, headers, nomeColuna) {
  let idx = buscarIndiceColuna(headers, nomeColuna);
  if (idx === -1) {
    let lastCol = sheet.getLastColumn();
    sheet.getRange(1, lastCol + 1).setValue(nomeColuna)
         .setFontWeight("bold")
         .setBackground("#274e13")
         .setFontColor("#ffffff");
    headers.push(nomeColuna);
    return lastCol;
  }
  return idx;
}

function formatarNumeroParaPDF(valor) {
  if (!valor || valor === "-") return "0,00";
  let texto = String(valor).replace("%", "").trim();
  if (!isNaN(parseFloat(texto.replace(",", ".")))) {
     let num = parseFloat(texto.replace(",", "."));
     if (num > 0 && num < 1 && String(valor).indexOf(",") === -1) {
       num = num * 100;
     }
     return num.toFixed(2).replace(".", ",");
  }
  return texto;
}

// =============================================================================
// MÓDULO 1: IMPORTAÇÃO E FILTRAGEM
// =============================================================================
function importarEmpresasFiscal() {
  const urlFonte = "https://docs.google.com/spreadsheets/d/ID_EXEMPLO/edit";
  const nomeAbaFonte = "empresas";
  const ssDestino = SpreadsheetApp.getActiveSpreadsheet();
  
  const colunasDesejadas = [
    "Código Dominio", "Situacao", "Empresa", "CNPJ", "E-mail", 
    "Matriz / Filial", "Tributacao", "Fiscal Responsavel", "DATA DE ABERTURA", "DATA EXEMPLO"
  ];
  
  // MUDANÇA: Inclusão da Coluna ENVIO automaticamente
  const colunasAutomáticas = [
    "ANEXO EXTRATO", "RBT12 EXTRATO", "STATUS RBT12", 
    "ALÍQUOTA ISS", "ALÍQUOTA ICMS", "ALÍQUOTA DAS", "COMPETÊNCIA", 
    "VENCIMENTO DAS", "STATUS PDF ISS", "STATUS PDF DAS", "ENVIO"
  ];
  const cabecalhoCompleto = ["☑️ SELECIONAR", ...colunasDesejadas, ...colunasAutomáticas];
  
  let ssFonte;
  try {
    ssFonte = SpreadsheetApp.openByUrl(urlFonte);
  } catch (e) {
    SpreadsheetApp.getUi().alert("❌ Erro ao acessar a planilha fonte: " + e.message);
    return;
  }
  
  const abaFonte = ssFonte.getSheetByName(nomeAbaFonte);
  if (!abaFonte) {
    SpreadsheetApp.getUi().alert("❌ Erro: Aba '" + nomeAbaFonte + "' não encontrada na fonte.");
    return;
  }
  
  const lastRow = abaFonte.getLastRow();
  if (lastRow < 2) return;
  
  const dados = abaFonte.getRange(1, 1, lastRow, abaFonte.getLastColumn()).getValues();
  const cabecalhoOriginal = dados[0];
  
  const mapeamentoIndices = colunasDesejadas.map(col => {
    let idx = cabecalhoOriginal.findIndex(h => h.toString().toLowerCase().trim() === col.toLowerCase().trim());
    if (idx === -1) idx = buscarIndiceColuna(cabecalhoOriginal, col);
    return idx;
  });
  
  const idxTributacao = buscarIndiceColuna(cabecalhoOriginal, "Tributacao");
  const linhasFiltradas = [];
  const indiceColunaO = 14; 
  
  for (let i = 1; i < dados.length; i++) {
    let AppLinha = dados[i];
    let valorO = String(AppLinha[indiceColunaO]).toLowerCase().trim();
    let tributacao = idxTributacao !== -1 ? String(AppLinha[idxTributacao]).toUpperCase().trim() : "";
    
    let ehFiscal4 = valorO.indexOf("fiscal 4") !== -1;
    let ehFiscal5 = valorO.indexOf("fiscal 5") !== -1;
    let ehFiscal2 = valorO.indexOf("fiscal 2") !== -1;
    
    let ehSimplesNacional = tributacao === "S" || tributacao === "SN" || tributacao.includes("SIMPLES");
    
    if (ehFiscal4 || ehFiscal5 || (ehFiscal2 && ehSimplesNacional)) {
      let novaLinhaSeletiva = mapeamentoIndices.map(idx => idx !== -1 ? AppLinha[idx] : "");
      let linhaFinal = [false, ...novaLinhaSeletiva];
      colunasAutomáticas.forEach(() => linhaFinal.push(""));
      linhasFiltradas.push(linhaFinal);
    }
  }
  
  if (linhasFiltradas.length === 0) {
    SpreadsheetApp.getUi().alert("⚠️ Nenhuma empresa compatível foi encontrada na fonte.");
    return;
  }
  
  let abaDestino = ssDestino.getSheetByName("EMPRESAS_FILTRADAS");
  if (!abaDestino) abaDestino = ssDestino.insertSheet("EMPRESAS_FILTRADAS");
  else abaDestino.clear();
  
  abaDestino.getRange(1, 1, 1, cabecalhoCompleto.length).setValues([cabecalhoCompleto]);
  abaDestino.getRange(2, 1, linhasFiltradas.length, cabecalhoCompleto.length).setValues(linhasFiltradas);
  abaDestino.getRange(2, 1, linhasFiltradas.length, 1).insertCheckboxes();
  
  abaDestino.getRange(1, 1, 1, cabecalhoCompleto.length).setFontWeight("bold").setBackground("#274e13").setFontColor("#ffffff");
  const totalLinhasTabela = linhasFiltradas.length + 1;
  for (let r = 3; r <= totalLinhasTabela; r += 2) {
    abaDestino.getRange(r, 1, 1, cabecalhoCompleto.length).setBackground("#efefef");
  }
  
  SpreadsheetApp.getUi().alert("✅ Sucesso! Aba 'EMPRESAS_FILTRADAS' atualizada.");
}

// =============================================================================
// MÓDULO 2: APURAÇÃO DE RBT12
// =============================================================================
function parseMoney(str) {
  if (!str) return 0;
  var limpo = str.toString().replace(/[^\d,.-]/g, '').replace(/\./g, '').replace(',', '.');
  return parseFloat(limpo) || 0;
}

function formatMoney(num) {
  return "R$ " + num.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function obterPastaPrincipal() {
  var id = "ID_EXEMPLO";
  try { return DriveApp.getFolderById(id); } catch (e) {
    var busca = DriveApp.getFoldersByName("1. Impostos");
    if (busca.hasNext()) return busca.next();
  }
  return null;
}

function encontrarPastaMes(pastaEmpresa, variacoes, anoPAStr) {
  var subPastas = pastaEmpresa.getFolders();
  var pastasAnoEncontradas = [];

  while (subPastas.hasNext()) {
    var sub = subPastas.next();
    var nome = sub.getName().toUpperCase().replace(/\s+/g, '');
    for (var i = 0; i < variacoes.length; i++) {
      if (nome.indexOf(variacoes[i]) !== -1) return sub;
    }
    if (anoPAStr && nome.indexOf(anoPAStr) !== -1) {
      pastasAnoEncontradas.push(sub);
    }
  }

  for (var k = 0; k < pastasAnoEncontradas.length; k++) {
    var subPastasAno = pastasAnoEncontradas[k].getFolders();
    while (subPastasAno.hasNext()) {
      var subMes = subPastasAno.next();
      var nomeMes = subMes.getName().toUpperCase().replace(/\s+/g, '');
      for (var j = 0; j < variacoes.length; j++) {
        if (nomeMes.indexOf(variacoes[j]) !== -1) return subMes;
      }
    }
  }
  return null;
}

function buscarPastaEmpresa(pastaPai, nomeEmpresa) {
  var nomeLimpo = nomeEmpresa.toString().replace(/[^\w\sÀ-ÿ]/gi, ' ').replace(/\s+/g, ' ').trim();
  var palavras = nomeLimpo.split(" ").slice(0, 3);
  if (palavras.length > 0) {
    var termoJunto = palavras.join(" ");
    var buscaJunta = pastaPai.searchFolders("title contains '" + termoJunto + "'");
    if (buscaJunta.hasNext()) return buscaJunta.next();
    var querySeparada = palavras.map(function(p) { return "title contains '" + p + "'"; }).join(" and ");
    var buscaSeparada = pastaPai.searchFolders(querySeparada);
    if (buscaSeparada.hasNext()) return buscaSeparada.next();
  }
  return null;
}

function buscarExtrato(pastaPai, pastaEmpresa, codigoDominio, compLimpa, variacoes, anoPAStr) {
  if (pastaEmpresa) {
    var pastaMes = encontrarPastaMes(pastaEmpresa, variacoes, anoPAStr);
    var alvo = pastaMes || pastaEmpresa;
    var arquivos = alvo.getFilesByType(MimeType.PDF);
    var fallback = null;
    
    while (arquivos.hasNext()) {
      var arq = arquivos.next();
      var nome = arq.getName().toUpperCase();
      if (nome.includes("RECIBO") && !nome.includes("EXTRATO") && !nome.includes("DECLARA")) continue;
      if (nome.includes("PGDASD") || nome.includes("EXTRATO") || nome.includes("DECLARA")) {
        if (nome.includes("PGDASD") || nome.includes("EXTRATO")) return arq;
        if (!fallback) fallback = arq;
      }
    }
    if (fallback) return fallback;
  }
  if (codigoDominio) {
    var arqGlobal = pastaPai.searchFiles("title contains '" + codigoDominio + "' and mimeType = 'application/pdf'");
    while (arqGlobal.hasNext()) {
      var arqG = arqGlobal.next();
      var n = arqG.getName().toUpperCase();
      if (n.includes("PGDASD") || n.includes("EXTRATO")) return arqG;
    }
  }
  return null;
}

function obterDadosCompetencia() {
  var ui = SpreadsheetApp.getUi();
  var prompt = ui.prompt('📅 Seleção de Competência', 'Digite o mês/ano do extrato no formato MM/AAAA\n(Exemplo: 05/2026):', ui.ButtonSet.OK_CANCEL);
  
  if (prompt.getSelectedButton() !== ui.Button.OK) return null;
  
  var input = prompt.getResponseText().trim();
  if (!/^\d{2}\/\d{4}$/.test(input)) {
    ui.alert('❌ Erro: Formato inválido! Use o formato MM/AAAA (Ex: 05/2026).');
    return null;
  }
  
  var partes = input.split('/');
  var mes = partes[0];
  var ano = partes[1];
  
  var compLimpa = mes + ano; 
  var variacoes = [compLimpa, ano + mes, mes + "-" + ano, ano + "-" + mes, mes + "." + ano];
  
  var mesesNomes = ["JANEIRO", "FEVEREIRO", "MARÇO", "ABRIL", "MAIO", "JUNHO", "JULHO", "AGOSTO", "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO"];
  var nomeMes = mesesNomes[parseInt(mes, 10) - 1];
  if (nomeMes) variacoes.push(nomeMes);
  
  var dataRef = new Date(parseInt(ano, 10), parseInt(mes, 10) - 1, 1);
  dataRef.setMonth(dataRef.getMonth() - 12);
  var mesAntigo = ("0" + (dataRef.getMonth() + 1)).slice(-2);
  var anoAntigo = dataRef.getFullYear();
  var textoAbatimento = mesAntigo + "/" + anoAntigo; 

  return { 
    compLimpa: compLimpa, 
    variacoes: variacoes, 
    textoAbatimento: textoAbatimento, 
    inputOriginal: input,
    mesPA: parseInt(mes, 10),
    anoPA: parseInt(ano, 10)
  };
}

function detetivePastas() {
  var ui = SpreadsheetApp.getUi();
  var sheet = SpreadsheetApp.getActiveSheet();
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  var linhasSel = obterLinhasParaProcessar(sheet, headers);
  if (linhasSel.length === 0) {
    ui.alert("❌ Nenhuma empresa selecionada. Marque as caixinhas na coluna A ou selecione com o mouse.");
    return;
  }

  var compDados = obterDadosCompetencia();
  if (!compDados) return; 

  var idxEmpresa = buscarIndiceColuna(headers, "Empresa");
  var idxCodigo = buscarIndiceColuna(headers, "Código Dominio");
  
  if (idxCodigo === -1 && headers.length > 0) idxCodigo = 1; 
  if (idxEmpresa === -1) {
    ui.alert("❌ Erro: Coluna 'Empresa' não identificada na aba atual.");
    return;
  }

  var pastaPai = obterPastaPrincipal();
  if (!pastaPai) { ui.alert("Erro de acesso ao Drive."); return; }
  
  var allData = sheet.getDataRange().getValues();
  var alertas = [];
  
  for (var i = 0; i < linhasSel.length; i++) {
    var row = linhasSel[i];
    var linha = allData[row - 1];
    
    var nome = linha[idxEmpresa], cod = linha[idxCodigo];
    if (!nome) continue;
    
    var pasta = buscarPastaEmpresa(pastaPai, nome);
    var pdf = buscarExtrato(pastaPai, pasta, cod, compDados.compLimpa, compDados.variacoes, String(compDados.anoPA));
    
    if (!pasta) alertas.push("❌ Linha " + row + " (" + nome + "): Pasta não encontrada.");
    else if (!pdf) alertas.push("⚠️ Linha " + row + " (" + nome + "): Arquivo não achado.");
  }
  ui.alert("Resultado", alertas.length > 0 ? alertas.join("\n") : "✅ Tudo OK nas linhas selecionadas!", ui.ButtonSet.OK);
}

function gravarStatusDinamico(sheet, row, idxAnexo, idxRbt, idxStatus, idxComp, txtAnexo, valorRBT, status, compTexto) {
  sheet.getRange(row, idxAnexo + 1).setValue(txtAnexo);
  sheet.getRange(row, idxRbt + 1).setValue(valorRBT);
  sheet.getRange(row, idxStatus + 1).setValue(status);
  if (idxComp !== -1 && compTexto) sheet.getRange(row, idxComp + 1).setValue(compTexto);
  SpreadsheetApp.flush();
}

function marcarEmProcessamentoDinamico(sheet, row, idxAnexo, idxRbt, idxStatus) {
  sheet.getRange(row, idxAnexo + 1).setValue("⏳ Processando...");
  sheet.getRange(row, idxRbt + 1).setValue("");
  sheet.getRange(row, idxStatus + 1).setValue("");
  SpreadsheetApp.flush();
}

function extrairSelecionados() {
  var ui = SpreadsheetApp.getUi();
  var sheet = SpreadsheetApp.getActiveSheet();
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  var idxEmpresa = buscarIndiceColuna(headers, "Empresa");
  var idxCodigo = buscarIndiceColuna(headers, "Código Dominio");
  var idxAbertura = buscarIndiceColuna(headers, "DATA DE ABERTURA");
  
  if (idxCodigo === -1 && headers.length > 0) idxCodigo = 1;
  if (idxEmpresa === -1) {
    ui.alert("❌ Erro: Coluna essencial 'Empresa' não encontrada na planilha.");
    return;
  }

  var idxAnexo = garantirColuna(sheet, headers, "ANEXO EXTRATO");
  var idxRbt = garantirColuna(sheet, headers, "RBT12 EXTRATO");
  var idxStatus = garantirColuna(sheet, headers, "STATUS RBT12");
  var idxComp = garantirColuna(sheet, headers, "COMPETÊNCIA");

  headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  var linhasSel = obterLinhasParaProcessar(sheet, headers);
  if (linhasSel.length === 0) {
    ui.alert("❌ Nenhuma empresa selecionada. Marque as caixinhas na coluna A ou selecione com o mouse.");
    return;
  }

  var compDados = obterDadosCompetencia();
  if (!compDados) return; 

  var pastaPai = obterPastaPrincipal();
  if (!pastaPai) return;

  var allData = sheet.getDataRange().getValues();

  for (var i = 0; i < linhasSel.length; i++) {
    var row = linhasSel[i];
    var linha = allData[row - 1];
    
    var nome = linha[idxEmpresa], cod = linha[idxCodigo];
    if (!nome) continue;

    marcarEmProcessamentoDinamico(sheet, row, idxAnexo, idxRbt, idxStatus);
    var tempFileId = null;

    try {
      var pasta = buscarPastaEmpresa(pastaPai, nome);
      var pdf   = buscarExtrato(pastaPai, pasta, cod, compDados.compLimpa, compDados.variacoes, String(compDados.anoPA));

      if (!pdf) {
        gravarStatusDinamico(sheet, row, idxAnexo, idxRbt, idxStatus, idxComp, "❌ PDF não encontrado", "", "ERRO", compDados.inputOriginal);
        continue;
      }

      var tempFile = Drive.Files.copy(
        { title: "Temp_" + nome, mimeType: MimeType.GOOGLE_DOCS },
        pdf.getId(),
        { convert: true, supportsAllDrives: true }
      );
      tempFileId = tempFile.id;

      var doc   = DocumentApp.openById(tempFileId);
      var texto = doc.getBody().getText();
      var textoLimpo = texto.replace(/\r?\n/g, " ");

      var rpa = 0, rbt12 = 0;
      var mRPA = texto.match(/Receita Bruta do PA \(RPA\)[\s\S]{0,50}?(?:Competência|Caixa)[\s\S]{0,30}?([\d\.]+(?:,\d{2}))/i);
      if (!mRPA) mRPA = texto.match(/Receita Bruta do PA \(RPA\)[\s\S]{0,80}?([\d\.]+(?:,\d{2}))/i);
      if (mRPA) rpa = parseMoney(mRPA[1]);

      var mRBT = texto.match(/RBT12\)[\s\S]{0,40}?([\d\.]+(?:,\d{2}))/i);
      if (mRBT) rbt12 = parseMoney(mRBT[1]);

      var idadeMeses = 13; 
      var matchAbertura = texto.match(/abertura(?: no CNPJ)?:?\s*(\d{2})\/(\d{2})\/(\d{4})/i);
      
      if (matchAbertura) {
        var mesAb = parseInt(matchAbertura[2], 10);
        var anoAb = parseInt(matchAbertura[3], 10);
        idadeMeses = (compDados.anoPA * 12 + compDados.mesPA) - (anoAb * 12 + mesAb) + 1;
      } else if (idxAbertura !== -1 && linha[idxAbertura]) {
        var dataSheet = linha[idxAbertura];
        if (dataSheet instanceof Date) {
          var mesAb = dataSheet.getMonth() + 1;
          var anoAb = dataSheet.getFullYear();
          idadeMeses = (compDados.anoPA * 12 + compDados.mesPA) - (anoAb * 12 + mesAb) + 1;
        } else {
          var stringData = String(dataSheet).match(/(\d{2})\/(\d{2})\/(\d{4})/);
          if (stringData) {
            var mesAb = parseInt(stringData[2], 10);
            var anoAb = parseInt(stringData[3], 10);
            idadeMeses = (compDados.anoPA * 12 + compDados.mesPA) - (anoAb * 12 + mesAb) + 1;
          }
        }
      }

      var finalRBTValue = 0;
      var statusConf    = "OK";

      if (idadeMeses < 12) {
        if (idadeMeses <= 0) idadeMeses = 1; 
        finalRBTValue = ((rbt12 + rpa) / idadeMeses) * 12;
        statusConf    = "OK-Prop (Projetado)";
      } else {
        var antigo = 0;
        var blocoReceitas = texto.match(/Receitas Brutas Anteriores[\s\S]{0,1500}/i);
        if (blocoReceitas) {
          var escapedAbat = compDados.textoAbatimento.replace("/", "\\/");
          var rAnt = new RegExp(escapedAbat + "[\\s\\S]{0,40}?([\\d\\.]+(?:,\\d{2}))", "i");
          var m = rAnt.exec(blocoReceitas[0]);
          if (m) {
            antigo = parseMoney(m[1]);
          }
        }
        finalRBTValue = (rbt12 + rpa) - antigo;
      }

      var anexos = [];
      var mAnx, rAnx = /Anexo\s+(VI|V|IV|III|II|I)\b/gi;
      while ((mAnx = rAnx.exec(texto)) !== null) {
        var nomeAnx = mAnx[1].toUpperCase(); 
        if (anexos.indexOf(nomeAnx) === -1) anexos.push(nomeAnx);
      }

      if (/Revenda de mercadoria/i.test(textoLimpo) && anexos.indexOf("I") === -1) {
        anexos.push("I");
      }
      if (/Venda de mercadorias industrializadas/i.test(textoLimpo) && anexos.indexOf("II") === -1) {
        anexos.push("II");
      }
      if ((/Serviços de comunicação/i.test(textoLimpo) || /transporte intermunicipal/i.test(textoLimpo) || /transporte interestadual/i.test(textoLimpo)) && anexos.indexOf("III") === -1) {
        anexos.push("III");
      }

      var txtAnexo = anexos.length > 0
        ? anexos.join(", ")
        : (pdf.getName().indexOf("DECLARA") !== -1 ? "SEM FATURAMENTO" : "NÃO ENCONTRADO");

      gravarStatusDinamico(sheet, row, idxAnexo, idxRbt, idxStatus, idxComp, txtAnexo, formatMoney(finalRBTValue), statusConf, compDados.inputOriginal);

    } catch (e) {
      gravarStatusDinamico(sheet, row, idxAnexo, idxRbt, idxStatus, idxComp, "❌ ERRO: " + e.message, "", "ERRO", compDados.inputOriginal);
    } finally {
      if (tempFileId) {
        try { DriveApp.getFileById(tempFileId).setTrashed(true); } catch (e2) {}
      }
    }
  }
  ui.alert("✅ Processamento Concluído nas linhas marcadas!");
}

// =============================================================================
// MÓDULO 3: CÁLCULO DE ALÍQUOTAS
// =============================================================================
function encontrarChaveAnexo(txt) {
  if (!txt) return null;
  let t = String(txt).toUpperCase();
  if (t.includes("III") || t.match(/\bIII\b/)) return "Anexo III";
  if (t.includes("IV") || t.match(/\bIV\b/)) return "Anexo IV";
  if (t.includes("II") || t.match(/\bII\b/)) return "Anexo II";
  if (t.includes("V") || t.match(/\bV\b/)) return "Anexo V";
  if (t.includes("I") || t.match(/\bI\b/)) return "Anexo I";
  return null;
}

function gerarAliquotasClientes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetClientes = ss.getActiveSheet(); 
  const allData = sheetClientes.getDataRange().getValues();
  const headers = allData[0];
  
  var linhasSel = obterLinhasParaProcessar(sheetClientes, headers);
  if (linhasSel.length === 0) {
    SpreadsheetApp.getUi().alert("❌ Marque as caixinhas na coluna A ou selecione com o mouse as linhas das empresas.");
    return;
  }
  
  var idxAnexo = buscarIndiceColuna(headers, "ANEXO EXTRATO");
  var idxRbt = buscarIndiceColuna(headers, "RBT12 EXTRATO");

  if (idxAnexo === -1 || idxRbt === -1) {
    SpreadsheetApp.getUi().alert("❌ Erro: Colunas de ANEXO ou RBT12 não encontradas.");
    return;
  }

  var idxISS = garantirColuna(sheetClientes, headers, "ALÍQUOTA ISS");
  var idxICMS = garantirColuna(sheetClientes, headers, "ALÍQUOTA ICMS");
  var idxDAS = garantirColuna(sheetClientes, headers, "ALÍQUOTA DAS");

  const TABELAS = {
    "Anexo I": [
      { max: 180000,    nom: 0.040, ded: 0,      iss: 0, icms: 0.340 },
      { max: 360000,    nom: 0.073, ded: 5940,   iss: 0, icms: 0.340 },
      { max: 720000,    nom: 0.095, ded: 13860,  iss: 0, icms: 0.340 },
      { max: 1800000,   nom: 0.107, ded: 22500,  iss: 0, icms: 0.340 },
      { max: 3600000,   nom: 0.143, ded: 87300,  iss: 0, icms: 0.340 },
      { max: 4800000,   nom: 0.190, ded: 378000, iss: 0, icms: 0     }
    ],
    "Anexo II": [
      { max: 180000,    nom: 0.045, ded: 0,      iss: 0, icms: 0.320 },
      { max: 360000,    nom: 0.078, ded: 5940,   iss: 0, icms: 0.320 },
      { max: 720000,    nom: 0.100, ded: 13860,  iss: 0, icms: 0.320 },
      { max: 1800000,   nom: 0.112, ded: 22500,  iss: 0, icms: 0.320 },
      { max: 3600000,   nom: 0.147, ded: 85500,  iss: 0, icms: 0.320 },
      { max: 4800000,   nom: 0.300, ded: 720000, iss: 0, icms: 0     }
    ],
    "Anexo III": [
      { max: 180000,    nom: 0.060, ded: 0,      iss: 0.335, icms: 0 },
      { max: 360000,    nom: 0.112, ded: 9360,   iss: 0.320, icms: 0 },
      { max: 720000,    nom: 0.135, ded: 17640,  iss: 0.325, icms: 0 },
      { max: 1800000,   nom: 0.160, ded: 35640,  iss: 0.325, icms: 0 },
      { max: 3600000,   nom: 0.210, ded: 125640, iss: 0.335, icms: 0 },
      { max: 4800000,   nom: 0.330, ded: 648000, iss: 0,     icms: 0 }
    ],
    "Anexo IV": [
      { max: 180000,    nom: 0.045, ded: 0,      iss: 0.445, icms: 0 },
      { max: 360000,    nom: 0.090, ded: 8100,   iss: 0.400, icms: 0 },
      { max: 720000,    nom: 0.102, ded: 12420,  iss: 0.400, icms: 0 },
      { max: 1800000,   nom: 0.140, ded: 39780,  iss: 0.400, icms: 0 },
      { max: 3600000,   nom: 0.220, ded: 183780, iss: 0.434, icms: 0 },
      { max: 4800000,   nom: 0.330, ded: 828000, iss: 0,     icms: 0 }
    ],
    "Anexo V": [
      { max: 180000,    nom: 0.155, ded: 0,      iss: 0.140, icms: 0 },
      { max: 360000,    nom: 0.180, ded: 4500,   iss: 0.170, icms: 0 },
      { max: 720000,    nom: 0.195, ded: 9900,   iss: 0.190, icms: 0 },
      { max: 1800000,   nom: 0.205, ded: 17100,  iss: 0.210, icms: 0 },
      { max: 3600000,   nom: 0.230, ded: 62100,  iss: 0.235, icms: 0 },
      { max: 4800000,   nom: 0.305, ded: 540000, iss: 0,     icms: 0 }
    ]
  };

  const TETO_TRIBUTO = 0.05;

  for (let i = 0; i < linhasSel.length; i++) {
    let row = linhasSel[i];
    let linha = allData[row - 1];
    
    const txtAnexo = String(linha[idxAnexo]).trim();
    const anexoNome = encontrarChaveAnexo(txtAnexo);

    let rbt12 = linha[idxRbt];
    if (typeof rbt12 === "string") {
      rbt12 = parseFloat(rbt12.replace(/R\$\s?/g, "").replace(/\./g, "").replace(",", ".").trim());
    }
    if (isNaN(rbt12) || rbt12 === "" || rbt12 === null) rbt12 = 0;

    let cellISS = sheetClientes.getRange(row, idxISS + 1);
    let cellICMS = sheetClientes.getRange(row, idxICMS + 1);
    let cellDAS = sheetClientes.getRange(row, idxDAS + 1);

    if (!anexoNome || !TABELAS[anexoNome] || rbt12 <= 0) {
      cellISS.setValue(""); cellICMS.setValue(""); cellDAS.setValue("");
      continue;
    }

    const faixa = TABELAS[anexoNome].find(function (f) { return rbt12 <= f.max; });

    if (!faixa) {
      cellISS.setValue("EXCEDEU LIMITE"); cellICMS.setValue("EXCEDEU LIMITE"); cellDAS.setValue("EXCEDEU LIMITE");
      continue;
    }

    const aliqDAS = ((rbt12 * faixa.nom) - faixa.ded) / rbt12;

    let aliqISS;
    if (faixa.iss > 0) {
      aliqISS = aliqDAS * faixa.iss;
      if (aliqISS > TETO_TRIBUTO) aliqISS = TETO_TRIBUTO;
    } else if (faixa.iss === 0 && anexoNome.match(/III|IV|V/)) {
      aliqISS = "NÃO DESTACADO";
    } else {
      aliqISS = "NÃO POSSUI";
    }

    let aliqICMS;
    if (faixa.icms > 0) {
      aliqICMS = aliqDAS * faixa.icms;
      if (aliqICMS > TETO_TRIBUTO) aliqICMS = TETO_TRIBUTO;
    } else if (faixa.icms === 0 && anexoNome.match(/I|II/)) {
      aliqICMS = "NÃO DESTACADO";
    } else {
      aliqICMS = "NÃO POSSUI";
    }

    cellISS.setValue(aliqISS);
    cellICMS.setValue(aliqICMS);
    cellDAS.setValue(aliqDAS);

    if (typeof aliqISS === "number") cellISS.setNumberFormat("0.0000%");
    else cellISS.setNumberFormat("@");
    
    if (typeof aliqICMS === "number") cellICMS.setNumberFormat("0.0000%");
    else cellICMS.setNumberFormat("@");
    
    if (typeof aliqDAS === "number") cellDAS.setNumberFormat("0.0000%");
    else cellDAS.setNumberFormat("@");
  }

  SpreadsheetApp.getUi().alert("✅ Alíquotas calculadas com sucesso!");
}

// =============================================================================
// MÓDULO 4: GERAÇÃO DE PDF - TIPO 1 (APENAS ISS)
// =============================================================================
function gerarPDF_ISS_Final() {
  const ID_MODELO_DOC = "ID_EXEMPLO"; 
  const ID_PASTA_DESTINO = "ID_EXEMPLO";
  const NOME_ABA = "EMPRESAS_FILTRADAS"; 
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName(NOME_ABA);
  if (!aba) { SpreadsheetApp.getUi().alert("❌ Erro: Aba 'EMPRESAS_FILTRADAS' não encontrada."); return; }

  const allData = aba.getDataRange().getValues();
  const headers = allData[0];
  
  var linhasSel = obterLinhasParaProcessar(aba, headers);
  if (linhasSel.length === 0) {
    SpreadsheetApp.getUi().alert("❌ Nenhuma empresa selecionada nas caixinhas ou com o mouse.");
    return;
  }
  
  const idxCnpj = buscarIndiceColuna(headers, "CNPJ");
  const idxEmpresa = buscarIndiceColuna(headers, "Empresa");
  const idxAnexo = buscarIndiceColuna(headers, "ANEXO EXTRATO");
  const idxIss = buscarIndiceColuna(headers, "ALÍQUOTA ISS");
  
  if (idxCnpj === -1 || idxEmpresa === -1 || idxAnexo === -1 || idxIss === -1) {
    SpreadsheetApp.getUi().alert("❌ Erro: Colunas bases não localizadas.");
    return;
  }
  
  const idxCompetencia = garantirColuna(aba, headers, "COMPETÊNCIA");
  const idxStatusPdf = garantirColuna(aba, headers, "STATUS PDF ISS");
  
  const dataHoje = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");
  
  let grupos = {};
  
  for (let i = 0; i < linhasSel.length; i++) {
    let row = linhasSel[i];
    let linha = allData[row - 1];
    
    let cnpj = String(linha[idxCnpj]);
    let empresa = String(linha[idxEmpresa]);
    let anexo = String(linha[idxAnexo]);
    let aliquota = linha[idxIss];
    let competencia = String(linha[idxCompetencia]).trim();
    let statusData = linha[idxStatusPdf];
    
    if (empresa && competencia && (statusData === "" || statusData === "(auto)")) {
      let chave = cnpj + "_" + competencia;
      if (!grupos[chave]) {
        grupos[chave] = { empresa: empresa, cnpj: cnpj, competencia: competencia, itens: [], linhasIndices: [] };
      }
      
      let issFormatado = formatarNumeroParaPDF(aliquota);
      grupos[chave].itens.push(`Anexo ${anexo} (${issFormatado}%)`);
      grupos[chave].linhasIndices.push(row); 
    }
  }
  
  let chaves = Object.keys(grupos);
  let processados = 0;
  if (chaves.length === 0) {
    SpreadsheetApp.getUi().alert("Nenhum PDF de ISS pendente nas linhas selecionadas.");
    return;
  }
  
  const pasta = DriveApp.getFolderById(ID_PASTA_DESTINO);
  const arquivoModelo = DriveApp.getFileById(ID_MODELO_DOC);
  
  for (let k = 0; k < chaves.length; k++) {
    let grupo = grupos[chaves[k]];
    let textoAnexos = grupo.itens.join("; "); 
    let compLimpa = grupo.competencia.replace(/\//g, '-');
    let nomeArquivo = `ALÍQUOTA DE ISS_${compLimpa}_${grupo.empresa}`;
    
    try {
      let copiaDoc = arquivoModelo.makeCopy("TEMP_" + nomeArquivo, pasta);
      let docId = copiaDoc.getId();
      Utilities.sleep(1500); 
      
      let doc = DocumentApp.openById(docId);
      let corpo = doc.getBody();
      
      corpo.replaceText("VAR_EMPRESA", grupo.empresa);
      corpo.replaceText("VAR_CNPJ", grupo.cnpj);
      corpo.replaceText("VAR_COMPETENCIA", grupo.competencia);
      corpo.replaceText("VAR_DATA", dataHoje); 
      corpo.replaceText("VAR_LISTA_ANEXOS", textoAnexos);
      corpo.replaceText("VAR_ANEXO", "");
      corpo.replaceText("VAR_ALIQUOTA", "");
      
      doc.saveAndClose();
      Utilities.sleep(2000); 
      
      let pdfBlob = copiaDoc.getAs(MimeType.PDF);
      pasta.createFile(pdfBlob).setName(nomeArquivo + ".pdf");
      DriveApp.getFileById(docId).setTrashed(true);
      
      grupo.linhasIndices.forEach(numLinha => { aba.getRange(numLinha, idxStatusPdf + 1).setValue(dataHoje); });
      processados++;
    } catch (e) {
      SpreadsheetApp.getUi().alert(`Erro na empresa ${grupo.empresa}: ${e.message}`);
    }
  }
  if (processados > 0) SpreadsheetApp.getUi().alert(`Sucesso! ${processados} documentos de ISS gerados.`);
}

// =============================================================================
// MÓDULO 5: GERAÇÃO DE PDF - TIPO 2 (CONJUNTO ISS/ICMS E DAS)
// =============================================================================
function gerarPDF_ISS_DAS_Final() {
  const ID_MODELO_DOC = "ID_EXEMPLO"; 
  const ID_PASTA_DESTINO = "ID_EXEMPLO";
  const NOME_ABA = "EMPRESAS_FILTRADAS"; 
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName(NOME_ABA);
  if (!aba) { SpreadsheetApp.getUi().alert("❌ Erro: Aba 'EMPRESAS_FILTRADAS' não encontrada."); return; }
  
  const allData = aba.getDataRange().getValues();
  const headers = allData[0];
  
  var linhasSel = obterLinhasParaProcessar(aba, headers);
  if (linhasSel.length === 0) {
    SpreadsheetApp.getUi().alert("❌ Nenhuma empresa selecionada nas caixinhas ou com o mouse.");
    return;
  }
  
  const idxCnpj = buscarIndiceColuna(headers, "CNPJ");
  const idxEmpresa = buscarIndiceColuna(headers, "Empresa");
  const idxAnexo = buscarIndiceColuna(headers, "ANEXO EXTRATO");
  const idxIss = buscarIndiceColuna(headers, "ALÍQUOTA ISS");
  const idxIcms = buscarIndiceColuna(headers, "ALÍQUOTA ICMS");
  const idxDas = buscarIndiceColuna(headers, "ALÍQUOTA DAS");
  
  if (idxCnpj === -1 || idxEmpresa === -1 || idxAnexo === -1 || idxIss === -1 || idxDas === -1) {
    SpreadsheetApp.getUi().alert("❌ Erro: Colunas bases não localizadas nesta aba para gerar PDF Conjunto.");
    return;
  }
  
  const idxCompetencia = garantirColuna(aba, headers, "COMPETÊNCIA");
  const idxVencimento = garantirColuna(aba, headers, "VENCIMENTO DAS");
  const idxStatusPdf = garantirColuna(aba, headers, "STATUS PDF DAS");
  
  const dataHoje = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");
  
  let grupos = {};
  
  function formatarDataSimples(valor) {
    if (valor instanceof Date) return Utilities.formatDate(valor, Session.getScriptTimeZone(), "dd/MM/yyyy");
    return String(valor); 
  }

  for (let i = 0; i < linhasSel.length; i++) {
    let row = linhasSel[i];
    let linha = allData[row - 1];
    
    let cnpj = String(linha[idxCnpj]);
    let empresa = String(linha[idxEmpresa]);
    let anexo = String(linha[idxAnexo]);
    let aliqISS = linha[idxIss];
    let aliqICMS = idxIcms !== -1 ? linha[idxIcms] : "NÃO POSSUI";
    let aliqDAS = linha[idxDas];
    let competencia = String(linha[idxCompetencia]).trim();
    let vencimentoManual = linha[idxVencimento]; 
    let statusData = linha[idxStatusPdf];       
    
    if (empresa && competencia && (statusData === "" || statusData === "(auto)")) {
      let chave = cnpj + "_" + competencia;
      if (!grupos[chave]) {
        let dataVencFormatada = formatarDataSimples(vencimentoManual);
        grupos[chave] = { empresa: empresa, cnpj: cnpj, competencia: competencia, vencimentoDAS: dataVencFormatada, itens: [], linhasIndices: [] };
      }
      
      let issFormatado = formatarNumeroParaPDF(aliqISS);
      let icmsFormatado = formatarNumeroParaPDF(aliqICMS);
      let dasFormatado = formatarNumeroParaPDF(aliqDAS);
      
      let infoTributo = "";
      if (aliqISS !== "NÃO POSSUI" && aliqISS !== "") {
        infoTributo = `ISS: ${issFormatado}% | `;
      } else if (aliqICMS !== "NÃO POSSUI" && aliqICMS !== "") {
        infoTributo = `ICMS: ${icmsFormatado}% | `;
      }

      grupos[chave].itens.push(`Anexo ${anexo} (${infoTributo}DAS: ${dasFormatado}%)`);
      grupos[chave].linhasIndices.push(row); 
    }
  }
  
  let chaves = Object.keys(grupos);
  let processados = 0;
  if (chaves.length === 0) {
    SpreadsheetApp.getUi().alert("Nenhum PDF de Conjunto pendente nas linhas marcadas.");
    return;
  }
  
  const pasta = DriveApp.getFolderById(ID_PASTA_DESTINO);
  const arquivoModelo = DriveApp.getFileById(ID_MODELO_DOC);
  
  for (let k = 0; k < chaves.length; k++) {
    let grupo = grupos[chaves[k]];
    let textoAnexos = grupo.itens.join("; "); 
    let nomeArquivo = `ALÍQUOTAS_${grupo.empresa}`; 
    
    try {
      let copiaDoc = arquivoModelo.makeCopy("TEMP_" + nomeArquivo, pasta);
      let docId = copiaDoc.getId();
      Utilities.sleep(1500); 
      
      let doc = DocumentApp.openById(docId);
      let corpo = doc.getBody();
      
      corpo.replaceText("VAR_EMPRESA", grupo.empresa);
      corpo.replaceText("VAR_CNPJ", grupo.cnpj);
      corpo.replaceText("VAR_COMPETENCIA", grupo.competencia);
      corpo.replaceText("VAR_DATA", dataHoje);
      corpo.replaceText("VAR_LISTA_ANEXOS", textoAnexos);
      corpo.replaceText("VAR_VENCIMENTO_DAS", grupo.vencimentoDAS); 
      corpo.replaceText("VAR_ANEXO", "");
      corpo.replaceText("VAR_ALIQUOTA", "");
      
      doc.saveAndClose();
      Utilities.sleep(2000); 
      
      let pdfBlob = copiaDoc.getAs(MimeType.PDF);
      pasta.createFile(pdfBlob).setName(nomeArquivo + ".pdf");
      DriveApp.getFileById(docId).setTrashed(true);
      
      grupo.linhasIndices.forEach(numLinha => { aba.getRange(numLinha, idxStatusPdf + 1).setValue(dataHoje); });
      processados++;
    } catch (e) {
      SpreadsheetApp.getUi().alert(`Erro na empresa ${grupo.empresa}: ${e.message}`);
    }
  }
  if (processados > 0) SpreadsheetApp.getUi().alert(`Sucesso! ${processados} PDFs de Documento Conjunto salvos.`);
}

// =============================================================================
// MÓDULO 6: INTEGRAÇÃO ACESSÓRIAS (NOVO)
// =============================================================================
function marcarAcessorias() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var aba = ss.getSheetByName("EMPRESAS_FILTRADAS");
  
  if (!aba) {
    ui.alert("❌ Erro: Aba 'EMPRESAS_FILTRADAS' não encontrada.");
    return;
  }
  
  var cnpjsEnviados = [];
  var abaAcessorias = ss.getSheetByName("ACESSORIAS");
  
  // 1. LER ABA ACESSORIAS SE EXISTIR
  if (abaAcessorias) {
    var dataAc = abaAcessorias.getDataRange().getValues();
    var headers = dataAc[0] || [];
    var idxCnpj = buscarIndiceColuna(headers, "CNPJ");
    
    if (idxCnpj !== -1) {
       for (var i = 1; i < dataAc.length; i++) {
         if (dataAc[i][idxCnpj]) {
            cnpjsEnviados.push(String(dataAc[i][idxCnpj]).replace(/[^\d]/g, ""));
         }
       }
    } else {
       var txt = JSON.stringify(dataAc);
       var matches = txt.match(/\d{2}\.\d{3}\.\d{3}\/\d{4}-\d{2}/g) || [];
       for (var i = 0; i < matches.length; i++) {
          cnpjsEnviados.push(matches[i].replace(/[^\d]/g, ""));
       }
    }
  } 
  
  // 2. SE NÃO ACHOU NADA NA ABA, PERGUNTA PELO GOOGLE DRIVE
  if (cnpjsEnviados.length === 0) {
    var prompt = ui.prompt(
      '📥 Importar Relatório do Acessórias', 
      'Opção 1: Cole o LINK do arquivo CSV do Acessórias que está no Google Drive.\n\nOpção 2: Crie uma aba chamada "ACESSORIAS", cole os dados do CSV nela e deixe esse campo em branco.', 
      ui.ButtonSet.OK_CANCEL
    );
    
    if (prompt.getSelectedButton() !== ui.Button.OK) return;
    var url = prompt.getResponseText().trim();
    
    if (!url) {
      ui.alert("⚠️ Nenhum link do Drive foi informado e a aba 'ACESSORIAS' está vazia ou não existe.");
      return;
    }
    
    var fileIdMatch = url.match(/[-\w]{25,}/);
    if (!fileIdMatch) {
       ui.alert("❌ Link inválido. Copie e cole um link de compartilhamento do Google Drive válido.");
       return;
    }
    
    try {
      var file = DriveApp.getFileById(fileIdMatch[0]);
      var mimeType = file.getMimeType();
      
      if (mimeType === MimeType.CSV || mimeType === MimeType.PLAIN_TEXT || mimeType.includes("csv")) {
        var txt = file.getBlob().getDataAsString("UTF-8");
        var matches = txt.match(/\d{2}\.\d{3}\.\d{3}\/\d{4}-\d{2}/g);
        if(!matches) {
           txt = file.getBlob().getDataAsString("ISO-8859-1");
           matches = txt.match(/\d{2}\.\d{3}\.\d{3}\/\d{4}-\d{2}/g);
        }
        if (matches) {
          for(var i=0; i < matches.length; i++) {
            cnpjsEnviados.push(matches[i].replace(/[^\d]/g, ""));
          }
        }
      } else {
        var tempSS = SpreadsheetApp.openById(fileIdMatch[0]);
        var sheetTemp = tempSS.getSheets()[0];
        var dataTemp = sheetTemp.getDataRange().getValues();
        var headersTemp = dataTemp[0];
        var idxCnpjTemp = buscarIndiceColuna(headersTemp, "CNPJ");
        if(idxCnpjTemp !== -1) {
           for(var i = 1; i < dataTemp.length; i++) {
              if(dataTemp[i][idxCnpjTemp]) {
                 cnpjsEnviados.push(String(dataTemp[i][idxCnpjTemp]).replace(/[^\d]/g, ""));
              }
           }
        }
      }
    } catch(e) {
      ui.alert("❌ Erro ao ler o arquivo do Drive: " + e.message + "\n\nO link pode estar restrito. Você liberou o acesso?");
      return;
    }
  }
  
  if (cnpjsEnviados.length === 0) {
     ui.alert("⚠️ Nenhum CNPJ válido encontrado no arquivo ou na aba do Acessórias.");
     return;
  }
  
  var headersDestino = aba.getRange(1, 1, 1, aba.getLastColumn()).getValues()[0];
  var idxCnpjDestino = buscarIndiceColuna(headersDestino, "CNPJ");
  
  if(idxCnpjDestino === -1) {
     ui.alert("❌ Coluna 'CNPJ' não encontrada na aba EMPRESAS_FILTRADAS."); 
     return;
  }
  
  var idxEnvio = garantirColuna(aba, headersDestino, "ENVIO");
  var dataDestino = aba.getDataRange().getValues();
  var marcados = 0;
  
  for (var i = 1; i < dataDestino.length; i++) {
     var cnpjDestino = String(dataDestino[i][idxCnpjDestino]).replace(/[^\d]/g, "");
     if (cnpjDestino && cnpjsEnviados.indexOf(cnpjDestino) !== -1) {
        aba.getRange(i + 1, idxEnvio + 1).setValue("ACESSORIAS");
        marcados++;
     }
  }
  
  ui.alert("✅ Sucesso! " + marcados + " empresas que constam no relatório foram carimbadas como 'ACESSORIAS' na coluna ENVIO.");
}