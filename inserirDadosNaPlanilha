function inserirDadosNaPlanilha() {
  try {
    // Carrega as configurações da aba "Configurações"
    const config = carregarConfiguracoes();
    if (!config) {
      Logger.log("Erro: Não foi possível carregar as configurações.");
      return;
    }

    // Carrega os dados da aba de origem
    const dados = carregarDadosDeOutraAba("Insercao", "A2:E");
    if (dados.length === 0) {
      Logger.log("Nenhum dado encontrado para inserir.");
      return;
    }

    // Valida os dados
    if (!validarDados(dados)) {
      Logger.log("Erro: Os dados não estão no formato esperado.");
      return;
    }

    // Remove duplicatas
    const dadosUnicos = removerDuplicatas(dados);

    // Acessa a aba principal
    const planilha = SpreadsheetApp.getActiveSpreadsheet();
    const abaPrincipal = planilha.getSheetByName(config.NOME_ABA_PRINCIPAL);
    if (!abaPrincipal) {
      Logger.log("A aba especificada não foi encontrada.");
      return;
    }

    // Encontra as linhas vazias no intervalo principal
    const intervaloRange = abaPrincipal.getRange(config.INTERVALO_PRINCIPAL);
    const valoresIntervalo = intervaloRange.getValues();
    const linhasVazias = [];

    for (let i = 0; i < valoresIntervalo.length; i++) {
      if (valoresIntervalo[i].every(cell => !cell)) { // Verifica se a linha está vazia
        linhasVazias.push(i);
      }
    }

    if (linhasVazias.length === 0) {
      Logger.log("Erro: O intervalo principal está completamente preenchido.");
      return;
    }

    // Limita a inserção ao número de linhas vazias disponíveis
    const dadosParaInserir = dadosUnicos.slice(0, linhasVazias.length).map((row, index) => [
      row[0], // Descrição (qualquer valor)
      row[1], // Categoria (deve estar no combo)
      row[2].toString(), // Parcela (convertido para string)
      row[3].toString(), // Preço (inserido como string)
      `=C${linhasVazias[index] + 12}*D${linhasVazias[index] + 12}` // Total (fórmula)
    ]);

    // Insere os dados nas linhas vazias
    for (let i = 0; i < dadosParaInserir.length; i++) {
      const linha = linhasVazias[i] + 12; // Ajusta o índice para o intervalo A12:E32
      abaPrincipal.getRange(linha, 1, 1, 5).setValues([dadosParaInserir[i]]);
    }

    // Aplica o combo fixo para o campo Categoria
    const categoriasPermitidas = ["Internet", "Cartão", "Assinatura", "Alimentação", "Entretenimento", "Outros"];
    atualizarMenuSuspenso(abaPrincipal, config.INTERVALO_CATEGORIA, categoriasPermitidas);

    Logger.log("Dados inseridos com sucesso!");
  } catch (error) {
    Logger.log("Erro durante a execução: " + error.toString());
  }
}

// Funções auxiliares
function carregarConfiguracoes() {
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  const abaConfig = planilha.getSheetByName("Configurações");
  if (!abaConfig) {
    Logger.log("Erro: A aba 'Configurações' não foi encontrada.");
    return null;
  }

  const dadosConfig = abaConfig.getRange("A2:B").getValues().filter(row => row[0]); // Filtra linhas vazias
  const config = {};

  dadosConfig.forEach(row => {
    config[row[0]] = row[1]; // Mapeia chave-valor
  });

  return config;
}

function carregarDadosDeOutraAba(nomeAba, intervalo) {
  const abaOrigem = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nomeAba);
  return abaOrigem.getRange(intervalo).getValues().filter(row => row.some(cell => cell));
}

function validarDados(dados) {
  return dados.every(row => row.length === 5); // Verifica se cada linha tem 5 colunas
}

function removerDuplicatas(dados) {
  const descricoesUnicas = new Set();
  return dados.filter(row => {
    if (!descricoesUnicas.has(row[0])) {
      descricoesUnicas.add(row[0]);
      return true;
    }
    return false;
  });
}

function atualizarMenuSuspenso(aba, intervalo, valoresNovos) {
  const intervaloRange = aba.getRange(intervalo);
  const novaRegra = SpreadsheetApp.newDataValidation()
    .requireValueInList(valoresNovos, true)
    .setAllowInvalid(false)
    .build();
  intervaloRange.setDataValidation(novaRegra);
}
