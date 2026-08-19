const CONFIG = {
  PLANILHA_ID: 'COLE_AQUI_O_ID_DA_PLANILHA',
  MODELO_ID: 'COLE_AQUI_O_ID_DO_MODELO_GOOGLE_DOCS',
  PASTA_ID: 'COLE_AQUI_O_ID_DA_PASTA_DO_DRIVE',
  ABA_REGRAS: 'Regras',
  ABA_REGISTROS: 'Registros',
  LIMITE_DOSES: 20,
  FUSO_HORARIO: 'America/Sao_Paulo'
};

function doGet() {
  return HtmlService
    .createHtmlOutputFromFile('Formulario')
    .setTitle('Emissão de Vale-Sêmen');
}

function prepararSistema() {
  const planilha = SpreadsheetApp.openById(CONFIG.PLANILHA_ID);
  let aba = planilha.getSheetByName(CONFIG.ABA_REGISTROS);

  if (!aba) {
    aba = planilha.insertSheet(CONFIG.ABA_REGISTROS);
  }

  const cabecalhos = [[
    'Data e hora', 'Número', 'Nome', 'IE', 'CPF', 'Produção anual',
    'Faixa', 'Litros por vale', 'Doses', 'Validade', 'Documento', 'PDF'
  ]];

  aba.getRange(1, 1, 1, cabecalhos[0].length).setValues(cabecalhos);
  aba.getRange(1, 1, 1, cabecalhos[0].length)
    .setFontWeight('bold')
    .setBackground('#1f4e78')
    .setFontColor('#ffffff');

  aba.setFrozenRows(1);
  aba.getRange('D:E').setNumberFormat('@');
}

function calcularVale(litrosInformados) {
  const litros = converterNumero(litrosInformados);

  if (!Number.isFinite(litros) || litros <= 0) {
    throw new Error('Informe uma produção anual maior que zero.');
  }

  const planilha = SpreadsheetApp.openById(CONFIG.PLANILHA_ID);
  const aba = planilha.getSheetByName(CONFIG.ABA_REGRAS);

  if (!aba) {
    throw new Error('A aba "Regras" não foi encontrada.');
  }

  const ultimaLinha = aba.getLastRow();

  if (ultimaLinha < 3) {
    throw new Error('A aba "Regras" não possui faixas cadastradas.');
  }

  const regras = aba.getRange(3, 1, ultimaLinha - 2, 5).getValues();
  const regra = regras.find(linha => {
    const minimo = Number(linha[1]);
    const maximo = Number(linha[2]);
    return litros >= minimo && litros <= maximo;
  });

  if (!regra) {
    throw new Error('Não foi encontrada uma regra para essa produção.');
  }

  const litrosPorVale = Number(regra[3]);
  const descricao = String(regra[4]);
  const doses = Math.min(
    CONFIG.LIMITE_DOSES,
    Math.ceil(litros / litrosPorVale)
  );

  return {
    litros: litros,
    faixa: descricao,
    litrosPorVale: litrosPorVale,
    doses: doses
  };
}

function visualizarCalculo(litros) {
  return calcularVale(litros);
}

function gerarVale(dados) {
  validarDados(dados);
  validarSenhaEmissao(dados.senha);
  delete dados.senha;

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const calculo = calcularVale(dados.litros);
    const agora = new Date();
    const ano = Number(
      Utilities.formatDate(agora, CONFIG.FUSO_HORARIO, 'yyyy')
    );

    const valeExistente = buscarValeExistente(dados.cpf, ano);

    if (valeExistente) {
      return {
        sucesso: false,
        jaRecebeu: true,
        mensagem: 'Este produtor já recebeu o Vale-Sêmen no ano de ' + ano + '.',
        numero: valeExistente.numero,
        nome: valeExistente.nome,
        data: valeExistente.data,
        doses: valeExistente.doses,
        documentoUrl: valeExistente.documentoUrl,
        pdfUrl: valeExistente.pdfUrl
      };
    }

    const numero = obterProximoNumero(ano);
    const numeroCompleto = numero + '/' + ano;
    const validade = '31/12/' + ano;
    const dataFormatada = Utilities.formatDate(
      agora,
      CONFIG.FUSO_HORARIO,
      'dd/MM/yyyy'
    );

    const pasta = DriveApp.getFolderById(CONFIG.PASTA_ID);
    const modelo = DriveApp.getFileById(CONFIG.MODELO_ID);
    const nomeArquivo =
      'Vale-Sêmen ' + numeroCompleto.replace('/', '-') +
      ' - ' + limparNomeArquivo(dados.nome);

    const copia = modelo.makeCopy(nomeArquivo, pasta);
    const documento = DocumentApp.openById(copia.getId());
    const corpo = documento.getBody();

    substituir(corpo, '<<NUMERO>>', numeroCompleto);
    substituir(corpo, '<<NOME>>', dados.nome);
    substituir(corpo, '<<CPF>>', dados.cpf);
    substituir(corpo, '<<IE>>', dados.ie);
    substituir(corpo, '<<LITROS>>', formatarNumero(calculo.litros));
    substituir(corpo, '<<FAIXA>>', calculo.faixa);
    substituir(corpo, '<<DOSES>>', String(calculo.doses));
    substituir(corpo, '<<DATA>>', dataFormatada);
    substituir(corpo, '<<VALIDADE>>', validade);

    documento.saveAndClose();

    const pdfBlob = copia.getAs(MimeType.PDF).setName(nomeArquivo + '.pdf');
    const pdf = pasta.createFile(pdfBlob);

    registrarEmissao({
      data: agora,
      numero: numeroCompleto,
      nome: dados.nome,
      ie: dados.ie,
      cpf: dados.cpf,
      litros: calculo.litros,
      faixa: calculo.faixa,
      litrosPorVale: calculo.litrosPorVale,
      doses: calculo.doses,
      validade: validade,
      documentoUrl: copia.getUrl(),
      pdfUrl: pdf.getUrl()
    });

    return {
      sucesso: true,
      numero: numeroCompleto,
      doses: calculo.doses,
      faixa: calculo.faixa,
      documentoUrl: copia.getUrl(),
      pdfUrl: pdf.getUrl()
    };
  } finally {
    lock.releaseLock();
  }
}

function validarSenhaEmissao(senhaInformada) {
  const senhaCorreta = PropertiesService
    .getScriptProperties()
    .getProperty('SENHA_EMISSAO');

  if (!senhaCorreta) {
    throw new Error('A senha de emissão ainda não foi configurada pelo administrador.');
  }

  const senha = String(senhaInformada || '');

  if (!senha) {
    throw new Error('Informe a senha para emitir o vale.');
  }

  if (senha !== senhaCorreta) {
    throw new Error('Senha de emissão incorreta.');
  }
}

function obterProximoNumero(ano) {
  const planilha = SpreadsheetApp.openById(CONFIG.PLANILHA_ID);
  const aba = planilha.getSheetByName(CONFIG.ABA_REGISTROS);
  const ultimaLinha = aba.getLastRow();

  if (ultimaLinha < 2) {
    return 1;
  }

  const numeros = aba
    .getRange(2, 2, ultimaLinha - 1, 1)
    .getDisplayValues()
    .flat();

  let maior = 0;

  numeros.forEach(valor => {
    const resultado = String(valor).match(/^(\d+)\/(\d{4})$/);

    if (resultado && Number(resultado[2]) === ano) {
      maior = Math.max(maior, Number(resultado[1]));
    }
  });

  return maior + 1;
}

function registrarEmissao(dados) {
  const planilha = SpreadsheetApp.openById(CONFIG.PLANILHA_ID);
  const aba = planilha.getSheetByName(CONFIG.ABA_REGISTROS);

  aba.appendRow([
    dados.data, dados.numero, dados.nome, dados.ie, dados.cpf,
    dados.litros, dados.faixa, dados.litrosPorVale, dados.doses,
    dados.validade, dados.documentoUrl, dados.pdfUrl
  ]);

  const linha = aba.getLastRow();
  aba.getRange(linha, 4, 1, 2).setNumberFormat('@');
}

function substituir(corpo, marcador, valor) {
  const textoSeguro = String(valor ?? '').replace(/\$/g, '\\$');
  corpo.replaceText(escaparRegex(marcador), textoSeguro);
}

function escaparRegex(texto) {
  return texto.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function validarDados(dados) {
  if (!dados) {
    throw new Error('Os dados do formulário não foram recebidos.');
  }

  dados.nome = String(dados.nome || '').trim();
  dados.ie = String(dados.ie || '').trim();
  dados.cpf = String(dados.cpf || '').trim();

  if (!dados.nome) {
    throw new Error('Informe o nome do produtor.');
  }

  if (!dados.ie) {
    throw new Error('Informe a Inscrição Estadual.');
  }

  if (!dados.cpf) {
    throw new Error('Informe o CPF.');
  }
}

function converterNumero(valor) {
  if (typeof valor === 'number') {
    return valor;
  }

  let texto = String(valor || '').trim().replace(/\s/g, '');

  if (texto.includes(',') && texto.includes('.')) {
    texto = texto.replace(/\./g, '').replace(',', '.');
  } else if (texto.includes(',')) {
    texto = texto.replace(',', '.');
  }

  return Number(texto);
}

function formatarNumero(numero) {
  return Number(numero).toLocaleString('pt-BR', {
    maximumFractionDigits: 0
  });
}

function limparNomeArquivo(nome) {
  return String(nome).replace(/[\\/:*?"<>|]/g, '').trim();
}

function testarPasta() {
  const pasta = DriveApp.getFolderById(CONFIG.PASTA_ID);
  Logger.log('Pasta encontrada: ' + pasta.getName());
}

function buscarValeExistente(cpfInformado, ano) {
  const cpf = normalizarCpf(cpfInformado);

  if (!cpf) {
    return null;
  }

  const planilha = SpreadsheetApp.openById(CONFIG.PLANILHA_ID);
  const aba = planilha.getSheetByName(CONFIG.ABA_REGISTROS);
  const ultimaLinha = aba.getLastRow();

  if (ultimaLinha < 2) {
    return null;
  }

  const registros = aba
    .getRange(2, 1, ultimaLinha - 1, 12)
    .getDisplayValues();

  for (let i = registros.length - 1; i >= 0; i--) {
    const linha = registros[i];
    const numero = linha[1];
    const cpfRegistrado = normalizarCpf(linha[4]);
    const correspondeAoAno = numero.endsWith('/' + ano);

    if (cpfRegistrado === cpf && correspondeAoAno) {
      return {
        data: linha[0],
        numero: numero,
        nome: linha[2],
        cpf: linha[4],
        doses: linha[8],
        documentoUrl: linha[10],
        pdfUrl: linha[11]
      };
    }
  }

  return null;
}

function normalizarCpf(cpf) {
  return String(cpf || '').replace(/\D/g, '');
}
