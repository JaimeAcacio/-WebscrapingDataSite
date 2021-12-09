function doGet(request) {
  return HtmlService.createTemplateFromFile("Index").evaluate();
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

const planilha = SpreadsheetApp.getActive();
const data = new Date();
const dia = ("0" + data.getDate()).slice(-2);
const mes = ("0" + (data.getMonth() + 1)).slice(-2);
const ano = data.getFullYear().toFixed();
const dataFormatada = `${dia}/${mes}/${ano}`;
const dataFormatadaCookie = `${mes}/${dia}/${ano}`;
const mesVigente = `${mes}/${ano}`;
const selecionarTabela = (nomeTabela) =>
  planilha.setActiveSheet(planilha.getSheetByName(nomeTabela), true);
const setarFormulaTabela = (stringFormula) =>
  planilha.getCurrentCell().setFormula(stringFormula);
const setarDadosTabela = (dados) => planilha.getDataRange().setValues(dados);

const username = USERNAMEHERE;
const password = PASSWORDHERE;

const novaHora = () => {
  const pad = (s) => (s < 10 ? "0" + s : s);
  const retorno = `${dataFormatada} ${[data.getHours(), data.getMinutes()]
    .map(pad)
    .join(":")}`;
  return retorno;
};

const novaHoraCookie = () => {
  const pad = (s) => (s < 10 ? "0" + s : s);
  const retorno = `${dataFormatadaCookie} ${[data.getHours(), data.getMinutes()]
    .map(pad)
    .join(":")}`;
  return retorno;
};

const escreveCookie = (value) => {
  const cookies = value;
  const cookieExpire = cookies[2].split("=")[1].split(";")[0];
  const cookieInteiro = cookies[1];

  selecionarTabela("Cookie");
  planilha.getRange("A1").setValue(cookieInteiro);
  planilha.getRange("A2").setValue(cookieExpire);
};

const pegaCookieTabela = () => {
  selecionarTabela("Cookie").hideSheet();
  const dados = planilha.getRange("A1:A2").getValues();
  const [cookie, expire] = dados;
  return {
    cookie,
    expire,
  };
};

const comparacaoExpireCookie = () => {
  const dados = pegaCookieTabela();
  const expireTabela = dados.expire[0];
  const horaAtual = novaHoraCookie();
  const options = {
    headers: {
      Cookie: dados.cookie[0],
    },
  };
  return expireTabela > horaAtual ? options : false;
};

const extrairDadosWkf = (url, username, password) => {
  const loginOptions = { OPTIONS_TO_BE_SENT_WITH_THE_REQUEST };

  const pegaCookie = () => {
    const loginResponse = UrlFetchApp.fetch(
      `LINK HERE.aspx?user=${username}&pass=${password}`,
      loginOptions
    );
    const headers = loginResponse.getAllHeaders();
    const cookies = headers["Set-Cookie"];
    const options = {
      headers: {
        Cookie: cookies[1],
      },
    };
    escreveCookie(cookies);
    return options;
  };

  const pegaDados = (optionsCookie = pegaCookie()) => {
    const wkfData = UrlFetchApp.fetch(url, optionsCookie);
    const charset = "ISO-8859-1";
    const data = wkfData.getBlob().getDataAsString(charset);
    const csv = Utilities.parseCsv(data, ";");

    return csv;
  };

  if (comparacaoExpireCookie()) {
    const cookie = comparacaoExpireCookie();
    return pegaDados(cookie);
  } else {
    const cookieAtualizado = pegaCookie();
    return pegaDados(cookieAtualizado);
  }
};

const getDataWkft = () => {
  const csvDadosWkf = extrairDadosWkf(
    `LINK HERE ${dataFormatada}`,
    username,
    password
  );
  const tratandoDados = csvDadosWkf.reduce((acc, item) => {
    const angel = item[21];
    const stonecode = item[5];

    acc[angel] = acc[angel] || [];
    acc[angel].push(stonecode);
    return acc;
  }, {});

  const conversaoObjArr = Object.entries(tratandoDados);

  const dadosTratados = conversaoObjArr.map(([key, value]) => {
    const angel = key;
    const totalOS = value.length;
    const scDiferentes = [...new Set(value)].length;

    return [angel, totalOS, scDiferentes];
  });

  const dadosLimpos = dadosTratados.filter(
    ([key, value]) =>
      (value !== "0") &
      (key !== "Técnico") &
      (key !== "") &
      (key !== "TECHNICAL_NAME") &
      (key !== "undefined")
  );

  return dadosLimpos;
};

const criaTabelaSite = () => {
  const ordenarPorStonecode = (a, b) =>
    a[2] > b[2] ? -1 : a[2] < b[2] ? 1 : 0;
  const dadosFn = getDataWkft();

  const dadosOrdenados = dadosFn.sort(ordenarPorStonecode);

  const moldeTabela = dadosOrdenados
    .map(
      ([nome, qtdOS, clientes]) =>
        `<tr>
          <td>${nome}</td>
          <td>${qtdOS}</td>
          <td>${clientes}</td>
        </tr>`
    )
    .toString()
    .split(",")
    .join("");

  const retornoTabela = moldeTabela
    ? moldeTabela
    : "Não Existem OS,s baixadas!";

  return retornoTabela;
};

const consolidacaoDeDados = () => {
  const dados = extrairDadosWkf(
    `LINK HERE ${dataFormatada}`,
    username,
    password
  );

  const tratandoDados = dados.reduce((acc, item) => {
    const angel = item[21];
    const stonecode = item[5];

    acc[angel] = acc[angel] || [];
    acc[angel].push(stonecode);

    return acc;
  }, {});

  const conversaoObjArrExample = Object.entries(tratandoDados);
  const dadosTratados = conversaoObjArrExample.map(([key, value]) => {
    const angel = key;
    const totalOS = value.length;
    const scDiferentes = [...new Set(value)].length;

    return [angel, totalOS, scDiferentes];
  });

  const dadosLimpos = dadosTratados.filter(
    ([key, value]) =>
      (value !== "0") &
      (key !== "Técnico") &
      (key !== "") &
      (key !== "TECHNICAL_NAME") &
      (key !== "undefined")
  );

  const bdOSs = selecionarTabela("BaseDeDados").hideSheet();

  dadosLimpos.forEach(([angel, totalOS, scdiferentes]) =>
    bdOSs.appendRow([angel, totalOS, scdiferentes, dataFormatada, mesVigente])
  );
};

const fazConsulta = (serial) => {
  const serialCheck = extrairDadosWkf(
    `LINK HERE ${serial}`,
    username,
    password
  );
  const status = serialCheck[1][2];
  return status;
};

const devolveStatus = (serial) =>
  fazConsulta(serial) ? fazConsulta(serial) : "Serial não Existe";

const estoqueConsolidado = () => {
  const dados = extrairDadosWkf(`LINK HERE`, username, password);

  const tratandoDados = dados.reduce((acc, item) => {
    const modelo = item[4];

    acc[modelo] = acc[modelo] || [];
    acc[modelo].push(modelo);

    return acc;
  }, {});

  const conversaoObjArr = Object.entries(tratandoDados);
  const dadosTratados = conversaoObjArr.map(([key, value]) => {
    const modelo = key;
    const total = value.length;

    return [modelo, total];
  });
  const dadosLimpos = dadosTratados.filter(
    ([key, value]) =>
      (value !== "") & (key !== "Modelo") & (key !== "") & (key !== "undefined")
  );
  const ordenarPorQuantidade = (a, b) =>
    a[1] > b[1] ? -1 : a[1] < b[1] ? 1 : 0;
  const dadosOrdenados = dadosLimpos.sort(ordenarPorQuantidade);

  const moldeTabela = dadosOrdenados
    .map(
      ([modelo, quantidade]) =>
        `<tr>
          <td>${modelo}</td>
          <td>${quantidade}</td>
        </tr>`
    )
    .toString()
    .split(",")
    .join("");

  return moldeTabela;
};