const tabela = document.getElementById("tabelaOSs");
const tbody = document.createElement("tbody");
const loadInput = document.getElementsByClassName("c-loaderinput");
const loadTabela = document.getElementsByClassName("c-loadertable");
const loadTabelaEstoque = document.getElementsByClassName(
  "c-loadertableEstoque"
);
const botaoPesquisar = document.getElementById("botaoPesquisar");
const serial = document.getElementById("input_text");
const status = document.getElementById("output");
const botaoOSs = document.getElementById("botaoOSs");
const botaoEstoqueConsolidado = document.getElementById("estoqueConsolidado");
const tabelaEstoqueConoslidado = document.getElementById("tabelaEstoque");

const mostrarElemento = (elemento) => (elemento.style.display = "block");
const ocultarElemento = (elemento) => (elemento.style.display = "none");
const ocultarLoadTable = (elemento) => (elemento[0].style.display = "none");
const mostrarLoadTable = (elemento) => (elemento[0].style.display = "block");
const limparInput = () => (serial.value = "");

const atualizaTabela = (valor) => {
  tbody.innerHTML += valor;
  tabela.append(tbody);
  ocultarLoadTable(loadTabela);
  mostrarElemento(tabela);
};

const init = () => {
  mostrarLoadTable(loadTabela);
  ocultarElemento(botaoOSs);
  google.script.run.withSuccessHandler(atualizaTabela).criaTabelaSite();
};

const onSuccess = (serial) => {
  if (serial.value == "") {
    return alert(
      "Preencha o campo abaixo para pesquisar o status de um serial!"
    );
  } else {
    status.textContent = "";
    mostrarLoadTable(loadInput);

    const escreveStatus = (valor) => {
      if ((serial.value = "")) {
        return alert("digite um serial");
      }
      limparInput();
      ocultarLoadTable(loadInput);
      status.innerHTML = valor;
    };

    google.script.run
      .withSuccessHandler(escreveStatus)
      .devolveStatus(serial.value);
  }
};

const buscaEstoqueConsolidado = () => {
  ocultarElemento(botaoEstoqueConsolidado);
  mostrarLoadTable(loadTabelaEstoque);
  const escreveEstoque = (valor) => {
    tbody.innerHTML = valor;
    tabelaEstoqueConoslidado.append(tbody);
    ocultarLoadTable(loadTabelaEstoque);
    mostrarElemento(tabelaEstoqueConoslidado);
  };

  google.script.run.withSuccessHandler(escreveEstoque).estoqueConsolidado();
};

botaoOSs.addEventListener("click", init);

botaoPesquisar.addEventListener("click", function () {
  onSuccess(serial);
});

botaoEstoqueConsolidado.addEventListener("click", function () {
  buscaEstoqueConsolidado();
});
