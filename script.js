const inputExcel = document.getElementById("inputExcel");
const tabelaBody = document.querySelector("#tabelaPatrimonio tbody");

const totalItens = document.getElementById("totalItens");
const valorTotal = document.getElementById("valorTotal");
const totalDepartamentos = document.getElementById("totalDepartamentos");
const ativos = document.getElementById("ativos");
const manutencao = document.getElementById("manutencao");
const estoque = document.getElementById("estoque");
const baixados = document.getElementById("baixados");
const alertaResponsavel = document.getElementById("alertaResponsavel");
const alertaValor = document.getElementById("alertaValor");
const alertaInventario = document.getElementById("alertaInventario");
const alertaManutencao = document.getElementById("alertaManutencao");

const pesquisa = document.getElementById("pesquisa");

const toggleDark = document.getElementById("toggleDark");
const filtroDepartamento = document.getElementById("filtroDepartamento");
const filtroFabricante = document.getElementById("filtroFabricante");
const filtroStatus = document.getElementById("filtroStatus");
const filtroResponsavel = document.getElementById("filtroResponsavel");

let dadosPlanilha = [];

let graficoStatus;
let graficoDepartamento;
let graficoRanking;
let graficoValorDepartamento;

inputExcel.addEventListener("change", carregarPlanilha);

function carregarPlanilha(event) {
  const arquivo = event.target.files[0];

  const reader = new FileReader();

  reader.onload = function (e) {
    const data = new Uint8Array(e.target.result);

    const workbook = XLSX.read(data, { type: "array" });

    const primeiraAba = workbook.SheetNames[0];

    const worksheet = workbook.Sheets[primeiraAba];

    const dados = XLSX.utils.sheet_to_json(worksheet);

    dadosPlanilha = dados;

    atualizarDashboard(dados);

    atualizarFiltroDepartamentos(dados);

    preencherFiltros(dados);
  };

  reader.readAsArrayBuffer(arquivo);
}

function atualizarDashboard(dados) {
  tabelaBody.innerHTML = "";

  let somaValores = 0;
  let totalAtivos = 0;
  let totalManutencao = 0;
  let totalEstoque = 0;
  let totalBaixados = 0;
  let semResponsavel = 0;
  let semValor = 0;
  let semInventario = 0;
  let emManutencao = 0;

  const departamentos = new Set();

  const statusContagem = {};
  const departamentoContagem = {};
  const valorDepartamento = {};

  dados.forEach((item) => {
    const tr = document.createElement("tr");

    tr.innerHTML = `
      <td>${item["Identificacao"] || item["Identificação"] || ""}</td>
      <td>${item["Produto"] || ""}</td>
      <td>${item["Fabricante"] || ""}</td>
      <td>${item["Modelo"] || ""}</td>
      <td>${item["Status"] || ""}</td>
      <td>${item["Departamento"] || ""}</td>
      <td>${item["Responsavel"] || item["Responsável"] || ""}</td>
      <td>R$ ${Number(item["Valor Total"] || 0).toLocaleString("pt-BR")}</td>
    `;

    tabelaBody.appendChild(tr);

    somaValores += Number(item["Valor Total"]) || 0;

    if (item["Departamento"]) {
      departamentos.add(item["Departamento"]);
    }

    const status = item["Status"] || "Sem Status";
    const statusTexto = status.toLowerCase();

    if (statusTexto.includes("ativo")) {
      totalAtivos++;
    } else if (
      statusTexto.includes("manutenção") ||
      statusTexto.includes("manutencao")
    ) {
      totalManutencao++;
    } else if (statusTexto.includes("estoque")) {
      totalEstoque++;
    } else if (statusTexto.includes("baixado")) {
      totalBaixados++;
    }

    if (!item["Responsavel"] && !item["Responsável"]) {
      semResponsavel++;
    }

    if (!item["Valor Total"] || Number(item["Valor Total"]) === 0) {
      semValor++;
    }

    if (!item["Data Inventário"]) {
      semInventario++;
    }

    if (
      statusTexto.includes("manutenção") ||
      statusTexto.includes("manutencao")
    ) {
      emManutencao++;
    }

    statusContagem[status] = (statusContagem[status] || 0) + 1;

    const departamento = item["Departamento"] || "Sem Departamento";
    const valorItem = Number(item["Valor Total"]) || 0;

    valorDepartamento[departamento] =
      (valorDepartamento[departamento] || 0) + valorItem;

    departamentoContagem[departamento] =
      (departamentoContagem[departamento] || 0) + 1;
  });

  totalItens.textContent = dados.length;

  valorTotal.textContent = somaValores.toLocaleString("pt-BR", {
    style: "currency",
    currency: "BRL",
  });

  totalDepartamentos.textContent = departamentos.size;
  ativos.textContent = totalAtivos;
  manutencao.textContent = totalManutencao;
  estoque.textContent = totalEstoque;
  baixados.textContent = totalBaixados;
  alertaResponsavel.textContent = semResponsavel;
  alertaValor.textContent = semValor;
  alertaInventario.textContent = semInventario;
  alertaManutencao.textContent = emManutencao;

  criarGraficoStatus(statusContagem);

  criarGraficoDepartamento(departamentoContagem);
  criarGraficoRanking(departamentoContagem);
  criarGraficoValorDepartamento(valorDepartamento);
}

function criarGraficoStatus(dados) {
  const ctx = document.getElementById("graficoStatus");

  if (graficoStatus) {
    graficoStatus.destroy();
  }

  graficoStatus = new Chart(ctx, {
    type: "pie",

    data: {
      labels: Object.keys(dados),

      datasets: [
        {
          data: Object.values(dados),

          backgroundColor: [
            "#facc15",
            "#eab308",
            "#ca8a04",
            "#fde047",
            "#f59e0b"
          ],

          borderColor: "#111",
          borderWidth: 2,
        },
      ],
    },

    options: {
      responsive: true,

      plugins: {
        legend: {
          labels: {
            color: "white",
          },
        },
      },
    },
  });
}

function criarGraficoDepartamento(dados) {
  const ctx = document.getElementById("graficoDepartamento");

  if (graficoDepartamento) {
    graficoDepartamento.destroy();
  }

  graficoDepartamento = new Chart(ctx, {
    type: "bar",

    data: {
      labels: Object.keys(dados),

      datasets: [
        {
          label: "Quantidade",
          data: Object.values(dados),

          backgroundColor: "#facc15",
          borderColor: "#eab308",
          borderWidth: 2,
          borderRadius: 8,
        },
      ],
    },

    options: {
      responsive: true,

      plugins: {
        legend: {
          labels: {
            color: "white",
          },
        },
      },

      scales: {
        x: {
          ticks: {
            color: "white",
          },
        },

        y: {
          ticks: {
            color: "white",
          },
        },
      },
    },
  });
}

pesquisa.addEventListener("input", () => {
  const texto = pesquisa.value.toLowerCase();

  const filtrados = dadosPlanilha.filter((item) => {
    return Object.values(item).some((valor) =>
      String(valor).toLowerCase().includes(texto),
    );
  });

  atualizarDashboard(filtrados);
});

toggleDark.addEventListener("click", () => {
  document.body.classList.toggle("dark-mode");
});

function atualizarFiltroDepartamentos(dados) {
  filtroDepartamento.innerHTML =
    '<option value="">Todos os Departamentos</option>';

  const departamentos = [
    ...new Set(dados.map((item) => item["Departamento"]).filter(Boolean)),
  ];

  departamentos.forEach((dep) => {
    const option = document.createElement("option");

    option.value = dep;
    option.textContent = dep;

    filtroDepartamento.appendChild(option);
  });
}

filtroDepartamento.addEventListener("change", aplicarFiltros);

function criarGraficoRanking(dados) {
  const ctx = document.getElementById("graficoRanking");

  if (graficoRanking) {
    graficoRanking.destroy();
  }

  const ordenado = Object.entries(dados).sort((a, b) => b[1] - a[1]);

  const labels = ordenado.map((item) => item[0]);

  const valores = ordenado.map((item) => item[1]);

  graficoRanking = new Chart(ctx, {
    type: "bar",

    data: {
      labels: labels,

      datasets: [
        {
          label: "Ranking de Patrimônios",
          data: valores,

          backgroundColor: "#facc15",
          borderColor: "#eab308",
          borderWidth: 2,
          borderRadius: 8,
        },
      ],
    },

    options: {
      responsive: true,
      indexAxis: "y",

      plugins: {
        legend: {
          labels: {
            color: "white",
          },
        },

        title: {
          display: true,
          text: "Ranking de Patrimônios",
          color: "white",
        },
      },

      scales: {
        x: {
          ticks: {
            color: "white",
          },
        },

        y: {
          ticks: {
            color: "white",
          },
        },
      },
    },
  });
}

function criarGraficoValorDepartamento(dados) {
  const ctx = document.getElementById("graficoValorDepartamento");

  if (graficoValorDepartamento) {
    graficoValorDepartamento.destroy();
  }

  const ordenado = Object.entries(dados).sort((a, b) => b[1] - a[1]);

  const labels = ordenado.map((item) => item[0]);

  const valores = ordenado.map((item) => item[1]);

  graficoValorDepartamento = new Chart(ctx, {
    type: "bar",

    data: {
      labels: labels,

      datasets: [
        {
          label: "Valor Patrimonial",
          data: valores,

          backgroundColor: "#facc15",
          borderColor: "#eab308",
          borderWidth: 2,
          borderRadius: 8,
        },
      ],
    },

    options: {
      responsive: true,

      plugins: {
        legend: {
          labels: {
            color: "white",
          },
        },

        tooltip: {
          callbacks: {
            label: function (context) {
              return context.raw.toLocaleString("pt-BR", {
                style: "currency",
                currency: "BRL",
              });
            },
          },
        },
      },

      scales: {
        x: {
          ticks: {
            color: "white",
          },
        },

        y: {
          ticks: {
            color: "white",
          },
        },
      },
    },
  });
}

function preencherFiltros(dados) {
  preencherSelect(
    filtroFabricante,
    dados.map((item) => item["Fabricante"]),
  );

  preencherSelect(
    filtroStatus,
    dados.map((item) => item["Status"]),
  );

  preencherSelect(
    filtroResponsavel,
    dados.map((item) => item["Responsavel"] || item["Responsável"]),
  );
}

function preencherSelect(select, valores) {
  const textoPadrao = select.options[0].text;

  select.innerHTML = `<option value="">${textoPadrao}</option>`;

  [...new Set(valores.filter(Boolean))].forEach((valor) => {
    const option = document.createElement("option");

    option.value = valor;
    option.textContent = valor;

    select.appendChild(option);
  });
}

function aplicarFiltros() {
  const departamento = filtroDepartamento.value;
  const fabricante = filtroFabricante.value;
  const status = filtroStatus.value;
  const responsavel = filtroResponsavel.value;

  const filtrados = dadosPlanilha.filter((item) => {
    return (
      (!departamento || item["Departamento"] === departamento) &&
      (!fabricante || item["Fabricante"] === fabricante) &&
      (!status || item["Status"] === status) &&
      (!responsavel ||
        (item["Responsavel"] || item["Responsável"]) === responsavel)
    );
  });

  atualizarDashboard(filtrados);
}

filtroFabricante.addEventListener("change", aplicarFiltros);
filtroStatus.addEventListener("change", aplicarFiltros);
filtroResponsavel.addEventListener("change", aplicarFiltros);
