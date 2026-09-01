import json

import webview  # mesma lib do gui.py - aqui só abre uma segunda janela sobre a mesma instância

import contratos_db
import janela_windows

class ApiContrato:
    # ponte Python <-> JS desta janela (window.pywebview.api.<metodo>), independente da Api
    # principal do gui.py - essa janela é autocontida (lista + formulário + banco local)
    def listar(self):
        return contratos_db.listar_contratos()

    def obter(self, contrato_id):
        return contratos_db.obter_contrato(contrato_id)

    def salvar(self, dados):
        try:
            if dados.get("id"):
                contratos_db.atualizar_contrato(dados["id"], dados)
                contrato_id = dados["id"]
            else:
                contrato_id = contratos_db.criar_contrato(dados)
            return {"ok": True, "id": contrato_id}
        except Exception as e:
            return {"ok": False, "erro": str(e)}

    def excluir(self, contrato_id):
        try:
            contratos_db.excluir_contrato(contrato_id)
            return {"ok": True}
        except Exception as e:
            return {"ok": False, "erro": str(e)}

HTML_CONTRATO = r"""
<!doctype html>
<html lang="pt-br">
<head>
<meta charset="utf-8">
<title>CCRGCI - Cadastro de Contratos</title>
<style>
  :root {
    --mist: #f2f5f3;
    --cloud: #ffffff;
    --hairline: #dde4e0;
    --ink: #16201b;
    --ink-soft: #56625b;
    --ink-faint: #8a958e;
    --pine: #178c4e;
    --pine-deep: #0f6b3b;
    --pine-tint: #e2f5ea;
    --pine-tint-strong: #c3ecd6;
    --status-error: #d1453d;
    --status-error-tint: #fbe9e8;
    --shadow-1: 0 1px 2px rgba(20,32,27,0.07), 0 1px 1px rgba(20,32,27,0.05);
  }

  * { box-sizing: border-box; }

  html, body {
    margin: 0; height: 100%; background: var(--mist); color: var(--ink);
    font-family: "Segoe UI Variable Display", "Segoe UI Variable Text", "Segoe UI", system-ui, sans-serif;
    font-size: 13.5px;
  }

  .pagina { max-width: 900px; margin: 0 auto; padding: 22px 26px 40px; }

  .cabecalho { display: flex; align-items: center; justify-content: space-between; gap: 10px; margin-bottom: 16px; }
  .cabecalho__marca { display: flex; align-items: center; gap: 10px; min-width: 0; }
  h1 { margin: 0; font-size: 19px; font-weight: 600; letter-spacing: -0.01em; }

  .marca-icone {
    width: 34px; height: 34px; border-radius: 8px;
    background: var(--pine);
    color: #fff;
    display: flex; align-items: center; justify-content: center;
    box-shadow: var(--shadow-1); flex: 0 0 auto;
  }
  .marca-icone svg { width: 19px; height: 19px; }

  .btn {
    font-family: inherit; font-size: 12.5px; font-weight: 600;
    border: 1px solid transparent; border-radius: 6px; padding: 7px 16px; cursor: pointer;
    white-space: nowrap; transition: background 120ms ease, border-color 120ms ease;
  }
  .btn--acento { background: var(--pine-tint); border-color: var(--pine-tint-strong); color: var(--pine-deep); }
  .btn--acento:hover { background: var(--pine-tint-strong); }
  .btn--outline { background: var(--cloud); border-color: var(--hairline); color: var(--ink); }
  .btn--outline:hover { border-color: var(--pine); color: var(--pine-deep); }
  .btn--perigo { background: var(--status-error-tint); border-color: #f3c9c6; color: var(--status-error); }
  .btn--perigo:hover { background: #f6d7d5; }
  .btn--mini { padding: 4px 10px; font-size: 11.5px; }

  table { width: 100%; border-collapse: collapse; background: var(--cloud); border: 1px solid var(--hairline); border-radius: 10px; overflow: hidden; box-shadow: var(--shadow-1); }
  th, td { text-align: left; padding: 9px 12px; font-size: 12.5px; border-bottom: 1px solid var(--hairline); }
  th { color: var(--ink); font-weight: 700; font-size: 11.5px; text-transform: uppercase; letter-spacing: 0.02em; }
  tr:last-child td { border-bottom: none; }
  tbody tr { transition: background 100ms ease; }
  tbody tr:hover { background: var(--pine-tint); }
  td.acoes { display: flex; gap: 6px; justify-content: flex-end; }
  .vazio { padding: 26px; text-align: center; color: var(--ink-faint); }

  .painel {
    border: 1px solid var(--hairline); border-radius: 10px; background: var(--pine-tint);
    box-shadow: var(--shadow-1); padding: 14px 16px; margin-bottom: 14px;
  }
  .painel h2 { margin: 0 0 10px; font-size: 13px; font-weight: 600; color: var(--pine-deep); }
  .painel--acento { border-color: var(--pine-tint-strong); background: var(--pine-tint); }

  .rotulo {
    font-weight: 600; font-size: 12px; margin: 0 0 8px;
    display: flex; align-items: center; gap: 6px;
    color: var(--pine);
  }
  .rotulo .ponto { width: 6px; height: 6px; border-radius: 50%; background: var(--pine); flex: 0 0 auto; }

  .grade { display: grid; grid-template-columns: repeat(3, 1fr); gap: 10px 12px; }
  .grade--2 { grid-template-columns: repeat(2, 1fr); }
  .campo { display: flex; flex-direction: column; gap: 4px; min-width: 0; }
  .campo--largo { grid-column: 1 / -1; }
  label { font-size: 11.5px; font-weight: 600; color: var(--ink-soft); }

  /* linhas de campos com largura sob medida (fora do grid de 3 colunas iguais) - usadas em
     "Dados do Contrato", onde alguns campos têm valores curtos (tipo, situação, planilha,
     datas) e outros precisam de mais espaço (nome da contratada, processos) */
  .linha-campos { display: flex; flex-wrap: wrap; gap: 10px 12px; margin-bottom: 10px; }
  .linha-campos .campo { min-width: 0; }
  .campo--estreito { flex: 0 0 140px; }
  .campo--flexivel { flex: 1 1 160px; }
  .campo--data { flex: 0 0 190px; }
  .campo--processo { flex: 0 0 320px; }

  input, select, textarea {
    font-family: inherit; font-size: 13px; padding: 7px 9px;
    border: 1px solid var(--hairline); border-radius: 6px; background: var(--cloud); color: var(--ink);
  }
  input:focus-visible, select:focus-visible, textarea:focus-visible { outline: none; border-color: var(--pine); box-shadow: 0 0 0 3px var(--pine-tint-strong); }
  textarea { resize: vertical; min-height: 56px; font-family: inherit; }

  .linha-dinamica { display: flex; gap: 8px; align-items: flex-end; margin-bottom: 8px; }
  .linha-dinamica .campo { flex: 1; }

  .bloco-empenho { border: 1px solid var(--hairline); border-radius: 8px; padding: 10px 12px; margin-bottom: 10px; background: var(--mist); }
  .bloco-empenho__topo { display: flex; gap: 8px; align-items: flex-end; margin-bottom: 8px; }
  .bloco-empenho__topo .campo { flex: 1; }
  .movimentacoes { padding-left: 10px; border-left: 2px solid var(--pine-tint-strong); }
  .saldo { font-size: 11.5px; color: var(--ink-soft); margin: 4px 0 8px 10px; }
  .saldo strong { color: var(--pine-deep); }

  .toggle-linha { display: flex; align-items: center; gap: 8px; margin-bottom: 8px; }
  .toggle-linha label { font-size: 12.5px; color: var(--ink); }
  .sub-grupo { display: none; }
  .sub-grupo.aberto { display: grid; }

  /* rodapé fixo no fim da viewport - o formulário é longo e a janela pode ficar maior que a área
     útil da tela; sem isso o botão Salvar some atrás da barra de tarefas */
  .rodape-form {
    display: flex; gap: 8px; justify-content: flex-end;
    position: sticky; bottom: 0; z-index: 5;
    margin-top: 12px; padding: 10px 0 8px;
    background: var(--mist); border-top: 1px solid var(--hairline);
  }
  .oculto { display: none !important; }
  .erro-form { color: var(--status-error); font-size: 12px; margin: 0 0 10px; }

  .overlay-modal {
    position: fixed; inset: 0; background: rgba(22,32,27,0.35);
    display: flex; align-items: center; justify-content: center; z-index: 100;
  }
  .modal-confirmacao {
    background: var(--cloud); border: 1px solid var(--hairline); border-radius: 10px;
    box-shadow: 0 12px 32px rgba(20,32,27,0.22), 0 4px 10px rgba(20,32,27,0.12);
    padding: 18px 20px; max-width: 340px; width: 90%;
  }
  .modal-mensagem { margin: 0 0 16px; font-size: 13px; color: var(--ink); line-height: 1.5; }
  .modal-acoes { display: flex; gap: 8px; justify-content: flex-end; }

  .link-visualizar { color: var(--pine-deep); font-weight: 600; text-decoration: none; cursor: pointer; }
  .link-visualizar:hover { text-decoration: underline; }

  .modal-visualizacao {
    background: var(--mist); border: 1px solid var(--hairline); border-radius: 10px;
    box-shadow: 0 12px 32px rgba(20,32,27,0.22), 0 4px 10px rgba(20,32,27,0.12);
    padding: 18px 20px; max-width: 640px; width: 92%; max-height: 86vh; overflow-y: auto;
  }
  .modal-visualizacao__cabecalho { display: flex; align-items: center; justify-content: space-between; gap: 10px; margin-bottom: 14px; }
  .modal-visualizacao__cabecalho h2 { margin: 0; font-size: 16px; color: var(--pine-deep); }
  .modal-visualizacao label { color: var(--pine-deep); }
  .valor-visualizacao { margin: 0; padding: 6px 0; font-size: 13px; color: var(--ink); }
  .rotulo-tributo { color: var(--pine-deep); font-weight: 600; }
  .vigencia-vencida { color: var(--status-error); font-weight: 600; }
</style>
</head>
<body>

<div class="pagina">
  <div class="cabecalho">
    <div class="cabecalho__marca">
      <div class="marca-icone"><svg viewBox="0 0 24 24" width="19" height="19" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8Z"/><path d="M14 2v6h6"/><path d="M12 12v6"/><path d="M9 15h6"/></svg></div>
      <h1>Cadastro de Contratos</h1>
    </div>
    <button class="btn btn--acento" id="botao-novo">+ Novo contrato</button>
  </div>

  <div id="vista-lista">
    <div class="painel painel--acento">
      <p class="rotulo"><span class="ponto"></span>Filtros</p>
      <div class="linha-campos" style="margin-bottom: 0;">
        <div class="campo campo--flexivel">
          <label>Buscar por contratada</label>
          <input type="text" id="filtro-nome" placeholder="Nome da contratada...">
        </div>
        <div class="campo campo--estreito">
          <label>Tipo de contrato</label>
          <select id="filtro-tipo">
            <option value="">Todos</option>
            <option value="servico">Serviço</option>
            <option value="almoxarifado">Almoxarifado</option>
          </select>
        </div>
        <div class="campo campo--estreito">
          <label>Situação</label>
          <select id="filtro-situacao">
            <option value="">Todas</option>
            <option value="vigente">Vigente</option>
            <option value="encerrado">Encerrado</option>
          </select>
        </div>
      </div>
    </div>
    <div id="lista-contratos"></div>
  </div>

  <div id="vista-form" class="oculto">
    <p class="erro-form oculto" id="erro-form"></p>
    <form id="form-contrato">
      <input type="hidden" id="campo-id">

      <div class="painel">
        <h2>Dados do Contrato</h2>

        <div class="linha-campos">
          <div class="campo campo--estreito">
            <label>Tipo de contrato</label>
            <select id="campo-tipo_contrato" required>
              <option value="servico">Serviço</option>
              <option value="almoxarifado">Almoxarifado</option>
            </select>
          </div>
          <div class="campo campo--estreito">
            <label>Situação</label>
            <select id="campo-situacao" required>
              <option value="vigente">Vigente</option>
              <option value="encerrado">Encerrado</option>
            </select>
          </div>
          <div class="campo campo--flexivel">
            <label>Nome da contratada</label>
            <input id="campo-nome_contratada" required>
          </div>
        </div>

        <div class="toggle-linha" id="grupo-mao-de-obra">
          <input type="checkbox" id="campo-tem_mao_de_obra">
          <label for="campo-tem_mao_de_obra">Contrato com mão de obra</label>
        </div>

        <div class="linha-campos">
          <div class="campo campo--estreito">
            <label>Planilha de controle</label>
            <input id="campo-nome_planilha_controle">
          </div>
          <div class="campo campo--flexivel">
            <label>CNPJ</label>
            <input id="campo-cnpj" placeholder="00.000.000/0000-00" maxlength="18">
          </div>
          <div class="campo campo--flexivel">
            <label>Nº do pregão</label>
            <input id="campo-numero_pregao">
          </div>
          <div class="campo campo--flexivel">
            <label>Nº do contrato</label>
            <input id="campo-numero_contrato">
          </div>
        </div>

        <div class="linha-campos">
          <div class="campo campo--data">
            <label>Vigência - início</label>
            <input type="date" id="campo-vigencia_inicio">
          </div>
          <div class="campo campo--data">
            <label>Vigência - fim</label>
            <input type="date" id="campo-vigencia_fim">
          </div>
        </div>

        <div class="linha-campos">
          <div class="campo campo--processo">
            <label>Processo de contratação</label>
            <input id="campo-processo_contratacao">
          </div>
          <div class="campo campo--processo" id="grupo-processo-empenho-anual">
            <label>Processo de empenho anual</label>
            <input id="campo-processo_empenho_anual">
          </div>
        </div>

        <div class="linha-campos">
          <div class="campo campo--flexivel">
            <label>Objeto Resumido</label>
            <textarea id="campo-objeto_resumido"></textarea>
          </div>
        </div>

        <div class="linha-campos">
          <div class="campo campo--flexivel">
            <label>Objeto Detalhado</label>
            <textarea id="campo-objeto_detalhado"></textarea>
          </div>
        </div>
      </div>

      <div class="painel">
        <h2>Domicílio Bancário</h2>
        <div class="grade">
          <div class="campo"><label>Banco</label><input id="campo-banco"></div>
          <div class="campo"><label>Agência</label><input id="campo-agencia"></div>
          <div class="campo"><label>Conta</label><input id="campo-conta"></div>
        </div>
      </div>

      <div class="painel" id="painel-valores-mensais">
        <h2>Valores Mensais</h2>
        <div id="lista-valores-mensais"></div>
        <button type="button" class="btn btn--outline btn--mini" id="botao-add-valor">+ Adicionar valor</button>
      </div>

      <div class="painel" id="painel-empenhos-servico">
        <h2>Empenhos</h2>
        <div id="lista-empenhos"></div>
        <button type="button" class="btn btn--outline btn--mini" id="botao-add-empenho">+ Adicionar empenho</button>
      </div>

      <div class="painel" id="painel-processos-empenho">
        <h2>Processos de Empenho</h2>
        <div id="lista-processos-empenho"></div>
        <button type="button" class="btn btn--outline btn--mini" id="botao-add-processo-empenho">+ Adicionar processo de empenho</button>
      </div>

      <div class="painel">
        <h2>Tributação</h2>

        <div class="toggle-linha">
          <input type="checkbox" id="campo-iss_incide">
          <label for="campo-iss_incide">Incide ISS</label>
        </div>
        <div class="grade grade--2 sub-grupo" id="grupo-iss">
          <div class="campo"><label>Alíquota ISS (%)</label><input type="number" step="0.01" id="campo-iss_aliquota"></div>
        </div>

        <div class="toggle-linha">
          <input type="checkbox" id="campo-previdenciaria_incide">
          <label for="campo-previdenciaria_incide">Incide Contribuição Previdenciária</label>
        </div>
        <div class="grade grade--2 sub-grupo" id="grupo-previdenciaria">
          <div class="campo"><label>Alíquota (%)</label><input type="number" step="0.01" id="campo-previdenciaria_aliquota"></div>
        </div>

        <div class="toggle-linha">
          <input type="checkbox" id="campo-federais_incide">
          <label for="campo-federais_incide">Incidem Tributos Federais</label>
        </div>
        <div class="grade grade--2 sub-grupo" id="grupo-federais">
          <div class="campo"><label>Código DARF</label><input id="campo-federais_codigo_darf"></div>
          <div class="campo"><label>Alíquota total (%)</label><input type="number" step="0.01" id="campo-federais_aliquota_total"></div>
        </div>
      </div>

      <div class="painel">
        <h2>Observação</h2>
        <div class="campo">
          <textarea id="campo-observacao"></textarea>
        </div>
      </div>

      <div class="rodape-form">
        <button type="button" class="btn btn--outline" id="botao-cancelar">Cancelar</button>
        <button type="submit" class="btn btn--acento">Salvar</button>
      </div>
    </form>
  </div>
</div>

<script>
  function formatarCnpj(digitos) {
    digitos = digitos.replace(/\D/g, "").slice(0, 14);
    let resultado = digitos;
    if (digitos.length > 12) resultado = digitos.replace(/^(\d{2})(\d{3})(\d{3})(\d{4})(\d{0,2})/, "$1.$2.$3/$4-$5");
    else if (digitos.length > 8) resultado = digitos.replace(/^(\d{2})(\d{3})(\d{3})(\d{0,4})/, "$1.$2.$3/$4");
    else if (digitos.length > 5) resultado = digitos.replace(/^(\d{2})(\d{3})(\d{0,3})/, "$1.$2.$3");
    else if (digitos.length > 2) resultado = digitos.replace(/^(\d{2})(\d{0,3})/, "$1.$2");
    return resultado;
  }

  document.getElementById("campo-cnpj").addEventListener("input", (evento) => {
    evento.target.value = formatarCnpj(evento.target.value);
  });

  function formatarMoeda(valor) {
    const numero = Number(valor) || 0;
    return numero.toLocaleString("pt-BR", { style: "currency", currency: "BRL" });
  }

  function formatarData(iso) {
    if (!iso) return "-";
    const [ano, mes, dia] = iso.split("-");
    return `${dia}/${mes}/${ano}`;
  }

  function hojeIso() {
    const agora = new Date();
    const mes = String(agora.getMonth() + 1).padStart(2, "0");
    const dia = String(agora.getDate()).padStart(2, "0");
    return `${agora.getFullYear()}-${mes}-${dia}`;
  }

  function contratoVencido(contrato) {
    // comparação de strings ISO (AAAA-MM-DD) equivale à comparação de datas
    return !!contrato.vigencia_fim && contrato.vigencia_fim < hojeIso();
  }

  function deveDestacarVencido(contrato) {
    // contrato marcado como encerrado não precisa do destaque vermelho - a data passada já
    // é esperada nesse caso, não é um alerta de algo que ficou pra trás sem ação
    return contrato.situacao !== "encerrado" && contratoVencido(contrato);
  }

  // ---------- vista de lista ----------

  let contratosCarregados = [];

  async function carregarLista() {
    contratosCarregados = await window.pywebview.api.listar();
    aplicarFiltros();
  }

  function aplicarFiltros() {
    const nomeFiltro = document.getElementById("filtro-nome").value.trim().toLocaleLowerCase("pt-BR");
    const tipoFiltro = document.getElementById("filtro-tipo").value;
    const situacaoFiltro = document.getElementById("filtro-situacao").value;

    const filtrados = contratosCarregados.filter((contrato) => {
      const bateNome = !nomeFiltro || (contrato.nome_contratada || "").toLocaleLowerCase("pt-BR").includes(nomeFiltro);
      const bateTipo = !tipoFiltro || contrato.tipo_contrato === tipoFiltro;
      const bateSituacao = !situacaoFiltro || contrato.situacao === situacaoFiltro;
      return bateNome && bateTipo && bateSituacao;
    });

    renderizarTabela(filtrados);
  }

  document.getElementById("filtro-nome").addEventListener("input", aplicarFiltros);
  document.getElementById("filtro-tipo").addEventListener("change", aplicarFiltros);
  document.getElementById("filtro-situacao").addEventListener("change", aplicarFiltros);

  function renderizarTabela(contratos) {
    const alvo = document.getElementById("lista-contratos");

    if (!contratosCarregados.length) {
      alvo.innerHTML = '<p class="vazio">Nenhum contrato cadastrado ainda.</p>';
      return;
    }
    if (!contratos.length) {
      alvo.innerHTML = '<p class="vazio">Nenhum contrato encontrado para o filtro atual.</p>';
      return;
    }

    const linhas = contratos.map((contrato) => `
      <tr>
        <td><a href="javascript:void(0)" class="link-visualizar" data-visualizar="${contrato.id}">${contrato.nome_planilha_controle || "-"}</a></td>
        <td>${contrato.numero_contrato || "-"}</td>
        <td>${contrato.tipo_contrato === "servico" ? "Serviço" : "Almoxarifado"}</td>
        <td>${contrato.situacao === "encerrado" ? "Encerrado" : "Vigente"}</td>
        <td>${formatarData(contrato.vigencia_inicio)} - <span class="${deveDestacarVencido(contrato) ? "vigencia-vencida" : ""}">${formatarData(contrato.vigencia_fim)}</span></td>
        <td class="acoes">
          <button class="btn btn--acento btn--mini" data-editar="${contrato.id}">Editar</button>
          <button class="btn btn--perigo btn--mini" data-excluir="${contrato.id}">Excluir</button>
        </td>
      </tr>
    `).join("");

    alvo.innerHTML = `
      <table>
        <thead><tr><th>Contratada</th><th>Nº Contrato</th><th>Tipo</th><th>Situação</th><th>Vigência</th><th></th></tr></thead>
        <tbody>${linhas}</tbody>
      </table>
    `;

    alvo.querySelectorAll("[data-editar]").forEach((botao) => {
      botao.addEventListener("click", () => abrirFormEditar(Number(botao.dataset.editar)));
    });
    alvo.querySelectorAll("[data-excluir]").forEach((botao) => {
      botao.addEventListener("click", () => excluirContrato(Number(botao.dataset.excluir)));
    });
    alvo.querySelectorAll("[data-visualizar]").forEach((link) => {
      link.addEventListener("click", () => abrirVisualizacao(Number(link.dataset.visualizar)));
    });
  }

  // ---------- visualização somente-leitura ----------

  function linhaVisualizacao(rotulo, valor, classeExtra) {
    return `<div class="campo"><label>${rotulo}</label><p class="valor-visualizacao ${classeExtra || ""}">${valor || "-"}</p></div>`;
  }

  function renderizarTributo(nome, incide, aliquota, extra) {
    const rotulo = `<span class="rotulo-tributo">${nome}:</span>`;
    if (!incide) return `<p class="valor-visualizacao">${rotulo} Não incide</p>`;
    return `<p class="valor-visualizacao">${rotulo} Sim${aliquota != null && aliquota !== "" ? ` — alíquota ${aliquota}%` : ""}${extra || ""}</p>`;
  }

  async function abrirVisualizacao(id) {
    const contrato = await window.pywebview.api.obter(id);
    if (!contrato) return;

    const ehServico = contrato.tipo_contrato === "servico";

    // almoxarifado não tem valor mensal (só serviço tem esse conceito)
    const valoresMensais = contrato.valores_mensais.length
      ? contrato.valores_mensais.map((item) => `<p class="valor-visualizacao">${item.descricao_servico || "-"}: ${formatarMoeda(item.valor)}</p>`).join("")
      : '<p class="valor-visualizacao">Nenhum valor cadastrado.</p>';

    const secaoEmpenhos = ehServico
      ? (contrato.empenhos.length
          ? contrato.empenhos.map((empenho) => {
              const movs = empenho.movimentacoes.length
                ? empenho.movimentacoes.map((mov) => `<p class="valor-visualizacao">${formatarData(mov.data)} — ${mov.tipo === "entrada" ? "Entrada" : "Saída"} de ${formatarMoeda(mov.valor_movimentado)} — saldo: ${formatarMoeda(mov.saldo_apos)}</p>`).join("")
                : '<p class="valor-visualizacao">Sem movimentações.</p>';
              const natureza = empenho.natureza_despesa ? ` — natureza ${empenho.natureza_despesa}` : "";
              return `<div class="bloco-empenho"><p class="valor-visualizacao"><strong>Empenho ${empenho.numero_empenho}</strong>${natureza}</p>${movs}</div>`;
            }).join("")
          : '<p class="valor-visualizacao">Nenhum empenho cadastrado.</p>')
      : (contrato.processos_empenho.length
          ? contrato.processos_empenho.map((processo) => {
              const empenhosDoProcesso = processo.empenhos.length
                ? processo.empenhos.map((empenho) => {
                    const natureza = empenho.natureza_despesa ? ` — natureza ${empenho.natureza_despesa}` : "";
                    return `<p class="valor-visualizacao">Empenho ${empenho.numero_empenho}${natureza}</p>`;
                  }).join("")
                : '<p class="valor-visualizacao">Nenhum empenho neste processo.</p>';
              return `<div class="bloco-empenho"><p class="valor-visualizacao"><strong>Processo ${processo.numero_processo}</strong></p>${empenhosDoProcesso}</div>`;
            }).join("")
          : '<p class="valor-visualizacao">Nenhum processo de empenho cadastrado.</p>');

    const overlay = document.createElement("div");
    overlay.className = "overlay-modal";
    overlay.innerHTML = `
      <div class="modal-visualizacao">
        <div class="modal-visualizacao__cabecalho">
          <h2>${contrato.nome_contratada}</h2>
          <button type="button" class="btn btn--outline btn--mini" data-fechar>Fechar</button>
        </div>

        <div class="painel">
          <h2>Dados do Contrato</h2>
          <div class="grade">
            ${linhaVisualizacao("Tipo de contrato", contrato.tipo_contrato === "servico" ? "Serviço" : "Almoxarifado")}
            ${linhaVisualizacao("Situação", contrato.situacao === "encerrado" ? "Encerrado" : "Vigente")}
            ${ehServico ? linhaVisualizacao("Mão de obra", contrato.tem_mao_de_obra == null ? "Não informado" : (contrato.tem_mao_de_obra ? "Com mão de obra" : "Sem mão de obra")) : ""}
            ${linhaVisualizacao("Planilha de controle", contrato.nome_planilha_controle)}
            ${linhaVisualizacao("CNPJ", formatarCnpj(contrato.cnpj || ""))}
            ${linhaVisualizacao("Nº do pregão", contrato.numero_pregao)}
            ${linhaVisualizacao("Nº do contrato", contrato.numero_contrato)}
            <div></div>
            ${linhaVisualizacao("Vigência - início", formatarData(contrato.vigencia_inicio))}
            ${linhaVisualizacao("Vigência - fim", formatarData(contrato.vigencia_fim), deveDestacarVencido(contrato) ? "vigencia-vencida" : "")}
            <div></div>
            ${linhaVisualizacao("Processo de contratação", contrato.processo_contratacao)}
            ${ehServico ? linhaVisualizacao("Processo de empenho anual", contrato.processo_empenho_anual) : ""}
            <div></div>
            <div class="campo campo--largo">${linhaVisualizacao("Objeto Resumido", contrato.objeto_resumido)}</div>
            <div class="campo campo--largo">${linhaVisualizacao("Objeto Detalhado", contrato.objeto_detalhado)}</div>
          </div>
        </div>

        <div class="painel">
          <h2>Domicílio Bancário</h2>
          <div class="grade">
            ${linhaVisualizacao("Banco", contrato.banco)}
            ${linhaVisualizacao("Agência", contrato.agencia)}
            ${linhaVisualizacao("Conta", contrato.conta)}
          </div>
        </div>

        ${ehServico ? `
        <div class="painel">
          <h2>Valores Mensais</h2>
          ${valoresMensais}
        </div>` : ""}

        <div class="painel">
          <h2>${ehServico ? "Empenhos" : "Processos de Empenho"}</h2>
          ${secaoEmpenhos}
        </div>

        <div class="painel">
          <h2>Tributação</h2>
          ${renderizarTributo("ISS", contrato.iss_incide, contrato.iss_aliquota)}
          ${renderizarTributo("Contribuição Previdenciária", contrato.previdenciaria_incide, contrato.previdenciaria_aliquota)}
          ${renderizarTributo("Tributos Federais", contrato.federais_incide, contrato.federais_aliquota_total, contrato.federais_codigo_darf ? ` — código DARF ${contrato.federais_codigo_darf}` : "")}
        </div>

        <div class="painel">
          <h2>Observação</h2>
          <p class="valor-visualizacao">${contrato.observacao || "-"}</p>
        </div>
      </div>
    `;

    const fechar = () => overlay.remove();
    overlay.addEventListener("click", (evento) => { if (evento.target === overlay) fechar(); });
    overlay.querySelector("[data-fechar]").addEventListener("click", fechar);

    document.body.appendChild(overlay);
  }

  function confirmarModal(mensagem, textoConfirmar) {
    // substitui o confirm() nativo do navegador (que no WebView2 aparece como uma caixa
    // escura fora do padrão visual do projeto) por um modal no mesmo estilo das telas
    return new Promise((resolve) => {
      const overlay = document.createElement("div");
      overlay.className = "overlay-modal";
      overlay.innerHTML = `
        <div class="modal-confirmacao">
          <p class="modal-mensagem"></p>
          <div class="modal-acoes">
            <button type="button" class="btn btn--outline" data-acao="cancelar">Cancelar</button>
            <button type="button" class="btn btn--perigo" data-acao="confirmar"></button>
          </div>
        </div>
      `;
      overlay.querySelector(".modal-mensagem").textContent = mensagem;
      overlay.querySelector('[data-acao="confirmar"]').textContent = textoConfirmar;

      const resolver = (valor) => { overlay.remove(); resolve(valor); };
      overlay.addEventListener("click", (evento) => { if (evento.target === overlay) resolver(false); });
      overlay.querySelector('[data-acao="cancelar"]').addEventListener("click", () => resolver(false));
      overlay.querySelector('[data-acao="confirmar"]').addEventListener("click", () => resolver(true));

      document.body.appendChild(overlay);
    });
  }

  async function excluirContrato(id) {
    const confirmado = await confirmarModal("Excluir este contrato? Essa ação não pode ser desfeita.", "Excluir");
    if (!confirmado) return;
    const resultado = await window.pywebview.api.excluir(id);
    if (!resultado.ok) {
      alert("Erro ao excluir: " + resultado.erro);
      return;
    }
    await carregarLista();
  }

  // ---------- linhas dinâmicas: valores mensais ----------

  function montarLinhaValor(item) {
    const linha = document.createElement("div");
    linha.className = "linha-dinamica";
    linha.innerHTML = `
      <div class="campo"><label>Serviço</label><input class="valor-descricao" value="${item?.descricao_servico ?? ""}"></div>
      <div class="campo"><label>Valor mensal (R$)</label><input type="number" step="0.01" class="valor-numero" value="${item?.valor ?? ""}"></div>
      <button type="button" class="btn btn--perigo btn--mini" title="Remover">✕</button>
    `;
    linha.querySelector("button").addEventListener("click", () => linha.remove());
    return linha;
  }

  document.getElementById("botao-add-valor").addEventListener("click", () => {
    document.getElementById("lista-valores-mensais").appendChild(montarLinhaValor());
  });

  // ---------- linhas dinâmicas: empenhos + movimentações ----------

  function recalcularSaldo(bloco) {
    const linhasMov = [...bloco.querySelectorAll(".linha-movimentacao")];
    const comData = linhasMov
      .map((linha) => ({
        linha,
        data: linha.querySelector(".mov-data").value,
        tipo: linha.querySelector(".mov-tipo").value,
        valor: Number(linha.querySelector(".mov-valor").value) || 0,
      }))
      .sort((a, b) => (a.data || "").localeCompare(b.data || ""));

    let saldo = 0;
    comData.forEach((item) => { saldo += item.tipo === "entrada" ? item.valor : -item.valor; });
    bloco.querySelector(".saldo strong").textContent = formatarMoeda(saldo);
  }

  function montarLinhaMovimentacao(bloco, mov) {
    const linha = document.createElement("div");
    linha.className = "linha-dinamica linha-movimentacao";
    linha.innerHTML = `
      <div class="campo"><label>Data</label><input type="date" class="mov-data" value="${mov?.data ?? ""}"></div>
      <div class="campo">
        <label>Tipo</label>
        <select class="mov-tipo">
          <option value="entrada" ${mov?.tipo !== "saida" ? "selected" : ""}>Entrada</option>
          <option value="saida" ${mov?.tipo === "saida" ? "selected" : ""}>Saída</option>
        </select>
      </div>
      <div class="campo"><label>Valor (R$)</label><input type="number" step="0.01" class="mov-valor" value="${mov?.valor_movimentado ?? ""}"></div>
      <button type="button" class="btn btn--perigo btn--mini" title="Remover">✕</button>
    `;
    linha.querySelectorAll("input, select").forEach((campo) => {
      campo.addEventListener("input", () => recalcularSaldo(bloco));
      campo.addEventListener("change", () => recalcularSaldo(bloco));
    });
    linha.querySelector("button").addEventListener("click", () => { linha.remove(); recalcularSaldo(bloco); });
    return linha;
  }

  function montarBlocoEmpenho(empenho) {
    const bloco = document.createElement("div");
    bloco.className = "bloco-empenho";
    bloco.innerHTML = `
      <div class="bloco-empenho__topo">
        <div class="campo"><label>Nº do empenho</label><input class="empenho-numero" value="${empenho?.numero_empenho ?? ""}"></div>
        <div class="campo"><label>Natureza de despesa</label><input class="empenho-natureza" value="${empenho?.natureza_despesa ?? ""}"></div>
        <button type="button" class="btn btn--perigo btn--mini" title="Remover empenho">✕</button>
      </div>
      <div class="movimentacoes"></div>
      <p class="saldo">Saldo atual: <strong>R$ 0,00</strong></p>
      <button type="button" class="btn btn--outline btn--mini botao-add-mov">+ Adicionar movimentação</button>
    `;
    const alvoMov = bloco.querySelector(".movimentacoes");
    (empenho?.movimentacoes ?? []).forEach((mov) => alvoMov.appendChild(montarLinhaMovimentacao(bloco, mov)));
    bloco.querySelector(".botao-add-mov").addEventListener("click", () => {
      alvoMov.appendChild(montarLinhaMovimentacao(bloco, null));
    });
    bloco.querySelector(".bloco-empenho__topo button").addEventListener("click", () => bloco.remove());
    recalcularSaldo(bloco);
    return bloco;
  }

  document.getElementById("botao-add-empenho").addEventListener("click", () => {
    document.getElementById("lista-empenhos").appendChild(montarBlocoEmpenho(null));
  });

  // ---------- linhas dinâmicas: processos de empenho + empenhos (almoxarifado, sem saldo) ----------

  function montarLinhaEmpenhoSimples(empenho) {
    const linha = document.createElement("div");
    linha.className = "linha-dinamica linha-empenho-simples";
    linha.innerHTML = `
      <div class="campo"><label>Nº do empenho</label><input class="empenho-simples-numero" value="${empenho?.numero_empenho ?? ""}"></div>
      <div class="campo"><label>Natureza de despesa</label><input class="empenho-simples-natureza" value="${empenho?.natureza_despesa ?? ""}"></div>
      <button type="button" class="btn btn--perigo btn--mini" title="Remover">✕</button>
    `;
    linha.querySelector("button").addEventListener("click", () => linha.remove());
    return linha;
  }

  function montarBlocoProcessoEmpenho(processo) {
    const bloco = document.createElement("div");
    bloco.className = "bloco-empenho";
    bloco.innerHTML = `
      <div class="bloco-empenho__topo">
        <div class="campo"><label>Nº do processo de empenho</label><input class="processo-empenho-numero" value="${processo?.numero_processo ?? ""}"></div>
        <button type="button" class="btn btn--perigo btn--mini" title="Remover processo">✕</button>
      </div>
      <div class="empenhos-do-processo"></div>
      <button type="button" class="btn btn--outline btn--mini botao-add-empenho-simples">+ Adicionar empenho</button>
    `;
    const alvoEmpenhos = bloco.querySelector(".empenhos-do-processo");
    (processo?.empenhos ?? []).forEach((empenho) => alvoEmpenhos.appendChild(montarLinhaEmpenhoSimples(empenho)));
    bloco.querySelector(".botao-add-empenho-simples").addEventListener("click", () => {
      alvoEmpenhos.appendChild(montarLinhaEmpenhoSimples(null));
    });
    bloco.querySelector(".bloco-empenho__topo button").addEventListener("click", () => bloco.remove());
    return bloco;
  }

  document.getElementById("botao-add-processo-empenho").addEventListener("click", () => {
    document.getElementById("lista-processos-empenho").appendChild(montarBlocoProcessoEmpenho(null));
  });

  // ---------- alterna Empenhos (serviço) x Processos de Empenho (almoxarifado) ----------

  function atualizarVisibilidadePorTipo() {
    const ehServico = document.getElementById("campo-tipo_contrato").value === "servico";
    document.getElementById("grupo-processo-empenho-anual").classList.toggle("oculto", !ehServico);
    document.getElementById("grupo-mao-de-obra").classList.toggle("oculto", !ehServico);
    document.getElementById("painel-empenhos-servico").classList.toggle("oculto", !ehServico);
    document.getElementById("painel-processos-empenho").classList.toggle("oculto", ehServico);
    // almoxarifado não tem valor mensal - só serviço (ex: limpeza, vigilância) tem esse conceito
    document.getElementById("painel-valores-mensais").classList.toggle("oculto", !ehServico);
  }
  document.getElementById("campo-tipo_contrato").addEventListener("change", atualizarVisibilidadePorTipo);

  // ---------- toggles de tributação ----------

  function ligarToggle(idCheckbox, idGrupo) {
    const checkbox = document.getElementById(idCheckbox);
    const grupo = document.getElementById(idGrupo);
    checkbox.addEventListener("change", () => grupo.classList.toggle("aberto", checkbox.checked));
  }
  ligarToggle("campo-iss_incide", "grupo-iss");
  ligarToggle("campo-previdenciaria_incide", "grupo-previdenciaria");
  ligarToggle("campo-federais_incide", "grupo-federais");

  // ---------- alternância lista/formulário ----------

  function mostrarLista() {
    document.getElementById("vista-lista").classList.remove("oculto");
    document.getElementById("vista-form").classList.add("oculto");
    carregarLista();
  }

  function mostrarForm() {
    document.getElementById("vista-lista").classList.add("oculto");
    document.getElementById("vista-form").classList.remove("oculto");
    document.getElementById("erro-form").classList.add("oculto");
  }

  function limparForm() {
    document.getElementById("form-contrato").reset();
    document.getElementById("campo-id").value = "";
    document.getElementById("lista-valores-mensais").innerHTML = "";
    document.getElementById("lista-empenhos").innerHTML = "";
    document.getElementById("lista-processos-empenho").innerHTML = "";
    ["iss", "previdenciaria", "federais"].forEach((grupo) => {
      document.getElementById(`grupo-${grupo}`).classList.remove("aberto");
    });
    atualizarVisibilidadePorTipo();
  }

  function abrirFormNovo() {
    limparForm();
    mostrarForm();
  }

  async function abrirFormEditar(id) {
    const contrato = await window.pywebview.api.obter(id);
    if (!contrato) return;
    limparForm();

    document.getElementById("campo-id").value = contrato.id;
    document.getElementById("campo-tipo_contrato").value = contrato.tipo_contrato;
    document.getElementById("campo-situacao").value = contrato.situacao || "vigente";
    document.getElementById("campo-tem_mao_de_obra").checked = !!contrato.tem_mao_de_obra;
    document.getElementById("campo-nome_contratada").value = contrato.nome_contratada || "";
    document.getElementById("campo-nome_planilha_controle").value = contrato.nome_planilha_controle || "";
    document.getElementById("campo-cnpj").value = formatarCnpj(contrato.cnpj || "");
    document.getElementById("campo-objeto_resumido").value = contrato.objeto_resumido || "";
    document.getElementById("campo-objeto_detalhado").value = contrato.objeto_detalhado || "";
    document.getElementById("campo-numero_pregao").value = contrato.numero_pregao || "";
    document.getElementById("campo-numero_contrato").value = contrato.numero_contrato || "";
    document.getElementById("campo-vigencia_inicio").value = contrato.vigencia_inicio || "";
    document.getElementById("campo-vigencia_fim").value = contrato.vigencia_fim || "";
    document.getElementById("campo-processo_contratacao").value = contrato.processo_contratacao || "";
    document.getElementById("campo-processo_empenho_anual").value = contrato.processo_empenho_anual || "";
    document.getElementById("campo-banco").value = contrato.banco || "";
    document.getElementById("campo-agencia").value = contrato.agencia || "";
    document.getElementById("campo-conta").value = contrato.conta || "";
    document.getElementById("campo-observacao").value = contrato.observacao || "";

    document.getElementById("campo-iss_incide").checked = !!contrato.iss_incide;
    document.getElementById("campo-iss_aliquota").value = contrato.iss_aliquota ?? "";
    document.getElementById("campo-previdenciaria_incide").checked = !!contrato.previdenciaria_incide;
    document.getElementById("campo-previdenciaria_aliquota").value = contrato.previdenciaria_aliquota ?? "";
    document.getElementById("campo-federais_incide").checked = !!contrato.federais_incide;
    document.getElementById("campo-federais_codigo_darf").value = contrato.federais_codigo_darf || "";
    document.getElementById("campo-federais_aliquota_total").value = contrato.federais_aliquota_total ?? "";
    ["iss", "previdenciaria", "federais"].forEach((grupo) => {
      document.getElementById(`grupo-${grupo}`).classList.toggle("aberto", document.getElementById(`campo-${grupo}_incide`).checked);
    });

    const alvoValores = document.getElementById("lista-valores-mensais");
    contrato.valores_mensais.forEach((item) => alvoValores.appendChild(montarLinhaValor(item)));

    const alvoEmpenhos = document.getElementById("lista-empenhos");
    contrato.empenhos.forEach((empenho) => alvoEmpenhos.appendChild(montarBlocoEmpenho(empenho)));

    const alvoProcessos = document.getElementById("lista-processos-empenho");
    contrato.processos_empenho.forEach((processo) => alvoProcessos.appendChild(montarBlocoProcessoEmpenho(processo)));

    atualizarVisibilidadePorTipo();
    mostrarForm();
  }

  function coletarDadosForm() {
    const tipo = document.getElementById("campo-tipo_contrato").value;

    // só coleta o que é relevante pro tipo atual - assim não sobra dado "fantasma" de uma
    // seção escondida se o usuário preencheu algo e depois trocou o tipo de contrato
    const valoresMensais = tipo === "servico"
      ? [...document.querySelectorAll("#lista-valores-mensais .linha-dinamica")].map((linha) => ({
          descricao_servico: linha.querySelector(".valor-descricao").value,
          valor: Number(linha.querySelector(".valor-numero").value) || 0,
        }))
      : [];

    const empenhos = tipo === "servico"
      ? [...document.querySelectorAll("#lista-empenhos .bloco-empenho")].map((bloco) => ({
          numero_empenho: bloco.querySelector(".empenho-numero").value,
          natureza_despesa: bloco.querySelector(".empenho-natureza").value,
          movimentacoes: [...bloco.querySelectorAll(".linha-movimentacao")].map((linha) => ({
            data: linha.querySelector(".mov-data").value,
            tipo: linha.querySelector(".mov-tipo").value,
            valor_movimentado: Number(linha.querySelector(".mov-valor").value) || 0,
          })),
        }))
      : [];

    const processosEmpenho = tipo === "almoxarifado"
      ? [...document.querySelectorAll("#lista-processos-empenho .bloco-empenho")].map((bloco) => ({
          numero_processo: bloco.querySelector(".processo-empenho-numero").value,
          empenhos: [...bloco.querySelectorAll(".linha-empenho-simples")].map((linha) => ({
            numero_empenho: linha.querySelector(".empenho-simples-numero").value,
            natureza_despesa: linha.querySelector(".empenho-simples-natureza").value,
          })),
        }))
      : [];

    return {
      id: document.getElementById("campo-id").value ? Number(document.getElementById("campo-id").value) : null,
      tipo_contrato: tipo,
      situacao: document.getElementById("campo-situacao").value,
      tem_mao_de_obra: tipo === "servico"
        ? document.getElementById("campo-tem_mao_de_obra").checked
        : null,
      nome_contratada: document.getElementById("campo-nome_contratada").value,
      nome_planilha_controle: document.getElementById("campo-nome_planilha_controle").value,
      cnpj: document.getElementById("campo-cnpj").value.replace(/\D/g, ""),
      objeto_resumido: document.getElementById("campo-objeto_resumido").value,
      objeto_detalhado: document.getElementById("campo-objeto_detalhado").value,
      numero_pregao: document.getElementById("campo-numero_pregao").value,
      numero_contrato: document.getElementById("campo-numero_contrato").value,
      vigencia_inicio: document.getElementById("campo-vigencia_inicio").value,
      vigencia_fim: document.getElementById("campo-vigencia_fim").value,
      processo_contratacao: document.getElementById("campo-processo_contratacao").value,
      processo_empenho_anual: tipo === "servico" ? document.getElementById("campo-processo_empenho_anual").value : null,
      banco: document.getElementById("campo-banco").value,
      agencia: document.getElementById("campo-agencia").value,
      conta: document.getElementById("campo-conta").value,
      iss_incide: document.getElementById("campo-iss_incide").checked,
      iss_aliquota: document.getElementById("campo-iss_aliquota").value || null,
      previdenciaria_incide: document.getElementById("campo-previdenciaria_incide").checked,
      previdenciaria_aliquota: document.getElementById("campo-previdenciaria_aliquota").value || null,
      federais_incide: document.getElementById("campo-federais_incide").checked,
      federais_codigo_darf: document.getElementById("campo-federais_codigo_darf").value,
      federais_aliquota_total: document.getElementById("campo-federais_aliquota_total").value || null,
      observacao: document.getElementById("campo-observacao").value,
      valores_mensais: valoresMensais,
      empenhos: empenhos,
      processos_empenho: processosEmpenho,
    };
  }

  document.getElementById("form-contrato").addEventListener("submit", async (evento) => {
    evento.preventDefault();
    const dados = coletarDadosForm();
    const resultado = await window.pywebview.api.salvar(dados);
    if (!resultado.ok) {
      const erro = document.getElementById("erro-form");
      erro.textContent = "Erro ao salvar: " + resultado.erro;
      erro.classList.remove("oculto");
      return;
    }
    mostrarLista();
  });

  document.getElementById("botao-novo").addEventListener("click", abrirFormNovo);
  document.getElementById("botao-cancelar").addEventListener("click", mostrarLista);

  window.addEventListener("pywebviewready", carregarLista);
</script>
</body>
</html>
"""

def abrir_janela():
    # cria a janela em cima da instância de webview já em execução (a principal do gui.py) -
    # não chama webview.start() de novo, pywebview permite criar janelas dinamicamente
    contratos_db.inicializar_db()
    x, y, largura, altura = janela_windows.geometria_para_tela(980, 760)
    webview.create_window("CCRGCI - Cadastro de Contratos", html=HTML_CONTRATO, js_api=ApiContrato(),
                          width=largura, height=altura, x=x, y=y)

def main(nome_planilha=None):
    # permite rodar este arquivo sozinho (fora do gui.py) pra testar a tela isolada
    contratos_db.inicializar_db()
    x, y, largura, altura = janela_windows.geometria_para_tela(980, 760)
    webview.create_window("CCRGCI - Cadastro de Contratos", html=HTML_CONTRATO, js_api=ApiContrato(),
                          width=largura, height=altura, x=x, y=y)
    webview.start()

if __name__ == "__main__":
    main()
