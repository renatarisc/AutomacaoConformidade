import webview  # mesma lib do gui.py - aqui só abre uma segunda janela sobre a mesma instância

import bolsas_db
import janela_windows

class ApiBolsa:
    # ponte Python <-> JS desta janela (window.pywebview.api.<metodo>), independente da Api
    # principal do gui.py - essa janela é autocontida (lista + formulário + banco local)
    def listar(self):
        return bolsas_db.listar_bolsas()

    def obter(self, bolsa_id):
        return bolsas_db.obter_bolsa(bolsa_id)

    def salvar(self, dados):
        try:
            if dados.get("id"):
                bolsas_db.atualizar_bolsa(dados["id"], dados)
                bolsa_id = dados["id"]
            else:
                bolsa_id = bolsas_db.criar_bolsa(dados)
            return {"ok": True, "id": bolsa_id}
        except Exception as e:
            return {"ok": False, "erro": str(e)}

    def excluir(self, bolsa_id):
        try:
            bolsas_db.excluir_bolsa(bolsa_id)
            return {"ok": True}
        except Exception as e:
            return {"ok": False, "erro": str(e)}

HTML_BOLSA = r"""
<!doctype html>
<html lang="pt-br">
<head>
<meta charset="utf-8">
<title>CCRGCI - Cadastro de Bolsas</title>
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
  .th-ordenar { font: inherit; color: inherit; text-transform: inherit; letter-spacing: inherit; background: none; border: 0; padding: 0; cursor: pointer; display: inline-flex; align-items: center; gap: 4px; }
  .th-ordenar:hover { color: var(--pine-deep); }
  .th-ordenar .seta { color: var(--ink-faint); font-size: 10px; }
  .th-ordenar[data-ativo="1"] .seta { color: var(--pine-deep); }
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

  .campo { display: flex; flex-direction: column; gap: 4px; min-width: 0; }
  label { font-size: 11.5px; font-weight: 600; color: var(--ink-soft); }

  .linha-campos { display: flex; flex-wrap: wrap; gap: 10px 12px; margin-bottom: 10px; }
  .linha-campos .campo { min-width: 0; }
  .campo--flexivel { flex: 1 1 200px; }
  .campo--estreito { flex: 0 0 160px; }

  input, select {
    font-family: inherit; font-size: 13px; padding: 7px 9px;
    border: 1px solid var(--hairline); border-radius: 6px; background: var(--cloud); color: var(--ink);
  }
  input:focus-visible, select:focus-visible { outline: none; border-color: var(--pine); box-shadow: 0 0 0 3px var(--pine-tint-strong); }

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
</style>
</head>
<body>

<div class="pagina">
  <div class="cabecalho">
    <div class="cabecalho__marca">
      <div class="marca-icone"><svg viewBox="0 0 24 24" width="19" height="19" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M12 2 2 7l10 5 10-5-10-5Z"/><path d="M2 17l10 5 10-5"/><path d="M2 12l10 5 10-5"/></svg></div>
      <h1>Cadastro de Bolsas</h1>
    </div>
    <button class="btn btn--acento" id="botao-novo">+ Nova bolsa</button>
  </div>

  <div id="vista-lista">
    <div class="painel painel--acento">
      <p class="rotulo"><span class="ponto"></span>Filtros</p>
      <div class="linha-campos" style="margin-bottom: 0;">
        <div class="campo campo--flexivel">
          <label>Tipo de bolsa</label>
          <input type="text" id="filtro-tipo" placeholder="Ex: Bolsa Atleta...">
        </div>
        <div class="campo campo--estreito">
          <label>Edital</label>
          <input type="text" id="filtro-edital" placeholder="Ex: 06/2026...">
        </div>
      </div>
    </div>
    <div id="lista-bolsas"></div>
  </div>

  <div id="vista-form" class="oculto">
    <p class="erro-form oculto" id="erro-form"></p>
    <form id="form-bolsa">
      <input type="hidden" id="campo-id">

      <div class="painel">
        <h2>Dados da Bolsa</h2>
        <div class="linha-campos">
          <div class="campo campo--flexivel">
            <label>Tipo de bolsa</label>
            <input id="campo-tipo_bolsa" required>
          </div>
          <div class="campo campo--estreito">
            <label>Processo anual</label>
            <input id="campo-processo_anual">
          </div>
          <div class="campo campo--estreito">
            <label>Edital</label>
            <input id="campo-edital">
          </div>
        </div>
        <div class="linha-campos">
          <div class="campo campo--flexivel">
            <label>Valor unitário</label>
            <input id="campo-valor_unitario" placeholder="Ex: 250,00 ou 250,00 e 300,00">
          </div>
          <div class="campo campo--flexivel">
            <label>Empenho</label>
            <input id="campo-empenho">
          </div>
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
  // ---------- vista de lista ----------

  let bolsasCarregadas = [];
  let ordenacao = { campo: null, dir: 1 };

  async function carregarLista() {
    bolsasCarregadas = await window.pywebview.api.listar();
    aplicarFiltros();
  }

  function ordenarPor(campo) {
    if (ordenacao.campo === campo) ordenacao.dir *= -1;
    else ordenacao = { campo, dir: 1 };
    aplicarFiltros();
  }

  function aplicarFiltros() {
    const tipoFiltro = document.getElementById("filtro-tipo").value.trim().toLocaleLowerCase("pt-BR");
    const editalFiltro = document.getElementById("filtro-edital").value.trim().toLocaleLowerCase("pt-BR");

    const filtrados = bolsasCarregadas.filter((bolsa) => {
      const bateTipo = !tipoFiltro || (bolsa.tipo_bolsa || "").toLocaleLowerCase("pt-BR").includes(tipoFiltro);
      const bateEdital = !editalFiltro || (bolsa.edital || "").toLocaleLowerCase("pt-BR").includes(editalFiltro);
      return bateTipo && bateEdital;
    });

    if (ordenacao.campo) {
      filtrados.sort((a, b) => ordenacao.dir * String(a[ordenacao.campo] || "").localeCompare(
        String(b[ordenacao.campo] || ""), "pt-BR", { numeric: true, sensitivity: "base" }));
    }

    renderizarTabela(filtrados);
  }

  document.getElementById("filtro-tipo").addEventListener("input", aplicarFiltros);
  document.getElementById("filtro-edital").addEventListener("input", aplicarFiltros);

  function renderizarTabela(bolsas) {
    const alvo = document.getElementById("lista-bolsas");

    if (!bolsasCarregadas.length) {
      alvo.innerHTML = '<p class="vazio">Nenhuma bolsa cadastrada ainda.</p>';
      return;
    }
    if (!bolsas.length) {
      alvo.innerHTML = '<p class="vazio">Nenhuma bolsa encontrada para o filtro atual.</p>';
      return;
    }

    const linhas = bolsas.map((bolsa) => `
      <tr>
        <td>${bolsa.tipo_bolsa}</td>
        <td>${bolsa.processo_anual || "-"}</td>
        <td>${bolsa.edital || "-"}</td>
        <td>${bolsa.valor_unitario || "-"}</td>
        <td>${bolsa.empenho || "-"}</td>
        <td class="acoes">
          <button class="btn btn--acento btn--mini" data-editar="${bolsa.id}">Editar</button>
          <button class="btn btn--perigo btn--mini" data-excluir="${bolsa.id}">Excluir</button>
        </td>
      </tr>
    `).join("");

    const thOrdenavel = (campo, rotulo) => {
      const ativo = ordenacao.campo === campo;
      const seta = ativo ? (ordenacao.dir === 1 ? "▲" : "▼") : "↕";
      return `<th><button type="button" class="th-ordenar" data-ordenar="${campo}" data-ativo="${ativo ? 1 : 0}">${rotulo}<span class="seta">${seta}</span></button></th>`;
    };

    alvo.innerHTML = `
      <table>
        <thead><tr>
          ${thOrdenavel("tipo_bolsa", "Tipo de bolsa")}
          ${thOrdenavel("processo_anual", "Processo anual")}
          ${thOrdenavel("edital", "Edital")}
          ${thOrdenavel("valor_unitario", "Valor unitário")}
          ${thOrdenavel("empenho", "Empenho")}
          <th></th>
        </tr></thead>
        <tbody>${linhas}</tbody>
      </table>
    `;

    alvo.querySelectorAll("[data-ordenar]").forEach((botao) => {
      botao.addEventListener("click", () => ordenarPor(botao.dataset.ordenar));
    });
    alvo.querySelectorAll("[data-editar]").forEach((botao) => {
      botao.addEventListener("click", () => abrirFormEditar(Number(botao.dataset.editar)));
    });
    alvo.querySelectorAll("[data-excluir]").forEach((botao) => {
      botao.addEventListener("click", () => excluirBolsa(Number(botao.dataset.excluir)));
    });
  }

  function confirmarModal(mensagem, textoConfirmar) {
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

  async function excluirBolsa(id) {
    const confirmado = await confirmarModal("Excluir esta bolsa? Essa ação não pode ser desfeita.", "Excluir");
    if (!confirmado) return;
    const resultado = await window.pywebview.api.excluir(id);
    if (!resultado.ok) {
      alert("Erro ao excluir: " + resultado.erro);
      return;
    }
    await carregarLista();
  }

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
    document.getElementById("form-bolsa").reset();
    document.getElementById("campo-id").value = "";
  }

  function abrirFormNovo() {
    limparForm();
    mostrarForm();
  }

  async function abrirFormEditar(id) {
    const bolsa = await window.pywebview.api.obter(id);
    if (!bolsa) return;
    limparForm();

    document.getElementById("campo-id").value = bolsa.id;
    document.getElementById("campo-tipo_bolsa").value = bolsa.tipo_bolsa || "";
    document.getElementById("campo-processo_anual").value = bolsa.processo_anual || "";
    document.getElementById("campo-edital").value = bolsa.edital || "";
    document.getElementById("campo-valor_unitario").value = bolsa.valor_unitario ?? "";
    document.getElementById("campo-empenho").value = bolsa.empenho || "";

    mostrarForm();
  }

  function coletarDadosForm() {
    return {
      id: document.getElementById("campo-id").value ? Number(document.getElementById("campo-id").value) : null,
      tipo_bolsa: document.getElementById("campo-tipo_bolsa").value,
      processo_anual: document.getElementById("campo-processo_anual").value,
      edital: document.getElementById("campo-edital").value,
      valor_unitario: document.getElementById("campo-valor_unitario").value || null,
      empenho: document.getElementById("campo-empenho").value,
    };
  }

  document.getElementById("form-bolsa").addEventListener("submit", async (evento) => {
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
    bolsas_db.inicializar_db()
    x, y, largura, altura = janela_windows.geometria_para_tela(900, 620)
    webview.create_window("CCRGCI - Cadastro de Bolsas", html=HTML_BOLSA, js_api=ApiBolsa(),
                          width=largura, height=altura, x=x, y=y)

def main():
    # permite rodar este arquivo sozinho (fora do gui.py) pra testar a tela isolada
    bolsas_db.inicializar_db()
    x, y, largura, altura = janela_windows.geometria_para_tela(900, 620)
    webview.create_window("CCRGCI - Cadastro de Bolsas", html=HTML_BOLSA, js_api=ApiBolsa(),
                          width=largura, height=altura, x=x, y=y)
    webview.start()

if __name__ == "__main__":
    main()
