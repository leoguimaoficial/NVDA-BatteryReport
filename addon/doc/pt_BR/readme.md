# NVDA Battery Report

O **NVDA Battery Report** é um complemento que gera e lê, de forma acessível, o **relatório de bateria do Windows** (`powercfg /batteryreport`) diretamente no NVDA. Ele apresenta **saúde da bateria**, **capacidades**, **histórico de uso** e **estimativas de duração**, com resultados falados, interface simples e histórico integrado.

---

## Funcionalidades

* **Gerar relatório oficial do Windows** com um clique, totalmente local.
* **Resumo falado** ao finalizar: saúde da bateria e capacidades.
* **Seções organizadas**:

  * Visão geral
  * Bateria instalada
  * Uso recente (7 dias)
  * **Consumo de bateria (7 dias)**
  * Histórico de capacidade
  * Histórico de uso
  * Estimativas de vida útil (com médias)
* **Listas acessíveis e descrições**: cada linha da tabela vira texto corrido com **legenda das colunas**.
* **Ordenação** (mais novo/mais antigo) e **limite de linhas** (10, 20, 30… até o máximo disponível).
* **Copiar item selecionado** para a área de transferência.
* **Abrir o HTML original** do Windows (para conferência).
* **Suporte a múltiplos idiomas** (.po/.mo).

---

## Conceitos importantes

Para evitar ambiguidade:

* **Capacidade de projeto (fábrica)**: quanto a bateria suportava **quando nova** (*Design capacity*).
* **Capacidade máxima atual (restante)**: quanto a bateria suporta **hoje**, após desgaste (*Full charge capacity*).

**Saúde da bateria** = Capacidade máxima atual ÷ Capacidade de projeto × 100%.

Em **Estimativas de vida útil**, você verá tempos previstos **Em carga completa** e **Na capacidade de projeto**, tanto para **Ativo** quanto para **Espera conectada**. O complemento também calcula **médias** com base nos períodos disponíveis.

---

## Como usar

1. Abra pelo NVDA: **Ferramentas → NVDA Battery Report**.
   *(Não há atalho padrão; veja “Atalhos” abaixo para configurar.)*
2. Na janela principal, clique em **Gerar relatório**.
   Ao terminar, o NVDA anuncia um resumo e o item aparece no **Histórico**.
3. Selecione um relatório no **Histórico** e clique em **Ver detalhes**.
4. Na janela de detalhes:

   * Selecione a **Seção** desejada.
   * Em tabelas grandes, ajuste **Linhas** (10, 20, 30…) e **Ordem** (Mais novo/Antigo).
   * A **lista** mostra cada linha pronta para leitura; a **descrição** repete em **texto corrido** e inclui a **legenda das colunas**.
   * Use **Copiar selecionado** para enviar o item para a área de transferência.
   * **Abrir HTML bruto** abre o relatório original do Windows no navegador.
5. Para remover relatórios, use **Excluir**; para apagar tudo, **Limpar histórico**.

---

## Dicas de leitura (seções)

* **Uso recente (7 dias)**: estados de energia ao longo do tempo (ativo/espera, fonte de energia e carga restante).
* **Consumo de bateria (7 dias)**: mostra **início**, **estado**, **duração** e **energia consumida**.
  *Se não houver consumo no período, o complemento informa claramente que não há entradas.*
* **Histórico de capacidade**: **Período**, **Capacidade máxima atual** e **Capacidade de projeto**.
* **Histórico de uso**: períodos com tempos **Em bateria** e **Em CA**.
* **Estimativas de vida útil**: por período, tempos **Em carga completa** e **Na capacidade de projeto**, em **Ativo** e **Espera conectada**, com **médias** no topo.

---

## Atalhos (definir no NVDA)

1. NVDA → **Preferências → Definir comandos**.
2. Procure por **“NVDA Battery Report”**.
3. Associe o gesto que preferir (ex.: `NVDA+Shift+B`).

---

## Onde ficam os arquivos

* Relatórios HTML: `…\addons\NVDABatteryReport\globalPlugins\battery_reports\`
* Histórico (JSON): `…\addons\NVDABatteryReport\globalPlugins\battery_history.json`

*(Caminhos dentro do perfil do NVDA do usuário.)*

---

## Perguntas frequentes (FAQ)

* **Preciso de internet?**
  Não. Tudo é gerado localmente pelo Windows (`powercfg`).

* **Funciona em desktop sem bateria?**
  O relatório pode indicar que não há bateria e algumas seções ficarão vazias — isso é esperado.

* **Por que aparecem “-” em alguns campos?**
  Significa que o Windows não registrou dados suficientes naquele período.

* **A ordem não combina com a do arquivo HTML.**
  O complemento ordena por **data real** (mais novos primeiro por padrão). Você pode inverter em **Ordem**.

* **As datas parecem “diferentes”.**
  Elas seguem o **formato regional do Windows** configurado na sua máquina.

---

## Solução de problemas

* **“powercfg.exe not found.”**
  Verifique a integridade do Windows. O utilitário fica em `System32/Sysnative`.

* **Erro ao abrir o HTML bruto.**
  Confirme se o arquivo existe na pasta de relatórios. Gere um novo se necessário.

* **Itens “desconhecidos” na lista.**
  Gere um relatório recente; seções sem dados mostram mensagens claras (ex.: “Sem entradas nos últimos 7 dias.”).

---

## Changelog

### 1.0

* Geração do relatório oficial do Windows.
* Seções com **legenda**, **ordem configurável** e **limite de linhas**.
* **Capacidade de projeto (fábrica)** vs **Capacidade máxima atual (restante)**, com **saúde** calculada.
* **Consumo de bateria (7 dias)** com mensagens claras quando não houver registros.
* **Estimativas de vida** com **médias**.
* **Histórico** integrado, **copiar selecionado** e **abrir HTML**.

---

## Suporte

Dúvidas ou sugestões? Abra uma **issue** no repositório:
[https://github.com/leoguimaoficial/NVDA-BatteryReport](https://github.com/leoguimaoficial/NVDA-BatteryReport)

---

Add-on criado por **Leo Guima** — 2025
