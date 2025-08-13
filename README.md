# Macro-Procv-Burndown
Passo a Passo no Excel
🧩 ABA: MACROS — Botões Filtrar e Valor
✅ O que esta aba demonstra
🖱️ Botão “Filtrar”: ordena alfabeticamente (A–Z) a 1ª coluna (NomeConsultorProduto).

💰 Botão “Valor”: ordena a 3ª coluna (valores) do menor para o maior.

🛠️ Como reproduzir
 Habilite a guia Desenvolvedor: Arquivo → Opções → Personalizar Faixa de Opções → marque Desenvolvedor.

 Defina o intervalo como Tabela (opcional, mas recomendado): selecione a região com cabeçalhos → Ctrl+T → “Minha tabela tem cabeçalhos”.

 Insira dois botões (formas): Inserir → Formas → retângulo; renomeie o texto para Filtrar e Valor.

 Grave/cole as macros abaixo no Editor do VBA (Alt+F11) em Módulo padrão.

 Atribua as macros aos botões: clique direito no botão → Atribuir Macro….
🧭 Uso
Clique em Filtrar para ordenar NomeConsultorProduto (A–Z).

Clique em Valor para ordenar a coluna 3 do menor → maior.

🔎 ABA: PROCV — Busca rápida Nome → Time
✅ O que esta aba demonstra
Uma forma rápida de trazer o Time a partir do Nome sem usar filtro — ideal para tabelas grandes.

🛠️ Como reproduzir
 Organize sua base com ao menos: Nome, Sexo, Time, Idade (ex.: na área B:D).

 Crie um campo de entrada para o Nome (ex.: H4) e uma célula de resultado (ex.: I4).

 Na célula de resultado, use PROCV (PT-BR) ou VLOOKUP (EN-US).

PT-BR:

excel
Copiar
Editar
=PROCV(H4;B:D;3;FALSO)
EN-US:

excel
Copiar
Editar
=VLOOKUP(H4,B:D,3,FALSE)
🔍 Busca o Nome (H4) na 1ª coluna do intervalo B:D e retorna a 3ª coluna (Time), correspondência exata (FALSO).

📉 ABA: BURNDOWN — Planejado x Realizado com Pontos
✅ O que esta aba demonstra
Como um gestor acompanha Data Planejada x Data Realizada e Pontos para avaliar se a atividade será entregue dentro do planejado.

🛠️ Como reproduzir
 Tabela de tarefas: Atividade | Pontos | Data Planejada | Data Realizada.

 Calendário do Sprint: lista de Dias em uma segunda tabela.

 Cálculos: somar Pontos Planejados (pela Data Planejada) e Pontos Realizados (pela Data Realizada).

 Curvas:

Plano Acumulado = backlog inicial − pontos planejados até o dia.

Real Acumulado = backlog inicial − pontos realizados até o dia.

 Gráfico de Linhas com 2 séries: Plano Acumulado e Real Acumulado.

📊 Leitura do gráfico
Real acima do Plano → atraso.

Real igual ou abaixo → no prazo ou adiantado.

📈 ABA: BURNDOWN2 — Gráfico simples (Azul = Planejado; Vermelho = Executado)
✅ O que esta aba demonstra
Um burndown mais simples com gráfico de linhas:

Azul = planejado.

Vermelho = executado.

Distância grande entre elas = risco.

🛠️ Como reproduzir
 3 colunas: Dia | Planejado (Azul) | Real (Vermelho).

 Inserir → Gráfico de Linhas (2D).

 Cores: azul para planejado, vermelho para real.
