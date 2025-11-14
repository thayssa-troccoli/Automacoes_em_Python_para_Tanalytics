Automatizador de Download e Unificação de Planilhas (Flask)

Este script foi criado para automatizar a tarefa de baixar várias planilhas .xlsx hospedadas em links externos e unificar todas elas em um único arquivo Excel, cada uma em sua própria aba.

O objetivo é facilitar processos repetitivos do dia a dia, principalmente em rotinas onde é necessário juntar múltiplos arquivos de coleta/avaliação em um único documento.

🚀 Funcionalidade Principal

✔️ Baixa automaticamente 6 arquivos Excel de URLs pré-definidas
✔️ Lê cada arquivo como um DataFrame do Pandas
✔️ Cria um único Excel contendo cada arquivo como uma aba separada
✔️ Interface super simples via Flask (um botão para baixar o arquivo final)
✔️ Tudo gerado instantaneamente sem precisar baixar manualmente cada planilha

🚀 Objetivo

✔️ Facilitar o trabalho das supervisoras
✔️ Centralizar todas as bases do Rotator em um único arquivo
✔️ Garantir agilidade e padronização nos processos
✔️ Eliminar risco de baixar arquivos desatualizados
✔️ Gerar o Excel final sempre com os dados mais recentes (download direto do link no momento da requisição)

Perfeito para automações do seu fluxo de trabalho!

🏗️ Como Funciona
- A aplicação Flask mostra uma página com um único botão: “Baixar Arquivo Excel”.
- Ao clicar, o script:
- Faz o download de cada link utilizando requests
- Converte o conteúdo em DataFrames
- Cria um arquivo Excel unificado com openpyxl
- Retorna o download final (todos_links.xlsx)
