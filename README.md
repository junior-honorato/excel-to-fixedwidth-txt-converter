📄 Excel → SIDE TXT Converter
Automação em Python para transformar planilhas Excel despadronizadas em arquivos TXT estruturados no padrão exigido pelo SIDE (FenaPrevi).

📌 Problema
No processo de portabilidade entre entidades previdenciárias, o analítico de contribuições precisa trafegar em um layout muito específico.
Porém:
Em cenários de ajustes ou retificações, os dados chegam das entidades em formato Excel totalmente despadronizado.
Isso exige tratamento manual de colunas, formatação de datas e conversão de valores (para centavos inteiros).
A estruturação linha a linha no padrão fixo de 1000 caracteres é complexa e suscetível a falhas.
O retrabalho manual atrasa a importação sistêmica e aumenta o risco de erro humano em dados financeiros críticos.

🎯 Objetivo da Ferramenta
Automatizar a padronização e geração do arquivo analítico, eliminando o trabalho manual de conversão e garantindo que os dados fiquem prontos para importação imediata no processo de portabilidade de Previdência VGBL.

A aplicação:
Lê a planilha Excel com as múltiplas contribuições na pasta `entradas/`
Detecta automaticamente a coluna de DATA e as colunas monetárias
Converte as datas para o padrão exigido (AAAAMMDD)
Transforma os valores monetários em centavos inteiros
Remove possíveis linhas de totalização no rodapé (lixo de formatação)
Estrutura e gera o arquivo TXT em layout fixo na pasta `saidas/`

🧪 Exemplo de Execução
Lendo arquivo recebido: analitico_portabilidade.xlsx…
✔ Coluna de datas identificada.
✔ Colunas de valores identificadas.
Convertendo dados e formatando layout…
✔ Sucesso! Arquivo TXT padrão SIDE gerado na pasta /saidas.

💼 Impacto no Negócio
A ferramenta contribui diretamente para:
Redução drástica do retrabalho manual no tratamento de planilhas
Mitigação de erro humano em dados financeiros (datas e valores)
Agilidade na geração do analítico para importação sistêmica
Maior confiabilidade e segurança no processo de portabilidade
Padronização consistente dos dados trocados entre entidades
Independência de layouts fixos de Excel, já que a detecção de colunas é inteligente

⚙️ Funcionalidades
✔ Detecção automática de colunas relevantes (reduz dependência de layout fixo)
✔ Conversão de datas e valores monetários para o padrão SIDE
✔ Tratamento de arquivos protegidos por senha
✔ Validação inteligente e remoção de linha totalizadora
✔ Estruturação em layout fixo com padding correto (1000 caracteres/linha)
✔ Interface simples e direta (CLI)

🛠 Tecnologias Utilizadas
Python 3
pandas
openpyxl / xlrd
msoffcrypto-tool (para arquivos protegidos)
CLI interativo

🖥️ Como usar
Coloque o arquivo Excel despadronizado na pasta `entradas/`.
Execute o script principal (`python seu_script.py`).
Siga as instruções na tela para escolher o arquivo.
O arquivo TXT formatado será gerado automaticamente na pasta `saidas/` pronto para uso.

📂 Estrutura do Projeto
conversor-excel-txt-side/
├── excel_to_TXT.py
├── entradas/
├── saidas/
└── README.md

🤖 Uso de Inteligência Artificial
A IA generativa foi utilizada como copiloto técnico, auxiliando principalmente em: estruturação da lógica de formatação, revisão de código e escrita de documentação.

O conhecimento do negócio (regras do SIDE, estrutura de portabilidade VGBL e tratamento das exceções das planilhas) foi aplicado manualmente.

👤 Autor
Arlindo Júnior Honorato
Product Owner | Automação | IA aplicada a processos financeiros e previdenciários