📄 Excel → SIDE TXT Converter
Automação para Estruturação de Analíticos de Contribuições (VGBL)

Ferramenta em Python desenvolvida para transformar planilhas Excel contendo múltiplas contribuições (N linhas) em um arquivo TXT estruturado no padrão exigido pelo SIDE – Sistema para Intercâmbio de Documentos Eletrônicos, implementado pela FenaPrevi.

O objetivo é automatizar a padronização e geração do arquivo analítico necessário para importação de contribuições no processo de portabilidade de Previdência VGBL.

🎯 Contexto de Negócio

No processo de portabilidade entre entidades previdenciárias, o analítico de contribuições precisa ser:

Estruturado em layout específico

Padronizado conforme regras acordadas no mercado

Formatado corretamente para importação sistêmica

O SIDE (Sistema para Intercâmbio de Documentos Eletrônicos), implementado pela FenaPrevi, padroniza essa troca entre entidades.

Em cenários de ajustes ou retificações, os dados podem chegar em formato Excel despadronizado, exigindo:

Tratamento manual

Reorganização de colunas

Conversão de datas

Conversão de valores

Estruturação linha a linha

Este projeto elimina esse retrabalho manual, estruturando automaticamente as N linhas de contribuições da proposta e gerando o TXT pronto para importação.

⚙️ O que a aplicação faz

✔ Detecta automaticamente a coluna de DATA
✔ Detecta duas colunas monetárias adjacentes
✔ Converte datas para o padrão AAAAMMDD
✔ Converte valores monetários para centavos inteiros
✔ Remove possíveis linhas de totalização no rodapé
✔ Gera arquivo TXT em layout fixo (1000 caracteres por linha)
✔ Estrutura corretamente todas as N contribuições da proposta
✔ Mantém compatibilidade com o padrão SIDE

📊 Estrutura do Fluxo

Recebe planilha Excel com múltiplas contribuições

Detecta automaticamente os campos relevantes

Normaliza e padroniza os dados

Estrutura linha a linha conforme layout fixo

Gera TXT pronto para importação no processo de portabilidade

🛠 Tecnologias Utilizadas

Python 3

pandas

openpyxl

xlrd

msoffcrypto-tool

CLI interativo

📂 Estrutura do Projeto
entradas/  → planilhas Excel recebidas  
saidas/    → TXT gerado no padrão SIDE  

▶️ Como Executar

Coloque o arquivo Excel na pasta entradas/

Execute o script

Escolha o arquivo desejado

O TXT será gerado automaticamente na pasta saidas/

💡 Diferenciais Técnicos

Detecção automática de colunas (reduz dependência de layout fixo no Excel)

Tratamento de arquivos protegidos por senha

Validação inteligente de linha totalizadora

Estruturação em layout fixo com padding correto

Compatibilidade com padrão de intercâmbio do mercado previdenciário

🚀 Impacto Operacional

Redução de retrabalho manual

Mitigação de erro humano

Agilidade na geração do analítico

Maior confiabilidade no processo de portabilidade

Padronização consistente dos dados

👤 Autor

Arlindo Júnior Honorato
Product Owner | Automação | IA aplicada a processos financeiros