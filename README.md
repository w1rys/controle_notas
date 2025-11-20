# Controle de Notas – Automação de NF-e (XML)
A loja de bicicletas onde trabalho não tinha nenhum sistema para acompanhar preços, produtos e histórico de compras. Este projeto foi criado para suprir essa ausência, automatizando a leitura de notas fiscais (XML) e atualizando automaticamente uma planilha estruturada com todas as informações necessárias.

Este projeto realiza a leitura automática de notas fiscais eletrônicas (NF-e) no formato XML, extrai os dados dos produtos e atualiza uma planilha Excel organizada para controle de compras e preços. O sistema foi desenvolvido para uso em lojas que não possuem um sistema de gestão e desejam automatizar o registro de produtos adquiridos.


## Funcionalidades

- Monitoramento automático da pasta de entrada de notas (via Watchdog)
- Leitura estruturada dos arquivos XML (NF-e)
- Extração de informações: código, nome, quantidade, preços e data de compra
- Prevenção de duplicidade via chave da nota fiscal
- Atualização da aba **Compras** no Excel
- Atualização da aba **Produtos** contendo:
  - último preço de compra
  - penúltimo preço automaticamente calculado
  - preservação do preço de venda editado manualmente
- Movimentação automática das notas processadas

---

## Estrutura do Projeto

controle_notas/
│
├── src/
│ ├── ler_notas.py
│ ├── atualizar_excel.py
│ ├── monitor.py
│ ├── utils.py
│
├── notas/
├── notas_processadas/
├── notas_invalidas/
│
├── produtos.xlsx (gerado automaticamente)
├── requirements.txt
└── README.md

---

## Instalação

1. Crie um ambiente virtual:

```bash
python3 -m venv venv
source venv/bin/activate      # macOS / Linux
venv\Scripts\activate         # Windows
Instale as dependências:
pip install -r requirements.txt
Uso
Executar o monitoramento de notas
python3 src/monitor.py
O sistema ficará observando a pasta notas/.
Sempre que um arquivo XML válido for adicionado, ele será processado.
Processamento manual de todas as notas
python3 src/ler_notas.py
Estrutura do Excel
Aba "Compras"
Contém o histórico completo de compras:
código
nome do produto
quantidade total comprada
último preço
penúltimo preço
data da última compra
chave da nota
Aba "Produtos"
Contém informações para consulta e uso no dia a dia:
código
nome do produto
último preço
penúltimo preço
preço de venda (preenchido manualmente e preservado)
Notas
Os arquivos XML processados são movidos para notas_processadas/.
Notas duplicadas (mesma chave) são ignoradas automaticamente.
Datas são convertidas para formato sem timezone para evitar erros de Excel.
Dependências
Listadas em requirements.txt:
pandas
openpyxl
xmltodict
watchdog
