
# 🛍️ Automação de Indicadores - Projeto OnePage

## 📌 Objetivo

Este projeto tem como foco o desenvolvimento de um sistema automatizado para gerar e enviar relatórios diários de desempenho (OnePages) das lojas de uma grande rede de varejo. Ele representa um exercício completo de automação de processos utilizando Python e manipulação de dados.

## 🧾 Descrição

Imagine que você trabalha em uma rede de lojas de roupas com 25 unidades espalhadas por todo o Brasil. Todas as manhãs, a equipe de análise de dados calcula os chamados **OnePages** — resumos diários contendo os principais indicadores de desempenho de cada loja.

O **OnePage** permite, em uma única página, visualizar o desempenho de cada loja, identificar metas atingidas e comparar os indicadores entre unidades. O desafio é automatizar completamente esse processo: do cálculo dos indicadores até o envio por e-mail personalizado para cada gerente.

## 📂 Estrutura e Arquivos do Projeto

- `Emails.xlsx`  
  Contém nome, loja e e-mail de cada gerente.  
  Obs: recomenda-se substituir os e-mails reais por e-mails de teste.

- `Vendas.xlsx`  
  Contém os dados de vendas de todas as lojas.

- `Lojas.csv`  
  Lista com o nome de cada loja.

## ✉️ Envio de E-mails

Ao final do processo, o sistema deve enviar:
- Um e-mail para **cada gerente** com:
  - O OnePage de sua loja no corpo do e-mail
  - Um anexo com os dados completos da loja
- Um e-mail consolidado para a **diretoria** com todos os arquivos

## 🛠️ Tecnologias e Bibliotecas Utilizadas

- Python 3
- Pandas
- OpenPyXL
- smtplib / email (bibliotecas para envio de e-mails)
- Jupyter Notebook

## 🚀 Como Executar

1. Instale os requisitos:
   ```bash
   pip install pandas openpyxl
   ```

2. Atualize os arquivos com dados reais (ou de teste).

3. Execute o notebook `Descrição do Projeto.ipynb` para processar os dados e enviar os e-mails.

## 🔐 Observações

- Certifique-se de permitir acesso de aplicativos menos seguros no serviço de e-mail que for utilizar para testes.
- Proteja credenciais e dados sensíveis, utilizando arquivos `.env` ou variáveis de ambiente.

---

📧 Projeto idealizado para fins educacionais e de prática em automação de processos de dados.
