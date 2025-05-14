{
 "cells": [
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "# Automação de Indicadores\n",
    "\n",
    "### Objetivo: Treinar e criar um Projeto Completo que envolva a automatização de um processo feito no computador\n",
    "\n",
    "### Descrição:\n",
    "\n",
    "Imagine que você trabalha em uma grande rede de lojas de roupa com 25 lojas espalhadas por todo o Brasil.\n",
    "\n",
    "Todo dia, pela manhã, a equipe de análise de dados calcula os chamados One Pages e envia para o gerente de cada loja o OnePage da sua loja, bem como todas as informações usadas no cálculo dos indicadores.\n",
    "\n",
    "Um One Page é um resumo muito simples e direto ao ponto, usado pela equipe de gerência de loja para saber os principais indicadores de cada loja e permitir em 1 página (daí o nome OnePage) tanto a comparação entre diferentes lojas, quanto quais indicadores aquela loja conseguiu cumprir naquele dia ou não.\n"
   ]
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "O seu papel, como Analista de Dados, é conseguir criar um processo da forma mais automática possível para calcular o OnePage de cada loja e enviar um email para o gerente de cada loja com o seu OnePage no corpo do e-mail e também o arquivo completo com os dados da sua respectiva loja em anexo."
   ]
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "### Arquivos e Informações Importantes\n",
    "\n",
    "- Arquivo Emails.xlsx com o nome, a loja e o e-mail de cada gerente. Obs: Sugiro substituir a coluna de e-mail de cada gerente por um e-mail seu, para você poder testar o resultado\n",
    "\n",
    "- Arquivo Vendas.xlsx com as vendas de todas as lojas. \n",
    "\n",
    "- Arquivo Lojas.csv com o nome de cada Loja\n",
    "\n",
    "- Ao final, sua rotina deve enviar ainda um e-mail para a diretoria (informações também estão no arquivo Emails.xlsx) com 2 rankings das lojas em anexo, 1 ranking do dia e outro ranking anual. O ranking de uma loja é dado pelo faturamento da loja.\n",
    "\n",
    "- As planilhas de cada loja devem ser salvas dentro da pasta da loja com a data da planilha, a fim de criar um histórico de backup.\n",
    "\n",
    "### Indicadores do OnePage\n",
    "\n",
    "- Faturamento -> Meta Ano: 1.650.000 / Meta Dia: 1000\n",
    "- Diversidade de Produtos (quantos produtos diferentes foram vendidos naquele período) -> Meta Ano: 120 / Meta Dia: 4\n",
    "- Ticket Médio por Venda -> Meta Ano: 500 / Meta Dia: 500\n"
   ]
  }
 ],
 "metadata": {
  "kernelspec": {
   "display_name": "Python 3",
   "language": "python",
   "name": "python3"
  },
  "language_info": {
   "codemirror_mode": {
    "name": "ipython",
    "version": 3
   },
   "file_extension": ".py",
   "mimetype": "text/x-python",
   "name": "python",
   "nbconvert_exporter": "python",
   "pygments_lexer": "ipython3",
   "version": "3.8.3"
  }
 },
 "nbformat": 4,
 "nbformat_minor": 4
}
