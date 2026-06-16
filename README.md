# MuLiRA Py-Ex - Multiple Linear Regression Analysis Python-Excel

Este software foi desenvolvido como objeto de estudo e produto prático para o Trabalho de Conclusão de Curso (TCC) de Rafael Mees, sob a orientação de Ana Lucia Sanchez Panico, para o curso de Bacharelado em Engenharia de Software na UniCesumar.

O **MuLiRA** é uma aplicação desktop com interface gráfica voltada para automação de processos de Regressão Linear Múltipla (RLM). Ele gerencia de forma inteligente a importação de dados textuais e hifens de planilhas eletrônicas, forçando conversões e aplicando validações matemáticas robustas antes da modelagem estatística.

## 🚀 Funcionalidades Principais

* **Fatiamento Dinâmico:** Seleção customizável de blocos contínuos e discretos de linhas e colunas para análise.
* **Limpeza e Higienização de Dados (Tratamento):** Conversão automática de dados corrompidos/texto em números (`NaN`) e expurgo de linhas inválidas através de Pandas.
* **Trava de Consistência Estatística:** Bloqueio preventivo de execuções com dados insuficientes (exige o mínimo de 8 linhas e garante que o número de observações superará o número de preditores para evitar falhas nos graus de liberdade).
* **Modelagem e Diagnóstico:** Geração automática do sumário estatístico completo (OLS/Statsmodels) e gráficos de diagnóstico integrados (Reais vs. Previstos e Homocedasticidade/Resíduos).

## 🛠️ Tecnologias e Dependências

A aplicação foi inteiramente desenvolvida em **Python 3** utilizando as seguintes bibliotecas:
* **Tkinter & TTK:** Interface Gráfica (GUI) nativa.
* **Pandas:** Estruturação de dados e higienização.
* **Statsmodels:** Motor de cálculos estatísticos e Regressão OLS.
* **Matplotlib:** Renderização de gráficos de diagnóstico integrados na GUI.

## 💻 Como Executar o Projeto

Para testar ou rodar a aplicação localmente a partir do código fonte, siga os passos abaixo:

### 1. Pré-requisitos
Certifique-se de ter o Python instalado em sua máquina.

### 2. Instalação das Dependências
Abra o seu terminal/prompt de comando na pasta do projeto e execute:
```bash
pip install -r requirements.txt
```

### 3. Inicialização
Após a instalação das dependências, execute o arquivo principal:

```bash
python TCC_MuLiRA.py
```
