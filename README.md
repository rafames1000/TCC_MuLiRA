# MuLiRA Py-Ex - Multiple Linear Regression Analysis Python-Excel

Este software foi desenvolvido como objeto de estudo e produto prático para o Trabalho de Conclusão de Curso (TCC) de Rafael Mees, sob as orientações de Nelidy Motizuki e Ana Lucia Sanchez Panico, para o curso de Bacharelado em Engenharia de Software na UniCesumar.

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

## 💻 Como Executar a Aplicação

Este projeto oferece duas formas de execução, dependendo do seu perfil de usuário:

### Opção 1: Versão Executável (Para Usuários Finais - Windows)
A forma mais fácil de utilizar o MuLiRA, sem necessidade de instalar o Python ou configurar ambientes.

1. Acesse a seção de Downloads (Releases) clicando [neste link](https://github.com/rafames1000/TCC_MuLiRA/releases/tag/v1.0.0).
2. Baixe o arquivo `MuLiRA.exe` (ou o arquivo .zip contendo o executável).
3. Dê um duplo clique no arquivo para abrir a aplicação (nenhuma instalação adicional é necessária).

### Opção 2: A Partir do Código Fonte (Para Avaliação Técnica/Desenvolvedores)
Para auditar o código ou rodar o projeto em modo de desenvolvimento, siga os passos:

**1. Pré-requisitos:** Certifique-se de ter o Python 3.x instalado.

**2. Instalação das Dependências:**
Abra o terminal na pasta raiz do projeto e instale as bibliotecas necessárias:
```bash
pip install -r requirements.txt
```
### Opção 3: Inicialização:
Execute o arquivo principal para abrir a interface gráfica:
```bash
python TCC_MuLiRA.py
```
