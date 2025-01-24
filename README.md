
# 🖥️ Conversor de Dados GeoVictoria

  [![Python](https://img.shields.io/badge/Python-3.10%2B-blue)](https://www.python.org/)

Aplicativo desktop para conversão de planilhas Excel arquivo retirado do sistema TOTVS em formato específico para importação no GeoVictoria.

## 🚀 Funcionalidades

-  **Conversão Automática**: Transforma arquivos Excel em formato padronizado

-  **Validação de Dados**: Verifica colunas obrigatórias e formatação

-  **Interface Moderna**: GUI com tema escuro e elementos interativos

-  **Gerenciamento de Erros**: Notificações detalhadas em caso de problemas

-  **Timestamp Automático**: Nomeia arquivos com hora/minuto da conversão

  

## 📋 Requisitos do Arquivo de Entrada

-  **Formato**: `.xlsx` ou `.xls`

-  **Cabeçalho**: Deve iniciar na **linha 3** com estas colunas exatas:
| Filial | Matricula | Nome | CPF | Email Princ | Data Admis. | Desc.Funcao | C.C. Movto | Desc. Compl

> Arquivo gerado pelo sistema TOTVS. > Consultas > Cadastros > Genéricos > Monte seu relatório com essas colunas.

## ⚙️ Instalação
1. Clone o repositório
2. Instale as dependências:

```shell
    pip install -r requirements.txt
```

## 🛠️ Tecnologias Utilizadas
-   `Python 3.10+`
-   `CustomTkinter`
-   `Pandas`
-   `CTkMessagebox`

## 🗂️Estrutura do Projeto
```
DataFormatGeoVictoria/
├── src
	└── main.py
├── requirements.txt
└── README.md
```

## 🖱️ Como Usar

1. Selecione Arquivo
	- Clique em "Selecionar Arquivo Excel"
	- Escolha um arquivo válido
2. Converter Dados
	- Clique em "Converter Arquivo"
	- Aguarde a mensagem de confirmação
3. Salvar Resultado
	- Escolha o local e nome do arquivo de saída
	- Nome gerado por padrão: `ImportGeoVictoria_FORMATADO_[HORA][MINUTO].xlsx`


## 📃 Estrutura do Arquivo de Saída
Estrutura padrão do arquivo para subir no GeoVictoria
| Coluna | Origem/Regra |
|--|--|
| Identificador | CPF sem pontuação |
| Sorebnome | "RE " + Matrícula |
| Data da contratação | Data Admis.
| Grupo | C.C. Movto |
| Perfil | Valor fixo "Usuário" |
| Marcação Web/App | Valores fixos "Sim" |

## 🚨 Tratamento de Erros
-   **Colunas Faltantes**: Lista exata das colunas não encontradas
-   **Formatação Inválida**: Detecta problemas em CPF, datas e campos obrigatórios
-   **Feedback Visual**: Notificações coloridas (✅ verdes para sucesso, ❌ vermelhas para erros)

## Licença
Este projeto não possui uma licença formal, mas você pode usá-lo e modificá-lo de acordo com suas necessidades. Se você modificar o código, por favor, compartilhe as melhorias com a comunidade.