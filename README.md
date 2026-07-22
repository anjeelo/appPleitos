# Sistema de Geração de Cartas de Pleito e Croquis

![Versão](https://img.shields.io/badge/version-2.0.0-blue)
![Python](https://img.shields.io/badge/python-3.11%2B-green)
![Flask](https://img.shields.io/badge/flask-2.3.0-red)
![License](https://img.shields.io/badge/license-MIT-yellow)

Sistema web para geração automatizada de documentos a partir de planilhas Excel. Reúne duas ferramentas:

- **Carta Pleito** — a partir de uma planilha de ocorrências, gera cartas consolidadas (Nacional e Regional) em Word e uma planilha de acompanhamento formatada.
- **Criação de Croqui** — a partir de uma planilha de entrada, gera um arquivo de croqui em Excel para cada obra, usando um modelo fixo já cadastrado no servidor.

## Funcionalidades

- Upload de planilhas Excel (.xlsx, .xls) por arrastar-e-soltar ou seleção manual
- Geração de cartas Word consolidadas, separadas por tipo (Nacional/Regional)
- Geração de um croqui por linha da planilha de entrada, nomeado automaticamente
- Planilha de acompanhamento formatada (cabeçalho, cores alternadas, filtros, largura de coluna automática)
- Download em lote via arquivo ZIP
- Menu de navegação fixo no topo, com layout responsivo e centralizado
- Sem armazenamento de dados: cada envio é processado e descartado ao final

## Tecnologias utilizadas

- **Backend**: Python 3.11+ com Flask
- **Processamento de dados**: Pandas
- **Manipulação de documentos Word**: python-docx
- **Manipulação de planilhas Excel**: openpyxl
- **Frontend**: HTML5, CSS3, JavaScript (embutidos no próprio `app.py`, sem dependência de pastas `templates/` ou `static/`)
- **Fontes**: Google Fonts (Inter)
- **Ícones**: Font Awesome 6

## Pré-requisitos

- Python 3.11 ou superior
- Pip (gerenciador de pacotes Python)
- Git (opcional, para clonar o repositório)

## Instalação local

### 1. Clone o repositório

```bash
git clone https://github.com/anjeelo/appPleitos
cd appPleitos
```

### 2. Instale as dependências

```bash
pip install -r requirements.txt
```

### 3. Verifique os modelos

A pasta `modelos/` precisa conter os três arquivos abaixo antes de rodar o sistema:

```
modelos/
├── Modelo Carta Pleito Backbone Nacional.docx
├── Modelo Carta Pleito Backbone Regional.docx
└── Modelo Croqui.xlsx
```

Os dois primeiros já fazem parte do projeto. O `Modelo Croqui.xlsx` precisa usar exatamente esse nome, ou a constante `MODELO_CROQUI` no início do `app.py` deve ser ajustada para apontar para o arquivo correto.

Se as células de destino do modelo de croqui forem diferentes das já configuradas (`C42`, `H53`, `C53`, `S32`, `C51`, `B56`, `H31`, `L43`), ajuste a lista `CROQUI_MAPPINGS` em `app.py`.

### 4. Rode o servidor

```bash
python app.py
```

Acesse `http://localhost:5000` — a página inicial redireciona para a ferramenta de Carta Pleito, com o menu no topo dando acesso às duas ferramentas.

## Formato das planilhas

### Carta Pleito

As colunas devem seguir esta ordem, a partir da coluna A (a primeira linha pode ser cabeçalho):

| Coluna | Campo | Descrição |
|---|---|---|
| A | Tipo | `Nacional` ou `Regional` |
| B | Sequência | Número da TA |
| C | Data de Criação | Data da ocorrência |
| D | Município | Localidade da ocorrência |

### Criação de Croqui

As colunas devem seguir esta ordem, a partir da coluna A (a primeira linha pode ser cabeçalho):

| Coluna | Campo |
|---|---|
| A | OR |
| B | TA |
| C | Obra (também usada para nomear o arquivo gerado) |
| D | Local |
| E | Causa |
| F | Tratativa |
| G | Endereço |
| H | Exec |

Linhas totalmente vazias nas primeiras 8 colunas são ignoradas.

## Estrutura do projeto

```
appPleitos/
├── app.py               # aplicação Flask completa (rotas, lógica e interface)
├── modelos/              # modelos .docx e .xlsx usados na geração
├── requirements.txt
└── README.md
```

## Licença

Distribuído sob licença MIT.