# 📡 Coinfra — Automação e Dashboard de Infraestrutura O&M

Pipeline em Python para consolidar dados de infraestrutura de rede móvel, enriquecer com informações de DDD por município, manter um histórico versionado em banco de dados e visualizar tudo em um dashboard interativo.

## 🎯 Sobre o projeto

Esse projeto nasceu da necessidade de automatizar um processo manual de tratamento de planilhas de infraestrutura (O&M), reduzindo tempo de trabalho repetitivo e permitindo acompanhar a evolução dos indicadores ao longo do tempo — algo que uma planilha estática sozinha não permite.

O fluxo funciona em duas etapas:

1. **`Coinfra.py`** — lê a planilha principal de infraestrutura, filtra os registros relevantes (UF e contratada responsável), cruza os municípios com uma planilha auxiliar de DDDs, e grava o resultado tanto em um Excel local quanto em uma tabela histórica no banco de dados (cada execução é uma nova "foto" no tempo, nunca sobrescrevendo dados anteriores).
2. **`dash.py`** — dashboard em Streamlit que lê essa tabela histórica e permite explorar os dados com filtros, gráficos e uma busca dedicada por site.

## ✨ Funcionalidades

- ✅ Leitura e filtro automatizado de planilhas (UF, contratada responsável)
- ✅ Enriquecimento automático com DDD via planilha auxiliar de municípios
- ✅ Histórico versionado em banco de dados (nunca sobrescreve, sempre acumula)
- ✅ Configuração de credenciais via `.env` (nada de senha exposta no código)
- ✅ Dashboard interativo com:
  - Filtros por UF, contratada, município e período
  - Indicadores e gráficos (distribuição por DDD, top municípios, evolução no tempo)
  - Busca dedicada por Site Central, trazendo todo o histórico daquele site específico
  - Exportação dos dados filtrados em CSV

## 🛠️ Tecnologias

- **Python** — linguagem principal
- **pandas** — leitura e tratamento das planilhas
- **SQLAlchemy + PyMySQL** — conexão e escrita no banco de dados
- **python-dotenv** — gerenciamento seguro de credenciais
- **Streamlit** — dashboard interativo
- **Plotly** — gráficos
- **MariaDB** — armazenamento do histórico

## 📁 Estrutura do projeto

```
├── Coinfra.py           # script de automação (filtro + DDD + histórico no banco)
├── dash.py              # dashboard Streamlit
├── requirements.txt     # dependências do projeto
├── .env.example         # modelo de variáveis de ambiente (sem dados reais)
└── .gitignore           # garante que .env e planilhas não sejam versionados
```

## 🚀 Como usar

### 1. Clone o repositório

```bash
git clone <url-do-seu-repositorio>
cd coinfra
```

### 2. Instale as dependências

```bash
pip install -r requirements.txt
```

### 3. Configure as variáveis de ambiente

Copie o arquivo de exemplo e preencha com seus dados reais:

```bash
cp .env.example .env
```

Variáveis necessárias no `.env`:

| Variável            | Descrição                                              |
| ------------------- | ------------------------------------------------------ |
| `CAMINHO_PRINCIPAL` | Caminho da planilha principal de infraestrutura        |
| `CAMINHO_AUXILIAR`  | Caminho da planilha auxiliar com os DDDs por município |
| `CAMINHO_SAIDA`     | Nome do arquivo Excel gerado como saída local          |
| `DB_HOST`           | Endereço do servidor MariaDB                           |
| `DB_PORT`           | Porta do MariaDB (padrão: 3306)                        |
| `DB_USER`           | Usuário do banco                                       |
| `DB_SENHA`          | Senha do banco                                         |
| `DB_NOME`           | Nome do banco de dados                                 |
| `DB_TABELA`         | Tabela onde o histórico é gravado                      |

### 4. Rode a automação

```bash
python Coinfra.py
```

### 5. Rode o dashboard

```bash
streamlit run dash.py
```

Se estiver publicando no Streamlit Community Cloud, configure as mesmas variáveis em **Settings → Secrets** ao invés do `.env` (que não é enviado ao repositório por segurança).

## 🔒 Segurança

- Nenhuma credencial fica exposta no código — tudo é lido via variáveis de ambiente.
- O `.gitignore` impede que o `.env` real e as planilhas de dados sejam versionados.
- ⚠️ Como o banco de dados fica em uma rede interna corporativa, o dashboard **não deve ser publicado publicamente** sem antes validar com a área de segurança da informação da empresa.

_Projeto pessoal de automação e visualização de dados, desenvolvido para otimizar um processo antes feito manualmente em planilhas._
