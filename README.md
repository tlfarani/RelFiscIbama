# ⚖️ FiscFlow — IBAMA: Gerador de Relatórios de Fiscalização de Acidentes Ambientais

Aplicativo web desenvolvido em **Streamlit** para automação, filtragem e geração em lote de **Relatórios de Fiscalização** em formato Word (`.docx`), integrado diretamente com bases de dados no **SharePoint via Power Automate**.

---

## 🔑 1. Como Obter e Configurar a URL da Planilha (Power Automate)

A comunicação entre o aplicativo e a planilha no SharePoint é realizada por meio de um **Webhook (HTTP POST)** do Power Automate. **Não deve ser utilizado o link direto do SharePoint ou do Excel**, mas sim o endpoint gerado pelo fluxo.

### 📋 Passo a Passo para Obter a URL do Fluxo:

```
[ Power Automate ] ──► [ Nó: Quando uma requisição HTTP for recebida ] ──► [ Salvar ] ──► [ Copiar HTTP POST URL ]

```

1. **Acesse o Power Automate:**
* Entre no [Power Automate](https://www.google.com/search?q=https://make.powerautomate.com/) com a sua conta institucional.


2. **Abra o Fluxo do Projeto:**
* Localize e abra o fluxo responsável por ler os dados da planilha da força-tarefa (ou crie/duplique um fluxo para a nova força-tarefa).


3. **Configure o Nó da Planilha:**
* No bloco de leitura do Excel (*Listar linhas presentes em uma tabela*), garanta que os campos **Local**, **Biblioteca de Documentos**, **Arquivo** e **Tabela** estejam apontando para a nova planilha do SharePoint.


4. **Gere a URL de Disparo no Primeiro Nó:**
* Clique no primeiro bloco do fluxo: **"Quando uma requisição HTTP for recebida"** (*When an HTTP request is received*).
* Se o campo **HTTP POST URL** estiver em branco com a mensagem *"A URL será gerada após o salvamento"*, clique no botão **Salvar** (*Save*) no canto superior direito da tela.
* Após salvar, abra o bloco novamente: o campo **HTTP POST URL** exibirá um endereço longo iniciado por `[https://...api.powerplatform.com](https://...api.powerplatform.com)...`.
* Clique no ícone de **Copiar URL** (duas folhas de papel ao lado da caixa de texto).



---

### 🛡️ Configuração Segura da URL no Streamlit (Secrets)

Para proteger dados sensíveis da operação, **nunca insira a URL copiada no código Python nem envie para o repositório público do GitHub**.

#### A. Na Nuvem (Streamlit Cloud):

1. No painel do seu app no Streamlit Cloud, clique em **Manage App** (canto inferior direito).
2. Vá em **Settings** ➔ **Secrets**.
3. Cole a URL no formato TOML:
```toml
[sharepoint]
url_planilha = "https://SUA_URL_COPIADA_DO_POWER_AUTOMATE_AQUI"

```


4. Clique em **Save**.

#### B. Em Desenvolvimento Local (no seu computador):

1. Crie o arquivo `.streamlit/secrets.toml` e insira o mesmo bloco acima.
2. Para evitar o envio acidental desse arquivo para o GitHub, certifique-se de que a linha abaixo esteja presente no seu arquivo **`.gitignore`**:
```text
.streamlit/secrets.toml

```



---

## 📐 2. Arquitetura do Sistema

```
[ SharePoint / Excel ] 
        ▲
        │ (Power Automate - Webhook HTTP POST)
        ▼
[ FiscFlow (Streamlit) ] ──► [ Processamento / Regras de Negócio ] ──► [ Modelos .docx / Pacote .ZIP ]

```

* **Frontend / Interface:** Streamlit (Python)
* **Backend & Regras:** Python (`pandas`, `python-docx`, `zipfile`, `requests`)
* **Integração de Dados:** Power Automate Webhook (retorno assíncrono em JSON)
* **Saída:** Arquivos `.docx` individuais ou compactados em lote via `.zip`

---

## 🛠️ 3. Configuração de Arquivos do Repositório

### 📦 Dependências (`requirements.txt`)

```text
streamlit
pandas
python-docx
requests

```

### 🎨 Tema Visual Institucional (`.streamlit/config.toml`)

Para aplicar a paleta institucional (**Verde Musgo e Cinza Claro**) e forçar o modo claro na tabela:

```toml
[theme]
primaryColor = "#4E5D30"
backgroundColor = "#F8F9F9"
secondaryBackgroundColor = "#4E5D30"
textColor = "#000000"
base = "light"

```

---

## 📊 4. Mapeamento da Base de Dados (SharePoint)

O aplicativo lê as colunas da planilha do SharePoint e realiza o mapeamento interno automático:

| Coluna Original | Nome Interno | Descrição / Uso |
| --- | --- | --- |
| `ID` | `num_doc` | Identificador único do processo |
| `PROCESSO` | `processo_sei` | Número do processo no SEI |
| `SIEMA` | `siema` | Código SIEMA (Trata "Siema fora do ar") |
| `SITUACAO` | `situacao` | Situação do processo (Filtro padrão: "Autuar") |
| `LAUDO_SEI` | `laudo_sei` | Número do Laudo Técnico SEI |
| `Fiscal` | `fiscal` | Fiscal responsável pela análise |
| `DATA_ACIDENTE` | `data_acid` | Data da ocorrência do acidente |
| `PRODUTO` | `produto` | Sustância ou produto derramado |
| `CLASS_OL` | `class_ol` | Classificação (`Oleoso` ou `Não Oleoso`) |
| `CLASS_RISCO` | `class_risco` | Classe de risco da empresa (`A`, `B`, `C`, `D`, `E`) |
| `VOL` | `vol_char` | Volume derramado em m³ (aceita notação científica) |
| `Lat` / `Lon` | `lat` / `lon` | Coordenadas do acidente (Vão para o relatório `.docx`) |
| `Lat_Auto` / `Lon_Auto` | `lat_auto` / `lon_auto` | Coordenadas para lavratura (Exibidas na tela para o fiscal) |
| `MULTA_PREVISTA` | `multa_prevista` | Valor da multa prevista |
| `Grandeza` | `grandeza` | Potencial, Reduzida, Fraca, Moderada, Grave |
| `Nivel` / `Nivel_Pontos` | `nivel` / `nivel_pontos` | Nível de gravidade para enquadramento |

---

## 🧠 5. Premissas e Regras de Negócio

### 🎯 Seleção Automática de Modelos Word (`.docx`)

A escolha do modelo na pasta `modelos/` ocorre com base nos parâmetros do evento:

1. **Produtos Oleosos (`CLASS_OL = Oleoso`):**
* Seleciona `Rel_Fisc_Oleoso_[CLASSE].docx` (Classes A, B ou D).


2. **Produtos Não Oleosos (`CLASS_OL = Não Oleoso`):**
* **Classe A:** Se `Volume > 8 m³` ➔ `Rel_Fisc_Nao_Oleoso_Art_61_A.docx`, caso contrário ➔ `Rel_Fisc_Nao_Oleoso_Art_62_A.docx`.
* **Classe B:** Se `Volume > 200 m³` ➔ `Rel_Fisc_Nao_Oleoso_Art_61_B.docx`, caso contrário ➔ `Rel_Fisc_Nao_Oleoso_Art_62_B.docx`.
* **Classe D:** ➔ `Rel_Fisc_Nao_Oleoso_Art_62_D.docx`.



### 🔢 Formatação de Volumes e Datas

* **Volumes Decimais:** Notações científicas (ex: `4.05E-05`) são convertidas para decimais no padrão brasileiro sem arredondar para zero (ex: `0,0000405`), preservando até 7 casas decimais.
* **Datas Excel:** Números seriais de data do Excel são convertidos para o formato `DD/MM/AAAA`.

### 🟡 Destaque Automático em Amarelo (Revisão Manual)

Campos com pendência ou marcados entre colchetes (ex: `[ SIEMA FORA DO AR - EDITAR MANUAL ]`, `[ DATA - EDITAR MANUAL ]`) recebem **realce em amarelo automático** (`WD_COLOR_INDEX.YELLOW`) no documento Word gerado.

---

## 🔄 6. Procedimento para Iniciar uma Nova Força-Tarefa

Para alterar a base de dados do sistema para um novo ciclo de fiscalização:

1. Suba a nova planilha no **SharePoint**.
2. Abra o **Power Automate** e atualize o fluxo para apontar para a nova planilha.
3. Clique em **Salvar** e copie a nova **HTTP POST URL**.
4. Atualize o campo `url_planilha` no painel de **Secrets** do Streamlit Cloud.
5. Reinicie o aplicativo no Streamlit Cloud.
