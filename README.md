Aqui está o **`README.md`** completo e atualizado, contemplando a separação dos links no `secrets.toml`, a barra de atalhos rápidos, o painel de métricas e a prévia do texto de infração.

```markdown
# ⚖️ FiscFlow — IBAMA: Gestão de Fila e Automação de Relatórios

Aplicativo web desenvolvido em **Streamlit** para automação, filtragem, visualização prévia e geração em lote de **Relatórios de Fiscalização** em formato Word (`.docx`), integrado com bases de dados do **SharePoint via Power Automate**.

---

## 🔑 1. Configuração dos Segredos (`secrets.toml`)

O aplicativo utiliza **dois links distintos** para interagir com o SharePoint. Essa separação garante a segurança da API e evita erros de navegação.

### 📋 As Duas Chaves do Sistema:
1. **`url_planilha` (API Webhook - Backend):** URL do gatilho HTTP POST do Power Automate. Usada internamente pelo código Python para consultar e carregar os dados.
2. **`url_visualizacao` (Navegação - Frontend):** Link de visualização/edição direta do arquivo Excel no SharePoint. Usado pelo botão da barra lateral para abrir a planilha no navegador.

### 🛡️ Estrutura nos Secrets do Streamlit Cloud (*Settings > Secrets*):

```toml
[sharepoint]
# URL de API (obtida no Power Automate - método POST)
url_planilha = "https://SUA_URL_DO_POWER_AUTOMATE_AQUI"

# URL Web direta da planilha no SharePoint (para o botão da barra lateral)
url_visualizacao = "[https://ibamagov.sharepoint.com/:x:/r/teams/](https://ibamagov.sharepoint.com/:x:/r/teams/)..."

```

> ⚠️ **Segurança:** Nunca comite o arquivo `.streamlit/secrets.toml` no GitHub. Certifique-se de mantê-lo listado no seu arquivo `.gitignore`.

---

## 📐 2. Arquitetura do Sistema

```
                        ┌──► [ requests.post() ] ──► (Leitura dos dados JSON)
                        │
[ Secrets ] ────────────┼──► [ url_planilha ] ────► Power Automate / Webhook (POST)
                        │
                        └──► [ url_visualizacao ] ──► Botão "Planilha de Controle" (GET / Navegador)

```

* **Frontend / Interface:** Streamlit (Python)
* **Backend & Regras:** Python (`pandas`, `python-docx`, `zipfile`, `requests`)
* **Integração de Dados:** Power Automate Webhook (JSON via HTTP POST)
* **Saída:** Arquivos `.docx` individuais ou compactados em lote em pacotes `.zip`

---

## 🛠️ 3. Dependências e Tema Visual

### 📦 Dependências (`requirements.txt`)

```text
streamlit
pandas
python-docx
requests

```

### 🎨 Tema Visual Institucional (`.streamlit/config.toml`)

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

| Coluna Original | Nome Interno | Descrição / Uso |
| --- | --- | --- |
| `ID` | `num_doc` | Identificador único do processo |
| `PROCESSO` | `processo_sei` | Número do processo no SEI |
| `SIEMA` | `siema` | Código SIEMA (Destaca em amarelo se "fora do ar") |
| `SITUACAO` | `situacao` | Situação do processo (Filtro padrão: "Autuar") |
| `LAUDO_SEI` | `laudo_sei` | Número do Laudo Técnico SEI |
| `Fiscal` | `fiscal` | Fiscal responsável pela análise |
| `DATA_ACIDENTE` | `data_acid` | Data da ocorrência do acidente |
| `PRODUTO` | `produto` | Substância ou produto envolvido |
| `CLASS_OL` | `class_ol` | Classificação (`Oleoso` ou `Não Oleoso`) |
| `CLASS_RISCO` | `class_risco` | Classe de risco da empresa (`A`, `B`, `C`, `D`, `E`) |
| `VOL` | `vol_char` | Volume derramado em m³ (preserva até 7 casas decimais) |
| `Lat` / `Lon` | `lat` / `lon` | Coordenadas do acidente (Inseridas no relatório `.docx`) |
| `Lat_Auto` / `Lon_Auto` | `lat_auto` / `lon_auto` | Coordenadas da lavratura (Exibidas na tela para o fiscal) |
| `MULTA_PREVISTA` | `multa_prevista` | Valor da multa prevista |
| `Grandeza` | `grandeza` | Potencial, Reduzida, Fraca, Moderada, Grave |
| `Nivel` / `Nivel_Pontos` | `nivel` / `nivel_pontos` | Nível de gravidade para enquadramento |

---

## 🧠 5. Premissas e Regras de Negócio

1. **Atribuição Automática de Modelos Word (`.docx`):**
* **Produtos Oleosos:** `Rel_Fisc_Oleoso_[CLASSE].docx` (Classes A, B ou D).
* **Produtos Não Oleosos:**
* Classe A: Volume > 8 m³ ➔ `Art_61_A`, caso contrário ➔ `Art_62_A`.
* Classe B: Volume > 200 m³ ➔ `Art_61_B`, caso contrário ➔ `Art_62_B`.
* Classe D: ➔ `Art_62_D`.




2. **Formatação de Dados:**
* **Volumes Decimais:** Converte notação científica (ex: `4.05E-05`) para o padrão decimal brasileiro (ex: `0,0000405`).
* **Datas Excel:** Converte seriais de data do Excel para `DD/MM/AAAA`.


3. **Destaque de Edição Manual (Amarelo):**
* Inconsistências ou dados pendentes (ex: `[ SIEMA FORA DO AR - EDITAR MANUAL ]`) recebem realce amarelo automático (`WD_COLOR_INDEX.YELLOW`) no Word.


4. **Prévia da Descrição da Infração:**
* O aplicativo gera uma visualização prévia da infração na tela antes do download, simulando a redação final do auto e do relatório.



---

## 🚀 6. Principais Funcionalidades da Interface

* **Barra Lateral com Atalhos:** Acesso direto ao ProMar e link seguro para a Planilha de Controle no SharePoint.
* **Painel de Métricas:** Exibição em tempo real do total de processos no fluxo, prontos para autuação e status da conexão.
* **Filtros e Seleção em Massa:** Filtros por Situação, Fiscal e Laudo SEI com suporte a marcadores "Selecionar Todos" reativos.
* **Download Unificado (.ZIP):** Ao selecionar múltiplos processos, gera automaticamente um pacote `.zip` contendo todas as minutas preenchidas.

---

## 🔄 7. Procedimento para Nova Força-Tarefa

1. Suba a nova planilha no **SharePoint**.
2. Atualize o fluxo no **Power Automate** para apontar para o novo arquivo/tabela.
3. Copie o **HTTP POST URL** do primeiro nó do Power Automate e atualize a chave `url_planilha` nos Secrets.
4. Copie o link de compartilhamento da nova planilha no SharePoint e atualize a chave `url_visualizacao` nos Secrets.

```

```
