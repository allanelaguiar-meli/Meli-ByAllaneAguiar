# 📦 Separador de Romaneios

App para separar PDFs de romaneios por transportadora, anotar número e motorista,
e gerar um PDF unificado de Envios Extra pronto para impressão 2 páginas/folha.

---

## 📁 Estrutura

```
romaneios_app/
├── app.py              ← App principal
├── requirements.txt    ← Dependências
└── README.md
```

---

## 🚀 Deploy no Streamlit Cloud

1. Crie um **novo repositório** no GitHub: `romaneios-srj9`
2. Faça upload dos 3 arquivos (`app.py`, `requirements.txt`, `README.md`)
3. Acesse https://share.streamlit.io → **New app**
4. Aponte para `romaneios-srj9` / `main` / `app.py`
5. Clique em **Deploy**

> ⚠️ Este é um app **separado** do app de rotas (Driver Assignment).
> Crie um repositório diferente para cada um.

---

## 📖 Como usar

### 1 · Planilha de rotas
Faça upload do xlsx (aba **PLAN**) ou csv.

Colunas esperadas (por índice, começando do zero):
| Índice | Coluna | Conteúdo |
|--------|--------|----------|
| 3 (D)  | ROTA   | Código da rota (ex: AM_1) |
| 4 (E)  | QR/ROMANEIO | Número do romaneio |
| 15 (P) | TRANSPORTADORA | Nome da transportadora |
| 16 (Q) | MOTORISTA | Nome do motorista |

### 2 · PDFs
Envie os PDFs dos romaneios ou um ZIP com todos.
O app detecta automaticamente cada romaneio pela palavra "Roteiro" no texto.

### 3 · Opção de separação
- **MLPs + Envios Extra**: separa tudo
- **Só MLPs**: ignora Envios Extra
- **Só Envios Extra**: ignora MLPs

### 4 · Downloads gerados
- **ZIP por transportadora**: cada pasta contém os PDFs anotados com número do romaneio e nome do motorista
- **PDF Unificado Envios Extra**: todos os romaneios da Envios Extra em um único PDF

---

## 🖨️ Como imprimir o PDF Unificado

1. Abra o PDF no seu visualizador
2. Vá em **Imprimir**
3. Selecione:
   - **"Múltiplas páginas por folha: 2"**
   - **"Frente e verso — virar pela borda longa"**
4. Imprima
5. Cada romaneio tem número **par** de páginas — dobre e grampeie cada bloco individualmente

---

## ⚙️ Ajustar colunas da planilha

Se suas colunas estiverem em posições diferentes, edite no início do `app.py`:

```python
IDX_COL_ROTA    = 3   # Coluna D (índice 3)
IDX_COL_QR      = 4   # Coluna E (índice 4)
IDX_COL_TRANSP  = 15  # Coluna P (índice 15)
IDX_COL_DRIVER  = 16  # Coluna Q (índice 16)
```
