# Relatório Semanal de Investimentos — Estratégia de Dividendos

Ferramenta Python para geração automática de relatório fundamentalista
das melhores ações dos setores de renda passiva (Bancos, Energia,
Saneamento, Seguros e Telecom) listadas no Ibovespa.

---

## Arquivos do projeto

| Arquivo | Descrição |
|---|---|
| `relatorio_semanal_besst.py` | Script principal — roda no Windows (VS Code) |
| `relatorio_besst_colab.ipynb` | Notebook — roda no Google Colab (celular/navegador) |
| `ranking_acoes_ibovespa.py` | Script de ranking standalone |
| `verificar_ambiente.py` | Diagnóstico do ambiente Python local |

---

## Opção 1 — Rodar no Windows (VS Code)

### Pré-requisitos

```bash
pip install python-docx yfinance requests pandas google-genai
```

### Configurar chave Gemini

```bash
setx GEMINI_API_KEY "sua-chave-aqui"
```

Feche e reabra o VS Code após o comando.

### Executar

```bash
python relatorio_semanal_besst.py
```

O relatório será salvo em `resultados/relatorios/`.

---

## Opção 2 — Rodar no Google Colab (celular ou navegador)

### Passo a passo

**1. Abrir o notebook**

Acesse [colab.research.google.com](https://colab.research.google.com),
clique em `Arquivo → Abrir notebook → GitHub` e cole a URL deste repositório.

Ou acesse diretamente pelo link do notebook no repositório e clique em
**"Open in Colab"**.

**2. Configurar a chave Gemini nos Secrets**

- Clique no ícone de cadeado na barra lateral esquerda
- Clique em `+ Adicionar novo secret`
- Nome: `GEMINI_API_KEY`
- Valor: sua chave do Google AI Studio
- Ative o acesso ao notebook

**3. Configurar a URL do GitHub na Célula 5**

Abra a Célula 5 e substitua a linha:

```python
GITHUB_RAW_URL = 'https://raw.githubusercontent.com/SEU_USUARIO/SEU_REPO/main/relatorio_semanal_besst.py'
```

Pela URL raw do seu repositório. Para obter essa URL:
- Abra o arquivo `relatorio_semanal_besst.py` no GitHub
- Clique em `Raw`
- Copie a URL da barra de endereços

**4. Executar**

`Ctrl+F9` ou `Runtime → Run all`

**5. Resultado**

- O arquivo `.docx` será salvo automaticamente em
  `Google Drive → relatorios_besst`
- O download iniciará automaticamente no navegador
- No celular: acesse o Google Drive para abrir o arquivo

---

## Como subir no GitHub

### Primeira vez

```bash
git init
git add relatorio_semanal_besst.py relatorio_besst_colab.ipynb README.md
git commit -m "Primeiro commit — relatório BESST"
git remote add origin https://github.com/SEU_USUARIO/SEU_REPOSITORIO.git
git push -u origin main
```

### Atualizações futuras

```bash
git add relatorio_semanal_besst.py
git commit -m "Descricao da alteracao"
git push
```

---

## Fontes de dados

| Dado | Fonte |
|---|---|
| Indicadores fundamentalistas | [Fundamentus.com.br](https://www.fundamentus.com.br) |
| Composição do Ibovespa | [B3](https://www.b3.com.br) |
| Histórico de dividendos | [Yahoo Finance (yfinance)](https://finance.yahoo.com) |
| Análise qualitativa | [Google Gemini API](https://aistudio.google.com) |

Todas as fontes são públicas e gratuitas.

---

## Aviso legal

Este projeto é de uso pessoal e educacional. Nenhuma informação
gerada constitui recomendação formal de investimento. O investidor
é o único responsável por suas decisões financeiras.
