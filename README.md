# Vindex — Preenchimento de Planilha

Aplicação Streamlit para preencher automaticamente uma planilha Excel.

## Regra

- Coluna **J**: quantidade prevista
- Coluna **M**: início do preenchimento
- Cada unidade recebe `1` em uma coluna
- O preenchimento continua até atingir a quantidade prevista de cada linha

## Arquivos

- `app.py` — aplicação Streamlit
- `requirements.txt` — dependências

## Rodar localmente

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Publicar no Streamlit Community Cloud

1. Crie um repositório no GitHub.
2. Faça upload de `app.py` e `requirements.txt`.
3. No Streamlit Community Cloud, escolha o repositório.
4. Selecione `app.py` como arquivo principal.
5. Faça o deploy.
