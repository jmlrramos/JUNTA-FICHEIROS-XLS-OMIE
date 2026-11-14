# Consolidador de ficheiros → Excel

Aplicação web (Streamlit). Fazes upload de vários ficheiros CSV, XLSX ou XLS e descarregas um único Excel. As linhas ficam ordenadas da data mais antiga para a mais recente.

## Testar localmente
```
pip install -r requirements.txt
streamlit run app.py
```

## Publicar grátis no Streamlit Community Cloud
1. Cria uma conta em github.com e um repositório novo (por exemplo `consolidador`).
2. Carrega `app.py` e `requirements.txt` para o repositório ("Add file" → "Upload files").
3. Entra em share.streamlit.io com a conta GitHub.
4. Clica em "Create app", escolhe o repositório, o branch `main` e o ficheiro `app.py`, e faz "Deploy".
5. Recebes um endereço do tipo `https://<nome>.streamlit.app`.
6. Para ser só para ti, vai a Settings → Sharing e limita o acesso a convidados (por email).

Notas
- O repositório pode ser privado.
- A aplicação adormece se ninguém a usar durante alguns dias. Quando voltas a abrir, demora cerca de 30 a 60 s a acordar.
- Os ficheiros são processados em memória e não ficam guardados. Mesmo assim, passam por um servidor de terceiros.

## Alternativa: Hugging Face Spaces
Cria um Space com o SDK "Streamlit" e visibilidade "Private". Depois carrega os mesmos dois ficheiros.
