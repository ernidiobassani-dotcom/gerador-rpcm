# Gerador de RPCM

Aplicação web para geração automática do Relatório de Prestação de Contas Mensal (RPCM) de contratos de credenciamento com o HMAPA (UG 167399).

## Como usar

1. Acesse o link do app
2. Faça upload do seu documento base (.docx ou .dotx)
3. Selecione o mês e o ano de referência
4. Clique em **Gerar Relatório**
5. Baixe o documento atualizado

O CNPJ é extraído automaticamente do documento. Os pagamentos são buscados diretamente no Portal da Transparência do Governo Federal.

## Como publicar no Streamlit Cloud (gratuito)

1. Crie uma conta em [streamlit.io](https://streamlit.io) com seu GitHub
2. Clique em **New app**
3. Selecione este repositório, branch `main`, arquivo `app.py`
4. Clique em **Deploy** — o link público fica disponível em segundos

## Dependências

Ver `requirements.txt`. Instalar localmente com:

```
pip install -r requirements.txt
streamlit run app.py
```
