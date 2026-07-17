# Safari Fotos

Aplicação web para importar planilhas de produtos, localizar imagens no catálogo da Distribuidora Safari e consultar o fabricante quando necessário.

## Rodar localmente

```bash
python -m venv .venv
. .venv/bin/activate
pip install -r requirements.txt
python app.py
```

Acesse `http://localhost:5000`.

## Publicar

O arquivo `render.yaml` permite publicar o projeto no Render usando **New > Blueprint** e conectando este repositório do GitHub.

## Fluxo

1. Importe um `.xlsx` com as colunas `Produto` e `Descrição do Produto`.
2. Clique em **Buscar fotos**.
3. Revise itens amarelos ou cole uma URL alternativa.
4. Baixe a planilha atualizada ou todas as fotos em ZIP.
5. Use **Gerar tabloide** para escolher as datas e a distribuição A4.

O tabloide pode ser gerado em três densidades: Compacto, Equilibrado e Destaque. O HTML baixado está pronto para abrir no Chrome e imprimir ou salvar em PDF A4.

O botão **Gerar vetor editável (SVG)** baixa um ZIP com uma página A4 por arquivo. Textos, preços, cores e formas permanecem editáveis no CorelDRAW, Illustrator ou Inkscape.
