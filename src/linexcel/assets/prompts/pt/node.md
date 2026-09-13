Documente cálculos do Excel para um leitor de negócio.
Respeite value_source: apenas engine significa recalculado; file é um cache, volatile um instantâneo e unknown uma origem desconhecida. Informe cached_agreement=differ e group_cached_agreement=differ; a concordância não prova exatidão em relação ao Excel. evaluated=false não fornece uma etapa calculada verificada. Vizinhos omitidos não estão ausentes do grafo. A extensão de um grupo é um retângulo envolvente que pode conter lacunas, não uma lista completa de membros. Não deduza significados de negócio apenas dos nomes das folhas, entradas ausentes por cálculo inverso nem valores de células não amostradas.
Para o nó fornecido, redija uma ficha curta em Markdown:
1. **Função** — uma frase sobre o que a fórmula calcula;
2. **Como** — a lógica, passo a passo, apoiando-se ESTRITAMENTE na
 decomposição fornecida (cite as subexpressões e os seus valores avaliados);
3. **Fontes** — de onde vêm os dados (precedentes, intervalos, nomes, VBA);
4. **Prova** — a fórmula exata e, se disponível, o valor calculado.
Regras absolutas: não invente dados; não afirme nada que não conste do dossiê;
se faltar informação, escreva «não determinado pela linhagem».
Tabelas: nunca escrevas tabelas Markdown com barras verticais. Quando uma tabela ajudar, coloca um marcador {{T1}}, {{T2}}… na sua própria linha e, depois do Markdown, adiciona um bloco ```json_tables — um array JSON de {"id": "T1", "caption": "…", "columns": ["…"], "rows": [["…", "…"]]}. Uma ferramenta determinista substitui cada marcador pela tabela final; JSON inválido descarta as tabelas, nunca o teu texto.
Responde com a ficha Markdown, seguida do bloco ```json_tables opcional; sem outros delimitadores.
