Documente uma pasta de trabalho do Excel para um leitor de negócio.
Redija uma síntese concisa em Markdown com estas secções:
1. **Finalidade** — o papel aparente da pasta, apenas se o dossiê o confirmar;
2. **Estrutura** — as suas folhas e a distribuição dos cálculos;
3. **Fluxo de cálculo** — os principais padrões de fórmulas, nomes definidos e
 ligações;
4. **Automação e limites** — VBA, referências externas, avisos e limites da
 análise;
5. **Questões a validar** — no máximo cinco pontos concretos indetermináveis.
Use apenas as informações do dossiê determinista. Títulos, rótulos e comentários
citados na pré-visualização de uma folha são provas e podem ser citados; o nome
de uma folha por si só não é, por isso nunca deduza uma finalidade apenas a
partir dos nomes. Escreva «não determinado pela linhagem» quando faltar
informação.
Tabelas: nunca escrevas tabelas Markdown com barras verticais. Quando uma tabela ajudar, coloca um marcador {{T1}}, {{T2}}… na sua própria linha e, depois do Markdown, adiciona um bloco ```json_tables — um array JSON de {"id": "T1", "caption": "…", "columns": ["…"], "rows": [["…", "…"]]}. Uma ferramenta determinista substitui cada marcador pela tabela final; JSON inválido descarta as tabelas, nunca o teu texto.
Responde com a síntese Markdown, seguida do bloco ```json_tables opcional; sem outros delimitadores.
