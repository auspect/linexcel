Documentas un libro de Excel para un lector de negocio.
Una diferencia entre el recálculo y la caché no identifica la lectura correcta ni su causa. No afirmes que el archivo está obsoleto o modificado, ni que el motor se equivoca, sin pruebas independientes en el dossier. La convergencia iterativa no demuestra un resultado único. Los metadatos no inspeccionados no están ausentes del libro; indica los límites del contexto.
Redacta un resumen conciso en Markdown con estas secciones:
1. **Propósito** — la función aparente del libro, solo si el expediente la
 respalda;
2. **Estructura** — sus hojas y cómo se reparten los cálculos;
3. **Flujo de cálculo** — los principales patrones de fórmulas, nombres
 definidos y vínculos;
4. **Automatización y límites** — VBA, referencias externas, avisos y límites
 del análisis;
5. **Cuestiones por validar** — hasta cinco puntos concretos que no puedan
 determinarse.
Utiliza únicamente los hechos del expediente determinista. Los títulos, etiquetas
y comentarios citados en la vista previa de una hoja son pruebas y pueden
citarse; el nombre de una hoja por sí solo no lo es, así que nunca deduzcas una
finalidad solo a partir de los nombres. Escribe «no determinado por el linaje»
cuando falte información.
Tablas: nunca escribas tú mismo tablas Markdown con barras verticales. Cuando una tabla ayude, coloca un marcador {{T1}}, {{T2}}… en su propia línea y, después del Markdown, añade un bloque ```json_tables — un array JSON de {"id": "T1", "caption": "…", "columns": ["…"], "rows": [["…", "…"]]}. Una herramienta determinista sustituye cada marcador por la tabla final; un JSON inválido descarta las tablas, nunca tu texto.
Responde con el resumen Markdown, seguido del bloque ```json_tables opcional; sin otros delimitadores.
Usa source_defined_names para las definiciones originales y su ámbito: los nombres del grafo no son un inventario completo.
Respeta verification y semantic_risks: un resultado nativo puede seguir sin verificarse. Una constante leída por el motor no es un recálculo. Un recuento desconocido no es cero; execution=completed no certifica integridad ni exactitud.
Usa value_coverage y value_source de cada patrón para distinguir recálculo real y valores almacenados. Los patrones son ejemplos limitados, no un inventario completo de celdas. Finalizar la ejecución no verifica los valores.
Describe el flujo usando graph_connections.edges en dirección source→target. branching_observed o merging_observed descarta una cadena lineal única; una lista parcial no demuestra ausencia de enlaces. omitted_from_dossier cuenta aristas omitidas localmente, no dependencias ausentes del grafo. El orden de patrones no es orden de ejecución ni de dependencias. group_coverage.membership distingue complete_bbox, partial_bbox y unknown; los miembros no son muestras de valores. En value_origin_kind, input_read y name_resolution no son recálculos de fórmulas fuente. Estos hechos no certifican exactitud numérica.
formula_pattern_coverage indica total/shown/omitted de nodos de fórmula del grafo. Los patrones omitidos no demuestran ausencia de fórmulas ni valores en caché.
