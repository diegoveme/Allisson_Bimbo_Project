# Manual del Sistema Bimbo (Excel)

Este manual recorre el archivo hoja por hoja: qué guarda cada columna, de dónde saca su información cuando es una fórmula, **paso a paso cómo se construye ese resultado** (no solo "qué función es", sino "cómo se llega ahí"), y cómo comprobar en el momento que esa conexión funciona. Cada bloque trae su propia comprobación justo después de explicarlo.

Como el equipo trabajó con 8 personas para 6 áreas, dos de ellas (Ventas y Compras/Producción) se hicieron en pareja, y las otras cuatro fueron individuales.

## Cómo viaja la información entre las 12 hojas

Bimbo fabrica su propio pan, no lo compra ya hecho. Por eso el sistema tiene **dos almacenes distintos** y una hoja de Producción en medio:

```
Compras (le compra materia prima a los proveedores: harina, azúcar, bolsas, etc.)
   → Almacén de Materia Prima (guarda los insumos que ya llegaron)
        → Producción (convierte materia prima en pan/producto terminado)
             → Almacén de Producto Terminado (guarda el pan ya hecho, listo para vender)
                  → Ventas (le vende ese producto terminado al cliente)
                       → Reparto (entrega los pedidos de Mayoreo/En línea)
Recursos Humanos calcula la comisión de cada vendedor con base en Ventas
Finanzas junta los ingresos de Ventas con los gastos de Compras, Nómina y Reparto
```

El sistema usa las funciones que pide el proyecto (SUMA, PROMEDIO, MAX, MIN, CONTAR, CONTARA, SI, CONTAR.SI, SUMAR.SI, PROMEDIO.SI), más **INDEX y MATCH** en 7 celdas específicas de Indicadores (se explica más abajo por qué esas sí y las demás no). Para todo lo demás, dos técnicas propias resuelven lo que el resto de las funciones no alcanza a cubrir:

- **Columna Clave**: pega dos condiciones en un solo texto (ej. Producto y Mes) para poder buscarlas juntas con un `SUMAR.SI` normal, que solo entiende una condición a la vez.
- **Columna Marca**: compara el valor de cada fila contra el `MAX`/`MIN` de su tabla usando `SI`, y esa fila se "autonombra" si es la ganadora o la perdedora.

---

## Catálogo Producto Terminado (45 productos, filas 3-47)

Solo marcas que **sí son de Bimbo actualmente**: Bimbo, Marinela, Barcel y Tía Rosa. Antes había productos Ricolino, pero Ricolino se vendió a Mondelez en 2022, así que ya no forma parte del portafolio y se sacaron del catálogo (se reemplazaron por productos reales de Marinela y Barcel: Suavicremas, Polvorones, Nito, Bombón, Chip's).

| Columna | Contenido |
|---|---|
| A - Código | Ej. `BM-001` |
| B - Producto | Nombre completo |
| C - Categoría | Pan de caja, Pan dulce, Galletas, etc. |
| D - Precio unitario | Precio de venta |
| E - Costo unitario | Referencia de costo |
| F - Stock mínimo | El número que compara la fórmula de Estado en Almacén de Producto Terminado |

No tiene ninguna fórmula: es la tabla de referencia.

## Catálogo Materia Prima (25 insumos, filas 3-27)

Aquí están los ingredientes (harina, azúcar, levadura, huevo, chocolate, etc.) **y también las bolsas de empaque específicas por producto** (Bolsa para Pan Blanco Grande, Bolsa para Donas, Bolsa para Conchas, etc.), igual que se maneja en una panificadora real donde cada tipo de pan lleva su propia bolsa.

| Columna | Contenido |
|---|---|
| A - Código | Ej. `MP-01` |
| B - Insumo | |
| C - Categoría | Granos y harinas, Empaque, etc. |
| D - Costo por unidad | |
| E - Consumo estimado por pieza producida | Ver explicación abajo |
| F - Stock mínimo | |

**La columna E es la pieza clave de todo el modelo de Materia Prima.** Es un número chiquito (por ejemplo `0.049`) que dice "cuánto de este insumo se gasta, en promedio, por cada pieza de CUALQUIER producto que se fabrica" (no es una receta específica por producto, es una tasa general simplificada). Se explica cómo se usa en la sección de Almacén de Materia Prima, más abajo.

---

## Recursos Humanos (17 empleados × 3 meses, filas 3-53)

| Columna | Contenido | Cómo se construye |
|---|---|---|
| A - ID empleado | | Dato capturado |
| B - Nombre | | Dato capturado |
| C - Área | | Dato capturado |
| D - Puesto | | Dato capturado |
| E - Turno | | Dato capturado |
| F - Mes | | Dato capturado |
| G - Días trabajados | | Dato capturado |
| H - Horas | | Dato capturado |
| I - Pago por hora | | Dato capturado |
| J - Bono | | Dato capturado |
| K - Comisión | `=IF(D5="Vendedor",SUMIF(Ventas!$P:$P,M5,Ventas!$K:$K)*0.02,0)` | Ver abajo, paso a paso |
| L - Pago total | `=H5*I5+J5+K5` | Horas × Pago por hora, más Bono, más Comisión |
| M - Clave Nombre-Mes | `=B5&"-"&F5` | Junta el Nombre y el Mes en un solo texto |

**Cómo se construye la Comisión, paso a paso:**
1. `D5="Vendedor"` pregunta si el puesto de esta fila es Vendedor.
2. Si la respuesta es NO, el resultado completo de la celda es `0` y ahí termina, sin calcular nada más.
3. Si la respuesta es SÍ, entra a `SUMIF(Ventas!$P:$P, M5, Ventas!$K:$K)`: busca en la columna P de Ventas (la Clave Vendedor-Mes) el mismo texto que tiene esta fila en su propia columna M (por ejemplo, `Erick Fernando Bautista Silva-2`), y suma la columna K (Total) de todas las filas de Ventas donde encuentre esa combinación exacta.
4. Ese resultado se multiplica por `0.02` (el 2% de comisión).

La misma fórmula, sin cambiar una letra, está en las 51 filas del archivo, sea la persona vendedora o no; es el `IF` del paso 1 el que decide qué hacer en cada caso.

**Panel de indicadores (desde la fila 56):**
- Nómina por mes: `=SUMIF($F$3:$F$53,1,$L$3:$L$53)` (filas 57-59, cambiando el 1 por 2 y 3)
- Nómina acumulada: `=D57+D58+D59` (fila 60)
- Pago máximo, mínimo, promedio: `MAX`, `MIN`, `AVERAGE` (filas 61-63)
- Empleados en Ventas / Reparto: `COUNTIF` dividido entre 3, porque cada empleado aparece 3 veces (una por mes) (fila 64-65)

**Cómo comprobarlo:** al subir el Total de una venta de cualquier vendedor, su Comisión del mes correspondiente debe subir sola, sin tocar nada en Recursos Humanos.

---

## Ventas (168 operaciones, filas 3-170)

| Columna | Contenido | Cómo se construye |
|---|---|---|
| A - Fecha | | Dato capturado |
| B - Folio | `F-2001` … `F-2168` | Dato capturado, es la llave que usa Reparto |
| C - Cliente | | Dato capturado |
| D - Producto | | Dato capturado, coincide con el Catálogo |
| E - Categoría | | Dato capturado |
| F - Vendedor | | Dato capturado |
| G - Cantidad | | Dato capturado |
| H - Precio unitario | | Dato capturado |
| I - Subtotal | `=G5*H5` | Cantidad × Precio unitario |
| J - IVA | `=I5*0.16` | 16% del Subtotal |
| K - Total | `=I5+J5` | Subtotal + IVA |
| L - Canal o sucursal | | Dato capturado |
| M - Semana | | Dato capturado |
| N - Mes | | Dato capturado |
| O - Clave Producto-Mes | `=D5&"-"&N5` | Junta Producto y Mes |
| P - Clave Vendedor-Mes | `=F5&"-"&N5` | Junta Vendedor y Mes |

**Panel de indicadores (desde la fila 173):**
- Ventas totales, promedio, máxima, mínima: `SUM`, `AVERAGE`, `MAX`, `MIN` sobre la columna K (filas 174-177)
- Número de operaciones: `=COUNTA(B3:B170)` (fila 178)
- Ventas por mes: `SUMIF` (filas 179-181)
- Tabla de semanas (filas 185-196): `SUMIF` por semana, más Marca con `IF` anidado
- Tabla de vendedores (filas 200-204): `SUMIF`, `COUNTIF`, `AVERAGEIF` por cada uno de los 5 vendedores
- Tabla de 45 productos (filas 208-252): `SUMIF` de cantidad vendida por producto

**Cómo comprobarlo:** al cambiar la Cantidad de cualquier fila por un número grande, el Total sube (columnas I, J, K), y en Almacén de Producto Terminado la Salida de ese producto en ese mes también debe subir.

---

## Producción (135 registros, filas 3-137)

Esta hoja es nueva respecto a un modelo de "solo comprar y revender": aquí se registra cuánto se **fabricó** de cada producto, cada mes. Cada uno de los 45 productos tiene una entrada de producción en cada uno de los 3 meses.

| Columna | Contenido | Cómo se construye |
|---|---|---|
| A - Producto | | Dato capturado |
| B - Cantidad producida | | Dato capturado |
| C - Fecha | | Dato capturado |
| D - Semana | | Dato capturado |
| E - Mes | | Dato capturado |
| F - Clave Producto-Mes | `=A3&"-"&E3` | Junta Producto y Mes |

**Panel de indicadores (desde la fila 140):**
- Piezas producidas por mes: `=SUMIF($E$3:$E$137,1,$B$3:$B$137)` (filas 141-143, cambiando el 1 por 2 y 3)

Esta hoja alimenta tanto a Almacén de Producto Terminado (como Entradas) como, de forma indirecta, a Almacén de Materia Prima (como base para calcular cuánto insumo se consumió).

---

## Almacén de Materia Prima (75 filas: 25 insumos × 3 meses)

Ordenado por insumo (3 filas seguidas por cada uno: mes 1, 2, 3).

| Columna | Contenido | Cómo se construye |
|---|---|---|
| A - Código | | Dato capturado |
| B - Insumo | | Dato capturado |
| C - Categoría | | Dato capturado |
| D - Mes | | Dato capturado |
| E - Stock inicial | Mes 1: dato capturado. Mes 2 y 3: `=I{fila anterior}` | Ver explicación abajo |
| F - Entradas | `=SUMIF(Compras!$K:$K,L3,Compras!$C:$C)` | Ver explicación abajo |
| G - Salidas | `='Catalogo Materia Prima'!$E$3*SUMIF(Produccion!$E:$E,D3,Produccion!$B:$B)` | Ver explicación abajo, es la fórmula más importante de esta hoja |
| H - Merma | | Dato capturado |
| I - Stock final | `=E3+F3-G3-H3` | Aritmética simple |
| J - Stock mínimo | | Dato capturado |
| K - Estado | `=IF(I3<=J3,"Reabastecer",IF(I3<=J3*1.5,"Stock bajo","Disponible"))` | IF anidado |
| L - Clave Insumo-Mes | `=B3&"-"&D3` | Junta Insumo y Mes |

**Cómo se construye la columna E (Stock inicial), paso a paso:** en el Mes 1 es un número capturado a mano (el inventario con el que arrancó la simulación). En el Mes 2 y el Mes 3, la fórmula es `=I{fila anterior}`, es decir, toma el Stock final de la fila justo arriba. Por ejemplo, la celda real `Almacén Materia Prima!E29` es exactamente `=I28`: el stock inicial de este mes es, literalmente, donde se quedó el inventario al cerrar el mes pasado.

**Cómo se construye la columna F (Entradas), paso a paso:**
1. `Compras!$K:$K` es la columna Clave de la hoja Compras, que solo tiene un texto real (como `Harina de trigo-1`) cuando el Estado de esa compra ya dice "Recibido"; si no, tiene el texto `"NA"`.
2. `SUMIF(Compras!$K:$K, L3, Compras!$C:$C)` busca, dentro de esa columna Clave, el mismo texto que tiene esta fila en su propia columna L (por ejemplo, también `Harina de trigo-1`), y si lo encuentra, suma la Cantidad solicitada (columna C de Compras) de esa fila.
3. Como las compras que siguen "En tránsito" o fueron "Canceladas" nunca llegan a tener una Clave real (tienen "NA"), nunca hacen match aquí, y por lo tanto nunca se cuentan como Entrada. Esto es automático, no hace falta una segunda condición aparte para filtrar por Estado.

**Cómo se construye la columna G (Salidas), paso a paso, la fórmula más importante de esta hoja:**
```
='Catalogo Materia Prima'!$E$3*SUMIF(Produccion!$E:$E,D3,Produccion!$B:$B)
```
1. `SUMIF(Produccion!$E:$E, D3, Produccion!$B:$B)` busca en la hoja Producción todas las filas donde el Mes (columna E de Producción) sea igual al Mes de esta fila (columna D), y suma la Cantidad producida (columna B) de esas filas. Esto da **el total de piezas de TODOS los productos que se fabricaron ese mes**, sin importar cuál producto específico fue.
2. `'Catalogo Materia Prima'!$E$3` es la tasa de consumo de este insumo específico (por ejemplo, `0.049` para la Harina de trigo), que está en el Catálogo de Materia Prima.
3. Se multiplican los dos: (tasa de consumo por pieza) × (piezas totales producidas ese mes) = cuánto de este insumo se gastó ese mes.

**Por qué se hizo así y no con una receta específica por producto:** una receta real diría "el Pan Blanco Grande usa 0.3kg de harina, 0 de chocolate" y "el Gansito usa 0.1kg de harina, 0.05kg de chocolate", es decir, una tabla de insumo × producto distinta para cada combinación. Eso son 25 insumos × 45 productos = 1,125 combinaciones posibles, demasiado para el alcance de este proyecto. La simplificación que se usó fue: cada insumo tiene una sola tasa de consumo "por pieza producida en general", sin importar cuál producto sea. No es una receta real, es una aproximación razonable para poder calcular el consumo con las funciones permitidas.

**Panel de indicadores (desde la fila 80):**
- Insumos que requieren reabastecimiento / en Stock bajo: `COUNTIF` (filas 81-82)
- Merma total: `SUM` (fila 83)
- Promedio de existencias: `AVERAGE` (fila 84)

**Cómo comprobarlo:** al cambiar en Compras el Estado de una fila de "En tránsito" a "Recibido", su columna Clave deja de decir "NA" y muestra el texto real; en Almacén de Materia Prima, ese insumo y ese mes deben ver subir su columna Entradas.

---

## Almacén de Producto Terminado (135 filas: 45 productos × 3 meses)

Ordenado por producto (3 filas seguidas por cada uno).

| Columna | Contenido | Cómo se construye |
|---|---|---|
| A - Código | | Dato capturado |
| B - Producto | | Dato capturado |
| C - Categoría | | Dato capturado |
| D - Mes | | Dato capturado |
| E - Stock inicial | Mes 1: dato capturado. Mes 2 y 3: `=I{fila anterior}` (celda real `E4` es `=I3`) | Igual que en Materia Prima |
| F - Entradas | `=SUMIF(Produccion!$F:$F,L3,Produccion!$B:$B)` | Ver explicación abajo |
| G - Salidas | `=SUMIF(Ventas!$O:$O,L3,Ventas!$G:$G)` | Busca en Ventas la Clave Producto-Mes y suma la Cantidad vendida |
| H - Merma | | Dato capturado |
| I - Stock final | `=E3+F3-G3-H3` | Aritmética simple |
| J - Stock mínimo | | Dato capturado |
| K - Estado | `=IF(I3<=J3,"Reabastecer",IF(I3<=J3*1.5,"Stock bajo","Disponible"))` | IF anidado |
| L - Clave Producto-Mes | `=B3&"-"&D3` | Junta Producto y Mes |

**Cómo se construye la columna F (Entradas):** a diferencia de Materia Prima (que usa una tasa general), aquí sí se sabe exactamente cuánto se produjo de CADA producto, porque la hoja Producción ya lo registra por producto. `SUMIF(Producción!$F:$F, L3, Producción!$B:$B)` busca la Clave Producto-Mes exacta de esta fila dentro de Producción, y suma la Cantidad producida de esas filas.

**Este es el punto donde antes había un problema real:** en una versión anterior del archivo, las Compras solo alcanzaban a cubrir 8 de 45 productos al azar cada mes, así que 121 de 135 filas tenían Entradas en cero. Ahora, como la hoja Producción tiene una entrada garantizada para cada uno de los 45 productos en cada uno de los 3 meses (con la única excepción a propósito de Pan Blanco Grande, que se deja corto como caso de estudio), ya no queda ninguna fila con Entradas en cero por accidente.

**Panel de indicadores (desde la fila 140):**
- Reabastecer / Stock bajo: `COUNTIF` (filas 141-142)
- Merma total: `SUM` (fila 143)
- Promedio de existencias: `AVERAGE` (fila 144)
- Tabla de stock final del Mes 3 por producto (filas 147-191): trae el Stock final de cada producto (ej. `=I5` para el primero), más Marca

**Caso de estudio para la defensa:** "Pan Blanco Grande Bimbo" queda con Stock final negativo los 3 meses (se agota más rápido de lo que se produce). El sistema lo marca solo como "Reabastecer" en la columna Estado.

**Cómo comprobarlo:** al subir la Cantidad producida de cualquier fila en Producción, en Almacén de Producto Terminado la columna Entradas de ese producto y ese mes debe subir.

---

## Reparto (79 pedidos, filas 3-81)

| Columna | Contenido | Cómo se construye |
|---|---|---|
| A - Folio | | Dato capturado |
| B - Cliente | | Dato capturado |
| C - Zona | | Dato capturado |
| D - Repartidor | | Dato capturado |
| E - Vehículo | | Dato capturado |
| F - Fecha de salida | `=SUMIF(Ventas!$B:$B,A3,Ventas!$A:$A)+1` | Ver explicación abajo |
| G - Días de entrega estimados | | Dato capturado |
| H - Fecha de entrega | `=IF(I3="Entregado",F3+G3,"")` | Solo si ya se entregó |
| I - Estado | | Dato capturado |
| J - Costo de reparto | | Dato capturado |
| K - Semana / L - Mes | | Copiados de la venta original |

**Cómo se construye la Fecha de salida, paso a paso, un uso poco común de SUMAR.SI:**
1. Cada Folio existe una sola vez en toda la hoja Ventas.
2. `SUMIF(Ventas!$B:$B, A3, Ventas!$A:$A)` busca ese Folio único en la columna de Folios de Ventas, y "suma" la Fecha (columna A de Ventas) de la única fila que coincide.
3. Como solo hay una fila que coincide, "sumar" un solo número es lo mismo que traerlo (una fecha, para Excel, es solo un número de serie por dentro).
4. Se le suma 1 día al final.

Este truco solo funciona con números (fechas incluidas); no sirve para traer texto, por eso el Cliente se captura directo como dato en vez de con una fórmula.

**Panel de indicadores (desde la fila 84):**
- Conteo por Estado: `COUNTIF` ×4 (filas 85-88)
- Pendientes o retrasados: `=D86+D87` (fila 89)
- Costo total y promedio: `SUM`, `AVERAGE` (filas 90-91)
- Tabla de repartidores y de zonas: `COUNTIF`/`AVERAGEIF` más Marca

**Cómo comprobarlo:** al cambiar en Ventas la Fecha de una venta cuyo Folio también exista en Reparto, la Fecha de salida de ese pedido debe moverse sola (con 1 día de diferencia).

---

## Compras (75 solicitudes, filas 3-77)

Ahora compra **materia prima**, no producto terminado ya hecho (eso no tendría sentido para una empresa que fabrica su propio pan).

| Columna | Contenido | Cómo se construye |
|---|---|---|
| A - Proveedor | | Dato capturado |
| B - Insumo | Debe existir en el Catálogo de Materia Prima | Dato capturado |
| C - Cantidad solicitada | | Dato capturado |
| D - Costo unitario | | Dato capturado |
| E - Total | `=C3*D3` | Cantidad × Costo unitario |
| F - Fecha de solicitud | | Dato capturado |
| G - Días de recepción estimados | | Dato capturado |
| H - Fecha de recepción | `=IF(I3="Recibido",F3+G3,"")` | Solo si ya llegó |
| I - Estado | | Dato capturado |
| J - Mes | | Dato capturado |
| K - Clave (si ya llegó) | `=IF(I3="Recibido",B3&"-"&J3,"NA")` | Ver explicación abajo |

**Cómo se construye la columna K, paso a paso:** primero `IF(I3="Recibido", ...)` pregunta si el Estado ya es "Recibido". Si SÍ, arma la Clave normal pegando el Insumo y el Mes (`B3&"-"&J3`). Si NO (sigue "En tránsito" o fue "Cancelado"), la Clave se reemplaza por el texto fijo `"NA"`, que nunca va a coincidir con ninguna combinación real de insumo-mes en Almacén de Materia Prima. Así, el `SUMIF` de Entradas en Almacén de Materia Prima excluye automáticamente lo que no ha llegado, sin necesitar una segunda condición.

**Panel de indicadores (desde la fila 80):**
- Gasto total: `SUM` (fila 81)
- Solicitudes recibidas/en tránsito/canceladas: `COUNTIF` (filas 82-84)
- Costo promedio: `AVERAGE` (fila 85)
- Tabla de proveedores: `SUMIF` más Marca (filas 88-97)

**Cómo comprobarlo:** al cambiar el Estado de cualquier fila a "Cancelado", su Clave cambia a "NA", y esa cantidad deja de sumarse en las Entradas de Almacén de Materia Prima.

---

## Finanzas (24 movimientos, filas 3-26)

| Columna | Contenido |
|---|---|
| A - Fecha | |
| B - Concepto | |
| C - Tipo de movimiento | Ingreso o Egreso |
| D - Área | |
| E - Monto | |
| F - Mes | |
| G - Clave Tipo-Mes | `=C3&"-"&F3` |

Los primeros 12 renglones jalan de otras hojas: `=SUMIF(Ventas!$N:$N,1,Ventas!$K:$K)` para ingresos, y patrones equivalentes para nómina, compras y reparto. Los últimos 12 (renta, servicios, mantenimiento de hornos, publicidad) son datos capturados a mano.

**Panel de indicadores (desde la fila 29):** para separar Ingresos y Egresos por mes hacen falta dos condiciones (Tipo Y Mes), así que se usa la Clave Tipo-Mes:
```
=SUMIF($G$3:$G$26,"Ingreso-1",$E$3:$E$26)
```
Se lee: busca en la columna Clave el texto exacto "Ingreso-1" (que ya combina las dos condiciones), y suma el Monto de esas filas.

- Ingresos y Egresos por mes: filas 31-32
- Resultado operativo: `=B31-B32` (fila 33)
- Acumulados de 3 meses: `SUM` (filas 35-37)
- Tabla de conceptos de egreso: `SUMIF` por concepto, más Marca (filas 40-46)
- Positivo/Negativo por mes: `IF` simple (filas 49-51)

**Dato para la defensa:** el margen entre ingresos y egresos es corto pero positivo en los 3 meses, y la nómina es, por mucho, el gasto de mayor peso del sistema (columna Marca en la fila 40).

---

## Indicadores

**Para las preguntas numéricas**, cada celda apunta directo a un resultado ya calculado (ej. `='Recursos Humanos'!D60` para la nómina acumulada).

**¿Qué mes tuvo la mayor y la menor venta?** (celda `Indicadores!C9`), fórmula real:
```
="Mayor: Mes "&IF(Ventas!D179>=Ventas!D180,IF(Ventas!D179>=Ventas!D181,1,3),IF(Ventas!D180>=Ventas!D181,2,3))&"   /   Menor: Mes "&IF(Ventas!D179<=Ventas!D180,IF(Ventas!D179<=Ventas!D181,1,3),IF(Ventas!D180<=Ventas!D181,2,3))
```
Solo son 3 meses, así que dos `IF` anidados alcanzan para comparar los tres totales directamente, sin necesitar Marca ni INDEX/MATCH.

**¿Qué semana tuvo la mayor y la menor venta?** (celda `Indicadores!C10-C11`): aquí sí son 12 semanas, y se usó INDEX/MATCH para traer el número de semana directo:
```
Semana con mayor venta: =INDEX(Ventas!$A$185:$A$196,MATCH(MAX(Ventas!$B$185:$B$196),Ventas!$B$185:$B$196,0))
Semana con menor venta: =INDEX(Ventas!$A$185:$A$196,MATCH(MIN(Ventas!$B$185:$B$196),Ventas!$B$185:$B$196,0))
```

**Para las preguntas de "quién/cuál" (producto más/menos vendido, vendedor mejor/peor), también se usó INDEX y MATCH**, a propósito, para poder mostrar el nombre real directamente en la celda en vez de solo indicar en qué otra hoja buscarlo. Esto es una excepción intencional: el resto del archivo completo evita INDEX/MATCH porque el proyecto no las pide en su lista de funciones, pero en estas 7 celdas específicas (producto más/menos vendido, vendedor mejor/peor, semana mayor/menor, y el producto con menor stock en la situación problemática) se decidió usarlas porque el equipo quería ver el resultado final directo, no un señalamiento a otra tabla. La explicación completa de cómo funciona INDEX/MATCH, paso a paso, está en la guía a fondo.

Las 3 situaciones problemáticas:
1. **Pedido mayor que el inventario disponible** (evidencia: INDEX/MATCH trae el nombre del producto con menor stock)
2. **Vendedor con bajo desempeño** (evidencia: INDEX/MATCH trae el nombre del vendedor con menor desempeño)
3. **Gastos que reducen el resultado operativo** (evidencia: columna Marca en la tabla de conceptos de egreso de Finanzas)

---

## Antes de la defensa

Cada quien debería poder, sin ayuda: abrir su hoja, señalar una fórmula al azar y explicar paso a paso de dónde sale cada parte del resultado (no solo decir "es un SUMAR.SI", sino explicar qué busca, dónde lo busca, y qué suma); hacer el cambio de "Cómo comprobarlo" de su sección; y explicar con sus palabras el flujo completo de materia prima → producción → producto terminado → venta.
