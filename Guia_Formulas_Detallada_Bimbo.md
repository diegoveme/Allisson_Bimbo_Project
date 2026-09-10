# Preguntas y respuestas sobre las fórmulas del Sistema Bimbo

Les explico cómo armé cada fórmula del sistema, como una serie de preguntas que a mí me hubieran surgido al verlas por primera vez. Cada respuesta usa un ejemplo real de una celda del archivo, con números concretos cuando ayuda a seguir el cálculo paso a paso.

## Parte 1: cómo armé las fórmulas básicas

### ¿Qué significan los símbolos `$` que ven en algunas celdas y en otras no?

Marcan si una referencia se mueve o se queda fija al copiar la fórmula. `B5` (sin `$`) cambia solo si copio la fórmula a otra fila; `$B$3` o `Ventas!$N:$N` (con `$`) se quedan exactamente ahí sin importar a dónde la copie.

Para que entiendan por qué importa: si `=SUMIF($N$3:$N$170,1,$K$3:$K$170)` no llevara los `$` en el rango, al copiarla una fila hacia abajo el rango se movería a `N4:N171`, dejando fuera la primera fila y agregando una que no debería estar. Como ese rango debe ser siempre el mismo (las 168 ventas completas), lleva `$` en las dos partes. Regla que seguí en todo el archivo: lo que debe mirar siempre el mismo lugar lo dejo con `$`; lo que se adapta a cada fila (como el número de mes, o la celda que identifica la fila actual) lo dejo sin `$`.

### ¿Cómo hice para que Excel sume solo las filas que cumplen una condición?

Con `SUMAR.SI`. Ejemplo real, celda `Ventas!D179`:
```
=SUMIF($N$3:$N$170,1,$K$3:$K$170)
```
Se lee en tres pasos: reviso la columna Mes (N) de la fila 3 a la 170; en cada fila donde el Mes sea exactamente 1; sumo el valor de esa misma fila pero de la columna Total (K).

`CONTAR.SI` funciona igual pero con solo dos argumentos (dónde buscar, y la condición), porque no suma nada, solo cuenta. `PROMEDIO.SI` tiene los mismos tres argumentos que `SUMAR.SI`, pero saca el promedio en vez de la suma.

### ¿Cómo resolví las preguntas que cruzan dos datos a la vez, como "cuánto produjo este insumo en este mes"?

Con una columna que pega las dos condiciones en un solo texto, usando `&` (un operador para unir texto, no una función). Celda real `Ventas!P5` (Clave Vendedor-Mes):
```
=F5&"-"&N5
```
Si el Vendedor es "Erick Fernando Bautista Silva" y el Mes es 2, esta celda queda como el texto `Erick Fernando Bautista Silva-2`. Con esa combinación armada, un `SUMAR.SI` de una sola condición puede buscar exactamente esa combinación. Esta es la técnica detrás de la Comisión en Recursos Humanos:
```
=IF(D5="Vendedor",SUMIF(Ventas!$P:$P,M5,Ventas!$K:$K)*0.02,0)
```

Usé esta misma idea en varios lugares: Ventas → Recursos Humanos (Clave Vendedor-Mes), Ventas y Producción → Almacén de Producto Terminado (Clave Producto-Mes), Compras → Almacén de Materia Prima (Clave Insumo-Mes), y dentro de Finanzas (Clave Tipo-Mes).

### ¿Cómo hace una hoja para saber quién tiene el valor más alto de una tabla, sin traer el nombre a otra celda?

Con una comparación de cada fila contra el `MAX`/`MIN` de toda su tabla, usando `SI`. En vez de preguntar "¿quién es el mejor?", cada fila se pregunta a sí misma "¿soy yo el mejor?". Ejemplo real, tabla de vendedores en Ventas:
```
=IF(B200=MAX($B$200:$B$204),"Mejor vendedor",IF(B200=MIN($B$200:$B$204),"Menor desempeño",""))
```
A esta columna le puse "Marca" en todo el archivo. El límite de esta técnica es que la respuesta se queda EN esa misma tabla (dice "Mejor vendedor" junto al nombre correcto, pero no puede traer ese nombre solo hasta otra hoja). Para eso usé INDEX/MATCH, que explico después.

### ¿Por qué a veces meto un `SI` dentro de otro `SI`?

Para elegir entre 3 opciones en vez de 2. Ejemplo real, columna Estado en Almacén de Producto Terminado:
```
=IF(I3<=J3,"Reabastecer",IF(I3<=J3*1.5,"Stock bajo","Disponible"))
```
Con números: si el Stock mínimo es 40 y el Stock final es 30, la primera pregunta (30≤40) es verdadera, sale "Reabastecer" y ahí se detiene. Si el Stock final fuera 55, la primera (55≤40) es falsa, pasa a la segunda: ¿55≤60 (40×1.5)? Sí, sale "Stock bajo". Si fuera 200, ninguna se cumple, sale "Disponible" por descarte.

### Hay una fórmula con `SUMAR.SI` que trae una fecha en vez de sumar dinero. ¿Cómo funciona?

Como una fecha es solo un número dentro de Excel, y cada Folio existe una sola vez en Ventas, "sumar" la fecha de la única fila que coincide es, en la práctica, traerla. Celda real, Fecha de salida en Reparto:
```
=SUMIF(Ventas!$B:$B,A5,Ventas!$A:$A)+1
```
Este truco solo funciona con números; el Cliente, al ser texto, se captura directo como dato.

---

## Parte 2: INDEX/MATCH, la excepción que sí usé (y por qué)

### ¿Por qué el archivo evita INDEX/MATCH en todos lados menos en 7 celdas de Indicadores?

Porque el proyecto no pide INDEX/MATCH en su lista de funciones requeridas, así que evité usarlas en las 12 hojas del sistema. Pero en Indicadores, para las preguntas de "¿cuál es el producto más vendido?", "¿quién es el mejor vendedor?" y "¿qué semana tuvo la mayor venta?", se decidió sí usarlas, porque de otra forma la respuesta solo podía señalar "ve a tal hoja, tal tabla, tal columna", sin traer el nombre o el número real a la vista. El equipo prefirió ver el resultado directo. Esto es una excepción intencional, no un descuido, y hay que poder explicarla así en la defensa si preguntan por qué esas 7 celdas sí las usan y las otras 3,000+ fórmulas del archivo no.

### ¿Cómo funciona INDEX/MATCH, paso a paso?

Es un buscador de dos partes. Ejemplo real, celda `Indicadores!C9` (producto más vendido):
```
=INDEX(Ventas!$A$208:$A$252,MATCH(MAX(Ventas!$B$208:$B$252),Ventas!$B$208:$B$252,0))
```

**Paso 1: `MAX(Ventas!$B$208:$B$252)`**. Esto encuentra el número más alto de la columna de cantidades vendidas (la columna B de la tabla de productos en Ventas, que va de la fila 208 a la 252, una fila por cada uno de los 45 productos). Supongamos que ese número más alto es 2,936.

**Paso 2: `MATCH(2936, Ventas!$B$208:$B$252, 0)`**. `MATCH` no trae ningún dato, solo dice EN QUÉ POSICIÓN, dentro de ese rango, está el número 2,936. Si 2,936 está en la fila 250 (que es la posición 43 dentro del rango, porque el rango empieza en la fila 208: 250-208+1=43), `MATCH` regresa el número **43**. El tercer argumento, el `0`, le dice a Excel "quiero coincidencia exacta".

**Paso 3: `INDEX(Ventas!$A$208:$A$252, 43)`**. `INDEX` sí trae un dato: va a la posición 43 pero de la columna A (los NOMBRES de los productos, que está al lado de la columna B con las cantidades), y trae lo que hay ahí. En este caso, "Muffins Bimbo caja 4pz".

Juntando los tres pasos: el resultado final de la celda es "Muffins Bimbo caja 4pz", el nombre completo del producto, sin tener que ir a buscarlo a mano.

### ¿Por qué funciona aunque cambien los datos?

Porque `MATCH` no busca un número fijo escrito a mano, busca el resultado de `MAX` (que se recalcula solo). Si mañana otro producto vende más, `MAX` cambia, `MATCH` encuentra una posición distinta, e `INDEX` trae un nombre distinto, todo automático.

### ¿Dónde exactamente se usó esta técnica?

En 6 celdas de la hoja Indicadores (fórmulas completas, reales):
```
Producto más vendido:
=INDEX(Ventas!$A$208:$A$252,MATCH(MAX(Ventas!$B$208:$B$252),Ventas!$B$208:$B$252,0))

Producto menos vendido:
=INDEX(Ventas!$A$208:$A$252,MATCH(MIN(Ventas!$B$208:$B$252),Ventas!$B$208:$B$252,0))

Mejor vendedor:
=INDEX(Ventas!$A$200:$A$204,MATCH(MAX(Ventas!$B$200:$B$204),Ventas!$B$200:$B$204,0))

Menor desempeño:
=INDEX(Ventas!$A$200:$A$204,MATCH(MIN(Ventas!$B$200:$B$204),Ventas!$B$200:$B$204,0))

Semana con mayor venta:
=INDEX(Ventas!$A$185:$A$196,MATCH(MAX(Ventas!$B$185:$B$196),Ventas!$B$185:$B$196,0))

Semana con menor venta:
=INDEX(Ventas!$A$185:$A$196,MATCH(MIN(Ventas!$B$185:$B$196),Ventas!$B$185:$B$196,0))
```
Y una séptima, en la tabla de situaciones problemáticas, para traer el nombre del producto con menor stock final (mismo patrón, apuntando a la tabla de stock en Almacén de Producto Terminado en vez de a la tabla de productos en Ventas).

**Nota sobre la pregunta del mes:** "¿qué mes tuvo la mayor y la menor venta?" NO usa INDEX/MATCH, porque solo son 3 meses: ahí bastan dos `IF` anidados comparando los tres totales directamente, sin necesitar ningún buscador. INDEX/MATCH solo hizo falta para la semana (12 valores) y para nombres de texto (producto, vendedor), donde comparar a mano con `IF` ya no es práctico.

---

## Parte 3: cómo armé cada área

**Catálogo Producto Terminado y Catálogo Materia Prima**: sin fórmulas, son las tablas de referencia. La tasa de consumo en Materia Prima (columna E) es un número capturado que se usa después en Almacén de Materia Prima.

**Producción** (Aarón Sánchez e Italia Jiménez): registra cuánto se fabricó de cada producto cada mes. Su única fórmula es la Clave Producto-Mes.

**Almacén de Materia Prima** (Allisson Cabello): el Stock inicial de meses 2 y 3 es el Stock final del mes anterior; las Entradas usan `SUMAR.SI` contra la Clave de Compras; las Salidas multiplican la tasa de consumo del Catálogo por el total de piezas producidas ese mes (`SUMAR.SI` sobre Producción, sin distinguir producto, una simplificación necesaria para no armar una receta insumo-por-producto).

**Almacén de Producto Terminado** (Allisson Cabello): mismo patrón de Stock inicial; las Entradas usan `SUMAR.SI` contra la Clave de Producción (aquí sí es específico por producto); las Salidas contra la Clave de Ventas.

**Recursos Humanos** (Sergio Chávez): la Comisión combina `SI` con `SUMAR.SI` sobre la Clave Vendedor-Mes, igual en las 51 filas.

**Ventas** (Paola Millán y Paola Erazo): Precio, Categoría y Cantidad son datos capturados; Subtotal, IVA y Total son aritmética simple; las columnas O y P son las Claves.

**Reparto** (Diego Huerta): la Fecha de salida usa el truco de `SUMAR.SI` como buscador de un valor único.

**Compras** (Aarón Sánchez e Italia Jiménez): compra materia prima; la Clave combina `SI` con la técnica de pegar texto, poniendo "NA" cuando el pedido no ha llegado.

**Finanzas** (Renata Iglesias): usa la Clave Tipo-Mes para separar Ingresos y Egresos por mes con una sola condición de `SUMAR.SI`.

**Indicadores**: casi todo son referencias directas a celdas ya calculadas; las 5 excepciones con INDEX/MATCH ya se explicaron en la Parte 2.

## Parte 4: cómo desarmar una fórmula que no expliqué

1. **¿Tiene `SUMIF`, `COUNTIF` o `AVERAGEIF`?** Suma, conteo o promedio con una condición. Ver Parte 1.
2. **¿La condición es un texto con guion en medio, como "Erick-2"?** Viene de una columna Clave; busquen de dónde sale con `&`.
3. **¿Tiene `IF` comparando una celda contra `MAX($rango)`/`MIN($rango)`?** Es una columna Marca.
4. **¿Tiene `INDEX` junto con `MATCH`?** Es de las 5 excepciones que sí traen el nombre directo; ver Parte 2.
5. **¿Tiene un `IF` metido dentro de otro?** Decisión de 3 resultados o más.
6. **¿Qué partes tienen `$` y cuáles no?** Lo que tiene `$` no se mueve; lo que no tiene, se adapta a la fila.
7. **¿Menciona el nombre de otra hoja antes de `!`?** Esa celda cruza información de una hoja distinta.

Con esas 7 preguntas se puede desarmar cualquier fórmula del archivo.
