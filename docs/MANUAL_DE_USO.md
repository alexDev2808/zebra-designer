# Manual de uso — Etiquetas Zebra

Guía paso a paso para diseñar e imprimir etiquetas en Zebra ZT411 (300 / 600 dpi) desde Excel.

## Contenido

1. [Abrir el programa](#1-abrir-el-programa)
2. [La pantalla principal](#2-la-pantalla-principal)
3. [Preparar el Excel](#3-preparar-el-excel)
4. [Diseñar la etiqueta](#4-diseñar-la-etiqueta)
5. [Códigos de barras](#5-códigos-de-barras)
6. [Tamaño y calidad de la etiqueta](#6-tamaño-y-calidad-de-la-etiqueta)
7. [Seleccionar la impresora](#7-seleccionar-la-impresora)
8. [Imprimir](#8-imprimir)
9. [Plantillas](#9-plantillas)
10. [Comprobación de impresora y drivers](#10-comprobación-de-impresora-y-drivers)
11. [Solución de problemas](#11-solución-de-problemas)

---

## 1. Abrir el programa

Doble clic en **`run.bat`**. La primera vez tarda unos segundos porque instala lo necesario.

> No es necesario activar el entorno virtual. Desde terminal también puede usar
> `.\.venv\Scripts\python.exe app.py`.

## 2. La pantalla principal

| Zona | Qué contiene |
|---|---|
| **Encabezado** (azul marino) | Selector de impresora y resolución activa (300 / 600 dpi) |
| **Barra de herramientas** | *Datos* (Excel y hoja), *Plantilla* (nueva, abrir, guardar), *ZPL* (ver, exportar) y botones de impresión |
| **Panel izquierdo** | Pestañas **Elementos**, **Etiqueta** e **Impresora** |
| **Vista previa** | La etiqueta con los datos del registro actual, con reglas en mm |
| **Datos** | Tabla con las filas del Excel, búsqueda y selección |
| **Barra de estado** | Mensajes y resumen: registros, etiquetas a imprimir, dpi e impresora |

## 3. Preparar el Excel

- Formatos admitidos: `.xlsx` y `.xlsm` (si tiene `.xls`, guárdelo como `.xlsx`).
- **La fila 1 son los encabezados** (nombres de columna). Cada fila siguiente es un registro.
- Las filas completamente vacías se ignoran.
- Se respeta el formato de número de la celda: un precio con formato `0.00` sale como `3.50`.
- Las fechas se muestran como `dd/mm/aaaa`.
- Para códigos largos (EAN, etc.) se recomienda que la columna tenga formato **Texto**,
  así Excel no los convierte a notación científica ni elimina ceros a la izquierda.

Ejemplo:

| Codigo | Descripcion | Precio | Lote | Cantidad |
|---|---|---|---|---|
| 7501234567890 | Tornillo hexagonal 1/4 x 2" | 3.50 | L-2401 | 2 |

Pasos: **Abrir Excel…** → elegir archivo → elegir la **Hoja** en la lista.

Use el cuadro **Buscar** para filtrar la tabla. Al hacer clic en una fila, la vista previa
muestra ese registro; también puede recorrerlos con ◀ ▶.

## 4. Diseñar la etiqueta

En la pestaña **Elementos**:

- **+ Texto** / **+ Código**: agregan un elemento.
- **Duplicar**, **Eliminar**, **↑ ↓** (orden de dibujo).

### Contenido

El contenido admite marcadores con el nombre de la columna entre llaves:

| Contenido | Resultado |
|---|---|
| `{Descripcion}` | Tornillo hexagonal 1/4 x 2" |
| `Lote: {Lote}` | Lote: L-2401 |
| `{Codigo}-{Lote}` | 7501234567890-L-2401 |

Puede escribir el marcador o elegir la columna y pulsar **Insertar columna**.
Si un marcador no coincide con ninguna columna, se ve tal cual (`{Columna}`) en la vista
previa y la barra de estado avisa al cargar el Excel.

### Mover y redimensionar

- **Arrastrar** un elemento en la vista previa lo mueve.
- **Flechas del teclado**: mueven 0.5 mm (con **Shift**, 0.1 mm). Haga clic en la vista previa primero.
- **Manijas** del elemento seleccionado:
  - Derecha → ancho (en códigos activa *ancho fijo*; en textos define el ancho del bloque).
  - Abajo → alto del código de barras.
- También puede escribir la posición exacta **X / Y** en mm.

### Propiedades del texto

| Propiedad | Uso |
|---|---|
| Tamaño letra (mm) | Altura de los caracteres |
| Ancho bloque (mm) | `0` = una sola línea. Con valor, el texto se ajusta en varias líneas |
| Máx. líneas | Líneas permitidas dentro del bloque |
| Alineación | Izquierda, centro o derecha (dentro del bloque) |
| Rotación | 0°, 90°, 180°, 270° |

## 5. Códigos de barras

| Tipo | Uso típico | Datos |
|---|---|---|
| Code 128 | General, alfanumérico | Cualquier texto |
| Code 39 | Industrial | Mayúsculas, números y `- . $ / + %` |
| EAN-13 | Productos de venta | 12 o 13 dígitos (el verificador lo calcula la impresora) |
| UPC-A | Productos (EE. UU.) | 11 o 12 dígitos |
| QR | URLs, textos largos | Cualquier texto |
| DataMatrix | Piezas pequeñas | Cualquier texto (vista previa aproximada) |

### Alto

**Alto (mm)** define la altura de las barras (no aplica a QR ni DataMatrix, que son cuadrados).

### Ancho del código

- **Por grosor de barra**: define la barra más delgada (módulo) en mm con el campo o el
  deslizador. Recomendado 0.25–0.5 mm. El ancho total depende de la longitud de los datos.
- **Ancho fijo**: escriba el ancho deseado en mm. Para cada registro se usa el grosor de
  barra más grande que cabe en esa medida, así códigos de distinta longitud no se salen.

El recuadro azul muestra el **ancho real** con el registro actual a 300 y 600 dpi. Puede ser
un poco menor que el pedido porque cada barra mide un número entero de puntos de la
impresora (a 300 dpi, 1 punto = 0.085 mm).

**Mostrar texto legible debajo** imprime los dígitos bajo las barras.

## 6. Tamaño y calidad de la etiqueta

Pestaña **Etiqueta**:

| Campo | Descripción |
|---|---|
| Predefinido | Tamaños comunes (4×6", 4×2", 100×50 mm, …) |
| Ancho / Alto (mm) | Medidas de la etiqueta física |
| Resolución | 300 o 600 dpi (se ajusta sola con la impresora) |
| Oscuridad | 0–30; `-1` usa la configuración de la impresora |
| Velocidad | pulgadas/s; `0` usa la configuración de la impresora |
| Desplaz. X / Y | Corrige si la impresión sale corrida (mm, admite negativos) |
| Columna cantidad | Columna del Excel con cuántas copias imprimir de cada fila |
| Copias por registro | Multiplicador de copias (p. ej. 2 = doble de cada fila) |

Las filas con cantidad `0` no se imprimen.

## 7. Seleccionar la impresora

En el **encabezado**, elija la impresora de la lista (↻ actualiza la lista). Si su nombre
contiene `300dpi` o `600dpi` (como los drivers ZDesigner), la resolución cambia sola.

**Impresora por red sin driver:** elija *Impresora de red (IP / puerto 9100)…* y escriba la
IP en la pestaña **Impresora**.

Herramientas de la pestaña **Impresora**:

- **Imprimir prueba**: una etiqueta del registro actual.
- **Calibrar sensor (~JC)**: después de cambiar de rollo o tamaño de etiqueta.
- **Imprimir configuración (~WC)**: imprime el reporte de configuración de la Zebra.

## 8. Imprimir

- **Imprimir seleccionados**: las filas marcadas en la tabla (Ctrl/Shift + clic, o
  *Seleccionar todo*).
- **Imprimir todos**: todas las filas visibles (respeta el filtro de **Buscar**).

Antes de enviar se muestra un resumen con el total de etiquetas, la impresora y la
resolución. La barra de estado confirma el envío.

**Ver código** muestra el ZPL del registro actual (puede validarse en
labelary.com/viewer.html). **Exportar…** guarda el ZPL de las filas seleccionadas (o de todas)
en un archivo `.zpl`.

## 9. Plantillas

Una plantilla guarda el tamaño, la configuración y los elementos de la etiqueta.

- **Guardar** / **Guardar como…** → archivo `.json`.
- **Abrir…** carga una plantilla existente.
- **Nueva** reinicia con el diseño de ejemplo.

La misma plantilla sirve para las ZT411 de 300 y de 600 dpi. Al cerrar, el programa recuerda
la última plantilla, Excel, hoja e impresora.

## 10. Comprobación de impresora y drivers

**¿Hace falta instalar un driver?**

- Impresión **por red (IP)**: no.
- Impresión **por USB o por una impresora de Windows**: sí, el driver **ZDesigner** de Zebra.

El programa lo revisa solo al abrir y cada vez que cambia de impresora. Si encuentra algo,
muestra un aviso amarillo (advertencia) o rojo (error) bajo la barra de herramientas.
Pulse **Ver detalles**, o **Comprobar** en el encabezado, para ver la lista completa:

| Símbolo | Significado |
|---|---|
| ✔ verde | Correcto |
| ℹ azul | Información (p. ej. existe una versión más reciente del driver) |
| ⚠ ámbar | Advertencia: puede imprimir, pero conviene corregirlo |
| ✖ rojo | Error: la impresión probablemente fallará |

Qué se revisa:

- Que la impresora exista en Windows y qué driver y versión usa.
- Si Windows la tiene **sin conexión**, en pausa o con error.
- Para impresoras en red: si responde y su estado real (**sin papel**, **en pausa**,
  **cabezal abierto**, **sin ribbon**).
- Si hay una **Zebra conectada por USB sin driver**.

**Instalar el driver:** cuando hace falta aparece el botón **Instalar driver**.

- Si TI configuró el instalador en `driver.json`, el programa lo descarga, comprueba que sea el
  archivo correcto y lo ejecuta. Windows pedirá permiso de administrador.
- Si no está configurado, se abre la página oficial de Zebra con las instrucciones y la
  **versión exacta** que debe descargar.

> Si la impresora es **compartida desde un servidor** (su nombre empieza con `\`), el driver
> se actualiza en ese servidor; avise a TI.

## 11. Solución de problemas

| Problema | Solución |
|---|---|
| No imprime nada | Pulse **Comprobar** en el encabezado y siga las indicaciones. Pruebe **Imprimir prueba**. |
| «Windows marca la impresora sin conexión» | Encienda la impresora y revise el cable USB; en Windows, desmarque *Usar impresora sin conexión*. |
| «Zebra USB sin driver» | Pulse **Instalar driver**. |
| «Driver ZDesigner antiguo» | Actualice a la versión indicada (en el servidor, si la cola es compartida). |
| Imprime texto con códigos `^XA…` | El driver no es ZPL; reinstale con ZDesigner o use la conexión por red (IP). |
| La etiqueta sale cortada o corrida | Revise ancho/alto, calibre el sensor (~JC) y use *Desplaz. X / Y*. |
| Sale a la mitad o al doble de tamaño | La resolución no corresponde a la impresora: elija 300 o 600 dpi correctamente. |
| Impresión muy clara u oscura | Ajuste **Oscuridad** (p. ej. 15–25). |
| El código no se lee | Aumente el grosor de barra (≥ 0.25 mm) o el alto; evite ancho fijo demasiado angosto. |
| Recuadro rojo en la vista previa | Los datos no son válidos para ese tipo de código (p. ej. EAN-13 con letras). |
| Un código EAN pierde ceros a la izquierda | Ponga la columna del Excel en formato **Texto**. |
| Error al abrir el Excel | Cierre el archivo en Excel si está bloqueado, o guárdelo como `.xlsx`. |
