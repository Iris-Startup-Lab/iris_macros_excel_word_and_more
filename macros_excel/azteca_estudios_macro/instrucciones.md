# Instrucciones de instalación y uso — Cotizaciones Azteca Estudios

Guía paso a paso para instalar y usar las macros de:

1. **UUID automático** (ID Cotización estable, celda `K4`).
2. **Colapso a demanda** de categorías y detalle (2 niveles).
3. **Generación de 2 PDFs** con un clic (Resumen + Detalle).

> Escrito para que cualquier persona pueda instalarlo desde cero, sin conocer VBA.
> Si algo falla, ve directo a la sección **9. Solución de problemas**.

---

## Índice

- [0. Antes de empezar (requisitos)](#0-antes-de-empezar-requisitos)
- [1. Habilitar macros y la pestaña Programador](#1-habilitar-macros-y-la-pestaña-programador)
- [2. Abrir el editor de VBA](#2-abrir-el-editor-de-vba)
- [3. Importar los 4 módulos](#3-importar-los-4-módulos)
- [4. Pegar el código de EVENTOS (paso delicado)](#4-pegar-el-código-de-eventos-paso-delicado)
- [5. Crear los botones en la hoja](#5-crear-los-botones-en-la-hoja)
- [6. Uso diario](#6-uso-diario)
- [7. Cómo funciona el UUID (referencia)](#7-cómo-funciona-el-uuid-referencia)
- [8. Estructura de un bloque (referencia)](#8-estructura-de-un-bloque-referencia)
- [9. Solución de problemas](#9-solución-de-problemas)
- [10. Pruebas de verificación](#10-pruebas-de-verificación)
- [Apéndice A. Lista de macros](#apéndice-a-lista-de-macros)
- [Apéndice B. Celdas y hojas usadas](#apéndice-b-celdas-y-hojas-usadas)

---

## 0. Antes de empezar (requisitos)

- **Microsoft Excel para Windows** (las macros usan una función del sistema Windows
  para generar el UUID; no funcionan en Excel para Mac ni en Excel Web).
- El archivo de trabajo debe ser **`.xlsm`** (Libro de Excel habilitado para macros).
  Si tu archivo es `.xlsx`, guárdalo primero como `.xlsm`
  (`Archivo > Guardar como > Tipo: Libro de Excel habilitado para macros`).
- Ten a la mano, en la misma carpeta, estos archivos de código:
  - `macro_generadora_uuid.vb`  *(generador base, ya existente — requerido)*
  - `asignar_uuid_cotizacion.vb`
  - `esquema_categorias.vb`
  - `generar_pdfs_resumen_detalle.vb`
  - `eventos_uuid.vb`  *(este NO se importa; se copia y pega — ver paso 4)*

> **Recomendación:** haz una **copia de respaldo** del `.xlsm` antes de empezar
> (ya existe `Plantilla Cotización AE 1_bkp.xlsm`, pero por si acaso).

---

## 1. Habilitar macros y la pestaña Programador

### 1.1 Mostrar la pestaña "Programador" (Developer)

1. `Archivo > Opciones > Personalizar cinta de opciones`.
2. En la lista de la derecha, marca la casilla **Programador**.
3. Acepta. Verás una nueva pestaña **Programador** en la cinta.

### 1.2 Permitir la ejecución de macros

1. `Archivo > Opciones > Centro de confianza > Configuración del Centro de confianza`.
2. Entra a **Configuración de macros**.
3. Selecciona **Deshabilitar todas las macros con notificación**
   (así Excel te preguntará y podrás habilitarlas al abrir el archivo).
4. Acepta todo.

> Al abrir el `.xlsm` verás una barra amarilla: **"Habilitar contenido"**. Haz clic
> en ella cada vez que abras el archivo (o marca el archivo como ubicación confiable).

---

## 2. Abrir el editor de VBA

- Con el `.xlsm` abierto, presiona **`Alt + F11`**.
- Se abre el **Editor de Visual Basic (VBA)**.
- Si no ves el panel de la izquierda (**Explorador de proyectos**), presiona **`Ctrl + R`**.

El Explorador de proyectos se ve así:

```
VBAProject (Plantilla Cotización AE 1.xlsm)
├── 📁 Microsoft Excel Objetos
│    ├── Hoja1 (Cotización )        ← objeto de la hoja de cotización
│    ├── Hoja16 (Plantillas)
│    ├── Hoja11 (Categorías)
│    ├── ...
│    └── ThisWorkbook               ← objeto del libro completo
└── 📁 Módulos
     └── (aquí aparecerán los módulos que importemos)
```

> **Importante:** el nombre entre paréntesis `(Cotización )` es el que ves en la
> pestaña de Excel. El `HojaN` de la izquierda es el nombre interno; ignóralo.

---

## 3. Importar los 4 módulos

Repite esto **una vez por cada archivo** de la lista de abajo:

1. En el editor VBA, menú **`Archivo > Importar archivo...`**
   (o clic derecho sobre `VBAProject` > **Importar archivo...**).
2. Navega a la carpeta `azteca_estudios_macro` y selecciona el archivo.
3. Clic en **Abrir**.

Importa **en este orden** (no es obligatorio, pero es lo más claro):

1. `macro_generadora_uuid.vb`
2. `asignar_uuid_cotizacion.vb`
3. `esquema_categorias.vb`
4. `generar_pdfs_resumen_detalle.vb`

Al terminar, en la carpeta **Módulos** del Explorador deberías ver los 4:

```
└── 📁 Módulos
     ├── macro_generadora_uuid
     ├── asignar_uuid_cotizacion
     ├── esquema_categorias
     └── generar_pdfs_resumen_detalle
```

> ⚠️ **NO importes `eventos_uuid.vb`.** Ese archivo es solo una referencia; su
> contenido se pega a mano en el paso 4.

---

## 4. Pegar el código de EVENTOS (paso delicado)

Este código son **eventos automáticos** ("cuando se abra el libro…", "cuando cambie
una celda…"). Solo funcionan si viven dentro de objetos específicos; por eso **no**
se importan como módulo, sino que se **copian y pegan** en dos lugares.

### 4.1 BLOQUE A → en `ThisWorkbook`

1. En el Explorador de proyectos, **doble clic en `ThisWorkbook`**.
2. Se abre una ventana de código (a la derecha), probablemente en blanco.
3. Copia y pega **exactamente** esto:

```vb
Private Sub Workbook_Open()
    On Error Resume Next
    AsegurarUUIDCotizacion
End Sub
```

### 4.2 BLOQUE B → en la hoja `Cotización`

1. En el Explorador, **doble clic en el objeto `HojaN (Cotización )`**
   (el que tiene `Cotización` entre paréntesis).
2. Se abre otra ventana de código en blanco.
3. Copia y pega **exactamente** esto:

```vb
Private Sub Worksheet_Change(ByVal Target As Range)
    ' Si se modifica C4 (nombre del proyecto), revisamos/reasignamos el UUID.
    If Intersect(Target, Me.Range("C4")) Is Nothing Then Exit Sub

    On Error GoTo salir
    Application.EnableEvents = False
    AsegurarUUIDCotizacion
salir:
    Application.EnableEvents = True
End Sub
```

### 4.3 Guardar

- Presiona **`Ctrl + S`** (o vuelve a Excel y guarda).
- Si Excel te ofrece guardar sin macros, di **No** y elige el formato **`.xlsm`**.

> **Cómo distinguir un módulo de un objeto de hoja:** los módulos están en la
> carpeta *Módulos* y sirven para código que llamas tú (botones). Los objetos
> `ThisWorkbook` y `HojaN (...)` sirven para código que Excel dispara solo (eventos).

---

## 5. Crear los botones en la hoja

Puedes usar botones de formulario para no tener que abrir el editor cada vez.

1. Ve a Excel, hoja **Cotización**.
2. Pestaña **Programador > Insertar > Botón (control de formulario)**.
3. Dibuja el botón donde quieras.
4. Aparece el cuadro **Asignar macro**: elige la macro y **Aceptar**.
5. Clic derecho en el botón > **Editar texto** para ponerle un nombre.

Botones recomendados:

| Texto del botón | Macro a asignar | Para qué sirve |
|---|---|---|
| **Generar PDFs** | `GenerarPDFsCotizacion` | Crea los 2 PDF (Resumen + Detalle). |
| **Reconstruir esquema** | `ConstruirEsquema` | Regenera los botones `[-]/[+]` de colapso. |
| **Marcar detalle** | `MarcarComoDetalle` | Marca las filas seleccionadas como detalle. |
| **Quitar marca detalle** | `QuitarMarcaDetalle` | Quita la marca de detalle. |
| **Ocultar detalle** | `OcultarDetalle` | Vista resumen (oculta filas de detalle). |
| **Mostrar todo** | `MostrarTodo` | Vuelve a mostrar todas las filas. |

> Si un botón ya existía (por ejemplo el de generar PDF viejo), clic derecho >
> **Asignar macro** y cámbialo a la macro nueva.

---

## 6. Uso diario

### 6.1 Llenar la cotización

- Escribe **Cliente** en `C3` y **Nombre del proyecto** en `C4`.
- En cuanto `C4` tenga texto, la celda `K4` ("ID Cotización") se llenará **sola**
  con el UUID. No lo edites a mano.
- Sube la **Versión** en `N4` cuando saques una nueva versión (el UUID NO cambia).

### 6.2 Marcar qué filas son "detalle colapsable"

1. Selecciona las filas de detalle (las de descripción larga que a veces el cliente
   no necesita ver). Puedes seleccionar varias filas a la vez.
2. Clic en el botón **Marcar detalle** (macro `MarcarComoDetalle`).
   - Esto escribe una `D` en la columna oculta `O` de esas filas.
3. Si te equivocas, selecciona y usa **Quitar marca detalle**.

### 6.3 Generar el esquema de colapso (botones `[-]/[+]`)

- Clic en **Reconstruir esquema** (macro `ConstruirEsquema`).
- Aparecerán los controles de esquema en el margen izquierdo:
  - **Nivel 1**: colapsa una categoría completa (queda solo título + SUMA TOTAL).
  - **Nivel 2**: colapsa solo el detalle marcado con `D`.
- Repite este paso cada vez que agregues/quites bloques o cambies marcas de detalle.

### 6.4 Generar los PDFs

- Clic en **Generar PDFs** (macro `GenerarPDFsCotizacion`).
- Se crean **2 archivos** en la misma carpeta del libro:
  - `Cotizacion_<Cliente>_<Proyecto>_<UUID>_v<Versión>_<Fecha>_Detalle.pdf`
  - `..._Resumen.pdf`  *(mismo contenido, pero con las filas `D` ocultas)*
- Al final aparece un mensaje con el nombre de los archivos y el UUID.

> Los **totales no cambian** entre Resumen y Detalle: el Resumen solo muestra
> menos renglones, pero el SUMA TOTAL / SUBTOTAL / IVA / TOTAL son los mismos.

---

## 7. Cómo funciona el UUID (referencia)

- **Dónde vive:** celda `K4` ("ID Cotización") de la hoja Cotización.
- **Llave de identidad:** el **nombre del proyecto** (`C4`).
- **Cuándo se genera uno nuevo:** solo si `C4` cambia (o si aún no había UUID).
- **Cuándo NO cambia:** al subir la Versión (`N4`), al reabrir el archivo, al
  generar PDFs, o al editar cualquier otra celda.
- **Dónde se recuerda:** en una hoja **muy oculta** llamada `Meta` que la macro
  crea sola. Guarda el par `(nombre de proyecto, uuid)`.
- **Por qué el nombre de archivo lo incluye:** para que el PDF y el ID interno
  siempre coincidan (trazabilidad).

Para ver/forzar el UUID manualmente puedes asignar un botón a `AsignarUUID_Click`.

---

## 8. Estructura de un bloque (referencia)

La macro reconoce cada categoría por su estructura (ejemplo real del archivo):

```
Fila 7-8   ┃ FOROS Y LOCACIONES              ← TÍTULO (barra negra, fusionada)
Fila 9     ┃ DESCRIPCIÓN | PRECIO | CANT...  ← encabezados de columna
Fila 10    ┃ 08/06/2027                      ← fecha
Fila 11    ┃ Foro 1                  175,000 ← concepto principal
Fila 12    ┃ CARPA/EXPLANADA: ...     45,000 ← DETALLE (se marca con "D")
Fila 13    ┃ SUMA TOTAL:            215,500 ← cierre del bloque
```

- **Inicio del bloque:** la macro compara la columna A contra la lista de la hoja
  `Categorías`. Si coincide (ej. `FOROS Y LOCACIONES`), ahí empieza un bloque.
- **Fin del bloque:** la fila que dice `SUMA TOTAL:`.
- **Nivel 1** agrupa todo el cuerpo entre el título y la SUMA.
- **Nivel 2** agrupa las filas marcadas con `D`.

> Por eso es importante que los títulos de categoría en la hoja Cotización estén
> escritos **igual** que en la hoja `Categorías` (mayúsculas/acentos no importan,
> pero el texto sí debe coincidir).

---

## 9. Solución de problemas

| Síntoma | Causa probable | Solución |
|---|---|---|
| No pasa nada al hacer clic en un botón | Macros deshabilitadas | Habilita el contenido (barra amarilla) y revisa el paso 1.2. |
| `K4` no se llena solo | Falta el BLOQUE A/B, o `C4` está vacío | Verifica el paso 4; escribe un proyecto en `C4`. |
| Error "Sub o Function no definida" | Falta importar un módulo | Revisa que estén los 4 módulos (paso 3); `macro_generadora_uuid` es obligatorio. |
| El UUID cambia solo sin querer | Cambió el texto de `C4` | Es el comportamiento esperado: el UUID se ata al nombre del proyecto. |
| Los botones `[-]/[+]` no aparecen | No se corrió `ConstruirEsquema` | Ejecuta **Reconstruir esquema**. |
| Una categoría no agrupa | Su título no coincide con la hoja `Categorías` | Corrige el texto del título o agrégalo a `Categorías`. |
| La columna `O` se ve en la hoja | Quedó visible | Se oculta sola al correr las macros; o clic derecho en col. O > Ocultar. |
| El PDF sale con filas de más/menos | Marcas de detalle mal puestas | Revisa qué filas tienen `D` (muestra col. O), corrige y regenera. |
| "No se encontró la hoja de Cotización" | Se renombró la hoja | La macro busca una hoja cuyo nombre contenga "COTIZACI"; no le quites esa palabra. |
| Excel avisa que hay una hoja oculta `Meta` | Es normal | No la borres: ahí se guarda el UUID. Está muy oculta a propósito. |

---

## 10. Pruebas de verificación

Haz este mini-recorrido la primera vez para confirmar que todo quedó bien:

1. **UUID al escribir proyecto:** borra `C4`, escribe "Prueba 1". → `K4` muestra un
   UUID. Anótalo.
2. **UUID estable con versión:** cambia `N4` de 1 a 2. → `K4` **no** cambia.
3. **UUID nuevo al cambiar proyecto:** cambia `C4` a "Prueba 2". → `K4` cambia a un
   UUID distinto.
4. **UUID persiste al reabrir:** guarda, cierra y reabre el archivo. → `K4` conserva
   el mismo UUID de "Prueba 2".
5. **Marcado y esquema:** marca 1-2 filas de detalle, corre `ConstruirEsquema`. →
   aparecen los `[-]/[+]` y al colapsar desaparece el detalle pero no la SUMA.
6. **PDFs:** corre `GenerarPDFsCotizacion`. → aparecen 2 archivos `_Detalle.pdf` y
   `_Resumen.pdf` en la carpeta del libro, y el Resumen no muestra las filas `D`.

Si los 6 pasos funcionan, la instalación está completa.

---

## Apéndice A. Lista de macros

| Macro | Módulo | Qué hace |
|---|---|---|
| `GENERAR_UUID` | `macro_generadora_uuid` | Genera un GUID/UUID (función base). |
| `AsegurarUUIDCotizacion` | `asignar_uuid_cotizacion` | Asigna/mantiene el UUID en `K4`. |
| `UUIDActual` | `asignar_uuid_cotizacion` | Devuelve el UUID actual (asegurándolo). |
| `AsignarUUID_Click` | `asignar_uuid_cotizacion` | Botón manual: fuerza y muestra el UUID. |
| `ConstruirEsquema` | `esquema_categorias` | Crea el esquema de 2 niveles. |
| `MostrarTodo` | `esquema_categorias` | Muestra todas las filas. |
| `OcultarDetalle` | `esquema_categorias` | Oculta las filas de detalle (`D`). |
| `MarcarComoDetalle` | `esquema_categorias` | Marca filas seleccionadas como detalle. |
| `QuitarMarcaDetalle` | `esquema_categorias` | Quita la marca de detalle. |
| `GenerarPDFsCotizacion` | `generar_pdfs_resumen_detalle` | Genera los 2 PDF. |
| `Workbook_Open` | `ThisWorkbook` (evento) | Asegura el UUID al abrir. |
| `Worksheet_Change` | Hoja Cotización (evento) | Revisa el UUID al cambiar `C4`. |

*(Se conservan además las macros originales: `suma`, `DuplicarBloque`,
`AgregarFilaItem`, `AgregarFilaProducto`, `AgregarFilaFecha`,
`BorrarFilaSeleccionada`, `LimpiarCotizacion`, `GenerarPDF`,
`GenerarPDFNombreInteligente`.)*

---

## Apéndice B. Celdas y hojas usadas

**Hoja Cotización — cabecera:**

| Celda | Contenido |
|---|---|
| `C3` | Cliente |
| `H3` | ID Cliente (lista) |
| `K3` | Fecha de cotización |
| `C4` | Nombre del proyecto (**llave del UUID**) |
| `H4` | ID Proyecto (lista) |
| `K4` | **ID Cotización (UUID) — automático** |
| `N4` | Versión |

**Columna auxiliar:**

| Columna | Uso |
|---|---|
| `O` | Marca de detalle (`D`). Oculta. Fuera del área de impresión (datos llegan a `N`). |

**Hojas del libro:**

| Hoja | Uso |
|---|---|
| `Cotización` | Documento principal. |
| `Plantillas` | Bloques/formatos base para duplicar. |
| `Categorías` | Lista oficial de categorías (usada para detectar bloques). |
| `ID Cliente`, `ID Proyecto`, `Versión` | Listas de validación. |
| `Meta` | **Creada por la macro.** Guarda `(proyecto, uuid)`. Muy oculta. |
| `Espacios`, `Instructivo` | Auxiliares existentes. |

