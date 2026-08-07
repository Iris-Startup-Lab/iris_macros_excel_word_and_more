# Excel VBA Automation Toolkit

Autores:

- Fernando Dorantes Nieto
- Azeneth Guadalupe García Méndez
- Javier Flores Martinez

Este proyecto contiene macros en VBA orientadas a automatizar dos procesos clave:

1. Generación de identificadores únicos (UUIDs)
2. Construcción y cálculo automático de cotizaciones estructuradas

---

# 1. Generación de UUIDs (VBA)

## Descripción técnica

La macro implementa un generador de identificadores únicos tipo GUID, produciendo strings con el formato estándar

---

# 3. UUID de cotización, colapso de detalle y PDFs (Resumen + Detalle)

Esta sección cubre la funcionalidad añadida para: (a) un **ID Cotización (UUID)
automático y estable**, (b) **colapsar/expandir** categorías y detalle a demanda,
y (c) generar **dos PDFs** (con y sin detalle) con un solo clic.

## Archivos (módulos a importar)

| Archivo | Tipo | Contenido |
| --- | --- | --- |
| `macro_generadora_uuid.vb` | Módulo | Generador base `GENERAR_UUID()` (ya existente, requerido). |
| `asignar_uuid_cotizacion.vb` | Módulo `M_UUID` | Asigna/mantiene el UUID en `K4` atado al proyecto (`C4`). Helpers y constantes compartidas. |
| `esquema_categorias.vb` | Módulo `M_Esquema` | Construye el esquema de 2 niveles y colapsa/expande/marca detalle. |
| `generar_pdfs_resumen_detalle.vb` | Módulo `M_PDF` | Genera `_Detalle.pdf` y `_Resumen.pdf` en un clic. |
| `eventos_uuid.vb` | **Eventos** | Código para `ThisWorkbook` y para la hoja `Cotización` (NO es un módulo normal). |

## Cómo funciona el UUID

- Se guarda en la celda **`K4` ("ID Cotización")** de la hoja `Cotización`.
- La llave es el **nombre del proyecto (`C4`)**. El UUID solo se **regenera si
  cambia `C4`**; se mantiene aunque cambie la **Versión (`N4`)** o se reabra el libro.
- El par `(proyecto, uuid)` se almacena en una hoja muy oculta llamada **`Meta`**
  que la macro crea automáticamente.
- El **nombre de los PDF incluye el UUID**, así archivo e ID interno coinciden.

## Colapso de detalle (2 niveles)

- **Nivel 1 – categoría:** se agrupa el cuerpo del bloque; al colapsar queda solo
  la barra de título y su `SUMA TOTAL`.
- **Nivel 2 – detalle:** las filas marcadas con `D` en la **columna oculta `O`**
  (el detalle largo de cada concepto) se agrupan y colapsan aparte.
- Marca filas con `MarcarComoDetalle` (o quítalas con `QuitarMarcaDetalle`),
  luego `ConstruirEsquema` para (re)generar los botones `[-]/[+]` del margen.
- Los **totales no cambian** al colapsar (son valores ya calculados por `suma`).

## Los dos PDFs

`GenerarPDFsCotizacion` produce en la carpeta del libro:

- `Cotizacion_<Cliente>_<Proyecto>_<UUID>_v<Versión>_<Fecha>_Detalle.pdf`
- `..._Resumen.pdf` (con las filas `D` ocultas)

## Instalación (una sola vez)

1. Abre el `.xlsm`, `Alt+F11` (editor VBA).
2. `Archivo > Importar archivo...` e importa los 4 módulos:
   `macro_generadora_uuid.vb`, `asignar_uuid_cotizacion.vb`,
   `esquema_categorias.vb`, `generar_pdfs_resumen_detalle.vb`.
3. Abre `eventos_uuid.vb` y copia **BLOQUE A** dentro de `ThisWorkbook` y
   **BLOQUE B** dentro del objeto de la hoja `Cotización` (doble clic en cada uno).
4. Asigna macros a botones en la hoja (clic derecho en el botón > *Asignar macro*):
   - Generar PDFs → `GenerarPDFsCotizacion`
   - Reconstruir esquema → `ConstruirEsquema`
   - Marcar detalle → `MarcarComoDetalle` / Quitar → `QuitarMarcaDetalle`
   - (opcional) Colapsar/Expandir detalle → `OcultarDetalle` / `MostrarTodo`
5. Guarda el libro como `.xlsm` (habilitado para macros).

> La primera vez que haya un nombre de proyecto en `C4`, el UUID sustituirá al
> número manual que hoy está en `K4`.

---

# 4. Blindaje de la plantilla (botones fijos, ID irrompible, una sola hoja)

Esta sección cubre tres fallas operativas reportadas en producción y su solución.

## Archivos

| Archivo | Tipo | Contenido |
| --- | --- | --- |
| `blindaje_botones.vb` / `bas/blindaje_botones.bas` | Módulo `M_Blindaje` | Anclaje, registro y bloqueo de botones; protección de estructura del libro. |
| `eventos_uuid.vb` | **Eventos** | Reescrito: ata las tres protecciones a `ThisWorkbook` y a la hoja `Cotización`. |

## 4.1 Los botones se movían a la hoja principal

**Causa raíz.** Por omisión todo objeto de Excel nace con
`Placement = xlMoveAndSize`, es decir **atado a las celdas que tiene debajo**.
Cada operación de las macros lo altera:

| Operación | Efecto sobre el botón |
| --- | --- |
| `Rows(x).Insert` (`AgregarFilaProducto`, `AgregarFilaFecha`, `DuplicarBloque`) | Lo arrastra hacia abajo. |
| `Rows(x).Copy` + `Insert` (`AgregarFilaItem`) | Lo **duplica**. |
| `Rows(x).Delete` (`BorrarFilaSeleccionada`) | Lo sube o lo recorta. |
| `EntireRow.Hidden = True` (PDF Resumen, `OcultarDetalle`) | Lo aplasta a altura ~0. |
| `.Group` / `ClearOutline` (`ConstruirEsquema`) | Lo desplaza al colapsar. |

**Solución (3 capas).**

1. **Anclaje** — `FijarBotones` pone `Placement = xlFreeFloating` en todos los
   objetos: dejan de estar ligados a celdas, así que insertar, borrar u ocultar
   filas ya no los toca ni los clona.
2. **Registro de posición** — `GuardarPosicionesBotones` guarda
   `nombre / Left / Top / Width / Height` de cada objeto en la hoja `Meta`
   (columnas `D:H`). `RestaurarPosicionesBotones` los devuelve ahí.
   Se llama solo al abrir el libro, al activar la hoja, al guardar, al imprimir
   y al final de cada macro de estructura. Es idempotente: si nada se movió no
   escribe nada y **no** marca el libro como modificado.
3. **Bloqueo** — `BloquearBotones` aplica
   `ws.Protect DrawingObjects:=True, Contents:=False`. El usuario ya no puede
   arrastrar, redimensionar ni borrar un botón, pero **las celdas siguen 100%
   editables** y ninguna macro existente se rompe.

`LimpiarBotonesDuplicados` detecta los objetos que no están en el registro
(los clones históricos) y los elimina previa confirmación.

## 4.2 El ID Cotización se borraba y no volvía

**Causa raíz.** `Worksheet_Change` solo vigilaba `C4`. Si el usuario borraba
`K4`, nada lo repuso.

**Solución.** La hoja `Meta` (`B2`) es la fuente de la verdad; `K4` es solo su
proyección. `VigilarIDCotizacion` se dispara con cambios en `C4` **o `K4`**:

- `K4` vacío o alterado a mano → se **reimprime el mismo UUID** (no se genera
  uno nuevo) y se avisa al usuario.
- `C4` cambió → se asigna el UUID que corresponde al nuevo proyecto.

Además `RestaurarIDSiFalta` corre al abrir, al guardar, al imprimir y antes de
generar los PDFs, así que ningún archivo se guarda ni exporta sin ID.

> Único caso en que `K4` queda vacío a propósito: cuando `C4` (proyecto) está
> vacío. Sin proyecto no hay identidad que registrar.

## 4.3 Se podían crear más hojas de cotización

**Solución (2 capas).**

1. **Nativa** — `ProtegerEstructura` aplica
   `ThisWorkbook.Protect Structure:=True`, con lo que Excel deshabilita
   *Insertar hoja*, *Eliminar hoja*, *Mover o copiar hoja*, *Cambiar nombre* y
   *Ocultar/Mostrar*. Se reaplica en `Workbook_Open` y `Workbook_BeforeSave`.
2. **Evento** — `Workbook_NewSheet` elimina al vuelo cualquier hoja que llegue a
   crearse y explica que una cotización = un archivo `.xlsm`.

La bandera pública `gCreandoHojaSistema` permite que las macros creen la hoja de
sistema `Meta` sin que el evento la borre.

## Puesta en marcha

Con el panel de botones ya acomodado como debe quedar, ejecuta **una vez**
`BlindarLibro`: ancla, registra posiciones, asegura el ID y protege la
estructura. Para mantenimiento, `DesblindarLibro` retira el blindaje (se vuelve
a aplicar solo al reabrir el archivo).

> La contraseña vive en `PWD_BLINDAJE` (módulo `M_Blindaje`). **Cámbiala** antes
> de repartir el archivo y anótala: sin ella no se puede desproteger.

---

# 2. Construcción y cálculo automático de cotizaciones estructuradas

## Descripción técnica

Este conjunto de macros implementa un sistema de cotización dinámico que:

Organiza servicios por categorías
Calcula subtotales, descuentos e impuestos

Permite manipulación dinámica del documento (filas, bloques, exportación)

Cada fila representa una entidad con atributos:

Concepto = {
    descripcion,
    precio_unitario,
    cantidad,
    dias,
    semanas,
    subtotal,
    descuento,
    total
}

Subtotal = PrecioUnitario *Cantidad* Dias * Semanas

Total = Subtotal * (1 - PorcentajeDescuento)

### Suma por categoría

La macro:

- Identifica bloques de categorías
- Recorre cada fila activa
- Acumula totales

SubtotalGeneral = Σ TotalesCategorias

IVA = SubtotalGeneral * 0.16

TotalFinal = SubtotalGeneral + IVA

### Manipulación dinámica del documento

Agregar filas

Inserta nuevas filas dentro de una categoría.
Replica formatos y fórmulas.
Ajusta referencias automáticamente.

#### Agregar bloques de categoría

Duplica una estructura completa:

Encabezado
Filas.
Fórmulas de suma.

#### Eliminación de filas

Borra filas activas.
Recalcula totales automáticamente.

### Manejo de fechas

Cada bloque puede incluir fechas específicas.

Las macros permiten insertar filas con fechas dinámicas.

#### Validaciones implícitas

Conversión numérica segura
Prevención de celdas vacías en cálculos
Control de errores básicos (On Error Resume Next en algunas rutinas)
