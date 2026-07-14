# Excel VBA Automation Toolkit

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
|---|---|---|
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

Subtotal = PrecioUnitario * Cantidad * Dias * Semanas


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


####  Agregar bloques de categoría

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