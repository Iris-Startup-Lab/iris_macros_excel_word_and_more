# AGENTS.md — Guía para Agentes de IA en Azteca Estudios Macro

Este documento proporciona contexto técnico, arquitectura, convenciones y reglas operativas para agentes de IA (y desarrolladores) que interactúen, mantengan o extiendan este repositorio.

---

## 1. Visión General del Proyecto

**Nombre:** Excel VBA Automation Toolkit — Cotizaciones Azteca Estudios  
**Propósito:** Sistema de automatización en VBA para Microsoft Excel en Windows orientado a la generación, estructuración, cálculo, colapso visual y exportación a PDF de cotizaciones comerciales de Azteca Estudios.

### Funcionalidades Clave
1. **UUID Estable y Automático:** Generación de un UUID único por cotización en la celda `K4`, vinculado al nombre del proyecto (`C4`) y persistido en una hoja de sistema muy oculta (`Meta`).
2. **Esquema Multinivel (Agrupación/Colapso):** Colapso y expansión a demanda de categorías completas (Nivel 1) y líneas de detalle técnico (Nivel 2, identificadas por la marca `"D"` en la columna oculta `O`).
3. **Exportación Dual a PDF en 1 Clic:** Generación automática de dos archivos PDF (`..._Resumen.pdf` y `..._Detalle.pdf`) con nomenclatura estandarizada.
4. **Cálculo Dinámico de Cotizaciones:** Recálculo automático de subtotales por bloque de categoría, Subtotal General, IVA (16%) y Total Final (`suma.vb`).
5. **Manipulación Dinámica de Estructura:** Inserción y eliminación de filas, adición de fechas/ítems y duplicación de bloques de categorías desde la hoja `Plantillas`.

---

## 2. Mapa y Estructura del Repositorio

```
azteca_estudios_macro/
├── AGENTS.md                                # Este documento de contexto para agentes de IA
├── README.md                                # Documentación técnica general y visión del proyecto
├── instrucciones.md                         # Guía paso a paso de instalación y solución de problemas para usuarios
├── GUÍA DE USO COTIZACIÓN.pdf              # Manual visual de usuario en PDF
│
├── Plantilla Cotización AE 1.xlsm           # Plantilla principal de trabajo (habilitada para macros)
├── Plantilla Cotización AE 1_para_demo.xlsm  # Versión de demostración
├── Plantilla Cotización AE 1_bkp.xlsm       # Backup de la plantilla
├── Macro_generadora_uuids.xlsm              # Libro auxiliar con prueba de UUID
│
├── bas/                                     # Exportaciones .bas para importar directamente al editor de VBA (Alt+F11)
│   ├── macro_generadora_uuid.bas
│   ├── asignar_uuid_cotizacion.bas
│   ├── esquema_categorias.bas
│   └── generar_pdfs_resumen_detalle.bas
│
├── Módulos Principales de la Nueva Versión (.vb):
│   ├── macro_generadora_uuid.vb            # Función base GENERAR_UUID() mediante API de Windows (ole32.dll)
│   ├── asignar_uuid_cotizacion.vb          # Módulo M_UUID (Gestión de K4, C4, hoja Meta y constantes globales)
│   ├── esquema_categorias.vb               # Módulo M_Esquema (Agrupación Outline nivel 1 y 2, columna O)
│   ├── generar_pdfs_resumen_detalle.vb     # Módulo M_PDF (Exportación de _Resumen.pdf y _Detalle.pdf)
│   └── eventos_uuid.vb                     # Código de eventos (ThisWorkbook y Worksheet_Change)
│
└── Macros Utilitarias Secundarias/Legacy (.vb):
    ├── suma.vb                             # Recálculo de totales por categoría, Subtotal, IVA y Total
    ├── duplicar_bloque.vb                  # Copia estructuras de bloque desde la hoja "Plantillas"
    ├── agregar_fila_item.vb                # Inserta fila simple dentro de una categoría
    ├── agregar_fila_producto.vb            # Inserta fila de producto formateada
    ├── agregar_fila_fecha.vb               # Inserta fila con campo fecha
    ├── borrar_fila_seleccionada.vb         # Elimina fila activa manteniendo consistencia
    ├── limpiar_cotizacion.vb               # Resetea campos de la cotización
    ├── generar_pdf.vb                      # Exportación simple a PDF
    └── generar_pdf_nombre_inteligente.vb   # Exportación PDF con nombre derivado de celdas
```

---

## 3. Arquitectura y Componentes Clave

### A. Módulo `M_UUID` (`asignar_uuid_cotizacion.vb` / `.bas`)
- **Celda Destino:** `K4` ("ID Cotización"). Formateada explícitamente como texto (`@`).
- **Celda Disparadora:** `C4` ("Nombre del Proyecto").
- **Celda Versión:** `N4`.
- **Hoja Meta (`Meta`):** Creada automáticamente si no existe. Configurada como `xlSheetVeryHidden` (invisible en la interfaz de Excel, no desocultable manualmente por el usuario). Guarda la relación `(Proyecto, UUID, Fecha)`.
- **Constantes Compartidas:**
  - `COL_MARCA = "O"`: Columna oculta que contiene la marca de detalle.
  - `MARCA_DETALLE = "D"`: Valor que identifica una fila como detalle técnico colapsable.
  - `FILA_INICIO = 7`: Primera fila a evaluar en el cuerpo de la cotización.
- **Helper de Búsqueda de Hojas (`HojaQueContiene` / `HojaCotizacion`):** Búsqueda flexible e insensible a mayúsculas/acentos/espacios para prevenir fallos por renombrado parcial de hojas (ej. `"Cotización "`).

### B. Módulo `M_Esquema` (`esquema_categorias.vb` / `.bas`)
- **Nivel 1 (Categoría):** Agrupa desde la fila posterior al título de la categoría hasta la fila previa a `SUMA TOTAL`. Permite colapsar el bloque completo dejando visible solo el encabezado y su total.
- **Nivel 2 (Detalle):** Detecta rachas consecutivas de filas con valor `"D"` en la columna `O` y aplica `.Group`.
- **Carga de Categorías (`CargarCategorias`):** Lee la hoja `Categorías` para identificar dinámicamente los nombres válidos de categoría.
- **MergeArea Handling:** `FilaFinEncabezado` determina correctamente el final de títulos fusionados (ej. `A7:N8`).

### C. Módulo `M_PDF` (`generar_pdfs_resumen_detalle.vb` / `.bas`)
- Exporta dos documentos PDF en la misma carpeta del archivo `.xlsm`:
  1. `..._Detalle.pdf`: Muestra todas las filas visibles y detalladas.
  2. `..._Resumen.pdf`: Oculta temporalmente las filas con marca `"D"` antes de exportar y restaura su visibilidad inmediatamente después.
- **Formato del Nombre:** `Cotizacion_<Cliente>_<Proyecto>_<UUID>_v<Version>_<Fecha>_<Tipo>.pdf`. Los caracteres no válidos para rutas Windows (`\`, `/`, `:`, `*`, `?`, `"`, `<`, `>`, `|`) son higienizados automáticamente.

### D. Eventos (`eventos_uuid.vb`)
- **`ThisWorkbook` (`Workbook_Open`):** Ejecuta `AsegurarUUIDCotizacion` al abrir el archivo.
- **`Worksheet_Change` (en la Hoja `Cotización`):** Monitorea cambios en `C4`. Si `C4` se modifica, desactiva `Application.EnableEvents = False`, ejecuta `AsegurarUUIDCotizacion` y reactiva los eventos.

### E. Recálculo Matemático (`suma.vb`)
- Recorre la columna `A` buscando etiquetas `"SUMA TOTAL"` o `"SUMA TOTAL:"`.
- Acumula los valores numéricos de la columna `M` (Total por renglón) hacia la celda `M` de la fila de suma.
- Busca en la columna `I` los rótulos `"SUBTOTAL"`, `"IVA"` y `"TOTAL"` para asignar en la columna `K`:
  - `SubtotalGeneral` = $\sum \text{Totales Categorías}$
  - `IVA` = $\text{SubtotalGeneral} \times 0.16$
  - `TotalFinal` = $\text{SubtotalGeneral} + \text{IVA}$

---

## 4. Convenciones Técnicas y Reglas para Agentes

Al modificar o crear código en este repositorio, **sigue estrictamente las siguientes reglas**:

### 1. Compatibilidad de Entorno
- **Sistema Operativo Target:** Windows. La generación de UUIDs depende de las funciones Windows API (`CoCreateGuid` y `StringFromGUID2` en `ole32.dll`).
- **Formato de Libro:** Todo archivo funcional debe ser `.xlsm` (Libro de Excel Habilitado para Macros). Nunca guardar versiones activas como `.xlsx`.

### 2. Sincronización Doble de Código (`.vb` y `bas/*.bas`)
- Los archivos `.vb` en la raíz son la fuente principal de código legible en UTF-8 / Markdown.
- Los archivos `.bas` en la carpeta `bas/` corresponden a los módulos estándar listos para importación en el editor VBA de Excel.
- **REGLA:** Si realizas un cambio en una macro modular (`asignar_uuid_cotizacion`, `esquema_categorias`, `generar_pdfs_resumen_detalle`, `macro_generadora_uuid`), **debes actualizar tanto el archivo `.vb` en la raíz como el correspondiente `.bas` dentro de `bas/`**.

### 3. Buenas Prácticas de Código VBA
- **`Option Explicit` OBLIGATORIO:** Todas las variables deben declararse explícitamente (`Dim`).
- **Manejo de Pantalla y Eventos:** Al realizar operaciones masivas o de ocultación temporal (como en esquemas o generación de PDFs):
  ```vba
  Application.ScreenUpdating = False
  Application.EnableEvents = False
  ' ... código ...
  Application.EnableEvents = True
  Application.ScreenUpdating = True
  ```
- **Control de Errores Limpio:** Usar `On Error Resume Next` / `On Error GoTo 0` o etiquetas de salida limpias (`On Error GoTo salir`) garantizando que `EnableEvents` y `ScreenUpdating` se restablezcan en `True` siempre.
- **Acceso Robusto a Hojas:** No asumas el nombre exacto de la pestaña ni el CodeName interno (`Hoja1`). Usa la función helper `HojaCotizacion()` / `HojaQueContiene("COTIZACI")` para tolerar variaciones de espacios o acentos.

### 4. Mapeo Rígido de Celdas y Columnas
| Elemento / Campo | Ubicación en Excel |
| --- | --- |
| Nombre del Cliente | Celda `C3` |
| Nombre del Proyecto | Celda `C4` (Dispara UUID) |
| ID Cotización (UUID) | Celda `K4` (Formato texto `@`) |
| Versión | Celda `N4` |
| Columna de Estructura / Categorías | Columna `A` |
| Rótulos Totales Generales | Columna `I` (`SUBTOTAL`, `IVA`, `TOTAL`) |
| Subtotales de Renglón / Categoría | Columna `M` |
| Valores Totales Generales | Columna `K` |
| Marca Oculta de Detalle | Columna `O` (`"D"` = Fila de detalle) |
| Hoja Meta de Sistema | `Worksheets("Meta")` (`xlSheetVeryHidden`) |

---

## 5. Guía de Pruebas y Verificación

Cuando realices modificaciones a las macros, efectúa las siguientes verificaciones:

1. **Verificación Sintáctica VBA:** Asegurar que no existan variables sin declarar ni errores de compilación (`Option Explicit`).
2. **Prueba de UUID:**
   - Escribir un nombre de proyecto en `C4`. `K4` debe poblarse con un UUID válido de 36 caracteres en formato `XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX`.
   - Modificar celdas no relacionadas (ej. `N4` o valores de ítems); `K4` debe conservar el mismo UUID.
   - Cambiar el valor de `C4`; el UUID en `K4` debe renovarse automáticamente y registrarse en la hoja `Meta`.
3. **Prueba de Esquema:**
   - Marcar filas con `"D"` en la columna `O`.
   - Ejecutar `ConstruirEsquema`. Verificar que se creen los agrupamientos de nivel 1 (Categoría) y nivel 2 (Detalle) sin alterar los totales acumulados.
4. **Prueba de PDF Dual:**
   - Ejecutar `GenerarPDFsCotizacion`.
   - Confirmar la creación de ambos archivos PDF en el directorio actual.
   - Verificar que `_Resumen.pdf` omita las filas marcadas con `"D"` y que `_Detalle.pdf` incluya la totalidad de las filas.

---

## 6. Solución de Problemas Frecuentes para Agentes

- **Error en llamada a `ole32.dll`:** Si el código se ejecuta en un entorno que no es Windows (ej. Mac OS o Linux), la API de Windows fallará. Notificar al usuario que la funcionalidad de UUID requiere Excel para Windows.
- **Pérdida de Reactividad en Eventos:** Si `Worksheet_Change` deja de responder, verificar si `Application.EnableEvents` quedó atascado en `False`. Ejecutar `Application.EnableEvents = True` en la ventana Inmediata de VBA (`Ctrl+G`).
- **Desalineación de Totales:** Si la macro `suma()` entrega un valor cero o incorrecto, verificar que los rótulos `"SUMA TOTAL"` en la columna `A` y `"SUBTOTAL"`, `"IVA"`, `"TOTAL"` en la columna `I` no contengan faltas de ortografía o espacios invisibles.
