# Evaluación Técnica y Plan de Mejora Integral

## Excel VBA Automation Toolkit — Cotizaciones Azteca Estudios

**Destinatario:** Equipo de Proyecto / Cliente Final  
**Objetivo del Plan:** Adecuar la solución para un usuario final **neófito** (sin conocimientos técnicos de VBA ni Excel avanzado), eliminando riesgos de fallas, pérdida de datos o pantallas de error técnico, y convirtiendo la plantilla en una herramienta **"a prueba de balas"** e intuitiva.

---

## 1. Diagnóstico Ejecutivo

El conjunto de macros actual cuenta con una base arquitectónica sólida en su versión modular (`M_UUID`, `M_Esquema`, `M_PDF`, `M_Blindaje`). Resuelve problemas complejos como la estabilidad de UUIDs, anclaje de botones e inconsistencia entre resumen y detalle en PDFs.

Sin embargo, **desde la perspectiva de un cliente sin conocimientos técnicos**, el sistema presenta **fricciones severas**:

1. **Riesgo de instalación:** El procedimiento de actualización requiere que el usuario abra el editor de VBA (`Alt + F11`), borre módulos, importe archivos `.bas`, pegue eventos en objetos del sistema y parchee código a mano. Para un usuario no técnico, esto es un punto de falla garantizado.
2. **Dependencia ciega de la celda activa (`ActiveCell`):** Las macros de inserción y borrado de filas no validan en qué parte de la hoja se encuentra el cursor. Si el usuario presiona un botón teniendo el cursor en los encabezados o en los totales, la estructura se corrompe.
3. **Falta de confirmación y borrado accidental:** `LimpiarCotizacion` no solicita confirmación previa y puede eliminar accidentalmente el trabajo de horas con un solo clic.
4. **Manejo de errores ausente en exportación:** Si el usuario tiene abierto el PDF de una cotización anterior e intenta generar uno nuevo, Excel arroja un error técnico `1004` en tiempo de ejecución con opción de "Depurar", dejando al usuario atrapado en el editor de código.
5. **Cálculos manuales y no reactivos:** La macro `suma()` requiere ser ejecutada manualmente y depende de coincidencias exactas de texto (`SUMA TOTAL`). Si el usuario cambia un valor y olvida recalcular, el PDF se genera con valores desactualizados.

---

## 2. Matriz de Fricciones Detectadas

| ID | Área / Macro | Fricción para el Usuario Neófito | Nivel de Riesgo | Impacto Real |
| --- | --- | --- | --- | --- |
| **F-01** | **Instalación / Despliegue** | Se le exige abrir `Alt+F11`, importar `.bas` y modificar código a mano en 5 macros. | 🔴 **CRÍTICO** | Frustración inmediata, imposibilidad de actualizar y riesgo de romper el libro. |
| **F-02** | **Acciones de Fila** (`AgregarFila*`, `BorrarFila`) | Se ejecutan en `ActiveCell.Row` sin validar si la celda activa está dentro de un bloque válido de conceptos. | 🔴 **CRÍTICO** | Inserción de filas dentro de totales o logos; borrado accidental de filas estructurales. |
| **F-03** | **Limpieza** (`limpiar_cotizacion.vb`) | Un clic involuntario borra la cotización sin advertencia (`MsgBox vbYesNo`). Además tiene errata ortográfica. | 🔴 **CRÍTICO** | Pérdida irreparable de cotizaciones en captura sin opción de deshacer (`Ctrl + Z` no sirve tras macros). |
| **F-04** | **PDFs** (`M_PDF`) | Si el archivo PDF ya está abierto en el lector de PDFs (Acrobat/Edge), Excel lanza error técnico 1004. | 🟠 **ALTO** | Pantalla de depuración VBA; el cliente piensa que el sistema se rompió. |
| **F-05** | **Recálculo** (`suma.vb`) | Los totales son estáticos (no fórmulas nativas). Si se añade una fila y no se oprime "Calcular", los montos quedan obsoletos. | 🟠 **ALTO** | Cotizaciones enviadas a clientes con sumas aritméticamente falsas. |
| **F-06** | **Nombre del Proyecto** (`M_UUID`) | Corregir una errata tipográfica en `C4` cambia el UUID automáticamente, perdiendo la trazabilidad. | 🟠 **ALTO** | Cotización cambia de folio inesperadamente al corregir ortografía del proyecto. |
| **F-07** | **Esquema de Detalle** (`M_Esquema`) | Marcar filas de detalle con "D" en columna oculta O es un concepto abstracto y confuso para un usuario básico. | 🟡 **MEDIO** | Usuario no entiende por qué ciertas filas no salen en el resumen o se colapsan mal. |
| **F-08** | **Duplicación de Bloque** (`duplicar_bloque.vb`) | Siempre duplica el rango rígido `A7:N12` de la hoja Plantillas, independientemente de qué categoría quería el usuario. | 🟡 **MEDIO** | Confusión del usuario si quería duplicar otra categoría o bloque. |
| **F-09** | **Feedback y Notificaciones** | Mensajes técnicos o falta de guías visuales paso a paso (barra de proceso, asistente de llenado). | 🟡 **MEDIO** | Incertidumbre sobre si el proceso terminó con éxito o si falta algún dato obligatorio. |

---

## 3. Análisis Técnico Detallado por Componente

### 3.1. Macros de Estructura de Filas (`agregar_fila_*.vb`, `borrar_fila_seleccionada.vb`)

* **Problema:** Emplean `ActiveCell.Row` directo:

  ```vba
  filaDestino = ActiveCell.Row + 1
  Rows(filaDestino).Insert Shift:=xlDown
  ```

* **Fricción:** Si el usuario tiene el cursor en la celda `C4` (Nombre del Proyecto) o en la fila de `SUBTOTAL`, la fila se insertará en el lugar incorrecto, descuadrando encabezados o fórmulas.
* **Solución requerida:** Crear una función guardiana `ValidarFilaEnBloque(fila As Long) As Boolean` que verifique que la fila activa esté comprendida estrictamente entre el primer bloque de conceptos y antes del área de totales finales. Si no lo está, mostrar un mensaje amigable: *"Por favor, selecciona una fila dentro de la tabla de conceptos para insertar o borrar."*

### 3.2. Macro de Limpieza (`limpiar_cotizacion.vb`)

* **Problemas:**
  1. No solicita confirmación al usuario.
  2. Borra `Cells(i, "A").ClearContents` de forma indiscriminada si no dice `"SUMA TOTAL:"`. Si el usuario puso un título de bloque o texto en columna A sin dos puntos, lo borra.
  3. Mensaje final con carácter corrupto: `"La cotización se limpió correctamente ?"`.
  4. No recalcula los totales (`Call suma`), dejando subtotales y totales congelados con los números anteriores.
* **Solución requerida:**
  * Agregar cuadro de confirmación modal: *"¿Estás seguro de que deseas limpiar los datos de esta cotización? Esta acción no se puede deshacer."*
  * Limpiar únicamente las celdas de captura de datos (cantidades, descripciones, precios unitarios) sin tocar la columna A ni los encabezados.
  * Ejecutar automáticamente `suma()` para que los totales queden en $0.00.

### 3.3. Cálculo Matemático y Totales (`suma.vb`)

* **Problemas:**
  1. Si el usuario agrega un concepto o edita un precio y va directo al botón "Generar PDFs", los PDFs se generarán con los totales viejos porque `GenerarPDFsCotizacion` no invoca a `suma`.
  2. Búsqueda dependiente de mayúsculas y textos literales sin normalizar: `"SUMA TOTAL"` o `"SUMA TOTAL:"`.
  3. El IVA está quemado al 16% en código (`subtotalGeneral * 0.16`). Si la tasa cambia o la celda en la hoja indica otro porcentaje, no se respeta.
* **Solución requerida:**
  * Invocar automáticamente `Call suma` dentro de `GenerarPDFsCotizacion` y tras cualquier inserción/borrado de filas.
  * Normalizar textos con `Replace(Replace(UCase(Trim(texto)), ":", ""), " ", "")` para tolerar variaciones accidentales de escritura.

### 3.4. Exportación a PDF (`generar_pdfs_resumen_detalle.vb`)

* **Problemas:**
  1. Si el archivo PDF ya existe y está abierto por el usuario o en su visor web, `ws.ExportAsFixedFormat` falla con `Error 1004` irrecuperable.
  2. Si el libro de Excel no ha sido guardado en disco (`ThisWorkbook.Path = ""`), la ruta de destino es inválida.
  3. No se abre la carpeta contenedora al finalizar, obligando al usuario a navegar por el explorador de Windows para buscar dónde quedaron los archivos.
* **Solución requerida:**
  * Validar si el libro tiene ruta (`If Len(ThisWorkbook.Path) = 0 Then MsgBox "Guarda el libro primero...", vbExclamation: Exit Sub`).
  * Implementar rutina de verificación o captura de error para PDF bloqueado: *"El archivo [Nombre.pdf] está abierto en otro programa. Por favor ciérralo y vuelve a intentar."*
  * Ofrecer al terminar abrir directamente la carpeta con `Shell "explorer.exe /select,""" & rutaArchivo & """", vbNormalFocus`.

### 3.5. Gestión de UUID e Identidad (`asignar_uuid_cotizacion.vb` / `eventos_uuid.vb`)

* **Problemas:**
  1. Si el usuario modifica el campo `C4` (Nombre del Proyecto), el UUID anterior se descarta y se genera uno nuevo de inmediato. Si fue un simple ajuste ortográfico, la cotización pierde su identificador original.
  2. La prohibición de nuevas hojas (`Workbook_NewSheet`) es drástica: borra la pestaña agregada inmediatamente. Aunque el principio de "un archivo = una cotización" es correcto, un usuario neófito puede entrar en pánico pensando que el programa "se volvió loco" o le borró algo que no debía.
* **Solución requerida:**
  * Si `C4` ya tenía un proyecto y un UUID asignado en `Meta!B2`, preguntar antes de reemplazar el UUID: *"El proyecto ha cambiado. ¿Deseas generar un nuevo ID de cotización o conservar el actual?"*
  * Para nuevas hojas, asegurar que el mensaje preventivo sea claro, pedagógico y sin jerga técnica.

---

## 4. Qué le Falta a la Macro (Funcionalidades Pendientes)

Para que la experiencia del cliente sea óptima y sin fricciones, faltan los siguientes elementos clave:

1. **Plantilla Maestra Empaquetada ("Llave en Mano"):**
   * El cliente **no debe editar código ni importar módulos**. Se le debe entregar el archivo `Plantilla Cotización AE 1.xlsm` con el 100% de los módulos ya integrados, probados y protegidos.
2. **Validación Contextual de Celda Activa:**
   * Filtro inteligente en todos los botones para que solo actúen si el usuario está ubicado en una zona válida de captura.
3. **Manejo Robusto de Excepciones en PDF:**
   * Detección de archivos en uso y enlace directo para abrir los PDFs generados.
4. **Recálculo Automático Integrado:**
   * Actualización de totales antes de cualquier exportación a PDF y después de cualquier modificación de renglones.
5. **Panel de Ayuda / Botón de Instrucciones Integrado en la Hoja:**
   * Un botón visible en Excel: `[ ? Guía Rápida ]` que abra una ventana simple explicando en 4 pasos cómo cotizar, sin tener que buscar archivos externos.
6. **Deshacer Seguro y Confirmaciones:**
   * Mensajes de confirmación con opciones `[Sí] / [No]` en todas las operaciones destructivas (limpiar hoja, borrar fila).

---

## 5. Plan de Acción y Hoja de Ruta (Roadmap)

### Fase 1: Blindaje Inmediato del Código (Técnica)

* [x] **Actualizar `generar_pdfs_resumen_detalle.vb` (y `bas/`):**
  * Manejo de errores implementado al exportar para capturar PDF bloqueado (Error 1004) y mostrar advertencia amigable sin pantallas de depuración.
  * Validación preventiva de `ThisWorkbook.Path` para evitar errores si el libro no ha sido guardado.
  * Invocación automática de `suma()` antes de exportar para sincronizar los totales.
  * Integración de cuadro de diálogo para abrir la carpeta contenedora con los PDFs generados al finalizar.
* [x] **Actualizar `limpiar_cotizacion.vb`:**
  * Confirmación modal `MsgBox vbYesNo` antes de limpiar.
  * Protección total de la columna A y filas de encabezados. Limpieza acotada a columnas de captura de datos (B:L y O).
  * Llamada automática a `suma()` al finalizar para dejar los totales en $0.00.
  * Errata ortográfica corregida en el mensaje de éxito.
* [x] **Actualizar `agregar_fila_*.vb` y `borrar_fila_seleccionada.vb`:**
  * Función guardiana `ValidarFilaEnBloque` creada en `M_Blindaje` e integrada en todas las macros de inserción y borrado.
  * Protección de filas de encabezados y sumas en `borrar_fila_seleccionada.vb`.
  * Recálculo automático de `suma()` tras agregar o borrar filas.
* [x] **Sincronizar exportaciones `.bas`:**
  * Sincronizados todos los módulos estándar en la carpeta `bas/` (`M_UUID`, `M_PDF`, `M_Blindaje`, `M_Esquema`, `macro_generadora_uuid`) con codificación Windows-1252 (cp1252) y saltos de línea CRLF.
* [x] **Actualizar `suma.vb`:**
  * Normalización de cadenas de búsqueda para `"SUMATOTAL"`, `"SUBTOTAL"`, `"IVA"`, `"TOTAL"`.
* [x] **Actualizar `asignar_uuid_cotizacion.vb` (y `bas/`):**
  * Diálogo de confirmación antes de generar nuevo UUID si cambia el nombre del proyecto en `C4`, preservando el folio en caso de corrección ortográfica.
* [x] **Actualizar `duplicar_bloque.vb`:**
  * Validación de fila activa con `ValidarFilaEnBloque` y recálculo automático con `suma()`.

### Fase 2: Experiencia de Usuario y Ergonomía (UX)

* [ ] **Diseño del Panel de Botones:**
  * Agrupar los botones lógicamente por pasos:
    * **Paso 1: Armado** (`+ Producto`, `+ Fecha`, `- Borrar fila`).
    * **Paso 2: Formato** (`Esquema / Colapsar`, `Marcar detalle`).
    * **Paso 3: Cierre** (`Calcular totales`, `Generar PDFs`).
    * **Utilidades** (`Limpiar cotización`, `Reacomodar botones`).
  * Colores distintivos (ej. Verde para PDFs, Azul para agregar filas, Rojo tenue para borrar).
* [x] **Apertura Asistida de PDFs:**
  * Al generar los PDFs con éxito, se muestra cuadro modal: *"¿Deseas abrir la carpeta con los PDFs generados?"*, abriendo el Explorador de Windows con el archivo seleccionado si el usuario acepta.

### Fase 3: Empaquetado y Entrega al Cliente (Despliegue)

* [ ] **Consolidar el `.xlsm` Oficial:**
  * Tomar la plantilla limpia, inyectar los módulos corregidos, compilar el proyecto VBA (`Depuración > Compilar VBAProject`).
  * Ejecutar `BlindarLibro` una sola vez para fijar posiciones definitivas.
  * Proteger la estructura y guardar el archivo maestro listo para producción.
* [ ] **Instrucciones para Bloqueo de Windows (Mark of the Web):**
  * Añadir una tarjeta visual de 1 página: cómo hacer clic derecho en el archivo `.xlsm` > *Propiedades* > marcar *Desbloquear* (Unblock) en caso de que Windows bloquee las macros por seguridad.

---

## 6. Checklist de Aceptación para el Cliente

Antes de liberar la versión al cliente final, verificar:
* [x] **Prueba de usuario distraído:** Hacer clic en "Agregar producto" teniendo seleccionada la celda del logotipo o del Total. *(Debe mostrar advertencia amigable sin modificar la hoja)*.
* [x] **Prueba de PDF abierto:** Abrir el PDF de cotización en Edge o Acrobat y hacer clic en "Generar PDFs". *(Debe avisar que se cierre el archivo sin abrir la pantalla de depuración de VBA)*.
* [x] **Prueba de borrado accidental:** Hacer clic en "Limpiar cotización" y seleccionar "No". *(No debe borrarse ningún dato)*.
* [x] **Prueba de cálculo desfasado:** Cambiar un precio y generar PDFs directamente. *(El total del PDF debe coincidir con el precio recién cambiado)*.
* [x] **Prueba de apertura limpia:** Abrir el archivo en otra máquina sin editor de VBA abierto. *(Todo debe funcionar con un clic en los botones)*.
