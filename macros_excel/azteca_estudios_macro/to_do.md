# Registro de Tareas y Cambios — Azteca Estudios Macro

## [2026-11-09] Cierre — Documentación, nombre final y limpieza para Git

- **README:** se agregó `Diana Berumen Estrada` en Autores y una sección de versión actual con módulos, botones y versiones.
- **AGENTS/instrucciones:** se documentaron los módulos actuales, el comportamiento del encabezado, los PDFs y la nota de módulos legacy conservados.
- **Archivo final:** `Plantilla Cotización AE.xlsm` (fuera de Git).
- **PDFs generados:** fuera de Git; la guía actualizada `GUÍA DE USO COTIZACIÓN.pdf` se conserva.
- **Legacy:** no se elimina ningún módulo `.vb` legacy.
- **Git:** se dejan preparadas las instrucciones; no se hace commit ni push.

## [2026-11-09] Ajuste — Comentarios de Fecha/ID hacia la izquierda

- **Problema:** los comentarios de Fecha e ID Cotización se abrían a la derecha del encabezado y se cortaban con el borde de la pantalla.
- **Corrección:**
  - Los comentarios de `K3` y `K4` se colocan a la izquierda de su celda.
  - Se mantienen 360×180 y texto con saltos de línea.
  - `H3`/`H4` siguen usando mensajes de validación de listas.
  - `VersionMBasePlantilla = 2026-11-09-COMMENTS-LEFT`.
- **Respaldo:** `Respaldos\Plantilla Cotización AE FUNCIONAL_FINAL_COMMENTS_LEFT_2026-11-09_012643.xlsm`.
- **Archivos actualizados:**
  - `base_plantilla.vb`
  - `bas/M_BasePlantilla.bas`
  - `to_do.md`
- **Pendiente:** reimportar `M_BasePlantilla` y verificar que Fecha/ID se vean completos al pasar el mouse.

## [2026-11-09] Ajuste — Ayuda mixta: comentarios y avisos de lista

- **Aclaración:** los campos con lista desplegable deben usar mensajes de validación; los campos de captura normal deben usar comentarios.
- **Comportamiento final:**
  - `H3` ID Cliente y `H4` ID Proyecto: `ShowInput=True`, sin comentario; el mensaje amarillo aparece al desplegar la lista.
  - `C3` Cliente, `C4` Proyecto, `K3` Fecha y `K4` ID Cotización: comentario con triángulo rojo y `ShowInput=False`.
  - `AplicarComentariosHeader` aplica el comportamiento mixto después de **Nueva cotización**.
- **Versión:** `VersionMBasePlantilla = 2026-11-09-MIX-COMMENTS-INPUT2`.
- **Comentarios largos:** C4, K3 y K4 ahora llevan saltos de línea manuales para que el texto no se corte.
- **Archivos actualizados:**
  - `base_plantilla.vb`
  - `bas/M_BasePlantilla.bas`
  - `to_do.md`
- **Pendiente:** reimportar `M_BasePlantilla`, compilar y probar las listas H3/H4 y los comentarios C3/C4/K3/K4.

## [2026-11-08] Corrección — Comentarios de Fecha/ID con tamaño fijo

- **Problema:** los comentarios de Fecha/ID seguían cortados; el ajuste anterior con `AutoSize=True` los redujo a una sola línea de ~25 pt de alto.
- **Corrección:**
  - `AutoSize=False`.
  - Tamaño fijo 360×180 para todos los comentarios del encabezado.
  - Se actualizaron los comentarios actuales y los de `_BaseAE`.
  - `VersionMBasePlantilla = 2026-11-08-COMMENTS-SIZE2`.
- **Respaldo:** `Respaldos\Plantilla Cotización AE FUNCIONAL_FINAL_COMMENTS_SIZE2_2026-11-08_235344.xlsm`.
- **Archivos actualizados:**
  - `base_plantilla.vb`
  - `bas/M_BasePlantilla.bas`
  - `to_do.md`
- **Pendiente:** reimportar `M_BasePlantilla` y verificar que el texto completo se vea en los comentarios.

## [2026-11-08] Ajuste — Tamaño de comentarios de Fecha/ID

- **Problema:** los comentarios amarillos de Fecha e ID Cotización mostraban el texto cortado/partido.
- **Causa:** el cuadro de comentario medía 280×120 y no alcanzaba para el texto largo.
- **Corrección:**
  - Todos los comentarios del encabezado ahora usan 360×180 y `TextFrame.AutoSize = True`.
  - Se actualizaron los comentarios actuales y los de `_BaseAE`.
  - `VersionMBasePlantilla = 2026-11-08-COMMENTS-SIZE`.
- **Respaldo:** `Respaldos\Plantilla Cotización AE FUNCIONAL_FINAL_COMMENTS_SIZE_2026-11-08_234226.xlsm`.
- **Archivos actualizados:**
  - `base_plantilla.vb`
  - `bas/M_BasePlantilla.bas`
  - `to_do.md`
- **Pendiente:** reimportar `M_BasePlantilla` y verificar que los comentarios de Fecha/ID se vean completos.

## [2026-11-08] Cambio — Ayuda del encabezado como comentarios de celda

- **Ajuste solicitado:** no ocultar la ayuda amarilla; mostrarla correctamente anclada a cada celda.
- **Solución:** los mensajes de validación se convirtieron en comentarios/notas de celda:
  - `C3` Cliente
  - `C4` Proyecto
  - `K3` Fecha
  - `K4` ID Cotización
  - `H3` ID Cliente
  - `H4` ID Proyecto
- Cada celda muestra un triángulo rojo; al pasar el mouse aparece el cuadro amarillo. No se muestran automáticamente al editar ni tapan `C:D`.
- `ShowInput` permanece desactivado para evitar los cuadros automáticos.
- Se agregó `ConfigurarComentariosHeader` en `M_BasePlantilla`; `RestaurarBasePlantilla` vuelve a crear los comentarios.
- **Versión:** `VersionMBasePlantilla = 2026-11-08-COMMENTS-HEADER`.
- **Respaldo:** `Respaldos\Plantilla Cotización AE FUNCIONAL_FINAL_COMMENTS_2026-11-08_221643.xlsm`.
- **Archivos actualizados:**
  - `base_plantilla.vb`
  - `bas/M_BasePlantilla.bas`
  - `to_do.md`
- **Pendiente:** reimportar `M_BasePlantilla` y comprobar el comportamiento de los comentarios al pasar el mouse y tras Nueva cotización.

## [2026-11-08] Corrección — Avisos de validación sobre el encabezado

- **Problema:** al editar Cliente, Proyecto u otros datos del encabezado, aparecían cuadros amarillos sobre `C:D` y otros campos.
- **Causa:** no eran comentarios; eran mensajes de entrada de las validaciones de datos (`ShowInput = True`) en `C3`, `C4`, `K3`, `K4`, `H3` y `H4`.
- **Corrección:**
  - Se desactivó `ShowInput` en esas celdas de `Cotización` y `_BaseAE`.
  - Se conservan las reglas de validación, listas y avisos de error.
  - Nueva cotización ya no restaura los cuadros amarillos.
- **Respaldo con el header corregido:** `Respaldos\Plantilla Cotización AE FUNCIONAL_FINAL_HEADER_2026-11-08_215943.xlsm`.
- **Pendiente:** probar edición de Cliente/Proyecto/Fecha/IDs y confirmar que no aparecen cuadros sobre `C:D`.

## [2026-11-08] Respaldo final y estado de módulos

- **Respaldo final:** `Respaldos\Plantilla Cotización AE FUNCIONAL_FINAL_2026-11-08_214327.xlsm` (324,873 bytes).
- **Origen:** copia de regresión activa, con los últimos ajustes probados.
- **Módulos actuales:** `M_BasePlantilla`, `M_Blindaje`, `M_Catalogos`, `M_Esquema`, `M_PDF`, `M_Producto`, `M_Recalculo`, `M_UUID`.
- **Módulos legacy:** `Módulo1` a `Módulo10`.
  - `Módulo9` contiene `suma` y está en uso.
  - `Módulo10` contiene `AgregarFilaFecha` y está en uso.
  - Los módulos que contienen `DuplicarBloque` y `BorrarFilaSeleccionada` también están en uso por los botones 3 y 5.
  - No eliminar `Módulo1` a `Módulo8` sin mapear primero sus procedimientos y referencias.
- **Pendiente:** inventariar procedimientos de `Módulo1` a `Módulo8`, eliminar solo los no referenciados y actualizar documentación.

## [2026-11-07] Corrección — Declaración duplicada de `i` en M_PDF

- **Problema:** al compilar `M_PDF` aparecía “Declaración duplicada en el ámbito actual”.
- **Causa:** `GenerarPDFsCotizacion` declaraba `Dim i As Long` al inicio y otra vez en `Dim ultima As Long, i As Long`.
- **Corrección:** se eliminó la segunda declaración de `i`.
- **Versión:** `VersionMPDF = 2026-11-07-ABRE-DETALLE-FIX`.
- **Archivos actualizados:**
  - `generar_pdfs_resumen_detalle.vb`
  - `bas/generar_pdfs_resumen_detalle.bas`
  - `to_do.md`
- **Pendiente:** reimportar `M_PDF`, compilar y probar la exportación con apertura automática del Detalle.

## [2026-11-07] Ajuste — Abrir PDF Detalle al terminar la exportación

- **Cambio:** al finalizar `GenerarPDFsCotizacion`, se abre automáticamente el PDF `_Detalle.pdf` en el visor predeterminado.
- **Mensaje:** se mantiene la confirmación con los nombres de ambos PDFs y la ruta queda en la carpeta del libro.
- **Respaldo:** si `FollowHyperlink` falla, se selecciona el archivo en Explorer.
- **Verificación:** `VersionMPDF = 2026-11-07-ABRE-DETALLE`.
- **Archivos actualizados:**
  - `generar_pdfs_resumen_detalle.vb`
  - `bas/generar_pdfs_resumen_detalle.bas`
  - `to_do.md`
- **Pendiente:** reimportar `M_PDF` y probar “Generar PDFs”; debe abrir el Detalle.

## [2026-11-07] Corrección — Pie de página con marcas D residuales tras Nueva cotización

- **Problema:** después de **Nueva cotización**, las filas del pie `Elaboró` y `Especificaciones` se veían rosadas y con formato alterado.
- **Causa:** `RestaurarBasePlantilla` copiaba `A1:N` desde `_BaseAE`, pero no limpiaba la columna `O`; quedaban marcas `D` de la cotización anterior y `ResaltarDetalle` pintaba esas filas.
- **Corrección:**
  - Después de restaurar y ajustar altos, se limpia `O1:O<última fila base>`.
  - `ResaltarDetalle` vuelve a aplicar solo las marcas válidas.
  - Se agregó `VersionMBasePlantilla = 2026-11-07-FIX-O-FOOTER`.
- **Archivos actualizados:**
  - `base_plantilla.vb`
  - `bas/M_BasePlantilla.bas`
  - `to_do.md`
- **Pendiente:** reimportar `M_BasePlantilla`, ejecutar Nueva cotización y verificar el pie.

## [2026-11-07] Corrección — Validación de fecha heredada en filas de producto/foro

- **Problema:** después de insertar producto/foro debajo de una fecha, la fila nueva conservaba la validación de fecha en `A`; al escribir descripción o foro aparecía “Este valor no coincide con las restricciones de validación de datos”.
- **Causa:** `Rows.Insert` hereda la validación de la fila superior y `PrepararFilaNueva` no la eliminaba.
- **Corrección:**
  - `PrepararFilaNueva` elimina la validación de `A` en la fila nueva antes de copiar formatos.
  - Aplica para producto y foro insertados debajo de una fecha.
  - `VersionMProducto` ahora es `2026-11-07-FIX-VALIDACION`.
- **Archivos actualizados:**
  - `producto_seguro.vb`
  - `bas/M_Producto.bas`
  - `to_do.md`
- **Pendiente:** reimportar `M_Producto`, borrar la fila con error, y probar producto/foro debajo de fecha.

## [2026-11-07] Corrección — Desfase de fila modelo al insertar arriba de ella

- **Problema:** con `M_Producto` y `M_Catalogos` actualizados, el producto/foro debajo de una fecha seguía tomando el formato de la fecha.
- **Causa:** cuando el producto modelo estaba **debajo** del punto de inserción, `Rows.Insert` lo desplazaba una fila y la macro seguía usando el número de fila anterior; terminaba copiando formato de otra fila.
- **Corrección:**
  - Después de insertar, si `filaModelo >= filaNueva`, se incrementa `filaModelo` en 1.
  - Se aplica en `AgregarProducto` y `AgregarEspacioForo`.
  - Versión de ambos módulos: `2026-11-07-FIX-MODELO-SHIFT`.
- **Archivos actualizados:**
  - `producto_seguro.vb`
  - `catalogo_categorias.vb`
  - `bas/M_Producto.bas`
  - `bas/M_Catalogos.bas`
  - `to_do.md`
- **Pendiente:** reimportar ambos módulos, verificar las dos versiones, borrar filas grises y repetir producto/foro debajo de fecha.

## [2026-11-07] Verificación — M_Catalogos actualizado para foro debajo de fecha

- **Problema:** producto y foro debajo de una fecha seguían tomando formato gris; `VersionMProducto` ya devolvía la versión correcta, pero el foro se inserta desde `M_Catalogos` y ese módulo podía seguir con una versión anterior.
- **Corrección/verificación:**
  - Se agregó `VersionMCatalogos` con el valor `2026-11-07-FIX-MODELO-FECHA`.
  - `AgregarEspacioForo` usa `FilaModeloProducto` y `PrepararFilaNueva` con la fila modelo válida.
  - Se regeneró `bas/M_Catalogos.bas`.
- **Pendiente:** eliminar cualquier `M_Catalogos`/`M_Catalogos1` del proyecto, importar limpio `bas/M_Catalogos.bas`, compilar y verificar `?VersionMCatalogos`.

## [2026-11-07] Corrección — Evitar que una fila de fecha sea modelo de producto

- **Problema:** al insertar producto/foro debajo de una fecha en un bloque sin productos, la fila nueva tomaba el formato gris de la fecha.
- **Causa probable:** la fila de fecha o una fila mal formada del bloque se estaba usando como modelo porque tenía fórmulas, aunque no tenía formato de producto.
- **Corrección:**
  - `EsFilaProductoModelo` ahora exige fórmulas en `K`/`M` **y** formato de moneda/numérico en `G`.
  - Las filas de fecha y filas mal formadas con formato General quedan excluidas como modelo.
  - Se mantiene el respaldo a un producto válido de otro bloque.
  - Se agregó `VersionMProducto` para verificar que el módulo actualizado esté cargado.
- **Archivos actualizados:**
  - `producto_seguro.vb`
  - `bas/M_Producto.bas`
  - `to_do.md`
- **Pendiente:** verificar `?VersionMProducto`, reimportar limpio `M_Producto`, borrar las filas grises de prueba y repetir producto/foro debajo de fecha.

## [2026-11-07] Corrección — Bloque con solo fecha y sin productos modelo

- **Problema:** al borrar todos los productos de un bloque y dejar solo la fecha, producto y foro se rechazaban porque no había fila modelo dentro del bloque.
- **Corrección:**
  - `FilaModeloProducto` mantiene la búsqueda local.
  - Si no encuentra producto en el bloque, busca en toda la hoja un producto con fórmulas en `K`/`M`.
  - Usa ese producto como modelo de formatos, combinaciones, alto y fórmulas.
- **Archivos actualizados:**
  - `producto_seguro.vb`
  - `bas/M_Producto.bas`
  - `to_do.md`
- **Pendiente:** reimportar `M_Producto` y probar un bloque con solo fecha usando un producto de otro bloque como modelo.

## [2026-11-07] Corrección — Formato G:M al insertar producto/foro debajo de fecha

- **Problema:** al insertar un producto debajo de una fila de fecha, la fila nueva conservaba fórmulas pero perdía los formatos de `G:M`; Excel heredaba el formato General de la fecha.
- **Causa:** `PrepararFilaNueva` copiaba formato de `A`, alto y bordes, pero no los formatos de las columnas `G:M` desde la fila modelo de producto.
- **Corrección:**
  - Se agregó `CopiarFormatoColumnasProducto` en `M_Producto`.
  - Copia desde la fila modelo a la fila nueva: `NumberFormat`, fuente, relleno, alineación, `WrapText` e `IndentLevel` para `G:M`.
  - Se llama desde `PrepararFilaNueva` antes de aplicar bordes.
  - Aplica tanto a producto como a foro insertados debajo de una fecha.
- **Archivos actualizados:**
  - `producto_seguro.vb`
  - `bas/M_Producto.bas`
  - `to_do.md`
- **Pendiente:** reimportar `M_Producto`, borrar las filas de prueba `14` y `15` de la copia de regresión, y repetir la inserción debajo de la fecha.

## [2026-11-07] Ajuste — Botones inferiores y alerta de Nueva cotización

- **Cambio visual:** `Reestablecer macro` quedó en la columna izquierda de la fila inferior.
- **Nueva cotización:** se movió a la columna derecha, alineada con `Reestablecer macro`.
- **Estilo unificado:** ambos botones usan `Avenir Next LT Pro`, tamaño 12, negrita. `Nueva cotización` conserva color rojo (`vbRed`).
- **Implementación:** `AcomodarBotonesInferiores` en `M_BasePlantilla` copia nombre, tamaño, negrita e itálica desde `Reestablecer macro` y aplica rojo solo a `Nueva cotización`.
- **Persistencia:** se actualizaron las posiciones de ambos botones en `Meta!D:H`; `RestaurarPosicionesBotones` ya no los mueve.
- **Verificación:** posiciones actuales `Reestablecer macro` en `L=982.5, T=551.6` y `Nueva cotización` en `L=1179.6, T=551.6`; ambas con `Avenir Next LT Pro 12` y negrita.
- **Archivos actualizados:**
  - `base_plantilla.vb`
  - `bas/M_BasePlantilla.bas`
  - `to_do.md`
- **Pendiente:** guardar el libro y probar `RestaurarPosicionesBotones`, cerrar/reabrir y ejecutar `Nueva cotización` (cancelar) para confirmar persistencia.

## [2026-11-07] Verificación — Limpieza M_Plantilla/M_LimpiarSeguro y Fase 6

- **Respaldo completo:** `Respaldos\Plantilla Cotización AE FUNCIONAL_ANTES_LIMPIEZA_2026-11-07.xlsm`.
- **Módulos viejos:** se revisó el proyecto VBA del respaldo; `M_Plantilla` y `M_LimpiarSeguro` **ya no existen como componentes**. No hay `.bas` que exportar; solo quedan cadenas en la caché compilada del proyecto, sin efecto.
- **Fase 6:** `M_Recalculo` ya está importado en el proyecto.
- **Prueba realizada:** se editó un producto de prueba con descripción y precio; sin ejecutar `suma` manualmente, el evento actualizó `SUMA TOTAL`, `SUBTOTAL`, `IVA` y `TOTAL`. Después se restauraron los valores de prueba y se confirmó `Application.EnableEvents = True`.
- **Resultado:** el recálculo automático de Fase 6 está activo y funcionando.
- **Pendiente:** guardar el libro, cerrar/reabrir y hacer la prueba final con datos reales.

## [2026-11-07] Corrección — Categoría editable con validación nativa

- **Problema:** el primer bloque abría un combo temporal de categorías que no desplegaba bien la lista y no permitía escribir; el segundo bloque sí permitía editar la celda.
- **Corrección:**
  - Se quitó del evento `Worksheet_BeforeDoubleClick` la rama que abría el combo temporal de categoría.
  - Las categorías usan directamente la validación nativa de `A:N`: se puede escribir o elegir de la lista con la flecha del extremo derecho.
  - El doble clic en filas de foro conserva su combo.
- **Archivos actualizados:**
  - `eventos_uuid.vb`
  - `to_do.md`
- **Pendiente en el libro:** reemplazar `Worksheet_BeforeDoubleClick` en la hoja `Cotización`; eliminar el control oculto `Drop Down 61` desde el Panel de selección; compilar y probar ambos bloques.

## [2026-11-07] Corrección — Producto/foro después de una fila de fecha

- **Problema:** después de agregar una fecha, los botones de producto y foro rechazaban la fila porque la fecha no tiene fórmulas en `K`/`M`.
- **Corrección:**
  - `M_Producto` agrega `EsFilaFecha` y `FilaModeloProducto`.
  - `EsFilaValidaParaProducto` y `EsFilaValidaParaForo` aceptan filas de fecha si existe una fila modelo de producto.
  - `AgregarProducto` y `AgregarEspacioForo` insertan debajo de la fecha, pero toman formato, combinaciones y fórmulas de la fila modelo.
  - `AgregarFilaFecha` selecciona la nueva fila de fecha después de insertarla.
- **Archivos actualizados:**
  - `producto_seguro.vb`
  - `catalogo_categorias.vb`
  - `agregar_fila_fecha.vb`
  - `bas/M_Producto.bas`
  - `bas/M_Catalogos.bas`
- **Estado:** implementado; pendiente reimportar `M_Producto`, `M_Catalogos` y actualizar `AgregarFilaFecha` en el libro; validar.

## [2026-11-06] Corrección — Vista dividida y botones aparentemente duplicados

- **Problema:** al agregar un foro, la ventana mostraba la hoja dividida en dos paneles y los botones parecían duplicados.
- **Causa:** la ventana quedó con una división (`SplitColumn`/`SplitRow`), no con botones duplicados; al dividir la vista, los mismos objetos se ven en ambos paneles.
- **Corrección:**
  - Se restauró la vista actual (`SplitColumn = 0`, `SplitRow = 0`).
  - Se agregó `RestaurarVistaCotizacion` en `M_Catalogos`.
  - `AgregarEspacioForo` y `AplicarForoCatalogo` ahora restauran la división de la ventana.
- **Archivos actualizados:**
  - `catalogo_categorias.vb`
  - `bas/M_Catalogos.bas`
- **Estado:** implementado; pendiente reimportar `M_Catalogos` y validar.

## [2026-11-06] Corrección — Ancho del combo de foro en celda A:F

- **Problema:** al agregar un foro, el combo medía 160 pt y solo cubría la mitad izquierda de la celda combinada A:F; la celda se veía partida.
- **Corrección:**
  - `CrearComboCatalogo` ahora recibe el ancho del combo.
  - El combo de foro usa el ancho real de `celda.MergeArea.Width - 4`.
  - Se posiciona en `celda.MergeArea.Left + 2`.
  - El combo de categorías conserva su ancho de 160 pt.
- **Archivos actualizados:**
  - `catalogo_categorias.vb`
  - `bas/M_Catalogos.bas`
- **Estado:** implementado; pendiente reimportar `M_Catalogos` y validar que el combo cubra A:F.

## [2026-11-06] Corrección — Salto de vista al agregar foro

- **Problema:** al agregar un foro, la hoja se desplazaba y la fila nueva se veía cortada.
- **Causa:** `ActiveWindow.ScrollRow = filaNueva - 4` forzaba la vista a mover la fila nueva hacia el borde superior.
- **Corrección:**
  - Se eliminó el `ScrollRow` forzado.
  - Ahora solo se desplaza la vista si la fila nueva queda fuera de la ventana visible.
  - Si la fila ya se ve, la hoja no se mueve.
- **Archivos actualizados:**
  - `catalogo_categorias.vb`
  - `bas/M_Catalogos.bas`
- **Estado:** implementado; pendiente reimportar `M_Catalogos` y validar.

## [2026-11-06] Corrección — PasteSpecial al copiar fórmulas K:M

- **Problema:** al agregar producto/foro aparecía:
  “No se pudo agregar … Error en el método PasteSpecial de la clase Range.”
- **Causa:** `K:M` incluye la celda `M`, que ahora está combinada con `N`. `PasteSpecial` falla al pegar en un rango que contiene celdas combinadas.
- **Corrección:**
  - Se agregó `CopiarFormulasProducto`, que copia las fórmulas de `K` y `M` directamente con `FormulaR1C1`.
  - Ya no se usa `PasteSpecial` para las fórmulas.
  - Se aplica en `AgregarProducto` y `AgregarEspacioForo`.
- **Archivos actualizados:**
  - `producto_seguro.vb`
  - `bas/M_Producto.bas`
  - `catalogo_categorias.vb`
  - `bas/M_Catalogos.bas`
- **Estado:** implementado; pendiente reimportar ambos módulos y validar.

## [2026-11-06] Corrección — Bordes verticales y combinación M:N

- **Formato correcto identificado:**
  - `A:F` sin bordes.
  - `G:L` con borde izquierdo y derecho, continuo, delgado, `RGB(173, 173, 173)`.
  - `M` con borde izquierdo, mismo estilo, dentro de `M:N` combinado.
  - Sin bordes horizontales ni diagonales.
- **Corrección:**
  - `PrepararFilaNueva` ahora asegura las combinaciones horizontales de la fila activa (`A:F`, `M:N`, etc.).
  - Se agregó `AplicarBordesProducto`, que limpia los bordes heredados y aplica explícitamente el estilo anterior.
  - Ya no se copian bordes de la fila origen, por lo que no importa si la base o el respaldo no los tienen.
- **Archivos actualizados:**
  - `producto_seguro.vb`
  - `bas/M_Producto.bas`
- **Estado:** implementado; pendiente reimportar `M_Producto` y validar desde “Nueva cotización”.

## [2026-11-06] Corrección — Bordes diagonales en filas nuevas

- **Problema:** las filas nuevas mostraban líneas diagonales y bordes incorrectos.
- **Causa:** se copiaban los bordes `xlDiagonalDown`/`xlDiagonalUp` y los interiores del rango combinado.
- **Corrección:**
  - `A:F` copia solo bordes exteriores.
  - `G:O` copia exteriores + interiores vertical/horizontal.
  - Nunca se copian bordes diagonales.
- **Archivos actualizados:**
  - `producto_seguro.vb`
  - `bas/M_Producto.bas`
- **Estado:** implementado; pendiente reimportar `M_Producto` y validar.

## [2026-11-06] Corrección — Bordes en filas nuevas de producto/foro

- **Problema:** las filas nuevas ya abarcaban `A:F`, pero perdían los bordes de las celdas.
- **Corrección:**
  - `PrepararFilaNueva` ahora copia los bordes de la fila activa al rango `A:F` y a `G:O`.
  - Se copian `LineStyle`, `Weight`, `Color`, `ColorIndex`, `ThemeColor` y `TintAndShade`.
- **Archivo actualizado:**
  - `producto_seguro.vb`
  - `bas/M_Producto.bas`
- **Estado:** implementado; pendiente reimportar `M_Producto` y validar producto/foro.

## [2026-11-06] Corrección — Filas nuevas pierden la combinación A:F

- **Problema:** al insertar producto o foro, la fila nueva perdía la combinación `A:F` y el texto quedaba solo en la columna `A`.
- **Causa:** `Rows.Insert` hereda formatos, pero no copia las celdas combinadas de la fila superior.
- **Corrección:**
  - Se agregó `PrepararFilaNueva` en `M_Producto`.
  - Vuelve a combinar `A:F` cuando la fila activa tiene esa combinación.
  - Copia alto de fila, `WrapText`, alineación, sangría, formato de número, fuente y color de fondo.
  - Se usa desde `AgregarProducto` y desde `AgregarEspacioForo`.
- **Archivos actualizados:**
  - `producto_seguro.vb`
  - `bas/M_Producto.bas`
  - `catalogo_categorias.vb`
  - `bas/M_Catalogos.bas`
- **Estado:** implementado; pendiente reimportar ambos módulos y validar.

## [2026-11-06] Corrección — Error PasteSpecial en “Agregar producto”

- **Problema:** el botón “Agregar producto” lanzaba:
  “Error 1004: Error en el método PasteSpecial de la clase Range.”
- **Causa:** el macro legado copiaba la fila completa de `Plantillas` y pegaba formatos en una fila con combinaciones distintas.
- **Corrección:**
  - Nuevo módulo `M_Producto` con `AgregarProducto`.
  - Valida que la fila activa sea de producto (fórmulas en `K`/`M`, sin encabezados, SUMA TOTAL ni totales).
  - Inserta la fila heredando formato de la fila activa.
  - Copia solo las fórmulas `K:M`.
  - `AsignarBotonProducto` reasigna el botón “Agregar producto”.
- **Archivos creados:**
  - `producto_seguro.vb`
  - `bas/M_Producto.bas` (cp1252, CRLF, `Attribute VB_Name = "M_Producto"`)
- **Estado:** implementado; pendiente importar y reasignar el botón.

## [2026-11-06] Corrección — Error PasteSpecial al agregar espacio/foro

- **Problema:** al pulsar “Agregar espacio/foro” se mostraba:
  “No se pudo agregar el espacio/foro: Error en el método PasteSpecial de la clase Range.”
- **Causa:** se copiaba la fila completa de `Plantillas` y se pegaban formatos en una fila cuya combinación de celdas no coincidía (encabezado, SUMA TOTAL, totales o fila con otro tipo de combinación).
- **Corrección:**
  - Se agregó `EsFilaValidaParaForo`: solo permite filas de producto/foro con fórmulas en `K` o `M`, dentro del bloque y fuera de encabezados, SUMA TOTAL y totales.
  - La fila nueva ya no copia la fila de `Plantillas`; hereda formato/combinación de la fila activa.
  - Solo se copian las fórmulas de `K:M` desde la fila activa a la fila nueva.
  - El mensaje de error ahora incluye el número de fila.
- **Archivos actualizados:**
  - `catalogo_categorias.vb`
  - `bas/M_Catalogos.bas`
- **Estado:** implementado; pendiente reimportar `M_Catalogos` y probar.

## [2026-11-06] Limpieza de módulos y Fase 6 — Recálculo automático

### Limpieza definitiva
- **Archivos eliminados del repositorio:**
  - `nueva_cotizacion.vb`
  - `bas/M_Plantilla.bas`
  - `limpiar_datos_seguro.vb`
  - `bas/M_LimpiarSeguro.bas`
  - `bas/M_LimpiarSeguro_ACTUAL.bas`
  - `bas/M_Catalogos_ACTUAL.bas`
- **Pendiente en el libro `Plantilla Cotización AE FUNCIONAL.xlsm`:**
  - Quitar los módulos `M_Plantilla` y `M_LimpiarSeguro`.
  - Verificar que el botón “Nueva cotización” use `RestaurarBasePlantilla`.
  - Compilar y guardar.

### Fase 6 — Recálculo automático
- **Archivos creados:**
  - `recalculo_automatico.vb`
  - `bas/M_Recalculo.bas` (cp1252, CRLF, `Attribute VB_Name = "M_Recalculo"`)
- **Macros:**
  - `RecalcularTrasCambio(Target)`: recalcula con `suma` cuando se editan o vacían celdas en `A:M` u `O` dentro de la zona de conceptos.
  - `RecalcularTodo()`: recálculo manual desde cualquier hoja.
- **Evento actualizado:**
  - `eventos_uuid.vb`: `Worksheet_Change` ahora llama a `RecalcularTrasCambio Target`.
- **Estado:** implementado en archivos; pendiente importar `M_Recalculo`, reemplazar el evento de la hoja Cotización y validar.

## [2026-11-06] Nueva solución: hoja base oculta `_BaseAE`

- **Motivo:** la limpieza en sitio siguió borrando encabezados; se reemplaza por una restauración desde una copia limpia interna.
- **Archivos creados:**
  - `base_plantilla.vb`
  - `bas/M_BasePlantilla.bas` (cp1252, CRLF, `Attribute VB_Name = "M_BasePlantilla"`)
- **Macros:**
  - `CrearBasePlantilla`: copia la hoja Cotización a la hoja `_BaseAE`, la deja visible para limpieza manual y copia anchos/altos.
  - `OcultarBasePlantilla`: oculta `_BaseAE` como `xlSheetVeryHidden`.
  - `RestaurarBasePlantilla`: crea respaldo rotativo, guarda `C3/C4/K3/K4/N4/H3/H4`, copia `_BaseAE` sobre Cotización, restaura el encabezado y recalcula.
  - `AsignarBotonNuevaCotizacion`: repurpone el botón actual como “Nueva cotización” con `OnAction = RestaurarBasePlantilla`.
- **Ventaja:** no requiere archivo maestro externo; la plantilla limpia viaja dentro del `.xlsm`.
- **Pendiente:** importar `M_BasePlantilla`, crear/limpiar/ocultar `_BaseAE`, reasignar el botón y eliminar los módulos viejos `M_Plantilla` y `M_LimpiarSeguro`.

## [2026-11-06] Fase 5 — Aislamiento total y bloqueo de eventos

- **Hallazgo:** el respaldo rotativo demuestra que los encabezados existían antes de la limpieza; la nueva macro no debería tocarlos. Se sospecha que un evento `Worksheet_Change` dispara una limpieza antigua durante el borrado.
- **Corrección:**
  - `LimpiarDatosSeguro` ahora guarda y desactiva `Application.EnableEvents` durante toda la limpieza y lo restaura al salir.
  - Se mantiene el módulo aislado `M_LimpiarSeguro` con el botón apuntando a `LimpiarDatosSeguro`.
  - Se recomienda eliminar el módulo viejo `M_Plantilla`; si al compilar aparece “Sub o Function no definida: LimpiarDatosCotizacion”, algún evento lo está llamando y ahí está la causa.
- **Archivos actualizados:**
  - `limpiar_datos_seguro.vb`
  - `bas/M_LimpiarSeguro.bas`
- **Estado:** implementado; pendiente recuperar del respaldo, importar el módulo, eliminar `M_Plantilla`, compilar y probar en una copia.

## [2026-11-06] Fase 5 — Nuevo módulo aislado M_LimpiarSeguro

- **Motivo:** el botón seguía borrando encabezados, posiblemente por ejecutar una versión vieja del módulo.
- **Solución:** crear un módulo nuevo, con nombres nuevos, que no reutiliza nada de `M_Plantilla`.
- **Archivos creados:**
  - `limpiar_datos_seguro.vb`
  - `bas/M_LimpiarSeguro.bas` (cp1252, CRLF, `Attribute VB_Name = "M_LimpiarSeguro"`)
- **Macros:**
  - `LimpiarDatosSeguro`: exige `C4` no vacío, crea respaldo rotativo, limpia solo filas con fórmula en `K`/`M` o fecha en `A`, y solo las columnas `A`, `G:M` y `O`. No toca encabezados, categorías, filas `SUMA TOTAL` ni fórmulas.
  - `AsignarBotonLimpiarSeguro`: reasigna el botón existente “Limpiar datos” a `LimpiarDatosSeguro` o lo crea si no existe.
- **Ventaja:** al tener un nombre de macro distinto, Excel no puede ejecutar por error la versión vieja.
- **Estado:** implementado; pendiente recuperar del respaldo, importar `M_LimpiarSeguro`, ejecutar `AsignarBotonLimpiarSeguro` y probar sobre una copia.

## [2026-11-06] Fase 5 — Corrección con lista blanca de filas

- **Incidente:** una prueba de “Limpiar datos” volvió a borrar encabezados.
- **Recuperación:** el respaldo rotativo `Respaldos\Plantilla Cotización AE 1 2_ANTES_FASE5_ULTIMO_RESPALDO.xlsm` conserva el estado anterior con encabezados, datos y fórmulas.
- **Corrección:**
  - Se agregó `EsFilaProductoOFecha`: solo se limpian filas que tengan fórmulas en `K`/`M`, una fecha en `A`, o capturas numéricas en `G/H/I/J/L`.
  - Los encabezados de categoría, encabezados de tabla y filas `SUMA TOTAL` quedan excluidos.
  - Se agregó una guardia: si `C4` (Nombre del proyecto) está vacío, no se limpia nada.
  - La confirmación muestra la versión segura `2026-11-06-v5`.
- **Archivos actualizados:**
  - `nueva_cotizacion.vb`
  - `bas/M_Plantilla.bas`
- **Estado:** implementado; pendiente reimportar `M_Plantilla` y probar sobre una copia con el respaldo rotativo.

## [2026-11-06] Fase 5 — Respaldo rotativo antes de “Limpiar datos”

- **Objetivo:** proteger los datos antes de la limpieza sin llenar la carpeta de respaldos.
- **Implementación:**
  - `CrearRespaldoAntesDeLimpiar` crea la subcarpeta `Respaldos` si no existe.
  - Guarda un único archivo `<NombreDelLibro>_ULTIMO_RESPALDO.xlsm`.
  - Cada limpieza sobrescribe ese mismo archivo; no se acumulan copias.
  - Si el respaldo falla, la macro no limpia nada y muestra el error.
  - El mensaje final incluye la ruta del respaldo.
- **Archivos actualizados:**
  - `nueva_cotizacion.vb`
  - `bas/M_Plantilla.bas`
- **Estado:** implementado; pendiente importar `M_Plantilla` y validar el respaldo rotativo.

## [2026-11-06] Fase 4 — Opción B: catálogo de foros sin validación nativa

- **Problema:** al insertar un producto debajo de una fila de foro, la fila nueva heredaba la validación nativa de `Espacios`; por eso “Agregar producto” mostraba el combo de foro.
- **Solución:**
  - Se eliminó la validación nativa de las filas de foro.
  - El catálogo de foros se abre con **doble clic** en una fila de foro, con el combo temporal.
  - `AgregarEspacioForo` ya no aplica validación; solo inserta la fila y abre el combo.
  - Se agregó `QuitarValidacionesForo` para limpiar las validaciones `Espacios` existentes en filas de producto, SUMA o foro.
  - `EsFilaForoEn` y `AbrirCatalogoForoEn` detectan filas cuyo texto empieza con `Foro ` y abren el combo.
  - En el libro abierto se eliminaron las validaciones de `Espacios` de las filas 40, 41, 42 y 43.
- **Archivos actualizados:**
  - `catalogo_categorias.vb`
  - `bas/M_Catalogos.bas`
  - `eventos_uuid.vb`
  - `to_do.md`
- **Estado:** implementado; pendiente reimportar `M_Catalogos`, repastar el bloque de eventos de Cotización y ejecutar `QuitarValidacionesForo` una vez.

## [2026-11-06] Fase 5 — Incidente y recuperación: limpieza borró demasiado

- **Incidente:** al probar “Limpiar datos” se borraron encabezados y fórmulas del libro de trabajo.
- **Causa probable:** se ejecutó una versión anterior de `LimpiarDatosCotizacion` (con `ClearContents` sobre `A:N`) o el libro tenía un `FILA_INICIO` distinto al esperado.
- **Recuperación:**
  - Se creó una copia restaurada desde el respaldo `Plantilla Cotizacion AE 1 2_backup_2026-11-06_092219.xlsm`.
  - Archivo: `Plantilla Cotizacion AE 1 2_RESTAURADO.xlsm`.
  - Se validó que conserva datos, fórmulas, UUID y el botón de foros apuntando a `AgregarEspacioForo`.
- **Blindaje de la macro:**
  - `LimpiarDatosCotizacion` ahora nunca toca filas anteriores a la 7 (`filaInicioSegura`).
  - `EsFilaProtegida` protege directamente los encabezados fusionados `A:N` de 2 o más filas.
  - Sigue usando `SpecialCells(xlCellTypeConstants)` para borrar solo valores y conservar fórmulas.
  - Si no encuentra la fila `SUBTOTAL`, aborta sin limpiar.
- **Pendiente:** reimportar la versión corregida de `M_Plantilla`, validar sobre la copia restaurada y volver a probar “Limpiar datos”.

## [2026-11-06] Fase 5 — Corrección: conservar fórmulas al limpiar datos

- **Problema:** `LimpiarDatosCotizacion` usaba `ClearContents` sobre `A:N`, lo que borraba también las fórmulas de Subtotal (`K`) y Total (`M`).
- **Corrección:**
  - Se agregó el helper `LimpiarConstantes`, que usa `SpecialCells(xlCellTypeConstants).ClearContents`.
  - En los renglones de producto/fecha se limpian solo los valores constantes de `A:O`; las fórmulas se conservan.
  - En el pie se limpian solo los valores constantes del área de `C`.
  - `suma` recalcula los totales a 0.
- **Archivos actualizados:**
  - `nueva_cotizacion.vb`
  - `bas/M_Plantilla.bas`
- **Estado:** implementado; pendiente reimportar `M_Plantilla` y validar que las fórmulas se conserven.

## [2026-11-06] Fase 5 — Corrección: creación del botón “Limpiar datos”

- **Problema:** `CrearBotonLimpiarDatos` no creaba el botón ni mostraba mensaje.
- **Causa:** el recorrido de `Shapes` consultaba `FormControlType` en formas que no son controles (por ejemplo, el comentario de `A1`), lo que provocaba un error controlado que salía sin mensaje.
- **Corrección:**
  - Se eliminó ese recorrido; el botón se crea en una posición fija visible: parte superior del panel (`Top = 45`, `Left = 974.5`).
  - Se captura `Err.Description` antes de volver a proteger y se muestra si algo falla.
  - Después de crearlo, se registra con `GuardarPosicionesBotones` y se lleva la vista al botón.
- **Archivos actualizados:**
  - `nueva_cotizacion.vb`
  - `bas/M_Plantilla.bas`
- **Estado:** implementado; pendiente reimportar `M_Plantilla` y ejecutar `CrearBotonLimpiarDatos`.

## [2026-11-06] Fase 5 — Botón “Limpiar datos”

- **Objetivo:** vaciar los datos capturados de la cotización sin eliminar el encabezado, la estructura de bloques ni el UUID.
- **Archivos creados:**
  - `nueva_cotizacion.vb`
  - `bas/M_Plantilla.bas` (cp1252, CRLF, `Attribute VB_Name = "M_Plantilla"`)
- **Macros:**
  - `LimpiarDatosCotizacion`: pide confirmación, expande filas, limpia `A:N` y `O` en renglones de producto/fecha, limpia los datos del pie (Solicitó, Elaboró, Incluye, Especificaciones), ejecuta `suma`, restaura botones y protege con `AE2026`.
  - `CrearBotonLimpiarDatos`: agrega el botón `btnLimpiarDatos` debajo del último botón del panel, con el texto “Limpiar datos”, y ejecuta `GuardarPosicionesBotones`.
- **Decisiones:**
  - Texto del botón: “Limpiar datos”.
  - Se conserva la fecha `K3`.
  - Se vacían los datos del pie.
  - El botón se agrega automáticamente debajo del último botón.
- **Estado:** implementado; pendiente importar `M_Plantilla`, ejecutar `CrearBotonLimpiarDatos` y validar sobre una copia.

## [2026-11-06] Fase 4 — Ajustes de altura, scroll y descripción al cambiar foro

- **Combo cortado / al final de la hoja:**
  - El combo de foros ahora usa como altura la altura real de la fila (`RowHeight`, mínimo 14 pt), para que no quede cortado.
  - Después de insertar la fila, se selecciona y se desplaza la vista para dejar la fila nueva cerca del inicio visible y con espacio para el desplegable.
- **Cambio de foro con la flecha nativa:**
  - Se agregó `AjustarDescripcionForo` en `M_Catalogos`.
  - El evento `Worksheet_Change` de la hoja Cotización ahora llama a `AjustarDescripcionForo Target`.
  - Si el usuario cambia el foro con la flecha nativa, la celda se completa automáticamente como `Foro N - Descripción`.
- **Archivos actualizados:**
  - `catalogo_categorias.vb`
  - `bas/M_Catalogos.bas`
  - `eventos_uuid.vb`
- **Estado:** implementado; pendiente reimportar `M_Catalogos`, repastar el bloque de eventos de la hoja Cotización y validar visualmente.

## [2026-11-06] Fase 4 — Corrección: evitar doble combo en la fila de foros

- **Problema:** al insertar un espacio/foro se mostraban a la vez el combo temporal y la flecha de la validación nativa, ambos con la lista de `Espacios`.
- **Corrección:**
  - `MostrarComboForo` elimina temporalmente la validación nativa de la celda `A` mientras el combo está abierto.
  - `LimpiarCatalogoTemporal` restaura la validación nativa de `Espacios` en la fila al cerrar el combo, ya sea por selección o por cancelación.
  - Si el combo no se puede abrir, se aplica directamente la validación nativa como respaldo.
- **Archivos actualizados:**
  - `catalogo_categorias.vb`
  - `bas/M_Catalogos.bas`
- **Limpieza:** se agregó `LimpiarControlesHuerfanos`; en el libro abierto se eliminaron los controles `Drop Down 26` y `TMP_CatalogoForo` (pendiente guardar).
- **Estado:** implementado; pendiente reimportar `M_Catalogos` y validar visualmente.

## [2026-11-06] Fase 4 — Catálogo de foros en “Agregar espacio/foro”

- **Objetivo:** al pulsar “Agregar espacio/foro”, insertar una fila de producto y abrir automáticamente el catálogo de `Espacios`.
- **Archivos actualizados:**
  - `catalogo_categorias.vb`
  - `bas/M_Catalogos.bas` (cp1252, CRLF, `Attribute VB_Name = "M_Catalogos"`)
- **Macros nuevas/actualizadas:**
  - `AgregarEspacioForo`: valida la fila, inserta el renglón, copia formatos/fórmulas de `Plantillas!11:11`, aplica validación de `Espacios` y abre el combo.
  - `AplicarForoCatalogo`: escribe `Foro N - Descripción` en la celda `A` de la fila nueva y recalcula con `suma`.
  - `InstalarValidacionForo`: validación nativa de `Espacios` como respaldo.
  - `AsignarBotonForo`: asigna la macro al `Button 9`, elimina el control huérfano `Drop Down 31` si existe y muestra un resumen.
  - `M_Catalogos` ahora distingue entre catálogo `CAT` y `FORO` con `gTipoCatalogo`.
- **Libro `.xlsm`:**
  - `Button 9` (“Agregar espacio/foro”) quedó asignado a `AgregarEspacioForo` en `xl/worksheets/sheet2.xml` (`controlPr macro`) y en `vmlDrawing1.vml`.
  - `Button 1` (“Agregar producto”) conserva `AgregarFilaProducto`.
  - Respaldo: `Plantilla Cotizacion AE 1 2_backup_2026-11-06_092219.xlsm`.
- **Validación técnica:** el libro abre sin reparación; `Button 9` reporta `OnAction = AgregarEspacioForo`; `Button 1` conserva su macro; `Plantillas!A7` mantiene la validación de categorías; la hoja sigue protegida con `AE2026`.
- **Estado:** implementado; pendiente reimportar `M_Catalogos`, ejecutar `AsignarBotonForo` y validar visualmente.

## [2026-11-05] Fase 3 — Corrección: validación nativa restaurada + combo izquierdo corregido

- **Validación nativa:**
  - `PrepararCatalogoCategorias` ahora **instala** la lista `Categorías!$A$2:$A$...` en todos los encabezados de `Cotización`, en lugar de eliminarla.
  - Se restauró la validación de lista en `Plantillas!A7` dentro del `.xlsm` para que los bloques nuevos la hereden.
  - La flecha nativa vuelve a aparecer en la columna `N`; se puede escribir o elegir de la lista.
- **Combo izquierdo:**
  - `MostrarComboCategoria` ya no borra la validación nativa.
  - El control se crea con zoom temporal al 100% y después se posiciona explícitamente con `.Left` y `.Top`, para corregir el desplazamiento que ocurría con zoom al 50%.
  - `A1` deja de quedar tapado por el combo.
- **Respaldo:** `Plantilla Cotizacion AE 1 2_backup_2026-11-05_182145.xlsm`.
- **Validación técnica:** el libro abre sin reparación; `Plantillas!A7` queda como lista (`Type=3`, `Formula1=Categorías!$A$2:$A$12`); los mensajes de `C3`, `C4`, `K4` y `N4` siguen intactos; la hoja sigue protegida con `AE2026`.
- **Estado:** implementado; pendiente reimportar `M_Catalogos`, ejecutar `PrepararCatalogoCategorias` y validar visualmente.

## [2026-11-05] Fase 3 — Catálogo de Categoría a la izquierda del encabezado

- **Objetivo:** reemplazar la flecha nativa de la validación de categorías (que aparecía en la columna `N`) por un desplegable temporal pegado a la izquierda del encabezado, activado con doble clic.
- **Archivos creados:**
  - `catalogo_categorias.vb`
  - `bas/M_Catalogos.bas` (cp1252, CRLF, `Attribute VB_Name = "M_Catalogos"`)
- **Archivos actualizados:**
  - `eventos_uuid.vb`: se agregaron `Worksheet_BeforeDoubleClick`, `Worksheet_SelectionChange` y `Worksheet_Deactivate` al bloque de la hoja Cotización.
  - `blindaje_botones.vb` y `bas/blindaje_botones.bas`: `GuardarPosicionesBotones` y `LimpiarBotonesDuplicados` ignoran los objetos `TMP_`.
  - `generar_pdfs_resumen_detalle.vb` y `bas/generar_pdfs_resumen_detalle.bas`: antes de exportar se eliminan los combos `TMP_Catalogo`.
- **Libro `.xlsm`:**
  - Se retiró la validación de categorías de `Plantillas!A7` para que los bloques nuevos no hereden la flecha nativa de la columna `N`.
  - Respaldo: `Plantilla Cotizacion AE 1 2_backup_2026-11-05_173118.xlsm`.
- **Macros nuevas:**
  - `PrepararCatalogoCategorias` (una sola vez: quita la validación nativa de los encabezados).
  - `AbrirCatalogoCategoria` / `AbrirCatalogoCategoriaEn`.
  - `AplicarCategoriaCatalogo`.
  - `LimpiarCatalogoTemporal` / `LimpiarCatalogoSiCorresponde`.
- **Estado:** implementado en archivos; pendiente de importación y validación en Excel.

## [2026-11-05] Fase 2 — Mensajes de entrada nativos en el encabezado

- **Celdas:** `C3` (Cliente), `C4` (Nombre del proyecto), `K4` (ID Cotización) y `N4` (Versión).
- **Implementación:** sin módulos VBA; se agregaron directamente en `xl/worksheets/sheet2.xml`.
  - `C3`, `C4`, `K4`: validaciones de solo mensaje de entrada (`Type=0`, `ShowInput=True`).
  - `N4`: se conservó su validación de lista (`Type=3`, `Formula1=Versión!$A$2:$A$11`) y solo se agregó el mensaje de entrada.
  - `H3`, `K3` y `H4` no se modificaron.
- **Respaldo:** `Plantilla Cotizacion AE 1 2_backup_2026-11-05_170928.xlsm`.
- **Validación:** el libro abre en Excel sin reparación, la hoja sigue protegida con `AE2026`, y los 4 mensajes se leen correctamente.
- **Estado:** implementado y validado técnicamente; pendiente validación visual del usuario.

## [2026-11-05] Limpieza posterior — Retiro del módulo M_Ayuda (Fase 1)

- **Motivo:** las notas de ayuda se implementarán con **Validación de datos → Mensaje de entrada** directamente en Excel, sin módulos VBA adicionales.
- **Archivos eliminados del repositorio:**
  - `ayuda_plantilla.vb`
  - `bas/ayuda.bas`
- **Módulo en el libro:** `M_Ayuda` fue retirado por el usuario desde el editor de VBA; el proyecto ya no lo lista entre sus módulos activos (solo quedan cadenas en la caché interna de compilación).
- **Se conserva:**
  - El texto visible `"Esta es una plantilla"` y su nota emergente en `A1`, ya guardados en la hoja.
  - El ajuste de `M_PDF` para ocultar temporalmente `A1` durante la exportación a PDF.
  - La protección de la hoja `Cotización` con `AE2026`.
- **Estado:** `M_Ayuda` retirado y libro guardado; `A1` y el ajuste de `M_PDF` se conservan.

## [2026-11-05] Fase 0 y Fase 1 — Nota de plantilla y limpieza de módulo duplicado

### Fase 0 — Eliminación del nombre ambiguo y corrección de protección
- **Objetivo:** dejar el proyecto VBA compilable y la hoja Cotización desprotegida con una contraseña conocida.
- **Resultado en el libro:** la copia de trabajo `Plantilla Cotización AE 1 2.xlsm` no contiene `Módulo11`; el usuario ejecutó **Depuración > Compilar VBAProject** y compiló sin errores.
- **Protección de la hoja:** la hoja `Cotización` estaba protegida con un hash SHA-512 distinto a `PWD_BLINDAJE` (`AE2026`). Se eliminó únicamente el elemento `<sheetProtection>` de `xl/worksheets/sheet2.xml`; el resto del libro y el proyecto VBA no se modificaron.
- **Respaldo creado:** `Plantilla Cotizacion AE 1 2_backup_2026-11-05_153620.xlsm`.
- **Validación estructural:** 74 entradas del ZIP conservadas, `vbaProject.bin` intacto, XML válido y apertura correcta en Excel (sin reparación). La hoja quedó con `ProtectDrawingObjects = False` y `ProtectContents = False`.
- **Pendiente:** volver a proteger la hoja con `AE2026` ejecutando `BloquearBotones` después de instalar la nota.

### Fase 1 — Nota de plantilla en `A1` con texto visible y ocultación en PDF (implementada)
- **Archivos creados / actualizados:**
  - `ayuda_plantilla.vb`
  - `bas/ayuda.bas` (cp1252, CRLF, `Attribute VB_Name = "M_Ayuda"`)
  - `generar_pdfs_resumen_detalle.vb`
  - `bas/generar_pdfs_resumen_detalle.bas` (cp1252, CRLF, `Attribute VB_Name = "M_PDF"`)
- **Macro:** `InstalarNotaPlantilla` en el módulo `M_Ayuda`.
- **Descripción:**
  - Escribe en `A1` el texto visible `"Esta es una plantilla"` en negrita y color negro (`RGB(0, 0, 0)`).
  - Conserva el comentario/nota emergente completo en `A1` indicando que el archivo es una plantilla y que no deben borrarse bloques ni filas de estructura.
  - Re-protege automáticamente los objetos de dibujo (`BloquearBotones ws`) al finalizar.
  - En `M_PDF.ExportarPDF`, oculta temporalmente el texto de `A1` (`NumberFormat = ";;;"`) durante la exportación a PDF y lo restaura inmediatamente después para que no aparezca impreso en los PDFs del cliente.
- **Dependencias reutilizadas:** `HojaCotizacion()`, `DesbloquearBotones()` y `BloquearBotones()` existentes en el libro.
- **Estado:** implementada y validada en Excel. El módulo `M_Ayuda` y sus archivos se retiraron después de validar la nota; ver la sección de limpieza del 2026-11-05.

## [2026-10-12] Implementación de Mejoras de Robustez, UX y Blindaje Operativo (Plan de Mejora Integral)

Basado en el documento [EVALUACION_Y_PLAN_MEJORA_MACRO.md](file:///e:/Users/1167486/Local/scripts/macros/macros_excel/azteca_estudios_macro/EVALUACION_Y_PLAN_MEJORA_MACRO.md) y las reglas operativas de [AGENTS.md](file:///e:/Users/1167486/Local/scripts/macros/macros_excel/azteca_estudios_macro/AGENTS.md).

### 1. Validación de Contexto de Celda Activa (`ValidarFilaEnBloque`)
- **Archivos modificados:**
  - [blindaje_botones.vb](file:///e:/Users/1167486/Local/scripts/macros/macros_excel/azteca_estudios_macro/blindaje_botones.vb)
  - [bas/blindaje_botones.bas](file:///e:/Users/1167486/Local/scripts/macros/macros_excel/azteca_estudios_macro/bas/blindaje_botones.bas)
  - [agregar_fila_producto.vb](file:///e:/Users/1167486/Local/scripts/macros/macros_excel/azteca_estudios_macro/agregar_fila_producto.vb)
  - [agregar_fila_fecha.vb](file:///e:/Users/1167486/Local/scripts/macros/macros_excel/azteca_estudios_macro/agregar_fila_fecha.vb)
  - [agregar_fila_item.vb](file:///e:/Users/1167486/Local/scripts/macros/macros_excel/azteca_estudios_macro/agregar_fila_item.vb)
  - [borrar_fila_seleccionada.vb](file:///e:/Users/1167486/Local/scripts/macros/macros_excel/azteca_estudios_macro/borrar_fila_seleccionada.vb)
  - [duplicar_bloque.vb](file:///e:/Users/1167486/Local/scripts/macros/macros_excel/azteca_estudios_macro/duplicar_bloque.vb)
- **Descripción:** Se implementó la función guardiana pública `ValidarFilaEnBloque(fila As Long)` dentro de `M_Blindaje`. Comprueba que la hoja activa sea `Cotización`, que la fila seleccionada sea $\ge \text{FILA\_INICIO}$ y estrictamente anterior al bloque de totales finales (`SUBTOTAL`/`IVA`/`TOTAL`). Si el usuario hace clic en los encabezados o en los totales, se muestra un mensaje informativo amigable y se aborta la operación sin alterar la hoja.

### 2. Blindaje y Manejo Robusto de Excepciones en PDF (`M_PDF`)
- **Archivos modificados:**
  - [generar_pdfs_resumen_detalle.vb](file:///e:/Users/1167486/Local/scripts/macros/macros_excel/azteca_estudios_macro/generar_pdfs_resumen_detalle.vb)
  - [bas/generar_pdfs_resumen_detalle.bas](file:///e:/Users/1167486/Local/scripts/macros/macros_excel/azteca_estudios_macro/bas/generar_pdfs_resumen_detalle.bas)
- **Descripción:**
  - **Detección de PDF en uso / bloqueado:** Captura el error de tiempo de ejecución `1004` al exportar. Si el PDF ya está abierto en Acrobat o el navegador, informa amigablemente al usuario para que lo cierre, evitando la pantalla de depuración de VBA.
  - **Validación de libro guardado:** Comprueba `Len(ThisWorkbook.Path) > 0` antes de proceder.
  - **Recálculo automático previo:** Invoca `suma()` antes de exportar para asegurar que los PDFs reflejen los datos más recientes.
  - **Apertura interactiva de carpeta:** Al completar con éxito la exportación de ambos PDFs, ofrece abrir el Explorador de Windows con el archivo seleccionado mediante `Shell explorer.exe /select...`.

### 3. Limpieza Segura con Confirmación y Recálculo (`limpiar_cotizacion.vb`)
- **Archivos modificados:**
  - [limpiar_cotizacion.vb](file:///e:/Users/1167486/Local/scripts/macros/macros_excel/azteca_estudios_macro/limpiar_cotizacion.vb)
- **Descripción:**
  - Cuadro de confirmación modal `MsgBox vbYesNo` para evitar borrado accidental.
  - Anclaje preventivo de botones (`FijarBotones` y `RestaurarPosicionesBotones`).
  - Protección de la columna A: se conservan intactos los títulos de categoría y la estructura. Solo se limpian las columnas de captura de datos (B a L y O).
  - Recálculo automático con `Call suma` para dejar los totales en $0.00.
  - Corrección de la errata tipográfica en el mensaje de éxito final.

### 4. Recálculo Matemático Tolerante (`suma.vb`)
- **Archivos modificados:**
  - [suma.vb](file:///e:/Users/1167486/Local/scripts/macros/macros_excel/azteca_estudios_macro/suma.vb)
- **Descripción:** Normalización de cadenas de búsqueda eliminando espacios, puntos y dos puntos (`Replace(Replace(UCase(texto), ":", ""), " ", "") = "SUMATOTAL"`) para tolerar variaciones accidentales de formato en las etiquetas. Distinción explícita de `TOTAL` frente a `SUBTOTAL` y cálculo redondeado de IVA.

### 5. Protección de Folio UUID ante Cambios de Proyecto (`M_UUID`)
- **Archivos modificados:**
  - [asignar_uuid_cotizacion.vb](file:///e:/Users/1167486/Local/scripts/macros/macros_excel/azteca_estudios_macro/asignar_uuid_cotizacion.vb)
  - [bas/asignar_uuid_cotizacion.bas](file:///e:/Users/1167486/Local/scripts/macros/macros_excel/azteca_estudios_macro/bas/asignar_uuid_cotizacion.bas)
- **Descripción:** En caso de que se modifique `C4` (Nombre del Proyecto) y ya existiera un UUID asignado en `Meta!B2`, se solicita confirmación al usuario antes de regenerar el folio. Esto permite corregir errores tipográficos en el nombre del proyecto conservando el mismo ID de cotización.

### 6. Sincronización Integral de Módulos `.bas` (cp1252 + CRLF)
- **Archivos modificados:**
  - [bas/asignar_uuid_cotizacion.bas](file:///e:/Users/1167486/Local/scripts/macros/macros_excel/azteca_estudios_macro/bas/asignar_uuid_cotizacion.bas) (`M_UUID`)
  - [bas/generar_pdfs_resumen_detalle.bas](file:///e:/Users/1167486/Local/scripts/macros/macros_excel/azteca_estudios_macro/bas/generar_pdfs_resumen_detalle.bas) (`M_PDF`)
  - [bas/blindaje_botones.bas](file:///e:/Users/1167486/Local/scripts/macros/macros_excel/azteca_estudios_macro/bas/blindaje_botones.bas) (`M_Blindaje`)
  - [bas/esquema_categorias.bas](file:///e:/Users/1167486/Local/scripts/macros/macros_excel/azteca_estudios_macro/bas/esquema_categorias.bas) (`M_Esquema`)
  - [bas/macro_generadora_uuid.bas](file:///e:/Users/1167486/Local/scripts/macros/macros_excel/azteca_estudios_macro/bas/macro_generadora_uuid.bas)
- **Descripción:** Regeneración de todos los módulos `.bas` en codificación Windows-1252 con saltos de línea CRLF y atributos `VB_Name` correspondientes para importación directa en el editor VBA de Excel.

### 7. Configuración de Entorno Conda en Documentación de Agentes
- **Archivos modificados:**
  - [AGENTS.md](file:///e:/Users/1167486/Local/scripts/macros/macros_excel/azteca_estudios_macro/AGENTS.md)
  - [.agents/AGENTS.md](file:///e:/Users/1167486/Local/scripts/macros/macros_excel/azteca_estudios_macro/.agents/AGENTS.md)
- **Descripción:** Se documentó la instrucción para activar el entorno de Anaconda `skills_env` cuando se requiera ejecutar herramientas en Python.
