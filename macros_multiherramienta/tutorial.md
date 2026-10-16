# 📖 Tutorial: Macro de Recordatorio y Checklist para Experimentos

Este tutorial explica paso a paso cómo configurar, instalar y ejecutar la macro multiherramienta que envía automáticamente correos electrónicos recordatorios y listas de verificación (checklist) a las personas involucradas **exactamente 1 día antes** de la realización de un experimento.

---

## 📑 Contenido
1. [Requisitos Previos](#1-requisitos-previos)
2. [Paso 1: Configurar la Hoja en Excel](#2-paso-1-configurar-la-hoja-en-excel)
3. [Paso 2: Instalar el Código VBA en Excel](#3-paso-2-instalar-el-código-vba-en-excel)
4. [Paso 3: Crear Botones de Control en Excel](#4-paso-3-crear-botones-de-control-en-excel)
5. [Paso 4: (Opcional) Usar Plantilla de Microsoft Word](#5-paso-4-opcional-usar-plantilla-de-microsoft-word)
6. [Paso 5: Automatización Diaria sin Intervención Manual](#6-paso-5-automatización-diaria-sin-intervención-manual)
7. [Preguntas Frecuentes y Solución de Problemas](#7-preguntas-frecuentes-y-solución-de-problemas)

---

## 1. Requisitos Previos

- **Microsoft Excel** (versión 2016, 2019, 2021 o Microsoft 365 para Windows).
- **Microsoft Outlook** con la cuenta de correo corporativa configurada e iniciada.
- **Microsoft Word** (solo requerido si deseas usar la variante de maquetación con archivo `.docx`).
- Permiso para habilitar macros en Excel (`.xlsm`).

---

## 2. Paso 1: Configurar la Hoja en Excel

1. Abre un libro nuevo o existente de Excel.
2. Nombra a la hoja principal exactamente: **`Experimentos`**.
3. En la **Fila 1**, escribe los siguientes encabezados de columna:

| Columna | Nombre del Encabezado | Ejemplo de Datos | ¿Para qué sirve? |
| :---: | :--- | :--- | :--- |
| **A** | `ID` | `EXP-001` | Identificador único del experimento. |
| **B** | `Titulo` | `Prueba de Resistencia de Polímeros` | Nombre formal del experimento. |
| **C** | `Responsable` | `Dr. Roberto Gómez` | Líder o encargado del experimento. |
| **D** | `Correos` | `investigador1@lab.com; tecnico@lab.com` | Correos de los destinatarios separados por punto y coma (`;`). |
| **E** | `Fecha` | `17/10/2026` | Fecha programada del experimento (**formato Fecha**). |
| **F** | `Hora` | `09:30 AM` | Horario o turno asignado. |
| **G** | `Ubicación` | `Laboratorio de Materiales Sala B` | Espacio o laboratorio asignado. |
| **H** | `Checklist / Requerimientos` | `Calibrar sensores de presión`<br>`Verificar temperatura a 22°C`<br>`Traer batas y gafas protectoras` | Lista de puntos o materiales. **Tip:** Presiona `Alt + Enter` dentro de la celda para separar cada punto con un salto de línea. |
| **I** | `Estado Envio` | `PENDIENTE` | Indica si ya se envió el correo. La macro lo cambiará a `ENVIADO`. |
| **J** | `Fecha Envio` | *(dejar en blanco)* | La macro registrará aquí la fecha y hora exacta del envío. |

> [!TIP]
> Puedes dar formato de tabla o aplicar color a los encabezados para mejorar la apariencia visual.

---

## 3. Paso 2: Instalar el Código VBA en Excel

1. Con el archivo de Excel abierto, presiona **`ALT + F11`** (o en Mac / laptops especiales: `Fn + Alt + F11`) para abrir el editor de Visual Basic para Aplicaciones (VBA).
2. En la barra de menú superior, selecciona **Insertar** > **Módulo**.
3. Se abrirá una ventana en blanco con el título `Módulo1`.
4. Abre el archivo de código [`recordatorio_checklist_experimentos.vb`](recordatorio_checklist_experimentos.vb), copia todo su contenido y pégalo en la ventana en blanco del módulo.
5. Haz clic en el ícono de **Guardar** (o `Ctrl + S`).
6. Si el archivo aún es `.xlsx`, Excel te advertirá que debes guardarlo como libro habilitado para macros:
   - Selecciona **No** en la advertencia.
   - En el tipo de archivo, elige: **`Libro de Excel habilitado para macros (*.xlsm)`**.
   - Guarda el archivo.

---

## 4. Paso 3: Crear Botones de Control en Excel

Para que cualquier usuario pueda ejecutar la macro con un solo clic:

### Botón 1: Envío General (Recordatorios de Mañana)
1. Ve a la hoja **`Experimentos`** en Excel.
2. En la cinta de opciones, ve a la pestaña **Insertar** > **Ilustraciones** > **Formas**.
3. Elige un rectángulo con esquinas redondeadas y dibújalo en una zona visible (por ejemplo, arriba de la tabla).
4. Escribe dentro del botón: **`🚀 Enviar Recordatorios Mañana`**.
5. Dale estilo y formato con tus colores preferidos.
6. Haz **clic derecho** sobre la forma y selecciona **Asignar macro...**.
7. Selecciona **`EnviarRecordatoriosExperimentos`** y presiona **Aceptar**.

### Botón 2: Probar Fila Seleccionada (Previsualización)
1. Inserta una segunda forma más pequeña al lado.
2. Escribe dentro: **`👁 Probar fila seleccionada`**.
3. Haz **clic derecho** > **Asignar macro...**.
4. Selecciona **`ProbarEnvioFilaSeleccionada`** y presiona **Aceptar**.
5. **Uso de este botón:** Haz clic en cualquier celda de una fila con datos y pulsa el botón; se abrirá Outlook mostrando cómo lucirá el correo antes de enviarlo, sin importar qué fecha tenga asignada.

---

## 5. Paso 4: (Opcional) Usar Plantilla de Microsoft Word

Si la organización cuenta con una plantilla institucional en Word (`.docx`) con logotipos, encabezados o diseño gráfico particular:

1. Crea un documento nuevo en **Microsoft Word**.
2. Diseña el formato a tu gusto y escribe las siguientes etiquetas en el texto donde quieras que se inserten los datos:
   - `{{TITULO}}`
   - `{{RESPONSABLE}}`
   - `{{FECHA}}`
   - `{{HORA}}`
   - `{{UBICACION}}`
   - `{{CHECKLIST}}`
3. Guarda el documento con el nombre `Plantilla_Checklist.docx`.
4. En Excel, ejecuta la macro **`EnviarUsandoPlantillaWord`**.
5. Se abrirá una ventana para que selecciones el archivo `.docx`. La macro leerá la plantilla, reemplazará las etiquetas con los datos del experimento y generará el correo en Outlook con el diseño de Word.

---

## 6. Paso 5: Automatización Diaria sin Intervención Manual

Si no deseas pulsar el botón manualmente cada día, puedes habilitar la ejecución automática:

### Opción A: Ejecución automática al abrir el archivo Excel
1. Presiona **`ALT + F11`** para abrir el Editor de VBA.
2. En el explorador de proyectos de la izquierda, busca tu archivo y haz doble clic sobre el objeto **`ThisWorkbook`** (o `EsteLibro`).
3. En la ventana de código que se abre, pega el siguiente bloque:

```vba
Private Sub Workbook_Open()
    ' Ejecuta automáticamente la macro al abrir el libro
    Call EnviarRecordatoriosExperimentos
End Sub
```
4. Guarda los cambios (`Ctrl + S`). Cada vez que alguien abra el archivo, se revisará la lista y se enviarán los correos de los experimentos que se llevarán a cabo al día siguiente.

### Opción B: Programador de Tareas de Windows (100% Desatendido)
Para que el proceso se ejecute todos los días hábiles a una hora fija (por ejemplo, a las 8:00 AM) sin que nadie tenga que abrir el archivo:

1. Presiona la tecla `Windows`, escribe **Programador de tareas** y ábrelo.
2. En el panel derecho, haz clic en **Crear tarea básica...**.
3. **Nombre:** `Recordatorio Experimentos Diario`.
4. **Desencadenador:** Selecciona *Diariamente* y define la hora (ejemplo: `08:00 AM`).
5. **Acción:** Selecciona *Iniciar un programa*.
   - **Programa o script:** `excel.exe` (o la ruta completa de Excel, ej. `C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE`).
   - **Agregar argumentos:** Escribe entre comillas la ruta completa a tu archivo `.xlsm`, por ejemplo:
     `"C:\MisProyectos\Experimentos_Lab.xlsm"`
6. Haz clic en **Finalizar**. Ahora, a las 8:00 AM, Windows abrirá el archivo, `Workbook_Open` ejecutará la macro y despachará los correos.

---

## 7. Preguntas Frecuentes y Solución de Problemas

### ❓ ¿Qué pasa si ejecuto la macro dos veces el mismo día?
La macro verifica la columna **`Estado Envio`**. Si ya dice `ENVIADO`, omitirá esa fila automáticamente para evitar correos duplicados.

### ❓ ¿Cómo coloco varios correos para un mismo experimento?
Separa los correos con punto y coma (`;`), por ejemplo:
`cientifico1@empresa.com; asistente2@empresa.com; supervisor@empresa.com`

### ❓ ¿Cómo escribo varios puntos en la celda de Checklist?
Entra en modo edición en la celda (con doble clic o presionando `F2`), escribe el primer punto, presiona **`Alt + Enter`** para crear un salto de línea y escribe el siguiente. La macro convertirá automáticamente cada línea en un elemento del checklist con su casilla (`☐`).

### ❓ ¿Outlook debe estar abierto?
No es obligatorio tenerlo abierto previamente; la macro iniciará el proceso de Outlook en segundo plano si está cerrado, siempre y cuando tengas un perfil de correo configurado en tu equipo.
