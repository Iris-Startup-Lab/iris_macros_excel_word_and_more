# Macros Multiherramienta

Esta carpeta contiene macros que integran múltiples aplicaciones del ecosistema Microsoft Office (Excel, Word, Outlook) para resolver flujos de trabajo colaborativos y automatizados.

---

> 📖 **Guía Completa**: Consulta el [tutorial.md](tutorial.md) para ver la guía paso a paso de instalación, configuración de columnas, asignación de botones y automatización diaria.

## 📌 Lista de Macros

### 1. `recordatorio_checklist_experimentos.vb`
Automatiza el envío de recordatorios y listas de verificación (checklist) por correo electrónico a todos los participantes de un experimento **exactamente 1 día antes** de su fecha programada.

#### 🛠 Herramientas Involucradas:
- **Microsoft Excel**: Base de datos de experimentos, fechas, participantes, requerimientos y estado de envío.
- **Microsoft Outlook**: Motor de despacho y distribución de correos electrónicos.
- **Microsoft Word (Opcional)**: Motor de maquetación de plantillas ricas (`.docx`) si el equipo prefiere diseñar el formato en Word.

#### 🚀 Modos de Uso Disponibles:
1. **`EnviarRecordatoriosExperimentos` (Recomendado)**:
   - Lee la hoja de Excel, calcula `Fecha - 1 = Hoy`.
   - Genera un correo con plantilla HTML moderna, limpia y casillas de verificación interactiva (`☐`).
   - Actualiza el estado a `ENVIADO` con fecha y hora para no duplicar correos.
2. **`ProbarEnvioFilaSeleccionada`**:
   - Permite previsualizar en pantalla (`.Display`) el correo de la fila activa seleccionada en Excel, sin importar la fecha, facilitando pruebas y validaciones visuales.
3. **`EnviarUsandoPlantillaWord`**:
   - Reemplaza marcadores (`{{TITULO}}`, `{{FECHA}}`, `{{CHECKLIST}}`, etc.) dentro de un documento Word prediseñado y lo inserta en el cuerpo de Outlook.

---

## 📋 Estructura Recomendada de la Hoja de Excel (`Experimentos`)

| Columna | Nombre de Campo | Ejemplo | Descripción |
| :--- | :--- | :--- | :--- |
| **A** | `ID` | `EXP-001` | Identificador único del experimento |
| **B** | `Titulo` | `Prueba de Resistencia de Materiales A` | Nombre del experimento |
| **C** | `Responsable` | `Dr. Roberto Gómez` | Líder del experimento |
| **D** | `Correos` | `analista1@empresa.com; tester2@empresa.com` | Correos de los participantes (separados por `;`) |
| **E** | `Fecha` | `17/10/2026` | Fecha de ejecución (Formato Fecha `DD/MM/AAAA`) |
| **F** | `Hora` | `09:30 AM` | Horario de la sesión |
| **G** | `Ubicación` | `Laboratorio de Física B2` | Sala o laboratorio asignado |
| **H** | `Checklist / Requerimientos` | `- Calibrar sensores de presión<br>- Traer batas y gafas de seguridad<br>- Muestra seca a 25°C` | Lista de puntos a verificar (separados por saltos de línea `Alt + Enter`) |
| **I** | `Estado Envio` | `PENDIENTE` / `ENVIADO` | Control de envío para evitar correos duplicados |
| **J** | `Fecha Envio` | `16/10/2026 08:00:15` | Registro de auditoría con fecha y hora del envío |

---

## ⚡ Automatización Diaria (Sin intervención humana)
Para que los correos se envíen automáticamente todas las mañanas:
1. **Opción A (Al abrir el libro)**: Agrega la llamada a `EnviarRecordatoriosExperimentos` en el evento `Workbook_Open` dentro de `ThisWorkbook`.
2. **Opción B (Programador de Tareas de Windows)**: Programa la apertura automática del libro a las 08:00 AM todos los días laborables.
