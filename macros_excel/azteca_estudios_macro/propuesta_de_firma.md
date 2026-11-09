# Propuesta: Firma Digital de las Macros

## Plantilla de Cotizaciones — Azteca Estudios

**Fecha:** 2026-11-09
**Estado:** Propuesta, pendiente de validar con el cliente
**Alcance:** Firma digital del proyecto VBA de `Plantilla Cotización AE.xlsm`. En un anexo se incluye una segunda propuesta pendiente: totales con fórmulas nativas.

---

## 1. Problema que resuelve

Hoy la plantilla depende de que las macros estén habilitadas. Sin ellas no funcionan el ID de cotización (UUID), el blindaje de botones y hojas, el recálculo ni la generación de PDFs.

En la práctica, el usuario se topa con estos obstáculos:

| Situación actual | Consecuencia |
|---|---|
| Al abrir el archivo aparece la barra amarilla **"Habilitar contenido"**. | Si el usuario no la pulsa, la plantilla funciona "a medias" y no avisa. |
| Si el archivo llega **por correo, Teams o descarga**, Windows lo marca como "de Internet" y Office **bloquea las macros**: barra roja, sin botón para habilitarlas. | El usuario tiene que ir a Propiedades del archivo → "Desbloquear". Casi nadie sabe hacerlo. |
| Cualquier persona puede modificar el código VBA. | No hay forma de saber si el archivo que circula es el original. |

## 2. Qué es la firma digital y cómo funciona

1. Con un **certificado digital** (un archivo de identidad emitido a nombre de una persona o de la empresa) se "sella" el código VBA desde el editor de macros.
2. El sello es una **huella matemática** del código. Si alguien cambia aunque sea una letra del código, la huella ya no coincide.
3. Cada vez que alguien abre el archivo, Excel comprueba en silencio dos cosas:
   - **¿El código está intacto?** La huella debe coincidir.
   - **¿Confío en quien firmó?** Su certificado debe estar en la lista de **"Editores de confianza"** de esa computadora.

## 3. Qué cambia al abrir el archivo

| Caso | Qué ve el usuario |
|---|---|
| **Firmado + editor de confianza** (el escenario objetivo) | Las macros **se activan solas**: sin barra amarilla y **sin el bloqueo de archivos de correo o descargados**. |
| **Firmado + editor aún no confiable** | Un aviso con la opción **"Confiar en todos los documentos de este editor"**. Se acepta **una sola vez** por computadora. |
| **Alguien alteró el código** | Excel retira la firma y el archivo vuelve a comportarse como hoy (aviso o bloqueo). Esto sirve como **alerta de manipulación**. |

**Lo que NO cambia:**

- Capturar datos, agregar filas, generar PDFs y guardar funcionan igual.
- **Editar celdas no rompe la firma.** Solo la rompe modificar el **código VBA**.
- **La firma no oculta el código.** Quien abra el editor (Alt+F11) lo sigue viendo. Para ocultarlo se necesita además el **bloqueo del proyecto VBA con contraseña**, que es una medida complementaria (ver sección 6).

## 4. Opciones de certificado

| Opción | Costo | Confianza | Esfuerzo de despliegue | Recomendación |
|---|---|---|---|---|
| **A. Certificado corporativo** (emitido por TI / la autoridad certificadora interna de la empresa) | Normalmente $0 (infraestructura existente) | Automática en todas las PCs del dominio, vía directiva de grupo (GPO) | Bajo para el usuario; requiere una solicitud a TI | ✅ **Recomendada** si TI lo ofrece |
| **B. Certificado comercial de firma de código** (DigiCert, Sectigo, etc.) | Aprox. USD 200–500 al año | Reconocido en cualquier PC, incluso fuera de la empresa | Medio (validación de identidad de la empresa; token USB) | Solo si el archivo se comparte con externos |
| **C. SelfCert** (autofirmado, incluido con Office) | $0 | **Solo en la PC que lo creó.** En las demás hay que instalar el certificado a mano | Alto: una instalación por cada computadora | Útil solo para pruebas o pocos usuarios |

## 5. Procedimiento propuesto (opción A)

**Una sola vez:**

1. Solicitar a TI un **certificado de firma de código** (*Code Signing*) a nombre del responsable de la plantilla o del área.
2. Pedir a TI que publique ese certificado como **Editor de confianza** en las PCs de los usuarios (por GPO).
3. Opcional pero recomendado: pedir que la directiva de macros de Office quede en **"Deshabilitar todas las macros excepto las firmadas digitalmente"**. Así ningún archivo alterado o ajeno ejecuta código.

**Cada vez que se actualicen las macros** (por ejemplo, al reimportar módulos):

1. Abrir el editor (Alt+F11) → menú **Herramientas → Firma digital**.
2. **Elegir** el certificado → **Aceptar**.
3. Guardar el `.xlsm`.
4. Distribuir esa versión firmada.

> ⚠️ Si alguien sin el certificado modifica y guarda el código, la firma desaparece. Solo el responsable con el certificado debe publicar versiones nuevas.

**Sello de tiempo:** al firmar con un certificado corporativo o comercial, conviene configurar un servidor de sello de tiempo (*timestamp*). Así la firma **sigue siendo válida aunque el certificado expire después**.

## 6. Medida complementaria: bloqueo del proyecto VBA

| Medida | Protege contra | Cómo |
|---|---|---|
| Firma digital | Código alterado; bloqueo de Office | Sección 5 |
| Bloqueo del proyecto VBA | Que se **vea o copie** el código, incluida la contraseña de blindaje (`PWD_BLINDAJE`) | Alt+F11 → Herramientas → Propiedades de VBAProject → pestaña **Protección** → "Bloquear proyecto para visualización" + contraseña |

**Recomendaciones asociadas:**

- Cambiar la contraseña de blindaje actual (`AE2026`) antes de distribuir.
- Dejar de guardarla en el repositorio Git.
- Guardar ambas contraseñas (blindaje y proyecto VBA) en un gestor de contraseñas del equipo.

## 7. Riesgos y consideraciones

| Riesgo | Mitigación |
|---|---|
| Se pierde o vence el certificado | Sello de tiempo al firmar; TI renueva y se vuelve a firmar |
| Se olvida volver a firmar tras una actualización | Agregar "Firmar" como paso obligatorio en la guía de actualización (`instrucciones.md`) |
| Usuarios con Excel para Mac | La firma VBA de Windows no aplica igual, y el UUID ya requiere Windows. Sin cambio respecto a hoy |
| La contraseña de blindaje sigue siendo una barrera débil | La protección de hojas de Excel es disuasoria, no de seguridad fuerte. Se acepta como prevención de errores, no de mala fe |

## 8. Preguntas para validar con el cliente

1. ¿TI de la empresa puede emitir un **certificado de firma de código** y distribuirlo por GPO? (Define si es la opción A o B/C.)
2. ¿El archivo se comparte **con personas fuera de la empresa**? (Si sí, evaluar la opción B.)
3. ¿Quién será el **responsable** de firmar y publicar versiones nuevas?
4. ¿Se acepta fijar la directiva **"solo macros firmadas"** en las PCs de los usuarios?
5. ¿Se aprueba **bloquear el proyecto VBA** y **cambiar la contraseña de blindaje**?

---

## Anexo — Propuesta pendiente: totales con fórmulas nativas

**Estado:** Pendiente de validar con el cliente. **No implementado.**

**Idea:** hoy los totales son **valores** que escribe la macro `suma()`. Si las macros no corren, o alguien cambia un dato sin que se recalcule, los totales quedan desactualizados. La propuesta es convertirlos en **fórmulas de Excel**, que se recalculan solas:

| Celda | Fórmula (Excel en español) | Qué hace |
|---|---|---|
| Columna `M` de cada fila **SUMA TOTAL** | `=SUMAR.SI.CONJUNTO(M7:INDICE(M:M,FILA()-1), A7:INDICE(A:A,FILA()-1), "<>")` | Suma el bloque, desde su título hasta la fila de arriba, ignorando filas sin concepto |
| Columna `K` de **SUBTOTAL** | `=SUMAR.SI(A$7:INDICE(A:A,FILA()-1), "SUMA TOTAL*", M$7:INDICE(M:M,FILA()-1))` | Suma todas las SUMA TOTAL |
| Columna `K` de **IVA** | `=REDONDEAR(subtotal * 0.16, 2)` | 16 % del subtotal |
| Columna `K` de **TOTAL** | `=subtotal + IVA` | Total final |

*`M7` y `A7` corresponden a la fila de título de cada bloque.* El uso de `INDICE(…FILA()-1)` garantiza que una fila insertada justo arriba de la SUMA TOTAL **sí** entre en el total; con un rango fijo quedaría fuera.

**Beneficios:**

- Totales siempre correctos, incluso con las macros deshabilitadas.
- El PDF coincide siempre con lo que se ve en pantalla.
- La captura es más fluida, porque ya no se recalcula por macro en cada cambio.

**Implementación prevista** (tras la validación):

1. `suma()` escribirá las fórmulas en lugar de valores, y repondrá las que alguien sobrescriba.
2. Se actualizarán el bloque modelo de la hoja `Plantillas` y la hoja `_BaseAE`.
3. Una macro de migración convertirá las cotizaciones existentes.

**Preguntas para validar con el cliente:**

1. ¿La columna `M` de cada renglón (importe) ya es una **fórmula** (p. ej. cantidad × precio) o se **captura a mano**?
2. ¿Las filas de **detalle** (marca "D", ocultas en el PDF Resumen) deben **sumar** al total? Hoy sí suman.
3. ¿El IVA es **siempre 16 %**, o hay casos de **8 % (frontera)** o **exentos**? Si hay variación, la tasa se tomaría de una celda editable.
4. ¿Se acepta que alguien pueda sobrescribir una fórmula a mano, sabiendo que la macro la repondrá al generar el PDF o al guardar? ¿O se prefiere bloquear esas celdas?
