# 📊 Excel VBA Automation Toolkit

Este proyecto contiene macros en VBA orientadas a automatizar dos procesos clave:

1. Generación de identificadores únicos (UUIDs)
2. Construcción y cálculo automático de cotizaciones estructuradas

---

# 1. Generación de UUIDs (VBA)

## Descripción técnica

La macro implementa un generador de identificadores únicos tipo GUID, produciendo strings con el formato estándar



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