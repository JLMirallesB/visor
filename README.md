# VISOR

Visor offline de registros de orientación para conservatorios profesionales de música. Lee archivos Excel generados por Microsoft Forms y muestra fichas del alumnado, histórico de actuaciones y estadísticas — todo en local, sin enviar datos a ningún servidor.

## ¿Qué es?

Un único archivo HTML que abres en el navegador. Eliges un Excel local con dos hojas (`Sheet1` con las actuaciones del Forms y `Ficha` con los datos del alumnado) y la app te muestra:

- Listado del alumnado con búsqueda multicampo y filtros por seguimiento, situación, curso, especialidad y curso académico.
- Ficha individual con datos académicos, contactos, diagnóstico e histórico de actuaciones.
- Pestaña de estadísticas con la carga del departamento (prioritarios, pendientes, alertas, actividad mensual, distribuciones, diagnósticos).
- Impresión / guardado como PDF de fichas, actuaciones y estadísticas, con cabecera personalizable.

## Cómo usarlo

1. **Descarga `visor.html`** desde este repositorio. Opciones:
   - Botón verde `Code → Download ZIP` y extrae `visor.html`.
   - O abre [`visor.html`](visor.html), pulsa `Raw` y guarda con clic derecho → "Guardar como…".
2. **Ábrelo con tu navegador** (doble clic). Funciona sin conexión, sin instalar nada y sin servidor. Recomendado: Chrome / Edge / Firefox actualizados.
3. Elige el módulo "Registro de Orientación" y carga tu archivo Excel.

> Los datos nunca salen de tu ordenador. SheetJS lee el archivo en el propio navegador y todo se mantiene en memoria mientras la pestaña está abierta. Al cerrarla o recargarla, se pierde lo cargado y hay que volver a abrir el Excel.

## Probarlo con datos de ejemplo

Si quieres ver cómo funciona sin necesidad de tu propio Excel:

1. Descarga también [`demo.xlsx`](demo.xlsx) del repositorio.
2. En la app, carga ese archivo en lugar del tuyo.

El demo incluye 15 alumnos ficticios con casos variados (prioritarios, alertas de inactividad, diagnósticos múltiples, actuaciones recientes y antiguas, datos de contacto incompletos…) para que puedas explorar todas las funciones del visor.

## Funcionalidades principales

- **Carga local de Excel** con dos hojas (`Sheet1` y `Ficha`).
- **Listado del alumnado** con:
  - Búsqueda por nombre, apellidos, NIA, tutor/a o especialidad.
  - Filtros visuales por Seguimiento y Situación (chips clicables).
  - Filtros desplegables por Curso, Especialidad y Curso Académico (por defecto el actual).
  - Orden por defecto Prioritario → Cerrado, con filas tintadas según el estado.
  - Contador de actuaciones por alumno.
- **Ficha del alumno** con cabecera grande, edad calculada, datos académicos (diagnóstico y campo abierto), datos de contacto colapsables organizados por persona, e histórico de actuaciones expandible.
- **Histórico de actuaciones** con icono clicable de adjunto, apertura del adjunto en ventana emergente y orden cronológico invertible.
- **Pestaña Estadística** con 6 bloques: distribución por seguimiento, actuaciones del mes/curso, pendientes propios, alertas de prioritarios sin actuación reciente, distribuciones por curso/especialidad/diagnóstico y actividad mensual del curso.
- **Generación de PDF** mediante el diálogo del navegador: ficha completa del alumno, actuación individual o panel de estadísticas. Cabecera personalizable con nombre del centro, fecha y nota de confidencialidad.
- **Aviso de pérdida de datos** al volver al inicio.

Consulta el detalle completo de cambios desde el botón `?` (esquina inferior derecha) → "Ver novedades", o en el [Changelog dentro de la app](#versionado).

## Estructura del Excel esperado

### Hoja `Sheet1` (registro generado por Microsoft Forms)

Columnas:

```
Id | Hora de inicio | Hora de finalización | Correo electrónico | Nombre |
NIA | Curso Académico | Tipo de Entrada | Título/Asunto | Descripción |
Profesional que registra | Enlace a documento (subir pdf) | Otros
```

- La **Fecha** se calcula automáticamente desde `Hora de inicio`.
- El **Curso Académico** se calcula automáticamente desde la fecha (no es necesario tener la columna).
- El **profesional** se lee de la columna `Nombre` (autorrellenada por el Forms).
- `Enlace a documento (subir pdf)` admite el enlace generado por Forms al adjuntar archivos.

### Hoja `Ficha` (mantenida manualmente)

Columnas:

```
NIA | Nombre | Apellidos | Fecha de nacimiento | Curso | Especialidad |
Centro ERG | Teléfono Alumno | Mail Alumno |
Nombre Madre | Teléfono Madre | Mail Madre |
Nombre Padre | Teléfono Padre | Mail Padre |
Tutor/a | Diagnóstico | Seguimiento | Situación | Campo Abierto
```

Valores admitidos:

- **Curso**: 1 EEM, 2 EEM, 3 EEM, 4 EEM, 1 EPM, 2 EPM, 3 EPM, 4 EPM, 5 EPM, 6 EPM
- **Seguimiento**: Prioritario, Activo, Pausa, Cerrado
- **Situación**: Espera de actuación/respuesta nuestra, En espera de actuación/respuesta de otros, Ninguno
- **Diagnóstico**: texto libre. Admite múltiples valores separados por coma o punto y coma (p. ej. `TDAH, ansiedad`) — el contador de la pestaña de Estadística los cuenta por separado.
- La **edad** se calcula automáticamente desde la fecha de nacimiento.

## Privacidad

- Toda la lectura del Excel ocurre en el navegador del usuario gracias a [SheetJS](https://sheetjs.com/) cargado desde cdnjs.
- La app **no escribe** en el Excel: es de solo lectura.
- No hay servidor, ni base de datos, ni tracking, ni telemetría. Cierras la pestaña y no queda nada.

## Versionado

- **Versión actual: 1.0**
- Consulta el changelog completo desde dentro de la app: botón `?` (esquina inferior derecha) → `Ver novedades`.
- Releases publicados en [`/releases`](https://github.com/JLMirallesB/visor/releases).

## Créditos

Diseñado por **José Luis Miralles Bono** con ayuda de Claude (Anthropic).

Si te resulta útil, puedes [invitarme a una orchata en Ko-fi](https://ko-fi.com/miralles) ☕.
