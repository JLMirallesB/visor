# VISOR

Visor offline de registros de orientación para conservatorios profesionales de música y danza. Lee archivos Excel generados por Microsoft Forms y muestra fichas del alumnado, histórico de actuaciones y estadísticas — todo en local, sin enviar datos a ningún servidor.

## ¿Qué es?

Un único archivo HTML que abres en el navegador. Eliges un Excel local con dos hojas obligatorias (`Sheet1` con las actuaciones del Forms y `Ficha` con los datos del alumnado) y dos opcionales (`Informes` y `Anuladas`) y la app te muestra:

- Listado del alumnado con búsqueda multicampo y filtros por seguimiento, situación, curso, especialidad y curso académico.
- Ficha individual con datos académicos, contactos, diagnóstico e histórico de actuaciones.
- Pestaña de estadísticas con la carga del departamento (prioritarios, pendientes, alertas, actividad mensual, distribuciones, diagnósticos).
- Impresión / guardado como PDF de fichas, actuaciones y estadísticas, con cabecera personalizable.

## Cómo usarlo

1. **Descarga `visor.html`** desde este repositorio. Opciones:
   - Botón verde `Code → Download ZIP` y extrae `visor.html`.
   - O abre [`visor.html`](visor.html), pulsa `Raw` y guarda con clic derecho → "Guardar como…".
2. **Ábrelo con tu navegador** (doble clic). Sin instalar nada y sin servidor. La primera vez necesita conexión para traer la librería que lee el Excel ([SheetJS](https://sheetjs.com/)); después el navegador suele tenerla en caché. Recomendado: Chrome / Edge / Firefox actualizados.
3. Elige el módulo "Registro de Orientación" y carga tu archivo Excel.

> Los datos nunca salen de tu ordenador. SheetJS lee el archivo en el propio navegador y todo se mantiene en memoria mientras la pestaña está abierta. Al cerrarla o recargarla, se pierde lo cargado y hay que volver a abrir el Excel.

## Probarlo con datos de ejemplo

Si quieres ver cómo funciona sin necesidad de tu propio Excel:

1. Descarga también [`demo.xlsx`](demo.xlsx) del repositorio.
2. En la app, carga ese archivo en lugar del tuyo.

También hay un [`ejemplo-ficticio.docx`](ejemplo-ficticio.docx) con los textos de **un solo caso inventado** —una ficha, un informe y una actuación— para comprobar tu propio montaje de principio a fin: que el formulario escribe donde debe, que las tres hojas se entienden por el NIA y que el visor lo muestra todo. Incluye las siete comprobaciones a hacer.

El demo incluye 15 alumnos ficticios con casos variados (prioritarios, alertas de inactividad, diagnósticos múltiples, actuaciones recientes y antiguas, datos de contacto incompletos…) para que puedas explorar todas las funciones del visor.

## Montar el sistema en tu centro

El sistema son tres piezas: **Microsoft Forms** escribe, un **libro de Excel en OneDrive** almacena, y `visor.html` lee. El visor no se conecta a OneDrive ni a SharePoint: lee un archivo que esté en el disco del ordenador.

Lo que sigue es el montaje completo. **El orden importa**, y por un motivo concreto: el emparejado entre un formulario y su libro de respuestas nace con la creación y **no se puede establecer después** — no existe en ninguna pantalla un «conectar este formulario a este libro». Y el libro nace donde vive el formulario.

### 1. Decidir quién es el propietario

Que el formulario pertenezca a una **cuenta institucional del centro**, no a la de una persona: si pertenece a una cuenta personal, el formulario, el libro y los archivos adjuntos se van con ella cuando esa persona deja el centro. Si en tu organización puedes crear un **grupo privado** con solo las personas del círculo de acceso, es aún más limpio, porque los permisos se gestionan por pertenencia.

> ⚠️ **No lo crees como «formulario de grupo» de un equipo de Teams amplio.** Un formulario de grupo lo pueden abrir, editar y leer completo **todos los miembros del grupo**, respuestas y adjuntos incluidos. Si el equipo es el claustro, el registro de orientación queda a la vista del profesorado.

### 2. Decidir quién puede escribir

Lo que se registra son datos de salud de menores. Y el enlace del formulario es el único control que existe: **no caduca, no se revoca por persona y se reenvía sin dejar rastro**. Repartirlo al claustro entero es, en la práctica, publicarlo.

La recomendación es dejarlo en **dos o tres personas nombradas** — por ejemplo jefatura de estudios y coordinación de orientación/inclusión. Quien detecte algo lo traslada por la vía ordinaria y lo registra quien tiene el enlace. Se pierde inmediatez y se gana un registro homogéneo y un círculo de acceso pequeño y nombrable.

### 3. Crear la carpeta

En el OneDrive de la cuenta propietaria, una carpeta para todo el sistema. Dos criterios: **sin tildes ni espacios**, porque el nombre acaba formando parte de una ruta de SharePoint; y **sin año ni curso** en el nombre, porque el registro es continuo.

Los adjuntos del formulario **no** caerán aquí: Forms usa una carpeta propia (paso 8).

### 4. Crear el formulario: seis preguntas

Con la sesión de la cuenta propietaria abierta en forms.office.com. **El título de cada pregunta se convierte en el nombre de su columna**, así que estos son los textos literales, en este orden:

```
NIA
Tipo de Entrada
Título/Asunto
Descripción
Enlace a documento (subir pdf)
Otros
```

| Pregunta | Tipo | Notas |
|---|---|---|
| NIA | Texto | Obligatoria. Es la llave que une las tres hojas |
| Tipo de Entrada | Opción | Obligatoria, **de una sola respuesta** |
| Título/Asunto | Texto | Obligatoria |
| Descripción | Texto, respuesta larga | |
| Enlace a documento (subir pdf) | Carga de archivos | Solo en cuentas de organización |
| Otros | Texto | |

Las siete opciones de `Tipo de Entrada` que el visor muestra **con color propio**: `Reunión`, `Derivación`, `Comunicación familia`, `Informe`, `Decisión`, `Suceso`, `Otro`. Puedes añadir más: funcionan igual, solo pierden el color.

**No preguntes la fecha, ni el curso académico, ni quién registra.** Los tres los pone el sistema: los dos primeros los calcula el visor desde `Hora de inicio`, y el tercero lo autorrellena Forms.

### 5. Los dos ajustes que sostienen la autoría

En la configuración del formulario, **antes de crear el libro**:

- Respuestas **restringidas a la organización**.
- **Registrar el nombre** de quien responde.

Sin los dos, las columnas `Nombre` y `Correo electrónico` llegan vacías y el registro pierde el dato de **quién hizo cada actuación**, que es la mitad de su valor. Y el orden importa: las columnas del libro se generan a partir del formulario, así que un libro creado mientras las respuestas son anónimas nace sin ellas.

> ⚠️ **Que nadie rellene el formulario con la sesión de la cuenta institucional abierta.** La columna `Nombre` recoge a quien **envía** la respuesta: enviada desde la cuenta del centro, la actuación queda firmada por el centro y se pierde quién la hizo.

### 6. Dejar que Forms cree el libro

**Respuestas → Abrir en Excel.** El libro nace sincronizado en vivo —cada respuesta nueva aparece sola, sin exportar nada— y con la hoja de respuestas llamada `Sheet1`, que es el nombre que busca el visor.

**No lo crees tú.** Excel para la web permite lo contrario, crear el libro primero y engancharle un formulario, pero un libro en blanco creado desde una interfaz en castellano tiene la hoja llamada `Hoja1`.

Después, **mueve el libro a tu carpeta**. El vínculo con el formulario va por identificador del elemento y no por ruta, así que sobrevive a un movimiento dentro del mismo OneDrive. Hazlo antes de recoger respuestas reales.

> Si el libro no se actualiza solo y hay que volver a pedir *Abrir en Excel* cada vez, no tienes sincronización en vivo y el resto del proceso no funciona como está descrito.

### 7. Las hojas `Ficha`, `Informes`, `Anuladas` y `Listas`

Se trasplantan desde `plantilla-hojas.xlsx` — ver [Estructura del Excel esperado](#estructura-del-excel-esperado). No las crees a mano.

### 8. El visor, los permisos y la sincronización

Guarda `visor.html` en la misma carpeta que el libro, y comparte esa carpeta con las personas del círculo **con permiso de edición** (tienen que escribir en `Ficha`). Cada una, en su ordenador: **Añadir acceso directo a Mis archivos** → si hay sincronización selectiva, **Elegir carpetas** y marcarla → clic derecho → **Mantener siempre en este dispositivo**.

A partir de ahí: doble clic en `visor.html`, elegir el libro de al lado, consultar. Sin exportar ni descargar nada.

**Y el permiso que nadie adivina: la carpeta de los adjuntos.** Forms no guarda los archivos subidos en tu carpeta, sino en una ruta propia del OneDrive de la cuenta propietaria:

```
Aplicaciones  ›  Microsoft Forms  ›  [nombre del formulario]  ›  [pregunta]
```

Quien tenga acceso a la carpeta del registro pero no a esta verá el libro, editará la ficha y usará el visor con normalidad… y al pulsar el icono de adjunto recibirá un «no tienes acceso». Es un fallo parcial y tardío: todo parece funcionar hasta el día en que hace falta leer precisamente ese documento.

Tres cosas que ayudan:

- Para llegar a la carpeta: en el formulario, **Respuestas → la pregunta de archivos → Más detalles → Ver en carpeta**. Mejor que navegar a mano, porque **la carpeta conserva el nombre que tenía el formulario cuando se creó la pregunta**, no el actual.
- **No existe hasta la primera subida**: comparte después de haber adjuntado algo.
- Comparte **la carpeta del nombre del formulario**, no la de la pregunta, con **lectura**: así una segunda pregunta de archivos heredará el permiso.

Microsoft no documenta el modelo de permisos de estas carpetas, así que la única comprobación que vale es que **la otra persona abra un adjunto desde su propia cuenta**. Que lo abra quien concedió el permiso no prueba nada.

### 9. Comprobar que funciona

Con [`ejemplo-ficticio.docx`](ejemplo-ficticio.docx): una ficha, un informe y una actuación inventados, y las siete comprobaciones a hacer en el visor. Al terminar se borra la fila de prueba de `Ficha` y de `Informes`, y **se deja la de `Sheet1`**: sin ficha, el visor no la muestra en ninguna parte, y así no hay que borrar filas de la hoja que escribe Forms.

### Reglas que no conviene romper

1. **`Sheet1` no se toca**: la escribe Forms. Ni renombrar columnas, ni reordenarlas, ni insertar una en medio, ni ordenar las filas. Si hace falta un cálculo, va en otra hoja.
   **Y borrar filas ahí no funciona**: el vínculo con Forms solo añade, y las restaura en cuanto llega la siguiente respuesta. Borrar la respuesta dentro de Forms tampoco las quita. Para retirar una actuación, su `Id` va a la hoja `Anuladas`.
2. **No se edita una copia descargada.** Las hojas mantenidas a mano se editan en Excel para la web, o abriendo el archivo de la carpeta sincronizada. Una copia aparte no vuelve a ningún sitio.
3. **El formulario no se recrea.** Uno nuevo nace con un libro nuevo y vacío, y `Ficha` e `Informes` se quedan en el viejo.
4. **El formulario no se abre a respuestas anónimas**, o se pierde la autoría.
5. **Nombres de columna exactos** en `Ficha` e `Informes`: un nombre distinto no da error, da una columna vacía.

### Lo que hay que revisar cada curso

- **`Ficha`**: alta del alumnado nuevo y repaso del `Seguimiento` del que ya estaba. Un caso antiguo que sigue en `Prioritario` dispara alertas falsas.
- **Las fichas pendientes de completar** (el conmutador del listado), hasta vaciar la lista.
- **Los permisos de las dos carpetas** —registro y adjuntos—, retirándolos a quien ya no esté en el círculo de acceso: mientras los tenga, el libro está en su disco.
- **Que la sincronización sigue viva**: una respuesta de prueba y ver si aparece sola.
- **`visor.html`**, sustituyéndolo si se ha publicado una versión nueva. Le llega a todo el que sincroniza.

## Funcionalidades principales

- **Carga local de Excel** con dos hojas obligatorias (`Sheet1` y `Ficha`) y dos opcionales (`Informes` y `Anuladas`).
- **Listado del alumnado** con:
  - Búsqueda por nombre, apellidos, NIA, tutor/a o especialidad.
  - Filtros visuales por Seguimiento y Situación (chips clicables).
  - Filtros desplegables por Curso, Especialidad y Curso Académico (por defecto el actual).
  - Orden por defecto Prioritario → Cerrado, con filas tintadas según el estado.
  - Contador de actuaciones por alumno.
- **Ficha del alumno** con cabecera grande, edad calculada, datos académicos (diagnóstico y campo abierto), datos de contacto colapsables organizados por persona, e histórico de actuaciones expandible.
- **Histórico de actuaciones** con icono clicable de adjunto, apertura del adjunto en ventana emergente y orden cronológico invertible.
- **Pestaña Estadística** con 6 bloques: distribución por seguimiento, actuaciones del mes/curso, pendientes propios, alertas de prioritarios sin actuación reciente, distribuciones por curso/especialidad/diagnóstico y actividad mensual del curso.
- **Informe de orientación** (si existe la hoja `Informes`): botón desplegable en la ficha con barreras y fortalezas de acceso, participación y aprendizaje, necesidades específicas y justificación, e impresión independiente en tipografía compacta.
- **Columnas de UEO y ERG** en la ficha, con chips de estado (Sí / No / En Proceso).
- **Anulación de actuaciones** listando su `Id` en la hoja `Anuladas`: desaparecen del histórico, del contador y de la estadística, sin tocar la hoja que escribe Forms.
- **Toggle «Incluir fichas pendientes de completar»** para mostrar u ocultar las fichas que solo tienen NIA (ocultas por defecto).
- **Generación de PDF** mediante el diálogo del navegador: ficha completa del alumno, actuación individual o panel de estadísticas. Cabecera personalizable con nombre del centro, fecha y nota de confidencialidad.
- **Aviso de pérdida de datos** al volver al inicio.

Consulta el detalle completo de cambios desde el botón `?` (esquina inferior derecha) → "Ver novedades", o en el [Changelog dentro de la app](#versionado).

## Estructura del Excel esperado

> **Atajo para no cometer erratas:** descarga [`plantilla-hojas.xlsx`](plantilla-hojas.xlsx), que trae cuatro hojas listas: `Ficha` con sus cabeceras y **siete columnas con desplegable**, `Informes` y `Anuladas` con las suyas, y `Listas` con los valores admitidos (incluidas las **40 especialidades oficiales**, 36 de música y 4 de danza). El visor ignora la hoja `Listas`.
>
> En el **Excel de escritorio**, con los dos libros abiertos: clic derecho en la pestaña de la plantilla → **Mover o copiar…** → *Al libro*: tu libro → marcar **Crear una copia**. Las cuatro, y `Listas` no es opcional: los desplegables de `Ficha` apuntan a sus rangos. Así las hojas llegan enteras, con su nombre ya puesto, y no hay que crearlas antes — si ya existe una hoja `Ficha`, la copia entra como `Ficha (2)` y el visor no la encuentra.
>
> En **Excel para la web** no existe *Mover o copiar* entre libros: hay que crear las hojas a mano y pegar **toda la hoja** de la plantilla en `A1`, empezando por `Listas` para que las validaciones encuentren sus rangos, y comprobando después que los desplegables han sobrevivido.
>
> Para **recortar** las listas a lo que imparte tu centro, borra **filas enteras** de la hoja `Listas`; para **añadir**, inserta una fila **dentro** de la lista, no después de la última. Los rangos se ajustan solos al insertar o borrar filas, pero una celda vaciada deja un hueco en el desplegable y una fila añadida al final se queda fuera.
>
> Conviene, porque las columnas de `Ficha` se leen **por nombre exacto**: un `Telefono Alumno` sin tilde o un `Tutor` sin `/a` dejan la columna vacía sin avisar de nada. Las de `Sheet1`, en cambio, admiten variantes.

### Hoja de registro (la que genera Microsoft Forms)

**El nombre de esta hoja no importa.** Forms la bautiza de maneras distintas según cómo se haya creado el libro y en qué idioma: `Sheet1`, `Hoja1`, `Form1`. El visor la busca primero por esos nombres y, si ninguno coincide, **la reconoce por sus columnas**: la hoja que tenga `Id` y `Hora de inicio` es el registro.

Las cinco primeras columnas las genera Forms solo. El resto son, en este orden, las preguntas del formulario:

```
Id | Hora de inicio | Hora de finalización | Correo electrónico | Nombre |
NIA | Tipo de Entrada | Título/Asunto | Descripción |
Enlace a documento (subir pdf) | Otros
```

- El **profesional que registra** se lee de la columna `Nombre`, que Forms autorrellena con la identidad de quien envía la respuesta. **No hace falta preguntarlo en el formulario** — para que venga rellena, el formulario debe estar restringido a la organización y no admitir respuestas anónimas.
- La **Fecha** se calcula automáticamente desde `Hora de inicio`.
- El **Curso Académico** se calcula automáticamente desde la fecha: no hace falta ni la columna ni la pregunta.
- `Enlace a documento (subir pdf)` admite el enlace generado por Forms al adjuntar archivos. Abrirlo exige sesión iniciada en Microsoft 365, porque el archivo se queda en el OneDrive del propietario del formulario.
- El `Correo electrónico` se guarda en el libro pero **no se muestra en ninguna pantalla del visor**.
- Si el libro arrastra columnas de versiones anteriores (`Curso Académico`, `Profesional que registra`), no estorban: mientras `Nombre` y la fecha vengan rellenas, el visor no las usa.

### Hoja `Ficha` (mantenida manualmente)

Columnas:

```
NIA | Nombre | Apellidos | Fecha de nacimiento | Curso | Especialidad |
Centro ERG | Teléfono Alumno | Mail Alumno |
Nombre Madre | Teléfono Madre | Mail Madre |
Nombre Padre | Teléfono Padre | Mail Padre |
Tutor/a | Diagnóstico | Seguimiento | Situación | Campo Abierto |
Activada UEO? | Autorización para Contactar ERG? | Contactado ERG?
```

Valores admitidos:

- **Curso**: 1 EEM, 2 EEM, 3 EEM, 4 EEM, 1 EPM, 2 EPM, 3 EPM, 4 EPM, 5 EPM, 6 EPM
- **Seguimiento**: Prioritario, Activo, Pausa, Cerrado
- **Situación**: Espera de actuación/respuesta nuestra, En espera de actuación/respuesta de otros, Ninguno
- **Especialidad**: las 40 del catálogo oficial — 36 de las enseñanzas profesionales de música ([Decreto 158/2007](https://dogv.gva.es/es/eli/es-vc/d/2007/09/21/158), art. 6, texto consolidado) y 4 de danza ([Decreto 156/2007](https://dogv.gva.es/es/eli/es-vc/d/2007/09/21/156), art. 6). Se respeta la grafía del decreto, que mezcla idiomas: `Cant valencià`, `Violoncello`. El visor no las valida —acepta cualquier texto— pero mantenerlas uniformes evita duplicados en los filtros y en la estadística.
- **Diagnóstico**: texto libre. Admite múltiples valores separados por coma o punto y coma (p. ej. `TDAH, ansiedad`) — el contador de la pestaña de Estadística los cuenta por separado.
- **Activada UEO?**, **Autorización para Contactar ERG?** y **Contactado ERG?**: Sí, No, En Proceso.
- La **edad** se calcula automáticamente desde la fecha de nacimiento.
- Una fila que solo tenga el NIA, sin nombre ni apellidos, se considera **ficha pendiente de completar** y queda oculta en el listado salvo que se active el toggle correspondiente.

### Hoja `Informes` (opcional)

Una fila por informe de orientación, vinculada al alumnado por el NIA.

```
NIA | Fecha Informe |
Barreras Acceso | Fortalezas Acceso |
Barreras Participación | Fortalezas Participación |
Barreras Aprendizaje | Fortalezas Aprendizaje |
Necesidades específicas de apoyo educativo | Justificación y orientaciones
```

- Si hay **varios informes con el mismo NIA**, el visor muestra solo el de `Fecha Informe` más reciente. Los anteriores permanecen en la hoja, pero no se ven.
- Si un NIA aparece aquí y **no** en `Ficha`, el visor crea una ficha mínima para que el informe no se pierda. Aparece como ficha pendiente de completar.
- El símbolo **¶** vale como salto de línea en cualquier campo de texto del libro. Es útil al volcar informes desde un PDF, donde los saltos de párrafo se pierden.

### Hoja `Anuladas` (opcional)

Una lista de los `Id` de las actuaciones que el visor **debe dejar de mostrar**.

```
Id | Motivo | Fecha
```

- El `Id` es el que Forms asigna a cada respuesta, la primera columna de `Sheet1`.
- Las actuaciones anuladas desaparecen del histórico del alumno, del contador de actuaciones y de toda la estadística, como si no se hubieran registrado.
- Se anula por `Id` y **no borrando la fila** de `Sheet1` porque borrarla no funciona: el vínculo entre Forms y el libro **solo añade**, y restaura las filas borradas en cuanto llega la siguiente respuesta. Borrar la respuesta dentro de Forms tampoco quita la fila del libro.
- `Motivo` y `Fecha` no los lee el visor: están para que quede constancia de por qué se retiró una actuación. En un registro que puede acabar justificando decisiones, eso vale más que una fila desaparecida.

## Privacidad

Del lado del visor:

- Toda la lectura del Excel ocurre en el navegador del usuario gracias a [SheetJS](https://sheetjs.com/) cargado desde cdnjs.
- La app **no escribe** en el Excel: es de solo lectura.
- No hay servidor, ni base de datos, ni tracking, ni telemetría. Cierras la pestaña y no queda nada — ni siquiera el nombre del centro que escribes para la cabecera del PDF, y por eso lo pregunta cada vez.

Del lado de tu centro, que es donde está el riesgo real. El libro contiene **datos de salud de menores** —diagnósticos, informes de orientación—, categoría especial del art. 9 del RGPD, y hay dos puntos por donde salen del entorno protegido:

- **El archivo en el disco.** Al sincronizar la carpeta con *Mantener siempre en este dispositivo*, el libro con los diagnósticos está físicamente en ese ordenador, de forma permanente. Actívalo **solo en equipos del centro, con cifrado de disco y sesión con contraseña**, y solo para las personas del círculo de acceso. Cuando alguien sale del círculo, retírale los permisos: el cliente de OneDrive le borrará la copia local en la siguiente sincronización.
- **Los PDF que se imprimen.** Una vez impresos son papel, o un fichero suelto como cualquier otro, fuera de todo control de acceso.

Aun así, la carpeta sincronizada es más segura que la alternativa de descargar una copia del libro cada vez: hay **un único archivo gestionado** en lugar de una copia por consulta acumulándose en la carpeta de Descargas, que alguien tiene que acordarse de borrar.

## Versionado

- **Versión actual: 1.3**
- Consulta el changelog completo desde dentro de la app: botón `?` (esquina inferior derecha) → `Ver novedades`.
- Releases publicados en [`/releases`](https://github.com/JLMirallesB/visor/releases).

## Créditos

Diseñado por **José Luis Miralles Bono** con ayuda de Claude (Anthropic).

Si te resulta útil, puedes [invitarme a una orchata en Ko-fi](https://ko-fi.com/miralles) ☕.
