---
title: "VISOR — Ejemplo ficticio para la comprobación del montaje"
lang: es
---

# Ejemplo ficticio para comprobar el montaje

> **Estos datos no corresponden a ninguna persona real.** Están inventados para comprobar que el libro, el formulario y el visor se entienden entre sí. En uso real, los datos de la ficha y del informe se obtienen de **ITACA**; aquí no se usa ningún caso real, porque el ejemplo va a pasar por manos de varias personas durante la formación.

El ejemplo recorre el sistema entero con **un solo alumno**: una ficha, un informe de orientación y una actuación registrada por el formulario. Al cargarlo en el visor deben encenderse todas las piezas a la vez.

Orden de trabajo: primero la ficha, después el informe, y al final la actuación desde el formulario.

\newpage

## 1. Hoja `Ficha` — una fila

Copiar en la **fila 2** de la hoja `Ficha`. Las columnas con desplegable están marcadas con ▾: en esas, elegir el valor en lugar de teclearlo.

| Columna | Valor |
|:--|:--|
| NIA | `99000001` |
| Nombre | Aitana |
| Apellidos | Ferrandis Oltra |
| Fecha de nacimiento | 14/03/2010 |
| Curso ▾ | 3 EPM |
| Especialidad ▾ | Clarinete |
| Centro ERG | IES Exemple |
| Teléfono Alumno | 600000001 |
| Mail Alumno | aitana.ferrandis@exemple.es |
| Nombre Madre | Empar Oltra Sanchis |
| Teléfono Madre | 600000002 |
| Mail Madre | empar.oltra@exemple.es |
| Nombre Padre | Vicent Ferrandis Beltran |
| Teléfono Padre | 600000003 |
| Mail Padre | *(dejar vacío a propósito)* |
| Tutor/a | Joan Pastor Micó |
| Diagnóstico | TDAH, ansiedad |
| Seguimiento ▾ | Prioritario |
| Situación ▾ | Espera de actuación/respuesta nuestra |
| Campo Abierto | Rinde por debajo de su nivel técnico en las audiciones con público.¶Se acordó con la familia revisar la carga de repertorio en noviembre. |
| Activada UEO? ▾ | Sí |
| Autorización para Contactar ERG? ▾ | Sí |
| Contactado ERG? ▾ | En Proceso |

Tres cosas están puestas a propósito:

- **`Mail Padre` se queda vacío**, para ver cómo el visor muestra un campo sin dato en lugar de dejar un hueco confuso.
- **`Diagnóstico` lleva dos valores separados por coma**, para ver cómo la estadística los cuenta por separado.
- **`Campo Abierto` lleva el símbolo `¶`**, que el visor convierte en un salto de línea. En el Excel se escribe en una sola línea; es el visor el que la parte.

\newpage

## 2. Hoja `Informes` — una fila

Copiar en la **fila 2** de la hoja `Informes`. El NIA debe ser **el mismo** que el de la ficha: es lo que une las dos hojas.

> **Quién escribe este informe.** No lo redacta el conservatorio. Lo redacta el departamento de orientación del **centro de régimen general** donde cursa la enseñanza obligatoria —el *Centro ERG* de la ficha—, y está escrito desde la perspectiva de ese centro: sus aulas, sus materias, sus pruebas. El conservatorio lo recibe y lo vuelca a esta hoja tal como llega.

| Columna | Valor |
|:--|:--|
| NIA | `99000001` |
| Fecha Informe | 19/06/2026 |
| Barreras Acceso | Le cuesta acceder a la información cuando se presenta solo de forma oral y extensa. Necesita el enunciado escrito y con las instrucciones destacadas. |
| Fortalezas Acceso | Buena comprensión lectora y buena retención cuando la información se presenta en fragmentos cortos. |
| Barreras Participación | Participa poco en las actividades de gran grupo y evita las exposiciones orales ante la clase. Se muestra inhibida en los tiempos no estructurados: entradas, cambios de aula y patio. |
| Fortalezas Participación | Buena relación con el grupo en tareas por parejas y en pequeño grupo. Acepta bien la mediación de la tutora. |
| Barreras Aprendizaje | Pierde la atención en tareas largas que no tienen pautas intermedias. Dificultad para planificar el trabajo de varias sesiones: entrega tareas incompletas y olvida el material. |
| Fortalezas Aprendizaje | Constante cuando los objetivos son cortos y comprobables. Responde bien al refuerzo positivo inmediato y a la agenda visual. |
| Necesidades específicas de apoyo educativo | Adaptaciones metodológicas y de acceso: segmentación de las tareas en pasos, tiempo adicional en las pruebas escritas y puesto preferente en el aula.¶Agenda visual y revisión semanal de la planificación con la tutora.¶Seguimiento por la UEO y coordinación periódica con la familia. |
| Justificación y orientaciones | Las dificultades de atención y de planificación, junto con una respuesta ansiosa ante la exposición pública, explican el desajuste entre su capacidad y su rendimiento académico.¶Se recomienda mantener las adaptaciones metodológicas en todas las materias, evitar las exposiciones orales no preparadas y anticipar los cambios de rutina.¶Revisión del informe al final del curso por el departamento de orientación del centro. |

Los dos campos largos llevan varios `¶`. Así se ve para qué sirve: al volcar un informe desde un PDF los saltos de párrafo se pierden y todo queda en una sola línea. Poniendo `¶` donde iba el salto, el visor lo restituye al mostrarlo y al imprimirlo.

\newpage

## 3. Una actuación desde el formulario

Esto **no se escribe en el Excel**: se rellena el formulario como se hará a diario. La fila debe aparecer sola en la hoja `Sheet1`.

| Pregunta | Respuesta |
|:--|:--|
| NIA | `99000001` |
| Tipo de Entrada | Reunión |
| Título/Asunto | Reunión con la familia y el tutor |
| Descripción | Se revisa el informe de orientación con la madre y el tutor del conservatorio. Se acuerda una audición interna de aula antes de la audición pública de diciembre y retirar una obra del programa. |
| Enlace a documento (subir pdf) | *(opcional: sube cualquier PDF para probar el adjunto)* |
| Otros | Pendiente de enviar al IES el resumen de los acuerdos. |

No hay que indicar la fecha ni el curso académico: los calcula el visor a partir de la hora en que se envía el formulario.

\newpage

## 4. Qué debe verse en el visor

Abrir `visor.html`, cargar el libro y comprobar estas siete cosas. Si alguna falla, el apartado correspondiente del montaje dice dónde mirar.

| # | Qué comprobar |
|:--|:--|
| 1 | En el listado, **Ferrandis Oltra, Aitana (99000001)** con el chip **Prioritario**, la fila tintada y un **1** en la columna *Act.* |
| 2 | En su ficha, la **edad calculada** a partir del 14/03/2010 — nadie la ha escrito |
| 3 | En *Campo Abierto*, el texto **partido en dos líneas** donde estaba el `¶` |
| 4 | El botón **📄 Ver informe de orientación**, con la fecha **08/09/2026** y los seis campos de barreras y fortalezas |
| 5 | En el histórico, la actuación **con la fecha de hoy**, la etiqueta **Reunión** en color, y **tu nombre** a la derecha de la línea |
| 6 | En *Estadística*: suma 1 en **Prioritario**, aparece en **pendientes propios** y **no** aparece en alertas (tiene actuación reciente) |
| 7 | En *Estadística*, el diagnóstico cuenta **dos** entradas —*TDAH* y *ansiedad*—, no una |

Si el **nombre del profesional llega vacío** en el punto 5, el formulario no está registrando la identidad: hay que revisar los ajustes de respuestas restringidas a la organización.

Si el **curso académico** de la actuación no sale como `26-27`, revisar que la columna *Hora de inicio* llega con fecha.

\newpage

## 5. Al terminar

Borrar la fila de prueba **de `Ficha` y de `Informes`**, y **dejar la de `Sheet1` donde está**.

Al no existir ya ese NIA en la ficha, el visor no muestra la actuación en ninguna parte: queda huérfana e invisible. Es preferible a borrar filas de la hoja que escribe Forms, que es la única regla que no conviene romper nunca.
