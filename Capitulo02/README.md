# Diagnóstico de ausentismo y detección de áreas de mejora.

## 1. Metadatos del Laboratorio

| Atributo | Detalle |
| :--- | :--- |
| **Duración** | 90 minutos |
| **Complejidad** | Intermedia |
| **Audiencia** | Gerentes de relaciones laborales, generalistas de recursos humanos, coordinadores de personal, analistas de desarrollo organizacional y asesores legales internos. |
| **Tecnologías** | Microsoft Copilot, Microsoft Excel, Microsoft Word y Microsoft PowerPoint. |
| **Enfoque** | Análisis de datos operativos, diagnóstico de causas raíz, toma de decisiones basada en indicadores y comunicación estratégica multinivel. |

---

## 2. Descripción Corta

Este laboratorio práctico de 90 minutos capacita a los profesionales de gestión humana para diagnosticar y abordar el ausentismo laboral de manera proactiva utilizando Microsoft Copilot como copiloto analítico. A través de un flujo estructurado de 7 pasos, los participantes aprenderán a preparar plantillas de datos en Excel, calcular de forma didáctica el Índice de Bradford, redactar informes ejecutivos en Word, generar imágenes de soporte para campañas internas, estructurar presentaciones en PowerPoint y diseñar KPIs de control para el área de relaciones laborales.

---

## 3. Objetivos del Laboratorio

Al finalizar este laboratorio, el estudiante será capaz de:
* **Estructurar y limpiar bases de datos de asistencia** en Excel utilizando la guía de la IA.
* **Comprender y aplicar el Índice de Bradford** para evaluar el impacto operativo de las faltas repetitivas.
* **Redactar informes de auditoría laboral estructurados** en Word con un tono neutral y preventivo.
* **Diseñar presentaciones ejecutivas en PowerPoint** que condensen propuestas de mejora para la alta dirección.
* **Crear apoyos visuales de comunicación** preventiva utilizando el generador de imágenes integrado de Copilot.

---

## 4. Prerrequisitos

* Cuenta activa de **Microsoft 365** con acceso a **Microsoft Copilot**.
* Aplicación de **Microsoft Excel** abierta.
* Aplicación de **Microsoft Word** abierta.
* Aplicación de **Microsoft PowerPoint** abierta.

---

## 5. Procedimiento Paso a Paso

### Paso 1: Estructuración y Limpieza de la Tabla de Asistencia en Excel

Antes de realizar cualquier análisis, debemos contar con una estructura de datos estandarizada en Excel para evitar errores de interpretación o registros inválidos.

1. Abra **Microsoft Excel** con un libro nuevo en blanco.
2. Para diseñar una tabla limpia de auditoría, abra el chat de **Microsoft Copilot** e introduzca el siguiente prompt:

```
Actúa como un Analista de Datos de Recursos Humanos experto en Excel. Necesito estructurar una tabla de registro de ausentismo mensual para una planta industrial que tiene 15 colaboradores en el área de empaque.

Dime exactamente:
1. Qué nombres de columnas debe tener mi tabla en Excel para capturar adecuadamente la información de ausentismo (necesito columnas clave como ID, Nombre, Departamento, Turno, Fecha de Inicio de la Ausencia, Días Totales, Motivo y si fue Justificada o no).
2. Proporcióname una lista de 5 motivos de ausentismo estandarizados que ayuden a clasificar las causas de forma rápida para un catálogo de validación de datos.
3. Dame 3 ejemplos prácticos en texto de cómo se verían tres filas de esta tabla ya llenas con datos realistas de faltas (una por enfermedad común, otra por causa personal injustificada y otra por accidente de trayecto).

Presenta la respuesta de forma clara y con los datos ordenados en una tabla de texto.
```

3. En su archivo de Excel, cree las columnas y capture los 3 registros de ejemplo sugeridos por Copilot para simular la estructura. Guarde el archivo como `Analisis_Ausentismo_Q2.xlsx`.

---

### Paso 2: Comprensión y Cálculo Didáctico del Índice de Bradford

El Índice de Bradford es una herramienta excelente para recursos humanos porque demuestra que **muchas faltas cortas e imprevistas causan más desorden operativo que una sola falta larga y programada**. 

Para entenderlo de forma muy didáctica, usaremos este ejemplo:
* **Colaborador A (Falta larga):** Falta 1 vez durante 10 días seguidos (Operaciones puede planificar su reemplazo fácilmente).
* **Colaborador B (Faltas repetidas):** Falta 10 veces de a 1 día a la semana de forma imprevista (Operaciones sufre para cubrir su puesto cada semana).

La fórmula matemática es:
$$B = S^2 \times D$$

Donde:
* $S$ (*Spells*): Número de veces (frecuencias) que el empleado se ausentó.
* $D$ (*Days*): Total de días acumulados de ausencia.

1. Introduzca el siguiente prompt en Copilot para aprender a calcularlo de forma muy sencilla en Excel:

```
Actúa como un Facilitador de Recursos Humanos. Necesito explicarle a mi equipo de Relaciones Laborales cómo calcular el Índice de Bradford de forma muy sencilla en Excel.

Por favor, dame:
1. Una explicación muy clara y didáctica usando el ejemplo de "Colaborador A" (1 falta de 10 días) vs. "Colaborador B" (10 faltas de 1 día). Haz los cálculos matemáticos simples de ambos para demostrar por qué el Colaborador B tiene un puntaje mucho más crítico si los dos faltaron 10 días en total.
2. La fórmula exacta que debo escribir en la celda de Excel si tengo la frecuencia (S) en la columna B y los días totales (D) en la columna C.
3. Una tabla de semáforo sencilla (Rango de puntos, Significado y qué acción tomar con el empleado) para saber qué hacer cuando un puntaje sale alto.

Preséntalo de forma muy visual, amable y libre de tecnicismos complejos.
```

2. Copie la fórmula de Excel provista por la IA, pruébela en su hoja de cálculo con los ejemplos y guarde los rangos del semáforo.

---

### Paso 3: Redacción del Informe de Diagnóstico en Word

Una vez detectados los patrones de ausentismo en Excel, el equipo de Relaciones Laborales debe presentar un informe formal que identifique las áreas de mejora y proponga soluciones sin sesgos subjetivos.

1. Abra **Microsoft Word** con un documento en blanco.
2. Ingrese el siguiente prompt en el chat de Copilot:

```
Actúa como un Especialista en Desarrollo Organizacional y Relaciones Laborales. Necesito redactar un "Informe Ejecutivo de Diagnóstico de Ausentismo" para el Comité de Planta, basado en los hallazgos del último trimestre.

El informe debe cubrir la siguiente situación observada en la planta:
- El área de "Embarques" concentra el 65% del ausentismo injustificado de la planta, ocurriendo principalmente los días lunes en el turno matutino.
- La causa raíz detectada a través de encuestas rápidas de clima es la falta de flexibilidad en los roles de fin de semana y una sobrecarga física acumulada.

Escribe el informe bajo la siguiente estructura formal:
1. **Introducción:** Declarando el propósito del reporte.
2. **Diagnóstico de Datos:** Describiendo de manera analítica los hallazgos del área de Embarques y los picos de los días lunes.
3. **Análisis de Causa Raíz:** Explicando los factores humanos y operativos encontrados.
4. **Propuestas de Solución:** Propón tres medidas viables (por ejemplo, rotación de descansos de fin de semana, incentivo de asistencia perfecta y brigadas de ergonomía/salud ocupacional).

Usa un tono formal, constructivo, orientado al bienestar del colaborador y a la productividad del negocio.
```

3. Copie el reporte estructurado que le entregue Copilot y péguelo en su documento de Word. Guárdelo como `Reporte_Diagnostico_Ausentismo.docx`.

---

### Paso 4: Generación de Ilustración de Soporte para Campaña de Clima y Bienestar

Para acompañar el plan de acción en la planta, lanzaremos una campaña interna sobre el valor del descanso y el bienestar. Necesitamos un recurso gráfico motivacional para los boletines oficiales de la compañía.

1. Introduzca el siguiente prompt de diseño gráfico en Copilot:

```
Genera una imagen en estilo de ilustración digital plana en 2D, moderna y corporativa, que represente un "Entorno de Trabajo Saludable y de Apoyo al Trabajador".

La escena debe mostrar a un equipo diverso de trabajadores industriales (hombres y mujeres con uniformes limpios de planta y equipo de protección personal como cascos y chalecos) sonriendo y colaborando alegremente en un área común de descanso de la planta o en una oficina de Recursos Humanos moderna y luminosa. Al fondo, se debe percibir un ambiente seguro y ordenado, con plantas verdes que sugieran frescura, bienestar y equilibrio entre la vida laboral y personal. 

Usa una paleta de colores brillantes pero profesionales (como azules suaves, verdes naturales y amarillos cálidos). La imagen no debe contener textos, marcas de agua ni deformaciones anatómicas en las personas.
```

2. Copilot generará las alternativas de imagen. Descargue su versión preferida, cópiela e insértela en el archivo de Word de su reporte de diagnóstico para darle una presentación visual impecable.

---

### Paso 5: Estructuración del Guión de Presentación en PowerPoint

Llegó el momento de convencer a la Dirección para aprobar el presupuesto de las medidas de mejora. Copilot nos ayudará a estructurar el contenido diapositiva por diapositiva para nuestra presentación de PowerPoint.

1. Introduzca el siguiente prompt de diseño instruccional en Copilot:

```
Actúa como un Consultor de Comunicación Corporativa y Relaciones Laborales de Alto Impacto. Necesito estructurar una presentación ejecutiva de 5 diapositivas para convencer al Director de Operaciones de autorizar el plan de mejora de asistencia y rotación de turnos.

Escríbeme el contenido detallado para cada diapositiva de la siguiente manera:
- **Diapositiva 1 (Título):** Nombre de la presentación sugerido y subtítulo de impacto.
- **Diapositiva 2 (El Problema):** Texto clave e indicador del costo operativo y clima en Embarques.
- **Diapositiva 3 (La Causa Raíz):** Puntos concisos que expliquen por qué faltan los lunes (fatiga, roles fijos).
- **Diapositiva 4 (Nuestra Solución):** El plan de 3 acciones estratégicas (Incentivos, ergonomía y flexibilidad).
- **Diapositiva 5 (Beneficios de Negocio / ROI):** Qué ganará la planta al reducir el ausentismo (estabilidad, ahorro de horas extra, clima laboral óptimo).

Para cada diapositiva, indícame:
1. El título exacto de la diapositiva.
2. 3 viñetas con textos muy cortos e impactantes para leer en segundos.
3. Una recomendación de qué tipo de imagen de soporte usar en esa diapositiva.

Una vez finalizada dejala lista para descargar y/o exportar a PowerPoint
```

2. Copie este esquema en su archivo de Word o utilícelo para armar rápidamente las diapositivas de su presentación en **Microsoft PowerPoint** con un diseño moderno.

---

### Paso 6: Establecimiento de Indicadores Clave de Desempeño (KPIs) de Control

Un plan sin métricas de control no se puede medir. Estandarizaremos las fórmulas de control de Relaciones Laborales que usaremos para monitorear el éxito del plan.

1. Introduzca este prompt en el chat de Copilot:

```
Actúa como un Especialista en Indicadores de Gestión Humana. Necesito definir los 3 KPIs críticos de control para medir el avance mensual de nuestro Plan de Mejora de Asistencia.

Para cada KPI, proporcióname:
1. El nombre del indicador (por ejemplo, Tasa de Ausentismo General, Índice de Retención del Turno, etc.).
2. La fórmula exacta para calcularlo utilizando palabras simples (ejemplo: [Faltas Totales / Días Laborables Programados] * 100).
3. La meta recomendada para una planta industrial de manufactura.
4. Con qué frecuencia se debe reportar este indicador al área de Relaciones Laborales e Industrial.

Preséntame los 3 KPIs ordenados en una tabla comparativa de texto para que sea muy fácil de entender.
```

2. Copie la tabla comparativa de KPIs y péguela en el anexo final de su informe de diagnóstico en Word.

---

### Paso 7: Reto de Aplicación Autónoma – Plan de Intervención para el "Viernes de Ausentismo"

**Instrucciones para el estudiante:** Las auditorías recientes del departamento de Relaciones Laborales revelaron un nuevo e inusual patrón de alerta: en el área de **Mantenimiento y Prensas**, las ausencias injustificadas se concentran de manera drástica los **viernes por la tarde (turno vespertino)**, coincidiendo con el día de pago semanal en efectivo que se realizaba anteriormente, lo cual genera graves retrasos en las entregas de fin de semana.

#### El Desafío:
Consolidando sus habilidades analíticas y el uso de Copilot para diagnosticar problemas laborales reales:

1. **Diseño de su propio Prompt:** Formule un prompt en Copilot asumiendo que el chat es su **Asesor de Relaciones Laborales y Legal Interno**.
2. **Requisitos de su instrucción a la IA:** Pídale que le diseñe una estrategia de intervención de 3 pasos para solucionar el ausentismo de los viernes en el turno vespertino, que incluya:
   * Una medida preventiva administrativa (ejemplo: migrar el pago a transferencia electrónica en lugar de efectivo en planta, o ajustar horarios de salida).
   * Un programa de reconocimiento positivo que premie la puntualidad de ese turno específico.
   * La estructura de un breve mensaje (correo o aviso de 1 párrafo) formal pero motivador para comunicar la importancia del turno del viernes.
3. **La Entrega:** Copie el prompt que usted diseñó y la respuesta con el plan sugerido por Copilot al final de su archivo de Word para completar el portafolio de evidencias de este capítulo.

---

## 6. Conceptos Clave para Recordar

* **Índice de Bradford:** Métrica analítica que pondera el daño operativo de las ausencias imprevistas y recurrentes frente a una ausencia prolongada pero planificada legalmente.
* **Diagnóstico Basado en Hechos:** Principio fundamental donde cada acción correctiva o política de asistencia debe estar respaldada por registros numéricos consolidados y patrones de datos comprobables.
* **ROI de Bienestar Laboral:** Evaluación que demuestra cómo invertir en flexibilidad laboral y ergonomía reduce de manera directa los gastos de la empresa por el pago de horas extra y reemplazos de emergencia.

---

## 7. Resultado Esperado del Estudiante

El estudiante consolidará su aprendizaje documentando en su portafolio:

1. **Archivo de Word (`Reporte_Diagnostico_Ausentismo.docx`) que contenga:**
   * El informe ejecutivo del diagnóstico del área de Embarques estructurado formalmente (Paso 3).
   * La imagen conceptual sobre entorno saludable generada por Copilot e integrada en el texto (Paso 4).
   * El guión detallado para las diapositivas de PowerPoint (Paso 5).
   * La tabla formal de los 3 KPIs de control del plan (Paso 6).
   * El prompt de desarrollo autónomo diseñado por el alumno y la propuesta de solución para el ausentismo del viernes (Paso 7).
