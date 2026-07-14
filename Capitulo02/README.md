---LAB_START---
LAB_ID: 02-00-01
---MARKDOWN---
# Diagnóstico de ausentismo y detección de áreas de mejora.

---

## 1. Metadatos

| Atributo | Detalle |
|---|---|
| **Duración estimada** | 90 minutos |
| **Complejidad** | Media |
| **Nivel Bloom** | Aplicar (Apply) |
| **Módulo** | 2 — Análisis estadístico de indicadores críticos |
| **Lección base** | 2.1 Análisis estadístico de indicadores críticos (Rotación y Clima) |
| **Modalidad** | Individual con revisión grupal al cierre |

---

> ⚠️ **ADVERTENCIA LEGAL CRÍTICA:** Copilot Chat es una herramienta de apoyo y productividad, **NO un sustituto de asesoría jurídica laboral certificada**. Todos los documentos y análisis generados durante esta práctica deben ser revisados por un profesional legal antes de su implementación en situaciones reales.
>
> 🔒 **PRIVACIDAD DE DATOS:** **No ingreses datos reales de empleados** (nombres, CURP, RFC, salarios reales ni información confidencial de tu empresa) en Copilot Chat. Todos los ejercicios utilizan datos **ficticios** proporcionados en esta guía.

---

## 2. Descripción General

En esta práctica simularás el rol de un **consultor interno de RR.HH.** que debe presentar un diagnóstico ejecutivo de ausentismo y clima organizacional a la dirección general de una empresa manufacturera ficticia. Utilizarás Microsoft 365 Copilot Chat como amplificador analítico para interpretar un dataset de ausentismo de 6 meses, calcular indicadores clave (tasa de ausentismo, tasa de rotación, eNPS de clima), identificar áreas de riesgo y redactar un informe ejecutivo de una página con hallazgos y recomendaciones priorizadas.

La práctica aplica directamente los conceptos de la Lección 2.1: cálculo de tasas de rotación, interpretación de favorabilidad top-box, eNPS, segmentación por área y técnicas de prompting estructurado para análisis cuantitativo con Copilot Chat.

---

## 3. Objetivos de Aprendizaje

Al finalizar esta práctica serás capaz de:

- [ ] Calcular e interpretar la **tasa de ausentismo mensual** y la **tasa de rotación** por departamento a partir de un dataset ficticio, usando Copilot Chat como asistente analítico.
- [ ] Interpretar indicadores de **clima laboral** (favorabilidad top-box, eNPS) y relacionarlos con patrones de ausentismo mediante prompts estructurados de análisis de datos.
- [ ] Diseñar un **checklist de auditoría interna de cumplimiento laboral** con al menos 20 ítems organizados por categoría, utilizando Copilot Chat como asistente de estructuración.
- [ ] Redactar un **informe ejecutivo de diagnóstico** de una página con hallazgos priorizados y recomendaciones accionables, aplicando técnicas de refinamiento iterativo de prompts.

---

## 4. Prerrequisitos

### Conocimiento previo

| Área | Nivel requerido |
|---|---|
| Uso básico de Copilot Chat (prompts, contexto, iteración) | Básico — completar Práctica 1 o equivalente |
| Indicadores de RR.HH.: tasa de rotación, índice de ausentismo, clima laboral | Básico — comprensión conceptual |
| Microsoft Excel: abrir, filtrar y copiar datos de tablas | Básico |
| Concepto de auditoría laboral y su propósito organizacional | Conceptual |

### Acceso y licencias

| Requisito | Verificación |
|---|---|
| Licencia **Microsoft 365 Copilot** activa (empresarial) | Confirmar con TI con 48 h de anticipación |
| Acceso a **Microsoft 365 Copilot Chat** en [microsoft365.com/copilot](https://microsoft365.com/copilot) | Iniciar sesión y verificar que aparece el ícono de Copilot |
| **Microsoft Excel** (versión 2301 o superior) | Verificar en Archivo → Cuenta → Acerca de Excel |
| **Microsoft Word** (versión 2301 o superior) | Verificar en Archivo → Cuenta → Acerca de Word |
| **OneDrive** sincronizado con cuenta organizacional | Verificar que el ícono de OneDrive aparece en la barra de tareas |

---

## 5. Entorno de Laboratorio

### Hardware recomendado

| Componente | Mínimo | Recomendado |
|---|---|---|
| Procesador | Intel Core i5 8ª gen / AMD Ryzen 5 | Intel Core i7 / AMD Ryzen 7 |
| RAM | 8 GB | 16 GB |
| Almacenamiento libre | 2 GB | 5 GB |
| Pantalla | 1366×768 | 1920×1080 |
| Conexión a Internet | 10 Mbps | 25 Mbps o superior |

### Software requerido

| Aplicación | Versión mínima | Propósito en esta práctica |
|---|---|---|
| Microsoft 365 Copilot Chat | Licencia empresarial activa | Análisis, cálculos, redacción asistida |
| Microsoft Excel (M365) | 2301 o superior | Preparación y visualización del dataset |
| Microsoft Word (M365) | 2301 o superior | Redacción del informe ejecutivo |
| Microsoft Edge o Chrome | 120 o superior | Acceso a Copilot Chat web |
| Adobe Acrobat Reader / Edge PDF | 2020 o superior | Consulta de esta guía en PDF |

### Configuración inicial (antes de comenzar)

Ejecuta los siguientes pasos de preparación **antes** de iniciar el Paso 1:

1. Abre tu navegador (Edge o Chrome) e inicia sesión en [https://microsoft365.com](https://microsoft365.com) con tu cuenta organizacional.
2. Navega a **Copilot Chat** desde el menú de aplicaciones de Microsoft 365 (ícono de cuadrícula → Copilot).
3. Verifica que aparece la interfaz de chat con el indicador **"Trabajo"** seleccionado (no "Web"). Esto garantiza que Copilot accede al contexto organizacional con protección de datos empresarial.
4. Abre **Microsoft Excel** y **Microsoft Word** en paralelo (puedes minimizarlos).
5. Crea una carpeta en OneDrive llamada `Lab02-Diagnostico-Ausentismo` donde guardarás todos los archivos generados.

> 💡 **Nota sobre variabilidad:** Las respuestas de Copilot Chat son probabilísticas y pueden variar entre sesiones. Si tu resultado difiere de los ejemplos de esta guía, eso es **normal y esperado**. Lo importante es evaluar la calidad y pertinencia de la respuesta, no replicarla exactamente.

---

## 6. Pasos del Laboratorio

---

### Paso 1: Preparar el Dataset Ficticio de Ausentismo en Excel

**Objetivo:** Construir el dataset de trabajo que usarás durante toda la práctica, familiarizarte con su estructura y realizar una exploración inicial antes de invocar a Copilot.

**Instrucciones:**

1. Abre **Microsoft Excel** y crea un nuevo libro en blanco. Guárdalo en tu carpeta de OneDrive como `Dataset_Ausentismo_Ficticio.xlsx`.

2. En la **Hoja 1** (renómbrala `Ausentismo`), crea la siguiente tabla con los encabezados en la fila 1 y los datos ficticios proporcionados a continuación. Escribe o copia exactamente esta información:

| ID_Colaborador | Departamento | Categoria | Antiguedad_Anos | Mes | Dias_Ausentes | Motivo_Ausencia | Tipo_Ausencia |
|---|---|---|---|---|---|---|---|
| EMP-001 | Producción | Operativo | 2 | Enero | 3 | Enfermedad general | Justificada |
| EMP-002 | Producción | Operativo | 1 | Enero | 5 | Enfermedad general | Justificada |
| EMP-003 | Producción | Técnico | 4 | Enero | 1 | Asunto personal | Injustificada |
| EMP-004 | Ventas | Ejecutivo | 3 | Enero | 2 | Cita médica | Justificada |
| EMP-005 | Ventas | Operativo | 1 | Enero | 4 | Enfermedad general | Justificada |
| EMP-006 | Administración | Administrativo | 6 | Enero | 0 | — | — |
| EMP-007 | Administración | Directivo | 10 | Enero | 1 | Asunto personal | Injustificada |
| EMP-008 | Producción | Operativo | 1 | Febrero | 6 | Accidente laboral | Justificada |
| EMP-009 | Producción | Técnico | 3 | Febrero | 2 | Enfermedad general | Justificada |
| EMP-010 | Producción | Operativo | 2 | Febrero | 4 | Enfermedad general | Justificada |
| EMP-011 | Ventas | Ejecutivo | 5 | Febrero | 1 | Cita médica | Justificada |
| EMP-012 | Ventas | Operativo | 1 | Febrero | 5 | Enfermedad general | Justificada |
| EMP-013 | Administración | Administrativo | 8 | Febrero | 0 | — | — |
| EMP-014 | Administración | Administrativo | 2 | Febrero | 2 | Asunto personal | Injustificada |
| EMP-015 | Producción | Operativo | 1 | Marzo | 7 | Enfermedad crónica | Justificada |
| EMP-016 | Producción | Operativo | 2 | Marzo | 5 | Accidente laboral | Justificada |
| EMP-017 | Producción | Técnico | 1 | Marzo | 3 | Enfermedad general | Justificada |
| EMP-018 | Ventas | Ejecutivo | 4 | Marzo | 0 | — | — |
| EMP-019 | Ventas | Operativo | 1 | Marzo | 6 | Enfermedad general | Justificada |
| EMP-020 | Administración | Directivo | 12 | Marzo | 0 | — | — |
| EMP-021 | Administración | Administrativo | 3 | Marzo | 1 | Cita médica | Justificada |
| EMP-022 | Producción | Operativo | 1 | Abril | 8 | Enfermedad crónica | Justificada |
| EMP-023 | Producción | Técnico | 5 | Abril | 2 | Asunto personal | Injustificada |
| EMP-024 | Producción | Operativo | 2 | Abril | 4 | Enfermedad general | Justificada |
| EMP-025 | Ventas | Ejecutivo | 3 | Abril | 3 | Estrés laboral | Justificada |
| EMP-026 | Ventas | Operativo | 1 | Abril | 5 | Estrés laboral | Justificada |
| EMP-027 | Administración | Administrativo | 7 | Abril | 1 | Cita médica | Justificada |
| EMP-028 | Administración | Administrativo | 1 | Abril | 3 | Enfermedad general | Justificada |
| EMP-029 | Producción | Operativo | 1 | Mayo | 9 | Accidente laboral | Justificada |
| EMP-030 | Producción | Técnico | 2 | Mayo | 4 | Enfermedad general | Justificada |
| EMP-031 | Producción | Operativo | 3 | Mayo | 3 | Enfermedad general | Justificada |
| EMP-032 | Ventas | Ejecutivo | 6 | Mayo | 2 | Asunto personal | Injustificada |
| EMP-033 | Ventas | Operativo | 1 | Mayo | 6 | Estrés laboral | Justificada |
| EMP-034 | Administración | Directivo | 9 | Mayo | 0 | — | — |
| EMP-035 | Administración | Administrativo | 2 | Mayo | 2 | Enfermedad general | Justificada |
| EMP-036 | Producción | Operativo | 1 | Junio | 10 | Enfermedad crónica | Justificada |
| EMP-037 | Producción | Técnico | 4 | Junio | 3 | Accidente laboral | Justificada |
| EMP-038 | Producción | Operativo | 2 | Junio | 5 | Enfermedad general | Justificada |
| EMP-039 | Ventas | Ejecutivo | 2 | Junio | 4 | Estrés laboral | Justificada |
| EMP-040 | Ventas | Operativo | 1 | Junio | 7 | Estrés laboral | Justificada |
| EMP-041 | Administración | Administrativo | 5 | Junio | 1 | Cita médica | Justificada |
| EMP-042 | Administración | Administrativo | 1 | Junio | 3 | Enfermedad general | Justificada |

3. Agrega también una **Hoja 2** (renómbrala `Headcount`) con la siguiente tabla de plantilla por departamento y mes (necesaria para calcular tasas):

| Departamento | Mes | HC_Inicio | HC_Cierre | Separaciones | Dias_Laborables |
|---|---|---|---|---|---|
| Producción | Enero | 45 | 44 | 1 | 22 |
| Ventas | Enero | 20 | 20 | 0 | 22 |
| Administración | Enero | 15 | 15 | 0 | 22 |
| Producción | Febrero | 44 | 43 | 1 | 20 |
| Ventas | Febrero | 20 | 19 | 1 | 20 |
| Administración | Febrero | 15 | 15 | 0 | 20 |
| Producción | Marzo | 43 | 42 | 1 | 21 |
| Ventas | Marzo | 19 | 19 | 0 | 21 |
| Administración | Marzo | 15 | 15 | 0 | 21 |
| Producción | Abril | 42 | 41 | 1 | 22 |
| Ventas | Abril | 19 | 18 | 1 | 22 |
| Administración | Abril | 15 | 14 | 1 | 22 |
| Producción | Mayo | 41 | 40 | 1 | 21 |
| Ventas | Mayo | 18 | 18 | 0 | 21 |
| Administración | Mayo | 14 | 14 | 0 | 21 |
| Producción | Junio | 40 | 39 | 1 | 22 |
| Ventas | Junio | 18 | 17 | 1 | 22 |
| Administración | Junio | 14 | 14 | 0 | 22 |

4. Agrega una **Hoja 3** (renómbrala `Clima`) con el siguiente resumen de encuesta de clima laboral (datos ficticios de una encuesta aplicada en mayo):

| Departamento | Dimension | Favorabilidad_Pct | Promedio_Likert | eNPS |
|---|---|---|---|---|
| Producción | Liderazgo | 48 | 3.1 | -15 |
| Producción | Carga de trabajo | 32 | 2.6 | -15 |
| Producción | Compensación | 41 | 2.9 | -15 |
| Producción | Seguridad psicológica | 38 | 2.8 | -15 |
| Ventas | Liderazgo | 55 | 3.4 | -8 |
| Ventas | Carga de trabajo | 44 | 3.1 | -8 |
| Ventas | Compensación | 50 | 3.3 | -8 |
| Ventas | Seguridad psicológica | 47 | 3.2 | -8 |
| Administración | Liderazgo | 72 | 4.1 | +22 |
| Administración | Carga de trabajo | 68 | 3.9 | +22 |
| Administración | Compensación | 65 | 3.8 | +22 |
| Administración | Seguridad psicológica | 70 | 4.0 | +22 |

5. Guarda el archivo (`Ctrl + S`).

**Resultado esperado:** Un archivo Excel con tres hojas (`Ausentismo`, `Headcount`, `Clima`) correctamente estructuradas y guardadas en OneDrive.

**Verificación:** Confirma que las tres pestañas aparecen en la parte inferior de Excel y que el archivo está guardado en la carpeta `Lab02-Diagnostico-Ausentismo` de OneDrive (el ícono de nube debe aparecer junto al nombre del archivo en la barra de título).

---

### Paso 2: Análisis Estadístico Descriptivo con Copilot Chat

**Objetivo:** Usar Copilot Chat para obtener un análisis estadístico descriptivo del dataset de ausentismo, identificando patrones por departamento y motivo de ausencia.

**Instrucciones:**

1. Abre **Copilot Chat** en tu navegador. Verifica que el modo **"Trabajo"** está seleccionado.

2. Inicia una nueva conversación. En el campo de texto, escribe el siguiente prompt. Cópialo y pégalo exactamente como aparece (puedes ajustar el formato de la tabla si es necesario, pero mantén todos los datos):

```text
Actúa como analista de Recursos Humanos especializado en diagnóstico de ausentismo laboral.

Tengo el siguiente dataset ficticio de ausentismo de una empresa manufacturera (6 meses, 3 departamentos). Analiza estadísticamente los datos y proporciona:

1. Total de días perdidos por departamento y por mes (tabla resumen).
2. Promedio de días ausentes por colaborador en cada departamento.
3. Distribución de motivos de ausencia (frecuencia y porcentaje del total).
4. Identificación del departamento con mayor ausentismo absoluto y relativo.
5. Tendencia mensual: ¿el ausentismo aumenta, disminuye o es estable?
6. Perfil de mayor riesgo: ¿qué categoría de colaborador y antigüedad concentra más ausencias?

DATOS (formato CSV):
ID_Colaborador,Departamento,Categoria,Antiguedad_Anos,Mes,Dias_Ausentes,Motivo_Ausencia,Tipo_Ausencia
EMP-001,Producción,Operativo,2,Enero,3,Enfermedad general,Justificada
EMP-002,Producción,Operativo,1,Enero,5,Enfermedad general,Justificada
EMP-003,Producción,Técnico,4,Enero,1,Asunto personal,Injustificada
EMP-004,Ventas,Ejecutivo,3,Enero,2,Cita médica,Justificada
EMP-005,Ventas,Operativo,1,Enero,4,Enfermedad general,Justificada
EMP-006,Administración,Administrativo,6,Enero,0,—,—
EMP-007,Administración,Directivo,10,Enero,1,Asunto personal,Injustificada
EMP-008,Producción,Operativo,1,Febrero,6,Accidente laboral,Justificada
EMP-009,Producción,Técnico,3,Febrero,2,Enfermedad general,Justificada
EMP-010,Producción,Operativo,2,Febrero,4,Enfermedad general,Justificada
EMP-011,Ventas,Ejecutivo,5,Febrero,1,Cita médica,Justificada
EMP-012,Ventas,Operativo,1,Febrero,5,Enfermedad general,Justificada
EMP-013,Administración,Administrativo,8,Febrero,0,—,—
EMP-014,Administración,Administrativo,2,Febrero,2,Asunto personal,Injustificada
EMP-015,Producción,Operativo,1,Marzo,7,Enfermedad crónica,Justificada
EMP-016,Producción,Operativo,2,Marzo,5,Accidente laboral,Justificada
EMP-017,Producción,Técnico,1,Marzo,3,Enfermedad general,Justificada
EMP-018,Ventas,Ejecutivo,4,Marzo,0,—,—
EMP-019,Ventas,Operativo,1,Marzo,6,Enfermedad general,Justificada
EMP-020,Administración,Directivo,12,Marzo,0,—,—
EMP-021,Administración,Administrativo,3,Marzo,1,Cita médica,Justificada
EMP-022,Producción,Operativo,1,Abril,8,Enfermedad crónica,Justificada
EMP-023,Producción,Técnico,5,Abril,2,Asunto personal,Injustificada
EMP-024,Producción,Operativo,2,Abril,4,Enfermedad general,Justificada
EMP-025,Ventas,Ejecutivo,3,Abril,3,Estrés laboral,Justificada
EMP-026,Ventas,Operativo,1,Abril,5,Estrés laboral,Justificada
EMP-027,Administración,Administrativo,7,Abril,1,Cita médica,Justificada
EMP-028,Administración,Administrativo,1,Abril,3,Enfermedad general,Justificada
EMP-029,Producción,Operativo,1,Mayo,9,Accidente laboral,Justificada
EMP-030,Producción,Técnico,2,Mayo,4,Enfermedad general,Justificada
EMP-031,Producción,Operativo,3,Mayo,3,Enfermedad general,Justificada
EMP-032,Ventas,Ejecutivo,6,Mayo,2,Asunto personal,Injustificada
EMP-033,Ventas,Operativo,1,Mayo,6,Estrés laboral,Justificada
EMP-034,Administración,Directivo,9,Mayo,0,—,—
EMP-035,Administración,Administrativo,2,Mayo,2,Enfermedad general,Justificada
EMP-036,Producción,Operativo,1,Junio,10,Enfermedad crónica,Justificada
EMP-037,Producción,Técnico,4,Junio,3,Accidente laboral,Justificada
EMP-038,Producción,Operativo,2,Junio,5,Enfermedad general,Justificada
EMP-039,Ventas,Ejecutivo,2,Junio,4,Estrés laboral,Justificada
EMP-040,Ventas,Operativo,1,Junio,7,Estrés laboral,Justificada
EMP-041,Administración,Administrativo,5,Junio,1,Cita médica,Justificada
EMP-042,Administración,Administrativo,1,Junio,3,Enfermedad general,Justificada

Presenta los resultados en tablas claras y concluye con un párrafo de interpretación ejecutiva de máximo 5 oraciones.
```

3. Presiona **Enter** y espera la respuesta de Copilot (aproximadamente 20–40 segundos).

4. Lee cuidadosamente la respuesta. Evalúa si:
   - Las tablas de resumen son coherentes con los datos ingresados.
   - La tendencia mensual identificada es consistente (los datos muestran una tendencia creciente en Producción).
   - El perfil de riesgo señala a colaboradores Operativos con antigüedad de 1–2 años.

5. Si algún punto no está suficientemente desarrollado, usa el siguiente **prompt de refinamiento iterativo**:

```text
Gracias. Por favor amplía el punto 6 (perfil de mayor riesgo): desglosa los días perdidos 
por combinación de Categoría + Antigüedad (0-1 año, 2-3 años, 4+ años) y señala cuál 
combinación representa el mayor riesgo de ausentismo acumulado. Presenta el resultado 
en una tabla con columnas: Categoría, Rango_Antigüedad, Total_Días_Perdidos, % del Total.
```

6. Copia la respuesta completa (tablas + interpretación) y pégala en un documento de Word que guardarás como `Hallazgos_Paso2_Analisis_Descriptivo.docx` en tu carpeta de OneDrive.

**Resultado esperado:** Copilot debe entregar tablas de resumen mostrando que **Producción** concentra el mayor ausentismo (estimado >60% del total de días perdidos), con una **tendencia creciente** mes a mes, y que los colaboradores **Operativos con 1–2 años de antigüedad** representan el perfil de mayor riesgo.

**Verificación:** Comprueba manualmente al menos una suma de la tabla (por ejemplo, total de días perdidos en Producción en Enero: EMP-001 + EMP-002 + EMP-003 = 3 + 5 + 1 = 9 días). Si la tabla de Copilot muestra 9 días para Producción en Enero, el cálculo es correcto.

---

### Paso 3: Cálculo de Indicadores Clave — Tasa de Ausentismo y Tasa de Rotación

**Objetivo:** Calcular la tasa de ausentismo mensual por departamento y la tasa de rotación semestral, comparándolas contra benchmarks del sector, con el apoyo analítico de Copilot Chat.

**Instrucciones:**

1. En Copilot Chat (continúa la misma conversación o inicia una nueva), envía el siguiente prompt estructurado que incluye los datos de headcount:

```text
Actúa como consultor de métricas de Recursos Humanos. Necesito calcular dos indicadores clave 
usando los siguientes datos ficticios de una empresa manufacturera.

FÓRMULAS A USAR:
- Índice de ausentismo mensual (%) = (Días perdidos en el mes / (HC_Promedio × Días_Laborables)) × 100
  donde HC_Promedio = (HC_Inicio + HC_Cierre) / 2
- Tasa de rotación mensual (%) = (Separaciones / HC_Promedio) × 100
- Tasa de rotación semestral acumulada (%) = suma de separaciones en 6 meses / HC_Promedio_semestral × 100

DATOS DE HEADCOUNT Y AUSENTISMO POR DEPARTAMENTO Y MES:

Departamento,Mes,HC_Inicio,HC_Cierre,Separaciones,Dias_Laborables,Dias_Perdidos
Producción,Enero,45,44,1,22,9
Ventas,Enero,20,20,0,22,6
Administración,Enero,15,15,0,22,1
Producción,Febrero,44,43,1,20,12
Ventas,Febrero,20,19,1,20,6
Administración,Febrero,15,15,0,20,2
Producción,Marzo,43,42,1,21,15
Ventas,Marzo,19,19,0,21,6
Administración,Marzo,15,15,0,21,1
Producción,Abril,42,41,1,22,14
Ventas,Abril,19,18,1,22,8
Administración,Abril,15,14,1,22,4
Producción,Mayo,41,40,1,21,16
Ventas,Mayo,18,18,0,21,8
Administración,Mayo,14,14,0,21,2
Producción,Junio,40,39,1,22,18
Ventas,Junio,18,17,1,22,11
Administración,Junio,14,14,0,22,4

TAREAS:
1. Calcula el índice de ausentismo mensual (%) para cada departamento en cada mes. 
   Muestra la fórmula aplicada con los valores reales en al menos un ejemplo por departamento.
2. Calcula la tasa de rotación mensual (%) para cada departamento en cada mes.
3. Calcula la tasa de rotación semestral acumulada por departamento.
4. Presenta todos los resultados en tablas organizadas por departamento.
5. Compara los resultados contra estos benchmarks del sector manufacturero mexicano:
   - Índice de ausentismo aceptable: 1.5% – 3.5% mensual
   - Tasa de rotación mensual aceptable: 1% – 2.5%
   - Tasa de rotación semestral aceptable: 6% – 15%
6. Señala qué departamentos y meses superan los benchmarks y clasifícalos como 
   ALERTA (supera benchmark) o NORMAL (dentro del rango).
7. Concluye con una interpretación ejecutiva de 3 oraciones sobre el estado general 
   de los indicadores.
```

2. Espera la respuesta y revisa que los cálculos sean coherentes. Para verificar manualmente, aplica la fórmula de ausentismo para **Producción en Enero**:
   - HC_Promedio = (45 + 44) / 2 = 44.5
   - Índice de ausentismo = (9 / (44.5 × 22)) × 100 = (9 / 979) × 100 ≈ **0.92%**

   > ⚠️ Nota: Este valor está por debajo del benchmark. Sin embargo, la tendencia creciente hacia junio es la señal de alerta. Evalúa si Copilot identifica correctamente esta tendencia.

3. Si Copilot no menciona la tendencia creciente, usa este prompt de seguimiento:

```text
Ahora analiza la tendencia del índice de ausentismo de Producción a lo largo de los 
6 meses. ¿Hay una tendencia estadísticamente significativa (creciente, decreciente o 
estable)? ¿En qué mes se cruza o se acerca más al umbral de alerta de 3.5%? 
¿Qué implicaciones tiene esta tendencia para la planificación de RR.HH. del segundo semestre?
```

4. Documenta los resultados: copia las tablas de indicadores en una nueva hoja de tu archivo Excel llamada `Indicadores_Calculados`. Puedes transcribir los valores manualmente o crear fórmulas Excel que repliquen los cálculos.

5. Guarda un registro de los prompts y respuestas en tu documento Word `Hallazgos_Paso2_Analisis_Descriptivo.docx` (agrega una nueva sección llamada "Paso 3 — Indicadores Clave").

**Resultado esperado:** Tablas con índices de ausentismo y tasas de rotación por departamento/mes, con clasificación ALERTA/NORMAL. Producción debe mostrar índices crecientes que se acercan al umbral de alerta en mayo-junio. La tasa de rotación semestral de Producción debe aproximarse al límite superior del benchmark.

**Verificación:** Confirma que la tabla de Copilot incluye al menos una columna de clasificación ALERTA/NORMAL y que la interpretación ejecutiva menciona explícitamente a Producción como el departamento de mayor riesgo.

---

### Paso 4: Análisis de Clima Laboral e Identificación de Factores de Insatisfacción

**Objetivo:** Interpretar los datos de encuesta de clima laboral usando Copilot Chat, correlacionarlos con los patrones de ausentismo identificados y detectar factores de insatisfacción prioritarios.

**Instrucciones:**

1. En Copilot Chat, inicia un nuevo prompt con el contexto de clima laboral:

```text
Actúa como especialista en clima organizacional y analítica de personas.

Tengo los siguientes datos ficticios de una encuesta de clima laboral aplicada en mayo 
a 3 departamentos de una empresa manufacturera. La escala Likert es de 1 a 5, donde 
4-5 es "favorable" (top box). El eNPS se calcula como % Promotores (9-10) menos 
% Detractores (0-6).

DATOS DE CLIMA:
Departamento,Dimension,Favorabilidad_Pct,Promedio_Likert,eNPS
Producción,Liderazgo,48,3.1,-15
Producción,Carga de trabajo,32,2.6,-15
Producción,Compensación,41,2.9,-15
Producción,Seguridad psicológica,38,2.8,-15
Ventas,Liderazgo,55,3.4,-8
Ventas,Carga de trabajo,44,3.1,-8
Ventas,Compensación,50,3.3,-8
Ventas,Seguridad psicológica,47,3.2,-8
Administración,Liderazgo,72,4.1,+22
Administración,Carga de trabajo,68,3.9,+22
Administración,Compensación,65,3.8,+22
Administración,Seguridad psicológica,70,4.0,+22

CONTEXTO ADICIONAL (resumen cualitativo de comentarios abiertos de la encuesta):
- Producción: "Mucha presión en línea, supervisores no escuchan, varios compañeros 
  han tenido accidentes por fatiga, siento que no me valoran."
- Ventas: "Las metas son altas pero alcanzables, aunque el estrés aumentó este año. 
  El sueldo podría mejorar."
- Administración: "Buen ambiente, jefes accesibles, carga de trabajo manejable."

TAREAS:
1. Identifica las 3 dimensiones más críticas (menor favorabilidad) y el departamento 
   con mayor riesgo de insatisfacción.
2. Interpreta el eNPS de cada departamento: ¿qué significa un eNPS de -15 en Producción?
3. Relaciona los datos de clima con los patrones de ausentismo del análisis previo: 
   ¿hay correlación entre baja favorabilidad y alto ausentismo? Explica el razonamiento.
4. Identifica 3 factores de insatisfacción prioritarios que podrían estar impulsando 
   el ausentismo en Producción.
5. ¿Qué dimensión de clima tiene mayor urgencia de intervención y por qué?
6. Presenta un semáforo de riesgo por departamento y dimensión 
   (🔴 Crítico <40%, 🟡 Atención 40-60%, 🟢 Favorable >60%).
```

2. Revisa la respuesta. Evalúa si Copilot:
   - Identifica correctamente "Carga de trabajo" en Producción (32%) como la dimensión más crítica.
   - Explica el eNPS negativo como indicador de que hay más detractores que promotores.
   - Establece una relación lógica (no causal) entre bajo clima y alto ausentismo.

3. Aplica un prompt de refinamiento para profundizar en la correlación:

```text
Profundiza en la relación entre clima y ausentismo. Considerando que:
- Producción tiene favorabilidad de Carga de trabajo = 32% y ausentismo creciente
- Administración tiene favorabilidad de Carga de trabajo = 68% y ausentismo estable y bajo

¿Qué hipótesis de causalidad podrías plantear? ¿Qué datos adicionales necesitarías 
para confirmar o refutar esa hipótesis? Menciona al menos 2 variables de control 
que deberías incluir en un análisis más riguroso.

IMPORTANTE: Recuerda que correlación no implica causalidad. Incluye esta advertencia 
metodológica en tu respuesta.
```

4. Guarda las respuestas en tu documento Word (nueva sección: "Paso 4 — Análisis de Clima").

**Resultado esperado:** Un semáforo de riesgo por departamento/dimensión, identificación de "Carga de trabajo" y "Seguridad psicológica" en Producción como dimensiones críticas (🔴), y una explicación de la correlación clima-ausentismo con advertencia metodológica explícita.

**Verificación:** Confirma que el semáforo clasifica correctamente: Producción/Carga de trabajo = 🔴 (32% < 40%), Ventas/Compensación = 🟡 (50% entre 40-60%), Administración/Liderazgo = 🟢 (72% > 60%).

---

### Paso 5: Diseño del Checklist de Auditoría Interna de Cumplimiento Laboral

**Objetivo:** Usar Copilot Chat como asistente de estructuración para diseñar un checklist de auditoría laboral con al menos 20 ítems organizados por categoría, aplicable al contexto de la empresa ficticia.

**Instrucciones:**

1. En Copilot Chat, envía el siguiente prompt:

```text
Actúa como especialista en auditoría de cumplimiento laboral en México.

Necesito diseñar un checklist de auditoría interna de cumplimiento laboral para una 
empresa manufacturera con 80 colaboradores (3 departamentos: Producción, Ventas y 
Administración). El checklist debe ser práctico, accionable y aplicable a una auditoría 
semestral interna de RR.HH.

INSTRUCCIONES:
1. Organiza el checklist en exactamente 5 categorías:
   A. Contratos y documentación laboral
   B. Jornada de trabajo y horas extras
   C. Prestaciones de ley y beneficios
   D. Seguridad e higiene en el trabajo
   E. Gestión del ausentismo y disciplina

2. Incluye MÍNIMO 4 ítems por categoría (total mínimo: 20 ítems).

3. Para cada ítem incluye:
   - Número de ítem (A1, A2, B1, B2, etc.)
   - Descripción del punto a verificar
   - Documento o evidencia requerida
   - Columna de estatus: Cumple / No cumple / En proceso / N/A

4. Los ítems deben ser específicos, verificables y basados en obligaciones del 
   empleador en México (sin citar artículos específicos de la LFT para evitar 
   errores de versión — solo menciona el tema general).

5. Al final, agrega una sección de "Notas del auditor" y una línea de firma 
   con: Auditor responsable, Fecha de auditoría, Próxima revisión programada.

Presenta el checklist en formato de tabla Markdown.

ADVERTENCIA: Este checklist es una herramienta de apoyo. Debe ser revisado y 
validado por un abogado laboralista certificado antes de su uso oficial.
```

2. Revisa el checklist generado. Verifica que:
   - Tiene al menos 20 ítems distribuidos en 5 categorías.
   - Cada ítem tiene descripción, evidencia requerida y columna de estatus.
   - La categoría D (Seguridad e higiene) incluye ítems relevantes para Producción (donde hay accidentes laborales en el dataset).

3. Si el checklist no incluye ítems específicos sobre **gestión de ausentismo recurrente** o **protocolo de accidentes laborales**, solicita una adición:

```text
Por favor agrega 3 ítems adicionales a la categoría E (Gestión del ausentismo y 
disciplina) específicamente sobre:
1. Protocolo de seguimiento a colaboradores con ausentismo recurrente (3+ ausencias 
   en 2 meses consecutivos)
2. Registro y reporte de accidentes de trabajo al IMSS
3. Programa de reintegración laboral para colaboradores con incapacidades prolongadas

Mantén el mismo formato de tabla con columnas: Número, Descripción, Evidencia requerida, Estatus.
```

4. Copia el checklist completo y pégalo en un nuevo documento Word llamado `Checklist_Auditoria_Laboral.docx`. Guárdalo en tu carpeta de OneDrive.

5. Agrega manualmente al inicio del documento el siguiente encabezado:

```
CHECKLIST DE AUDITORÍA INTERNA DE CUMPLIMIENTO LABORAL
Empresa: [Nombre Ficticio de la Empresa]
Periodo auditado: Enero – Junio [Año]
Elaborado con apoyo de: Microsoft 365 Copilot Chat
ADVERTENCIA: Documento de práctica. Requiere validación por abogado laboralista 
certificado antes de uso oficial.
```

**Resultado esperado:** Un checklist en formato tabla con mínimo 23 ítems (20 base + 3 adicionales), organizado en 5 categorías, con columnas de evidencia y estatus, guardado en Word.

**Verificación:** Cuenta los ítems en el documento. Verifica que la categoría D incluye al menos un ítem sobre registro de accidentes de trabajo y que la categoría E incluye el protocolo de seguimiento a ausentismo recurrente.

---

### Paso 6: Redacción del Informe Ejecutivo de Diagnóstico

**Objetivo:** Integrar todos los hallazgos de los pasos anteriores en un informe ejecutivo de una página, utilizando técnicas de prompting estructurado y refinamiento iterativo para producir un documento listo para presentación a la dirección.

**Instrucciones:**

1. Antes de escribir el prompt, recopila en un texto breve los hallazgos clave de los pasos anteriores. Usa este resumen como contexto para Copilot:

```text
HALLAZGOS CLAVE DEL DIAGNÓSTICO (para incluir en el informe):
- Producción: ausentismo creciente (tendencia al alza en 6 meses), 6 separaciones en el semestre
- Perfil de riesgo: Operativos con 1-2 años de antigüedad concentran >60% de días perdidos
- Motivos principales: enfermedad general (40%), accidentes laborales (22%), estrés laboral (20%)
- Clima Producción: eNPS = -15, favorabilidad Carga de trabajo = 32% (crítico), 
  Seguridad psicológica = 38% (crítico)
- Clima Administración: eNPS = +22, favorabilidad promedio >65% (favorable)
- Tasa de rotación semestral Producción: ~14% (límite superior del benchmark)
- Checklist de auditoría: 5 áreas identificadas con puntos de cumplimiento a verificar
```

2. En Copilot Chat, envía el siguiente prompt de redacción estructurada:

```text
Actúa como consultor senior de Recursos Humanos redactando un informe ejecutivo 
para presentar a la Dirección General.

Usando los hallazgos del diagnóstico que te proporciono, redacta un INFORME EJECUTIVO 
DE UNA PÁGINA con la siguiente estructura exacta:

ESTRUCTURA DEL INFORME:
1. ENCABEZADO: Título, empresa ficticia, periodo, fecha, elaborado por (RR.HH.)
2. RESUMEN EJECUTIVO (3-4 oraciones): situación general del ausentismo y clima
3. HALLAZGOS PRINCIPALES (máximo 6 puntos en viñetas, con dato cuantitativo):
   - Uno por indicador clave (ausentismo, rotación, clima, perfil de riesgo)
4. ÁREAS CRÍTICAS (tabla de 2 columnas: Área de Riesgo | Indicador de Alerta)
5. RECOMENDACIONES PRIORIZADAS (tabla con 3 columnas: Prioridad | Acción | Impacto Esperado)
   - Incluye mínimo 5 recomendaciones ordenadas de mayor a menor urgencia
   - Las recomendaciones deben ser específicas y accionables (no genéricas)
6. PRÓXIMOS PASOS (3 acciones concretas con responsable sugerido y plazo)
7. PIE DE PÁGINA: Advertencia de confidencialidad y nota de que el documento 
   fue elaborado con apoyo de IA y requiere validación profesional

HALLAZGOS A INTEGRAR:
[Pega aquí el resumen de hallazgos del punto 1 de estas instrucciones]

TONO: Ejecutivo, directo, basado en datos. Máximo 600 palabras en total.
FORMATO: Markdown con encabezados claros.
```

3. Revisa el informe generado. Evalúa:
   - ¿Las recomendaciones son específicas y accionables (no solo "mejorar el clima")?
   - ¿Cada hallazgo tiene al menos un dato cuantitativo?
   - ¿El tono es apropiado para una presentación a la Dirección General?

4. Si las recomendaciones son genéricas, aplica este refinamiento:

```text
Las recomendaciones son demasiado generales. Por favor, reescribe las 5 recomendaciones 
siendo más específicas. Para cada una incluye:
- Qué acción concreta se debe tomar (verbo de acción + objeto específico)
- A quién va dirigida (RR.HH., Supervisores de Producción, Dirección, IMSS)
- En qué plazo: Inmediato (0-30 días), Corto plazo (1-3 meses), Mediano plazo (3-6 meses)
- Qué indicador se espera mejorar y en qué porcentaje estimado

Ejemplo del nivel de especificidad esperado:
"Implementar programa de pausas activas y rotación de tareas en línea de Producción, 
coordinado por el supervisor de turno, en los próximos 30 días, con el objetivo de 
reducir el ausentismo por estrés laboral en un 20% para el tercer trimestre."
```

5. Una vez satisfecho con el informe, cópialo en Microsoft Word. Crea un nuevo documento llamado `Informe_Ejecutivo_Diagnostico_Ausentismo.docx`. Aplica el siguiente formato:
   - Fuente: Calibri 11 para cuerpo, Calibri 14 negrita para títulos de sección.
   - Márgenes: 2.5 cm todos los lados.
   - Encabezado de página: nombre de la empresa ficticia + "CONFIDENCIAL — DATOS FICTICIOS".
   - Pie de página: número de página + advertencia legal.

6. Guarda el documento en tu carpeta de OneDrive.

**Resultado esperado:** Un informe ejecutivo de una página (máximo 600 palabras) con estructura completa, 5 recomendaciones específicas y accionables con responsable y plazo, y formato profesional en Word.

**Verificación:** Comprueba que el informe tiene exactamente las 7 secciones especificadas. Verifica que al menos 3 de las 5 recomendaciones mencionan un responsable específico (no solo "RR.HH." de forma genérica) y un plazo definido.

---

### Paso 7: Evaluación Crítica de las Respuestas de Copilot

**Objetivo:** Aplicar criterio profesional para evaluar la pertinencia, precisión y limitaciones de las respuestas generadas por Copilot Chat a lo largo de la práctica.

**Instrucciones:**

1. En Copilot Chat, envía el siguiente prompt de metacognición analítica:

```text
A lo largo de esta sesión de análisis de ausentismo y clima laboral, has generado 
varios análisis y documentos. Ahora necesito que actúes como revisor crítico de 
tu propio trabajo.

Responde las siguientes preguntas de evaluación:

1. ¿Qué supuestos asumiste al analizar los datos que podrían no ser válidos en 
   un contexto real? (menciona al menos 3)

2. ¿Qué información adicional necesitarías para hacer un diagnóstico más robusto 
   y estadísticamente confiable?

3. ¿En qué puntos del análisis existe mayor riesgo de error o interpretación incorrecta?

4. ¿Qué limitaciones tiene el uso de IA generativa para este tipo de análisis de 
   Recursos Humanos? (menciona al menos 4 limitaciones específicas)

5. ¿Qué validaciones debería hacer un especialista humano de RR.HH. antes de 
   presentar estos hallazgos a la Dirección General?

Sé honesto y específico. El objetivo es desarrollar pensamiento crítico sobre el 
uso responsable de IA en la gestión de personas.
```

2. Lee la respuesta y anota en tu documento Word (nueva sección: "Evaluación Crítica") al menos **3 limitaciones** que consideres más relevantes para tu contexto organizacional.

3. Reflexiona individualmente (2 minutos) y responde por escrito en tu documento:
   - *"¿En qué parte del análisis confié demasiado en Copilot sin verificar los datos?"*
   - *"¿Qué haría diferente en un diagnóstico real con datos de mi organización?"*

**Resultado esperado:** Una sección de evaluación crítica en tu documento Word con las limitaciones identificadas por Copilot y tu reflexión personal sobre el uso responsable de IA en RR.HH.

**Verificación:** Confirma que tu reflexión personal tiene al menos 3 oraciones propias (no copiadas de Copilot) que conectan las limitaciones identificadas con tu contexto profesional real.

---

## 7. Validación y Pruebas

Al finalizar todos los pasos, realiza la siguiente verificación integral:

### Lista de verificación de entregables

| # | Entregable | Ubicación en OneDrive | Criterio de éxito |
|---|---|---|---|
| 1 | `Dataset_Ausentismo_Ficticio.xlsx` | `Lab02-Diagnostico-Ausentismo/` | 3 hojas: Ausentismo (42 filas), Headcount (18 filas), Clima (12 filas) |
| 2 | `Hallazgos_Paso2_Analisis_Descriptivo.docx` | `Lab02-Diagnostico-Ausentismo/` | Incluye secciones de Pasos 2, 3, 4 y 7 con respuestas de Copilot |
| 3 | `Checklist_Auditoria_Laboral.docx` | `Lab02-Diagnostico-Ausentismo/` | Mínimo 23 ítems en 5 categorías, con encabezado de advertencia |
| 4 | `Informe_Ejecutivo_Diagnostico_Ausentismo.docx` | `Lab02-Diagnostico-Ausentismo/` | 7 secciones, ≤600 palabras, 5 recomendaciones con responsable y plazo |

### Verificación de cálculos clave

Comprueba manualmente los siguientes valores de referencia:

| Indicador | Departamento | Mes | Valor esperado (aprox.) |
|---|---|---|---|
| Índice de ausentismo | Producción | Enero | ~0.92% |
| Índice de ausentismo | Producción | Junio | ~2.05% |
| Tasa de rotación mensual | Producción | Enero | ~2.25% |
| Tasa de rotación semestral | Producción | Ene–Jun | ~14.0% |
| eNPS | Producción | Mayo | -15 |
| Favorabilidad top-box | Producción / Carga de trabajo | Mayo | 32% |

> 💡 **Nota:** Los valores de Copilot pueden diferir ligeramente de los valores de referencia por redondeo o método de cálculo. Una variación de ±0.5% es aceptable. Si la diferencia es mayor, revisa el prompt y verifica que los datos ingresados son correctos.

### Pregunta de validación final

Antes de cerrar la práctica, responde esta pregunta en tu documento de hallazgos:

> *"Si los datos de ausentismo de Producción fueran reales y los presentaras a la Dirección General mañana, ¿qué tres acciones de verificación adicionales realizarías antes de la presentación?"*

---

## 8. Solución de Problemas

### Problema 1: Copilot no procesa correctamente los datos en formato CSV pegado en el chat

**Síntoma:** Copilot devuelve una respuesta genérica, dice que no puede ver los datos, o los totales calculados no coinciden con los datos ingresados (por ejemplo, suma incorrecta de días perdidos).

**Causa probable:** El texto con formato de tabla o CSV puede truncarse si supera el límite de contexto del prompt, o Copilot interpreta incorrectamente los separadores de columna si hay caracteres especiales (como la "ó" en "Producción") que afectan el parseo.

**Solución:**
1. Divide el dataset en bloques más pequeños. En lugar de pegar los 42 registros en un solo prompt, envía primero solo los datos de **Producción** (15 registros), obtén el análisis, y luego envía los datos de **Ventas** y **Administración** por separado.
2. Si el problema persiste, simplifica los nombres: reemplaza "Producción" por "Prod", "Administración" por "Admin" en el texto del prompt (solo para el prompt, no en tu archivo Excel).
3. Verifica manualmente una suma simple antes de confiar en los totales de Copilot: suma la columna `Dias_Ausentes` para Producción en Enero (EMP-001 + EMP-002 + EMP-003 = 3 + 5 + 1 = 9) y compara con el valor reportado.
4. Si el error persiste, usa Excel para calcular las sumas por departamento/mes usando una tabla dinámica, y pega solo los **totales resumidos** (no los datos individuales) en el prompt de Copilot.

---

### Problema 2: El informe ejecutivo generado por Copilot supera una página o las recomendaciones son demasiado genéricas

**Síntoma:** El informe tiene más de 800 palabras, las recomendaciones usan frases como "mejorar la comunicación" o "fortalecer el liderazgo" sin especificidad, o la estructura no sigue el formato solicitado.

**Causa probable:** Los modelos de lenguaje tienden a generar texto extenso y recomendaciones genéricas cuando el prompt no establece restricciones de longitud y nivel de especificidad de manera suficientemente explícita, o cuando el contexto previo de la conversación es muy largo y Copilot "pierde" las instrucciones originales.

**Solución:**
1. **Para el problema de extensión:** Agrega al inicio del prompt de refinamiento: `"RESTRICCIÓN ESTRICTA: El informe completo NO debe superar 600 palabras. Cuenta las palabras antes de responder y ajusta si es necesario."` Copilot respetará esta restricción en la mayoría de los casos.
2. **Para recomendaciones genéricas:** Usa el prompt de refinamiento del Paso 6 (instrucción 4) que solicita explícitamente el formato: verbo de acción + objeto específico + responsable + plazo + indicador a mejorar + porcentaje estimado. Proporciona siempre un ejemplo del nivel de especificidad esperado dentro del mismo prompt.
3. **Para estructura incorrecta:** Inicia una **nueva conversación** en Copilot (esto limpia el contexto acumulado) y pega el prompt de redacción del Paso 6 nuevamente, incluyendo todos los hallazgos en un solo mensaje. Las conversaciones largas pueden degradar la calidad de las instrucciones de formato.
4. Si después de dos intentos el resultado sigue siendo insatisfactorio, redacta manualmente el esqueleto del informe en Word y usa Copilot solo para **completar secciones específicas**, por ejemplo: `"Redacta únicamente la sección de Recomendaciones Priorizadas con este formato [especifica formato]. Los demás datos ya los tengo."`.

---

## 9. Limpieza del Entorno

Al finalizar la práctica, realiza los siguientes pasos de cierre y organización:

1. **Organiza tu carpeta de OneDrive:** Verifica que los 4 entregables están en `Lab02-Diagnostico-Ausentismo/`. Crea una subcarpeta `Prompts_Usados/` y guarda un documento de texto con los prompts principales que utilizaste (esto te servirá como referencia para futuras prácticas).

2. **Cierra las conversaciones de Copilot:** Aunque Copilot Chat no almacena conversaciones de forma permanente en todos los contextos, es buena práctica cerrar la sesión de chat al terminar. No compartas capturas de pantalla de las conversaciones en canales públicos.

3. **Revisa la privacidad:** Confirma que en ningún momento ingresaste datos reales de empleados de tu organización. Si accidentalmente lo hiciste, contacta a tu administrador de TI para revisar las políticas de retención de datos de Copilot en tu tenant.

4. **Guarda los archivos Excel y Word:** Asegúrate de que todos los archivos están guardados y sincronizados en OneDrive (el ícono de nube debe mostrar una palomita ✓, no un símbolo de sincronización en progreso).

5. **Comparte con el instructor** (si aplica): Comparte la carpeta `Lab02-Diagnostico-Ausentismo/` con el correo del instructor a través de OneDrive con permisos de **solo lectura**. Esto permite la revisión sin riesgo de modificación.

---

## 10. Resumen

### Puntos Clave de la Práctica

En esta práctica aplicaste los conceptos de la Lección 2.1 en un escenario simulado de consultoría interna de RR.HH.:

- **Análisis descriptivo con Copilot:** Aprendiste a estructurar prompts de *data summarization* para extraer patrones de ausentismo de datasets tabulares, identificando que Producción concentra el mayor riesgo con una tendencia creciente.

- **Cálculo de indicadores clave:** Calculaste el índice de ausentismo mensual usando la fórmula `(Días perdidos / (HC_Promedio × Días_Laborables)) × 100` y la tasa de rotación semestral, comparándolas contra benchmarks del sector manufacturero mexicano.

- **Interpretación de clima laboral:** Aplicaste los conceptos de favorabilidad top-box (4–5 en escala Likert) y eNPS para identificar que Producción tiene un eNPS de -15 y una favorabilidad crítica en "Carga de trabajo" (32%), correlacionada con el alto ausentismo.

- **Checklist de auditoría:** Diseñaste con Copilot un checklist de 23+ ítems en 5 categorías (contratos, jornada, prestaciones, seguridad, ausentismo), aplicando el principio de que Copilot es un **asistente de estructuración**, no un sustituto del juicio experto.

- **Informe ejecutivo:** Integraste todos los hallazgos en un documento de una página con recomendaciones priorizadas y accionables, aplicando técnicas de refinamiento iterativo de prompts para mejorar la especificidad.

- **Evaluación crítica:** Desarrollaste pensamiento crítico sobre las limitaciones de la IA generativa en análisis de personas, incluyendo el riesgo de asumir causalidad donde solo existe correlación y la necesidad de validación profesional.

### Reflexión Final

> *"Copilot Chat es un amplificador analítico, no un oráculo. Su valor radica en acelerar el procesamiento de información y la estructuración de documentos, pero el criterio técnico, ético y legal siempre debe residir en el especialista de RR.HH."*

### Recursos Adicionales

| Recurso | Descripción | URL |
|---|---|---|
| ISO 30414 | Directrices para informes de capital humano | [iso.org/standard/69338.html](https://www.iso.org/standard/69338.html) |
| SHRM — Cálculo de rotación | Guía práctica de métricas de rotación | [shrm.org](https://www.shrm.org/resourcesandtools/tools-and-samples/how-to-guides/pages/how-to-calculate-employee-turnover.aspx) |
| Gallup State of the Workplace | Benchmarks globales de engagement y clima | [gallup.com/workplace](https://www.gallup.com/workplace/349484/state-of-the-global-workplace.aspx) |
| Microsoft Support — Copilot en Excel | Uso de Copilot para análisis de datos en Excel | [support.microsoft.com](https://support.microsoft.com/es-es/office/usar-copilot-en-excel-para-analizar-y-explorar-datos-165544b1-b1e1-4f95-953e-9bbf09c3e963) |
| DOF — Ley Federal del Trabajo | Versión vigente de la LFT mexicana | [dof.gob.mx](https://www.dof.gob.mx) |
| Nightingale DVS — Gráficos Likert | Mejores prácticas para visualizar datos de encuestas | [nightingaledvs.com](https://nightingaledvs.com/how-to-visualize-likert-scale-data/) |

---

> ⚠️ **Recordatorio final:** Los documentos generados en esta práctica son **exclusivamente para fines educativos** con datos ficticios. Antes de adaptar cualquier checklist, análisis o recomendación a una situación laboral real, consulta con un abogado laboralista certificado y verifica la versión vigente de la Ley Federal del Trabajo en el Diario Oficial de la Federación.

---
LAB_END---
