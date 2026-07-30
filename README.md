
 ╔══════════════════════════════════════════════════════════════════════════════════════════════╗
 ║                                                                                              ║
 ║    ██████╗██████╗ ██████╗ ██████╗ ██╗██████╗     █████╗ ██████╗ ██╗                          ║
 ║   ██╔════╝██╔══██╗██╔══██╗██╔══██╗██║██╔═══██╗   ██╔══██╗██╔══██╗██║                          ║
 ║   ██║     ███████║██████╔╝██║  ██║██║██║   ██║ - ███████║██████╔╝██║                          ║
 ║   ██║     ██╔══██║██╔══██╗██║  ██║██║██║   ██║   ██╔══██║██╔═══╝ ██║                          ║
 ║   ╚██████╗██║  ██║██║  ██║██████╔╝██║╚██████╔╝   ██║  ██║██║     ██║                          ║
 ║    ╚═════╝╚═╝  ╚═╝╚═╝  ╚═╝╚═════╝ ╚═╝ ╚═════╝    ╚═╝  ╚═╝╚═╝     ╚═╝                          ║
 ║                                                                                              ║
 ║                    Plataforma de Inteligencia de Prensa & Reputación Médica                      ║
 ║                              Desarrollado para La Cardio (Colombia)                          ║
 ╚══════════════════════════════════════════════════════════════════════════════════════════════╝
```

[![Streamlit](https://img.shields.io/badge/Streamlit-1.30+-FF4B4B?style=for-the-badge&logo=Streamlit&logoColor=white)](https://streamlit.io/)
[![OpenAI](https://img.shields.io/badge/OpenAI-GPT--4.1--nano-412991?style=for-the-badge&logo=openai&logoColor=white)](https://openai.com/)
[![Python](https://img.shields.io/badge/Python-3.10+-3776AB?style=for-the-badge&logo=python&logoColor=white)](https://python.org)
[![License](https://img.shields.io/badge/Licencia-Propietario-orange?style=for-the-badge)](LICENSE)

---

### ◈ Análisis Automatizado de Noticias y Reputación Institucional con IA

**Realizado por Johnathan Cortés** · *Versión 18.2*

</div>

---

## 📖 Tabla de Contenidos

1. [Acerca del Proyecto](#-acerca-del-proyecto)
2. [Características Principales](#-características-principales)
3. [Flujo de Procesamiento y Arquitectura](#-flujo-de-procesamiento-y-arquitectura)
4. [Taxonomía de Clasificación](#-taxonomía-de-clasificación)
5. [Requisitos Previos](#-requisitos-previos)
6. [Instalación y Configuración](#-instalación-y-configuración)
7. [Estructura del Proyecto](#-estructura-del-proyecto)
8. [Modelos de IA y Costos](#-modelos-de-ia-y-costos)
9. [Uso de la Aplicación](#-uso-de-la-aplicación)

---

## 🚀 Acerca del Proyecto

**Cardio-API** es una solución de analítica avanzada de medios desarrollada para la **Fundación CardioInfantil (La Cardio)** en Colombia. Su objetivo principal es automatizar el procesamiento, limpieza, deduplicación y análisis contextual reputacional de clipping de noticias provenientes de diversos tipos de medio: Prensa, Internet, Radio, Televisión y Revistas.

El sistema implementa un motor híbrido que combina **reglas heurísticas locales**, **embeddings vectoriales simétricos con caché** y **modelos de lenguaje generativo de última generación (GPT-4.1-nano)** para clasificar con precisión el tono reputacional, la categoría temática y la narrativa estratégica asignada a cada impacto de marca.

---

## ✨ Características Principales

* 📥 **Ingesta Inteligente de Dossiers:** Normalización automatizada de datos provenientes de archivos Excel de agencias de clipping.
* 🧹 **Limpieza y Homologación:**
  * Eliminación de prefijos/sufijos de archivo (`La Cardio 26 - `, etc.).
  * Mapeo dinámico regional y de medios web mediante `Configuracion.xlsx`.
  * Detección y homologación inteligente de sinónimos de marcas competidoras y aliadas (`FCV` $\leftrightarrow$ `Cardiovascular`, `Country` $\leftrightarrow$ `Clínica del Country`, `CardioInfantil` $\leftrightarrow$ `La Cardio`).
* 🔍 **Deduplicación Multinivel:**
  * Deduplicación estricta por URLs (Internet) y enlaces de streaming.
  * Deduplicación por coincidencia de emisión broadcast (Medio + Franja Horaria en Radio/TV).
  * Deduplicación por similitud de título utilizando `SequenceMatcher` y correspondencia de tokens Jaccard ($\ge 93\%$).
* 🤖 **Análisis Contextual Inteligente (ThreadPoolExecutor Concurrent):**
  * Procesamiento paralelo multi-hilo de alto rendimiento.
  * Evaluación simultánea e integrada de **Tono**, **Categoría** y **Narrativa**.
  * Filtros locales inmediatos para descartar boletines genéricos, noticias de tránsito/accidentes y coyuntura política ajena a temas de salud.
* ⚡ **Optimización y Caché:**
  * Caché vectorial en memoria (`EmbeddingCache` MD5) para minimizar consumo de la API de OpenAI.
  * Agrupamiento previo por similitud de títulos para garantizar consistencia absoluta y reducir tokens procesados.
* 📊 **Exportación Excel de Nivel Ejecutivo:**
  * Formatos numéricos condicionales (Moneda CPE, millares en audiencias, hipervínculos funcionales y fechas estandarizadas).

---

## 🏗 Flujo de Procesamiento y Arquitectura

```text
               ┌─────────────────────────────────────────┐
               │    Entrada: Dossier de Noticias (.xlsx) │
               └────────────────────┬────────────────────┘
                                    │
                                    ▼
       ┌────────────────────────────────────────────────────────┐
       │ 1. Normalización de Campos y Mapeo Regional/Internet   │
       │    (Lectura de Configuracion.xlsx)                     │
       └────────────────────────────┬───────────────────────────┘
                                    │
                                    ▼
       ┌────────────────────────────────────────────────────────┐
       │ 2. Deduplicación Multinivel                             │
       │    - Coincidencia de URLs / Streaming                   │
       │    - Broadcast (Medio + Hora)                          │
       │    - Similitud de Títulos (SequenceMatcher >= 0.93)    │
       └────────────────────────────┬───────────────────────────┘
                                    │
                                    ▼
       ┌────────────────────────────────────────────────────────┐
       │ 3. Agrupamiento por Títulos Homólogos                  │
       │    - Generación de Embeddings (text-embedding-3-small) │
       │    - Búsqueda en EmbeddingCache en memoria            │
       └────────────────────────────┬───────────────────────────┘
                                    │
                                    ▼
       ┌────────────────────────────────────────────────────────┐
       │ 4. Clasificación Contextual Inteligente                │
       │    - Exclusión Local (Tránsito, Política, Boletines)   │
       │    - Evaluación Paralela LLM (gpt-4.1-nano-2025-04-14) │
       └────────────────────────────┬───────────────────────────┘
                                    │
                                    ▼
               ┌─────────────────────────────────────────┐
               │  Salida: Informe Formateado (.xlsx)     │
               └─────────────────────────────────────────┘
```

---

## 🏷 Taxonomía de Clasificación

El sistema evalúa cada impacto bajo tres ejes principales:

### 1. Tono Reputacional
| Tono | Criterio de Asignación |
| :--- | :--- |
| **Positivo** | Galardones, acreditaciones (Joint Commission International), avances científicos, cirugías exitosas o hitos en salud. |
| **Negativo** | Demandas, fallas operativas o clínicas, quejas directas o crisis reputacionales. |
| **Neutro** | Menciones contextuales de marco sectorial, referencias secundarias o datos informativos. |

### 2. Categoría Institucional
* 🏥 **Core:** Servicios enfocados en la estrategia cardiovascular y trasplantes (exclusivo para la marca).
* 🩺 **Especialidades:** Oncología, pediatría, neurología, ortopedia, urgencias, entre otras.
* 🏆 **Ranking:** Posicionamiento en monitores de reputación (Merco, América Economía, P&M).
* 🏛 **Reforma:** Proyectos de ley, discusiones en el Congreso y regulaciones a la salud.
* 💼 **Sector:** Situación financiera de EPS, embargos, decisiones de Minsalud y Supersalud.
* 🤝 **Corporativo:** Alianzas, asambleas de socios, eventos institucionales (ej. Latidos Futuros).
* 🚨 **Sucesos:** Hechos fortuitos o cierres viales vinculados por cercanía geográfica.

### 3. Narrativas Estratégicas
* 🌿 **Sostenibilidad:** Responsabilidad social, brigadas médicas, voluntariado, donaciones.
* 🥇 **Excelencia Médica:** Reacreditación de calidad, experiencia clínica de alta complejidad.
* 💡 **Innovación + Desarrollo:** Investigaciones (Nature Index), nuevos equipos, patentes.
* 👥 **Marca Empleadora:** Logros y bienestar de colaboradores y cuerpo médico.
* 📋 **Portafolio:** Consejos de salud, promoción de la salud y prevención de la enfermedad.
* 📂 **Otras:** Referencial general o menciones de contexto no estratégico.

---

## ⚙️ Requisitos Previos

* **Python:** 3.10 o superior.
* **API Key de OpenAI:** Requerida para generar los embeddings vectoriales y la clasificación LLM.
* **Archivo de Configuración:** Archivo Excel `Configuracion.xlsx` ubicado en la raíz del proyecto.

---

## 📦 Instalación y Configuración

1. **Clonar el repositorio:**
   ```bash
   git clone https://github.com/tu-usuario/cardio-api.git
   cd cardio-api
   ```

2. **Crear y activar un entorno virtual:**
   ```bash
   # En Linux / macOS
   python3 -m venv venv
   source venv/bin/activate

   # En Windows
   python -m venv venv
   venv\Scripts\activate
   ```

3. **Instalar dependencias:**
   ```bash
   pip install -r requirements.txt
   ```
   *(O instalación directa de librerías clave:)*
   ```bash
   pip install streamlit pandas openpyxl openai scikit-learn unidecode numpy requests
   ```

4. **Configurar los secretos de Streamlit (`.streamlit/secrets.toml`):**
   Crea la carpeta `.streamlit` en la raíz y añade el archivo `secrets.toml`:
   ```toml
   OPENAI_API_KEY = "sk-proj-tu-api-key-aqui"
   APP_PASSWORD = "admin123"  # Contraseña de acceso a la interfaz web
   ```

5. **Verificar el archivo `Configuracion.xlsx`:**
   Debe contener las pestañas con la siguiente estructura:
   * **`Regiones`**: Columna A (Nombre del medio) | Columna B (Región).
   * **`Internet`**: Columna A (Medio fuente) | Columna B (Nombre normalizado).

6. **Ejecutar la aplicación:**
   ```bash
   streamlit run app.py
   ```

---

## 📂 Estructura del Proyecto

```text
cardio-api/
├── .streamlit/
│   └── secrets.toml          # Secretos (API Key de OpenAI y password de acceso)
├── app.py                    # Aplicación principal Streamlit & Engine
├── Configuracion.xlsx        # Mapeo de Regiones y Medios de Internet
├── README.md                 # Documentación del proyecto
└── requirements.txt          # Dependencias del proyecto
```

---

## 🤖 Modelos de IA y Estructura de Costos

La solución hace uso de los endpoints de la API de OpenAI optimizando el gasto de tokens:

| Componente | Modelo Utilizado | Costo Estimado / 1M Tokens |
| :--- | :--- | :--- |
| **Embeddings Vectoriales** | `text-embedding-3-small` | $0.02 USD |
| **Clasificación Contextual** | `gpt-4.1-nano-2025-04-14` | $0.10 USD (Input) / $0.40 USD (Output) |

> 💡 **Ahorro de Tokens:** Mediante el agrupamiento previo por similitud de títulos y el uso de caché en memoria (`EmbeddingCache`), el gasto total por dossier se reduce entre un **60% y un 80%**.

---

## 🖥 Uso de la Aplicación

1. **Ingreso:** Autentícate con la contraseña definida en `secrets.toml`.
2. **Pestaña "Análisis Completo":**
   * Configura la **Marca principal** (ej: `La Cardio`) y sus **Alias** (ej: `Fundación CardioInfantil;LaCardio;Cardio Infantil;FVDL;Country`).
   * Selecciona el modo `API de OpenAI`.
   * Carga el archivo `.xlsx` del Dossier.
   * Haz clic en **`▶ Iniciar análisis`**.
   * Revisa la tabla de métricas (Noticias Totales, Únicas, Duplicadas, Tiempo y Costo estimado en USD) y descarga el informe generado.
3. **Pestaña "Análisis Rápido":**
   * Permite procesar archivos Excel genéricos seleccionando manualmente las columnas de título y resumen para obtener análisis express.

---

<div align="center">

**Cardio-API · Inteligencia de Prensa & Reputación Médica**  
*Johnathan Cortés © 2026*

</div>
```
