```text
███████╗ ██████╗  ██╗
██╔════╝██╔════╝  ██║
█████╗  ██║       ██║
██╔══╝  ██║       ██║
██║     ╚██████╗  ██║
╚═╝      ╚═════╝  ╚═╝

> **Sistema de Inteligencia de Medios, Monitoreo Reputacional y Clasificación Contextual con IA para la Fundación CardioInfantil (La Cardio)**  
> *Realizado por **Johnathan Cortés***

[![Live App](https://img.shields.io/badge/Streamlit_App-Online-FF4B4B?style=for-the-badge&logo=streamlit&logoColor=white)](https://cardio-api.streamlit.app/)
[![Python](https://img.shields.io/badge/Python-3.10%2B-3776AB?style=for-the-badge&logo=python&logoColor=white)](https://python.org)
[![OpenAI](https://img.shields.io/badge/OpenAI-API-412991?style=for-the-badge&logo=openai&logoColor=white)](https://openai.com)
[![Status](https://img.shields.io/badge/Status-v18.2--Stable-brightgreen?style=for-the-badge)]()

---

## 🚀 App en Vivo

🔗 **Acceso a la plataforma web:** [https://cardio-api.streamlit.app/](https://cardio-api.streamlit.app/)

---

## 📋 Tabla de Contenidos

- [Descripción General](#-descripción-general)
- [Características Principales](#-características-principales)
- [Flujo de Procesamiento](#-flujo-de-procesamiento)
- [Requisitos e Instalación](#-requisitos-e-instalación)
- [Configuración del Sistema](#-configuración-del-sistema)
- [Estructura del Proyecto](#-estructura-del-proyecto)
- [Taxonomía y Criterios de Clasificación](#-taxonomía-y-criterios-de-clasificación)
- [Modos de Análisis](#-modos-de-análisis)
- [Formato de Salida y Métricas](#-formato-de-salida-y-métricas)

---

## 🏥 Descripción General

**FCI (Análisis de Noticias - La Cardio)** es una plataforma web inteligente de procesamiento masivo, análisis reputacional y desduplicación de impactos noticiosos desarrollada para la **Fundación CardioInfantil (La Cardio)**.

El sistema recibe dossieres de noticias en formato Excel (`.xlsx`), estandariza la información de medios y regiones, elimina publicaciones repetidas mediante algoritmos de similitud y clasifica automáticamente con modelos de **OpenAI** (`gpt-4.1-nano-2025-04-14` y `text-embedding-3-small`) el **Tono reputacional**, la **Categoría del servicio** y la **Narrativa estratégica** asociada a La Cardio o a instituciones competidoras.

---

## ✨ Características Principales

* ⚡ **Desduplicación Inteligente Multinivel**:
  * **Prensa / Digital**: Comparación difusa (Fuzzy Matching) de títulos con `SequenceMatcher` (umbral 0.93) y Jaccard de tokens.
  * **Internet**: Normalización y cruce exacto de enlaces web (`Link Nota`, `Link Streaming`).
  * **Radio / Televisión**: Identificación de duplicados por coincidencia exacta de `(Mención, Medio, Hora)`.

* 🧠 **Embeddings Semánticos y Caché de Alto Rendimiento**:
  * Vectorización mediante `text-embedding-3-small` con sistema de caché en memoria (`EmbeddingCache`) basado en hashes MD5 para optimizar el consumo de tokens.
  * Agrupamiento jerárquico (`AgglomerativeClustering`) y algoritmo **DSU (Disjoint Set Union)**.

* 🏷️ **Clasificador Contextual con IA & Filtros Locales**:
  * Integración con `gpt-4.1-nano-2025-04-14` ejecutado mediante procesamiento concurrente en multihilos (`ThreadPoolExecutor`).
  * Homologación automática de marcas y sinónimos (ej: *Cardiovascular $\Leftrightarrow$ FCV*, *La Cardio $\Leftrightarrow$ CardioInfantil*, *Country $\Leftrightarrow$ Clínica del Country*).
  * Reglas locales de descarte para boletines genéricos (*"Noticias en 60 segundos"*), accidentes de tránsito, obras viales o sucesos políticos ajenos al negocio de salud.

* 🧹 **Limpieza Automática de Metadatos**:
  * Eliminación del prefijo `"La Cardio 26 - "` en la columna de menciones de empresa.
  * Corrección de tildes y eñes mediante diccionarios ortográficos (`_TILDE_MAP` y `_ENIE_MAP`).
  * Limpieza de caracteres y etiquetas HTML (`html.unescape`).

* 🔐 **Autenticación Integrada**:
  * Control de acceso mediante contraseña configurable en `st.secrets`.

* 📊 **Exportación Profesional a Excel**:
  * Generación de reportes con hipervínculos funcionales activados.
  * Aplicación de formatos numéricos para fechas (`DD/MM/YYYY`), millares (`#,##0`) y valores monetarios de nota/CPE (`$#,##0`).

---

## 🔄 Flujo de Procesamiento

```text
┌─────────────────────────────────────────────────────────────────┐
│                   1. DOSSIER ENTRADA (.XLSX)                    │
│    Lectura de celdas, extracción de links y normalización UTF8  │
└────────────────────────────────┬────────────────────────────────┘
                                 │
                                 ▼
┌─────────────────────────────────────────────────────────────────┐
│                 2. MAPEO REGIONAL & INTERNET                    │
│     Cruce dinámico con diccionario local (Configuracion.xlsx)    │
└────────────────────────────────┬────────────────────────────────┘
                                 │
                                 ▼
┌─────────────────────────────────────────────────────────────────┐
│                 3. DESDUPLICACIÓN MULTINIVEL                    │
│  Filtro por URL streaming, horario de emisión y fuzzy de título │
└────────────────────────────────┬────────────────────────────────┘
                                 │
                                 ▼
┌─────────────────────────────────────────────────────────────────┐
│                4. AGRUPAMIENTO Y EMBEDDINGS                     │
│    Vectorización semántica (OpenAI) + Agrupamiento por Títulos  │
└────────────────────────────────┬────────────────────────────────┘
                                 │
                                 ▼
┌─────────────────────────────────────────────────────────────────┐
│                5. ANÁLISIS INTEGRADO (LLM + REGLAS)             │
│   Multihilo síncrono para evaluar Tono, Categoría y Narrativa   │
└────────────────────────────────┬────────────────────────────────┘
                                 │
                                 ▼
┌─────────────────────────────────────────────────────────────────┐
│                   6. INFORME FINAL EXCEL                        │
│   Aplicación de estilos, hipervínculos y métricas de costo/tiempo│
└─────────────────────────────────────────────────────────────────┘
```

---

## 🛠️ Requisitos e Instalación

### Pre-requisitos
* **Python 3.10** o superior.
* API Key de OpenAI habilitada.

### Pasos para Ejecución Local

1. **Clonar el repositorio:**
   ```bash
   git clone https://github.com/tu-usuario/FC-API.git
   cd FC-API
   ```

2. **Crear y activar el entorno virtual:**
   ```bash
   # En macOS/Linux:
   python3 -m venv venv
   source venv/bin/activate

   # En Windows:
   python -m venv venv
   venv\Scripts\activate
   ```

3. **Instalar dependencias:**
   ```bash
   pip install streamlit pandas openpyxl openai unidecode numpy scikit-learn joblib requests
   ```

4. **Ejecutar la aplicación:**
   ```bash
   streamlit run app.py
   ```

---

## ⚙️ Configuración del Sistema

### 1. Variables Secretas (`.streamlit/secrets.toml`)

Crea la carpeta `.streamlit` en la raíz del proyecto y agrega el archivo `secrets.toml`:

```toml
OPENAI_API_KEY = "sk-proj-tu-api-key-de-openai-aqui"
APP_PASSWORD = "tu_contraseña_de_acceso"
```

> **Nota:** Si no se define `APP_PASSWORD`, la aplicación utilizará por defecto `admin123`.

### 2. Archivo de Mapeo de Medios (`Configuracion.xlsx`)

El sistema consulta el archivo `Configuracion.xlsx` en la raíz para asociar automáticamente los medios con su ubicación geográfica y canal correspondiente. Debe contener dos hojas:
* **Regiones**: Mapeo `Medio` $\rightarrow$ `Región`.
* **Internet**: Normalización de nombres de portales web.

---

## 📁 Estructura del Proyecto

```text
FC-API/
├── .streamlit/
│   └── secrets.toml         # Variables de entorno y credenciales (OpenAI/Password)
├── app.py                   # Aplicación principal Streamlit e integración de IA
├── Configuracion.xlsx       # Matriz de consulta Medio -> Región / Internet
├── README.md                # Documentación del repositorio
└── requirements.txt         # Lista de dependencias del proyecto
```

---

## 🏷️ Taxonomía y Criterios de Clasificación

### **1. Tono Reputacional (`Tono IA`)**
* 🟢 **Positivo**: Reconocimientos, acreditaciones internacionales (JCI), avances médicos, cirugías exitosas o aportes científicos.
* 🔴 **Negativo**: Demandas, fallas operativas, quejas médicas o afectación directa a la reputación corporativa.
* ⚪ **Neutro**: Menciones referenciales, contexto de sector o informaciones sin sesgo implícito.
* ⬛ **Duplicada**: Filas identificadas como copias de otra noticia ya procesada.

### **2. Categorías (`Categoría`)**
| Categoría | Descripción / Criterio |
| :--- | :--- |
| **Core** | Servicios enfocados en el core cardiovascular y trasplantes (ej. *Infarto, Arritmias, Marcapasos, Válvulas*). |
| **Especialidades** | Otras áreas médicas de la institución (ej. *Neurología, Pediatría, Oncología, Urgencias*). |
| **Ranking** | Posicionamiento en escalafones de reputación y calidad (ej. *Merco, América Economía, P&M*). |
| **Reforma** | Proyectos de ley, debates normativos y regulaciones del sector (ej. *Reforma a la Salud, Minsalud*). |
| **Sector** | Coyuntura del sistema de salud, crisis de EPS, embargos o contexto general de clínicas. |
| **Corporativo** | Alianzas, asambleas de socios, acreditaciones, convenios o eventos institucionales. |
| **Sucesos** | Noticia donde se cita la marca por cercanía geográfica pero ajena al negocio (ej. *Atracos, cierres viales*). |

### **3. Narrativas Estratégicas (`Narrativas`)**
| Narrativa | Descripción / Criterio |
| :--- | :--- |
| **Excelencia médica** | Experticia clínica, alta complejidad, calidad en la atención y acreditación internacional. |
| **Innovación + Desarrollo** | Tecnología de punta, robótica (Da Vinci), investigación (Nature Index), patentes y telemedicina. |
| **Sostenibilidad** | Acciones de responsabilidad social, brigadas médicas, voluntariado y proyectos con propósito. |
| **Marca empleadora** | Reconocimiento a colaboradores, médicos "Orgullo Cardio", bienestar y talento humano. |
| **Portafolio** | Chequeos ejecutivos, prevenciones de salud, nutrición y recomendaciones de vida saludable. |
| **Otras** | Menciones puramente referenciales o menciones secundarias. |

---

## 💻 Modos de Análisis

1. **Análisis Completo**:
   * Procesa el dossier completo de prensa en formato Excel oficial.
   * Aplica desduplicación multinivel por URL, horario y título.
   * Asigna mapeos regionales, estilos visuales e hipervínculos funcionales en el reporte final.
   * Genera un informe detallado de costos y tokens consumidos.

2. **Análisis Rápido**:
   * Permite evaluar rápidamente datasets simples seleccionando únicamente las columnas de `Título` y `Resumen`.

---

## 📊 Formato de Salida y Métricas

Al finalizar la ejecución, la aplicación muestra métricas clave del proceso:
* **Total de Filas Procesadas**
* **Noticias Únicas vs. Duplicadas**
* **Tiempo Total de Procesamiento**
* **Costo Estimado en API (USD)**
* **Tasa de Eficiencia de Caché de Embeddings**

---

<p align="center">
  <sub>Desarrollado para la <b>Fundación CardioInfantil (La Cardio)</b> por <b>Johnathan Cortés</b> | App: <a href="https://cardio-api.streamlit.app/">cardio-api.streamlit.app</a></sub>
</p>
```
