```
 ██████╗  █████╗  ██████╗  ██████╗  ██╗  ██████╗   █████╗  ██████╗  ██╗
██╔════╝ ██╔══██╗ ██╔══██╗ ██╔══██╗ ██║ ██╔═══██╗ ██╔══██╗ ██╔══██╗ ██║
██║      ███████║ ██████╔╝ ██║  ██║ ██║ ██║   ██║ ███████║ ██████╔╝ ██║
██║      ██╔══██╗ ██╔══██╗ ██║  ██║ ██║ ██║   ██║ ██╔══██╗ ██╔═══╝  ██║
╚██████╗ ██║  ██║ ██║  ██║ ██████╔╝ ██║ ╚██████╔╝ ██║  ██║ ██║      ██║
 ╚═════╝ ╚═╝  ╚═╝ ╚═╝  ╚═╝ ╚═════╝  ╚═╝  ╚═════╝  ╚═╝  ╚═╝ ╚═╝      ╚═╝
```

<div align="center">

# CardioAPI · Análisis de Noticias

**Motor de reputación y clasificación de prensa impulsado por IA para la Fundación CardioInfantil**

[![Streamlit](https://img.shields.io/badge/Streamlit-App-FF4B4B?logo=streamlit&logoColor=white)](https://cardio-api.streamlit.app/)
[![Python](https://img.shields.io/badge/Python-3.10+-3776AB?logo=python&logoColor=white)](https://www.python.org/)
[![OpenAI](https://img.shields.io/badge/OpenAI-GPT--4.1--nano-412991?logo=openai&logoColor=white)](https://openai.com/)
[![Version](https://img.shields.io/badge/version-v18.2-blueviolet)]()
[![License](https://img.shields.io/badge/License-Privado-lightgrey)]()

**[🔗 Ver aplicación en vivo](https://cardio-api.streamlit.app/)**

</div>

---

## 📋 Descripción

**CardioAPI** es una aplicación **Streamlit** que automatiza el análisis de reputación de prensa para la **Fundación CardioInfantil (La Cardio)** y su ecosistema de marcas asociadas. A partir de dossiers de monitoreo de medios (prensa, radio, televisión, internet y revistas), aplica un pipeline de IA para determinar, por cada mención de marca:

- **Tono** de reputación (Positivo / Negativo / Neutro)
- **Categoría** de involucramiento institucional (Sucesos, Core, Especialidades, Ranking, Sector, Reforma, Corporativo)
- **Narrativa** estratégica del mensaje (Sostenibilidad, Excelencia médica, Innovación + Desarrollo, Marca empleadora, Portafolio, Obras, Otras)

---

## ✨ Funciones actuales

- 🏥 **Clasificador contextual multi-marca** — Detecta menciones de "La Cardio" y marcas relacionadas (Country, Cardiovascular/FCV, Santa Fe, Shaio, etc.) mediante un diccionario de sinónimos, evaluando el tono, la categoría y la narrativa específicos hacia cada entidad.
- 🧩 **Agrupamiento estricto por título + mención** — Combina similitud de texto (`SequenceMatcher` + Jaccard de tokens) con normalización de acentos para que noticias homólogas reciban la misma clasificación.
- ⚡ **Procesamiento paralelo** — Utiliza `ThreadPoolExecutor` para ejecutar múltiples llamadas a la API de OpenAI de forma simultánea y estable dentro de Streamlit.
- 🧠 **Caché de embeddings con hash MD5** — Evita recalcular vectores para textos ya procesados, reduciendo costo y tiempo de ejecución.
- 🚫 **Filtros heurísticos locales** — Excluye titulares de "última hora", boletines genéricos, sucesos viales/accidentes y coyuntura política ajena al sector salud, sin invocar la IA.
- 🔍 **Detección avanzada de duplicados** — Combina URL normalizada, medio + mención y similitud de título, con lógica distinta para impresos/digitales frente a radio y televisión.
- 🧹 **Limpieza y homologación de menciones** — Quita el prefijo `La Cardio 26 - ` y homologa automáticamente la marca Country; expande filas cuando hay varias menciones separadas por `;`.
- 🗺️ **Mapeo de medios** — Usa `Configuracion.xlsx` (hojas `Regiones` e `Internet`) para asignar región y tipo de medio al dossier institucional.
- 🔐 **Acceso protegido por contraseña** — Autenticación simple vía `st.secrets` con fallback local para desarrollo.
- 💰 **Estimador de costos en tiempo real** — Calcula el gasto en tokens de entrada, salida y embeddings por cada corrida, junto con métricas de total, únicas, duplicados y duración.
- 📥 **Dos modos de uso** — *Análisis Completo* (dossier estructurado, formato institucional) y *Análisis Rápido* (cualquier Excel con columnas de título/resumen).
- 🎨 **Interfaz moderna estilo Material** — Diseño personalizado con tipografía Google Sans, tarjetas de métricas y estados de progreso por pasos.

---

## 🛠️ Stack tecnológico

| Componente             | Tecnología                                |
|-------------------------|---------------------------------------------|
| Framework web            | Streamlit                                   |
| Procesamiento de datos   | pandas, NumPy                               |
| Embeddings               | OpenAI `text-embedding-3-small`             |
| Clasificación             | OpenAI `gpt-4.1-nano`                       |
| Clustering semántico      | scikit-learn (`AgglomerativeClustering`, `cosine_similarity`) |
| Concurrencia              | `concurrent.futures.ThreadPoolExecutor`     |
| Manejo de Excel            | openpyxl                                    |
| Normalización de texto     | unidecode, regex, mapas de tildes/eñes propios |

---

## 🚀 Uso

1. Ingresa a la aplicación: **[cardio-api.streamlit.app](https://cardio-api.streamlit.app/)**
2. Autentícate con la contraseña configurada.
3. Elige el modo de trabajo:
   - **Análisis Completo**: carga el dossier institucional en `.xlsx`, define la marca principal y sus alias, y ejecuta el pipeline completo (normalización → duplicados → embeddings → clasificación IA → informe).
   - **Análisis Rápido**: sube cualquier Excel, selecciona las columnas de título y resumen, define marca y alias, y obtén una clasificación exprés.
4. Descarga el informe `.xlsx` con las columnas `Tono IA`, `Categoría` y `Narrativas` completadas.

---

## ⚙️ Configuración local (desarrollo)

```bash
git clone https://github.com/johnathanacortesd/CardioAPI.git
cd CardioAPI
pip install -r requirements.txt
```

Crea un archivo `.streamlit/secrets.toml`:

```toml
OPENAI_API_KEY = "tu-api-key-aqui"
APP_PASSWORD = "tu-contraseña-de-acceso"
```

Ejecuta la aplicación:

```bash
streamlit run app.py
```

> El proyecto requiere un archivo `Configuracion.xlsx` en la raíz del repositorio con las hojas `Regiones` e `Internet` para el mapeo de medios.

---

## 📁 Categorías y narrativas de clasificación

**Categorías institucionales:**
`Sucesos` · `Core` · `Especialidades` · `Ranking` · `Sector` · `Reforma` · `Corporativo`

**Narrativas estratégicas (en main):**
`Sostenibilidad` · `Excelencia médica` · `Innovación + Desarrollo` · `Marca empleadora` · `Portafolio` · `Obras` · `Otras`

`Obras` cubre construcción, ampliación, remodelación o infraestructura física de LaCardio (por ejemplo, ampliación de la torre quirúrgica o remodelación de urgencias). No se confunde con la categoría `Sucesos`, que sigue reservada a eventos ajenos al negocio (cierres u obras *viales* por cercanía geográfica).

---

## 👤 Autoría

Desarrollado por **Johnathan Cortés** para **GlobalNews Group**, como parte de sus soluciones de inteligencia de medios y análisis reputacional aplicadas al sector salud.

---

<div align="center">

*CardioAPI · Análisis de Noticias · Fundación CardioInfantil · GlobalNews Group*

</div>
