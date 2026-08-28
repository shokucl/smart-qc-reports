# Smart QC Reports (Gestor automatizado PNC)

[Read in English](#english) | [Leer en Español](#español)

---

<a name="english"></a>
## About this Project

**Note on Language:** *The source code, variables, and graphical user interface (GUI) of this application are written in Spanish. This is because it is not a theoretical exercise, but a real-world software tool designed to eliminate repetitive manual work in quality control.*

### What is it?
It is a desktop application developed in Python for the automated management of production Non-Conforming Product (PNC) data.

### The Problem it Solves
This system was created out of the need to eliminate repetitive manual work in the Quality area. Previously, creating reports required a much more manual and slow process of generating pivot tables. This tool made the workflow much easier and simpler:
* **Synchronization:** Downloads data directly from SharePoint.
* **Error-Proof Cleaning:** The cleaning logic does not look for exact column names, but text fragments. If a header is mistyped in the source Excel upon data entry, the program detects it, corrects it internally, and prevents crashes.
* **Dynamic Charts:** Generates Pareto charts (80/20), Area vs. Defect heatmaps, and evolutionary time trends.
* **AI Integration:** Connects via the Llama-3 API (Groq) to analyze KPI data and automatically draft an executive summary for management.
* **Export:** Renders an HTML template and converts it invisibly into a corporate PDF.

### Tech Stack
* **Python 3**
* **Pandas / NumPy** (Data Aggregation & Cleaning)
* **Matplotlib** (Data Visualization)
* **Tkinter** (GUI)
* **Jinja2** (HTML to PDF rendering)
* **PyInstaller** (Executable compilation)

---

<a name="español"></a>
## Acerca de este Proyecto

### ¿Qué es?
Es una aplicación de escritorio desarrollada en Python para la gestión automatizada de datos de Producto No Conforme (PNC) de producción.

### El Problema que Resuelve
Este sistema fue creado por la necesidad de eliminar el trabajo manual repetitivo en el área de Calidad. Antes, requería hacer reportes de forma más manual y lenta generando tablas dinámicas. Esta herramienta hizo mi trabajo más fácil y simple:
* **Sincronización:** Descarga la data directamente desde SharePoint.
* **Limpieza a Prueba de Errores:** La lógica de limpieza no busca nombres de columnas exactos, sino fragmentos de texto. Si al ingresar un dato se escribe mal un encabezado en el Excel de origen, el programa lo detecta, lo corrige internamente y no se rompe.
* **Gráficos Dinámicos:** Genera diagramas de Pareto (80/20), mapas de calor de Área vs. Defecto y tendencias evolutivas temporales.
* **Uso de IA:** Se conecta a través de la API de Llama-3 (Groq) para analizar los datos de KPIs y redactar un resumen ejecutivo para gerencia de forma automática.
* **Exportación:** Renderiza una plantilla HTML y la convierte a un PDF corporativo de forma invisible.

### Tech Stack
* **Python 3**
* **Pandas / NumPy** (Agregación y Limpieza de Datos)
* **Matplotlib** (Visualización de Datos)
* **Tkinter** (Interfaz Gráfica)
* **Jinja2** (Renderizado de HTML a PDF)
* **PyInstaller** (Compilación de Ejecutable)
