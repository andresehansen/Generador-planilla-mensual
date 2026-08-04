# Generador de Planillas de Distribución de Guardia y Turnos

Herramienta web estática desarrollada en **Python (PyScript)** para la automatización de planillas de turnos, francos, Home Office y guardias de monitores en Excel (`.xlsx`), para la oficina de Herrera.

---

## 📌 Características Principales

1. **Reglas de Cobertura Estrictas:**
   - **Mañana (mínimo 4 personas):** Garantiza al menos 1 integrante de `Grupo GIS` y al menos 1 integrante de `Dúos/Tríos`.
   - **Tarde (mínimo 2 personas):** Garantiza al menos 1 integrante de `Grupo GIS`. `Herrera` e `Itzi` están bloqueados del turno tarde.
2. **Home Office (AHO / BHO):**
   - **AHO:** Aranda, Alegre, Nahuel, Fuentes, Romero.
   - **BHO:** Ledesma, Lezcano, Itzi, Navarro, Gonzalote.
   - **HO Fijo:** Aranda los martes por estudio.
3. **Francos e Incompatibilidades:**
   - 1 franco por semana.
   - Bloqueo de francos los días inmediatos **anteriores** o **posteriores** al Home Office.
   - **Anulación automática de francos:** Si la semana contiene un feriado nacional, se anulan los francos individuales de esa semana para todo el personal.
4. **Turno de Monitores (`< 07:00` marcado con `*`):**
   - Asignación de 1 persona rotativa diaria en la mañana.
   - Excluidos permanentemente: Itzi, Alegre, Nahuel, Gonzalote, Coronel, Tamara, Herrera.
   - Rotación entre elegibles: Aranda, Fuentes, Romero, Ledesma, Lezcano, Navarro.
5. **Consulta de Feriados en Vivo:**
   - Consume automáticamente la API de feriados nacionales de Argentina (`nolaborables.com.ar`).

---

## 🚀 Despliegue en GitHub Pages (Paso a Paso)

1. **Crear Repositorio:**
   Crea un nuevo repositorio en GitHub (ej: `generador-planillas`).

2. **Subir los Archivos:**
   Sube los siguientes archivos a la raíz de tu repositorio:
   - `index.html`
   - `main.py`
   - `style.css`
   - `README.md`

3. **Activar GitHub Pages:**
   - Ve a **Settings** > **Pages** en tu repositorio.
   - En **Source**, selecciona **Deploy from a branch**.
   - En **Branch**, elige `main` (o `master`) y guarda (`Save`).

4. **Acceso Web:**
   En 1 a 2 minutos la URL estará activa:
   `https://tu-usuario.github.io/generador-planillas/`

Herrera o cualquier operador podrá ingresar a ese enlace desde su navegador, configurar las licencias o feriados del mes y descargar directamente el Excel formateado.
