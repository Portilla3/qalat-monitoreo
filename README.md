# QALAT · Sistema de Monitoreo de Resultados de Tratamiento
## App de análisis automático TOP / IRT

### Cómo instalar y correr (computador local)

#### 1. Instalar Python
Si no tienes Python, descárgalo de https://www.python.org (versión 3.10 o superior)

#### 2. Instalar dependencias
Abre la terminal (o cmd en Windows), navega a esta carpeta y ejecuta:
```
pip install -r requirements.txt
```

#### 3. Correr la app
```
streamlit run app.py
```
Se abre automáticamente en el navegador en http://localhost:8501

---

### Cómo publicar en la web (Streamlit Cloud) — gratis

1. Sube esta carpeta a un repositorio GitHub
2. Ve a https://share.streamlit.io
3. Conecta tu repositorio
4. Selecciona `app.py` como archivo principal
5. Clic en Deploy

La app queda disponible en una URL pública que puedes compartir con los países.

---

### Estructura
```
qalat_app/
├── app.py                  # Interfaz Streamlit
├── pipeline/
│   ├── wide_top.py         # Motor TOP (basado en SCRIPT_TOP_Universal_Wide_v3_6)
│   └── wide_irt.py         # Motor IRT (próxima versión)
├── requirements.txt
└── README.md
```

---

### Qué genera la app

| Output | Formato | Descripción |
|--------|---------|-------------|
| Base Wide | Excel (.xlsx) | 6 hojas: Base Wide · Resumen · Alertas · Calidad · Por Centro · Pendientes |
| Gráficos | PNG | Seguimiento · Semáforo · Sustancia principal |
| Pendientes | CSV | Lista de pacientes con TOP2 urgente o próximo |

---

### Versiones futuras
- [ ] Módulo IRT
- [ ] Reporte PDF automático
- [ ] Presentación PPT automática
- [ ] Tablero comparativo entre países
- [ ] Login por país

---
Desarrollado para Proyecto QALAT · UNODC · 2026
