# Verificación de Niños en ONN Bolivia

Flujo de trabajo reutilizable para consultar en `https://onnbolivia.com` qué personas de un archivo `.xlsx` están registradas en **Lista de niños** y generar un reporte PDF con el resultado.

## Estructura del proyecto

- `.env` / `.env.example`: credenciales de acceso al portal ONN Bolivia (`ONN_BASE_URL`, `ONN_USERNAME`, `ONN_PASSWORD`). `.env` está excluido en `.gitignore`.
- `.agents/skills/verificar-ninos-onn/SKILL.md`: skill de Antigravity invocable en el chat con `/verificar-ninos-onn`.
- `.agents/skills/verificar-ninos-onn/scripts/leer_excel.py`: lector `.xlsx` sin dependencias externas.
- `.agents/skills/verificar-ninos-onn/scripts/verificar_onn.py`: script principal de consulta (solo lectura) y cruce de datos.
- `.agents/skills/verificar-ninos-onn/scripts/generar_pdf.py`: generador de reporte PDF con Chromium headless.
- `reportes/`: carpeta donde se guardan los JSON de resultados y los PDFs generados.

## Uso desde Antigravity

Escribe en el chat:

```text
/verificar-ninos-onn yaque.xlsx
```

O con cualquier otro archivo `.xlsx` del proyecto.

## Uso desde la terminal

1. Asegúrate de tener configurado `.env` (puedes copiar `.env.example` a `.env`).
2. Ejecuta:

```bash
python3 .agents/skills/verificar-ninos-onn/scripts/verificar_onn.py yaque.xlsx
```

El reporte se guardará en `reportes/reporte_yaque.pdf` y el detalle estructurado en `reportes/resultados_yaque.json`.
# workflow
