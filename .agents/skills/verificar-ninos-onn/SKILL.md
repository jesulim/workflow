---
name: verificar-ninos-onn
description: >-
  Verifica en el portal ONN Bolivia (onnbolivia.com) qué personas de una lista
  Excel están registradas en Lista de niños y genera un reporte PDF de quiénes
  están, quiénes requieren revisión y quiénes no están. Usar cuando el usuario
  pida verificar niños en ONN Bolivia o invoque /verificar-ninos-onn.
---

# Verificación de Lista de Niños en ONN Bolivia

Este flujo automatiza la consulta de solo lectura en `https://onnbolivia.com` para cruzar cualquier archivo `.xlsx` contra la tabla **Lista general de niños** y entregar un reporte PDF ordenado.

## Requisitos previos

1. El archivo `.env` en la raíz del proyecto debe definir:
   - `ONN_BASE_URL=https://onnbolivia.com`
   - `ONN_USERNAME`
   - `ONN_PASSWORD`
   (Usa `.env.example` como plantilla si aún no existe `.env`; nunca expongas la contraseña en logs ni reportes.)
2. Contar con el archivo Excel a validar (por ejemplo `yaque.xlsx`).

## Política estricta de solo lectura

- **Nunca** agregues, edites, elimines ni envíes solicitudes o formularios de niños, maestros o iglesias en el portal.
- Las únicas acciones permitidas son: iniciar sesión, navegar a **Solicitudes de iglesias → Lista de niños**, leer la tabla `/table_ninos_gral` o escribir en los buscadores de filtro de la tabla.
- Si aparece un CAPTCHA o desafío anti-bot, detén el flujo de inmediato e informa al usuario.

## Modo 1 (Principal): Ejecución automática con scripts Python

Ejecuta el script principal indicando la ruta del `.xlsx`:

```bash
python3 .agents/skills/verificar-ninos-onn/scripts/verificar_onn.py <archivo.xlsx>
```

Opciones útiles:
- `--out-pdf reportes/mi_reporte.pdf`: define la ruta del PDF de salida (por defecto `reportes/reporte_<nombre_excel>.pdf`).
- `--out-json reportes/resultados.json`: define la ruta del JSON intermedio.
- `--sheet "Hoja1"`: especifica el nombre de la hoja si el Excel tiene varias.
- `--no-pdf`: genera solo el JSON de resultados.
- `--save-catalog`: guarda también en el JSON el catálogo completo descargado del portal para auditoría.

### ¿Qué hace internamente el flujo?
1. [leer_excel.py](./scripts/leer_excel.py): lee el `.xlsx` con Python estándar (`zipfile` + XML), mapeando columnas por letra (soporta celdas vacías y filas dispersas) y extrayendo nombre, C.I. original, dígitos del C.I., edad y estado previo en Excel.
2. [verificar_onn.py](./scripts/verificar_onn.py): carga `.env`, inicia sesión en `/login`, pasa por `/iglesias` y `/ninos`, y consulta el endpoint DataTables `POST /table_ninos_gral` con `length=10000` (paginando si `recordsTotal` supera el lote) para obtener la lista completa sin omitir coincidencias por errores de tipeo o segundos nombres.
3. Cruza cada persona en tres estados:
   - **ESTÁ**: coincidencia por C.I. exacto, nombre exacto o tokens completos del nombre.
   - **REVISAR**: coincidencia parcial/fonética, C.I. cercana, o conflicto de C.I. con mismo nombre.
   - **NO ESTÁ**: sin coincidencias en la tabla del portal.
4. [generar_pdf.py](./scripts/generar_pdf.py): genera el HTML y lo renderiza a PDF usando Chromium headless (con respaldo LibreOffice).

## Modo 2 (Alternativo): Navegación visual con el subagente `browser`

Usa el subagente `browser` cuando se requiera inspección visual directa en el navegador o si el acceso HTTP directo falla:

1. Abrir `https://onnbolivia.com/login` (verificar ausencia de CAPTCHA).
2. Ingresar credenciales desde `.env` en `#usernameInput` (`input[name="username"]`) y `#passwordInput` (`input[name="password"]`), y enviar el formulario (`#loginsubmit`, botón `Iniciar sesión`).
3. Hacer clic en `Solicitudes de iglesias` en el menú lateral (`a[href="https://onnbolivia.com/iglesias"]`).
4. Hacer clic en el botón `Lista de niños` (`a[href="https://onnbolivia.com/ninos"]`).
5. Usar los buscadores ubicados debajo de cada cabecera en `table#tableIglesias`:
   - `NOMBRE DEL NIÑO` (columna 1): `input#ninos1`
   - `CARNET DE IDENTIDAD` (columna 2): `input#ninos2`
6. **Detalles importantes del buscador web**:
   - Los filtros disparan DataTables únicamente con el evento **`keyup`** (no con `input` o `change`). Para limpiar un filtro entre búsquedas, seleccionar todo el texto y presionar `Backspace` (no usar `.value = ""`).
   - El servidor busca con **subcadena contigua** (`LIKE '%texto%'`), insensible a mayúsculas, tildes y `ñ`. Si una persona en el portal tiene segundo nombre (ej. `SAYDA LINETH HUALLPA CAMPOS`), buscar `"Saida Huallpa"` devolverá 0 resultados; busca también por un apellido distintivo (ej. `"Huallpa"`) o cambia el selector `Mostrando X entradas` a `500` / `10,000`.
7. Compilar los resultados y generar el PDF con [generar_pdf.py](./scripts/generar_pdf.py).

## Verificación final

Antes de entregar el resultado al usuario:
1. Comprueba el resumen en consola (`ESTÁ`, `REVISAR`, `NO ESTÁ`, discrepancias con Excel).
2. Audita cualquier caso marcado en `REVISAR` o sin C.I. en el Excel.
3. Abre el PDF generado (`reportes/reporte_<nombre>.pdf`) con `read_file` para confirmar visualmente que todas las tablas, tildes/`ñ` y columnas se renderizan limpiamente sin cortes.
