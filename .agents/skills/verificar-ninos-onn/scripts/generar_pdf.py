#!/usr/bin/env python3
"""Genera un reporte PDF en español a partir del JSON producido por verificar_onn.py."""

from __future__ import annotations

import argparse
import html
import json
import os
import shutil
import subprocess
import tempfile
from pathlib import Path
from typing import Any


def _esc(val: Any) -> str:
    return html.escape(str(val if val is not None else ""), quote=True)


def _format_fecha(iso_str: str) -> str:
    if not iso_str:
        return "-"
    try:
        date_part, time_part = iso_str.split("T", 1)
        yyyy, mm, dd = date_part.split("-")
        return f"{dd}/{mm}/{yyyy} {time_part[:5]}"
    except Exception:
        return iso_str


def _build_html(data: dict[str, Any]) -> str:
    meta = data.get("metadata", {})
    resumen = data.get("resumen", {})
    resultados = data.get("resultados", [])

    no_estan = [r for r in resultados if r.get("estado") == "NO ESTÁ"]
    revisar = [r for r in resultados if r.get("estado") == "REVISAR"]
    estan = [r for r in resultados if r.get("estado") == "ESTÁ"]

    alertas: list[dict[str, Any]] = []
    for r in resultados:
        nota_cruce = r.get("nota_cruce", "")
        es_alerta_ci = "C.I. difiere" in nota_cruce or "nombre difiere bastante" in nota_cruce
        if r.get("discrepancia_excel") or es_alerta_ci:
            alertas.append(r)

    rows_no_html = ""
    for r in no_estan:
        ex_sist = _esc(r.get("en_sistema_excel") or "-")
        if r.get("discrepancia_excel"):
            ex_sist += "<br><span class='badge badge-warn'>Difiere</span>"
        rows_no_html += f"""
        <tr>
          <td class="num">{_esc(r.get("numero") or "-")}</td>
          <td><strong>{_esc(r.get("nombre"))}</strong></td>
          <td class="mono">{_esc(r.get("carnet_raw") or "(sin C.I.)")}</td>
          <td class="num">{_esc(r.get("edad_excel") or "-")}</td>
          <td>{ex_sist}</td>
          <td>{_esc(r.get("estado_invitacion") or "-")}</td>
          <td>{_esc(r.get("observaciones_excel") or "-")}</td>
        </tr>"""
    if not rows_no_html:
        rows_no_html = '<tr><td colspan="7" class="empty">Todas las personas del Excel fueron localizadas en el sistema.</td></tr>'

    sec_revisar_html = ""
    if revisar:
        rows_rev = ""
        for r in revisar:
            cands = r.get("candidatos_onn") or ([] if not r.get("match_onn") else [r["match_onn"]])
            cands_html = ""
            for c in cands:
                cands_html += (
                    f"<div class='cand-item'>"
                    f"<strong>{_esc(c.get('nombre'))}</strong> "
                    f"<span class='mono'>(C.I. {_esc(c.get('carnet') or '-')})</span> · "
                    f"{_esc(c.get('edad') or '-')} · "
                    f"{_esc(c.get('codigo_iglesia') or '')} {_esc(c.get('nombre_iglesia') or '')}"
                    f"</div>"
                )
            if not cands_html:
                cands_html = "-"
            rows_rev += f"""
            <tr>
              <td class="num">{_esc(r.get("numero") or "-")}</td>
              <td><strong>{_esc(r.get("nombre"))}</strong><br><span class="sub">Excel C.I.: {_esc(r.get("carnet_raw") or "(sin C.I.)")} | Edad: {_esc(r.get("edad_excel") or "-")} | Excel ¿En sist.?: {_esc(r.get("en_sistema_excel") or "-")}</span></td>
              <td>{cands_html}</td>
              <td>{_esc(r.get("nota_cruce"))}</td>
            </tr>"""
        sec_revisar_html = f"""
        <section class="section">
          <div class="section-header amber">
            <h2>2. Posibles coincidencias / Casos para revisar ({len(revisar)})</h2>
          </div>
          <table>
            <thead>
              <tr>
                <th style="width: 4%">N°</th>
                <th style="width: 30%">Registro en Excel</th>
                <th style="width: 36%">Candidato(s) en ONN Bolivia</th>
                <th style="width: 30%">Motivo de revisión</th>
              </tr>
            </thead>
            <tbody>{rows_rev}</tbody>
          </table>
        </section>"""

    rows_si_html = ""
    for r in estan:
        m = r.get("match_onn") or {}
        ig_txt = m.get("nombre_iglesia") or m.get("iglesia_texto") or "-"
        cod_txt = m.get("codigo_iglesia") or ""
        est_sol = m.get("estado_solicitud") or ""
        iglesia_cell = f"<strong>{_esc(cod_txt)}</strong> {_esc(ig_txt)}" if cod_txt else _esc(ig_txt)
        if est_sol:
            iglesia_cell += f"<br><span class='sub'>{_esc(est_sol)}</span>"
        disc_badge = ""
        if r.get("discrepancia_excel"):
            disc_badge = "<br><span class='badge badge-warn'>Difiere de Excel</span>"
        rows_si_html += f"""
        <tr>
          <td class="num">{_esc(r.get("numero") or "-")}</td>
          <td><strong>{_esc(r.get("nombre"))}</strong></td>
          <td class="mono">{_esc(r.get("carnet_raw") or "(sin C.I.)")}</td>
          <td>{_esc(m.get("nombre") or "-")}</td>
          <td class="mono">{_esc(m.get("carnet") or "-")}</td>
          <td>{_esc(m.get("edad") or "-")}</td>
          <td>{iglesia_cell}</td>
          <td class="mono">{_esc(m.get("created_at") or "-")}</td>
          <td><span class="badge badge-ok">{_esc(r.get("metodo_cruce"))}</span>{disc_badge}<div class="sub">{_esc(r.get("nota_cruce"))}</div></td>
        </tr>"""
    if not rows_si_html:
        rows_si_html = '<tr><td colspan="9" class="empty">No se encontraron personas registradas en el sistema.</td></tr>'

    alertas_html = ""
    if alertas:
        items_li = ""
        for a in alertas:
            detalle = a.get("nota_discrepancia") or a.get("nota_cruce")
            items_li += (
                f"<li><strong>N° {_esc(a.get('numero') or '-')} — {_esc(a.get('nombre'))}:</strong> "
                f"{_esc(detalle)}</li>"
            )
        alertas_html = f"""
        <div class="alert-box">
          <div class="alert-title">Observaciones y discrepancias respecto al Excel ({len(alertas)})</div>
          <ul>{items_li}</ul>
        </div>"""

    sec_si_num = "3" if revisar else "2"

    return f"""<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="utf-8">
  <title>Reporte de Verificación ONN Bolivia — {_esc(meta.get("archivo_excel", ""))}</title>
  <style>
    @page {{
      size: letter landscape;
      margin: 10mm 11mm;
    }}
    * {{
      box-sizing: border-box;
    }}
    body {{
      font-family: "DejaVu Sans", "Liberation Sans", Arial, sans-serif;
      font-size: 9.5px;
      line-height: 1.35;
      color: #1f2937;
      margin: 0;
      padding: 0;
      -webkit-print-color-adjust: exact;
      print-color-adjust: exact;
    }}
    .header {{
      border-bottom: 2px solid #0f766e;
      padding-bottom: 8px;
      margin-bottom: 10px;
      display: flex;
      justify-content: space-between;
      align-items: flex-end;
      gap: 12px;
    }}
    .header h1 {{
      font-size: 16px;
      margin: 0 0 3px 0;
      color: #0f766e;
    }}
    .header .subtitle {{
      font-size: 9.5px;
      color: #4b5563;
      margin: 0;
    }}
    .meta-right {{
      text-align: right;
      font-size: 8.5px;
      color: #374151;
    }}
    .kpis {{
      display: grid;
      grid-template-columns: repeat(4, 1fr);
      gap: 8px;
      margin-bottom: 10px;
    }}
    .kpi {{
      border: 1px solid #d1d5db;
      border-radius: 6px;
      padding: 7px 10px;
      background: #f9fafb;
    }}
    .kpi.green {{ border-left: 4px solid #15803d; background: #f0fdf4; }}
    .kpi.amber {{ border-left: 4px solid #b45309; background: #fffbeb; }}
    .kpi.red {{ border-left: 4px solid #b91c1c; background: #fef2f2; }}
    .kpi.slate {{ border-left: 4px solid #0f766e; background: #f0fdfa; }}
    .kpi-label {{
      font-size: 8px;
      text-transform: uppercase;
      letter-spacing: 0.4px;
      color: #4b5563;
      margin-bottom: 2px;
    }}
    .kpi-val {{
      font-size: 16px;
      font-weight: bold;
      color: #111827;
    }}
    .kpi-sub {{
      font-size: 8px;
      color: #4b5563;
    }}
    .alert-box {{
      border: 1px solid #f59e0b;
      background: #fffbeb;
      border-radius: 6px;
      padding: 7px 10px;
      margin-bottom: 12px;
    }}
    .alert-title {{
      font-weight: bold;
      color: #92400e;
      margin-bottom: 4px;
      font-size: 9.5px;
    }}
    .alert-box ul {{
      margin: 0;
      padding-left: 16px;
    }}
    .alert-box li {{
      margin-bottom: 2px;
    }}
    .section {{
      margin-bottom: 14px;
    }}
    .section-header {{
      padding: 5px 9px;
      border-radius: 5px 5px 0 0;
      color: #ffffff;
    }}
    .section-header h2 {{
      margin: 0;
      font-size: 11px;
      font-weight: bold;
    }}
    .section-header.red {{ background: #b91c1c; }}
    .section-header.amber {{ background: #b45309; }}
    .section-header.green {{ background: #15803d; }}
    table {{
      width: 100%;
      border-collapse: collapse;
      table-layout: fixed;
      border: 1px solid #d1d5db;
    }}
    thead {{
      display: table-header-group;
    }}
    tr {{
      break-inside: avoid;
      page-break-inside: avoid;
    }}
    th {{
      background: #f3f4f6;
      color: #111827;
      font-weight: bold;
      text-align: left;
      padding: 5px 6px;
      border-bottom: 1.5px solid #9ca3af;
      border-right: 1px solid #e5e7eb;
      font-size: 8.5px;
    }}
    td {{
      padding: 4px 6px;
      border-bottom: 1px solid #e5e7eb;
      border-right: 1px solid #f3f4f6;
      vertical-align: top;
      word-wrap: break-word;
      overflow-wrap: break-word;
    }}
    tbody tr:nth-child(even) {{
      background: #fafafa;
    }}
    .num {{
      text-align: center;
      font-variant-numeric: tabular-nums;
    }}
    .mono {{
      font-family: "DejaVu Sans Mono", monospace;
      font-size: 8.5px;
    }}
    .sub {{
      font-size: 8px;
      color: #4b5563;
      margin-top: 2px;
    }}
    .cand-item {{
      margin-bottom: 3px;
      padding-bottom: 3px;
      border-bottom: 1px dashed #e5e7eb;
    }}
    .cand-item:last-child {{
      margin-bottom: 0;
      padding-bottom: 0;
      border-bottom: none;
    }}
    .badge {{
      display: inline-block;
      padding: 1px 5px;
      border-radius: 3px;
      font-size: 7.5px;
      font-weight: bold;
    }}
    .badge-ok {{
      background: #dcfce7;
      color: #166534;
      border: 1px solid #86efac;
    }}
    .badge-warn {{
      background: #fef3c7;
      color: #92400e;
      border: 1px solid #fcd34d;
      margin-top: 2px;
    }}
    .empty {{
      text-align: center;
      color: #6b7280;
      padding: 10px;
      font-style: italic;
    }}
  </style>
</head>
<body>
  <header class="header">
    <div>
      <h1>Reporte de Verificación — Lista de Niños (ONN Bolivia)</h1>
      <p class="subtitle">
        Archivo verificado: <strong>{_esc(meta.get("archivo_excel", "-"))}</strong>
        (Hoja: {_esc(meta.get("hoja_excel", "-"))}) ·
        Portal: {_esc(meta.get("portal_url", "https://onnbolivia.com/ninos"))}
      </p>
    </div>
    <div class="meta-right">
      <div><strong>Fecha de reporte:</strong> {_esc(_format_fecha(meta.get("generado_en", "")))}</div>
      <div><strong>Cuenta:</strong> {_esc(meta.get("usuario_portal", "-"))} ({_esc(meta.get("cuenta_login", "-"))})</div>
      <div><strong>Registros en ONN:</strong> {_esc(meta.get("total_onn_portal", "-"))}</div>
    </div>
  </header>

  <section class="kpis">
    <div class="kpi slate">
      <div class="kpi-label">Total en lista Excel</div>
      <div class="kpi-val">{_esc(resumen.get("total_excel", 0))}</div>
      <div class="kpi-sub">Personas evaluadas</div>
    </div>
    <div class="kpi green">
      <div class="kpi-label">SÍ están en sistema</div>
      <div class="kpi-val">{_esc(resumen.get("estan", 0))}</div>
      <div class="kpi-sub">{_esc(resumen.get("pct_estan", 0))}% del total</div>
    </div>
    <div class="kpi amber">
      <div class="kpi-label">Por revisar</div>
      <div class="kpi-val">{_esc(resumen.get("revisar", 0))}</div>
      <div class="kpi-sub">{_esc(resumen.get("pct_revisar", 0))}% del total</div>
    </div>
    <div class="kpi red">
      <div class="kpi-label">NO están en sistema</div>
      <div class="kpi-val">{_esc(resumen.get("no_estan", 0))}</div>
      <div class="kpi-sub">{_esc(resumen.get("pct_no_estan", 0))}% del total</div>
    </div>
  </section>

  {alertas_html}

  <section class="section">
    <div class="section-header red">
      <h2>1. Personas que NO ESTÁN en el sistema ({len(no_estan)})</h2>
    </div>
    <table>
      <thead>
        <tr>
          <th style="width: 4%">N°</th>
          <th style="width: 28%">Nombre(s) y Apellido(s)</th>
          <th style="width: 13%">Carnet / C.I.</th>
          <th style="width: 7%">Edad</th>
          <th style="width: 11%">Excel ¿En sist.?</th>
          <th style="width: 12%">Estado Inv.</th>
          <th style="width: 25%">Observaciones (Excel)</th>
        </tr>
      </thead>
      <tbody>{rows_no_html}</tbody>
    </table>
  </section>

  {sec_revisar_html}

  <section class="section">
    <div class="section-header green">
      <h2>{sec_si_num}. Personas que SÍ ESTÁN en el sistema ({len(estan)})</h2>
    </div>
    <table>
      <thead>
        <tr>
          <th style="width: 3.5%">N°</th>
          <th style="width: 16%">Nombre (Excel)</th>
          <th style="width: 8.5%">C.I. (Excel)</th>
          <th style="width: 16%">Nombre (ONN)</th>
          <th style="width: 8.5%">C.I. (ONN)</th>
          <th style="width: 6.5%">Edad</th>
          <th style="width: 15%">Iglesia / Estado</th>
          <th style="width: 8%">Registro</th>
          <th style="width: 18%">Coincidencia / Nota</th>
        </tr>
      </thead>
      <tbody>{rows_si_html}</tbody>
    </table>
  </section>
</body>
</html>
"""


def _render_chromium(html_path: Path, pdf_path: Path, no_sandbox: bool) -> tuple[bool, str]:
    chromium_bin = shutil.which("chromium") or "/usr/bin/chromium"
    if not Path(chromium_bin).exists():
        return False, "No se encontró el binario de Chromium."

    with tempfile.TemporaryDirectory(prefix="chromium_pdf_") as tmp_dir:
        profile_dir = Path(tmp_dir) / "profile"
        profile_dir.mkdir(parents=True, exist_ok=True)
        env = os.environ.copy()
        env["HOME"] = tmp_dir
        env["XDG_CONFIG_HOME"] = tmp_dir
        env["XDG_CACHE_HOME"] = tmp_dir

        cmd = [
            chromium_bin,
            "--headless=new",
            "--disable-gpu",
            "--no-pdf-header-footer",
            f"--user-data-dir={profile_dir}",
            f"--print-to-pdf={pdf_path.resolve()}",
        ]
        if no_sandbox:
            cmd.extend(["--no-sandbox", "--disable-dev-shm-usage"])
        cmd.append(html_path.resolve().as_uri())

        proc = subprocess.run(cmd, capture_output=True, text=True, env=env, timeout=60)
        ok = proc.returncode == 0 and pdf_path.is_file() and pdf_path.stat().st_size > 1000
        msg = (proc.stderr or proc.stdout or "").strip()
        return ok, msg


def generar_pdf(
    json_path: str | Path,
    out_pdf_path: str | Path | None = None,
    keep_html: bool = False,
) -> Path:
    """Genera el reporte PDF desde un archivo JSON de resultados."""
    src = Path(json_path)
    if not src.is_file():
        raise FileNotFoundError(f"No se encontró el archivo JSON: {src}")

    data = json.loads(src.read_text(encoding="utf-8"))
    if out_pdf_path is None:
        out_pdf = src.parent / (src.stem.replace("resultados_", "reporte_") + ".pdf")
    else:
        out_pdf = Path(out_pdf_path)

    out_pdf.parent.mkdir(parents=True, exist_ok=True)
    html_content = _build_html(data)

    with tempfile.NamedTemporaryFile(
        mode="w",
        encoding="utf-8",
        suffix=".html",
        prefix="onn_report_",
        dir=str(out_pdf.parent),
        delete=False,
    ) as tmp:
        tmp.write(html_content)
        tmp_html = Path(tmp.name)

    try:
        ok, err1 = _render_chromium(tmp_html, out_pdf, no_sandbox=False)
        if not ok:
            ok, err2 = _render_chromium(tmp_html, out_pdf, no_sandbox=True)
            if not ok:
                lo_bin = shutil.which("libreoffice") or "/usr/bin/libreoffice"
                if Path(lo_bin).exists():
                    with tempfile.TemporaryDirectory(prefix="lo_pdf_") as lo_tmp:
                        lo_env = os.environ.copy()
                        lo_env["HOME"] = lo_tmp
                        lo_profile = (Path(lo_tmp) / "profile").resolve().as_uri()
                        proc = subprocess.run(
                            [
                                lo_bin,
                                "--headless",
                                f"-env:UserInstallation={lo_profile}",
                                "--convert-to",
                                "pdf",
                                "--outdir",
                                str(out_pdf.parent),
                                str(tmp_html),
                            ],
                            capture_output=True,
                            text=True,
                            env=lo_env,
                            timeout=60,
                        )
                        generated = out_pdf.parent / f"{tmp_html.stem}.pdf"
                        if proc.returncode == 0 and generated.is_file():
                            generated.replace(out_pdf)
                            ok = True
                if not ok:
                    raise RuntimeError(
                        f"No se pudo renderizar el PDF con Chromium ni LibreOffice. Detalle: {err1} | {err2}"
                    )
    finally:
        if not keep_html and tmp_html.exists():
            tmp_html.unlink()

    return out_pdf


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Genera un reporte PDF desde el JSON de verificación de ONN Bolivia."
    )
    parser.add_argument("json_file", help="Ruta al archivo JSON generado por verificar_onn.py")
    parser.add_argument("--out-pdf", default=None, help="Ruta de salida para el archivo PDF")
    parser.add_argument(
        "--keep-html",
        action="store_true",
        help="Conserva el archivo HTML intermedio junto al PDF",
    )
    args = parser.parse_args()

    pdf = generar_pdf(args.json_file, args.out_pdf, keep_html=args.keep_html)
    print(f"PDF generado en: {pdf} ({pdf.stat().st_size} bytes)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
