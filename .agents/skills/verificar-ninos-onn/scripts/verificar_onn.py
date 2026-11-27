#!/usr/bin/env python3
"""Flujo de verificación de niños en ONN Bolivia (onnbolivia.com/ninos) contra un Excel."""

from __future__ import annotations

import argparse
import html as html_lib
import json
import os
import re
import sys
import unicodedata
from datetime import datetime
from difflib import SequenceMatcher
from html.parser import HTMLParser
from pathlib import Path
from typing import Any

import requests

SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

from generar_pdf import generar_pdf
from leer_excel import leer_personas_excel

SPANISH_PARTICLES = {"de", "del", "la", "las", "los", "y", "e", "da", "do"}


class _HTMLTextExtractor(HTMLParser):
    """Extrae texto plano de fragmentos HTML (maneja atributos con tags anidados)."""

    def __init__(self) -> None:
        super().__init__()
        self._chunks: list[str] = []

    def handle_data(self, data: str) -> None:
        self._chunks.append(data)

    def get_text(self) -> str:
        return re.sub(r"\s+", " ", html_lib.unescape("".join(self._chunks))).strip()


def _strip_html(raw_html: str) -> str:
    """Convierte HTML a texto limpio sin perder contenido por comillas o tags en atributos."""
    raw = str(raw_html or "")
    if "<" not in raw:
        return re.sub(r"\s+", " ", html_lib.unescape(raw)).strip()
    extractor = _HTMLTextExtractor()
    extractor.feed(raw)
    extractor.close()
    return extractor.get_text()


class _PageInspector(HTMLParser):
    """Extrae meta csrf-token, inputs _token y valores por defecto de selects."""

    def __init__(self) -> None:
        super().__init__()
        self.meta_csrf: str | None = None
        self.hidden_tokens: list[str] = []
        self.selects: dict[str, dict[str, Any]] = {}
        self._cur_select: str | None = None
        self._cur_option_val: str | None = None
        self._cur_option_selected = False
        self._cur_option_text: list[str] = []

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        attr_map = {k.lower(): (v or "") for k, v in attrs}
        t = tag.lower()
        if t == "meta" and attr_map.get("name", "").lower() == "csrf-token":
            if attr_map.get("content"):
                self.meta_csrf = attr_map["content"]
        elif t == "input" and attr_map.get("name") == "_token":
            if attr_map.get("value"):
                self.hidden_tokens.append(attr_map["value"])
        elif t == "select":
            key = attr_map.get("name") or attr_map.get("id") or ""
            if key:
                self._cur_select = key
                self.selects[key] = {"options": [], "selected": None}
        elif t == "option" and self._cur_select is not None:
            self._cur_option_val = attr_map.get("value")
            self._cur_option_selected = any(k.lower() == "selected" for k, _ in attrs)
            self._cur_option_text = []

    def handle_data(self, data: str) -> None:
        if self._cur_select is not None and self._cur_option_val is not None:
            self._cur_option_text.append(data)

    def handle_endtag(self, tag: str) -> None:
        t = tag.lower()
        if t == "option" and self._cur_select is not None:
            text = re.sub(r"\s+", " ", "".join(self._cur_option_text)).strip()
            val = text if self._cur_option_val is None else self._cur_option_val
            entry = self.selects[self._cur_select]
            entry["options"].append((val, text))
            if self._cur_option_selected:
                entry["selected"] = val
            self._cur_option_val = None
            self._cur_option_selected = False
            self._cur_option_text = []
        elif t == "select":
            self._cur_select = None

    def default_select_value(self, key: str, fallback: str = "") -> str:
        info = self.selects.get(key)
        if not info:
            return fallback
        if info["selected"] is not None:
            return str(info["selected"])
        if info["options"]:
            return str(info["options"][0][0])
        return fallback


def _find_env_file(explicit_env: str | None) -> Path | None:
    """Busca el archivo .env en la ruta indicada, CWD o ancestros del script."""
    if explicit_env:
        p = Path(explicit_env)
        if not p.is_file():
            raise FileNotFoundError(f"No se encontró el archivo .env en: {p}")
        return p

    candidates = [Path.cwd() / ".env"]
    for parent in SCRIPT_DIR.parents:
        candidates.append(parent / ".env")
    for cand in candidates:
        if cand.is_file():
            return cand
    return None


def load_credentials(explicit_env: str | None = None) -> tuple[str, str, str]:
    """Carga ONN_BASE_URL, ONN_USERNAME y ONN_PASSWORD desde .env con respaldo de os.environ."""
    env_vars: dict[str, str] = {}
    env_path = _find_env_file(explicit_env)
    if env_path is not None:
        for raw_line in env_path.read_text(encoding="utf-8").splitlines():
            line = raw_line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, val = line.split("=", 1)
            key = key.strip()
            val = val.strip()
            if len(val) >= 2 and ((val[0] == val[-1] == '"') or (val[0] == val[-1] == "'")):
                val = val[1:-1]
            env_vars[key] = val

    base_url = (
        env_vars.get("ONN_BASE_URL")
        or os.environ.get("ONN_BASE_URL")
        or "https://onnbolivia.com"
    ).rstrip("/")
    username = env_vars.get("ONN_USERNAME") or os.environ.get("ONN_USERNAME") or ""
    password = env_vars.get("ONN_PASSWORD") or os.environ.get("ONN_PASSWORD") or ""

    if not username or not password:
        raise RuntimeError(
            "Faltan ONN_USERNAME o ONN_PASSWORD. Configura el archivo .env (puedes usar .env.example como base)."
        )
    return base_url, username, password


def _has_captcha(page_html: str) -> bool:
    """Detecta presencia de CAPTCHA o desafío anti-bot explícito."""
    lower = page_html.lower()
    markers = ("g-recaptcha", "recaptcha/api", "hcaptcha.com", "cf-turnstile", "challenges.cloudflare")
    return any(m in lower for m in markers)


def _extract_user_banner(page_html: str) -> str:
    """Extrae el nombre del usuario autenticado mostrado en el portal si está disponible."""
    patterns = [
        r"BIENVENIDO\s*</[^>]+>\s*<[^>]+>\s*([^<]{3,80})<",
        r"Solicitante de iglesias.*?<[^>]+>([^<]{3,80})<",
        r'class="[^"]*user-name[^"]*"[^>]*>\s*([^<]{3,80})<',
    ]
    for pat in patterns:
        m = re.search(pat, page_html, flags=re.IGNORECASE | re.DOTALL)
        if m:
            txt = re.sub(r"\s+", " ", html_lib.unescape(m.group(1))).strip()
            if txt:
                return txt
    return ""


def _parse_iglesia_html(raw_html: str) -> dict[str, str]:
    """Extrae código, nombre de iglesia y estado de solicitud del HTML de la columna iglesia."""
    raw = raw_html or ""
    code_m = re.search(r"<b>\s*([^<]+?)\s*</b>", raw, flags=re.IGNORECASE)
    name_m = re.search(r"<i>\s*([^<]+?)\s*</i>", raw, flags=re.IGNORECASE)
    plain = re.sub(r"<[^>]+>", " | ", html_lib.unescape(raw))
    parts = [re.sub(r"\s+", " ", p).strip() for p in plain.split("|")]
    parts = [p for p in parts if p]

    codigo = html_lib.unescape(code_m.group(1)).strip() if code_m else (parts[0] if parts else "")
    iglesia = html_lib.unescape(name_m.group(1)).strip() if name_m else (parts[1] if len(parts) > 1 else "")
    estado = ""
    for p in reversed(parts):
        if p != codigo and p != iglesia:
            estado = p
            break

    return {
        "codigo_iglesia": codigo,
        "nombre_iglesia": iglesia,
        "estado_solicitud": estado,
        "iglesia_texto": " - ".join(p for p in (codigo, iglesia, estado) if p),
    }


def fetch_onn_children(base_url: str, username: str, password: str) -> dict[str, Any]:
    """Inicia sesión en ONN Bolivia (solo lectura) y descarga la tabla completa de niños."""
    session = requests.Session()
    session.headers.update(
        {
            "User-Agent": (
                "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
                "(KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36"
            ),
            "Accept-Language": "es-BO,es;q=0.9,en;q=0.8",
        }
    )

    login_url = f"{base_url}/login"
    r_login = session.get(login_url, timeout=30)
    r_login.raise_for_status()
    if _has_captcha(r_login.text):
        raise RuntimeError(
            "El portal muestra un CAPTCHA o desafío anti-bot en /login. Proceso detenido."
        )

    inspector = _PageInspector()
    inspector.feed(r_login.text)
    login_token = (
        inspector.hidden_tokens[0]
        if inspector.hidden_tokens
        else (inspector.meta_csrf or "")
    )
    if not login_token:
        raise RuntimeError("No se pudo extraer el token CSRF (_token) de /login.")

    r_post = session.post(
        login_url,
        data={
            "_token": login_token,
            "username": username,
            "password": password,
        },
        headers={"Referer": login_url, "Origin": base_url},
        allow_redirects=True,
        timeout=30,
    )
    r_post.raise_for_status()
    if _has_captcha(r_post.text):
        raise RuntimeError("El portal presentó un CAPTCHA tras enviar el login.")
    if "/login" in r_post.url.rstrip("/") or 'id="loginsubmit"' in r_post.text:
        raise RuntimeError(
            "Error de autenticación en ONN Bolivia: verifica las credenciales en .env."
        )

    usuario_portal = _extract_user_banner(r_post.text)

    r_iglesias = session.get(f"{base_url}/iglesias", timeout=30)
    r_iglesias.raise_for_status()
    if "/login" in r_iglesias.url:
        raise RuntimeError("La sesión expiró al visitar /iglesias.")

    ninos_url = f"{base_url}/ninos"
    r_ninos = session.get(ninos_url, timeout=30)
    r_ninos.raise_for_status()
    if "/login" in r_ninos.url:
        raise RuntimeError("La sesión expiró al visitar /ninos.")

    ninos_inspector = _PageInspector()
    ninos_inspector.feed(r_ninos.text)
    csrf_ninos = ninos_inspector.meta_csrf or (
        ninos_inspector.hidden_tokens[0] if ninos_inspector.hidden_tokens else login_token
    )
    if not csrf_ninos:
        raise RuntimeError("No se encontró el token CSRF en /ninos.")

    state_val = ninos_inspector.default_select_value("state", "")
    region_val = ninos_inspector.default_select_value("region", "")
    graduados_val = ninos_inspector.default_select_value("graduados", "")

    columns = [
        "iglesia",
        "nombre",
        "carnet",
        "edad",
        "genero",
        "created_at",
        "rango",
        "region",
    ]

    endpoint = f"{base_url}/table_ninos_gral"
    all_rows: list[dict[str, Any]] = []
    start = 0
    draw = 1
    page_len = 10000
    records_total = None
    records_filtered = None

    while True:
        payload: dict[str, str] = {
            "_token": csrf_ninos,
            "draw": str(draw),
            "start": str(start),
            "length": str(page_len),
            "search[value]": "",
            "search[regex]": "false",
            "state": state_val,
            "region": region_val,
            "graduados": graduados_val,
        }
        for idx, col_name in enumerate(columns):
            payload[f"columns[{idx}][data]"] = col_name
            payload[f"columns[{idx}][name]"] = ""
            payload[f"columns[{idx}][searchable]"] = "true"
            payload[f"columns[{idx}][orderable]"] = "false"
            payload[f"columns[{idx}][search][value]"] = ""
            payload[f"columns[{idx}][search][regex]"] = "false"

        r_dt = session.post(
            endpoint,
            data=payload,
            headers={
                "X-Requested-With": "XMLHttpRequest",
                "X-CSRF-TOKEN": csrf_ninos,
                "Referer": ninos_url,
                "Origin": base_url,
                "Accept": "application/json, text/javascript, */*; q=0.01",
            },
            timeout=60,
        )
        r_dt.raise_for_status()
        try:
            dt_json = json.loads(r_dt.text)
        except json.JSONDecodeError as exc:
            snippet = r_dt.text[:250].replace("\n", " ")
            raise RuntimeError(
                f"Respuesta inesperada de /table_ninos_gral (no es JSON): {snippet}"
            ) from exc

        records_total = int(dt_json.get("recordsTotal", 0))
        records_filtered = int(dt_json.get("recordsFiltered", records_total))
        batch = dt_json.get("data", [])
        if not isinstance(batch, list):
            raise RuntimeError("El campo 'data' de /table_ninos_gral no es una lista.")

        all_rows.extend(batch)
        if not batch or len(all_rows) >= records_filtered or len(all_rows) >= records_total:
            break
        start += len(batch)
        draw += 1

    registros: list[dict[str, Any]] = []
    for row in all_rows:
        ig_info = _parse_iglesia_html(str(row.get("iglesia", "")))
        nombre = _strip_html(str(row.get("nombre", "")))
        carnet = _strip_html(str(row.get("carnet", "")))
        edad = _strip_html(str(row.get("edad", "")))
        genero = _strip_html(str(row.get("genero", "")))
        created_at = _strip_html(str(row.get("created_at", "")))
        rango = _strip_html(str(row.get("rango", "")))
        region = _strip_html(str(row.get("region", "")))

        registros.append(
            {
                "nombre": nombre,
                "carnet": carnet,
                "carnet_digits": re.sub(r"\D+", "", carnet),
                "edad": edad,
                "genero": genero,
                "created_at": created_at,
                "rango": rango,
                "region": region,
                **ig_info,
            }
        )

    return {
        "usuario_portal": usuario_portal or username,
        "username": username,
        "records_total": records_total or len(registros),
        "records_filtered": records_filtered or len(registros),
        "filtros_default": {
            "state": state_val,
            "region": region_val,
            "graduados": graduados_val,
        },
        "registros": registros,
    }


def norm_text(text: str) -> str:
    """Normaliza texto: sin tildes/diacríticos (ñ->n), minúsculas y alfanumérico."""
    nfd = unicodedata.normalize("NFD", text or "")
    sin_tildes = "".join(ch for ch in nfd if unicodedata.category(ch) != "Mn")
    return re.sub(r"[^a-z0-9]+", " ", sin_tildes.lower()).strip()


def norm_ci(text: str) -> str:
    """Extrae solo dígitos de un C.I. y elimina ceros a la izquierda si aplica."""
    digits = re.sub(r"\D+", "", str(text or ""))
    if not digits:
        return ""
    return digits.lstrip("0") or digits


def phonetic_es(text: str) -> str:
    """Normalización fonética española suave para detectar posibles coincidencias."""
    base = norm_text(text)
    if not base:
        return ""
    tokens_out: list[str] = []
    for tok in base.split():
        t = tok
        t = t.replace("ch", "\x01").replace("sh", "\x02")
        t = t.replace("th", "t").replace("ph", "f").replace("ck", "k")
        t = re.sub(r"qu([ei])", r"k\1", t)
        t = t.replace("q", "k")
        t = re.sub(r"gu([ei])", r"g\1", t)
        t = re.sub(r"g([ei])", r"j\1", t)
        t = re.sub(r"c([ei])", r"s\1", t)
        t = t.replace("c", "k")
        t = t.replace("z", "s").replace("v", "b").replace("ll", "i").replace("y", "i")
        t = t.replace("h", "")
        t = t.replace("\x01", "ch").replace("\x02", "sh")
        t = re.sub(r"(.)\1+", r"\1", t)
        if t:
            tokens_out.append(t)
    return " ".join(tokens_out)


def _levenshtein_max1(a: str, b: str) -> bool:
    """Devuelve True si la distancia de edición entre a y b es <= 1."""
    if a == b:
        return True
    la, lb = len(a), len(b)
    if abs(la - lb) > 1:
        return False
    if la == lb:
        return sum(ca != cb for ca, cb in zip(a, b)) <= 1
    if la > lb:
        a, b = b, a
        la, lb = lb, la
    i = j = diffs = 0
    while i < la and j < lb:
        if a[i] == b[j]:
            i += 1
            j += 1
        else:
            diffs += 1
            if diffs > 1:
                return False
            j += 1
    return True


def _extract_int_age(val: str) -> int | None:
    m = re.search(r"\d+", str(val or ""))
    return int(m.group(0)) if m else None


def _tokens_meaningful(tokens: list[str]) -> list[str]:
    return [t for t in tokens if t not in SPANISH_PARTICLES and len(t) >= 2]


def _all_tokens_present(needles: list[str], haystack: list[str]) -> bool:
    """Comprueba si todos los tokens de needles aparecen en haystack (multiconjunto)."""
    if not needles:
        return False
    pool = list(haystack)
    for tok in needles:
        if tok in pool:
            pool.remove(tok)
        else:
            return False
    return True


def _format_onn_match(rec: dict[str, Any]) -> dict[str, Any]:
    return {
        "nombre": rec.get("nombre", ""),
        "carnet": rec.get("carnet", ""),
        "edad": rec.get("edad", ""),
        "genero": rec.get("genero", ""),
        "created_at": rec.get("created_at", ""),
        "rango": rec.get("rango", ""),
        "region": rec.get("region", ""),
        "codigo_iglesia": rec.get("codigo_iglesia", ""),
        "nombre_iglesia": rec.get("nombre_iglesia", ""),
        "estado_solicitud": rec.get("estado_solicitud", ""),
        "iglesia_texto": rec.get("iglesia_texto", ""),
    }


def _norm_excel_en_sistema(val: str) -> str:
    n = norm_text(val)
    if n in ("si", "s", "yes", "1", "true"):
        return "SI"
    if n in ("no", "n", "0", "false"):
        return "NO"
    return ""


def _evaluate_discrepancy(en_sistema_excel: str, estado: str) -> tuple[bool, str]:
    ex = _norm_excel_en_sistema(en_sistema_excel)
    if ex == "SI" and estado == "NO ESTÁ":
        return True, "Excel indica 'SÍ', pero NO aparece en ONN"
    if ex == "NO" and estado == "ESTÁ":
        return True, "Excel indica 'NO', pero SÍ está registrado en ONN"
    if ex == "SI" and estado == "REVISAR":
        return True, "Excel indica 'SÍ', pero en ONN solo hay coincidencia parcial/dudosa"
    if ex == "NO" and estado == "REVISAR":
        return True, "Excel indica 'NO', pero existe un posible candidato en ONN"
    return False, ""


def cross_reference_lists(
    personas_excel: list[dict[str, Any]],
    registros_onn: list[dict[str, Any]],
) -> dict[str, Any]:
    """Cruza los niños del Excel contra el catálogo completo descargado de ONN."""
    indexed_onn: list[dict[str, Any]] = []
    by_ci: dict[str, list[dict[str, Any]]] = {}
    by_name: dict[str, list[dict[str, Any]]] = {}

    for rec in registros_onn:
        n_txt = norm_text(rec.get("nombre", ""))
        c_norm = norm_ci(rec.get("carnet", ""))
        ph_txt = phonetic_es(rec.get("nombre", ""))
        item = {
            "raw": rec,
            "norm_name": n_txt,
            "tokens": n_txt.split(),
            "ph_name": ph_txt,
            "ph_tokens": ph_txt.split(),
            "ci_norm": c_norm,
            "edad_int": _extract_int_age(rec.get("edad", "")),
        }
        indexed_onn.append(item)
        if len(c_norm) >= 5:
            by_ci.setdefault(c_norm, []).append(item)
        if n_txt:
            by_name.setdefault(n_txt, []).append(item)

    resultados: list[dict[str, Any]] = []

    for persona in personas_excel:
        ex_name_raw = persona.get("nombre", "")
        ex_name_norm = norm_text(ex_name_raw)
        ex_tokens = ex_name_norm.split()
        ex_tokens_sig = _tokens_meaningful(ex_tokens)
        ex_ph_name = phonetic_es(ex_name_raw)
        ex_ph_tokens = ex_ph_name.split()
        ex_ci = norm_ci(persona.get("carnet_digits") or persona.get("carnet_raw") or "")
        ex_edad_int = _extract_int_age(persona.get("edad_excel", ""))

        estado = "NO ESTÁ"
        metodo = "Sin coincidencia"
        nota_cruce = "No se encontró coincidencia por C.I. ni por nombre en la tabla de ONN."
        match_principal: dict[str, Any] | None = None
        candidatos: list[dict[str, Any]] = []

        # 1. C.I. exacto
        if len(ex_ci) >= 5 and ex_ci in by_ci:
            matches = by_ci[ex_ci]
            best = max(
                matches,
                key=lambda m: SequenceMatcher(None, ex_name_norm, m["norm_name"]).ratio(),
            )
            estado = "ESTÁ"
            metodo = "C.I. exacto"
            match_principal = _format_onn_match(best["raw"])
            sim = SequenceMatcher(None, ex_name_norm, best["norm_name"]).ratio()
            if ex_name_norm == best["norm_name"]:
                nota_cruce = "Coincide por C.I. exacto y nombre exacto."
            elif sim >= 0.72:
                nota_cruce = (
                    f"Coincide por C.I. exacto (variación en nombre: "
                    f"ONN '{best['raw']['nombre']}')."
                )
            else:
                nota_cruce = (
                    f"Coincide por C.I. exacto, pero el nombre difiere bastante "
                    f"(Excel: '{ex_name_raw}' vs ONN: '{best['raw']['nombre']}')."
                )

        # 2. Nombre normalizado exacto
        elif ex_name_norm and ex_name_norm in by_name:
            matches = by_name[ex_name_norm]
            best = matches[0]
            if ex_edad_int is not None:
                for m in matches:
                    if m["edad_int"] == ex_edad_int:
                        best = m
                        break
            estado = "ESTÁ"
            metodo = "Nombre exacto"
            match_principal = _format_onn_match(best["raw"])
            onn_ci = best["ci_norm"]
            if ex_ci and onn_ci and ex_ci != onn_ci:
                nota_cruce = (
                    f"Coincide por nombre exacto; C.I. difiere "
                    f"(Excel: {persona.get('carnet_raw') or ex_ci} vs ONN: {best['raw']['carnet']})."
                )
            elif not ex_ci and onn_ci:
                nota_cruce = (
                    f"Coincide por nombre exacto (sin C.I. en Excel; ONN: {best['raw']['carnet']})."
                )
            else:
                nota_cruce = "Coincide por nombre exacto."

        else:
            # 3. Todos los tokens exactos del nombre en ONN
            token_matches: list[dict[str, Any]] = []
            if len(ex_tokens_sig) >= 2:
                for item in indexed_onn:
                    if _all_tokens_present(ex_tokens_sig, item["tokens"]):
                        token_matches.append(item)

            if token_matches:
                best = max(
                    token_matches,
                    key=lambda m: (
                        1 if (ex_edad_int is not None and m["edad_int"] == ex_edad_int) else 0,
                        SequenceMatcher(None, ex_name_norm, m["norm_name"]).ratio(),
                    ),
                )
                onn_ci = best["ci_norm"]
                ci_conflict = bool(ex_ci and onn_ci and ex_ci != onn_ci)
                if len(ex_tokens_sig) >= 3 and not ci_conflict and len(token_matches) == 1:
                    estado = "ESTÁ"
                    metodo = "Nombre completo (tokens)"
                    match_principal = _format_onn_match(best["raw"])
                    if not ex_ci and onn_ci:
                        nota_cruce = (
                            f"Coinciden todos los tokens del nombre (sin C.I. en Excel; ONN C.I.: {best['raw']['carnet']})."
                        )
                    else:
                        nota_cruce = "Coinciden todos los tokens del nombre en ONN."
                else:
                    estado = "REVISAR"
                    metodo = "Tokens exactos (verificar)"
                    candidatos = [_format_onn_match(m["raw"]) for m in token_matches[:3]]
                    match_principal = candidatos[0]
                    if ci_conflict:
                        nota_cruce = (
                            f"Coincide por tokens del nombre, pero C.I. difiere "
                            f"(Excel: {persona.get('carnet_raw') or ex_ci} vs ONN: {best['raw']['carnet']})."
                        )
                    elif len(token_matches) > 1:
                        nota_cruce = (
                            f"Coinciden los tokens del nombre con {len(token_matches)} registros en ONN; verificar cuál corresponde."
                        )
                    else:
                        nota_cruce = (
                            f"Solo {len(ex_tokens_sig)} tokens en Excel coinciden con '{best['raw']['nombre']}'; confirmar identidad."
                        )

            # 4. Candidatos aproximados / fonéticos / C.I. cercana
            else:
                scored: list[tuple[float, str, dict[str, Any]]] = []
                ex_ph_sig = [t for t in ex_ph_tokens if len(t) >= 2]
                ex_first_ph = ex_ph_sig[0] if ex_ph_sig else ""
                ex_surname_cands = []
                if len(ex_ph_sig) >= 2:
                    ex_surname_cands.append(ex_ph_sig[-1])
                if len(ex_ph_sig) >= 3:
                    ex_surname_cands.append(ex_ph_sig[-2])

                for item in indexed_onn:
                    reasons: list[str] = []
                    score = 0.0

                    sim_norm = SequenceMatcher(None, ex_name_norm, item["norm_name"]).ratio()
                    sim_ph = SequenceMatcher(None, ex_ph_name, item["ph_name"]).ratio()
                    onn_ph = item["ph_tokens"]
                    shared_ph = set(ex_ph_sig) & set(onn_ph)

                    # C.I. cercana (1 dígito de error o subcadena >=7 dígitos),
                    # exigiendo compatibilidad mínima de nombre para evitar colisiones secuenciales.
                    onn_ci = item["ci_norm"]
                    ci_close = False
                    has_min_name_compat = bool(shared_ph) or sim_ph >= 0.45
                    if len(ex_ci) >= 6 and len(onn_ci) >= 6 and has_min_name_compat:
                        if _levenshtein_max1(ex_ci, onn_ci):
                            ci_close = True
                            score = max(score, 0.92)
                            reasons.append(f"C.I. muy similar (ONN: {item['raw']['carnet']})")
                        elif min(len(ex_ci), len(onn_ci)) >= 7 and (
                            ex_ci in onn_ci or onn_ci in ex_ci
                        ):
                            ci_close = True
                            score = max(score, 0.88)
                            reasons.append(f"C.I. contenida (ONN: {item['raw']['carnet']})")

                    # Nombre fonético / parcial (apellido en tokens[1:] + primer nombre)
                    onn_tail_ph = onn_ph[1:] if len(onn_ph) > 1 else onn_ph
                    surname_hit = any(
                        sc in onn_tail_ph
                        or any(SequenceMatcher(None, sc, ot).ratio() >= 0.88 for ot in onn_tail_ph)
                        for sc in ex_surname_cands
                    )
                    first_hit = bool(ex_first_ph) and (
                        ex_first_ph in onn_ph[:2]
                        or any(
                            SequenceMatcher(None, ex_first_ph, ot).ratio() >= 0.82
                            for ot in onn_ph[:2]
                        )
                    )
                    both_have_distinct_ci = bool(ex_ci and onn_ci and ex_ci != onn_ci and not ci_close)
                    if surname_hit and first_hit:
                        # Si ambos tienen C.I. válidos y distintos, exigir al menos 2 tokens o alta similitud
                        if not both_have_distinct_ci or len(shared_ph) >= 2 or sim_ph >= 0.82:
                            s_val = max(0.82, 0.75 + 0.25 * sim_ph)
                            score = max(score, s_val)
                            reasons.append("nombre y apellido coincidentes o fonéticos")

                    # Similitud global de nombre
                    if sim_norm >= 0.80 or sim_ph >= 0.82:
                        if not both_have_distinct_ci or sim_norm >= 0.86 or sim_ph >= 0.88:
                            score = max(score, max(sim_norm, sim_ph))
                            reasons.append(f"nombre similar ({max(sim_norm, sim_ph):.0%})")

                    if score > 0.0 or ci_close:
                        if ex_edad_int is not None and item["edad_int"] == ex_edad_int:
                            score += 0.05
                            reasons.append("misma edad")
                        scored.append((score, ", ".join(dict.fromkeys(reasons)), item))

                if scored:
                    scored.sort(key=lambda x: x[0], reverse=True)
                    top = scored[:3]
                    estado = "REVISAR"
                    metodo = "Posible coincidencia"
                    candidatos = [_format_onn_match(it[2]["raw"]) for it in top]
                    match_principal = candidatos[0]
                    best_score, best_reason, best_item = top[0]
                    ci_note = ""
                    if ex_ci and best_item["ci_norm"] and ex_ci != best_item["ci_norm"]:
                        ci_note = f"; C.I. difiere (Excel: {persona.get('carnet_raw')} vs ONN: {best_item['raw']['carnet']})"
                    elif not ex_ci:
                        ci_note = " (sin C.I. en Excel)"
                    nota_cruce = (
                        f"Posible coincidencia con '{best_item['raw']['nombre']}' "
                        f"({best_reason}{ci_note})."
                    )

        discrepancia, nota_disc = _evaluate_discrepancy(
            persona.get("en_sistema_excel", ""), estado
        )

        resultados.append(
            {
                **persona,
                "estado": estado,
                "metodo_cruce": metodo,
                "nota_cruce": nota_cruce,
                "discrepancia_excel": discrepancia,
                "nota_discrepancia": nota_disc,
                "match_onn": match_principal,
                "candidatos_onn": candidatos,
            }
        )

    total = len(resultados)
    si_estan = sum(1 for r in resultados if r["estado"] == "ESTÁ")
    revisar = sum(1 for r in resultados if r["estado"] == "REVISAR")
    no_estan = sum(1 for r in resultados if r["estado"] == "NO ESTÁ")
    discrepancias = sum(1 for r in resultados if r["discrepancia_excel"])

    return {
        "resumen": {
            "total_excel": total,
            "estan": si_estan,
            "revisar": revisar,
            "no_estan": no_estan,
            "discrepancias_excel": discrepancias,
            "pct_estan": round((si_estan / total * 100), 1) if total else 0.0,
            "pct_revisar": round((revisar / total * 100), 1) if total else 0.0,
            "pct_no_estan": round((no_estan / total * 100), 1) if total else 0.0,
        },
        "resultados": resultados,
    }


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Verifica una lista Excel de niños contra el portal ONN Bolivia y genera reporte PDF."
    )
    parser.add_argument("excel", help="Ruta al archivo .xlsx a verificar")
    parser.add_argument("--env", default=None, help="Ruta opcional al archivo .env")
    parser.add_argument("--sheet", default=None, help="Nombre opcional de la hoja Excel")
    parser.add_argument("--out-json", default=None, help="Ruta de salida para el JSON de resultados")
    parser.add_argument("--out-pdf", default=None, help="Ruta de salida para el reporte PDF")
    parser.add_argument("--no-pdf", action="store_true", help="Solo guardar JSON, sin generar PDF")
    parser.add_argument(
        "--save-catalog",
        action="store_true",
        help="Incluye los registros completos descargados de ONN dentro del JSON",
    )
    args = parser.parse_args()

    excel_path = Path(args.excel)
    stem = excel_path.stem
    reportes_dir = excel_path.resolve().parent / "reportes"
    out_json = Path(args.out_json) if args.out_json else reportes_dir / f"resultados_{stem}.json"
    out_pdf = Path(args.out_pdf) if args.out_pdf else reportes_dir / f"reporte_{stem}.pdf"

    excel_data = leer_personas_excel(excel_path, sheet_name=args.sheet)
    base_url, username, password = load_credentials(args.env)
    onn_data = fetch_onn_children(base_url, username, password)
    cruce = cross_reference_lists(excel_data["personas"], onn_data["registros"])

    payload: dict[str, Any] = {
        "metadata": {
            "generado_en": datetime.now().isoformat(timespec="seconds"),
            "archivo_excel": excel_data["nombre_archivo"],
            "ruta_excel": excel_data["archivo"],
            "hoja_excel": excel_data["hoja"],
            "portal_url": f"{base_url}/ninos",
            "usuario_portal": onn_data["usuario_portal"],
            "cuenta_login": onn_data["username"],
            "total_onn_portal": onn_data["records_total"],
            "total_onn_descargados": len(onn_data["registros"]),
            "filtros_default_portal": onn_data["filtros_default"],
        },
        "resumen": cruce["resumen"],
        "resultados": cruce["resultados"],
    }
    if args.save_catalog:
        payload["catalogo_onn"] = onn_data["registros"]

    out_json.parent.mkdir(parents=True, exist_ok=True)
    out_json.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

    res = cruce["resumen"]
    print(
        f"Verificación completada ({excel_data['nombre_archivo']}): "
        f"Total={res['total_excel']} | ESTÁ={res['estan']} | "
        f"REVISAR={res['revisar']} | NO ESTÁ={res['no_estan']} | "
        f"Discrepancias Excel={res['discrepancias_excel']}"
    )
    print(f"JSON guardado en: {out_json}")

    if not args.no_pdf:
        pdf_path = generar_pdf(out_json, out_pdf)
        print(f"PDF generado en: {pdf_path}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
