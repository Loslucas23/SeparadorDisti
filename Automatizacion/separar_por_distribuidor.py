"""
Separa el backup de ventas en un Excel por distribuidor (datos completos, sin resúmenes).

- Lee TODOS los .xlsb / .xlsx de la carpeta "Backup Distribuidores" (no las subcarpetas).
- Cada hoja de datos es un año fiscal (FY24, FY25, FY28...). Si una hoja no se llama FYxx,
  el año fiscal se calcula por la fecha (julio a junio). Las hojas que no tienen las columnas
  DISTRIBUIDOR / Fecha / NETO s/IMP (resúmenes, tablas dinámicas) se ignoran con aviso.
- Combinación de archivos: para cada MES + DISTRIBUIDOR se usan los datos del archivo
  modificado más recientemente que tenga ese mes para ese distribuidor. Así:
    * un archivo incremental (solo noviembre) agrega noviembre sin borrar los meses previos,
    * una corrección de un mes reemplaza solo ese mes,
    * un archivo con un solo distribuidor no borra a los demás.
- Cada archivo se lee una sola vez y se guarda en caché (_cache); si no cambia, no se relee.
- Si un Excel de salida está abierto, se saltea, se avisa y se sigue con los demás.
- Al final deja control.txt (origen vs. archivos generados) y agrega la corrida a log.txt.

Salida: Backup Distribuidores/Por Distribuidor/<DISTRIBUIDOR>.xlsx  (una hoja por FY)
Códigos de salida: 0 = todo OK · 2 = terminó con avisos (ver control.txt) · 1 = error
"""
import os
import re
import sys
import time
import glob
import shutil
import socket
import hashlib
import tempfile
import traceback
from datetime import datetime
from multiprocessing import Pool

import pandas as pd

FROZEN = getattr(sys, "frozen", False)          # True cuando corre como .exe
if FROZEN:   # el .exe vive en la carpeta "Backup Distribuidores"
    BASE = os.path.dirname(os.path.abspath(sys.executable))
    HERE = os.path.join(BASE, "Automatizacion")
else:        # el .py vive en "Backup Distribuidores/Automatizacion"
    HERE = os.path.dirname(os.path.abspath(__file__))
    BASE = os.path.dirname(HERE)
if os.environ.get("BACKUP_BASE"):    # solo para pruebas: usar otra carpeta
    BASE = os.environ["BACKUP_BASE"]
    HERE = os.path.join(BASE, "Automatizacion")
CACHE_DIR = os.path.join(HERE, "_cache")
OUT_DIR = os.path.join(BASE, "Por Distribuidor")
LOG_FILE = os.path.join(OUT_DIR, "log.txt")
CONTROL_FILE = os.path.join(OUT_DIR, "control.txt")
LOCK_FILE = os.path.join(OUT_DIR, "_EN_USO.txt")

KEY = "DISTRIBUIDOR"     # columna que identifica al distribuidor
NETO = "NETO s/IMP"      # columna que se suma en el control
REQUERIDAS = [KEY, "Fecha", NETO]
PROCESOS = 4             # archivos que se escriben en paralelo
EXCEL_MAX_ROWS = 1_048_000
LOG_MAX_MB = 5
CACHE_VERSION = "v3"     # cambiarlo obliga a releer los orígenes

NUMERIC = ['Vendedor Id.', 'Cliente Id.', 'Producto Cod.', 'Anio', 'Mes', 'Nro. Linea',
           'CANTIDAD', 'CAJAS', 'VOLUMEN', 'BRUTO s/IMP', 'BRUTO c/IMP', 'DTOS s/IMP',
           'DTOS c/IMP', 'NETO s/IMP', 'NETO c/IMP', 'IMPUESTOS']
INTERNAS = ["_src", "_mes"]   # columnas auxiliares, no se escriben
TXT = "_txt_"                 # _txt_<col>: texto original de valores no numéricos en columnas numéricas


# ------------------------------------------------------------------ log
_log_fh = None
AVISOS = []   # requieren revisar algo -> la corrida termina en "ATENCION" (código 2)
NOTAS = []    # informativas (reemplazos normales, hojas de resumen ignoradas)


def log(msg=""):
    linea = f"[{datetime.now():%H:%M:%S}] {msg}" if msg else ""
    print(linea, flush=True)
    if _log_fh:
        _log_fh.write(linea + "\n")
        _log_fh.flush()


def aviso(msg):
    AVISOS.append(msg)
    log("  ! " + msg)


def nota(msg):
    NOTAS.append(msg)
    log("  · " + msg)


def abrir_log():
    global _log_fh
    os.makedirs(OUT_DIR, exist_ok=True)
    if os.path.exists(LOG_FILE) and os.path.getsize(LOG_FILE) > LOG_MAX_MB * 1e6:
        os.replace(LOG_FILE, LOG_FILE.replace("log.txt", "log_anterior.txt"))
    _log_fh = open(LOG_FILE, "a", encoding="utf-8")
    _log_fh.write("\n" + "=" * 78 + "\n")
    _log_fh.write(f"CORRIDA {datetime.now():%d/%m/%Y %H:%M:%S} · usuario: {os.environ.get('USERNAME', '?')}"
                  f" · equipo: {os.environ.get('COMPUTERNAME', '?')}\n")
    _log_fh.write("=" * 78 + "\n")


# ------------------------------------------------------------------ bloqueo de corridas simultáneas
def _pid_vivo(pid):
    """True si el proceso sigue corriendo en esta PC (sin enviarle ninguna señal)."""
    try:
        import ctypes
        k = ctypes.windll.kernel32
        h = k.OpenProcess(0x1000, False, int(pid))       # PROCESS_QUERY_LIMITED_INFORMATION
        if not h:
            return False
        code = ctypes.c_ulong()
        k.GetExitCodeProcess(h, ctypes.byref(code))
        k.CloseHandle(h)
        return code.value == 259                         # STILL_ACTIVE
    except Exception:
        return True                                      # ante la duda, respetar el bloqueo


def tomar_bloqueo():
    if os.path.exists(LOCK_FILE):
        edad_h = (time.time() - os.path.getmtime(LOCK_FILE)) / 3600
        try:
            usuario, equipo, pid, desde = open(LOCK_FILE, encoding="utf-8").read().strip().split("|")
        except (OSError, ValueError):
            usuario, equipo, pid, desde = "?", "?", "0", "?"
        abandonado = equipo == socket.gethostname() and not _pid_vivo(pid)   # ventana cerrada a mitad
        if edad_h < 3 and not abandonado:
            raise SystemExit(f"Ya hay una corrida en curso ({usuario} en {equipo} desde {desde}). "
                             f"Esperar a que termine. Si no hay ninguna, borrar el archivo {LOCK_FILE}")
        log(f"  · bloqueo viejo de {usuario} en {equipo} ({desde}) descartado")
    with open(LOCK_FILE, "w", encoding="utf-8") as fh:
        fh.write(f"{os.environ.get('USERNAME', '?')}|{socket.gethostname()}|{os.getpid()}|"
                 f"{datetime.now():%d/%m/%Y %H:%M}")


def soltar_bloqueo():
    try:
        os.remove(LOCK_FILE)
    except OSError:
        pass


# ------------------------------------------------------------------ lectura
def _to_text(v):
    if v is None or (isinstance(v, float) and v != v):
        return None
    if isinstance(v, float):
        return str(int(v)) if v.is_integer() else repr(v)
    if isinstance(v, datetime):
        return v.strftime("%d/%m/%Y")
    return str(v).strip() or None


def _a_numero(s, col, donde):
    """Convierte a número. Devuelve (números, textos_originales_no_numéricos o None)."""
    num = pd.to_numeric(s, errors="coerce")
    mask = num.isna() & s.notna()
    textos = None
    if mask.any():   # textos tipo "1.234,56" o "$ 100"
        txt = s[mask].astype(str).str.strip().str.replace(r"[\s$]", "", regex=True)
        coma = txt.str.contains(",", regex=False)
        txt = txt.where(~coma, txt.str.replace(".", "", regex=False).str.replace(",", ".", regex=False))
        num.loc[mask] = pd.to_numeric(txt, errors="coerce")
        perdidos = num.isna() & s.notna() & (s.astype(str).str.strip() != "")
        if perdidos.any():
            ej = ", ".join(repr(v) for v in s[perdidos].astype(str).unique()[:3])
            aviso(f"{donde}: columna '{col}' tiene {perdidos.sum():,} valor(es) que no son número "
                  f"(ej: {ej}). Se conservan como texto en la salida, pero no suman en el control")
            textos = pd.Series(None, index=s.index, dtype=object)
            textos[perdidos] = s[perdidos].astype(str).str.strip()
    return num, textos


def _a_fecha(s, donde):
    if pd.api.types.is_datetime64_any_dtype(s):
        return s
    num = pd.to_numeric(s, errors="coerce")
    out = pd.to_datetime(num, unit="D", origin="1899-12-30")
    resto = num.isna() & s.notna()
    if resto.any():   # fechas reales (desde .xlsx) o texto
        r = s[resto]
        iso = pd.to_datetime(r, errors="coerce", format="ISO8601")
        dmy = pd.to_datetime(r[iso.isna()], errors="coerce", dayfirst=True, format="mixed")
        iso.loc[iso.isna()] = dmy
        out.loc[resto] = iso
    sin = out.isna() & s.notna()
    if sin.any():
        ej = ", ".join(repr(v) for v in s[sin].astype(str).unique()[:3])
        aviso(f"{donde}: {sin.sum():,} fecha(s) no se pudieron leer y quedan vacías (ej: {ej})")
    return out


def _fy_de_fecha(fechas):
    y = fechas.dt.year + (fechas.dt.month >= 7).astype("Int64")
    fy = "FY" + (y % 100).astype("string").str.zfill(2)
    return fy.fillna("SIN_FY").astype(object)


def _normalizar(df, hoja, archivo):
    donde = f"{archivo} / hoja {hoja}"
    for h in list(df.columns):
        if h in NUMERIC:
            df[h], textos = _a_numero(df[h], h, donde)
            if textos is not None:
                df[TXT + h] = textos
        elif h == "Fecha":
            df[h] = _a_fecha(df[h], donde)
        else:
            df[h] = df[h].map(_to_text).astype(object)
    m = re.fullmatch(r"\s*(FY\s*\d{2})\s*", str(hoja), re.I)
    df.insert(0, "FY", m.group(1).upper().replace(" ", "") if m else _fy_de_fecha(df["Fecha"]))
    return df


def _encabezado(crudo, donde):
    """Limpia encabezados: sin título -> se descartan; repetidos -> se renombran."""
    header, vistos = [], {}
    for i, h in enumerate(crudo):
        h = None if h is None or (isinstance(h, float) and h != h) else str(h).strip()
        if not h or h.startswith("Unnamed:"):
            header.append(None)
            continue
        if h in vistos:
            vistos[h] += 1
            aviso(f"{donde}: la columna '{h}' está repetida; la copia se guarda como '{h} ({vistos[h]})'")
            h = f"{h} ({vistos[h]})"
        else:
            vistos[h] = 1
        header.append(h)
    return header


def _leer_xlsb(path):
    from pyxlsb import open_workbook
    arch = os.path.basename(path)
    with open_workbook(path) as wb:
        for name in wb.sheets:
            t0 = time.time()
            with wb.get_sheet(name) as sh:
                it = sh.rows()
                try:
                    crudo = [c.v for c in next(it)]
                except StopIteration:
                    continue
                header = _encabezado(crudo, f"{arch} / hoja {name}")
                while header and header[-1] is None:
                    header.pop()
                ncol = len(header)
                usar = [j for j, h in enumerate(header) if h]
                cols = {header[j]: [] for j in usar}
                listas = [cols[header[j]] for j in usar]
                n = huerfanas = 0
                for row in it:
                    vals = [c.v for c in row]
                    if not vals or all(v is None for v in vals):
                        continue
                    if len(vals) < ncol:
                        vals += [None] * (ncol - len(vals))
                    if any(vals[j] is not None for j in range(ncol, len(vals))) or \
                            any(vals[j] is not None for j in range(ncol) if header[j] is None):
                        huerfanas += 1
                    for lst, j in zip(listas, usar):
                        lst.append(vals[j])
                    n += 1
                    if n % 500_000 == 0:
                        log(f"    {name}: {n:,} filas...")
            if huerfanas:
                aviso(f"{arch} / hoja {name}: {huerfanas:,} fila(s) tienen datos en columnas SIN ENCABEZADO "
                      f"(esos valores se ignoran; ponerle título a la columna para conservarlos)")
            log(f"    hoja {name}: {n:,} filas ({time.time() - t0:.0f}s)")
            yield name, pd.DataFrame(cols)


def _leer_xlsx(path):
    arch = os.path.basename(path)
    for name, df in pd.read_excel(path, sheet_name=None, dtype=object, header=None).items():
        if df.empty:
            continue
        header = _encabezado(list(df.iloc[0]), f"{arch} / hoja {name}")
        df = df.iloc[1:]
        sin_titulo = [j for j, h in enumerate(header) if h is None]
        if sin_titulo:
            huerfanas = df.iloc[:, sin_titulo].notna().any(axis=1).sum()
            if huerfanas:
                aviso(f"{arch} / hoja {name}: {huerfanas:,} fila(s) tienen datos en columnas SIN ENCABEZADO "
                      f"(esos valores se ignoran; ponerle título a la columna para conservarlos)")
        usar = [j for j, h in enumerate(header) if h]
        df = df.iloc[:, usar]
        df.columns = [header[j] for j in usar]
        df = df.dropna(how="all").reset_index(drop=True)
        log(f"    hoja {name}: {len(df):,} filas")
        yield name, df


def leer_archivo(path):
    """Devuelve un DataFrame con todas las hojas de datos del archivo (usa caché si no cambió)."""
    st = os.stat(path)
    firma = hashlib.md5(f"{CACHE_VERSION}|{st.st_size}|{st.st_mtime}".encode()).hexdigest()[:10]
    base = re.sub(r"[^\w\-]+", "_", os.path.splitext(os.path.basename(path))[0])
    cache = os.path.join(CACHE_DIR, f"{base}__{firma}.parquet")
    if os.path.exists(cache):
        log(f"  {os.path.basename(path)}: sin cambios (caché)")
        return pd.read_parquet(cache)
    log(f"  {os.path.basename(path)}: leyendo (puede tardar varios minutos)...")
    lector = _leer_xlsb if path.lower().endswith(".xlsb") else _leer_xlsx
    partes = []
    for hoja, df in lector(path):
        faltan = [c for c in REQUERIDAS if c not in df.columns]
        if faltan:
            msg = (f"{os.path.basename(path)} / hoja {hoja}: no es una hoja de datos "
                   f"(faltan columnas {', '.join(faltan)}), se ignora")
            # si la hoja se llama FYxx o tiene muchas filas, probablemente ES de datos -> avisar fuerte
            if re.fullmatch(r"\s*FY\s*\d{2}\s*", str(hoja), re.I) or len(df) > 1000:
                aviso(msg + "  <-- ¿le cambiaron el nombre a una columna?")
            else:
                nota(msg)
            continue
        partes.append(_normalizar(df, hoja, os.path.basename(path)))
    df = pd.concat(partes, ignore_index=True) if partes else pd.DataFrame()
    os.makedirs(CACHE_DIR, exist_ok=True)
    for viejo in glob.glob(os.path.join(CACHE_DIR, f"{base}__*.parquet")):
        os.remove(viejo)
    if not df.empty:
        df.to_parquet(cache, index=False)
    return df


def limpiar_cache(archivos):
    vigentes = {re.sub(r"[^\w\-]+", "_", os.path.splitext(os.path.basename(p))[0]) for p in archivos}
    for c in glob.glob(os.path.join(CACHE_DIR, "*.parquet")):
        if os.path.basename(c).rsplit("__", 1)[0] not in vigentes:
            os.remove(c)


def cargar_todo():
    archivos = [p for p in glob.glob(os.path.join(BASE, "*.xls[bx]"))
                if not os.path.basename(p).startswith("~$")]
    if not archivos:
        raise SystemExit(f"No hay archivos .xlsb/.xlsx en {BASE}")
    archivos.sort(key=os.path.getmtime)          # el más nuevo queda al final
    nombres = [os.path.basename(p) for p in archivos]
    log(f"{len(archivos)} archivo(s) de origen (del más viejo al más nuevo):")
    partes, cols_ref = [], None
    for i, p in enumerate(archivos):
        df = leer_archivo(p)
        if df.empty:
            aviso(f"{nombres[i]}: no tiene hojas de datos, se ignora")
            continue
        cols = [c for c in df.columns if c != "FY" and not str(c).startswith(TXT)]
        if cols_ref is None:
            cols_ref = cols
        else:
            faltan = [c for c in cols_ref if c not in cols]
            sobran = [c for c in cols if c not in cols_ref]
            if faltan or sobran:
                aviso(f"{nombres[i]}: columnas distintas al resto"
                      + (f" · faltan: {', '.join(faltan)}" if faltan else "")
                      + (f" · de más: {', '.join(sobran)}" if sobran else ""))
        df["_src"] = i
        partes.append(df)
    limpiar_cache(archivos)
    if not partes:
        raise SystemExit("Ningún archivo tiene hojas de datos válidas")
    todo = pd.concat(partes, ignore_index=True)
    todo[KEY] = todo[KEY].fillna("(SIN DISTRIBUIDOR)")
    todo["_mes"] = todo["Fecha"].dt.strftime("%Y-%m").fillna(todo["FY"] + " sin fecha")

    # para cada mes + distribuidor gana el archivo más nuevo que lo tenga
    ganador = todo.groupby(["_mes", KEY])["_src"].transform("max")
    queda = todo["_src"] == ganador
    if (~queda).any():
        desc = todo.loc[~queda, ["_src", "_mes", KEY]].assign(_gan=ganador[~queda])
        for (src, gan), g in desc.groupby(["_src", "_gan"]):
            meses = sorted(g["_mes"].unique())
            dists = g[KEY].nunique()
            nota(f"'{nombres[src]}': {len(g):,} fila(s) reemplazadas por la versión más nueva de "
                 f"'{nombres[gan]}' ({dists} distribuidor(es), meses {', '.join(meses[:6])}"
                 f"{'...' if len(meses) > 6 else ''})")
        # ¿el archivo nuevo trae mucho menos que el que reemplaza? -> posible archivo incompleto
        cnt = todo.groupby(["_mes", KEY, "_src"]).size().rename("n").reset_index()
        cnt["_gan"] = cnt.groupby(["_mes", KEY])["_src"].transform("max")
        nuevo = cnt[cnt["_src"] == cnt["_gan"]].set_index(["_mes", KEY])["n"]
        viejo = cnt[cnt["_src"] != cnt["_gan"]].groupby(["_mes", KEY])["n"].max()
        comp = pd.concat([nuevo.rename("nuevo"), viejo.rename("viejo")], axis=1, join="inner")
        sosp = comp[(comp["viejo"] >= 20) & (comp["nuevo"] < 0.5 * comp["viejo"])]
        for (mes, dist), r in sosp.head(15).iterrows():
            aviso(f"{dist} {mes}: el archivo más nuevo trae {int(r['nuevo']):,} filas y reemplaza a uno "
                  f"que tenía {int(r['viejo']):,}. ¿Archivo incompleto?")
        if len(sosp) > 15:
            aviso(f"... y {len(sosp) - 15} combinaciones mes/distribuidor más con la misma situación")
    for i, n in enumerate(nombres):
        x = todo[queda & (todo["_src"] == i)]
        if len(x):
            log(f"  usa {len(x):>10,} filas de {n}  ({', '.join(sorted(x['FY'].unique()))})")
    return todo, queda, nombres


# ------------------------------------------------------------------ escritura
class ArchivoAbierto(Exception):
    pass


def nombre_archivo(s):
    s = re.sub(r'[<>:"/\\|?*\x00-\x1f]+', "-", str(s)).strip(" .")
    if s.upper() in {"CON", "PRN", "AUX", "NUL"} or re.fullmatch(r"(COM|LPT)\d", s.upper()):
        s = "_" + s
    return s[:80] or "SIN NOMBRE"


def esta_bloqueado(path):
    """True si el archivo existe y otro programa (Excel) lo tiene abierto."""
    if not os.path.exists(path):
        return False
    try:
        with open(path, "r+b"):
            return False
    except PermissionError:
        return True


def escribir(d, path):
    """Escribe el Excel del distribuidor. Devuelve {FY: [filas, neto]} de lo efectivamente escrito."""
    import xlsxwriter
    if esta_bloqueado(path):
        raise ArchivoAbierto()
    fd, tmp = tempfile.mkstemp(suffix=".xlsx", prefix="dist_")
    os.close(fd)
    stats = {}
    try:
        wb = xlsxwriter.Workbook(tmp, {"constant_memory": True})
        f_h = wb.add_format({"bold": True, "bg_color": "#1F3864", "font_color": "white"})
        f_date = wb.add_format({"num_format": "dd/mm/yyyy"})
        cols = [c for c in d.columns if c != "FY" and c not in INTERNAS and not str(c).startswith(TXT)]
        tipos = ["d" if c == "Fecha" else ("n" if pd.api.types.is_numeric_dtype(d[c]) else "s") for c in cols]
        j_neto = cols.index(NETO) if NETO in cols else -1
        for fy in sorted(d["FY"].unique()):
            x = d.loc[d["FY"] == fy, cols]
            if "Fecha" in cols:
                x = x.sort_values("Fecha", kind="stable")
            stats[fy] = [0, 0.0]
            for k, i0 in enumerate(range(0, max(len(x), 1), EXCEL_MAX_ROWS)):
                t = x.iloc[i0:i0 + EXCEL_MAX_ROWS]
                ws = wb.add_worksheet(fy if len(x) <= EXCEL_MAX_ROWS else f"{fy} ({k + 1})")
                ws.write_row(0, 0, cols, f_h)
                ws.freeze_panes(1, 0)
                ws.autofilter(0, 0, max(len(t), 1), len(cols) - 1)
                ws.set_column(0, len(cols) - 1, 14)

                def texto(r, j, v, _ws=ws):
                    # códigos numéricos (EAN, RUT...) vuelven a ser número como en el origen
                    if v.isascii() and v.isdigit() and len(v) <= 15 and (v[0] != "0" or len(v) == 1):
                        _ws.write_number(r, j, int(v))
                    else:
                        _ws.write_string(r, j, v)

                datos, escritores = [], []
                for j, (c, tp) in enumerate(zip(cols, tipos)):
                    if tp == "d":
                        datos.append([v.to_pydatetime() if pd.notna(v) else None for v in t[c]])
                        ws.set_column(j, j, 11, f_date)
                        escritores.append(ws.write_datetime)
                    elif tp == "n":
                        vals = [None if v != v or v in (float("inf"), float("-inf")) else v
                                for v in t[c].tolist()]
                        if TXT + c in t.columns or TXT + c in d.columns:
                            orig = d.loc[t.index, TXT + c].tolist()   # devolver el texto original ("CC")
                            vals = [o if v is None and isinstance(o, str) else v for v, o in zip(vals, orig)]
                            escritores.append(lambda r, j, v, _ws=ws: _ws.write_string(r, j, v)
                                              if isinstance(v, str) else _ws.write_number(r, j, v))
                        else:
                            escritores.append(ws.write_number)
                        datos.append(vals)
                    else:
                        datos.append(t[c].tolist())
                        escritores.append(texto)
                neto = 0.0
                r = 0
                for r, fila in enumerate(zip(*datos), 1):
                    for j, v in enumerate(fila):
                        if v is None or v != v:
                            continue
                        escritores[j](r, j, v)
                    if j_neto >= 0:
                        v = fila[j_neto]
                        if v is not None and not isinstance(v, str) and v == v:
                            neto += v
                stats[fy][0] += r
                stats[fy][1] += neto
        wb.close()
        try:
            os.replace(tmp, path)
        except PermissionError:
            raise ArchivoAbierto()
        except OSError:                          # temp en otro disco
            try:
                shutil.copyfile(tmp, path)
            except PermissionError:
                raise ArchivoAbierto()
    finally:
        if os.path.exists(tmp):
            os.remove(tmp)
    return stats


def _tarea(args):
    nombre, d, path = args
    t0 = time.time()
    try:
        return nombre, "OK", escribir(d, path), time.time() - t0, ""
    except ArchivoAbierto:
        return nombre, "ABIERTO", {}, time.time() - t0, ""
    except Exception as e:
        return nombre, "ERROR", {}, time.time() - t0, f"{type(e).__name__}: {e}"


def asignar_archivos(grupos):
    """Nombre de archivo por distribuidor, sin choques (Windows no distingue mayúsculas)."""
    usados, tareas = {}, []
    for n, d in grupos:
        base = nombre_archivo(n)
        nom, k = base, 1
        while nom.upper() in usados:
            k += 1
            nom = f"{base} ({k})"
        if k > 1:
            aviso(f"'{n}' y '{usados[base.upper()]}' darían el mismo nombre de archivo; "
                  f"'{n}' se guarda como '{nom}.xlsx' (¿es el mismo distribuidor escrito distinto?)")
        usados[nom.upper()] = n
        tareas.append((n, d, os.path.join(OUT_DIR, nom + ".xlsx")))
    return tareas


# ------------------------------------------------------------------ control
def fmt_n(v, dec=0):
    s = f"{v:,.{dec}f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def escribir_control(todo, queda, nombres, tareas, resultados, filtro, completa):
    leido = todo.groupby("FY").agg(filas=("FY", "size"), neto=(NETO, "sum"))
    reempl = todo[~queda].groupby("FY").agg(filas=("FY", "size"), neto=(NETO, "sum"))
    gen = {}
    for _, estado, stats, _, _ in resultados:
        for fy, (f, n) in stats.items():
            g = gen.setdefault(fy, [0, 0.0])
            g[0] += f
            g[1] += n
    L = []
    w = L.append
    w(f"CONTROL DE INTEGRIDAD · {datetime.now():%d/%m/%Y %H:%M:%S}")
    w(f"Distribuidores: {'solo los que contienen ' + repr(filtro) if filtro else 'todos'}")
    w("Archivos de origen (del más viejo al más nuevo): " + " · ".join(nombres))
    w("")
    w("Filas leídas = todas las filas de todos los archivos de origen")
    w("Reemplazadas = filas de un archivo más viejo cuyo mes+distribuidor vino en un archivo más nuevo")
    w("Esperadas    = leídas - reemplazadas  ->  deben coincidir con lo escrito en los archivos")
    w("")
    enc = (f"{'FY':<8}{'Leídas':>12}{'Reemplaz.':>11}{'Esperadas':>12}{'En archivos':>13}{'Dif.':>8}"
           f"{'NETO esperado':>20}{'NETO archivos':>20}{'Dif.':>12}  Estado")
    w(enc)
    w("-" * len(enc))
    todo_ok = True
    tot = [0] * 4 + [0.0] * 2
    for fy in sorted(set(leido.index) | set(gen)):
        fl = int(leido.loc[fy, "filas"]) if fy in leido.index else 0
        fr = int(reempl.loc[fy, "filas"]) if fy in reempl.index else 0
        ne = (float(leido.loc[fy, "neto"]) if fy in leido.index else 0.0) - \
             (float(reempl.loc[fy, "neto"]) if fy in reempl.index else 0.0)
        fe = fl - fr
        fg, ng = gen.get(fy, [0, 0.0])
        ok = fe == fg and abs(ne - ng) < 0.01 * max(1, fe) ** 0.5 + 0.01
        todo_ok &= ok
        tot = [tot[0] + fl, tot[1] + fr, tot[2] + fe, tot[3] + fg, tot[4] + ne, tot[5] + ng]
        w(f"{fy:<8}{fmt_n(fl):>12}{fmt_n(fr):>11}{fmt_n(fe):>12}{fmt_n(fg):>13}{fmt_n(fg - fe):>8}"
          f"{fmt_n(ne, 2):>20}{fmt_n(ng, 2):>20}{fmt_n(ng - ne, 2):>12}  {'OK' if ok else 'DIFERENCIA'}")
    w("-" * len(enc))
    w(f"{'TOTAL':<8}{fmt_n(tot[0]):>12}{fmt_n(tot[1]):>11}{fmt_n(tot[2]):>12}{fmt_n(tot[3]):>13}"
      f"{fmt_n(tot[3] - tot[2]):>8}{fmt_n(tot[4], 2):>20}{fmt_n(tot[5], 2):>20}{fmt_n(tot[5] - tot[4], 2):>12}")
    w("")

    # verificación en disco: cada distribuidor generado tiene su propio archivo
    rutas = {n: p for n, _, p in tareas}
    faltan_disco = [n for n, e, *_ in resultados if e == "OK" and not os.path.exists(rutas[n])]
    if faltan_disco or len(set(p.upper() for p in rutas.values())) != len(rutas):
        todo_ok = False
        w("ERROR: hay distribuidores generados cuyo archivo no está en disco: " + ", ".join(faltan_disco))
        w("")

    abiertos = [r for r in resultados if r[1] == "ABIERTO"]
    errores = [r for r in resultados if r[1] == "ERROR"]
    if abiertos:
        w("ARCHIVOS NO ACTUALIZADOS (estaban abiertos) - su diferencia explica la de arriba:")
        for nombre, *_ in abiertos:
            d = todo[queda & (todo[KEY] == nombre)]
            w(f"  - {nombre}: {fmt_n(len(d))} filas, NETO {fmt_n(d[NETO].sum(), 2)}")
        w("")
    if errores:
        w("ARCHIVOS CON ERROR:")
        for nombre, _, _, _, msg in errores:
            w(f"  - {nombre}: {msg}")
        w("")
    if completa:
        generados = {os.path.basename(p).upper() for p in rutas.values()}
        viejos = sorted(f for f in os.listdir(OUT_DIR) if f.lower().endswith(".xlsx")
                        and not f.startswith("~$") and f.upper() not in generados)
        if viejos:
            w("ARCHIVOS VIEJOS en 'Por Distribuidor' que ya NO corresponden a ningún distribuidor del origen")
            w("(¿distribuidor renombrado o que dejó de existir?). No se tocaron; revisar y borrar si sobran:")
            for f in viejos:
                w(f"  - {f}")
                AVISOS.append(f"archivo viejo sin distribuidor en el origen: {f}")
            w("")
    if AVISOS:
        w("AVISOS (revisar):")
        for a in AVISOS:
            w(f"  ! {a}")
        w("")
    if NOTAS:
        w("NOTAS (informativas):")
        for a in NOTAS:
            w(f"  · {a}")
        w("")
    resultado = todo_ok and not abiertos and not errores
    w("RESULTADO: " + ("OK - no se perdió ninguna fila" if resultado
                       else "REVISAR - hay diferencias o archivos sin actualizar"))
    with open(CONTROL_FILE, "w", encoding="utf-8") as fh:
        fh.write("\n".join(L) + "\n")
    for l in L:
        log(l)
    return resultado


# ------------------------------------------------------------------ main
def menu_exe():
    """Menú interactivo cuando se abre el .exe con doble clic."""
    print("=" * 60)
    print("  BACKUP POR DISTRIBUIDOR")
    print("=" * 60)
    print()
    print(" Toma TODOS los archivos .xlsb / .xlsx de esta carpeta")
    print(' y crea un Excel por distribuidor en "Por Distribuidor\\"')
    print(" (una hoja por año fiscal: FY24, FY25, ... FY28).")
    print()
    print(" - FY nuevo, FY anteriores o un mes suelto: copiar el archivo")
    print("   acá (hojas llamadas FY28, FY24, etc.).")
    print(" - Por cada mes y distribuidor se usa el archivo más nuevo.")
    print(" - Si un Excel de distribuidor está abierto, se saltea y se avisa.")
    print(" - Al final: control.txt (origen vs generado) y log.txt (historial).")
    print()
    print(" 1 = Generar TODOS los distribuidores")
    print(" 2 = Generar UN distribuidor")
    print()
    limpiar = lambda s: s.replace("﻿", "").replace("ï»¿", "").strip()
    opc = re.sub(r"\D", "", input("Opción [1]: "))
    if opc == "2":
        nombre = ""
        while not nombre:
            nombre = limpiar(input("Nombre (o parte del nombre) del distribuidor: "))
        return nombre
    return ""


def main(filtro=None):
    abrir_log()
    tomar_bloqueo()
    try:
        t_all = time.time()
        log(f"Carpeta: {BASE}")
        todo, queda, nombres = cargar_todo()
        if filtro is None:
            filtro = " ".join(sys.argv[1:])
        filtro = filtro.strip().upper()
        if filtro:
            sel = todo[KEY].astype(str).str.upper().str.contains(filtro, regex=False)
            todo, queda = todo[sel], queda[sel]
            if todo.empty:
                raise SystemExit(f"Ningún distribuidor contiene '{filtro}'")
        df = todo[queda]
        grupos = sorted(df.groupby(KEY), key=lambda g: (-len(g[1]), str(g[0])))   # los grandes primero
        tareas = asignar_archivos(grupos)
        log(f"Generando {len(tareas)} archivo(s) en: {OUT_DIR}")
        resultados = []
        with Pool(min(PROCESOS, len(tareas))) as pool:
            for i, (nombre, estado, stats, seg, msg) in enumerate(pool.imap_unordered(_tarea, tareas), 1):
                resultados.append((nombre, estado, stats, seg, msg))
                filas = sum(f for f, _ in stats.values())
                if estado == "OK":
                    log(f"  [{i}/{len(tareas)}] {nombre}: {fmt_n(filas)} filas ({', '.join(sorted(stats))}) - {seg:.0f}s")
                elif estado == "ABIERTO":
                    log(f"  [{i}/{len(tareas)}] !! {nombre}: el archivo está ABIERTO, se saltea (cerrarlo y volver a correr)")
                else:
                    log(f"  [{i}/{len(tareas)}] !! {nombre}: ERROR {msg}")
        log("")
        ok = escribir_control(todo, queda, nombres, tareas, resultados, filtro, completa=not filtro)
        log(f"Terminado en {time.time() - t_all:.0f}s. Archivos en: {OUT_DIR}")
        return 0 if ok and not AVISOS else 2
    finally:
        soltar_bloqueo()


if __name__ == "__main__":
    from multiprocessing import freeze_support
    freeze_support()                     # necesario para el .exe con procesos en paralelo
    interactivo = FROZEN and len(sys.argv) == 1
    try:
        codigo = main(menu_exe() if interactivo else None)
    except SystemExit as e:
        log(f"DETENIDO: {e}")
        codigo = 1
    except Exception:
        log("ERROR INESPERADO:\n" + traceback.format_exc())
        codigo = 1
    finally:
        if _log_fh:
            _log_fh.close()
    if interactivo:
        print()
        if codigo == 0:
            print('[OK] Todo generado y controlado. Ver "Por Distribuidor\\control.txt"')
        elif codigo == 2:
            print("[ATENCION] Terminó con avisos: archivos abiertos, reemplazos o diferencias.")
            print('           Revisar "Por Distribuidor\\control.txt".')
        else:
            print('[ERROR] Algo falló. Revisar el mensaje de arriba o "Por Distribuidor\\log.txt"')
        if codigo in (0, 2) and os.path.isdir(OUT_DIR):
            os.startfile(OUT_DIR)
        input("\nPresione Enter para cerrar...")
    sys.exit(codigo)
