"""
pipeline/runner.py — importa funciones directamente, sin subprocesos ni rutas en código.
"""
import sys, os, re, tempfile, shutil, importlib.util, types
from io import BytesIO
from pathlib import Path

PIPELINE_DIR = Path(__file__).parent

OUTPUTS = {
    'caract_excel': ('TOP_Caracterizacion_Ingreso.xlsx','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'),
    'seg_excel':    ('TOP_Seguimiento.xlsx','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'),
    'pdf_caract':   ('TOP_Informe_Caracterizacion.docx','application/vnd.openxmlformats-officedocument.wordprocessingml.document'),
    'pdf_seg':      ('TOP_Informe_Seguimiento.docx','application/vnd.openxmlformats-officedocument.wordprocessingml.document'),
    'pptx_caract':  ('TOP_Presentacion_Caracterizacion.pptx','application/vnd.openxmlformats-officedocument.presentationml.presentation'),
    'pptx_seg':     ('TOP_Presentacion_Seguimiento.pptx','application/vnd.openxmlformats-officedocument.presentationml.presentation'),
}

SCRIPT_FILES = {
    'caract_excel': 'caract_excel.py',
    'seg_excel':    'seg_excel.py',
    'pdf_caract':   'word_caract.py',
    'pdf_seg':      'word_seg.py',
    'pptx_caract':  'pptx_caract.py',
    'pptx_seg':     'pptx_seg.py',
}


def _load_mod(script_key, wide_path, out_path):
    """
    Carga un script como módulo Python sin ejecutar código de nivel superior
    que dependa de rutas de archivos.
    """
    import types, re as _re
    src = open(str(PIPELINE_DIR / SCRIPT_FILES[script_key]), encoding='utf-8').read()

    # 1. Neutralizar INPUT_FILE = auto_archivo_wide()
    src = _re.sub(
        r'INPUT_FILE\s*=\s*auto_archivo_wide\(\)',
        'INPUT_FILE = None',
        src
    )
    # 2. Neutralizar OUTPUT_FILE hardcodeado
    for old in ["'/home/claude/TOP_Caracterizacion_Ingreso.xlsx'",
                "'/home/claude/TOP_Informe_Caracterizacion.pdf'",
                "'/home/claude/TOP_Informe_Seguimiento.pdf'",
                "'/home/claude/TOP_Presentacion_Caracterizacion.pptx'",
                "'/home/claude/TOP_Presentacion_Seguimiento.pptx'",
                "'/home/claude/TOP_Seguimiento.xlsx'"]:
        src = src.replace(old, 'None')

    # 3. Neutralizar llamadas a nivel de módulo que usen INPUT_FILE
    #    (como _detectar_pais(INPUT_FILE) o _pais_detectado = ...)
    src = _re.sub(
        r'^(_pais_detectado\s*=\s*_detectar_pais\(INPUT_FILE\))',
        r'_pais_detectado = None  # neutralizado',
        src, flags=_re.MULTILINE
    )

    # Crear módulo limpio y ejecutar
    mod = types.ModuleType('_qmod')
    mod.__file__ = str(PIPELINE_DIR / SCRIPT_FILES[script_key])
    try:
        exec(compile(src, '<qalat>', 'exec'), mod.__dict__)
    except SystemExit:
        pass
    except Exception:
        pass

    # Inyectar rutas correctas después de cargar
    mod.__dict__['INPUT_FILE']  = wide_path
    mod.__dict__['OUTPUT_FILE'] = out_path
    mod.__dict__['auto_archivo_wide'] = lambda: wide_path

    return mod


def run_script(script_key, wide_path, filtro_centro=None):
    out_filename, mimetype = OUTPUTS[script_key]
    if filtro_centro:
        base, ext = out_filename.rsplit('.', 1)
        out_filename = f'{base}_{filtro_centro}.{ext}'

    # Archivo de salida temporal
    suffix = '.' + out_filename.rsplit('.', 1)[1]
    fd, out_path = tempfile.mkstemp(suffix=suffix, prefix='qalat_out_')
    os.close(fd)

    try:
        if script_key == 'caract_excel':
            from openpyxl import Workbook
            mod = _load_mod(script_key, wide_path, out_path)
            d, N = mod.cargar_ingreso()
            DC = mod.detectar_columnas(d.columns.tolist())
            wb = Workbook()
            mod.build_report(wb, d, N, DC)
            wb.save(out_path)

        elif script_key == 'seg_excel':
            from openpyxl import Workbook
            mod = _load_mod(script_key, wide_path, out_path)
            seg, N_total, N_seg, seg_tiempo = mod.cargar_datos()
            DC = mod.detectar_columnas(seg.columns.tolist())
            wb = Workbook()
            mod.build_seguimiento(wb, seg, N_total, N_seg, DC, seg_tiempo)
            mod.build_cambio_consumo(wb, seg, N_seg, DC)
            wb.save(out_path)

        elif script_key == 'pdf_caract':
            mod = _load_mod(script_key, wide_path, out_path)
            R = mod.cargar_datos()
            mod.build_word(R)

        elif script_key == 'pdf_seg':
            mod = _load_mod(script_key, wide_path, out_path)
            R = mod.cargar_datos()
            mod.build_word(R)

        elif script_key in ('pptx_caract', 'pptx_seg'):
            import subprocess
            src = open(str(PIPELINE_DIR / SCRIPT_FILES[script_key]),
                       encoding='utf-8').read()

            # Calcular directorio temporal para json/js auxiliares
            tmp_dir = os.path.dirname(out_path).replace('\\','/')

            # 1. Parchear auto_archivo_wide
            src = re.sub(
                r'def auto_archivo_wide\(\):.*?return [^\n]+\n',
                'def auto_archivo_wide():\n    import os as _os\n    return _os.environ["QALAT_WIDE"]\n',
                src, flags=re.DOTALL
            )
            # 2. Parchear glob directo
            src = src.replace(
                "glob.glob('/home/claude/TOP_Base_Wide.xlsx')",
                '[__import__("os").environ["QALAT_WIDE"]]'
            )
            # 3. Parchear OUTPUT_FILE
            for old_path in [
                "'/home/claude/TOP_Presentacion_Caracterizacion.pptx'",
                "'/home/claude/TOP_Presentacion_Seguimiento.pptx'",
                "'/home/claude/TOP_Informe_Caracterizacion.pdf'",
                "'/home/claude/TOP_Informe_Seguimiento.pdf'",
                "'/home/claude/TOP_Seguimiento.xlsx'",
                "'/home/claude/TOP_Caracterizacion_Ingreso.xlsx'",
            ]:
                src = src.replace(old_path, '__import__("os").environ["QALAT_OUT"]')

            # 4. Parchear rutas de archivos auxiliares JSON y JS a directorio tmp
            for old_aux, new_aux in [
                ("'/home/claude/_top_car_data.json'",
                 f'"{tmp_dir}/_top_car_data.json"'),
                ("'/home/claude/_top_data.json'",
                 f'"{tmp_dir}/_top_data.json"'),
                ("'/home/claude/_top_car_builder.js'",
                 f'"{tmp_dir}/_top_car_builder.js"'),
                ("'/home/claude/_top_builder.js'",
                 f'"{tmp_dir}/_top_builder.js"'),
                # Rutas dentro del JS (readFileSync)
                ("fs.readFileSync('/home/claude/_top_car_data.json'",
                 f"fs.readFileSync('{tmp_dir}/_top_car_data.json'"),
                ("fs.readFileSync('/home/claude/_top_data.json'",
                 f"fs.readFileSync('{tmp_dir}/_top_data.json'"),
            ]:
                src = src.replace(old_aux, new_aux)

            # 5. Parchear OUTPUT_FILE en bloque JS
            src = src.replace(
                'OUTPUT_FILE + r"""',
                '__import__("os").environ["QALAT_OUT"] + r"""'
            )

            # Guardar script parcheado
            fd2, tmp_py = tempfile.mkstemp(suffix='.py', prefix='qs_')
            os.close(fd2)
            with open(tmp_py, 'w', encoding='utf-8') as f:
                f.write(src)

            env = os.environ.copy()
            env['QALAT_WIDE'] = wide_path
            env['QALAT_OUT']  = out_path

            try:
                r = subprocess.run(
                    [sys.executable, tmp_py],
                    capture_output=True, text=True,
                    timeout=180, env=env
                )
                if r.returncode != 0:
                    raise RuntimeError(r.stderr[-1000:] or r.stdout[-1000:])
            finally:
                try: os.unlink(tmp_py)
                except: pass

        if not os.path.exists(out_path) or os.path.getsize(out_path) == 0:
            raise FileNotFoundError('El script no generó salida')

        with open(out_path, 'rb') as f:
            data = f.read()
        return BytesIO(data), out_filename, mimetype

    finally:
        try: os.unlink(out_path)
        except: pass


def run_all(wide_path, progress_cb=None):
    results = {}
    keys = list(OUTPUTS.keys())
    for i, key in enumerate(keys):
        if progress_cb: progress_cb(i, len(keys), key)
        try:
            buf, fname, mime = run_script(key, wide_path)
            results[key] = {'ok': True, 'buf': buf, 'fname': fname, 'mime': mime}
        except Exception as e:
            results[key] = {'ok': False, 'error': str(e)}
    if progress_cb: progress_cb(len(keys), len(keys), 'listo')
    return results
