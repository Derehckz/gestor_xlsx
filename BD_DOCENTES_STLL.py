import pandas as pd
from tabulate import tabulate
from pathlib import Path
import shutil
import datetime
import logging
import json
import argparse
import re
import sys
import time

try:
    from unidecode import unidecode
except ImportError:
    unidecode = None

try:
    from colorama import Fore, Style, init
    init(autoreset=True)
except ImportError:
    print("Instala colorama con: pip install colorama")
    sys.exit(1)


def mostrar_mensaje(msg, tipo="info"):
    colores = {
        "info": Fore.CYAN + "ℹ️ ",
        "exito": Fore.GREEN + "✅ ",
        "advertencia": Fore.YELLOW + "⚠️ ",
        "error": Fore.RED + "❌ ",
        "pregunta": Fore.BLUE + "❓ ",
    }
    print(colores.get(tipo, "") + msg + Style.RESET_ALL)


def animacion_carga(msg="Cargando", duracion=2):
    for i in range(duracion * 2):
        sys.stdout.write(Fore.YELLOW + f"\r⏳ {msg}{'.' * (i % 4)}   ")
        sys.stdout.flush()
        time.sleep(0.5)
    print("\r", end='')


def explorar_directorios(inicio: Path = Path.home()) -> Path:
    current = inicio
    while True:
        try:
            items = sorted([p for p in current.iterdir()])
        except PermissionError:
            mostrar_mensaje(f"No se puede acceder a {current}. Subiendo nivel.", "advertencia")
            current = current.parent
            continue

        print(f"\n📂 Directorio: {current}")
        for i, p in enumerate(items, 1):
            tipo = "📁" if p.is_dir() else "📄"
            print(f"{i:2d}. {tipo} {p.name}")
        print(" 0. 🔙 Subir nivel")
        choice = input("Selecciona número (o 'q' para cancelar): ").strip().lower()
        if choice == 'q':
            mostrar_mensaje("Selección cancelada por el usuario.", "advertencia")
            sys.exit(0)
        if not choice.isdigit():
            mostrar_mensaje("Opción inválida.", "error")
            continue
        idx = int(choice)
        if idx == 0:
            if current.parent != current:
                current = current.parent
            else:
                mostrar_mensaje("Ya estás en la raíz del sistema de archivos.", "advertencia")
        elif 1 <= idx <= len(items):
            sel = items[idx-1]
            if sel.is_dir():
                current = sel
            elif sel.is_file() and sel.suffix.lower() in ('.xlsx', '.xlsm', '.xls'):
                return sel
            else:
                mostrar_mensaje("No es un archivo Excel válido (.xlsx/.xlsm/.xls). Elige otro.", "advertencia")
        else:
            mostrar_mensaje("Índice fuera de rango.", "error")


def input_validado(prompt, validacion_func, mensaje_error, formato_func=None, opcional=False):
    while True:
        valor = input(prompt).strip()
        if opcional and valor == '':
            return ''
        if formato_func:
            try:
                valor_formateado = formato_func(valor)
            except Exception:
                mostrar_mensaje("Formato inválido, intente de nuevo.", "error")
                continue
        else:
            valor_formateado = valor
        if validacion_func(valor_formateado):
            return valor_formateado
        else:
            mostrar_mensaje(mensaje_error, "error")


class GestorDocentes:
    def __init__(self, ruta: Path, backup_dir: Path = None, lock_timeout: int = 300):
        self.ruta = ruta
        self.backup_dir = backup_dir or ruta.parent / "backups"
        self.backup_dir.mkdir(parents=True, exist_ok=True)
        self.columnas = []  # dinámico
        self.col_rut = None
        self.col_email = None
        self.col_tel = None

        logging.basicConfig(
            filename=str(ruta.parent / "gestor_docentes.log"),
            level=logging.INFO,
            format="%(asctime)s - %(levelname)s - %(message)s"
        )
        self.lock_file = ruta.with_suffix('.lock')
        self.lock_timeout = lock_timeout

    def acquire_lock(self):
        start = time.time()
        while self.lock_file.exists():
            if time.time() - start > self.lock_timeout:
                mostrar_mensaje("El archivo está bloqueado por mucho tiempo. Verifique manualmente.", "advertencia")
                logging.warning("Timeout esperando lock para %s", self.ruta)
                return False
            time.sleep(1)
        try:
            self.lock_file.write_text(str(time.time()))
            return True
        except Exception as e:
            logging.error("No se pudo crear archivo de lock: %s", e)
            return False

    def release_lock(self):
        try:
            if self.lock_file.exists():
                self.lock_file.unlink()
        except Exception as e:
            logging.error("No se pudo eliminar archivo de lock: %s", e)

    def clean_rut(self, rut: str) -> str:
        if not isinstance(rut, str):
            rut = str(rut or "")
        return re.sub(r"[.\-\s]", "", rut).upper()

    def format_rut(self, rut: str) -> str:
        rn = self.clean_rut(rut)
        if len(rn) < 2:
            return rut
        num, dv = rn[:-1], rn[-1]
        return f"{num}-{dv}"

    def leer(self) -> pd.DataFrame:
        if not self.ruta.exists():
            mostrar_mensaje(f"No se encontró {self.ruta}.", "advertencia")
            crear = input("¿Deseas crear un nuevo archivo Excel con columnas personalizadas? (s/n): ").strip().lower()
            if crear == 's':
                cols = input("Ingresa los nombres de las columnas, separados por coma (ej: RUT,NOMBRE,Email,Teléfono): ").strip()
                nombres = [c.strip() for c in cols.split(',') if c.strip()]
                if not nombres:
                    mostrar_mensaje("No se definieron columnas. Abortando.", "error")
                    sys.exit(1)
                df = pd.DataFrame(columns=nombres)
                self.columnas = nombres
                try:
                    df.to_excel(self.ruta, index=False)
                    mostrar_mensaje(f"Archivo creado con columnas: {self.columnas}", "exito")
                except Exception as e:
                    mostrar_mensaje(f"No se pudo crear el archivo: {e}", "error")
                    logging.error("Error al crear nuevo Excel: %s", e, exc_info=True)
                    sys.exit(1)
                return df
            else:
                mostrar_mensaje("No se creó archivo. Saliendo.", "info")
                sys.exit(0)
        try:
            df = pd.read_excel(self.ruta, dtype=str)
            df = df.fillna("")
            self.columnas = df.columns.tolist()
            return df
        except Exception as e:
            logging.error(f"Error al leer Excel: {e}", exc_info=True)
            mostrar_mensaje(f"Error al leer Excel: {e}", "error")
            sys.exit(1)

    def backup(self):
        if not self.ruta.exists():
            return
        ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        destino = self.backup_dir / f"{self.ruta.stem}_{ts}{self.ruta.suffix}"
        try:
            shutil.copy2(self.ruta, destino)
            logging.info(f"Backup creado en {destino}")
        except Exception as e:
            logging.error(f"Error al crear backup: {e}", exc_info=True)

    def guardar(self, df: pd.DataFrame):
        if not self.acquire_lock():
            mostrar_mensaje("No se pudo obtener lock para guardar. Abortando.", "error")
            return
        try:
            if self.ruta.exists():
                self.backup()
            animacion_carga("Guardando")
            df.to_excel(self.ruta, index=False)
            logging.info("Archivo guardado correctamente.")
            mostrar_mensaje("Archivo guardado exitosamente.", "exito")
        except Exception as e:
            logging.error(f"Error al guardar Excel: {e}", exc_info=True)
            mostrar_mensaje(f"No se pudo guardar el archivo: {e}", "error")
        finally:
            self.release_lock()

    def validar_rut(self, rut: str) -> bool:
        rut_clean = self.clean_rut(rut)
        if not re.match(r"^\d{7,8}[0-9K]$", rut_clean):
            return False
        num = rut_clean[:-1]
        dv = rut_clean[-1]
        try:
            reversed_digits = map(int, reversed(num))
            factors = [2, 3, 4, 5, 6, 7]
            total = 0
            factor_index = 0
            for d in reversed_digits:
                total += d * factors[factor_index]
                factor_index = (factor_index + 1) % len(factors)
            mod = 11 - (total % 11)
            dv_calculado = 'K' if mod == 10 else '0' if mod == 11 else str(mod)
            return dv_calculado == dv
        except Exception:
            return False

    def validar_email(self, email: str) -> bool:
        return re.match(r'^[\w\.-]+@[\w\.-]+\.\w+$', email) is not None

    def validar_telefono(self, telefono: str) -> bool:
        t = re.sub(r"[ \-()]+", "", telefono)
        return t.isdigit() and 7 <= len(t) <= 15

    def mapear_columnas_clave(self, df: pd.DataFrame):
        cols = self.columnas
        mostrar_mensaje(f"Columnas detectadas: {cols}", "info")

        def detectar_y_preguntar(tipo, patrones):
            candidatos = [c for c in cols if any(pat.lower() in c.lower() for pat in patrones)]
            elegido = None
            if candidatos:
                for c in candidatos:
                    resp = input(f"¿La columna '{c}' corresponde a {tipo}? (s/n): ").strip().lower()
                    if resp == 's':
                        elegido = c
                        break
            if not elegido:
                resp = input(f"Ingrese el nombre de la columna para {tipo} (o Enter para omitir validación): ").strip()
                if resp:
                    if resp in cols:
                        elegido = resp
                    else:
                        mostrar_mensaje(f"No existe columna '{resp}'. Se omitirá validación de {tipo}.", "advertencia")
                        elegido = None
            return elegido

        self.col_rut = detectar_y_preguntar("RUT (para validar formato chileno)", ["rut"])
        if self.col_rut:
            mostrar_mensaje(f"Validación de RUT activada en columna '{self.col_rut}'.", "exito")
        else:
            mostrar_mensaje("No se validará RUT (columna no mapeada).", "advertencia")

        self.col_email = detectar_y_preguntar("Email", ["email", "correo"])
        if self.col_email:
            mostrar_mensaje(f"Validación de email activada en columna '{self.col_email}'.", "exito")
        else:
            mostrar_mensaje("No se validará email (columna no mapeada).", "advertencia")

        self.col_tel = detectar_y_preguntar("Teléfono", ["tel", "fono", "telefono"])
        if self.col_tel:
            mostrar_mensaje(f"Validación de teléfono activada en columna '{self.col_tel}'.", "exito")
        else:
            mostrar_mensaje("No se validará teléfono (columna no mapeada).", "advertencia")

    def buscar(self, df: pd.DataFrame, criterio: str) -> pd.DataFrame:
        if unidecode:
            crit = unidecode(criterio.strip().lower())
            def match(val):
                return crit in unidecode(str(val or '').lower())
        else:
            crit = criterio.strip().lower()
            def match(val):
                return crit in str(val or '').lower()
        mask = df.apply(lambda row: any(match(row[c]) for c in df.columns), axis=1)
        return df[mask]

    def paginar(self, df: pd.DataFrame, page_size: int = 20):
        total = len(df)
        if total == 0:
            mostrar_mensaje("No hay registros para mostrar.", "advertencia")
            return
        pages = (total + page_size - 1) // page_size
        for i in range(pages):
            start = i * page_size
            end = start + page_size
            subset = df.iloc[start:end]
            print(tabulate(subset, headers='keys', tablefmt='fancy_grid', showindex=False))
            mostrar_mensaje(f"Página {i+1}/{pages} ({start+1}-{min(end,total)} de {total})", "info")
            if i < pages - 1:
                cont = input("Presione Enter para continuar, 'q' para salir, 's' para cambiar tamaño página: ").strip().lower()
                if cont == 'q':
                    break
                elif cont == 's':
                    nuevo_size = input("Nuevo tamaño de página (número): ").strip()
                    if nuevo_size.isdigit() and int(nuevo_size) > 0:
                        page_size = int(nuevo_size)
                        pages = (total + page_size - 1) // page_size
                        i = -1  # reiniciar paginación
                        continue

    def menu_ayuda(self):
        print(Fore.BLUE + Style.BRIGHT + """
╔═══════════════════════╗
║       AYUDA RÁPIDA    ║
╠═══════════════════════╣
║ 1 / v : Ver registros paginados
║ 2 / b : Buscar registros por texto
║ 3 / a : Agregar nuevo registro
║ 4 / u : Actualizar registro existente
║ 5 / d : Eliminar registro
║ 6 / g : Guardar y salir
║ h / ? : Mostrar esta ayuda
║ q     : Salir sin guardar
╚═══════════════════════╝
""")

    def run_interactivo(self):
        print(Fore.BLUE + Style.BRIGHT + """
╔═════════════════════════════════════════════╗
║    SISTEMA DE GESTIÓN DE REGISTROS EXCEL   ║
╚═════════════════════════════════════════════╝
""")
        try:
            df = self.leer()
            if not self.columnas:
                self.columnas = df.columns.tolist()
            mostrar_mensaje("Archivo cargado correctamente.", "exito")
            self.mapear_columnas_clave(df)
            mostrar_mensaje(f"Columnas finales: {self.columnas}", "info")
            mostrar_mensaje(f"Total de registros: {len(df)}", "info")

            while True:
                print("\n" + Fore.MAGENTA + "="*60)
                print("📚 " + Style.BRIGHT + "MENÚ CRUD GENÉRICO".center(58))
                print(Fore.MAGENTA + "="*60)
                print(Fore.CYAN + "1️⃣  Ver registros (v)")
                print(Fore.CYAN + "2️⃣  Buscar registro (b)")
                print(Fore.CYAN + "3️⃣  Agregar nuevo registro (a)")
                print(Fore.CYAN + "4️⃣  Actualizar registro existente (u)")
                print(Fore.CYAN + "5️⃣  Eliminar registro (d)")
                print(Fore.CYAN + "6️⃣  Guardar y salir (g)")
                print(Fore.CYAN + "h️⃣  Ayuda (h/?))")
                print(Fore.CYAN + "q️⃣  Salir sin guardar (q)")
                print(Fore.MAGENTA + "-"*60)
                opcion = input("Seleccione opción (número o letra): ").strip().lower()

                if opcion in ['1', 'v']:
                    self.paginar(df)
                elif opcion in ['2', 'b']:
                    criterio = input("🔍 Ingrese término de búsqueda: ").strip()
                    if not criterio:
                        mostrar_mensaje("Debe ingresar un criterio de búsqueda.", "advertencia")
                    else:
                        filtrado = self.buscar(df, criterio)
                        if filtrado.empty:
                            mostrar_mensaje("No se encontraron coincidencias.", "advertencia")
                        else:
                            print(tabulate(filtrado, headers='keys', tablefmt='fancy_grid', showindex=False))
                elif opcion in ['3', 'a']:
                    nuevo = {}
                    mostrar_mensaje("📝 Ingrese los datos del nuevo registro:", "info")
                    for col in self.columnas:
                        if self.col_rut and col == self.col_rut:
                            valor = input_validado(f"{col}: ", self.validar_rut, "RUT inválido. Intenta nuevamente.", self.format_rut)
                        elif self.col_email and col == self.col_email:
                            valor = input_validado(f"{col} (opcional): ", self.validar_email, "Email inválido.", opcional=True)
                        elif self.col_tel and col == self.col_tel:
                            valor = input_validado(f"{col} (opcional): ", self.validar_telefono, "Teléfono inválido.", opcional=True)
                        else:
                            valor = input(f"{col}: ").strip()
                        nuevo[col] = valor
                    df = pd.concat([df, pd.DataFrame([nuevo])], ignore_index=True)
                    mostrar_mensaje("Registro agregado correctamente.", "exito")
                    logging.info(f"Agregado registro: {nuevo}")
                elif opcion in ['4', 'u']:
                    idx = None
                    if self.col_rut:
                        clave = input(f"✏️ Ingrese el {self.col_rut} del registro a actualizar: ").strip()
                        clave_norm = self.clean_rut(clave)
                        ser_norm = df[self.col_rut].fillna("").astype(str).apply(self.clean_rut)
                        matches = df[ser_norm == clave_norm]
                        if matches.empty:
                            mostrar_mensaje(f"No se encontró un registro con {self.col_rut} = {clave}.", "error")
                            continue
                        idx_list = matches.index.tolist()
                        if len(idx_list) > 1:
                            mostrar_mensaje("Hay múltiples coincidencias. Se mostrará la primera.", "advertencia")
                        idx = idx_list[0]
                    else:
                        mostrar_mensaje("No está configurado el campo RUT para búsqueda. Abortando actualización.", "error")
                        continue
                    mostrar_mensaje(f"Registro actual:\n{tabulate([df.loc[idx].to_dict()], headers='keys', tablefmt='fancy_grid')}", "info")
                    for col in self.columnas:
                        valor_actual = df.at[idx, col]
                        nuevo_valor = input(f"{col} [{valor_actual}]: ").strip()
                        if nuevo_valor != '':
                            if self.col_rut and col == self.col_rut:
                                if not self.validar_rut(nuevo_valor):
                                    mostrar_mensaje("RUT inválido. Se mantiene el valor anterior.", "advertencia")
                                    continue
                                nuevo_valor = self.format_rut(nuevo_valor)
                            elif self.col_email and col == self.col_email:
                                if not self.validar_email(nuevo_valor):
                                    mostrar_mensaje("Email inválido. Se mantiene el valor anterior.", "advertencia")
                                    continue
                            elif self.col_tel and col == self.col_tel:
                                if not self.validar_telefono(nuevo_valor):
                                    mostrar_mensaje("Teléfono inválido. Se mantiene el valor anterior.", "advertencia")
                                    continue
                            df.at[idx, col] = nuevo_valor
                    mostrar_mensaje("Registro actualizado correctamente.", "exito")
                    logging.info(f"Actualizado registro idx={idx}")
                elif opcion in ['5', 'd']:
                    if self.col_rut:
                        clave = input(f"🗑️ Ingrese el {self.col_rut} del registro a eliminar: ").strip()
                        clave_norm = self.clean_rut(clave)
                        ser_norm = df[self.col_rut].fillna("").astype(str).apply(self.clean_rut)
                        matches = df[ser_norm == clave_norm]
                        if matches.empty:
                            mostrar_mensaje(f"No se encontró un registro con {self.col_rut} = {clave}.", "error")
                            continue
                        idx_list = matches.index.tolist()
                        idx = idx_list[0]
                        mostrar_mensaje(f"Registro a eliminar:\n{tabulate([df.loc[idx].to_dict()], headers='keys', tablefmt='fancy_grid')}", "advertencia")
                        confirm = input("¿Confirmar eliminación? (s/n): ").strip().lower()
                        if confirm == 's':
                            df = df.drop(idx).reset_index(drop=True)
                            mostrar_mensaje("Registro eliminado.", "exito")
                            logging.info(f"Eliminado registro idx={idx}")
                        else:
                            mostrar_mensaje("Eliminación cancelada.", "info")
                    else:
                        mostrar_mensaje("No está configurado el campo RUT para búsqueda. Abortando eliminación.", "error")
                elif opcion in ['6', 'g']:
                    self.guardar(df)
                    mostrar_mensaje("Saliendo del sistema. ¡Hasta luego!", "info")
                    break
                elif opcion in ['h', '?']:
                    self.menu_ayuda()
                elif opcion == 'q':
                    confirmar = input("¿Salir sin guardar? (s/n): ").strip().lower()
                    if confirmar == 's':
                        mostrar_mensaje("Saliendo sin guardar. ¡Hasta pronto!", "advertencia")
                        break
                else:
                    mostrar_mensaje("Opción inválida. Escribe 'h' para ayuda.", "error")

        except KeyboardInterrupt:
            mostrar_mensaje("\nOperación cancelada por usuario.", "advertencia")
            sys.exit(0)


def main():
    mostrar_mensaje("Selecciona el archivo Excel para gestionar:", "info")
    archivo = explorar_directorios(Path.home())
    gestor = GestorDocentes(archivo)
    gestor.run_interactivo()


if __name__ == "__main__":
    main()
