# Gestor Docentes Excel

**Descripción:**  
Este proyecto es un sistema interactivo de gestión de registros en archivos Excel, diseñado para facilitar el CRUD de datos de docentes con validaciones específicas para RUT chileno, emails y teléfonos. Implementa un menú en consola con paginación, búsqueda, respaldo y bloqueo para evitar conflictos de escritura.

---

## Características principales

- Exploración dinámica de directorios para seleccionar archivos Excel.
- Validación de RUT chileno con formato y dígito verificador.
- Validación de emails y teléfonos.
- Backup automático antes de guardar cambios.
- Manejo de bloqueo para evitar accesos simultáneos.
- Interfaz de usuario en consola con mensajes coloridos e intuitivos.
- Registro de eventos con archivo de log.
- Código modular y orientado a facilitar futuras extensiones.

---

## Tecnologías

- Python 3.x
- Pandas
- Colorama para consola
- Unidecode (opcional)

---

## Uso

Ejecutar el script principal:

```bash
python gestor_docentes.py
