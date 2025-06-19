# Gestor Interactivo de Registros Excel para Docentes

---

## 📋 Descripción

Este proyecto es un **sistema interactivo de gestión de registros en archivos Excel** diseñado para simplificar la administración de datos de docentes. Ofrece un CRUD completo con validaciones específicas para el contexto chileno, incluyendo:

- Validación y formateo de RUT con dígito verificador.
- Validación de emails y números de teléfono.
- Navegación dinámica para seleccionar archivos Excel.
- Sistema de backup automático y manejo de bloqueo para evitar conflictos de escritura concurrentes.
- Menú en consola con paginación, búsqueda avanzada y edición intuitiva.
- Registro de eventos detallado para auditoría y seguimiento.

---

## 🚀 Características destacadas

- **Interfaz amigable:** Colores y símbolos en consola para facilitar la experiencia de usuario.
- **Flexibilidad:** Detecta columnas automáticamente y permite mapear campos clave para validaciones.
- **Seguridad:** Bloqueo de archivo durante escritura para evitar corrupciones.
- **Portabilidad:** Funciona en cualquier sistema con Python 3 y dependencias básicas.
- **Extensible:** Código modular y preparado para futuras integraciones con bases de datos o interfaces gráficas.

---

## 🛠 Tecnologías y dependencias

- **Python 3.x**
- [pandas](https://pandas.pydata.org/) para manipulación de datos
- [colorama](https://pypi.org/project/colorama/) para colores en consola
- [unidecode](https://pypi.org/project/Unidecode/) (opcional) para normalización de texto
- [tabulate](https://pypi.org/project/tabulate/) para presentación tabular

---
