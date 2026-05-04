# Editor de Horario - BUAP

Aplicación de escritorio para armar y visualizar horarios escolares. Permite registrar materias con profesor, salón, días y horas, asignarles colores, detectar choques de horario automáticamente y exportar el resultado en PDF o JPG.

---

## ¿Qué hace cada parte del código?

### `__init__` — Inicialización
Configura la ventana principal y declara las variables globales que usa toda la aplicación: la lista de materias registradas, los horarios temporales mientras se arma una materia, los días de la semana disponibles y el color por defecto.

### `setup_ui` — Interfaz gráfica
Construye toda la interfaz visual. Está dividida en dos paneles:
- **Panel izquierdo:** contiene los campos de registro (nombre, salón, profesor), el selector de día y hora, el botón de color, los botones de exportar y la lista de visibilidad de materias.
- **Panel derecho:** contiene el canvas (lienzo) donde se dibuja la tabla del horario.

### `crear_campo` — Campos de texto
Función auxiliar que crea un par etiqueta + campo de texto de forma reutilizable. Se usa para no repetir código al crear los campos de nombre, salón y profesor.

### `elegir_color` — Selector de color
Abre el selector de color del sistema operativo y guarda el color elegido. Ese color se usará como fondo de la materia en la tabla y en el archivo exportado.

### `agregar_horario_temporal` — Agregar días a una materia
Permite asignar múltiples días y rangos de hora a una misma materia antes de registrarla. Por ejemplo, una materia puede tener Lunes de 8 a 10 y Miércoles de 10 a 12. Cada combinación se guarda temporalmente hasta que se registra la materia.

### `validar_y_registrar` — Registrar materia
Valida que el nombre y al menos un horario estén completos, luego guarda la materia en la lista principal con todos sus datos (nombre, salón, profesor, horarios, color). Después limpia los campos y actualiza la tabla y la lista de visibilidad.

### `limpiar_campos` — Limpiar formulario
Borra todos los campos de texto y reinicia la lista de horarios temporales después de registrar una materia.

### `verificar_choque` — Detección de choques
Cada vez que se activa o desactiva una materia en la lista de visibilidad, esta función revisa si hay traslape de horario con alguna otra materia que ya esté visible. La condición de choque es que coincidan el día y que las horas se sobrepongan (no cuenta como choque si una termina exactamente cuando la otra empieza). Si hay choque, desmarca la materia automáticamente y muestra una advertencia.

### `actualizar_lista_toggles` — Lista de visibilidad
Reconstruye la lista de checkboxes del panel izquierdo cada vez que se registra una nueva materia. Cada checkbox permite mostrar u ocultar una materia en la tabla.

### `dibujar_cuadricula` — Tabla base
Dibuja la estructura fija del horario en el canvas: los encabezados de días (Lunes a Sábado) y las filas de horas (de 7:00 a 20:00). Esta cuadrícula siempre está visible, independientemente de qué materias estén activas.

### `actualizar_tabla` — Pintar materias en la tabla
Recorre todas las materias visibles y dibuja un rectángulo de color en las celdas correspondientes según su día y rango de horas. Encima del rectángulo escribe el nombre y el salón. Esta función se llama cada vez que cambia la visibilidad de alguna materia.

### `exportar_proceso_completo` — Exportar a PDF o JPG
Es la función más compleja. Su flujo es el siguiente:

1. Recopila solo las materias visibles y calcula el rango mínimo de días y horas que ocupan.
2. Crea un archivo Excel (`.xlsx`) con `openpyxl`, aplicando colores, bordes, tamaños de columna y configuración de página.
3. Abre ese archivo con Excel real (a través de `win32com`) para aprovechar su motor de renderizado.
4. **Si es PDF:** exporta directamente con el método `ExportAsFixedFormat` de Excel, con orientación horizontal y ajuste automático al ancho de página.
5. **Si es JPG:** copia el rango usado como imagen al portapapeles, lo captura con `ImageGrab.grabclipboard()` de Pillow y lo guarda como archivo `.jpg`.
6. Cierra Excel y elimina el archivo temporal.

---

## Dependencias

Instala todo con un solo comando:

```bash
pip install openpyxl pillow pywin32
```

| Librería | Uso |
|---|---|
| `tkinter` | Interfaz gráfica (incluida con Python, no se instala) |
| `openpyxl` | Genera el archivo Excel que se usa para exportar |
| `Pillow` | Captura la imagen del portapapeles al exportar JPG |
| `pywin32` | Controla Excel desde Python para generar el PDF/JPG |

> La PC donde se use la aplicación necesita tener **Microsoft Excel instalado**.

---

## Convertir a .exe

### 1. Instalar PyInstaller
```bash
pip install pyinstaller
```

### 2. Ir a la carpeta del proyecto
```bash
cd C:\ruta\de\tu\proyecto
```

### 3. Ejecutar este comando
```bash
pyinstaller --onefile --windowed --name "EditorHorario" --hidden-import win32com --hidden-import win32com.client --hidden-import win32com.shell --hidden-import pythoncom --hidden-import pywintypes editor_de_horario.py
```

### 4. Buscar el ejecutable

Cuando termine, el archivo estará en:
```
dist/
└── EditorHorario.exe
```

Ese `.exe` es el único archivo que necesitas para distribuir la aplicación.
