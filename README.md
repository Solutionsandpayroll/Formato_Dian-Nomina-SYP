# Formato DIAN - Nómina

Aplicación web profesional para convertir archivos Excel de nómina de diferentes sistemas (Novasoft y Midasoft) al formato requerido por la DIAN.

## 🚀 Características

- **Interfaz Profesional y Corporativa**: Diseño moderno y limpio
- **Drag & Drop**: Arrastra archivos o selecciónalos con un clic
- **Selector de Sistema**: Elige entre Novasoft o Midasoft
- **Procesamiento de Excel**: Convierte archivos Excel automáticamente
- **Descarga Inmediata**: Obtén el archivo procesado al instante
- **Vista Previa**: Visualiza el archivo cargado antes de procesar
- **100% Frontend**: No requiere backend para el procesamiento básico

## �️ Tecnologías

- **React 18.2.0** - Framework principal
- **Vite 5.0.8** - Build tool y servidor de desarrollo
- **XLSX (SheetJS) 0.18.5** - Lectura de archivos Excel de entrada
- **ExcelJS 4.4.0** - Manipulación de plantilla Excel preservando formatos
- **file-saver 2.0.5** - Descarga de archivos en el navegador
- **CSS moderno** - Animaciones y diseño responsive

## �📋 Requisitos

- Node.js 16 o superior
- npm o yarn

## 🛠️ Instalación

1. Instala las dependencias:
```bash
npm install
```

## 🎯 Uso

1. Inicia el servidor de desarrollo:
```bash
npm run dev
```

2. Abre tu navegador en `http://localhost:5173`

3. Selecciona el sistema de origen (Novasoft o Midasoft)

4. Arrastra tu archivo Excel o haz clic para seleccionarlo

5. Haz clic en "Procesar y Descargar"

6. El archivo convertido se descargará automáticamente

## 🏗️ Estructura del Proyecto

```
Formato DIAN - Nomina/
├── public/               # Archivos públicos estáticos
│   ├── logo-syp.png     # Logo de la empresa
│   ├── Formato DIAN 2276.xlsx  # Plantilla DIAN para el mapeo
│   └── README.md        # Instrucciones para el logo
├── src/
│   ├── App.jsx          # Componente principal
│   ├── App.css          # Estilos de la aplicación
│   ├── main.jsx         # Punto de entrada
│   └── index.css        # Estilos globales
├── index.html           # HTML base
├── package.json         # Dependencias
└── vite.config.js       # Configuración de Vite
```

## ⚙️ Lógica de Mapeo

La aplicación trabaja **directamente sobre la plantilla DIAN** preservando todos sus formatos, fórmulas y estilos.

### 🎨 Preservación de Formatos con ExcelJS

La aplicación utiliza **ExcelJS** para garantizar que **todos los formatos de la plantilla se mantengan intactos**:

✅ **Colores de celdas y fondos** (incluyendo el color verde de la fila 2)  
✅ **Altura de filas y ancho de columnas** (preservando medidas específicas)  
✅ **Bordes y estilos de celda** (líneas, grosores, colores)  
✅ **Fórmulas** (se mantienen funcionales)  
✅ **Formatos de número** (decimales, moneda, porcentajes)  
✅ **Alineación y fuentes** (tamaños, tipos, negrita, cursiva)  
✅ **Validaciones de datos y formato condicional**  
✅ **Imágenes, gráficos y configuraciones de impresión**

**Importante:** Solo se modifican los **valores** de las celdas, nunca el formato. Esto garantiza que el archivo resultante mantiene exactamente la misma apariencia profesional que la plantilla original.

### Novasoft
El sistema lee los datos del archivo de entrada y los inserta en las celdas correspondientes de la plantilla:
- **Datos de entrada:** Empiezan en la fila 2
- **Datos de salida:** Se escriben desde la fila 3 de la plantilla
- **Mapeo de columnas:** Cada columna se desplaza una posición a la derecha
  - Columna A → Columna B
  - Columna B → Columna C
  - ... hasta Columna AS → Columna AT
- **Preservación:** Todos los formatos, fórmulas y estilos de la plantilla se mantienen intactos

### Midasoft
Pendiente de implementar.

### Documentación Detallada
Para más información sobre el mapeo, consulta [MAPEO.md](MAPEO.md)
## 🎨 Personalización del Logo

El logo actual es un placeholder. Para usar tu propio logo:

1. Coloca tu logo en la carpeta [public/](public/)
2. Nombra el archivo como `logo-syp.png`, `logo-syp.jpg` o `logo-syp.svg`
3. Si cambias el formato, actualiza la ruta en [src/App.jsx](src/App.jsx) línea ~141

Ver [public/README.md](public/README.md) para más detalles.

## 📦 Dependencias Principales

- **React**: Framework de UI
- **Vite**: Build tool y dev server
- **xlsx (SheetJS)**: Procesamiento de archivos Excel
- **file-saver**: Descarga de archivos

## 🔧 Próximos Pasos

Para implementar el mapeo de Midasoft:

1. Edita la función `mapDataToTemplate()` en [src/App.jsx](src/App.jsx)
2. Agrega el caso para 'Midasoft' con sus reglas específicas de mapeo
3. Sigue el mismo patrón usado en Novasoft

Para personalizar el mapeo de Novasoft:

1. Ajusta los índices de columnas y filas en [src/App.jsx](src/App.jsx) línea ~135
2. Modifica el desplazamiento de columnas según sea necesario
3. La función trabaja directamente sobre las celdas de la plantilla preservando todos los formatos

## 📝 Notas

- Los archivos deben ser formato .xlsx o .xls
- El procesamiento se realiza completamente en el navegador
- No se envían datos a ningún servidor

## 🎨 Personalización

Para cambiar el logo, edita el SVG en el componente Header en [src/App.jsx](src/App.jsx)

Para modificar los colores corporativos, ajusta las variables CSS en [src/App.css](src/App.css):
```css
:root {
  --primary-color: #2563eb;
  --primary-dark: #1e40af;
  /* ... más colores */
}
```

## 📄 Licencia

© 2026 Solutions & Payroll. Todos los derechos reservados.
