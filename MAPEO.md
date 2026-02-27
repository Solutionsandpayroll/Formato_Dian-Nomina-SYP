# Documentación de Mapeo - Formato DIAN

## 📋 Descripción General

Este documento detalla cómo funciona el mapeo de datos desde los archivos Excel de entrada (Novasoft/Midasoft) hacia la plantilla estándar de la DIAN.

## 📁 Plantilla Base

La aplicación utiliza la plantilla **"Formato DIAN 2276.xlsx"** ubicada en la carpeta `public/` como base para generar el archivo resultante.

**Importante:** La plantilla se modifica directamente usando **ExcelJS**, preservando:
- ✅ Todos los formatos de celda (colores, bordes, fuentes)
- ✅ Altura de filas y ancho de columnas
- ✅ Fórmulas existentes
- ✅ Estilos y colores de fondo
- ✅ Validaciones de datos
- ✅ Configuraciones de impresión
- ✅ Imágenes y gráficos
- ✅ Formato condicional

Los datos del archivo de entrada se insertan en las celdas correspondientes sin afectar ningún formato de la plantilla.

## 🔄 Mapeo: NOVASOFT

### Reglas de Mapeo

#### Filas:
- **Archivo de entrada:** Los datos comienzan en la **fila 2** (la fila 1 es el encabezado)
- **Archivo de salida:** Los datos se escriben desde la **fila 3** de la plantilla DIAN

#### Columnas:
Cada columna del archivo de entrada se desplaza **una posición a la derecha** en el archivo de salida:

| Columna Entrada | → | Columna Salida |
|-----------------|---|----------------|
| A (índice 0)    | → | B (índice 1)   |
| B (índice 1)    | → | C (índice 2)   |
| C (índice 2)    | → | D (índice 3)   |
| D (índice 3)    | → | E (índice 4)   |
| ...             | → | ...            |
| AS (índice 44)  | → | AT (índice 45) |

### Rango de Columnas
- **Desde:** Columna A
- **Hasta:** Columna AS (45 columnas en total)

### Ejemplo Visual

```
ENTRADA (Novasoft):
Fila 1: [Encabezados]
Fila 2: [Dato A1, Dato B1, Dato C1, ...]
Fila 3: [Dato A2, Dato B2, Dato C2, ...]

SALIDA (Plantilla DIAN):
Fila 1: [Encabezados DIAN - se mantienen]
Fila 2: [Información DIAN - se mantiene]
Fila 3: [, Dato A1, Dato B1, Dato C1, ...]  ← Columna A vacía
Fila 4: [, Dato A2, Dato B2, Dato C2, ...]  ← Los datos empiezan en B
```

## 🔄 Mapeo: MIDASOFT

### Estado
⚠️ **Pendiente de implementación**

Para implementar el mapeo de Midasoft:
1. Abre [src/App.jsx](src/App.jsx)
2. Localiza la función `mapDataAccordingToSystem`
3. Completa el bloque `else if (system === 'Midasoft')`
4. Define las reglas específicas de mapeo

## 🛠️ Modificar el Mapeo

### Ubicación del Código
Archivo: `src/App.jsx`  
Función: `mapDataToTemplateExcelJS(inputData, worksheet, system)`  
Línea aproximada: ~130

### Cómo Funciona

El código usa **ExcelJS** para modificar directamente las celdas de la plantilla Excel:

```javascript
// Obtener la fila (ExcelJS usa 1-indexed: fila 1, 2, 3...)
const row = worksheet.getRow(outputRowNumber)

// Obtener la celda (columnas también 1-indexed: A=1, B=2, C=3...)
const cell = row.getCell(outputCol)

// Cambiar solo el valor, el formato se preserva automáticamente
cell.value = inputRow[j]

// Confirmar cambios
row.commit()
```

### Ventajas de ExcelJS

✅ **Preservación total de formatos:** Colores, bordes, fuentes, alturas, anchos  
✅ **Detección automática de tipos:** Números, texto, fechas, booleanos  
✅ **Gratuito y open source:** Sin limitaciones de la versión community  
✅ **API moderna y fácil de usar:** Promises, async/await  
✅ **Excelente documentación:** https://github.com/exceljs/exceljs

### Parámetros Configurables

**Importante:** ExcelJS usa **1-indexed** (comienza en 1, no en 0)
- Fila 1 = primera fila de Excel
- Columna 1 = columna A, columna 2 = columna B, etc.

```javascript
// Fila de inicio en el archivo de entrada (0-indexed para el array JS)
for (let i = 1; i < inputData.length; i++) {
  // i = 1 significa fila 2 del array (primera fila de datos)
  
  // Fila de destino en Excel (1-indexed)
  const outputRowNumber = i + 2 // Fila 3 en adelante
  
  // Columna de destino (1-indexed: A=1, B=2, C=3...)
  const outputCol = j + 2 // Columna B (2) en adelante
}
```

### Cambios Comunes

**Para cambiar la fila de inicio de lectura:**
```javascript
for (let i = 2; i < inputData.length; i++) {
  // Ahora empieza desde la fila 3 del archivo de entrada
```

**Para cambiar la fila de inicio de escritura:**
```javascript
const outputRowNumber = i + 3 // Ahora escribe desde la fila 4
```

**Para cambiar el desplazamiento de columnas:**
```javascript
const outputCol = j + 2 // Desplazar 2 columnas a la derecha
```

**Para mapeo sin desplazamiento (columna por columna):**
```javascript
const outputCol = j // Sin desplazamiento
```

**Para mapear a una columna específica:**
```javascript
// Siempre escribir en la columna D (índice 3)
const outputCol = 3
const cellAddress = columnToLetter(outputCol) + outputRowNumber
```

## 📝 Notas Importantes

1. **Índices en JavaScript:** Los arrays empiezan en 0
   - Fila 1 del Excel = índice 0
   - Fila 2 del Excel = índice 1
   - Columna A = índice 0
   - Columna B = índice 1

2. **Preservación de la Plantilla:** Las primeras filas de la plantilla DIAN (encabezados y formato) se mantienen intactas.

3. **Límite de Columnas:** El mapeo actual está limitado a 45 columnas (hasta AS). Para extender:
```javascript
for (let j = 0; j < inputRow.length && j <= 50; j++) {
  // Ahora mapea hasta la columna AY (índice 50)
```

## 🐛 Troubleshooting

### Los datos no aparecen en el archivo resultante
- Verifica que la fila de inicio sea correcta (`i = 1` para leer desde la fila 2)
- Asegúrate de que `outputRowNumber` apunte a la fila correcta (recuerda que Excel usa 1-indexed)
- Revisa la consola del navegador para ver posibles errores

### Los datos están en las columnas incorrectas
- Revisa el cálculo de `outputCol`: debe ser `j + 1` para desplazar una columna
- Verifica la función `columnToLetter()` - debe convertir correctamente (0→A, 1→B, 25→Z, 26→AA)

### Los números aparecen como texto
- La función `createCell()` detecta automáticamente tipos de datos
- Si un número se guarda como texto, revisa si tiene espacios o caracteres especiales

### Los formatos de la plantilla se pierden
- Esto NO debería ocurrir con el método actual
- El código modifica solo el valor de las celdas, preservando todos los formatos
- Si ocurre, verifica que la plantilla se esté cargando correctamente

### Faltan datos al final
- Aumenta el límite en `j <= 44` al número de columnas necesario
- Para más de 45 columnas: `j <= 50` (hasta columna AY)

## 📧 Soporte

Para cambios o consultas sobre el mapeo, revisar el código en:
- [src/App.jsx](../src/App.jsx) - Función `mapDataToTemplate`

### Ventajas del Método Actual

✅ **Preserva formatos:** Todos los estilos, colores y formatos de la plantilla se mantienen  
✅ **Mantiene fórmulas:** Las fórmulas existentes en la plantilla no se pierden  
✅ **Detección automática:** Los tipos de datos (número, texto) se detectan automáticamente  
✅ **Eficiente:** Solo modifica las celdas necesarias, no reconstruye todo el archivo  
✅ **Confiable:** Trabaja directamente con las estructuras de datos de Excel
