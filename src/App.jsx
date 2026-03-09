import { useState, useEffect } from 'react'
import * as XLSX from 'xlsx'
import ExcelJS from 'exceljs'
import { saveAs } from 'file-saver'
import './App.css'

function App() {
  const [selectedSystem, setSelectedSystem] = useState('Novasoft')
  const [uploadedFile, setUploadedFile] = useState(null)
  const [fileName, setFileName] = useState('')
  const [customFileName, setCustomFileName] = useState('Cliente_2026_Medios Magneticos')
  const [isProcessing, setIsProcessing] = useState(false)
  const [dragActive, setDragActive] = useState(false)
  const [showModal, setShowModal] = useState(false)
  const [modalMessage, setModalMessage] = useState('')
  const [modalType, setModalType] = useState('success') // 'success' o 'error'
  const [isHelpExpanded, setIsHelpExpanded] = useState(false)

  const showNotification = (message, type = 'success') => {
    setModalMessage(message)
    setModalType(type)
    setShowModal(true)
  }

  const closeModal = () => {
    setShowModal(false)
  }

  // Auto-cerrar modal de éxito después de 3 segundos
  useEffect(() => {
    if (showModal && modalType === 'success') {
      const timer = setTimeout(() => {
        closeModal()
      }, 3000)
      return () => clearTimeout(timer)
    }
  }, [showModal, modalType])

  const handleSystemChange = (e) => {
    setSelectedSystem(e.target.value)
  }

  const handleFileUpload = (e) => {
    const file = e.target.files[0]
    if (file) {
      validateAndSetFile(file)
    }
  }

  const handleDrag = (e) => {
    e.preventDefault()
    e.stopPropagation()
    if (e.type === "dragenter" || e.type === "dragover") {
      setDragActive(true)
    } else if (e.type === "dragleave") {
      setDragActive(false)
    }
  }

  const handleDrop = (e) => {
    e.preventDefault()
    e.stopPropagation()
    setDragActive(false)
    
    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      validateAndSetFile(e.dataTransfer.files[0])
    }
  }

  const validateAndSetFile = (file) => {
    const validTypes = [
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'application/vnd.ms-excel'
    ]
    
    if (validTypes.includes(file.type) || file.name.endsWith('.xlsx') || file.name.endsWith('.xls')) {
      setUploadedFile(file)
      setFileName(file.name)
    } else {
      showNotification('Por favor, sube un archivo Excel válido (.xlsx o .xls)', 'error')
    }
  }

  const processExcel = async () => {
    if (!uploadedFile) {
      showNotification('Por favor, selecciona un archivo Excel primero', 'error')
      return
    }

    setIsProcessing(true)

    try {
      // Leer archivo de entrada con XLSX (para obtener datos rápidamente)
      const inputData = await uploadedFile.arrayBuffer()
      const inputWorkbook = XLSX.read(inputData, { type: 'array' })
      const inputSheetName = inputWorkbook.SheetNames[0]
      const inputWorksheet = inputWorkbook.Sheets[inputSheetName]
      // raw: false mantiene el formato original de las celdas (preserva "001" como texto)
      const inputJsonData = XLSX.utils.sheet_to_json(inputWorksheet, { header: 1, raw: false, defval: '' })

      // Cargar plantilla DIAN con ExcelJS (preserva TODOS los formatos)
      const templateResponse = await fetch('/Formato DIAN 2276.xlsx')
      if (!templateResponse.ok) {
        throw new Error('No se pudo cargar la plantilla DIAN')
      }
      const templateData = await templateResponse.arrayBuffer()
      
      // Usar ExcelJS para manipular el archivo preservando formatos
      const workbook = new ExcelJS.Workbook()
      await workbook.xlsx.load(templateData)
      
      const worksheet = workbook.worksheets[0]

      // Aplicar mapeo según el sistema - modificando directamente la plantilla
      await mapDataToTemplateExcelJS(inputJsonData, worksheet, selectedSystem)

      // Generar y descargar archivo con ExcelJS (preserva TODO)
      const buffer = await workbook.xlsx.writeBuffer()
      const blob = new Blob([buffer], { 
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
      })
      
      // Usar nombre personalizado si existe, sino usar el formato por defecto
      let downloadFileName
      if (customFileName && customFileName.trim() !== '') {
        // Asegurar que termine en .xlsx
        downloadFileName = customFileName.trim().endsWith('.xlsx') 
          ? customFileName.trim() 
          : `${customFileName.trim()}.xlsx`
      } else {
        downloadFileName = 'Cliente_2026_Medios Magneticos.xlsx'
      }
      saveAs(blob, downloadFileName)

      showNotification('¡Archivo procesado y descargado exitosamente!', 'success')
    } catch (error) {
      console.error('Error al procesar el archivo:', error)
      showNotification(`Error al procesar el archivo: ${error.message}`, 'error')
    } finally {
      setIsProcessing(false)
    }
  }

  const mapDataToTemplateExcelJS = async (inputData, worksheet, system) => {
    // Función auxiliar para convertir número de columna a letra Excel (1=A, 2=B, ..., 13=M, 46=AT)
    const getColumnLetter = (columnNumber) => {
      let letter = ''
      while (columnNumber > 0) {
        const remainder = (columnNumber - 1) % 26
        letter = String.fromCharCode(65 + remainder) + letter
        columnNumber = Math.floor((columnNumber - 1) / 26)
      }
      return letter
    }

    if (system === 'Novasoft') {
      // Mapeo para Novasoft:
      // - Datos de entrada empiezan en fila 2 (índice 1)
      // - Datos de salida empiezan en fila 3
      // - Columna A (0) del input → Columna B (2) del output (ExcelJS usa 1-indexed)
      // - Columna B (1) del input → Columna C (3) del output
      // - ... hasta Columna AS (44) del input → Columna AT (46) del output
      
      // Procesar cada fila de datos del archivo de entrada (desde la fila 2)
      for (let i = 1; i < inputData.length; i++) {
        const inputRow = inputData[i]
        const outputRowNumber = i + 2 // Fila 3 en adelante (Excel/ExcelJS usa 1-indexed)
        
        // Obtener o crear la fila en la hoja
        const row = worksheet.getRow(outputRowNumber)
        
        // Mapear cada columna (desplazar una posición a la derecha)
        // Columna A (0) → B (2), B (1) → C (3), etc. (ExcelJS usa 1-indexed)
        for (let j = 0; j < inputRow.length && j <= 44; j++) { // Hasta columna AS (índice 44)
          const outputCol = j + 2 // Desplazar una columna a la derecha (B es columna 2 en ExcelJS)
          const cell = row.getCell(outputCol)
          
          // Solo cambiar el valor, ExcelJS preserva automáticamente el formato de la celda
          const value = inputRow[j]
          
          if (value != null && value !== '') {
            // Columnas M a AO (13-41) deben ser números para las fórmulas de suma
            if (outputCol >= 13 && outputCol <= 41) {
              // Convertir a número si es posible
              if (typeof value === 'number') {
                cell.value = value
              } else if (typeof value === 'string' && !isNaN(value) && value.trim() !== '') {
                cell.value = parseFloat(value)
              } else {
                cell.value = 0 // Si no se puede convertir, usar 0
              }
            } else {
              // Para otras columnas, mantener el valor tal cual (preserva "001" como texto)
              cell.value = value
            }
          } else {
            cell.value = ''
          }
        }
        
        // Confirmar cambios en la fila
        row.commit()
      }

      // Agregar fila de totales con fórmulas de suma (columnas M a AO)
      // Calcular la última fila de datos y la fila donde irán las sumas
      const lastDataRow = inputData.length + 1 // Última fila con datos
      const sumRowNumber = lastDataRow + 1 // Fila siguiente para las sumas
      
      const sumRow = worksheet.getRow(sumRowNumber)
      
      // Columnas M (13) a AO (41) deben tener fórmulas de suma
      for (let col = 13; col <= 41; col++) {
        const columnLetter = getColumnLetter(col)
        const cell = sumRow.getCell(col)
        
        // Crear fórmula de suma: =SUM(M3:M{lastDataRow})
        cell.value = {
          formula: `SUM(${columnLetter}3:${columnLetter}${lastDataRow})`,
          result: 0 // Valor inicial, Excel lo calculará al abrir
        }
      }
      
      sumRow.commit()
    }
  }

  const mapDataToTemplate = (inputData, templateWorksheet, system) => {
    // Función auxiliar para convertir número de columna a letra (0 -> A, 1 -> B, etc.)
    const columnToLetter = (col) => {
      let letter = ''
      while (col >= 0) {
        letter = String.fromCharCode((col % 26) + 65) + letter
        col = Math.floor(col / 26) - 1
      }
      return letter
    }
    
    // Función para determinar el tipo de celda y formatear el valor
    const createCell = (value, existingCell) => {
      // Si hay una celda existente, preservar sus propiedades de formato
      const cell = existingCell ? { ...existingCell } : {}
      
      if (value == null || value === '') {
        cell.t = 's'
        cell.v = ''
        return cell
      }
      
      // Si es un número
      if (typeof value === 'number') {
        cell.t = 'n'
        cell.v = value
        return cell
      }
      
      // Si es texto que representa un número
      if (typeof value === 'string' && !isNaN(value) && value.trim() !== '') {
        cell.t = 'n'
        cell.v = parseFloat(value)
        return cell
      }
      
      // Si es booleano
      if (typeof value === 'boolean') {
        cell.t = 'b'
        cell.v = value
        return cell
      }
      
      // Por defecto, texto
      cell.t = 's'
      cell.v = value.toString()
      return cell
    }
    
    if (system === 'Novasoft') {
      // Mapeo para Novasoft:
      // - Datos de entrada empiezan en fila 2 (índice 1)
      // - Datos de salida empiezan en fila 3 (índice 2)
      // - Columna A (0) del input → Columna B (1) del output
      // - Columna B (1) del input → Columna C (2) del output
      // - ... hasta Columna AS (44) del input → Columna AT (45) del output
      
      // Procesar cada fila de datos del archivo de entrada (desde la fila 2)
      for (let i = 1; i < inputData.length; i++) {
        const inputRow = inputData[i]
        const outputRowNumber = i + 2 // Fila 3 en adelante (Excel usa 1-indexed)
        
        // Mapear cada columna (desplazar una posición a la derecha)
        // Columna A (0) → B (1), B (1) → C (2), etc.
        for (let j = 0; j < inputRow.length && j <= 44; j++) { // Hasta columna AS (índice 44)
          const outputCol = j + 1 // Desplazar una columna a la derecha
          const cellAddress = columnToLetter(outputCol) + outputRowNumber
          
          // Obtener la celda existente (si existe) para preservar formatos
          const existingCell = templateWorksheet[cellAddress]
          
          // Escribir el valor en la celda correspondiente de la plantilla
          templateWorksheet[cellAddress] = createCell(inputRow[j], existingCell)
        }
      }
    }
  }

  const resetForm = () => {
    setUploadedFile(null)
    setFileName('')
    const fileInput = document.getElementById('fileInput')
    if (fileInput) fileInput.value = ''
  }

  const formatFileSize = (bytes) => {
    if (bytes === 0) return '0 Bytes'
    const k = 1024
    const sizes = ['Bytes', 'KB', 'MB']
    const i = Math.floor(Math.log(bytes) / Math.log(k))
    return Math.round(bytes / Math.pow(k, i) * 100) / 100 + ' ' + sizes[i]
  }

  return (
    <div className="app">
      <header className="header">
        <div className="container">
          <div className="header-content">
            <div className="logo-container">
              <div className="logo">
                <img 
                  src="/Logo syp.png" 
                  alt="Solutions & Payroll Logo" 
                  width="60" 
                  height="60"
                />
              </div>
              <div className="header-text">
                <h1>Solutions & Payroll</h1>
                <p className="subtitle">Medios Magneticos - Nómina</p>
              </div>
            </div>
            <div className="welcome-box">
              <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                <path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"/>
                <circle cx="12" cy="7" r="4"/>
              </svg>
              <span>Bienvenido, Usuario de Nómina</span>
            </div>
          </div>
        </div>
      </header>

      <main className="main-content">
        <div className="container">
          {/* Sección de ayuda colapsable */}
          <div className="help-section">
            <button 
              className="help-toggle"
              onClick={() => setIsHelpExpanded(!isHelpExpanded)}
              aria-expanded={isHelpExpanded}
            >
              <div className="help-toggle-header">
                <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                  <circle cx="12" cy="12" r="10"/>
                  <line x1="12" y1="16" x2="12" y2="12"/>
                  <line x1="12" y1="8" x2="12.01" y2="8"/>
                </svg>
                <span>¿Cómo usar esta aplicación?</span>
              </div>
              <svg 
                className={`chevron ${isHelpExpanded ? 'expanded' : ''}`}
                width="20" 
                height="20" 
                viewBox="0 0 24 24" 
                fill="none" 
                stroke="currentColor" 
                strokeWidth="2"
              >
                <polyline points="6 9 12 15 18 9"/>
              </svg>
            </button>
            <div className={`help-content ${isHelpExpanded ? 'expanded' : ''}`}>
              <ol className="help-list">
                <li>
                  <span className="step-number">1</span>
                  <div>
                    <strong>Ingresa el nombre del archivo (opcional)</strong>
                    <p>Puedes personalizar el nombre del archivo resultante</p>
                  </div>
                </li>
                <li>
                  <span className="step-number">2</span>
                  <div>
                    <strong>Carga tu archivo Excel</strong>
                    <p>Arrastra o selecciona el archivo con los datos que vienen de Novasoft</p>
                  </div>
                </li>
                <li>
                  <span className="step-number">3</span>
                  <div>
                    <strong>Procesa y descarga</strong>
                    <p>El sistema convertirá automáticamente al formato DIAN</p>
                  </div>
                </li>
              </ol>
            </div>
          </div>

          <div className="card">
            <div className="card-header">
              <h2>Convertidor de Formato</h2>
              <p className="description">
                Transforma tus archivos de nómina al formato requerido por la DIAN
              </p>
            </div>

            <div className="card-body">
              <div className="form-section">
                <div className="form-group">
                  <label htmlFor="systemSelect" className="label">
                    <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                      <rect x="3" y="3" width="18" height="18" rx="2"/>
                      <path d="M9 3v18M15 3v18M3 9h18M3 15h18"/>
                    </svg>
                    Sistema de Origen
                  </label>
                  <select 
                    id="systemSelect" 
                    value={selectedSystem} 
                    onChange={handleSystemChange}
                    className="select-input"
                  >
                    <option value="Novasoft">Novasoft</option>
                  </select>
                </div>

                <div className="form-group">
                  <label htmlFor="fileNameInput" className="label">
                    <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                      <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
                      <polyline points="14 2 14 8 20 8"/>
                    </svg>
                    Nombre del Archivo Final
                  </label>
                  <input
                    type="text"
                    id="fileNameInput"
                    value={customFileName}
                    onChange={(e) => setCustomFileName(e.target.value)}
                    placeholder="Nombre del archivo"
                    className="select-input"
                  />
                  <p className="hint">Si no especificas un nombre, se generará con el nombre "Cliente_2026_Medios Magneticos"</p>
                </div>

                <div className="form-group">
                  <label className="label">
                    <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                      <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
                      <polyline points="14 2 14 8 20 8"/>
                      <line x1="12" y1="18" x2="12" y2="12"/>
                      <line x1="9" y1="15" x2="15" y2="15"/>
                    </svg>
                    Archivo Excel
                  </label>
                  
                  <div 
                    className={`drop-zone ${dragActive ? 'drag-active' : ''} ${fileName ? 'has-file' : ''}`}
                    onDragEnter={handleDrag}
                    onDragLeave={handleDrag}
                    onDragOver={handleDrag}
                    onDrop={handleDrop}
                  >
                    <input
                      type="file"
                      id="fileInput"
                      accept=".xlsx, .xls"
                      onChange={handleFileUpload}
                      className="file-input"
                    />
                    
                    {!fileName ? (
                      <label htmlFor="fileInput" className="drop-zone-content">
                        <svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5">
                          <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
                          <polyline points="17 8 12 3 7 8"/>
                          <line x1="12" y1="3" x2="12" y2="15"/>
                        </svg>
                        <div className="drop-zone-text">
                          <p className="drop-zone-title">Arrastra tu archivo aquí</p>
                          <p className="drop-zone-subtitle">o haz clic para seleccionar</p>
                          <p className="drop-zone-hint">Formatos soportados: .xlsx, .xls</p>
                        </div>
                      </label>
                    ) : (
                      <div className="file-preview">
                        <div className="file-icon">
                          <svg width="40" height="40" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                            <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
                            <polyline points="14 2 14 8 20 8"/>
                            <line x1="9" y1="15" x2="15" y2="15"/>
                            <line x1="12" y1="18" x2="12" y2="12"/>
                          </svg>
                        </div>
                        <div className="file-details">
                          <p className="file-name">{fileName}</p>
                          <p className="file-size">
                            {uploadedFile && formatFileSize(uploadedFile.size)}
                          </p>
                        </div>
                        <button onClick={resetForm} className="btn-remove" title="Eliminar archivo">
                          <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                            <line x1="18" y1="6" x2="6" y2="18"/>
                            <line x1="6" y1="6" x2="18" y2="18"/>
                          </svg>
                        </button>
                      </div>
                    )}
                  </div>
                </div>

                <button 
                  onClick={processExcel} 
                  disabled={!uploadedFile || isProcessing}
                  className="btn-primary"
                >
                  {isProcessing ? (
                    <>
                      <svg className="spinner" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                        <circle cx="12" cy="12" r="10"/>
                      </svg>
                      Procesando...
                    </>
                  ) : (
                    <>
                      <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                        <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
                        <polyline points="7 10 12 15 17 10"/>
                        <line x1="12" y1="15" x2="12" y2="3"/>
                      </svg>
                      Procesar y Descargar
                    </>
                  )}
                </button>
              </div>
            </div>
          </div>
        </div>
      </main>

      <footer className="footer">
        <div className="container">
          <p>&copy; {new Date().getFullYear()} Solutions & Payroll. Todos los derechos reservados.</p>
        </div>
      </footer>

      {/* Modal de Notificación */}
      {showModal && (
        <div className="modal-overlay" onClick={closeModal}>
          <div className={`modal-content ${modalType}`} onClick={(e) => e.stopPropagation()}>
            <div className="modal-icon">
              {modalType === 'success' ? (
                <svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                  <path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/>
                  <polyline points="22 4 12 14.01 9 11.01"/>
                </svg>
              ) : (
                <svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                  <circle cx="12" cy="12" r="10"/>
                  <line x1="15" y1="9" x2="9" y2="15"/>
                  <line x1="9" y1="9" x2="15" y2="15"/>
                </svg>
              )}
            </div>
            <h3 className="modal-title">
              {modalType === 'success' ? '¡Éxito!' : 'Error'}
            </h3>
            <p className="modal-message">{modalMessage}</p>
            <button onClick={closeModal} className="modal-button">
              Aceptar
            </button>
            {modalType === 'success' && (
              <div className="modal-progress">
                <div className="modal-progress-bar"></div>
              </div>
            )}
          </div>
        </div>
      )}
    </div>
  )
}

export default App