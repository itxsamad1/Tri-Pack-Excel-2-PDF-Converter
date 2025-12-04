import { useState } from 'react'
import { convertExcelToPdf } from './utils/pdfGenerator'
import { parseExcelFile } from './utils/excelParser'
import './App.css'

function App() {
  const [file, setFile] = useState(null)
  const [loading, setLoading] = useState(false)
  const [error, setError] = useState(null)

  const handleFileChange = (e) => {
    const selectedFile = e.target.files[0]
    if (selectedFile) {
      if (selectedFile.name.endsWith('.xlsx') || selectedFile.name.endsWith('.xls')) {
        setFile(selectedFile)
        setError(null)
      } else {
        setError('Please select a valid Excel file (.xlsx or .xls)')
        setFile(null)
      }
    }
  }

  const handleConvert = async () => {
    if (!file) {
      setError('Please select an Excel file first')
      return
    }

    setLoading(true)
    setError(null)

    try {
      const excelData = await parseExcelFile(file)
      console.log('Parsed Excel Data:', excelData)
      console.log('Headers:', excelData.headers)
      console.log('First Row:', excelData.rows[0])
      await convertExcelToPdf(excelData, file.name)
    } catch (err) {
      setError(`Error converting file: ${err.message}`)
      console.error(err)
    } finally {
      setLoading(false)
    }
  }

  return (
    <div className="app-container">
      <div className="card">
        <div className="logo-container">
          <img src="/tri-pack-logo.jpeg" alt="Tri-Pack Films Limited" className="logo" />
        </div>
        <h1>Excel to PDF Converter</h1>
        <p className="subtitle">Convert your Excel pallet tags to PDF format</p>
        
        <div className="upload-section">
          <label htmlFor="file-upload" className="upload-label">
            <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
              <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path>
              <polyline points="17 8 12 3 7 8"></polyline>
              <line x1="12" y1="3" x2="12" y2="15"></line>
            </svg>
            Choose Excel File
          </label>
          <input
            id="file-upload"
            type="file"
            accept=".xlsx,.xls"
            onChange={handleFileChange}
            className="file-input"
          />
          {file && (
            <div className="file-info">
              <span className="file-name">📄 {file.name}</span>
            </div>
          )}
        </div>

        {error && <div className="error-message">{error}</div>}

        <button
          onClick={handleConvert}
          disabled={!file || loading}
          className="convert-button"
        >
          {loading ? 'Converting...' : 'Convert to PDF'}
        </button>
      </div>
    </div>
  )
}

export default App

