import * as XLSX from 'xlsx'

export async function parseExcelFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader()
    
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result)
        const workbook = XLSX.read(data, { type: 'array' })
        
        // Get the first sheet
        const firstSheetName = workbook.SheetNames[0]
        const worksheet = workbook.Sheets[firstSheetName]
        
        // Convert to JSON with header row
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { 
          header: 1,
          defval: '' // Default value for empty cells
        })
        
        // Find the header row (usually first row, but skip empty rows)
        let headerRowIndex = 0
        for (let i = 0; i < Math.min(5, jsonData.length); i++) {
          if (jsonData[i] && jsonData[i].some(cell => cell && String(cell).trim() !== '')) {
            headerRowIndex = i
            break
          }
        }
        
        const headers = jsonData[headerRowIndex] || []
        
        // Clean headers - remove any trailing colons and trim
        const cleanedHeaders = headers.map(h => {
          if (!h) return ''
          let cleaned = String(h).trim()
          // Keep the colon if it's part of the column name (like "Size MM:")
          return cleaned
        })
        
        // Extract data rows (skip header row and empty rows)
        const rows = jsonData.slice(headerRowIndex + 1).filter(row => 
          row && row.some(cell => cell !== '' && cell !== null && cell !== undefined)
        )
        
        // Map data to object format
        const mappedData = rows.map(row => {
          const obj = {}
          cleanedHeaders.forEach((header, index) => {
            if (header) {
              obj[header] = row[index] !== undefined && row[index] !== null ? row[index] : ''
            }
          })
          return obj
        })
        
        resolve({
          headers: cleanedHeaders,
          rows: mappedData,
          rawData: jsonData
        })
      } catch (error) {
        reject(new Error(`Failed to parse Excel file: ${error.message}`))
      }
    }
    
    reader.onerror = () => {
      reject(new Error('Failed to read file'))
    }
    
    reader.readAsArrayBuffer(file)
  })
}

