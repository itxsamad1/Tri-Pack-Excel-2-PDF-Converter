import jsPDF from 'jspdf'

// Helper function to get value from Excel data (case-insensitive and flexible matching)
function getValue(data, keys, defaultValue = '') {
  if (!data || typeof data !== 'object') return defaultValue
  
  // keys can be a string or array of possible key names
  const keyArray = Array.isArray(keys) ? keys : [keys]
  
  for (const key of keyArray) {
    const normalizedKey = key.toLowerCase().trim()
    
    // Try exact match first
    if (data[key] !== undefined && data[key] !== null && data[key] !== '') {
      return String(data[key])
    }
    
    // Try case-insensitive match (with and without trailing colon)
    const foundKey = Object.keys(data).find(k => {
      const normalizedK = k.toLowerCase().trim()
      // Match with or without trailing colon
      return normalizedK === normalizedKey || 
             normalizedK === normalizedKey + ':' ||
             normalizedK.replace(':', '') === normalizedKey.replace(':', '')
    })
    
    if (foundKey && data[foundKey] !== undefined && data[foundKey] !== null && data[foundKey] !== '') {
      return String(data[foundKey])
    }
  }
  
  return defaultValue
}

// Parse address string into customer name, address, and country
function parseAddress(addressStr) {
  if (!addressStr) return { customer: '', address: '', country: 'Spain' }
  
  const parts = addressStr.split(',').map(p => p.trim()).filter(p => p)
  
  if (parts.length >= 3) {
    return {
      customer: parts[0] || '',
      address: parts.slice(1, -1).join(', ') || '',
      country: parts[parts.length - 1] || 'Spain'
    }
  } else if (parts.length === 2) {
    return {
      customer: parts[0] || '',
      address: '',
      country: parts[1] || 'Spain'
    }
  } else {
    return {
      customer: parts[0] || '',
      address: '',
      country: 'Spain'
    }
  }
}

// Extract container number from filename or data
function extractContainerNumber(filename, data) {
  // Try to get from data first (column "CONT #03" or similar)
  const contNum = getValue(data, [
    'CONT #03',
    'CONT #',
    'CONT#',
    'CONT #',
    'Container Number',
    'CONT'
  ])
  
  if (contNum) {
    // Extract number from "CONT #03" format
    const match = String(contNum).match(/(\d+)/)
    if (match) {
      return match[1].padStart(2, '0')
    }
    return String(contNum).padStart(2, '0')
  }
  
  // Try to extract from filename (e.g., "TAG - QUIMINDROGA- CONT # 03.xlsx" -> "03")
  const match = filename.match(/CONT\s*#\s*(\d+)/i)
  if (match) {
    return match[1].padStart(2, '0')
  }
  
  return '01' // Default
}

// Load image and convert to base64
function loadImageAsBase64(url) {
  return new Promise((resolve, reject) => {
    const img = new Image()
    img.crossOrigin = 'anonymous'
    
    img.onload = () => {
      try {
        const canvas = document.createElement('canvas')
        canvas.width = img.width
        canvas.height = img.height
        const ctx = canvas.getContext('2d')
        ctx.drawImage(img, 0, 0)
        const base64 = canvas.toDataURL('image/jpeg')
        resolve(base64)
      } catch (e) {
        console.error('Error converting image to base64:', e)
        reject(e)
      }
    }
    
    img.onerror = (error) => {
      console.error('Error loading image:', url, error)
      reject(new Error(`Failed to load image from ${url}`))
    }
    
    // Try the provided URL
    img.src = url
  })
}

// Helper function to generate a single page for a row of data
async function generatePageForRow(doc, data, filename, pageWidth, pageHeight, margin, contentWidth, leftCol, rightCol, colWidth, labelValueGap, logoBase64, logoWidth, logoHeight) {
  let yPosition = margin
  const headerY = yPosition
  
  // Add logo on the left side
  if (logoBase64 && logoWidth && logoHeight) {
    try {
      doc.addImage(logoBase64, 'JPEG', leftCol, yPosition, logoWidth, logoHeight)
    } catch (error) {
      console.warn('Could not add logo to page:', error)
    }
  }
  
  // Set font - using default sans-serif
  doc.setFont('helvetica')
  
  // Title: PALLET TAG (centered) - slightly bigger
  doc.setFontSize(13)
  doc.setFont('helvetica', 'bold')
  const titleText = 'PALLET TAG'
  const titleWidth = doc.getTextWidth(titleText)
  doc.text(titleText, (pageWidth - titleWidth) / 2, headerY + 3)
  
  // Container Number: CONT # XX (centered, below title) - slightly bigger
  const containerNum = extractContainerNumber(filename, data)
  doc.setFontSize(11)
  doc.setFont('helvetica', 'bold')
  const contText = `CONT # ${containerNum}`
  const contWidth = doc.getTextWidth(contText)
  doc.text(contText, (pageWidth - contWidth) / 2, headerY + 7)
  
  // Start content below header - moved a little lower
  yPosition = headerY + 13
  
  // Set default font size for body - adjusted for landscape 4x6 format - slightly bigger
  doc.setFontSize(7.5)
  doc.setFont('helvetica', 'normal')
  
  // Note: rightCol, colWidth, and labelValueGap are already passed as parameters
  
  // Left column: LC / PO #
  doc.setFont('helvetica', 'bold')
  doc.text('LC / PO #', leftCol, yPosition)
  const lcPoValue = getValue(data, ['LC / PO #', 'LC/PO#', 'LC/PO', 'PO#', 'LC / PO'])
  doc.setFont('helvetica', 'normal')
  doc.text(lcPoValue, leftCol + labelValueGap, yPosition)
  yPosition += 4
  
  // Parse address field to get customer, address, and country
  const addressStr = getValue(data, ['ADDRESS', 'Address', 'Address Line 1'])
  const addressParts = parseAddress(addressStr)
  
  // Customer Name (left column)
  doc.setFont('helvetica', 'bold')
  const customerName = addressParts.customer || getValue(data, ['Customer', 'Customer Name', 'QUIMIDROGA'])
  doc.text(customerName, leftCol, yPosition)
  yPosition += 4
  
  // Address Line 1 (left column)
  doc.setFont('helvetica', 'bold')
  const address1 = addressParts.address || getValue(data, ['Address Line 1', 'C/TUSET 26 08006, 08006 BARCELONA,'])
  const addressLines = doc.splitTextToSize(address1, colWidth)
  doc.text(addressLines, leftCol, yPosition)
  yPosition += Math.max(addressLines.length * 3.5, 3.5)
  
  // Country (left column)
  doc.setFont('helvetica', 'bold')
  const country = addressParts.country || getValue(data, ['Country', 'Spain'])
  doc.text(country, leftCol, yPosition)
  yPosition += 4
  
  // PROFORMA INVOICE NUMBER (left column, below address) - make label bold, value normal
  doc.setFont('helvetica', 'bold')
  const invoiceLabel = 'PROFORMA INVOICE NUMBER:'
  const invoiceNum = getValue(data, ['PROFORMA INVOICE NUMBER:', 'PROFORMA INVOICE NUMBER', 'Invoice Number', 'Invoice'])
  doc.text(invoiceLabel, leftCol, yPosition)
  doc.setFont('helvetica', 'normal')
  doc.text(invoiceNum, leftCol + doc.getTextWidth(invoiceLabel) + 2, yPosition)
  yPosition += 4.5
  
  // FILM DESCP (left column, full width) - remove semicolon
  doc.setFont('helvetica', 'bold')
  doc.text('FILM DESCP:', leftCol, yPosition)
  const filmDesc = getValue(data, ['Film', 'FILM DESCP', 'Film Description', 'Description'])
  doc.setFont('helvetica', 'normal')
  // Handle long descriptions with word wrap
  const filmDescLines = doc.splitTextToSize(filmDesc, contentWidth - 20)
  doc.text(filmDescLines, leftCol + labelValueGap, yPosition)
  yPosition += Math.max(filmDescLines.length * 3.5, 4)
  
  // SIZE MM (left column)
  doc.setFont('helvetica', 'bold')
  doc.text('SIZE MM:', leftCol, yPosition)
  const sizeMm = getValue(data, ['Size MM:', 'SIZE MM', 'Size', 'Size MM'])
  doc.setFont('helvetica', 'normal')
  doc.text(sizeMm, leftCol + labelValueGap, yPosition)
  
  // No. OF REELS / PALLET (right column, aligned with SIZE MM)
  doc.setFont('helvetica', 'bold')
  const reelsLabel = 'No. OF REELS / PALLET'
  doc.text(reelsLabel, rightCol, yPosition)
  const numReels = getValue(data, ['NO. Of Reels / Pallet', 'NO. Of Reels / Pallet:', 'No. OF REELS / PALLET', 'Reels', 'Number of Reels'])
  doc.setFont('helvetica', 'normal')
  doc.text(numReels, rightCol + doc.getTextWidth(reelsLabel) + 2, yPosition)
  yPosition += 4
  
  // NET WEIGHT (PALLET) - left column
  doc.setFont('helvetica', 'bold')
  const netWeightLabel = 'NET WEIGHT (PALLET):'
  doc.text(netWeightLabel, leftCol, yPosition)
  const netWeight = getValue(data, ['Net Weight (PALLET):', 'NET WEIGHT (PALLET):', 'NET WEIGHT (PALLET)', 'Net Weight', 'Net Weight (Pallet)'])
  
  // GROSS WEIGHT (PALLET) - left column (below NET WEIGHT) - format to 2 decimal places
  const grossWeightLabel = 'GROSS WEIGHT (PALLET):'
  const grossWeightRaw = getValue(data, ['Gross Weight (PALLET):', 'GROSS WEIGHT (PALLET):', 'GROSS WEIGHT (PALLET)', 'Gross Weight', 'Gross Weight (Pallet)'])
  // Format to 2 decimal places
  let grossWeight = grossWeightRaw
  const grossWeightNum = parseFloat(grossWeightRaw)
  if (!isNaN(grossWeightNum)) {
    grossWeight = grossWeightNum.toFixed(2)
  }
  
  // Calculate the maximum label width to align values at the same x position
  const maxLabelWidth = Math.max(doc.getTextWidth(netWeightLabel), doc.getTextWidth(grossWeightLabel))
  const valueX = leftCol + maxLabelWidth + 3 // Aligned x position for both values
  
  // Display NET WEIGHT value and KGS aligned
  doc.setFont('helvetica', 'normal')
  doc.text(netWeight, valueX, yPosition)
  doc.text('KGS.', valueX + doc.getTextWidth(netWeight) + 2, yPosition)
  yPosition += 4
  
  // Display GROSS WEIGHT value and KGS aligned (same x position)
  doc.setFont('helvetica', 'bold')
  doc.text(grossWeightLabel, leftCol, yPosition)
  doc.setFont('helvetica', 'normal')
  doc.text(grossWeight, valueX, yPosition)
  doc.text('KGS.', valueX + doc.getTextWidth(grossWeight) + 2, yPosition)
  yPosition += 4
  
  // PALLET DIMENSIONS MM - left column
  doc.setFont('helvetica', 'bold')
  doc.text('PALLET DIMENSIONS MM:', leftCol, yPosition)
  yPosition += 4 // Add space after heading to prevent overlap
  
  // Parse dimensions from data
  let length = getValue(data, ['Length', 'LENGTH'])
  let width = getValue(data, ['Width', 'WIDTH'])
  let height = getValue(data, ['Height', 'HEIGHT'])
  
  // If not found as individual columns, try to parse from combined dimension string
  if (!length || !width || !height) {
    const dimensions = getValue(data, ['Pallet Dimension MM:', 'PALLET DIMENSIONS MM:', 'PALLET DIMENSIONS MM', 'Dimensions', 'Pallet Dimensions'])
    if (dimensions) {
      // Parse format like "725 X 895 X 2625" or "725X895X2625"
      const parts = dimensions.split(/[xX]/).map(p => p.trim()).filter(p => p)
      if (parts.length >= 3) {
        length = parts[0]
        width = parts[1]
        height = parts[2]
      }
    }
  }
  
  // Default if still not found
  if (!length || !width || !height) {
    length = '725'
    width = '895'
    height = '2625'
  }
  
  // Display labels and values - labels on left, values aligned on right - slightly bigger
  const dimLabelX = leftCol + 5
  const dimValueX = leftCol + 42 // Align all values at same x position
  
  doc.setFont('helvetica', 'normal')
  doc.setFontSize(7)
  
  // Width (first dimension shown)
  doc.text('Width', dimLabelX, yPosition)
  doc.text(width, dimValueX, yPosition)
  yPosition += 3.5
  
  // Height (second dimension)
  doc.text('Height', dimLabelX, yPosition)
  doc.text(height, dimValueX, yPosition)
  yPosition += 3.5
  
  // Length (third dimension - shown last)
  doc.text('Length', dimLabelX, yPosition)
  doc.text(length, dimValueX, yPosition)
  
  // Reset font size
  doc.setFontSize(7.5)
  yPosition += 4
  
  // PALLET NUMBER - right column (aligned with PALLET DIMENSIONS)
  doc.setFont('helvetica', 'bold')
  doc.text('PALLET NUMBER:', rightCol, yPosition)
  const palletNum = getValue(data, ['Pallet No.', 'PALLET NUMBER', 'Pallet Number', 'Pallet #', 'Pallet'])
  doc.setFont('helvetica', 'normal')
  doc.text(palletNum, rightCol + 36, yPosition)
  yPosition += 5
  
  // MADE IN PAKISTAN (centered, bold, underlined, at bottom) - slightly bigger
  doc.setFontSize(7.5)
  doc.setFont('helvetica', 'bold')
  const originText = 'MADE IN PAKISTAN'
  const originWidth = doc.getTextWidth(originText)
  const originX = (pageWidth - originWidth) / 2
  const originY = pageHeight - margin - 1.5
  doc.text(originText, originX, originY)
  
  // Underline
  doc.setLineWidth(0.3)
  doc.line(originX - 1, originY + 0.8, originX + originWidth + 1, originY + 0.8)
}

export async function convertExcelToPdf(excelData, filename) {
  // Tag dimensions: height=4", width=6" (landscape orientation)
  // In mm: height=101.6mm (4"), width=152.4mm (6")
  const doc = new jsPDF({
    orientation: 'landscape',
    unit: 'mm',
    format: [152.4, 101.6] // width x height: 6" x 4"
  })

  // Page dimensions in mm
  const pageWidth = 152.4  // 6 inches (width)
  const pageHeight = 101.6 // 4 inches (height)
  const margin = 6 // Reduced margin for smaller format
  const contentWidth = pageWidth - (margin * 2)
  const leftCol = margin
  const rightCol = pageWidth / 2 + 3
  const colWidth = (pageWidth - margin * 2) / 2 - 3
  const labelValueGap = 15 // Gap between label and value
  
  // Load logo once to reuse for all pages
  let logoBase64 = null
  let logoWidth = 0
  let logoHeight = 0
  
  try {
    const logoUrl = '/tri-pack-logo.jpeg'
    
    // Load image first to get dimensions
    const img = new Image()
    img.crossOrigin = 'anonymous'
    
    await new Promise((resolve, reject) => {
      img.onload = resolve
      img.onerror = reject
      img.src = logoUrl
    })
    
    // Get base64 for PDF
    logoBase64 = await loadImageAsBase64(logoUrl)
    
    // Calculate dimensions maintaining aspect ratio
    const maxLogoHeight = 8
    const aspectRatio = img.width / img.height
    logoHeight = maxLogoHeight
    logoWidth = logoHeight * aspectRatio
    
    console.log('Logo loaded successfully')
  } catch (error) {
    console.warn('Could not load logo:', error)
  }
  
  // Get all rows from Excel data
  const rows = excelData.rows || []
  
  if (rows.length === 0) {
    console.warn('No data rows found in Excel file')
    return
  }
  
  console.log(`Generating PDF with ${rows.length} page(s)`)
  
  // Generate a page for each row
  for (let i = 0; i < rows.length; i++) {
    const data = rows[i]
    
    // Add new page for all rows except the first
    if (i > 0) {
      doc.addPage()
    }
    
    // Generate page for this row
    await generatePageForRow(doc, data, filename, pageWidth, pageHeight, margin, contentWidth, leftCol, rightCol, colWidth, labelValueGap, logoBase64, logoWidth, logoHeight)
  }
  
  // Save the PDF
  const pdfFilename = filename.replace(/\.(xlsx|xls)$/i, '.pdf')
  doc.save(pdfFilename)
  
  console.log(`PDF generated successfully with ${rows.length} page(s)`)
}

