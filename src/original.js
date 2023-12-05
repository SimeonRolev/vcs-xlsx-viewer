import Dayjs from 'dayjs'
import ExcelJS from 'exceljs'

function getPrototype(value) {
  return Object.prototype.toString.call(value).replace(/^\[object (\S+)\]$/, '$1').toLowerCase()
}

function isNumeric(str) {
  if (typeof str === 'number') return true;
  if (typeof str != "string") return false // we only process strings!  
  return !isNaN(str) && // use type coercion to parse the _entirety_ of the string (`parseFloat` alone does not do this)...
         !isNaN(parseFloat(str)) // ...and ensure strings of whitespace fail
}

function addImageToCell({mediaImage, tdElement}) {
  const img = document.createElement('img');
  const { buffer, extension } = mediaImage;
  img.src = `data:image/${extension};base64,${buffer.toString('base64')}`
  img.style.maxHeight = '100%';
  img.style.maxWidth = '100%';
  tdElement.appendChild(img);
}

async function renderXlsx({
  arrayBuffer: xlsxData,
  node: xlsxElement,
  options: xlsxOptions = {}
}) {
  const {
    initialSheetIndex = 0,
    frameRenderSize = 500,
    onLoad = () => {},
    onRender = () => {},
    onSwitch = () => {}
  } = xlsxOptions
  if (!['blob', 'file', 'arraybuffer'].includes(getPrototype(xlsxData))) {
    throw new Error(`renderXlsx ${xlsxData} is not a file`)
  }
  if (getPrototype(xlsxElement).indexOf('element') === -1) {
    throw new Error(`renderXlsx ${xlsxElement} is not a element`)
  }
  if (getPrototype(xlsxOptions) !== 'object') {
    throw new Error(`renderXlsx ${xlsxOptions} is not a object`)
  }
  if (getPrototype(initialSheetIndex) !== 'number') {
    throw new Error('renderXlsx \'initialSheetIndex\' is not a number')
  }
  if (getPrototype(frameRenderSize) !== 'number') {
    throw new Error('renderXlsx \'frameRenderSize\' is not a number')
  }
  if (getPrototype(onLoad).indexOf('function') === -1) {
    throw new Error('renderXlsx \'onLoad\' is not a function')
  }
  if (getPrototype(onRender).indexOf('function') === -1) {
    throw new Error('renderXlsx \'onRender\' is not a function')
  }
  if (getPrototype(onSwitch).indexOf('function') === -1) {
    throw new Error('renderXlsx \'onSwitch\' is not a function')
  }
  // viewer params init
  const viewerParams = {
    arrayBuffer: undefined,
    sheetList: [],
    currentSheetId: undefined
  }
  // viewer elements init
  const viewerElements = {
    xlsxElement: undefined,
    containerElement: undefined,
    tipElement: undefined,
    tableElement: undefined
  }
  // viewer methods init
  const viewerMethods = {
    async loadXlsxDataWorkbook() {
      return new Promise((resolve) => {
        try {
          // load workbook
          (new ExcelJS.Workbook().xlsx.load((viewerParams.arrayBuffer))).then((workbook) => {
            const worksheet = workbook.getWorksheet(initialSheetIndex + 1)
            const sheetItem = {
              id: xlsxOptions.initialSheetIndex,
              name: worksheet.name,
              columns: [],
              rows: [],
              merges: [],
              worksheet,
              workbook,
              images: worksheet.getImages(),
              rendered: false
            }
            // set sheet column
            for (let i = 0; i < worksheet.columnCount; i++) {
              const column = worksheet.getColumn(i + 1)
              sheetItem.columns.push(column)
            }
            // set sheet row
            for (let i = 0; i < worksheet.rowCount; i++) {
              const row = worksheet.getRow(i + 1)
              // set sheet row cell merges
              for (let j = 0; j < row.cellCount; j++) {
                const cell = row.getCell(j + 1)
                if (cell.isMerged) {
                  const targetAddress = sheetItem.merges.find((item) => item.address === cell.master._address)
                  if (targetAddress) {
                    targetAddress.cells.push(cell)
                  } else {
                    sheetItem.merges.push({
                      address: cell._address,
                      master: cell,
                      cells: [cell]
                    })
                  }
                }
              }
              sheetItem.rows.push(row)
            }
            viewerParams.sheetList.push(sheetItem)
            if (viewerElements.tipElement && getPrototype(viewerElements.tipElement).indexOf('element') !== -1) {
              viewerElements.tipElement.style.display = 'none'
              onLoad(viewerParams.sheetList)
            }
            resolve()
          })
        } catch (err) {
          if (viewerElements.tipElement && getPrototype(viewerElements.tipElement).indexOf('element') !== -1) {
            viewerElements.tipElement.innerText = `Load error: ${err}`
          }
          console.error('[xlsx-viewer] load error: ', err)
        }
      })
    },
    createXlsxContainerElement() {
      const xlsxViewerContainerElement = document.createElement('div')
      const xlsxViewerTableElement = document.createElement('div')
      const xlsxViewerTipElement = document.createElement('div')
      const oldXlsxViewerContainerElement = xlsxElement.querySelector('.xlsx-viewer-container')
      xlsxViewerContainerElement.classList.add('xlsx-viewer-container')
      xlsxViewerTableElement.classList.add('xlsx-viewer-table')
      xlsxViewerTipElement.classList.add('xlsx-viewer-tip')
      xlsxViewerTipElement.innerText = 'Loading...'
      viewerElements.xlsxElement = xlsxElement
      viewerElements.tableElement = xlsxViewerTableElement
      viewerElements.tipElement = xlsxViewerTipElement
      viewerElements.containerElement = xlsxViewerContainerElement
      viewerElements.containerElement.appendChild(xlsxViewerTableElement)
      viewerElements.containerElement.appendChild(xlsxViewerTipElement)
      if (oldXlsxViewerContainerElement) {
        viewerElements.xlsxElement.replaceChild(viewerElements.containerElement, oldXlsxViewerContainerElement)
      } else {
        viewerElements.xlsxElement.appendChild(viewerElements.containerElement)
      }
    },
    createTableContainerElement() {
      for (let i = 0; i < viewerParams.sheetList.length; i++) {
        const sheetItem = viewerParams.sheetList[initialSheetIndex]
        const xlsxViewerTableItemElement = document.createElement('div')
        xlsxViewerTableItemElement.classList.add('xlsx-viewer-table-content')
        viewerMethods.createTableContentElement(sheetItem, xlsxViewerTableItemElement)
        viewerElements.tableElement?.appendChild(xlsxViewerTableItemElement)
      }
    },
    createTableContentElement(
      sheetItem,
      xlsxViewerTableItemElement
    ) {
      // set table element
      const tableElement = document.createElement('table')
      const theadElement = document.createElement('thead')
      const tbodyElement = document.createElement('tbody')
      const tbodyTrElementArr = []
      const appendTrElementToTbodyElement = (currentPage = 0) => {
        requestAnimationFrame(() => {
          for (let i = 0; i < frameRenderSize; i++) {
            const trElement = tbodyTrElementArr[currentPage * frameRenderSize + i]
            if (trElement) {
              tbodyElement.appendChild(trElement)
            } else {
              break
            }
          }
          if (currentPage * frameRenderSize < tbodyTrElementArr.length) {
            appendTrElementToTbodyElement(currentPage + 1)
          } else {
            sheetItem.rendered = true
            onRender(sheetItem)
          }
        })
      }
      // set sheet columns element
      if (sheetItem.columns.length > 0) {
        const trElement = document.createElement('tr')
        const firstThElement = document.createElement('th')
        let tableWidth = 50
        firstThElement.style.width = '50px'
        trElement.appendChild(firstThElement)
        for (let i = 0; i < sheetItem.columns.length; i++) {
          const column = sheetItem.columns[i]
          const thElement = document.createElement('th')
          const columnWidth = column.width > 0 ? column.width / 0.125 : 100
          tableWidth = tableWidth + columnWidth
          thElement.style.width = `${columnWidth}px`
          thElement.innerText = column.letter
          trElement.appendChild(thElement)
        }
        theadElement.appendChild(trElement)
        tableElement.style.width = `${tableWidth}px`
      }
      // set sheet rows element
      if (sheetItem.rows.length > 0) {
        for (let i = 0; i < sheetItem.rows.length; i++) {
          const row = sheetItem.rows[i]
          const trElement = document.createElement('tr')
          const firstTdElement = document.createElement('td')
          firstTdElement.innerText = (i + 1).toString()
          trElement.appendChild(firstTdElement)
          for (let j = 0; j < sheetItem.columns.length; j++) {
            const cell = row.getCell(j + 1)
            if (cell.isMerged && cell.master._address !== cell._address) {
              continue
            }
            const tdElement = document.createElement('td')
            if (cell.isMerged && cell.master._address === cell._address) {
              const merge = sheetItem.merges.find(item => item.address === cell._address)
              if (merge) {
                const maxCol = Math.max.apply(Math, merge.cells.map((cell) => cell.col))
                const maxRow = Math.max.apply(Math, merge.cells.map((cell) => cell.row))
                const colSpan = maxCol - cell.col + 1
                const rowSpan = maxRow - cell.row + 1
                tdElement.setAttribute('colspan', colSpan.toString())
                tdElement.setAttribute('rowSpan', rowSpan.toString())
              }
            }
            // set row size
            if (row.height) {
              tdElement.style.height = `${row.height / 0.75}px`
            }
            // set cell alignment
            if (cell.style?.alignment) {
              const { horizontal, vertical } = cell.style.alignment
              tdElement.style.textAlign = horizontal
              tdElement.style.verticalAlign = vertical
            }
            // set cell background
            if (cell.style?.fill) {
              const { fgColor } = cell.style.fill
              tdElement.style.backgroundColor = fgColor?.argb ? (utilMethods.parseARGB(fgColor?.argb)?.color) : '#fff'
            }
            // set cell border
            if (cell.style?.border) {
              const { top, bottom, left, right } = cell.style.border
              tdElement.style.borderTop = top?.color?.argb ? '1px solid ' + (utilMethods.parseARGB(top?.color?.argb)?.color) : ''
              tdElement.style.borderBottom = bottom?.color?.argb ? '1px solid ' + (utilMethods.parseARGB(bottom?.color?.argb)?.color) : ''
              tdElement.style.borderLeft = left?.color?.argb ? '1px solid ' + (utilMethods.parseARGB(left?.color?.argb)?.color) : ''
              tdElement.style.borderRight = right?.color?.argb ? '1px solid ' + (utilMethods.parseARGB(right?.color?.argb)?.color) : ''
            }
            // set cell font
            if (cell.style?.font) {
              const { color, name, size, bold, italic, underline } = cell.style.font
              tdElement.style.color = color?.argb ? (utilMethods.parseARGB(color?.argb)?.color) : '#333'
              tdElement.style.fontFamily = name
              tdElement.style.fontSize = size ? `${size / 0.75}px` : '14px'
              tdElement.style.fontWeight = bold ? 'bold' : 'normal'
              tdElement.style.fontStyle = italic ? 'italic' : 'normal'
              tdElement.style.textDecoration = underline ? 'underline' : 'none'
            }

            // Add images
            sheetItem.images
              .filter(image => image.range.tl.nativeCol + 1 === cell.col && image.range.tl.nativeRow + 1 === cell.row)
              .forEach(image => {
                const mediaImage = sheetItem.workbook.model.media.find(m => m.index === image.imageId)
                addImageToCell({ mediaImage, tdElement });
              })
            
            // set cell value
            if (getPrototype(cell.value) === 'object') {
              const { richText, hyperlink } = cell.value
              if (richText && getPrototype(richText) === 'array') {
                for (const span of richText) {
                  const spanElement = document.createElement('span')
                  if (span?.font) {
                    const { color, name, size, bold, italic, underline } = span.font
                    spanElement.style.color = color?.argb ? (utilMethods.parseARGB(color?.argb)?.color) : '#333'
                    spanElement.style.fontFamily = name
                    spanElement.style.fontSize = size ? `${size / 0.75}px` : '14px'
                    spanElement.style.fontWeight = bold ? 'bold' : 'normal'
                    spanElement.style.fontStyle = italic ? 'italic' : 'normal'
                    spanElement.style.textDecoration = underline ? 'underline' : 'none'
                  }
                  spanElement.innerText = span.text
                  tdElement.appendChild(spanElement)
                }
              } else if (getPrototype(hyperlink) === 'string') {
                const link = cell.value
                const linkElement = document.createElement('a')
                linkElement.setAttribute('href', link.hyperlink)
                linkElement.setAttribute('target', '_blank')
                linkElement.style.color = '#2f63c1'
                linkElement.style.textDecoration = 'underline'
                linkElement.innerText = link.text
                tdElement.appendChild(linkElement)
              }
            } else if (getPrototype(cell.value) === 'date') {
              tdElement.innerText = Dayjs(cell.value).format('YYYY-MM-DD HH:mm:ss')
            } else {
              const span = document.createElement('span')
              span.innerText = isNumeric(cell.value) ? parseFloat(cell.value).toFixed(2) : cell.value
              tdElement.appendChild(span)
            }
            trElement.appendChild(tdElement)
          }
          tbodyTrElementArr.push(trElement)
        }
        appendTrElementToTbodyElement()
      }
      tableElement.appendChild(theadElement)
      tableElement.appendChild(tbodyElement)
      xlsxViewerTableItemElement.appendChild(tableElement)
    }
  }
  // util methods init
  const utilMethods = {
    blobOrFileToArrayBuffer(blob) {
      return new Promise(resolve => {
        const fileReader = new FileReader()
        fileReader.onload = (e) => {
          resolve(e.target.result)
        }
        fileReader.readAsArrayBuffer(blob)
      })
    },
    parseARGB(argb) {
      if (getPrototype(argb) !== 'string' || argb.length !== 8) {
        return undefined
      }
      let result
      const color = []
      for (let i = 0; i < 4; i++) {
        color.push(argb.substr(i * 2, 2))
      }
      const [a, r, g, b] = color.map((v) => parseInt(v, 16))
      result = {
        argb: { a, r, g, b },
        color: `rgba(${r}, ${g}, ${b}, ${a / 255})`
      }
      return result
    }
  }
  // check browser compatibility
  viewerMethods.createXlsxContainerElement()
  if (
    (window.navigator.userAgent.indexOf('MSIE') !== -1 || 'ActiveXObject' in window) &&
    (viewerElements.tipElement && getPrototype(viewerElements.tipElement).indexOf('element') !== -1)
  ) {
    viewerElements.tipElement.innerText = 'Browser incompatibility.'
    return
  }
  // load xlsx data
  if (['blob', 'file'].includes(getPrototype(xlsxData))) {
    viewerParams.arrayBuffer = await utilMethods.blobOrFileToArrayBuffer(xlsxData)
  } else if (getPrototype(xlsxData) === 'arraybuffer') {
    viewerParams.arrayBuffer = (xlsxData)
  }
  await viewerMethods.loadXlsxDataWorkbook()
  viewerMethods.createTableContainerElement()
}

export { renderXlsx }