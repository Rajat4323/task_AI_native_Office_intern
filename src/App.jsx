import { useState, useRef, useCallback, useMemo, useEffect } from 'react'
import './App.css'
import { createEngine } from './engine/core.js'

const TOTAL_ROWS = 50
const TOTAL_COLS = 50

export default function App() {
  // Engine instance is created once and reused across renders
  // Note: The engine maintains its own internal state, so React state is only used for UI updates
  const [engine] = useState(() => createEngine(TOTAL_ROWS, TOTAL_COLS))
  const [version, setVersion] = useState(0)
  const originalRowOrderRef = useRef(Array.from({ length: engine.rows }, (_, i) => i))
const [rowOrder, setRowOrder] = useState(() => [...originalRowOrderRef.current])
  const [selectedCell, setSelectedCell] = useState(null)
  const [editingCell, setEditingCell] = useState(null)
  const [editValue, setEditValue] = useState('')
  // Cell styles are stored separately from engine data
  // Format: { "row,col": { bold: bool, italic: bool, ... } }
  const [cellStyles, setCellStyles] = useState({})
  const cellInputRef = useRef(null)

  const [sortConfig, setSortConfig] = useState({ column: null, direction: 'none' })
   const [filters, setFilters] = useState({})
   const [openFilterCol, setOpenFilterCol] = useState(null)

  const forceRerender = useCallback(() => setVersion(v => v + 1), [])

  // Reset sort/rowOrder when engine row count changes
useEffect(() => {
  originalRowOrderRef.current = Array.from({ length: engine.rows }, (_, i) => i)
  setSortConfig({ column: null, direction: 'none' })
  setRowOrder([...originalRowOrderRef.current])
}, [engine.rows])

  // ────── Clipboard (copy / paste) ──────

  useEffect(() => {
    async function handleClipboardKeys(event) {
      // Only handle when spreadsheet has a selected cell and we're not actively editing
      if (!selectedCell || editingCell) return

      const isMac = navigator.platform.toLowerCase().includes('mac')
      const ctrlOrCmd = isMac ? event.metaKey : event.ctrlKey
      if (!ctrlOrCmd) return

      const key = event.key.toLowerCase()

      // Copy: Ctrl/Cmd + C
      if (key === 'c') {
        event.preventDefault()

        const { r, c } = selectedCell
        const cell = engine.getCell(r, c)
        const value = cell.computed
        const textToCopy =
          value === null || typeof value === 'undefined' ? '' : String(value)

        try {
          await navigator.clipboard.writeText(textToCopy)
        } catch {
          // Silently ignore clipboard errors (permissions, http context, etc.)
        }
      }

      // Paste: Ctrl/Cmd + V
      if (key === 'v') {
        event.preventDefault()

        let text = ''
        try {
          text = await navigator.clipboard.readText()
        } catch {
          return
        }

        if (!text) return

        // Normalize newlines and trim trailing blank lines
        text = text.replace(/\r\n/g, '\n')
        const rawRows = text.split('\n')
        const rows = rawRows.filter((row, idx) => row.length > 0 || idx < rawRows.length - 1)
        if (rows.length === 0) return

        const startRow = selectedCell.r
        const startCol = selectedCell.c

        for (let rowOffset = 0; rowOffset < rows.length; rowOffset++) {
          const cols = rows[rowOffset].split('\t')
          for (let colOffset = 0; colOffset < cols.length; colOffset++) {
            const targetRow = startRow + rowOffset
            const targetCol = startCol + colOffset

            if (targetRow >= engine.rows || targetCol >= engine.cols) continue

            const rawValue = cols[colOffset]
            // Preserve empty strings; engine.setCell will handle parsing / formulas
            engine.setCell(targetRow, targetCol, rawValue)
          }
        }

        forceRerender()
      }
    }

    document.addEventListener('keydown', handleClipboardKeys)
    return () => document.removeEventListener('keydown', handleClipboardKeys)
  }, [selectedCell, editingCell, engine, forceRerender])

  // ────── Sorting helpers ──────

  const getCellSortKey = useCallback((row, col) => {
    const cell = engine.getCell(row, col)

    if (cell.error) {
      return { kind: 'error', value: cell.error }
    }

    const value = cell.computed

    if (value === null || value === '' || typeof value === 'undefined') {
      return { kind: 'empty', value: '' }
    }

    if (typeof value === 'number') {
      return { kind: 'number', value }
    }

    // Attempt numeric coercion for numeric-looking strings
    const numeric = Number(value)
    if (!Number.isNaN(numeric)) {
      return { kind: 'number', value: numeric }
    }

    return { kind: 'string', value: String(value) }
  }, [engine])

  const compareRows = useCallback((rowA, rowB, colIndex) => {
    const a = getCellSortKey(rowA, colIndex)
    const b = getCellSortKey(rowB, colIndex)

    const rank = { number: 0, string: 1, empty: 2, error: 3 }

    if (a.kind !== b.kind) {
      return rank[a.kind] - rank[b.kind]
    }

    let diff = 0
    if (a.kind === 'number') {
      diff = a.value - b.value
    } else if (a.kind === 'string') {
      diff = a.value.localeCompare(b.value)
    } else {
      diff = 0
    }

    return diff
  }, [getCellSortKey])

  const handleSortClick = useCallback((colIndex) => {
    setSortConfig(prev => {
      let nextDirection = 'asc'
      let nextColumn = colIndex

      if (prev.column === colIndex) {
        if (prev.direction === 'asc') {
          nextDirection = 'desc'
        } else if (prev.direction === 'desc') {
          nextDirection = 'none'
          nextColumn = null
        }
      }

      return { column: nextColumn, direction: nextDirection }
    })
  }, [])

  // Derive rowOrder from sortConfig to keep sorting view-layer only and avoid
  // inconsistent state when React batches updates.
// NEW — paste this in
useEffect(() => {
  if (sortConfig.direction === 'none' || sortConfig.column === null) {
    setRowOrder([...originalRowOrderRef.current])
  } else {
    const sign = sortConfig.direction === 'asc' ? 1 : -1
    const sorted = [...originalRowOrderRef.current].sort(
      (a, b) => sign * compareRows(a, b, sortConfig.column)
    )
    setRowOrder(sorted)
  }
}, [sortConfig, compareRows])

  // ────── Filtering helpers ──────

  const getUniqueValuesForColumn = useCallback((colIndex) => {
    const unique = new Map()
    for (let r = 0; r < engine.rows; r++) {
      const value = engine.getCell(r, colIndex).computed
      const key =
        value === null || typeof value === 'undefined'
          ? 'null|'
          : `${typeof value}|${String(value)}`
      if (!unique.has(key)) {
        unique.set(key, value)
      }
    }
    return Array.from(unique.values())
  }, [engine])

  const toggleFilterValue = useCallback((colIndex, value) => {
    setFilters(prev => {
      const key = String(colIndex)
      const current = prev[key] || []
      const exists = current.some(v => Object.is(v, value))
      const nextForCol = exists
        ? current.filter(v => !Object.is(v, value))
        : [...current, value]

      const next = { ...prev }
      if (nextForCol.length === 0) {
        delete next[key]
      } else {
        next[key] = nextForCol
      }
      return next
    })
  }, [])

  const clearFilter = useCallback((colIndex) => {
    setFilters(prev => {
      if (!(String(colIndex) in prev)) return prev
      const next = { ...prev }
      delete next[String(colIndex)]
      return next
    })
  }, [])

  const passesFilters = useCallback((rowIndex) => {
    const entries = Object.entries(filters)
    if (entries.length === 0) return true

    for (const [colKey, selectedValues] of entries) {
      if (!selectedValues || selectedValues.length === 0) continue
      const colIndex = parseInt(colKey, 10)
      const cell = engine.getCell(rowIndex, colIndex)
      const value = cell.computed
      const match = selectedValues.some(v => Object.is(v, value))
      if (!match) return false
    }
    return true
  }, [filters, engine])

  // ────── Local storage persistence ──────

  const STORAGE_KEY = 'spreadsheet-state-v1'

  // Restore engine + styles on initial load
  useEffect(() => {
    if (typeof window === 'undefined') return
    try {
      const raw = window.localStorage.getItem(STORAGE_KEY)
      if (!raw) return
      const parsed = JSON.parse(raw)
      if (!parsed || typeof parsed !== 'object') return

      const { engineSnapshot, stylesSnapshot } = parsed
      if (engineSnapshot) {
        const ok = engine.deserialize(engineSnapshot)
        if (!ok) return
      }
      if (stylesSnapshot && typeof stylesSnapshot === 'object') {
        setCellStyles(stylesSnapshot)
      }
      // Force UI to reflect restored state
      forceRerender()
    } catch {
      // Corrupted or inaccessible storage; ignore and start fresh
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [engine])

  // Debounced auto-save when engine data or styles change
  useEffect(() => {
    if (typeof window === 'undefined') return

    const handleSave = () => {
      try {
        const engineSnapshot = engine.serialize()
        const stylesSnapshot = cellStyles
        const payload = { engineSnapshot, stylesSnapshot }
        window.localStorage.setItem(STORAGE_KEY, JSON.stringify(payload))
      } catch {
        // Ignore storage errors (quota exceeded, etc.)
      }
    }

    const timeoutId = window.setTimeout(handleSave, 500)
    return () => window.clearTimeout(timeoutId)
  }, [engine, cellStyles, version])

  // ────── Cell style helpers ──────

  const getCellStyle = useCallback((row, col) => {
    const key = `${row},${col}`
    return cellStyles[key] || {
      bold: false, italic: false, underline: false,
      bg: 'white', color: '#202124', align: 'left', fontSize: 13
    }
  }, [cellStyles])

  const updateCellStyle = useCallback((row, col, updates) => {
    const key = `${row},${col}`
    setCellStyles(prev => ({
      ...prev,
      [key]: { ...getCellStyle(row, col), ...updates }
    }))
  }, [getCellStyle])

  // ────── Cell editing ──────

  const startEditing = useCallback((row, col) => {
    setSelectedCell({ r: row, c: col })
    setEditingCell({ r: row, c: col })
    const cellData = engine.getCell(row, col)
    setEditValue(cellData.raw)
    setTimeout(() => cellInputRef.current?.focus(), 0)
  }, [engine])

  const commitEdit = useCallback((row, col) => {
    // Only commit if the value actually changed to avoid unnecessary recalculations
    const currentCell = engine.getCell(row, col)
    if (currentCell.raw !== editValue) {
      engine.setCell(row, col, editValue)
      forceRerender()
    }
    setEditingCell(null)
  }, [engine, editValue, forceRerender])

  const handleCellClick = useCallback((row, col) => {
    if (editingCell && (editingCell.r !== row || editingCell.c !== col)) {
      commitEdit(editingCell.r, editingCell.c)
    }
    if (!editingCell || editingCell.r !== row || editingCell.c !== col) {
      startEditing(row, col)
    }
  }, [editingCell, commitEdit, startEditing])

  // ────── Keyboard navigation ──────

  const handleKeyDown = useCallback((event, row, col) => {
    if (event.key === 'Enter') {
      event.preventDefault()
      commitEdit(row, col)
      startEditing(Math.min(row + 1, engine.rows - 1), col)
    } else if (event.key === 'Tab') {
      event.preventDefault()
      commitEdit(row, col)
      startEditing(row, Math.min(col + 1, engine.cols - 1))
    } else if (event.key === 'Escape') {
      setEditValue(engine.getCell(row, col).raw)
      setEditingCell(null)
    } else if (event.key === 'ArrowDown') {
      event.preventDefault()
      commitEdit(row, col)
      startEditing(Math.min(row + 1, engine.rows - 1), col)
    } else if (event.key === 'ArrowUp') {
      event.preventDefault()
      commitEdit(row, col)
      startEditing(Math.max(row - 1, 0), col)
    } else if (event.key === 'ArrowLeft') {
      event.preventDefault()
      commitEdit(row, col)
      if (col > 0) {
        startEditing(row, col - 1)
      } else if (row > 0) {
        startEditing(row - 1, engine.cols - 1)
      }
    } else if (event.key === 'ArrowRight') {
      event.preventDefault()
      commitEdit(row, col)
      startEditing(row, Math.min(col + 1, engine.cols - 1))
    }
  }, [engine, commitEdit, startEditing])

  // ────── Formula bar handlers ──────

  const handleFormulaBarKeyDown = useCallback((event) => {
    if (!editingCell) return
    handleKeyDown(event, editingCell.r, editingCell.c)
  }, [editingCell, handleKeyDown])

  const handleFormulaBarFocus = useCallback(() => {
    if (selectedCell && !editingCell) {
      setEditingCell(selectedCell)
      setEditValue(engine.getCell(selectedCell.r, selectedCell.c).raw)
    }
  }, [selectedCell, editingCell, engine])

  const handleFormulaBarChange = useCallback((value) => {
    if (!editingCell && selectedCell) setEditingCell(selectedCell)
    setEditValue(value)
  }, [editingCell, selectedCell])

  // ────── Undo / Redo ──────

  const handleUndo = useCallback(() => { if (engine.undo()) forceRerender() }, [engine, forceRerender])
  const handleRedo = useCallback(() => { if (engine.redo()) forceRerender() }, [engine, forceRerender])

  // ────── Formatting toggles ──────

  const toggleBold = useCallback(() => {
    if (!selectedCell) return
    const style = getCellStyle(selectedCell.r, selectedCell.c)
    updateCellStyle(selectedCell.r, selectedCell.c, { bold: !style.bold })
  }, [selectedCell, getCellStyle, updateCellStyle])

  const toggleItalic = useCallback(() => {
    if (!selectedCell) return
    const style = getCellStyle(selectedCell.r, selectedCell.c)
    updateCellStyle(selectedCell.r, selectedCell.c, { italic: !style.italic })
  }, [selectedCell, getCellStyle, updateCellStyle])

  const toggleUnderline = useCallback(() => {
    if (!selectedCell) return
    const style = getCellStyle(selectedCell.r, selectedCell.c)
    updateCellStyle(selectedCell.r, selectedCell.c, { underline: !style.underline })
  }, [selectedCell, getCellStyle, updateCellStyle])

  const changeFontSize = useCallback((size) => {
    if (!selectedCell) return
    updateCellStyle(selectedCell.r, selectedCell.c, { fontSize: size })
  }, [selectedCell, updateCellStyle])

  const changeAlignment = useCallback((align) => {
    if (!selectedCell) return
    updateCellStyle(selectedCell.r, selectedCell.c, { align })
  }, [selectedCell, updateCellStyle])

  const changeFontColor = useCallback((color) => {
    if (!selectedCell) return
    updateCellStyle(selectedCell.r, selectedCell.c, { color })
  }, [selectedCell, updateCellStyle])

  const changeBackgroundColor = useCallback((color) => {
    if (!selectedCell) return
    updateCellStyle(selectedCell.r, selectedCell.c, { bg: color })
  }, [selectedCell, updateCellStyle])

  // ────── Clear operations ──────

  const clearSelectedCell = useCallback(() => {
    if (!selectedCell) return
    engine.setCell(selectedCell.r, selectedCell.c, '')
    forceRerender()
    // Remove style entry for cleared cell
    // Note: This deletes the style object entirely - if you need to preserve default styles,
    // you may want to set them explicitly rather than deleting
    const key = `${selectedCell.r},${selectedCell.c}`
    setCellStyles(prev => { const next = { ...prev }; delete next[key]; return next })
    setEditValue('')
  }, [selectedCell, engine, forceRerender])

  const clearAllCells = useCallback(() => {
    for (let r = 0; r < engine.rows; r++) {
      for (let c = 0; c < engine.cols; c++) {
        engine.setCell(r, c, '')
      }
    }
    forceRerender()
    setCellStyles({})
    setSelectedCell(null)
    setEditingCell(null)
    setEditValue('')
  }, [engine, forceRerender])

  // ────── Row / Column operations ──────

  const insertRow = useCallback(() => {
    if (!selectedCell) return
    engine.insertRow(selectedCell.r)
    forceRerender()
    setSelectedCell({ r: selectedCell.r + 1, c: selectedCell.c })
  }, [selectedCell, engine, forceRerender])

  const deleteRow = useCallback(() => {
    if (!selectedCell) return
    engine.deleteRow(selectedCell.r)
    forceRerender()
    if (selectedCell.r >= engine.rows) {
      setSelectedCell({ r: engine.rows - 1, c: selectedCell.c })
    }
  }, [selectedCell, engine, forceRerender])

  const insertColumn = useCallback(() => {
    if (!selectedCell) return
    engine.insertColumn(selectedCell.c)
    forceRerender()
    setSelectedCell({ r: selectedCell.r, c: selectedCell.c + 1 })
  }, [selectedCell, engine, forceRerender])

  const deleteColumn = useCallback(() => {
    if (!selectedCell) return
    engine.deleteColumn(selectedCell.c)
    forceRerender()
    if (selectedCell.c >= engine.cols) {
      setSelectedCell({ r: selectedCell.r, c: engine.cols - 1 })
    }
  }, [selectedCell, engine, forceRerender])

  // ────── Derived state ──────

  const selectedCellStyle = useMemo(() => {
    return selectedCell ? getCellStyle(selectedCell.r, selectedCell.c) : null
  }, [selectedCell, getCellStyle])

  const getColumnLabel = useCallback((col) => {
    let label = ''
    let num = col + 1
    while (num > 0) {
      num--
      label = String.fromCharCode(65 + (num % 26)) + label
      num = Math.floor(num / 26)
    }
    return label
  }, [])

  const selectedCellLabel = selectedCell
    ? `${getColumnLabel(selectedCell.c)}${selectedCell.r + 1}`
    : 'No cell'

  // Formula bar shows the raw formula text, not the computed value
  // When editing, show the current editValue; otherwise show the cell's raw content
  // Note: This is different from the cell display, which shows computed values
  const formulaBarValue = editingCell
    ? editValue
    : (selectedCell ? engine.getCell(selectedCell.r, selectedCell.c).raw : '')

  // ────── Render ──────

  return (
    <div className="app-wrapper">
      <div className="app-header">
        <h2 className="app-title">📊 Spreadsheet App</h2>
      </div>

      <div className="main-content">

        {/* ── Toolbar ── */}
        <div className="toolbar">
          <div className="toolbar-group">
            <button className={`toolbar-btn bold-btn ${selectedCellStyle?.bold ? 'active' : ''}`} onClick={toggleBold} title="Bold">B</button>
            <button className={`toolbar-btn italic-btn ${selectedCellStyle?.italic ? 'active' : ''}`} onClick={toggleItalic} title="Italic">I</button>
            <button className={`toolbar-btn underline-btn ${selectedCellStyle?.underline ? 'active' : ''}`} onClick={toggleUnderline} title="Underline">U</button>
          </div>

          <div className="toolbar-group">
            <span className="toolbar-label">Size:</span>
            <select className="toolbar-select" value={selectedCellStyle?.fontSize || 13} onChange={(e) => changeFontSize(parseInt(e.target.value))}>
              {[8, 10, 11, 12, 13, 14, 16, 18, 20, 24].map(s => (
                <option key={s} value={s}>{s}</option>
              ))}
            </select>
          </div>

          <div className="toolbar-group">
            <button className={`align-btn ${selectedCellStyle?.align === 'left' ? 'active' : ''}`} onClick={() => changeAlignment('left')} title="Align Left">⬤←</button>
            <button className={`align-btn ${selectedCellStyle?.align === 'center' ? 'active' : ''}`} onClick={() => changeAlignment('center')} title="Align Center">⬤</button>
            <button className={`align-btn ${selectedCellStyle?.align === 'right' ? 'active' : ''}`} onClick={() => changeAlignment('right')} title="Align Right">⬤→</button>
          </div>

          <div className="toolbar-group">
            <span className="toolbar-label">Text:</span>
            <input
              type="color"
              value={selectedCellStyle?.color || '#000000'}
              onChange={(e) => changeFontColor(e.target.value)}
              title="Font color"
              style={{ width: '32px', height: '32px', border: '1px solid #dadce0', cursor: 'pointer', borderRadius: '4px' }}
            />
          </div>

          <div className="toolbar-group">
            <span className="toolbar-label">Fill:</span>
            <select className="toolbar-select" value={selectedCellStyle?.bg || 'white'} onChange={(e) => changeBackgroundColor(e.target.value)}>
              <option value="white">White</option>
              <option value="#ffff99">Yellow</option>
              <option value="#99ffcc">Green</option>
              <option value="#ffcccc">Red</option>
              <option value="#cce5ff">Blue</option>
              <option value="#e0ccff">Purple</option>
              <option value="#ffd9b3">Orange</option>
              <option value="#f0f0f0">Gray</option>
            </select>
          </div>

          <div className="toolbar-group">
            <button className="toolbar-btn" onClick={handleUndo} disabled={!engine.canUndo()} title="Undo">↶ Undo</button>
            <button className="toolbar-btn" onClick={handleRedo} disabled={!engine.canRedo()} title="Redo">↷ Redo</button>
          </div>

          <div className="toolbar-group">
            <button className="toolbar-btn" onClick={insertRow} title="Insert Row">+ Row</button>
            <button className="toolbar-btn" onClick={deleteRow} title="Delete Row">- Row</button>
            <button className="toolbar-btn" onClick={insertColumn} title="Insert Column">+ Col</button>
            <button className="toolbar-btn" onClick={deleteColumn} title="Delete Column">- Col</button>
          </div>

          <div className="toolbar-group">
            <button className="toolbar-btn danger" onClick={clearSelectedCell}>✕ Cell</button>
            <button className="toolbar-btn danger" onClick={clearAllCells}>✕ All</button>
          </div>
        </div>

        {/* ── Formula Bar ── */}
        <div className="formula-bar">
          <span className="formula-bar-label">{selectedCellLabel}</span>
          <input
            className="formula-bar-input"
            value={formulaBarValue}
            onChange={(e) => handleFormulaBarChange(e.target.value)}
            onKeyDown={handleFormulaBarKeyDown}
            onFocus={handleFormulaBarFocus}
            placeholder="Select a cell then type, or enter a formula like =SUM(A1:A5)"
          />
        </div>

        {/* ── Grid ── */}
        <div className="grid-scroll">
          <table className="grid-table">
            <thead>
              <tr>
                <th className="col-header-blank"></th>
                {Array.from({ length: engine.cols }, (_, colIndex) => (
                  <th
                    key={colIndex}
                    className="col-header"
                    onClick={() => handleSortClick(colIndex)}
                  >
                    <div className="col-header-inner">
                      <span className="col-header-label">
                        {getColumnLabel(colIndex)}
                        {sortConfig.column === colIndex && sortConfig.direction !== 'none' && (
                          <span className="sort-indicator">
                            {sortConfig.direction === 'asc' ? '▲' : '▼'}
                          </span>
                        )}
                      </span>
                      <button
                        className="filter-btn"
                        onClick={(e) => {
                          e.stopPropagation()
                          setOpenFilterCol(openFilterCol === colIndex ? null : colIndex)
                        }}
                        title="Filter"
                      >
                        ⋮
                      </button>
                    </div>
                    {openFilterCol === colIndex && (
                      <div
                        className="filter-dropdown"
                        onClick={(e) => e.stopPropagation()}
                      >
                        <div className="filter-dropdown-header">
                          Filter by value
                        </div>
                        <div className="filter-options">
                          {getUniqueValuesForColumn(colIndex).map((val, idx) => {
                            const selectedList = filters[String(colIndex)] || []
                            const selected = selectedList.some(v => Object.is(v, val))
                            const label =
                              val === null || val === '' || typeof val === 'undefined'
                                ? '(blank)'
                                : String(val)
                            return (
                              <label key={idx} className="filter-option">
                                <input
                                  type="checkbox"
                                  checked={selected}
                                  onChange={() => toggleFilterValue(colIndex, val)}
                                />
                                {label}
                              </label>
                            )
                          })}
                        </div>
                        <button
                          className="filter-clear-btn"
                          onClick={() => {
                            clearFilter(colIndex)
                            setOpenFilterCol(null)
                          }}
                        >
                          Clear Filter
                        </button>
                      </div>
                    )}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {rowOrder
                .filter((rowIndex) => passesFilters(rowIndex))
                .map((rowIndex, displayIndex) => (
                <tr key={rowIndex}>
                  <td className="row-header">{displayIndex + 1}</td>
                  {Array.from({ length: engine.cols }, (_, colIndex) => {
                    const isSelected = selectedCell?.r === rowIndex && selectedCell?.c === colIndex
                    const isEditing = editingCell?.r === rowIndex && editingCell?.c === colIndex
                    const cellData = engine.getCell(rowIndex, colIndex)
                    const style = cellStyles[`${rowIndex},${colIndex}`] || {}
                    const displayValue = cellData.error
                      ? cellData.error
                      : (cellData.computed !== null && cellData.computed !== '' ? String(cellData.computed) : cellData.raw)

                    return (
                      <td
                        key={colIndex}
                        className={`cell ${isSelected ? 'selected' : ''}`}
                        style={{ background: style.bg || 'white' }}
                        onMouseDown={(e) => { e.preventDefault(); handleCellClick(rowIndex, colIndex) }}
                      >
                        {isEditing ? (
                          <input
                            autoFocus
                            className="cell-input"
                            value={editValue}
                            onChange={(e) => setEditValue(e.target.value)}
                            onBlur={() => commitEdit(rowIndex, colIndex)}
                            onKeyDown={(e) => handleKeyDown(e, rowIndex, colIndex)}
                            ref={isSelected ? cellInputRef : undefined}
                            style={{
                              fontWeight: style.bold ? 'bold' : 'normal',
                              fontStyle: style.italic ? 'italic' : 'normal',
                              textDecoration: style.underline ? 'underline' : 'none',
                              color: style.color || '#202124',
                              fontSize: (style.fontSize || 13) + 'px',
                              textAlign: style.align || 'left',
                              background: style.bg || 'white',
                            }}
                          />
                        ) : (
                          <div
                            className={`cell-display align-${style.align || 'left'} ${cellData.error ? 'error' : ''}`}
                            style={{
                              fontWeight: style.bold ? 'bold' : 'normal',
                              fontStyle: style.italic ? 'italic' : 'normal',
                              textDecoration: style.underline ? 'underline' : 'none',
                              color: cellData.error ? '#d93025' : (style.color || '#202124'),
                              fontSize: (style.fontSize || 13) + 'px',
                            }}
                          >
                            {displayValue}
                          </div>
                        )}
                      </td>
                    )
                  })}
                </tr>
              ))}
            </tbody>
          </table>
        </div>

        <p className="footer-hint">
          Click a cell to edit · Enter/Tab/Arrow keys to navigate · Formulas: =A1+B1 · =SUM(A1:A5) · =AVG(A1:A5) · =MAX(A1:A5) · =MIN(A1:A5)
        </p>
      </div>
    </div>
  )
}
