<script setup>
import { computed, nextTick, onBeforeUnmount, onMounted, reactive, ref } from 'vue'
import { removeDuplicateRows, switchSheet, uploadWorkbook, validateWorkbook } from './services/api'

const API_BASE = import.meta.env.VITE_API_BASE ?? '/api'
const fileInput = ref(null)
const sessionId = ref('')
const fileName = ref('')
const sheetNames = ref([])
const sheetName = ref('')
const columns = ref([])
const rows = ref([])
const errors = ref([])
const duplicateGroups = ref([])
const dragActive = ref(false)
const statusMessage = ref('Upload an Excel workbook to begin validating.')
const statusTone = ref('info')
const isUploading = ref(false)
const isValidating = ref(false)
const history = reactive({
  past: [],
  future: [],
  limit: 5,
})

const typeOptions = [
  { value: 'decimal', label: 'Decimal number', base: 'float' },
  { value: 'integer', label: 'Whole number', base: 'integer' },
  { value: 'percentage', label: 'Percentage', base: 'float' },
  { value: 'currency', label: 'Currency', base: 'float' },
  { value: 'datetime', label: 'Date/time', base: 'date' },
  { value: 'date', label: 'Date', base: 'date' },
  { value: 'text', label: 'Text', base: 'string' },
  { value: 'boolean', label: 'True/False', base: 'boolean' },
]
const typeOptionMap = typeOptions.reduce((map, option) => {
  map[option.value] = option
  return map
}, {})
const baseTypeLabels = {
  string: 'Text',
  integer: 'Whole number',
  float: 'Decimal number',
  boolean: 'True/False',
  date: 'Date',
}
const formatLabels = {
  lowercase: 'lowercase',
  uppercase: 'uppercase',
  capitalise: 'capitalized',
  trim: 'trimmed',
}
const hasSession = computed(() => !!sessionId.value)
const datasetLoaded = computed(() => rows.value.length > 0)
const invalidCellMap = computed(() => {
  const set = new Set()
  errors.value.forEach((error) => {
    set.add(`${error.rowId}::${error.column}`)
  })
  return set
})
const duplicateRowIds = computed(() => {
  const ids = new Set()
  duplicateGroups.value.forEach((group) => group.forEach((rowId) => ids.add(rowId)))
  return ids
})
const errorRowIds = computed(() => {
  const ids = new Set()
  errors.value.forEach((error) => ids.add(error.rowId))
  return ids
})
const filterPairs = computed(() =>
  Object.entries(columnFilters).map(([column, values]) => [column, values]),
)
const filteredRows = computed(() => {
  let result = rows.value
  if (filterPairs.value.length) {
    result = result.filter((row) =>
      filterPairs.value.every(([column, values]) => {
        const cell = row.values[column]
        if (cell === undefined || cell === null) return false
        if (!values.size) return true
        return values.has(cell.toString())
      }),
    )
  }
  if (showErrorsOnly.value) {
    result = result.filter((row) => errorRowIds.value.has(row.rowId))
  }
  return result
})
const overrides = computed(() => {
  const map = {}
  columns.value.forEach((column) => {
    if (column.overrideType) {
      map[column.name] = getBaseType(column.overrideType)
    }
  })
  return map
})
const stats = computed(() => ({
  errors: errors.value.length,
  duplicates: duplicateGroups.value.length,
  rows: filteredRows.value.length,
  columns: columns.value.length,
}))
const editingColumn = ref('')
const editingValue = ref('')
const headerInputRef = ref(null)
const focusedDuplicateRows = ref(new Set())
const focusedDuplicateGroups = ref([])
const selectedColumns = ref(new Set())
const lastSelectedIndex = ref(-1)
const selectedRows = ref(new Set())
const lastSelectedRowIndex = ref(-1)
const cellWarnings = ref([])
const showErrorsOnly = ref(false)
const columnFilters = reactive({})
const filterDrafts = reactive({})
const filterSearchTerms = reactive({})
const filterMenuOpen = ref('')
const selectedColumnLabel = computed(() => {
  const names = Array.from(selectedColumns.value)
  if (!names.length) {
    return 'None'
  }
  if (names.length <= 3) {
    return names.join(', ')
  }
  return `${names.length} columns selected`
})
const hasSelectedColumns = computed(() => selectedColumns.value.size > 0)
const hasSelectedRows = computed(() => selectedRows.value.size > 0)
const canUndo = computed(() => history.past.length > 1)
const canRedo = computed(() => history.future.length > 0)
const hasMultipleSheets = computed(() => sheetNames.value.length > 1)
const handleGlobalPointerDown = (event) => {
  const target = event.target
  if (
    filterMenuOpen.value &&
    !(target instanceof Element && (target.closest('.filter-popover') || target.closest('.header-filter')))
  ) {
    filterMenuOpen.value = ''
  }
  if (!editingColumn.value) return
  if (headerInputRef.value && target instanceof Element && headerInputRef.value.contains(target)) {
    return
  }
  finishColumnRename(true)
}

onMounted(() => {
  document.addEventListener('pointerdown', handleGlobalPointerDown, true)
})

onBeforeUnmount(() => {
  document.removeEventListener('pointerdown', handleGlobalPointerDown, true)
})

const uploadZoneHandlers = reactive({
  dragover: (event) => {
    event.preventDefault()
    dragActive.value = true
  },
  dragleave: (event) => {
    event.preventDefault()
    dragActive.value = false
  },
  drop: (event) => {
    event.preventDefault()
    dragActive.value = false
    const [file] = event.dataTransfer.files
    if (file) {
      processFile(file)
    }
  },
})

function openFileDialog() {
  fileInput.value?.click()
}

async function handleFileInput(event) {
  const [file] = event.target.files
  if (file) {
    await processFile(file)
    event.target.value = ''
  }
}

async function processFile(file) {
  try {
    isUploading.value = true
    setStatus(`Uploading ${file.name}...`, 'info')
    const payload = await uploadWorkbook(file)
    fileName.value = file.name
    sessionId.value = payload.sessionId
    applyPayload(payload, true)
    setStatus('Upload complete. Review highlighted cells or adjust column types.', 'success')
  } catch (error) {
    console.error(error)
    setStatus(error.message || 'Upload failed', 'error')
  } finally {
    isUploading.value = false
  }
}

async function handleSheetChange(event) {
  const nextSheet = event.target.value
  if (!sessionId.value || !nextSheet || nextSheet === sheetName.value) return
  try {
    setStatus(`Loading sheet "${nextSheet}"...`, 'info')
    const payload = await switchSheet(sessionId.value, nextSheet)
    applyPayload(payload, true)
    setStatus(`Loaded sheet "${nextSheet}".`, 'success')
  } catch (error) {
    console.error(error)
    setStatus(error.message || 'Failed to switch sheet', 'error')
  }
}

async function revalidateData() {
  if (!sessionId.value) return
  try {
    isValidating.value = true
    setStatus('Re-validating grid...', 'info')
    const payload = await validateWorkbook(
      sessionId.value,
      serializeRows(rows.value),
      overrides.value,
      columns.value,
    )
    applyPayload(payload)
    setStatus('Validation updated. Download a report or continue fixing cells.', 'success')
  } catch (error) {
    console.error(error)
    setStatus(error.message || 'Validation failed', 'error')
  } finally {
    isValidating.value = false
  }
}

async function removeDuplicates(group) {
  if (!sessionId.value || group.length <= 1) return
  const [, ...toRemove] = group
  if (!toRemove.length) {
    setStatus('Nothing to remove. The first row in each group is preserved.', 'info')
    return
  }
  try {
    setStatus('Removing selected duplicates...', 'info')
    const payload = await removeDuplicateRows(sessionId.value, toRemove)
    applyPayload(payload)
    pushHistoryState()
    setStatus('Duplicates removed. Re-run validation to confirm clean data.', 'success')
  } catch (error) {
    console.error(error)
    setStatus(error.message || 'Could not remove duplicates', 'error')
  }
}

async function triggerReportDownload() {
  try {
    setStatus('Preparing Excel report...', 'info')
    const response = await fetch(`${API_BASE}/report`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        rows: serializeRows(rows.value),
        columns: columns.value,
        errors: errors.value,
      }),
    })
    if (!response.ok) {
      const detail = await response.json().catch(() => ({}))
      throw new Error(detail?.detail || 'Failed to download report')
    }
    const blob = await response.blob()
    downloadBlob(blob, `${fileName.value?.replace(/\.[^.]+$/, '') || 'validation'}-report.xlsx`)
    setStatus('Report downloaded successfully.', 'success')
  } catch (error) {
    console.error(error)
    setStatus(error.message || 'Failed to download report', 'error')
  }
}

async function triggerEditedSheetDownload() {
  if (!sessionId.value) return
  try {
    setStatus('Exporting edited sheet...', 'info')
    const response = await fetch(`${API_BASE}/export`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        sessionId: sessionId.value,
        rows: serializeRows(rows.value),
        columns: columns.value,
      }),
    })
    if (!response.ok) {
      const detail = await response.json().catch(() => ({}))
      throw new Error(detail?.detail || 'Failed to export sheet')
    }
    const blob = await response.blob()
    downloadBlob(blob, `${fileName.value?.replace(/\\.[^.]+$/, '') || 'validation'}-edited.xlsx`)
    setStatus('Edited sheet downloaded successfully.', 'success')
  } catch (error) {
    console.error(error)
    setStatus(error.message || 'Failed to export sheet', 'error')
  }
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob)
  const anchor = document.createElement('a')
  anchor.href = url
  anchor.download = filename
  document.body.appendChild(anchor)
  anchor.click()
  document.body.removeChild(anchor)
  URL.revokeObjectURL(url)
}

function detectTrailingSpaces() {
  const offenders = []
  rows.value.forEach((row, rowIndex) => {
    Object.entries(row.values).forEach(([column, value]) => {
      if (typeof value === 'string' && value !== value.trim()) {
        offenders.push({ rowId: row.rowId ?? rowIndex, column })
      }
    })
  })
  if (!offenders.length) {
    setStatus('No trailing spaces detected.', 'success')
    return
  }
  const formatted = offenders
    .slice(0, 5)
    .map((entry) => `Row ${entry.rowId + 1} • ${entry.column}`)
    .join(' | ')
  const suffix = offenders.length > 5 ? ` …and ${offenders.length - 5} more.` : '.'
  setStatus(`Trailing spaces found: ${formatted}${suffix}`, 'error')
  cellWarnings.value = offenders
}

function applyPayload(payload, resetHistory = false) {
  columns.value = (payload.columns || []).map((column) => ({ ...column }))
  rows.value = (payload.rows || []).map((row, index) => ({
    rowId: row.rowId ?? index,
    values: { ...row.values },
  }))
  errors.value = payload.errors || []
  duplicateGroups.value = payload.duplicateGroups || []
  clearDuplicateFocus()
  selectedColumns.value = new Set()
  lastSelectedIndex.value = -1
  selectedRows.value = new Set()
  lastSelectedRowIndex.value = -1
  cellWarnings.value = []
  Object.keys(columnFilters).forEach((key) => delete columnFilters[key])
  filterMenuOpen.value = ''
  showErrorsOnly.value = false
  if (Array.isArray(payload.sheetNames) && payload.sheetNames.length) {
    sheetNames.value = payload.sheetNames.slice()
  }
  if (typeof payload.sheetName === 'string') {
    sheetName.value = payload.sheetName
  } else if (!sheetName.value && sheetNames.value.length) {
    [sheetName.value] = sheetNames.value
  }
  if (resetHistory) {
    history.past.splice(0, history.past.length, snapshotState())
    history.future.splice(0, history.future.length)
  }
}

function updateCell(rowId, columnName, newValue) {
  const targetRow = rows.value.find((row) => row.rowId === rowId)
  if (!targetRow) return
  pushHistoryState()
  targetRow.values[columnName] = newValue
}

function setStatus(message, tone = 'info') {
  statusMessage.value = message
  statusTone.value = tone
}

function snapshotState() {
  return {
    rows: serializeRows(rows.value),
    columns: columns.value.map((column) => ({ ...column })),
    selectedColumns: Array.from(selectedColumns.value),
    selectedRows: Array.from(selectedRows.value),
    columnFilters: Object.fromEntries(
      Object.entries(columnFilters).map(([key, value]) => [key, new Set(value)]),
    ),
    warnings: cellWarnings.value.map((warning) => ({ ...warning })),
    sheetName: sheetName.value,
    sheetNames: sheetNames.value.slice(),
  }
}

function pushHistoryState(state = null) {
  const snapshot = state ?? snapshotState()
  history.past.push(snapshot)
  if (history.past.length > history.limit) {
    history.past.shift()
  }
  history.future.splice(0, history.future.length)
}

function restoreState(state) {
  rows.value = state.rows.map((row) => ({ rowId: row.rowId, values: { ...row.values } }))
  columns.value = state.columns.map((column) => ({ ...column }))
  selectedColumns.value = new Set(state.selectedColumns)
  selectedRows.value = new Set(state.selectedRows)
  Object.keys(columnFilters).forEach((key) => delete columnFilters[key])
  Object.entries(state.columnFilters).forEach(([key, value]) => {
    columnFilters[key] = new Set(value)
  })
  lastSelectedIndex.value = -1
  lastSelectedRowIndex.value = -1
  clearDuplicateFocus()
  cellWarnings.value = state.warnings ? state.warnings.map((warning) => ({ ...warning })) : []
  if (state.sheetName) {
    sheetName.value = state.sheetName
  }
  if (Array.isArray(state.sheetNames) && state.sheetNames.length) {
    sheetNames.value = state.sheetNames.slice()
  }
}

function undo() {
  if (history.past.length <= 1) {
    setStatus('Nothing to undo.', 'info')
    return
  }
  const current = history.past.pop()
  history.future.push(current)
  const previous = history.past[history.past.length - 1]
  restoreState(previous)
  setStatus('Undid last change.', 'info')
}

function redo() {
  if (!history.future.length) {
    setStatus('Nothing to redo.', 'info')
    return
  }
  const next = history.future.pop()
  history.past.push(next)
  restoreState(next)
  setStatus('Redid change.', 'info')
}

function getBaseType(typeValue) {
  if (!typeValue) return typeValue
  const option = typeOptionMap[typeValue]
  return option ? option.base || typeValue : typeValue
}

function getTypeLabel(typeValue) {
  if (!typeValue) return 'Auto'
  const option = typeOptionMap[typeValue]
  if (option) return option.label
  return baseTypeLabels[typeValue] || typeValue
}

function serializeRows(sourceRows) {
  return sourceRows.map((row, index) => ({
    rowId: row.rowId ?? index,
    values: { ...row.values },
  }))
}

function cellClass(rowId, columnName) {
  const classes = []
  if (invalidCellMap.value.has(`${rowId}::${columnName}`)) {
    classes.push('cell-invalid')
  }
  if (duplicateRowIds.value.has(rowId)) {
    classes.push('cell-duplicate')
  }
  if (focusedDuplicateGroups.value.length) {
    const group = focusedDuplicateGroups.value.find(
      (entry) => entry.column === columnName && entry.rows.includes(rowId),
    )
    if (group) {
      classes.push(`cell-duplicate-group-${group.colorIndex}`)
    }
  } else if (focusedDuplicateRows.value.has(rowId)) {
    classes.push('cell-duplicate-focus')
  }
  if (selectedColumns.value.has(columnName)) {
    classes.push('cell-selected-column')
  }
  if (cellWarnings.value.some((warning) => warning.rowId === rowId && warning.column === columnName)) {
    classes.push('cell-warning')
  }
  return classes.join(' ')
}

function startColumnRename(columnName) {
  editingColumn.value = columnName
  editingValue.value = columnName
  nextTick(() => {
    headerInputRef.value?.focus()
    headerInputRef.value?.select()
  })
}

function finishColumnRename(submit = true) {
  const original = editingColumn.value
  const nextValue = editingValue.value.trim()
  editingColumn.value = ''
  if (!submit || !original) {
    return
  }
  if (!nextValue || nextValue === original) {
    return
  }
  applyColumnRename(original, nextValue)
}

function applyColumnRename(oldName, newName) {
  const column = columns.value.find((col) => col.name === oldName)
  if (!column) return
  const exists = columns.value.some((col) => col.name === newName && col !== column)
  if (exists) {
    setStatus('A column with that name already exists.', 'error')
    return
  }
  column.name = newName
  rows.value = rows.value.map((row) => {
    const value = row.values[oldName]
    const updatedValues = { ...row.values }
    delete updatedValues[oldName]
    updatedValues[newName] = value
    return { ...row, values: updatedValues }
  })
  errors.value = errors.value.map((error) =>
    error.column === oldName ? { ...error, column: newName } : error,
  )
  duplicateGroups.value = duplicateGroups.value.map((group) => [...group])
  if (selectedColumns.value.has(oldName)) {
    const updatedSelection = new Set(selectedColumns.value)
    updatedSelection.delete(oldName)
    updatedSelection.add(newName)
    selectedColumns.value = updatedSelection
  }
  pushHistoryState()
}

async function deleteColumn(columnName) {
  if (!columns.value.length) return
  const confirmed = window.confirm(`Delete column "${columnName}"?`)
  if (!confirmed) return
  columns.value = columns.value.filter((column) => column.name !== columnName)
  rows.value = rows.value.map((row) => {
    const updated = { ...row.values }
    delete updated[columnName]
    return { ...row, values: updated }
  })
  errors.value = errors.value.filter((error) => error.column !== columnName)
  if (selectedColumns.value.has(columnName)) {
    const updatedSelection = new Set(selectedColumns.value)
    updatedSelection.delete(columnName)
    selectedColumns.value = updatedSelection
    if (!updatedSelection.size) {
      lastSelectedIndex.value = -1
    }
  }
  pushHistoryState()
  await revalidateData()
  setStatus(`Column "${columnName}" deleted.`, 'success')
}

function clearDuplicateFocus() {
  focusedDuplicateRows.value = new Set()
  focusedDuplicateGroups.value = []
  cellWarnings.value = []
}

function highlightColumnsDuplicates(columnNames) {
  const groups = []
  let groupIndex = 0
  columnNames.forEach((columnName) => {
    const valueMap = new Map()
    rows.value.forEach((row) => {
      const value = row.values[columnName]
      const key = value ?? '__undefined__'
      if (!valueMap.has(key)) {
        valueMap.set(key, [])
      }
      valueMap.get(key).push(row.rowId)
    })
    valueMap.forEach((ids) => {
      if (ids.length > 1) {
        groups.push({ column: columnName, rows: ids, colorIndex: groupIndex % 10 })
        groupIndex += 1
      }
    })
  })
  if (!groups.length) {
    clearDuplicateFocus()
    setStatus(`No duplicates found in ${formatColumnList(columnNames)}.`, 'success')
    return
  }
  focusedDuplicateRows.value = new Set(groups.flatMap((group) => group.rows))
  focusedDuplicateGroups.value = groups
}

function selectColumn(columnName, event) {
  if (editingColumn.value) return
  const index = columns.value.findIndex((column) => column.name === columnName)
  if (index === -1) return
  const ctrl = event?.ctrlKey || event?.metaKey
  const shift = event?.shiftKey
  let nextSelection = new Set(selectedColumns.value)
  if (shift && lastSelectedIndex.value !== -1) {
    const start = Math.min(lastSelectedIndex.value, index)
    const end = Math.max(lastSelectedIndex.value, index)
    nextSelection = new Set()
    for (let i = start; i <= end; i += 1) {
      const targetColumn = columns.value[i]
      if (targetColumn) {
        nextSelection.add(targetColumn.name)
      }
    }
  } else if (ctrl) {
    if (nextSelection.has(columnName)) {
      nextSelection.delete(columnName)
    } else {
      nextSelection.add(columnName)
    }
    lastSelectedIndex.value = index
  } else {
    nextSelection = new Set([columnName])
    lastSelectedIndex.value = index
  }
  selectedColumns.value = nextSelection
  if (!nextSelection.size) {
    lastSelectedIndex.value = -1
  }
  if (!ctrl && !shift) {
    clearDuplicateFocus()
  }
  if (!selectedColumns.value.size) {
    clearDuplicateFocus()
  }
}

function toggleFilterMenu(columnName) {
  if (filterMenuOpen.value === columnName) {
    filterMenuOpen.value = ''
    return
  }
  filterSearchTerms[columnName] = ''
  if (columnFilters[columnName]) {
    filterDrafts[columnName] = new Set(columnFilters[columnName])
  } else {
    filterDrafts[columnName] = new Set(getColumnUniqueValues(columnName))
  }
  filterMenuOpen.value = columnName
}

function getColumnUniqueValues(columnName) {
  const search = (filterSearchTerms[columnName] || '').trim().toLowerCase()
  const values = new Set()
  rows.value.forEach((row) => {
    const value = row.values[columnName]
    if (value === undefined || value === null) return
    const text = value.toString()
    if (!search || text.toLowerCase().includes(search)) {
      values.add(text)
    }
  })
  return Array.from(values).sort((a, b) => a.localeCompare(b))
}

function toggleFilterValue(columnName, value, checked) {
  if (!filterDrafts[columnName]) {
    filterDrafts[columnName] = new Set()
  }
  if (checked) {
    filterDrafts[columnName].add(value)
  } else {
    filterDrafts[columnName].delete(value)
  }
}

function applyFilter(columnName) {
  const draft = getFilterDraft(columnName)
  const values = getColumnUniqueValues(columnName)
  if (!draft.size || draft.size === values.length) {
    delete columnFilters[columnName]
  } else {
    columnFilters[columnName] = new Set(draft)
  }
  delete filterDrafts[columnName]
  filterMenuOpen.value = ''
}

function clearFilter(columnName) {
  delete columnFilters[columnName]
  delete filterDrafts[columnName]
  delete filterSearchTerms[columnName]
  if (filterMenuOpen.value === columnName) {
    filterMenuOpen.value = ''
  }
}

function selectRow(rowId, event) {
  const visibleRows = filteredRows.value
  const index = visibleRows.findIndex((row) => row.rowId === rowId)
  if (index === -1) return
  const ctrl = event?.ctrlKey || event?.metaKey
  const shift = event?.shiftKey
  let nextSelection = new Set(selectedRows.value)
  if (shift && lastSelectedRowIndex.value !== -1) {
    const start = Math.min(lastSelectedRowIndex.value, index)
    const end = Math.max(lastSelectedRowIndex.value, index)
    nextSelection = new Set()
    for (let i = start; i <= end; i += 1) {
      const targetRow = visibleRows[i]
      if (targetRow) {
        nextSelection.add(targetRow.rowId)
      }
    }
  } else if (ctrl) {
    if (nextSelection.has(rowId)) {
      nextSelection.delete(rowId)
    } else {
      nextSelection.add(rowId)
    }
    lastSelectedRowIndex.value = index
  } else {
    nextSelection = new Set([rowId])
    lastSelectedRowIndex.value = index
  }
  selectedRows.value = nextSelection
  if (!nextSelection.size) {
    lastSelectedRowIndex.value = -1
  }
}

function ensureColumnSelection() {
  if (!selectedColumns.value.size) {
    setStatus('Select at least one column first.', 'error')
    return false
  }
  return true
}

function getSelectedColumnNames() {
  return Array.from(selectedColumns.value)
}

function formatColumnList(names) {
  if (!names.length) return ''
  if (names.length <= 3) return names.join(', ')
  return `${names.length} columns`
}

function getFilterDraft(columnName) {
  if (!filterDrafts[columnName]) {
    if (columnFilters[columnName]) {
      filterDrafts[columnName] = new Set(columnFilters[columnName])
    } else {
      filterDrafts[columnName] = new Set(getColumnUniqueValues(columnName))
    }
  }
  return filterDrafts[columnName]
}

function isSelectAllChecked(columnName) {
  const values = getColumnUniqueValues(columnName)
  if (!values.length) return true
  const draft = getFilterDraft(columnName)
  return values.every((value) => draft.has(value))
}

function handleSelectAllToggle(columnName, checked) {
  const draft = getFilterDraft(columnName)
  if (checked) {
    getColumnUniqueValues(columnName).forEach((value) => draft.add(value))
  } else {
    draft.clear()
  }
}

async function deleteSelectedColumns() {
  if (!ensureColumnSelection()) return false
  const names = getSelectedColumnNames()
  const confirmed = window.confirm(`Delete the selected column(s)?\n${names.join(', ')}`)
  if (!confirmed) return false
  const removalSet = new Set(names)
  columns.value = columns.value.filter((column) => !removalSet.has(column.name))
  rows.value = rows.value.map((row) => {
    const updated = { ...row.values }
    names.forEach((name) => {
      delete updated[name]
    })
    return { ...row, values: updated }
  })
  errors.value = errors.value.filter((error) => !removalSet.has(error.column))
  selectedColumns.value = new Set()
  lastSelectedIndex.value = -1
  clearDuplicateFocus()
  pushHistoryState()
  await revalidateData()
  setStatus(`Deleted ${formatColumnList(names)}.`, 'success')
  return true
}

async function deleteSelectedRows() {
  if (!selectedRows.value.size) return false
  const rowIds = Array.from(selectedRows.value).sort((a, b) => a - b)
  const confirmed =
    rowIds.length === 1
      ? window.confirm(`Delete row ${rowIds[0] + 1}?`)
      : window.confirm(`Delete ${rowIds.length} selected rows?`)
  if (!confirmed) return false
  const removalSet = new Set(rowIds)
  const remaining = rows.value.filter((row) => !removalSet.has(row.rowId))
  rows.value = remaining.map((row, index) => ({
    rowId: index,
    values: { ...row.values },
  }))
  selectedRows.value = new Set()
  lastSelectedRowIndex.value = -1
  pushHistoryState()
  await revalidateData()
  setStatus(
    rowIds.length === 1 ? `Deleted row ${rowIds[0] + 1}.` : `Deleted ${rowIds.length} rows.`,
    'success',
  )
  return true
}

function toggleErrorFilter() {
  if (!errors.value.length) return
  showErrorsOnly.value = !showErrorsOnly.value
  setStatus(
    showErrorsOnly.value ? 'Showing only rows with errors.' : 'Showing all rows.',
    'info',
  )
}

function checkSelectedColumnsDuplicates() {
  if (!ensureColumnSelection()) return
  highlightColumnsDuplicates(getSelectedColumnNames())
}

async function deleteSelection() {
  if (await deleteSelectedRows()) return
  if (await deleteSelectedColumns()) return
  setStatus('Select at least one column or row first.', 'error')
}

async function applyToSelectedColumns(transform, messageTemplate) {
  if (!ensureColumnSelection()) return
  const names = getSelectedColumnNames()
  rows.value = rows.value.map((row) => {
    const updated = { ...row.values }
    names.forEach((name) => {
      updated[name] = transform(updated[name], name)
    })
    return { ...row, values: updated }
  })
  pushHistoryState()
  await revalidateData()
  setStatus(messageTemplate.replace('{columns}', formatColumnList(names)), 'success')
}

function isColumnSelected(columnName) {
  return selectedColumns.value.has(columnName)
}

function isRowSelected(rowId) {
  return selectedRows.value.has(rowId)
}

function setSelectedColumnsType(type) {
  if (!ensureColumnSelection()) return null
  const names = getSelectedColumnNames()
  names.forEach((name) => {
    const column = columns.value.find((col) => col.name === name)
    if (column) {
      column.overrideType = type
    }
  })
  pushHistoryState()
  return names
}

async function handleRibbonTypeChange(event) {
  const value = event.target.value
  event.target.value = ''
  if (!value) return
  const type = value === 'auto' ? null : value
  const names = setSelectedColumnsType(type)
  if (!names) return
  await revalidateData()
  if (type) {
    setStatus(`Set ${formatColumnList(names)} to ${getTypeLabel(type)}.`, 'info')
  } else {
    setStatus(`Reverted ${formatColumnList(names)} to auto detection.`, 'info')
  }
}

async function handleFormatChange(event) {
  const value = event.target.value
  event.target.value = ''
  if (!value) return
  const formatters = {
    lowercase: (text) => text.toLowerCase(),
    uppercase: (text) => text.toUpperCase(),
    capitalise: (text) =>
      text
        .toLowerCase()
        .split(' ')
        .map((word) => word.charAt(0).toUpperCase() + word.slice(1))
        .join(' '),
    trim: (text) => text.trim(),
  }
  const formatter = formatters[value]
  if (!formatter) return
  await applyToSelectedColumns(
    (cell) => (typeof cell === 'string' ? formatter(cell) : cell),
    `Applied ${formatLabels[value] || value} format to {columns}.`,
  )
}
</script>

<template>
  <div class="page">
    <header class="page-header glass-panel">
  <div>
        <h1>Upload, validate, and fix your spreadsheets in one place.</h1>
      </div>
      <div class="stats">
        <div class="stat">
          <span>{{ stats.rows }}</span>
          <label>rows</label>
        </div>
        <div class="stat">
          <span>{{ stats.columns }}</span>
          <label>columns</label>
        </div>
        <div class="stat">
          <span>{{ stats.errors }}</span>
          <label>errors</label>
        </div>
        <div class="stat">
          <span>{{ stats.duplicates }}</span>
          <label>dup groups</label>
        </div>
      </div>
    </header>

    <section class="upload-zone glass-panel" :class="{ 'upload-zone--active': dragActive }" v-on="uploadZoneHandlers" @click="openFileDialog">
      <input ref="fileInput" class="sr-only" type="file" accept=".xls,.xlsx,.xlsm,.csv" @change="handleFileInput" />
      <p class="upload-title">{{ hasSession ? 'Upload a new workbook' : 'Drop your Excel workbook' }}</p>
      <p class="upload-subtitle">
        Drag and drop Excel sheet or click to browse. All files are stores locally in your browser.
      </p>
      <button class="primary" :disabled="isUploading" type="button">{{ isUploading ? 'Uploading...' : 'Select File' }}</button>
      <p v-if="fileName" class="file-name">Current file: {{ fileName }}</p>
      <div v-if="sheetNames.length" class="sheet-switcher" @click.stop>
        <label for="sheet-select" @click.stop>Sheet:</label>
        <select id="sheet-select" :value="sheetName" @click.stop @change="handleSheetChange">
          <option v-for="name in sheetNames" :key="name" :value="name">
            {{ name }}
          </option>
        </select>
      </div>
    </section>

    <section v-if="datasetLoaded" class="layout">
      <div class="grid-panel glass-panel">
        <div class="status" :data-tone="statusTone">
          <span>{{ statusMessage }}</span>
          <div class="status-actions">
            <button class="ghost" :disabled="!sessionId" type="button" @click="triggerReportDownload">
              Download error report
            </button>
            <button class="ghost" :disabled="!sessionId" type="button" @click="triggerEditedSheetDownload">
              Export edited sheet
            </button>
          </div>
        </div>


        <div class="ribbon glass-panel">
          
          <div class="ribbon-info">
            <div class="ribbon-actions ribbon-actions--top">
            <button class="ghost icon-button" :disabled="!canUndo" type="button" @click="undo" title="Undo">
              ↺
            </button>
            <button class="ghost icon-button" :disabled="!canRedo" type="button" @click="redo" title="Redo">
              ↻
            </button>
            
          </div>
            <p>Selected columns:</p>
            <strong>{{ selectedColumnLabel }}</strong>
          </div>
          <div class="ribbon-actions">
            <select class="ribbon-select" :disabled="!hasSelectedColumns" @change="handleRibbonTypeChange">
              <option value="">Set data type</option>
              <option value="auto">Auto detect</option>
              <option v-for="option in typeOptions" :key="option.value" :value="option.value">
                {{ option.label }}
              </option>
            </select>
            <select class="ribbon-select" :disabled="!hasSelectedColumns" @change="handleFormatChange">
              <option value="">Format</option>
              <option value="lowercase">lowercase</option>
              <option value="uppercase">UPPERCASE</option>
              <option value="capitalise">Capitalise Each Word</option>
              <option value="trim">Trim</option>
            </select>
            <button class="ghost" :disabled="!hasSelectedColumns" type="button" @click="checkSelectedColumnsDuplicates">
              Check duplicates
            </button>
            <button class="ghost" :disabled="!sessionId" type="button" @click="detectTrailingSpaces">
              Detect trailing spaces
            </button>
            <button
              class="ghost"
              :class="{ 'ghost--active': showErrorsOnly }"
              :disabled="!errors.length"
              type="button"
              @click="toggleErrorFilter"
            >
              {{ showErrorsOnly ? 'Show all rows' : 'Only errors' }}
            </button>
            <button
              class="danger"
              :disabled="!hasSelectedColumns && !hasSelectedRows"
              type="button"
              @click="deleteSelection"
            >
              Delete
            </button>
          </div>
        </div>
        <div class="grid-window">
          <div class="grid-window__controls">
            
          </div>
          <div class="grid-window__body">
            <div class="table-wrapper grid-window__scroll">
              <table>
                <thead>
                  <tr>
                    <th>#</th>
                    <th
                      v-for="column in columns"
                      :key="column.name"
                      :class="{ 'column-selected': isColumnSelected(column.name) }"
                      @click="selectColumn(column.name, $event)"
                    >
                      <div class="header-cell">
                        <div class="header-cell__titles">
                          <template v-if="editingColumn === column.name">
                            <input
                              ref="headerInputRef"
                              v-model="editingValue"
                              class="header-input"
                              type="text"
                              @click.stop
                              @blur="finishColumnRename(true)"
                              @keydown.enter.prevent="finishColumnRename(true)"
                              @keydown.esc.prevent="finishColumnRename(false)"
                            />
                          </template>
                          <template v-else>
                            <span
                              class="header-cell__title"
                              :class="{ 'header-cell__title--selected': isColumnSelected(column.name) }"
                              @dblclick.stop="startColumnRename(column.name)"
                            >
                              {{ column.name }}
                            </span>
                          </template>
                          <small>Type: {{ getTypeLabel(column.overrideType || column.detectedType) }}</small>
                        </div>
                        <button
                          class="header-filter"
                          type="button"
                          title="Filter column"
                          @click.stop="toggleFilterMenu(column.name)"
                        >
                          ▾
                        </button>
                        <div v-if="filterMenuOpen === column.name" class="filter-popover">
                          <div class="filter-search">
                            <input
                              class="filter-input"
                              type="text"
                              placeholder="Search values"
                              :value="filterSearchTerms[column.name] || ''"
                              @input="filterSearchTerms[column.name] = $event.target.value"
                            />
                            <button
                              class="filter-clear"
                              type="button"
                              @click.stop="filterSearchTerms[column.name] = ''"
                            >
                              ×
                            </button>
                          </div>
                          <div class="filter-values">
                            <div class="filter-values-inner">
                              <label class="filter-value filter-select-all">
                                <input
                                  type="checkbox"
                                  :checked="isSelectAllChecked(column.name)"
                                  @change="handleSelectAllToggle(column.name, $event.target.checked)"
                                />
                                <span>Select all</span>
                              </label>
                              <label
                                v-for="value in getColumnUniqueValues(column.name)"
                                :key="value"
                                class="filter-value"
                              >
                                <input
                                  type="checkbox"
                                  :checked="getFilterDraft(column.name).has(value)"
                                  @change="toggleFilterValue(column.name, value, $event.target.checked)"
                                />
                                <span>{{ value }}</span>
                              </label>
                            </div>
                          </div>
                          <div class="filter-actions">
                            <button class="filter-apply" type="button" @click.stop="applyFilter(column.name)">
                              Apply
                            </button>
                            <button class="filter-clear" type="button" @click.stop="clearFilter(column.name)">
                              Clear
                            </button>
                          </div>
                        </div>
                      </div>
                    </th>
                  </tr>
                </thead>
              <tbody>
                <tr
                  v-for="row in filteredRows"
                  :key="row.rowId"
                  :class="{
                    'row-duplicate': duplicateRowIds.has(row.rowId),
                    'row-focus': focusedDuplicateRows.has(row.rowId),
                    'row-selected': isRowSelected(row.rowId),
                  }"
                >
                  <td
                    class="index"
                    :class="{ 'index-selected': isRowSelected(row.rowId) }"
                    @click="selectRow(row.rowId, $event)"
                  >
                    <span>{{ row.rowId + 1 }}</span>
                  </td>
                    <td v-for="column in columns" :key="column.name">
                      <input
                        class="cell-input"
                        :class="cellClass(row.rowId, column.name)"
                        :value="row.values[column.name] ?? ''"
                        type="text"
                        @input="updateCell(row.rowId, column.name, $event.target.value)"
                      />
                    </td>
                  </tr>
                </tbody>
              </table>
            </div>
          </div>
        </div>
      </div>
    </section>
  </div>
</template>
