const API_BASE = import.meta.env.VITE_API_BASE ?? '/api'

async function handleResponse(responsePromise) {
  const response = await responsePromise
  if (!response.ok) {
    const detail = await safeParseError(response)
    throw new Error(detail || 'Request failed')
  }
  return response.json()
}

async function safeParseError(response) {
  try {
    const payload = await response.json()
    return payload?.detail || payload?.message
  } catch (error) {
    return null
  }
}

export async function uploadWorkbook(file) {
  const formData = new FormData()
  formData.append('file', file)
  return handleResponse(
    fetch(`${API_BASE}/upload`, {
      method: 'POST',
      body: formData,
    }),
  )
}

export async function validateWorkbook(sessionId, rows, columnTypes, columns) {
  return handleResponse(
    fetch(`${API_BASE}/validate`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        sessionId,
        rows,
        columnTypes,
        columns,
      }),
    }),
  )
}

export async function removeDuplicateRows(sessionId, rowIds) {
  return handleResponse(
    fetch(`${API_BASE}/duplicates/remove`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        sessionId,
        rowIds,
      }),
    }),
  )
}

export async function switchSheet(sessionId, sheetName) {
  return handleResponse(
    fetch(`${API_BASE}/sheet`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        sessionId,
        sheetName,
      }),
    }),
  )
}

