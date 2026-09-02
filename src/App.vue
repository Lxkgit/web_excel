<script setup>
import { computed, onBeforeUnmount, onMounted, ref } from 'vue'
import * as XLSX from 'xlsx'

const input = ref(null)
const files = ref([])
const activeFile = ref(null)
const activeSheet = ref(null)
const isParsing = ref(false)
const error = ref('')
const selection = ref(null)
const copied = ref(false)
const isDragging = ref(false)
const directoryHandle = ref(null)
const archiveDirectoryHandle = ref(null)
const fileMenu = ref(null)
const isRefreshing = ref(false)
const tableWrap = ref(null)
const sortMode = ref('modified')
const fileCreatedAt = ref(new Map())
let copiedTimer
let refreshTimer
let autoScrollFrame
let lastMouse = { x: 0, y: 0 }

const selectedBook = computed(() => files.value.find((file) => file.id === activeFile.value))
const selectedSheet = computed(() => selectedBook.value?.sheets.find((sheet) => sheet.name === activeSheet.value))
const totalSheets = computed(() => files.value.reduce((sum, file) => sum + file.sheets.length, 0))
const columnLabels = computed(() => Array.from({ length: selectedSheet.value?.columnCount ?? 0 }, (_, index) => XLSX.utils.encode_col(index)))

function isExcel(file) {
  return /\.(xlsx|xls|xlsm|xlsb)$/i.test(file.name) && !file.name.startsWith('~$')
}
function formatCell(cell) {
  if (!cell) return ''
  if (cell.w !== undefined) return cell.w
  if (cell.v === null || cell.v === undefined) return ''
  return String(cell.v)
}
function visualStyle(cell) {
  const fill = cell?.s?.fill?.fgColor?.rgb
  const color = cell?.s?.font?.color?.rgb
  const style = {}
  if (fill && !/^0{8}$/i.test(fill)) style.backgroundColor = `#${fill.slice(-6)}`
  if (color && !/^0{8}$/i.test(color)) style.color = `#${color.slice(-6)}`
  if (cell?.s?.font?.bold) style.fontWeight = 700
  if (cell?.s?.alignment?.horizontal) style.textAlign = cell.s.alignment.horizontal
  if (cell?.s?.alignment?.vertical) style.verticalAlign = cell.s.alignment.vertical
  if (cell?.s?.alignment?.wrapText) style.whiteSpace = 'pre-wrap'
  return style
}
function parseSheet(worksheet, name) {
  const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1')
  const merges = worksheet['!merges'] || []
  const covered = new Set()
  const mergeStarts = new Map()
  merges.forEach((merge) => {
    const start = `${merge.s.r}:${merge.s.c}`
    mergeStarts.set(start, { rowspan: merge.e.r - merge.s.r + 1, colspan: merge.e.c - merge.s.c + 1 })
    for (let row = merge.s.r; row <= merge.e.r; row += 1) {
      for (let col = merge.s.c; col <= merge.e.c; col += 1) {
        if (row !== merge.s.r || col !== merge.s.c) covered.add(`${row}:${col}`)
      }
    }
  })
  const rows = []
  for (let row = range.s.r; row <= range.e.r; row += 1) {
    const cells = []
    for (let col = range.s.c; col <= range.e.c; col += 1) {
      const key = `${row}:${col}`
      if (covered.has(key)) continue
      const cell = worksheet[XLSX.utils.encode_cell({ r: row, c: col })]
      const merge = mergeStarts.get(key)
      cells.push({ key, row, col, value: formatCell(cell), colspan: merge?.colspan ?? 1, rowspan: merge?.rowspan ?? 1, style: visualStyle(cell), type: cell?.t })
    }
    rows.push({ number: row + 1, cells })
  }
  return { name, rows, startRow: range.s.r, startCol: range.s.c, endRow: range.e.r, endCol: range.e.c, columnCount: range.e.c - range.s.c + 1, rowCount: range.e.r - range.s.r + 1, worksheet }
}
function naturalCompare(a, b) {
  return a.localeCompare(b, undefined, { numeric: true, sensitivity: 'base' })
}
function compareFiles(a, b) {
  if (sortMode.value === 'modified') {
    return b.lastModified - a.lastModified || naturalCompare(a.path, b.path)
  }
  if (sortMode.value === 'created') {
    return (fileCreatedAt.value.get(b.id) ?? 0) - (fileCreatedAt.value.get(a.id) ?? 0) || naturalCompare(a.path, b.path)
  }
  return naturalCompare(a.name, b.name) || naturalCompare(a.path, b.path)
}
function sortFiles(list) {
  return [...list].sort(compareFiles)
}
function updateCreatedTimes(entries) {
  const now = Date.now()
  const next = new Map(fileCreatedAt.value)
  entries.forEach((entry) => {
    const id = entry.path
    if (!next.has(id)) next.set(id, now)
  })
  fileCreatedAt.value = next
}
function sameFileMeta(a, b) {
  return a && b && a.id === b.id && a.size === b.size && a.lastModified === b.lastModified
}
async function buildFile(entry, previous = null) {
  const file = entry.file || entry
  const path = entry.path || file.webkitRelativePath || file.name
  if (previous && sameFileMeta({ ...previous, id: path }, { ...entry, id: path })) {
    return { ...previous, fileHandle: entry.fileHandle, parentHandle: entry.parentHandle, source: file }
  }
  const workbook = XLSX.read(await file.arrayBuffer(), { type: 'array', cellDates: true, cellStyles: true })
  return { id: path, name: file.name, path, size: file.size, lastModified: file.lastModified || 0, source: file, fileHandle: entry.fileHandle, parentHandle: entry.parentHandle, sheets: workbook.SheetNames.map((name) => parseSheet(workbook.Sheets[name], name)) }
}
async function parseFiles(entries, preserveActive = false) {
  const previousFile = activeFile.value
  const previousSheet = activeSheet.value
  updateCreatedTimes(entries)
  const previousMap = new Map(files.value.map((file) => [file.id, file]))
  const parsed = await Promise.all(entries.map((entry) => buildFile(entry, previousMap.get(entry.path || entry.file?.webkitRelativePath || entry.file?.name))))
  files.value = sortFiles(parsed)
  const retained = preserveActive && files.value.find((file) => file.id === previousFile)
  if (retained) {
    activeFile.value = retained.id
    activeSheet.value = retained.sheets.some((sheet) => sheet.name === previousSheet) ? previousSheet : retained.sheets[0]?.name
  } else if (files.value[0]) chooseFile(files.value[0])
  else { activeFile.value = null; activeSheet.value = null }
}
async function collectDirectoryFiles(handle, prefix = '') {
  const entries = []
  for await (const [name, child] of handle.entries()) {
    const path = prefix ? `${prefix}/${name}` : name
    if (child.kind === 'directory') {
      if (archiveDirectoryHandle.value && await child.isSameEntry(archiveDirectoryHandle.value)) continue
      entries.push(...await collectDirectoryFiles(child, path))
    } else if (isExcel({ name })) entries.push({ file: await child.getFile(), fileHandle: child, parentHandle: handle, path })
  }
  return entries
}
function stopDirectoryWatch() { clearInterval(refreshTimer); refreshTimer = undefined }
function startDirectoryWatch() { stopDirectoryWatch(); refreshTimer = setInterval(() => refreshDirectory(true), 5000) }
async function refreshDirectory(preserveActive = true) {
  if (!directoryHandle.value || isRefreshing.value) return
  isRefreshing.value = true
  try { await parseFiles(await collectDirectoryFiles(directoryHandle.value), preserveActive); error.value = '' }
  catch (reason) { console.error(reason); error.value = '无法刷新目录。请重新选择文件夹并授予读取和写入权限。'; stopDirectoryWatch() }
  finally { isRefreshing.value = false }
}
async function chooseFolder() {
  if (!window.showDirectoryPicker) { input.value?.click(); return }
  try {
    const handle = await window.showDirectoryPicker({ mode: 'readwrite' })
    directoryHandle.value = handle
    error.value = ''
    await refreshDirectory(false)
    startDirectoryWatch()
  } catch (reason) { if (reason?.name !== 'AbortError') console.error(reason) }
}
async function selectFolder(event) {
  const picked = [...event.target.files].filter(isExcel)
  event.target.value = ''
  directoryHandle.value = null
  stopDirectoryWatch()
  error.value = ''
  files.value = []
  activeFile.value = null
  activeSheet.value = null
  selection.value = null
  if (!picked.length) { error.value = '这个文件夹里没有可读取的 Excel 文件。请尝试选择另一个文件夹。'; return }
  isParsing.value = true
  try { await parseFiles(picked) }
  catch (reason) { console.error(reason); error.value = '读取文件时出现问题。请确认表格没有损坏或受密码保护。' }
  finally { isParsing.value = false }
}
async function chooseArchiveFolder() {
  if (!window.showDirectoryPicker) { error.value = '当前浏览器不支持归档目录配置，请使用 Chrome 或 Edge。'; return }
  try {
    const handle = await window.showDirectoryPicker({ mode: 'readwrite' })
    if (directoryHandle.value && await handle.isSameEntry(directoryHandle.value)) { error.value = '归档目录不能与当前 Excel 源目录相同。'; return }
    archiveDirectoryHandle.value = handle
    error.value = ''
    await saveArchiveHandle(handle)
  } catch (reason) {
    if (reason?.name !== 'AbortError') { console.error(reason); error.value = '配置归档目录失败，请确认已授予目录读写权限。' }
  }
}
function changeSortMode(event) {
  sortMode.value = event.target.value
  files.value = sortFiles(files.value)
}
function chooseFile(file) { activeFile.value = file.id; activeSheet.value = file.sheets[0]?.name ?? null; selection.value = null }
function bounds() {
  if (!selection.value) return null
  const { startRow, startCol, endRow, endCol } = selection.value
  return { top: Math.min(startRow, endRow), bottom: Math.max(startRow, endRow), left: Math.min(startCol, endCol), right: Math.max(startCol, endCol) }
}
function selectRange(startRow, startCol, endRow = startRow, endCol = startCol) { selection.value = { startRow, startCol, endRow, endCol } }
function startCellSelection(cell, event) {
  isDragging.value = true
  if (event.shiftKey && selection.value) selectRange(selection.value.startRow, selection.value.startCol, cell.row + cell.rowspan - 1, cell.col + cell.colspan - 1)
  else selectRange(cell.row, cell.col, cell.row + cell.rowspan - 1, cell.col + cell.colspan - 1)
  lastMouse = { x: event.clientX, y: event.clientY }
  updateAutoScroll()
}
function extendCellSelection(cell, event) {
  if (!isDragging.value || !(event.buttons & 1) || !selection.value) return
  selectRange(selection.value.startRow, selection.value.startCol, cell.row + cell.rowspan - 1, cell.col + cell.colspan - 1)
  lastMouse = { x: event.clientX, y: event.clientY }
}
function selectRow(row, event) { const sheet = selectedSheet.value; if (sheet) selectRange(event.shiftKey && selection.value ? selection.value.startRow : row, sheet.startCol, row, sheet.endCol) }
function selectColumn(col, event) { const sheet = selectedSheet.value; if (sheet) selectRange(sheet.startRow, event.shiftKey && selection.value ? selection.value.startCol : col, sheet.endRow, col) }
function cellSelected(cell) { const area = bounds(); return area && cell.row <= area.bottom && cell.row + cell.rowspan - 1 >= area.top && cell.col <= area.right && cell.col + cell.colspan - 1 >= area.left }
function rowSelected(row) { const area = bounds(); return area && row >= area.top && row <= area.bottom && area.left === selectedSheet.value?.startCol && area.right === selectedSheet.value?.endCol }
function columnSelected(col) { const area = bounds(); return area && col >= area.left && col <= area.right && area.top === selectedSheet.value?.startRow && area.bottom === selectedSheet.value?.endRow }
function selectionText() {
  const area = bounds(); const sheet = selectedSheet.value
  if (!area || !sheet) return ''
  const output = []
  for (let row = area.top; row <= area.bottom; row += 1) {
    const values = []
    for (let col = area.left; col <= area.right; col += 1) values.push(formatCell(sheet.worksheet[XLSX.utils.encode_cell({ r: row, c: col })]))
    output.push(values.join('\t'))
  }
  return output.join('\n')
}
async function copySelection() {
  const text = selectionText(); if (!text) return
  try { await navigator.clipboard.writeText(text) } catch { const textarea = document.createElement('textarea'); textarea.value = text; document.body.append(textarea); textarea.select(); document.execCommand('copy'); textarea.remove() }
  copied.value = true; clearTimeout(copiedTimer); copiedTimer = setTimeout(() => { copied.value = false }, 1600)
}
function selectAll() { const sheet = selectedSheet.value; if (sheet) selectRange(sheet.startRow, sheet.startCol, sheet.endRow, sheet.endCol) }
function clearSelection() { selection.value = null }
function handleKeydown(event) { if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'c' && selection.value) { event.preventDefault(); copySelection() } }
function updateAutoScroll() {
  cancelAnimationFrame(autoScrollFrame)
  if (!isDragging.value || !tableWrap.value) return
  const wrap = tableWrap.value
  const rect = wrap.getBoundingClientRect()
  const edge = 72
  const maxSpeed = 20
  let deltaY = 0
  let deltaX = 0
  if (lastMouse.y < rect.top + edge) deltaY = -Math.ceil((rect.top + edge - lastMouse.y) / edge * maxSpeed)
  else if (lastMouse.y > rect.bottom - edge) deltaY = Math.ceil((lastMouse.y - (rect.bottom - edge)) / edge * maxSpeed)
  if (lastMouse.x < rect.left + edge) deltaX = -Math.ceil((rect.left + edge - lastMouse.x) / edge * maxSpeed)
  else if (lastMouse.x > rect.right - edge) deltaX = Math.ceil((lastMouse.x - (rect.right - edge)) / edge * maxSpeed)
  if (!deltaX && !deltaY) return
  wrap.scrollTop += deltaY; wrap.scrollLeft += deltaX
  extendSelectionAtPointer()
  autoScrollFrame = requestAnimationFrame(updateAutoScroll)
}
function extendSelectionAtPointer() {
  if (!isDragging.value || !selection.value) return
  const element = document.elementFromPoint(lastMouse.x, lastMouse.y)
  const cell = element?.closest?.('td[data-row][data-col]')
  if (!cell) return
  const row = Number(cell.dataset.row); const col = Number(cell.dataset.col)
  if (Number.isInteger(row) && Number.isInteger(col)) selectRange(selection.value.startRow, selection.value.startCol, row, col)
}
function handleTableMouseMove(event) { if (isDragging.value) { lastMouse = { x: event.clientX, y: event.clientY }; updateAutoScroll() } }
function stopAutoScroll() { cancelAnimationFrame(autoScrollFrame); autoScrollFrame = undefined }
function stopDragging() { isDragging.value = false; stopAutoScroll() }
onMounted(async () => {
  window.addEventListener('keydown', handleKeydown); window.addEventListener('mouseup', stopDragging); window.addEventListener('click', closeFileMenu)
  archiveDirectoryHandle.value = await loadArchiveHandle()
})
onBeforeUnmount(() => {
  window.removeEventListener('keydown', handleKeydown); window.removeEventListener('mouseup', stopDragging); window.removeEventListener('click', closeFileMenu)
  stopDirectoryWatch(); stopAutoScroll(); clearTimeout(copiedTimer)
})
function openPicker() { chooseFolder() }
function openFile(file) {
  if (!file?.source) return
  window.open(URL.createObjectURL(file.source), '_blank', 'noopener')
}
function showFileMenu(file, event) { fileMenu.value = { file, x: event.clientX, y: event.clientY } }
function closeFileMenu() { fileMenu.value = null }
async function renameFile(file) {
  closeFileMenu()
  if (!file?.fileHandle || !file.parentHandle) { error.value = '当前浏览器的兼容模式不支持重命名，请使用 Chrome 或 Edge 重新选择文件夹。'; return }
  const name = window.prompt('输入新的文件名', file.name)
  if (!name || name === file.name) return
  if (!isExcel({ name })) { error.value = '新文件名必须使用 Excel 文件扩展名。'; return }
  try {
    const target = await file.parentHandle.getFileHandle(name, { create: true })
    const writable = await target.createWritable(); await writable.write(await file.fileHandle.getFile()); await writable.close(); await file.fileHandle.remove(); await refreshDirectory(true)
  } catch (reason) { console.error(reason); error.value = '重命名失败。请确认文件未被其它程序占用，并允许网站写入该文件夹。' }
}
async function archiveFile(file) {
  closeFileMenu()
  if (!file?.fileHandle || !file.parentHandle) { error.value = '当前浏览器的兼容模式不支持归档，请使用 Chrome 或 Edge 重新选择文件夹。'; return }
  if (!archiveDirectoryHandle.value) { error.value = '请先配置归档目录。'; return }
  try {
    let permission = await archiveDirectoryHandle.value.queryPermission({ mode: 'readwrite' })
    if (permission !== 'granted') permission = await archiveDirectoryHandle.value.requestPermission({ mode: 'readwrite' })
    if (permission !== 'granted') { error.value = '没有归档目录的写入权限。'; return }
    const dateFolder = new Date().toLocaleDateString('sv-SE')
    const dateDirectory = await archiveDirectoryHandle.value.getDirectoryHandle(dateFolder, { create: true })
    let target
    try {
      target = await dateDirectory.getFileHandle(file.name, { create: false })
      if (!window.confirm(`归档目录中已存在“${file.name}”，是否覆盖？`)) return
    } catch { target = await dateDirectory.getFileHandle(file.name, { create: true }) }
    const writable = await target.createWritable(); await writable.write(await file.fileHandle.getFile()); await writable.close()
    await file.fileHandle.remove()
    await refreshDirectory(true)
  } catch (reason) { console.error(reason); error.value = '归档失败。请确认源文件未被其它程序占用，并允许归档目录写入。' }
}
async function deleteFile(file) {
  closeFileMenu()
  if (!file?.fileHandle) { error.value = '当前浏览器的兼容模式不支持删除，请使用 Chrome 或 Edge 重新选择文件夹。'; return }
  if (!window.confirm(`确定删除“${file.name}”吗？此操作不可撤销。`)) return
  try { await file.fileHandle.remove(); await refreshDirectory(true) } catch (reason) { console.error(reason); error.value = '删除失败。请确认文件未被其它程序占用，并允许网站写入该文件夹。' }
}
function formatBytes(bytes) { return bytes < 1024 * 1024 ? `${Math.max(1, Math.round(bytes / 1024))} KB` : `${(bytes / 1024 / 1024).toFixed(1)} MB` }
const HANDLE_DB_NAME = 'web-excel-settings'
const HANDLE_STORE_NAME = 'settings'
function openHandleDb() {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(HANDLE_DB_NAME, 1)
    request.onupgradeneeded = () => request.result.createObjectStore(HANDLE_STORE_NAME)
    request.onsuccess = () => resolve(request.result)
    request.onerror = () => reject(request.error)
  })
}
async function saveArchiveHandle(handle) {
  if (!('indexedDB' in window)) return
  try {
    const db = await openHandleDb()
    await new Promise((resolve, reject) => { const transaction = db.transaction(HANDLE_STORE_NAME, 'readwrite'); transaction.objectStore(HANDLE_STORE_NAME).put(handle, 'archiveDirectory'); transaction.oncomplete = resolve; transaction.onerror = () => reject(transaction.error) })
    db.close()
  } catch (reason) { console.warn('无法持久化归档目录配置', reason) }
}
async function loadArchiveHandle() {
  if (!('indexedDB' in window)) return null
  try {
    const db = await openHandleDb()
    const handle = await new Promise((resolve, reject) => { const transaction = db.transaction(HANDLE_STORE_NAME, 'readonly'); const request = transaction.objectStore(HANDLE_STORE_NAME).get('archiveDirectory'); request.onsuccess = () => resolve(request.result || null); request.onerror = () => reject(request.error) })
    db.close(); return handle
  } catch (reason) { console.warn('无法读取归档目录配置', reason); return null }
}
</script>

<template>
  <main class="app-shell" :class="{ 'has-content': files.length }">
    <header class="topbar">
      <a class="brand" href="#" aria-label="表格阅览室首页"><span class="brand-mark">▦</span> 表格阅览室</a>
      <span class="topbar-label">EXCEL FOLDER READER · LOCAL ONLY</span>
      <div v-if="files.length" class="header-stats"><b>{{ files.length }}</b> 个文件 <i></i> <b>{{ totalSheets }}</b> 个工作表</div>
      <button class="outline-button" type="button" @click="chooseArchiveFolder">归档目录 <span>▣</span></button>
      <button class="outline-button" type="button" @click="openPicker">选择文件夹 <span>↗</span></button>
    </header>
    <section v-if="!files.length && !isParsing" class="welcome">
      <div class="hero"><div><p class="eyebrow">LOCAL · PRIVATE · IN BROWSER</p><h1>让文件夹里的表格<br /><em>一目了然。</em></h1><p class="intro">选择一个文件夹，快速读取其中所有 Excel 文件并在浏览器中预览。文件不会上传，数据始终留在你的设备上。</p></div><div class="hero-stat"><strong>01</strong><span>选择文件夹</span><i></i><b>02</b><span>开始预览</span></div></div>
      <section class="drop-card"><div class="grid-orb">▦</div><p class="eyebrow">START HERE</p><h2>选择一个表格文件夹</h2><p>支持读取文件夹及子文件夹中的 <b>.xlsx、.xls、.xlsm、.xlsb</b> 文件。</p><button class="primary-button" type="button" @click="openPicker">浏览文件夹 <span>→</span></button><small>Chrome、Edge 等现代浏览器支持文件夹选择</small></section>
    </section>
    <section v-else-if="isParsing" class="loading-card"><div class="loader"></div><strong>正在整理表格内容…</strong><span>请稍候，正在安全地在本地读取文件。</span></section>
    <section v-else class="workspace">
      <aside class="file-panel">
        <div class="panel-heading"><span>文件列表</span><b>{{ files.length }}</b><select class="sort-select" :value="sortMode" aria-label="文件排序方式" @change="changeSortMode"><option value="modified">最近修改时间</option><option value="created">创建时间</option><option value="name">文字首字母</option></select></div>
        <button v-for="file in files" :key="file.id" class="file-item" :class="{ active: activeFile === file.id }" type="button" title="双击在浏览器中打开；右键显示文件操作" @click="chooseFile(file)" @dblclick="openFile(file)" @contextmenu.prevent.stop="showFileMenu(file, $event)"><span class="file-icon">X</span><span class="file-name">{{ file.name }}<small>{{ file.sheets.length }} 个工作表 · {{ formatBytes(file.size) }}</small></span></button>
        <button class="add-folder" type="button" @click="openPicker">＋ 更换文件夹</button>
        <button class="add-folder" type="button" @click="chooseArchiveFolder">＋ {{ archiveDirectoryHandle ? '更换归档目录' : '设置归档目录' }}</button>
        <small v-if="archiveDirectoryHandle" class="archive-path">归档：{{ archiveDirectoryHandle.name }}</small>
      </aside>
      <section class="preview-panel">
        <div class="file-meta"><div><p class="eyebrow">WORKBOOK PREVIEW</p><h2>{{ selectedBook?.name }}</h2><span>{{ selectedBook?.path }}</span></div><span class="local-badge">● 本地解析，未上传</span></div>
        <div class="sheet-tabs" role="tablist"><button v-for="sheet in selectedBook?.sheets" :key="sheet.name" type="button" :class="{ active: activeSheet === sheet.name }" @click="activeSheet = sheet.name; clearSelection()">{{ sheet.name }} <small>{{ sheet.rowCount }} 行</small></button></div>
        <div ref="tableWrap" class="table-wrap" @mousemove="handleTableMouseMove" @mouseleave="stopAutoScroll">
          <table v-if="selectedSheet?.rows.length" @dragstart.prevent>
            <thead><tr><th class="corner-cell" title="点击全选" @mousedown.prevent="selectAll">#</th><th v-for="(label, index) in columnLabels" :key="label" :class="{ selected: columnSelected(selectedSheet.startCol + index) }" @mousedown.prevent="selectColumn(selectedSheet.startCol + index, $event)">{{ label }}</th></tr></thead>
            <tbody><tr v-for="row in selectedSheet.rows" :key="row.number"><th class="row-num" :class="{ selected: rowSelected(row.number - 1) }" @mousedown.prevent="selectRow(row.number - 1, $event)">{{ row.number }}</th><td v-for="cell in row.cells" :key="cell.key" :data-row="cell.row" :data-col="cell.col" :colspan="cell.colspan" :rowspan="cell.rowspan" :style="cell.style" :class="{ 'number-cell': cell.type === 'n', selected: cellSelected(cell) }" @mousedown.prevent="startCellSelection(cell, $event)" @mouseenter="extendCellSelection(cell, $event)">{{ cell.value }}</td></tr></tbody>
          </table>
          <div v-else class="empty-sheet">这个工作表没有可展示的数据。</div>
        </div>
        <div v-if="selection" class="selection-menu"><span>{{ copied ? '已复制到剪贴板' : '已选区域' }}</span><button type="button" @click="copySelection">复制 <kbd>Ctrl C</kbd></button><button type="button" @click="selectAll">全选</button><button type="button" @click="clearSelection">取消</button></div>
      </section>
    </section>
    <div v-if="fileMenu" class="file-menu" :style="{ left: `${fileMenu.x}px`, top: `${fileMenu.y}px` }" @click.stop>
      <strong>{{ fileMenu.file.name }}</strong>
      <button type="button" @click="openFile(fileMenu.file); closeFileMenu()">在浏览器打开</button>
      <button type="button" @click="archiveFile(fileMenu.file)">归档到 {{ archiveDirectoryHandle?.name || '归档目录' }}</button>
      <button type="button" @click="renameFile(fileMenu.file)">重命名</button>
      <button type="button" class="danger" @click="deleteFile(fileMenu.file)">删除</button>
    </div>
    <p v-if="directoryHandle" class="watch-status">● 正在每 5 秒检查当前目录的新文件</p>
    <p v-if="error" class="error-message">{{ error }}</p>
    <input ref="input" class="visually-hidden" type="file" webkitdirectory directory multiple accept=".xlsx,.xls,.xlsm,.xlsb" @change="selectFolder" />
  </main>
</template>
