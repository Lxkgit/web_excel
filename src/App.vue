<script setup>
import { computed, ref } from 'vue'
import * as XLSX from 'xlsx'

const input = ref(null)
const files = ref([])
const activeFile = ref(null)
const activeSheet = ref(null)
const isParsing = ref(false)
const error = ref('')

const selectedBook = computed(() => files.value.find((file) => file.id === activeFile.value))
const selectedSheet = computed(() => selectedBook.value?.sheets.find((sheet) => sheet.name === activeSheet.value))
const totalSheets = computed(() => files.value.reduce((sum, file) => sum + file.sheets.length, 0))
const columnLabels = computed(() => Array.from({ length: selectedSheet.value?.columnCount ?? 0 }, (_, index) => XLSX.utils.encode_col(index)))

function isExcel(file) {
  return /\.(xlsx|xls|xlsm|xlsb)$/i.test(file.name)
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
      cells.push({
        key,
        value: formatCell(cell),
        colspan: merge?.colspan ?? 1,
        rowspan: merge?.rowspan ?? 1,
        style: visualStyle(cell),
        type: cell?.t,
      })
    }
    rows.push({ number: row + 1, cells })
  }

  return { name, rows, columnCount: range.e.c - range.s.c + 1, rowCount: range.e.r - range.s.r + 1 }
}

async function selectFolder(event) {
  const picked = [...event.target.files].filter(isExcel)
  error.value = ''
  files.value = []
  activeFile.value = null
  activeSheet.value = null
  event.target.value = ''

  if (!picked.length) {
    error.value = '这个文件夹里没有可读取的 Excel 文件。请尝试选择另一个文件夹。'
    return
  }

  isParsing.value = true
  try {
    const parsed = await Promise.all(picked.map(async (file, index) => {
      const workbook = XLSX.read(await file.arrayBuffer(), { type: 'array', cellDates: true, cellStyles: true })
      return {
        id: `${index}-${file.name}`,
        name: file.name,
        path: file.webkitRelativePath || file.name,
        size: file.size,
        sheets: workbook.SheetNames.map((name) => parseSheet(workbook.Sheets[name], name)),
      }
    }))
    files.value = parsed
    chooseFile(parsed[0])
  } catch (reason) {
    console.error(reason)
    error.value = '读取文件时出现问题。请确认表格没有损坏或受密码保护。'
  } finally {
    isParsing.value = false
  }
}

function chooseFile(file) {
  activeFile.value = file.id
  activeSheet.value = file.sheets[0]?.name ?? null
}
function openPicker() { input.value?.click() }
function formatBytes(bytes) { return bytes < 1024 * 1024 ? `${Math.max(1, Math.round(bytes / 1024))} KB` : `${(bytes / 1024 / 1024).toFixed(1)} MB` }
</script>

<template>
  <main class="app-shell" :class="{ 'has-content': files.length }">
    <header class="topbar">
      <a class="brand" href="#" aria-label="表格阅览室首页"><span class="brand-mark">▦</span> 表格阅览室</a>
      <span class="topbar-label">EXCEL FOLDER READER · LOCAL ONLY</span>
      <div v-if="files.length" class="header-stats"><b>{{ files.length }}</b> 个文件 <i></i> <b>{{ totalSheets }}</b> 个工作表</div>
      <button class="outline-button" type="button" @click="openPicker">选择文件夹 <span>↗</span></button>
    </header>

    <section v-if="!files.length && !isParsing" class="welcome">
      <div class="hero"><div><p class="eyebrow">LOCAL · PRIVATE · IN BROWSER</p><h1>让文件夹里的表格<br /><em>一目了然。</em></h1><p class="intro">选择一个文件夹，快速读取其中所有 Excel 文件并在浏览器中预览。文件不会上传，数据始终留在你的设备上。</p></div><div class="hero-stat"><strong>01</strong><span>选择文件夹</span><i></i><b>02</b><span>开始预览</span></div></div>
      <section class="drop-card"><div class="grid-orb">▦</div><p class="eyebrow">START HERE</p><h2>选择一个表格文件夹</h2><p>支持读取文件夹及子文件夹中的 <b>.xlsx、.xls、.xlsm、.xlsb</b> 文件。</p><button class="primary-button" type="button" @click="openPicker">浏览文件夹 <span>→</span></button><small>Chrome、Edge 等现代浏览器支持文件夹选择</small></section>
    </section>

    <section v-else-if="isParsing" class="loading-card"><div class="loader"></div><strong>正在整理表格内容…</strong><span>请稍候，正在安全地在本地读取文件。</span></section>

    <section v-else class="workspace">
      <aside class="file-panel"><div class="panel-heading"><span>文件列表</span><b>{{ files.length }}</b></div><button v-for="file in files" :key="file.id" class="file-item" :class="{ active: activeFile === file.id }" type="button" @click="chooseFile(file)"><span class="file-icon">X</span><span class="file-name">{{ file.name }}<small>{{ file.sheets.length }} 个工作表 · {{ formatBytes(file.size) }}</small></span></button><button class="add-folder" type="button" @click="openPicker">＋ 更换文件夹</button></aside>
      <section class="preview-panel"><div class="file-meta"><div><p class="eyebrow">WORKBOOK PREVIEW</p><h2>{{ selectedBook?.name }}</h2><span>{{ selectedBook?.path }}</span></div><span class="local-badge">● 本地解析，未上传</span></div><div class="sheet-tabs" role="tablist"><button v-for="sheet in selectedBook?.sheets" :key="sheet.name" type="button" :class="{ active: activeSheet === sheet.name }" @click="activeSheet = sheet.name">{{ sheet.name }} <small>{{ sheet.rowCount }} 行</small></button></div>
        <div class="table-wrap"><table v-if="selectedSheet?.rows.length"><thead><tr><th class="corner-cell">#</th><th v-for="label in columnLabels" :key="label">{{ label }}</th></tr></thead><tbody><tr v-for="row in selectedSheet.rows" :key="row.number"><th class="row-num">{{ row.number }}</th><td v-for="cell in row.cells" :key="cell.key" :colspan="cell.colspan" :rowspan="cell.rowspan" :style="cell.style" :class="{ 'number-cell': cell.type === 'n' }">{{ cell.value }}</td></tr></tbody></table><div v-else class="empty-sheet">这个工作表没有可展示的数据。</div></div>
      </section>
    </section>
    <p v-if="error" class="error-message">{{ error }}</p><input ref="input" class="visually-hidden" type="file" webkitdirectory directory multiple accept=".xlsx,.xls,.xlsm,.xlsb" @change="selectFolder" />
  </main>
</template>
