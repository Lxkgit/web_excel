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
const displayRows = computed(() => selectedSheet.value?.rows ?? [])
const displayHeaders = computed(() => {
  const columns = Math.max(1, ...displayRows.value.map((row) => row.length))
  return Array.from({ length: columns }, (_, index) => `列 ${index + 1}`)
})

function isExcel(file) {
  return /\.(xlsx|xls|xlsm|xlsb)$/i.test(file.name)
}

function cellValue(value) {
  if (value === null || value === undefined) return ''
  return String(value)
}

async function selectFolder(event) {
  const picked = [...event.target.files].filter(isExcel)
  error.value = ''
  files.value = []
  activeFile.value = null
  activeSheet.value = null

  if (!picked.length) {
    error.value = '这个文件夹里还没有可读取的 Excel 文件。请尝试选择另一个文件夹。'
    return
  }

  isParsing.value = true
  try {
    const parsed = await Promise.all(picked.map(async (file, index) => {
      const data = await file.arrayBuffer()
      const workbook = XLSX.read(data, { type: 'array', cellDates: true })
      return {
        id: `${index}-${file.name}`,
        name: file.name,
        path: file.webkitRelativePath || file.name,
        size: file.size,
        sheets: workbook.SheetNames.map((name) => ({
          name,
          rows: XLSX.utils.sheet_to_json(workbook.Sheets[name], { header: 1, defval: '' }),
        })),
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

function openPicker() {
  input.value?.click()
}

function formatBytes(bytes) {
  if (bytes < 1024 * 1024) return `${Math.max(1, Math.round(bytes / 1024))} KB`
  return `${(bytes / 1024 / 1024).toFixed(1)} MB`
}
</script>

<template>
  <main class="app-shell">
    <header class="topbar">
      <a class="brand" href="#" aria-label="表格阅览室首页"><span class="brand-mark">▦</span> 表格阅览室</a>
      <span class="topbar-label">EXCEL FOLDER READER</span>
      <button class="outline-button" type="button" @click="openPicker">选择文件夹 <span>↗</span></button>
    </header>

    <section class="hero">
      <div>
        <p class="eyebrow">LOCAL · PRIVATE · IN BROWSER</p>
        <h1>让文件夹里的表格<br /><em>一目了然。</em></h1>
        <p class="intro">选择一个文件夹，快速读取其中所有 Excel 文件并在浏览器中预览。文件不会上传，数据始终留在你的设备上。</p>
      </div>
      <div class="hero-stat" aria-label="已解析文件统计">
        <strong>{{ files.length || '00' }}</strong>
        <span>已读取文件</span>
        <i></i>
        <b>{{ totalSheets || '00' }}</b><span>工作表</span>
      </div>
    </section>

    <section v-if="!files.length && !isParsing" class="drop-card">
      <div class="grid-orb">▦</div>
      <p class="eyebrow">START HERE</p>
      <h2>选择一个表格文件夹</h2>
      <p>支持读取文件夹及子文件夹中的 <b>.xlsx、.xls、.xlsm、.xlsb</b> 文件。</p>
      <button class="primary-button" type="button" @click="openPicker">浏览文件夹 <span>→</span></button>
      <small>Chrome、Edge 等现代浏览器支持文件夹选择</small>
    </section>

    <section v-else-if="isParsing" class="loading-card">
      <div class="loader"></div><strong>正在整理表格内容…</strong><span>请稍候，正在安全地在本地读取文件。</span>
    </section>

    <section v-else class="workspace">
      <aside class="file-panel">
        <div class="panel-heading"><span>文件列表</span><b>{{ files.length }}</b></div>
        <button v-for="file in files" :key="file.id" class="file-item" :class="{ active: activeFile === file.id }" type="button" @click="chooseFile(file)">
          <span class="file-icon">X</span><span class="file-name">{{ file.name }}<small>{{ file.sheets.length }} 个工作表 · {{ formatBytes(file.size) }}</small></span>
        </button>
        <button class="add-folder" type="button" @click="openPicker">＋ 更换文件夹</button>
      </aside>

      <section class="preview-panel">
        <div class="file-meta"><div><p class="eyebrow">PREVIEW</p><h2>{{ selectedBook?.name }}</h2><span>{{ selectedBook?.path }}</span></div><span class="local-badge">● 仅本地读取</span></div>
        <div class="sheet-tabs" role="tablist"><button v-for="sheet in selectedBook?.sheets" :key="sheet.name" type="button" :class="{ active: activeSheet === sheet.name }" @click="activeSheet = sheet.name">{{ sheet.name }}</button></div>
        <div class="table-wrap">
          <table v-if="displayRows.length"><thead><tr><th class="corner-cell">#</th><th v-for="header in displayHeaders" :key="header">{{ header }}</th></tr></thead><tbody><tr v-for="(row, rowIndex) in displayRows" :key="rowIndex"><th class="row-num">{{ rowIndex + 1 }}</th><td v-for="(_, colIndex) in displayHeaders" :key="colIndex">{{ cellValue(row[colIndex]) }}</td></tr></tbody></table>
          <div v-else class="empty-sheet">这个工作表没有可展示的数据。</div>
        </div>
      </section>
    </section>

    <p v-if="error" class="error-message">{{ error }}</p>
    <input ref="input" class="visually-hidden" type="file" webkitdirectory directory multiple accept=".xlsx,.xls,.xlsm,.xlsb" @change="selectFolder" />
  </main>
</template>
