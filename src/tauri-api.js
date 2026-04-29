import { open } from '@tauri-apps/plugin-dialog'
import { invoke } from '@tauri-apps/api/core'

export async function selectFile({ filters }) {
  const result = await open({
    multiple: false,
    filters: filters.map(f => ({
      name: f.name,
      extensions: f.extensions
    }))
  })
  return result || null
}

export async function selectDirectory() {
  const result = await open({ directory: true, multiple: false })
  return result || null
}

export async function readFile(filePath) {
  const result = await invoke('read_file', { path: filePath })
  return new Uint8Array(result)
}

export async function writeFile(filePath, data) {
  try {
    await invoke('write_file', { path: filePath, data: Array.from(data) })
    return { ok: true }
  } catch (e) {
    return { ok: false, error: e.toString() }
  }
}

export function getFilename(filePath) {
  const parts = filePath.replace(/\\/g, '/').split('/')
  return parts[parts.length - 1]
}
