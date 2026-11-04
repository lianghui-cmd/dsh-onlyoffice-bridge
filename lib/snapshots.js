// @ts-check
// 文档版本快照：不可变副本 + 索引 + 回滚。
//
// 存储布局（都在缓存目录下，随 TTL 一起被清理）：
//   <key>.<ext>                       当前版本
//   <key>.meta.json                   文档元信息
//   snapshots/<key>/<snapId>.<ext>    快照副本（只读，不再修改）
//   snapshots/index.json              快照索引（fileKey → 快照列表）
//
// snapId 由时间戳 + 随机串组成，字典序即时间序，便于排序与定位。

import crypto from 'node:crypto'
import fs from 'node:fs'
import path from 'node:path'

/** 快照子目录。 */
const SNAPSHOT_DIR = 'snapshots'
/** 索引文件名。 */
const INDEX_FILE = 'index.json'

/**
 * @param {string} storageDir
 * @param {string} fileKey
 */
function assertKey(storageDir, fileKey) {
  const name = String(fileKey ?? '').trim()
  if (!name || name.includes('/') || name.includes('\\') || name.includes('..')) {
    throw new Error(`fileKey 非法：${fileKey}`)
  }
  const target = path.join(storageDir, name)
  const relative = path.relative(storageDir, target)
  if (!relative || relative.startsWith('..') || path.isAbsolute(relative)) {
    throw new Error(`fileKey 非法：${fileKey}`)
  }
  return target
}

/** @param {string} storageDir */
function indexPath(storageDir) {
  return path.join(storageDir, SNAPSHOT_DIR, INDEX_FILE)
}

/**
 * 读取快照索引。
 * @param {string} storageDir
 * @returns {Record<string, Array<{ snapId: string, label: string, createTime: string, ext: string, size: number }>>}
 */
export function readSnapshotIndex(storageDir) {
  try {
    const parsed = JSON.parse(fs.readFileSync(indexPath(storageDir), 'utf8'))
    return parsed && typeof parsed === 'object' ? parsed : {}
  } catch {
    return {}
  }
}

/**
 * 写入快照索引。
 * @param {string} storageDir
 * @param {Record<string, any>} index
 */
function writeSnapshotIndex(storageDir, index) {
  const file = indexPath(storageDir)
  fs.mkdirSync(path.dirname(file), { recursive: true })
  fs.writeFileSync(file, JSON.stringify(index, null, 2), 'utf8')
}

/**
 * 为一个文档打快照。
 * @param {{ storageDir: string, fileKey: string, label?: string }} input
 */
export function createSnapshot(input) {
  const target = assertKey(input.storageDir, input.fileKey)
  if (!fs.existsSync(target)) throw new Error(`fileKey 不存在：${input.fileKey}`)

  const ext = path.extname(target).slice(1).toLowerCase() || 'bin'
  // snapId 用「毫秒时间戳 + 同毫秒内的自增序号 + 随机串」：
  // 只用秒级时间戳时，同一秒内连续创建的两个快照只能靠随机串排序，
  // 结果 snapshot-list 的顺序会与创建顺序不一致。
  const now = new Date()
  const stamp =
    now.toISOString().replace(/[-:.TZ]/g, '').slice(0, 17) // 精确到毫秒
  const stem = input.fileKey.replace(/\.[^.]+$/, '')
  const dir = path.join(input.storageDir, SNAPSHOT_DIR, stem)
  fs.mkdirSync(dir, { recursive: true })

  // 同一毫秒内多次创建时，用已有快照数做序号，保证单调递增
  let seq = 0
  while (seq < 999) {
    const candidate = `${stamp}${String(seq).padStart(3, '0')}-${crypto.randomBytes(3).toString('hex')}`
    if (!fs.existsSync(path.join(dir, `${candidate}.${ext}`))) break
    seq += 1
  }
  const snapId = `${stamp}${String(seq).padStart(3, '0')}-${crypto.randomBytes(3).toString('hex')}`
  const snapPath = path.join(dir, `${snapId}.${ext}`)
  fs.copyFileSync(target, snapPath)

  const stat = fs.statSync(snapPath)
  const entry = {
    snapId,
    label: String(input.label ?? '').trim(),
    createTime: new Date().toISOString(),
    ext,
    size: stat.size,
  }
  const index = readSnapshotIndex(input.storageDir)
  const list = index[input.fileKey] ?? []
  list.push(entry)
  index[input.fileKey] = list
  writeSnapshotIndex(input.storageDir, index)

  return { snapId, fileKey: input.fileKey, label: entry.label, createTime: entry.createTime, size: entry.size }
}

/**
 * 列出某个文档的全部快照（按时间正序）。
 * @param {{ storageDir: string, fileKey: string }} input
 */
export function listSnapshots(input) {
  assertKey(input.storageDir, input.fileKey)
  const list = readSnapshotIndex(input.storageDir)[input.fileKey] ?? []
  return list
    .slice()
    .sort((a, b) => {
      // 主排序：创建时间；时间相同（同一毫秒）时按 snapId 兜底，
      // snapId 内含毫秒时间戳与自增序号，能保证与创建顺序一致。
      const byTime = String(a.createTime ?? '').localeCompare(String(b.createTime ?? ''))
      return byTime !== 0 ? byTime : String(a.snapId).localeCompare(String(b.snapId))
    })
    .map((entry) => ({
      snapId: entry.snapId,
      label: entry.label,
      createTime: entry.createTime,
      size: entry.size,
    }))
}

/**
 * 回滚到指定快照。
 *
 * 回滚前会先给「当前版本」打一个备份快照，保证不会把现有内容弄丢。
 * @param {{ storageDir: string, fileKey: string, snapId: string }} input
 */
export function restoreSnapshot(input) {
  const target = assertKey(input.storageDir, input.fileKey)
  if (!fs.existsSync(target)) throw new Error(`fileKey 不存在：${input.fileKey}`)

  const index = readSnapshotIndex(input.storageDir)
  const list = index[input.fileKey] ?? []
  const entry = list.find((item) => item.snapId === input.snapId)
  if (!entry) throw new Error(`快照不存在：${input.snapId}`)

  const stem = input.fileKey.replace(/\.[^.]+$/, '')
  const snapPath = path.join(input.storageDir, SNAPSHOT_DIR, stem, `${entry.snapId}.${entry.ext}`)
  if (!fs.existsSync(snapPath)) throw new Error(`快照文件缺失：${entry.snapId}`)

  // 先备份当前版本
  const backup = createSnapshot({
    storageDir: input.storageDir,
    fileKey: input.fileKey,
    label: `回滚前自动备份（回滚到 ${entry.snapId}）`,
  })

  fs.copyFileSync(snapPath, target)

  return {
    fileKey: input.fileKey,
    backupSnapId: backup.snapId,
    restoredSnapId: entry.snapId,
    restoredLabel: entry.label,
    size: fs.statSync(target).size,
  }
}

/**
 * 清理快照：删除某个文档的全部快照（供 purge 调用）。
 * @param {string} storageDir
 * @param {string} fileKey
 */
export function purgeSnapshots(storageDir, fileKey) {
  const stem = String(fileKey).replace(/\.[^.]+$/, '')
  const dir = path.join(storageDir, SNAPSHOT_DIR, stem)
  fs.rmSync(dir, { recursive: true, force: true })
  const index = readSnapshotIndex(storageDir)
  if (index[fileKey]) {
    delete index[fileKey]
    writeSnapshotIndex(storageDir, index)
  }
}

/**
 * 清理过期快照：按快照文件 mtime 与给定 TTL。
 * @param {string} storageDir
 * @param {number} ttlMs
 */
export function cleanExpiredSnapshots(storageDir, ttlMs) {
  if (!(ttlMs > 0)) return []
  const root = path.join(storageDir, SNAPSHOT_DIR)
  if (!fs.existsSync(root)) return []
  /** @type {string[]} */
  const removed = []
  const index = readSnapshotIndex(storageDir)
  const cutoff = Date.now() - ttlMs
  for (const stem of fs.readdirSync(root)) {
    if (stem === INDEX_FILE) continue
    const dir = path.join(root, stem)
    if (!fs.statSync(dir).isDirectory()) continue
    for (const file of fs.readdirSync(dir)) {
      const filePath = path.join(dir, file)
      if (fs.statSync(filePath).mtimeMs < cutoff) {
        fs.rmSync(filePath, { force: true })
        removed.push(`${stem}/${file}`)
      }
    }
    if (!fs.readdirSync(dir).length) fs.rmdirSync(dir)
  }
  // 索引里去掉已删除的快照
  let changed = false
  for (const [fileKey, list] of Object.entries(index)) {
    const stem = fileKey.replace(/\.[^.]+$/, '')
    const keep = list.filter((entry) => {
      const exists = fs.existsSync(path.join(root, stem, `${entry.snapId}.${entry.ext}`))
      if (!exists) changed = true
      return exists
    })
    if (keep.length) index[fileKey] = keep
    else {
      delete index[fileKey]
      changed = true
    }
  }
  if (changed) writeSnapshotIndex(storageDir, index)
  return removed
}
