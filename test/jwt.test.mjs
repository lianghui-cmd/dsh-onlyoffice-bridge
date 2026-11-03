import assert from 'node:assert/strict'
import crypto from 'node:crypto'
import test from 'node:test'
import { authHeaders, parseFileKey, readEnvConfig, resolveConfig, signJwt, verifyJwt } from '../index.js'

test('signJwt 产出标准 HS256 三段式 token，并能被 verifyJwt 解回', () => {
  const secret = 'super-secret'
  const token = signJwt({ document: { key: 'k1' }, n: 1 }, secret)
  assert.equal(token.split('.').length, 3)
  const payload = verifyJwt(token, secret)
  assert.deepEqual(payload, { document: { key: 'k1' }, n: 1 })
  // 第三段是 base64url + HMAC-SHA256，和独立实现一致
  const [header, body, signature] = token.split('.')
  const expected = crypto.createHmac('sha256', secret).update(`${header}.${body}`).digest('base64')
    .replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '')
  assert.equal(signature, expected)
})

test('verifyJwt 拒绝被篡改的 payload 与错误密钥', () => {
  const token = signJwt({ a: 1 }, 'secret-a')
  assert.equal(verifyJwt(token, 'secret-b'), undefined)
  const [header, , signature] = token.split('.')
  const forgedBody = Buffer.from(JSON.stringify({ a: 2 }), 'utf8').toString('base64')
    .replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '')
  assert.equal(verifyJwt(`${header}.${forgedBody}.${signature}`, 'secret-a'), undefined)
  assert.equal(verifyJwt('not-a-jwt', 'secret-a'), undefined)
})

test('authHeaders 同时给出 Authorization 与自定义头', () => {
  const config = { jwtHeader: 'X-Office', jwtHeaderPrefixed: true, jwtSecret: 's' }
  const { token, headers } = authHeaders({ url: 'http://x/y' }, config)
  assert.equal(headers['X-Office'], `Bearer ${token}`)
  assert.equal(headers.Authorization, `Bearer ${token}`)
})

test('parseFileKey 只接受 <32hex>-<16hex>.<ext>', () => {
  const key = `${'a'.repeat(32)}-${'b'.repeat(16)}.docx`
  assert.deepEqual(parseFileKey(key), { stem: key.slice(0, -5), ext: 'docx', name: key })
  assert.equal(parseFileKey('../../etc/passwd'), undefined)
  assert.equal(parseFileKey('a'.repeat(32) + '.docx'), undefined)
  assert.equal(parseFileKey(`${'a'.repeat(32)}-${'b'.repeat(16)}`), undefined)
  assert.equal(parseFileKey(`${'a'.repeat(32)}-${'b'.repeat(16)}.docx/../x`), undefined)
})

test('readEnvConfig 读取环境变量并做地址归一化', () => {
  const config = readEnvConfig({
    cwd: 'C:\\dsh',
    dotenv: {
      ONLYOFFICE_DS_URL: 'office.local:8080/',
      ONLYOFFICE_SECRETS: 'legacy-secret',
      ONLYOFFICE_JWT_HEADER: 'X-Open',
      DSH_PUBLIC_URL: '192.168.1.20:11006',
      STORAGE_DIR: './cache',
      CACHE_TTL_HOUR: '6',
      MAX_FILE_MB: '32',
      ALLOWED_ORIGINS: 'https://a.example, https://b.example/',
    },
  })
  assert.equal(config.dsUrl, 'http://office.local:8080')
  assert.equal(config.jwtSecret, 'legacy-secret')
  assert.equal(config.jwtHeader, 'X-Open')
  assert.equal(config.publicUrl, 'http://192.168.1.20:11006')
  assert.equal(config.storageDir, 'C:\\dsh\\cache')
  assert.equal(config.cacheTtlHour, 6)
  assert.equal(config.maxFileMb, 32)
  assert.deepEqual(config.allowedOrigins, ['https://a.example', 'https://b.example'])
})

test('resolveConfig：ONLYOFFICE_SECRET 优先于 ONLYOFFICE_SECRETS，config 覆盖默认值但不覆盖环境变量', () => {
  const previous = { ONLYOFFICE_SECRET: process.env.ONLYOFFICE_SECRET, ONLYOFFICE_SECRETS: process.env.ONLYOFFICE_SECRETS }
  delete process.env.ONLYOFFICE_SECRET
  delete process.env.ONLYOFFICE_SECRETS
  try {
    const config = resolveConfig(
      { dotenv: { ONLYOFFICE_SECRETS: 'from-env-secrets', ONLYOFFICE_SECRET: 'from-env-secret' } },
      { jwtSecret: 'from-config', cacheTtlHour: 3, allowedOrigins: ['https://c.example/'] },
    )
    assert.equal(config.jwtSecret, 'from-env-secret')
    assert.equal(config.cacheTtlHour, 3)
    assert.deepEqual(config.allowedOrigins, ['https://c.example'])
  } finally {
    for (const [key, value] of Object.entries(previous)) {
      if (value === undefined) delete process.env[key]
      else process.env[key] = value
    }
  }
})
