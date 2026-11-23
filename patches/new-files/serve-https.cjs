// Minimal HTTPS static server for the self-hosted Word GPT Plus build.
// Uses the PFX produced by AmbientWriter\make-certs.ps1 (no openssl / PEM needed).
const https = require('https')
const fs = require('fs')
const path = require('path')
const os = require('os')

const root = path.join(__dirname, 'dist')
const pfxPath = path.join(process.env.LOCALAPPDATA, 'AmbientWriter', 'certs', 'server.pfx')

const mime = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'application/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.svg': 'image/svg+xml',
  '.ico': 'image/x-icon',
  '.woff': 'font/woff',
  '.woff2': 'font/woff2',
  '.map': 'application/json',
}

const server = https.createServer(
  { pfx: fs.readFileSync(pfxPath), passphrase: 'ambient' },
  (req, res) => {
    // ---- bridge API: lets the task pane drive the AmbientWriter watcher ----
    if (req.url.split('?')[0] === '/api/bridge') {
      const bridgePath = path.join(process.env.LOCALAPPDATA, 'AmbientWriter', 'bridge.json')
      if (req.method === 'GET') {
        let data = {}
        try {
          if (fs.existsSync(bridgePath)) data = JSON.parse(fs.readFileSync(bridgePath, 'utf8').replace(/^\uFEFF/, ''))
        } catch (e) {}
        res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' })
        res.end(JSON.stringify(data))
        return
      }
      if (req.method === 'POST') {
        let body = ''
        req.on('data', (c) => (body += c))
        req.on('end', () => {
          try {
            let cur = {}
            if (fs.existsSync(bridgePath)) cur = JSON.parse(fs.readFileSync(bridgePath, 'utf8').replace(/^\uFEFF/, ''))
            const patch = JSON.parse((body || '{}').replace(/^\uFEFF/, ''))
            const next = Object.assign(cur, patch)
            fs.mkdirSync(path.dirname(bridgePath), { recursive: true })
            fs.writeFileSync(bridgePath, JSON.stringify(next, null, 2), 'utf8')
            res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' })
            res.end(JSON.stringify(next))
          } catch (e) {
            res.writeHead(400, { 'Content-Type': 'application/json; charset=utf-8' })
            res.end(JSON.stringify({ error: String(e) }))
          }
        })
        return
      }
      res.writeHead(405).end()
      return
    }

    if (req.url.split('?')[0] === '/api/status') {
      const statusPath = path.join(process.env.LOCALAPPDATA, 'AmbientWriter', 'status.json')
      let data = {}
      try {
        if (fs.existsSync(statusPath)) data = JSON.parse(fs.readFileSync(statusPath, 'utf8').replace(/^\uFEFF/, ''))
      } catch (e) {}
      res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' })
      res.end(JSON.stringify(data))
      return
    }

    let p = decodeURIComponent(req.url.split('?')[0])
    if (p === '/' || p === '') p = '/index.html'
    let file = path.join(root, p)
    if (!file.startsWith(root)) {
      res.writeHead(403).end()
      return
    }
    if (!fs.existsSync(file) || fs.statSync(file).isDirectory()) {
      file = path.join(root, 'index.html') // SPA fallback
    }
    res.writeHead(200, { 'Content-Type': mime[path.extname(file).toLowerCase()] || 'application/octet-stream' })
    fs.createReadStream(file).pipe(res)
  },
)

server.listen(3000, '0.0.0.0', () => {
  console.log('listening on https://0.0.0.0:3000  (host=' + os.hostname() + ')')
})
