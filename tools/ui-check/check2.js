// End-to-end check: type an API key in the pane -> POST /api/bridge -> bridge.json
const puppeteer = require('puppeteer-core')
const fs = require('fs')
const os = require('os')
const path = require('path')

const EDGE = 'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe'
const URL = `https://${os.hostname()}:3000`
const BRIDGE = path.join(process.env.LOCALAPPDATA, 'AmbientWriter', 'bridge.json')

;(async () => {
  const browser = await puppeteer.launch({
    executablePath: EDGE,
    headless: 'new',
    args: ['--ignore-certificate-errors', '--disable-gpu', '--no-sandbox'],
  })
  const page = await browser.newPage()
  await page.setViewport({ width: 430, height: 940 })
  await page.goto(URL, { waitUntil: 'networkidle2', timeout: 30000 })
  await new Promise((r) => setTimeout(r, 2000))

  await page.evaluate(() => {
    const b = Array.from(document.querySelectorAll('button')).find((x) =>
      ['\u8bbe\u7f6e', 'Settings'].includes(x.getAttribute('title') || ''),
    )
    if (b) b.click()
  })
  await new Promise((r) => setTimeout(r, 1200))
  await page.evaluate(() => {
    const b = Array.from(document.querySelectorAll('button')).find((x) =>
      ['\u4f34\u5199', 'Companion'].includes(x.getAttribute('title') || ''),
    )
    if (b) b.click()
  })
  await new Promise((r) => setTimeout(r, 1200))

  const typed = await page.evaluate(() => {
    const input = Array.from(document.querySelectorAll('input')).find((i) => (i.placeholder || '').includes('sk-'))
    if (!input) return false
    const setter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set
    setter.call(input, 'sk-e2e-12345')
    input.dispatchEvent(new Event('input', { bubbles: true }))
    input.dispatchEvent(new Event('change', { bubbles: true }))
    return true
  })

  await new Promise((r) => setTimeout(r, 2000))
  const file = fs.existsSync(BRIDGE) ? fs.readFileSync(BRIDGE, 'utf8') : '(missing)'
  console.log(JSON.stringify({ typed, bridge: JSON.parse(file) }, null, 2))
  await browser.close()
})().catch((e) => {
  console.error('FAIL', e)
  process.exit(1)
})
