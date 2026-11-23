// Headless UI check for the self-hosted Word GPT Plus pane.
// Opens the app in the installed Edge, clicks the settings gear, then the 伴写 tab,
// and saves screenshots so we can prove the new panel renders.
const puppeteer = require('puppeteer-core')
const path = require('path')
const os = require('os')

const EDGE = 'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe'
const URL = `https://${os.hostname()}:3000`

;(async () => {
  const browser = await puppeteer.launch({
    executablePath: EDGE,
    headless: 'new',
    args: ['--ignore-certificate-errors', '--disable-gpu', '--no-sandbox'],
  })
  const page = await browser.newPage()
  await page.setViewport({ width: 430, height: 940 })
  await page.goto(URL, { waitUntil: 'networkidle2', timeout: 30000 })
  await new Promise((r) => setTimeout(r, 2500))
  await page.screenshot({ path: path.join(__dirname, 'shot-home.png') })

  // settings gear: the header buttons carry a title attribute
  const clickedGear = await page.evaluate(() => {
    const b = Array.from(document.querySelectorAll('button')).find((x) => ['设置','Settings'].includes(x.getAttribute('title') || ''))
    if (b) {
      b.click()
      return true
    }
    return false
  })
  await new Promise((r) => setTimeout(r, 1500))
  await page.screenshot({ path: path.join(__dirname, 'shot-settings.png') })

  const tabs = await page.evaluate(() =>
    Array.from(document.querySelectorAll('button'))
      .map((b) => b.getAttribute('title') || '')
      .filter((t) => t),
  )

  const clickedTab = await page.evaluate(() => {
    const b = Array.from(document.querySelectorAll('button')).find((x) => ['伴写','Companion'].includes(x.getAttribute('title') || ''))
    if (b) {
      b.click()
      return true
    }
    return false
  })
  await new Promise((r) => setTimeout(r, 1200))
  await page.screenshot({ path: path.join(__dirname, 'shot-ambient.png') })

  const texts = await page.evaluate(() => document.body.innerText)

  console.log(JSON.stringify({ clickedGear, clickedTab, tabs, hasAmbientText: texts.includes('启用伴写'), snippet: texts.slice(0, 400) }, null, 2))
  await browser.close()
})().catch((e) => {
  console.error('FAIL', e)
  process.exit(1)
})
