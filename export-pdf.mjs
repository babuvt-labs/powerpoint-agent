import puppeteer from 'puppeteer';
import path from 'path';
import { fileURLToPath } from 'url';
import fs from 'fs';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const sourceArg = process.argv[2];
const outputArg = process.argv[3];

const sourcePath = sourceArg
  ? path.resolve(__dirname, sourceArg)
  : path.join(__dirname, 'OCP-SOW-pptx', 'powerpnt.html');
const outputPath = outputArg
  ? path.resolve(__dirname, outputArg)
  : path.join(__dirname, 'OCP-SOW-pptx', 'powerpnt.pdf');
const tempPath = path.join(path.dirname(outputPath), '_print_tmp.html');

async function enforceMinFontSize(page, minPx = 12) {
  await page.evaluate((minFontPx) => {
    const textTags = new Set([
      'P', 'SPAN', 'A', 'LI', 'TD', 'TH', 'CAPTION', 'LABEL', 'SMALL',
      'STRONG', 'EM', 'B', 'I', 'U', 'BUTTON', 'INPUT', 'TEXTAREA',
      'SELECT', 'OPTION', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'DIV'
    ]);

    const all = Array.from(document.querySelectorAll('*'));
    for (const el of all) {
      if (!textTags.has(el.tagName)) continue;
      const style = window.getComputedStyle(el);
      const fontSize = parseFloat(style.fontSize);
      if (!Number.isFinite(fontSize) || fontSize >= minFontPx) continue;

      el.style.fontSize = `${minFontPx}px`;

      const lineHeight = parseFloat(style.lineHeight);
      if (Number.isFinite(lineHeight) && lineHeight < minFontPx * 1.2) {
        el.style.lineHeight = `${Math.ceil(minFontPx * 1.25)}px`;
      }
    }
  }, minPx);
}

(async () => {
  if (!fs.existsSync(sourcePath)) {
    throw new Error(`Source file not found: ${sourcePath}`);
  }

  const browser = await puppeteer.launch({
    headless: 'new',
    args: ['--no-sandbox', '--disable-setuid-sandbox']
  });

  const page = await browser.newPage();
  await page.setViewport({ width: 1600, height: 900, deviceScaleFactor: 1 });
  await page.goto(`file://${sourcePath}`, { waitUntil: 'networkidle0', timeout: 60000 });
  await page.emulateMediaType('screen');

  // Ensure web fonts and external images are fully ready before rendering.
  await page.evaluate(async () => {
    await document.fonts.ready;
    const imgs = Array.from(document.images);
    await Promise.all(
      imgs.map((img) => {
        if (img.complete) return Promise.resolve();
        return new Promise((resolve) => {
          img.addEventListener('load', resolve, { once: true });
          img.addEventListener('error', resolve, { once: true });
        });
      })
    );
  });
  await enforceMinFontSize(page, 12);

  const { styleTag, slidesMarkup, hasSlides } = await page.evaluate(() => {
    const style = document.querySelector('style');
    const slides = Array.from(document.querySelectorAll('.slide')).map((slide) => slide.outerHTML);

    return {
      styleTag: style ? style.outerHTML : '',
      slidesMarkup: slides.length > 0 ? slides.join('\n') : document.body.innerHTML,
      hasSlides: slides.length > 0
    };
  });

  await page.close();

  if (!hasSlides) {
    const directPage = await browser.newPage();
    await directPage.setViewport({ width: 1600, height: 900, deviceScaleFactor: 1 });
    await directPage.goto(`file://${sourcePath}`, { waitUntil: 'networkidle0', timeout: 60000 });
    await directPage.emulateMediaType('screen');
    await directPage.evaluate(async () => {
      await document.fonts.ready;
      const imgs = Array.from(document.images);
      await Promise.all(
        imgs.map((img) => {
          if (img.complete) return Promise.resolve();
          return new Promise((resolve) => {
            img.addEventListener('load', resolve, { once: true });
            img.addEventListener('error', resolve, { once: true });
          });
        })
      );
    });
    await enforceMinFontSize(directPage, 12);

    // Normalize centered presentation layouts so exported content is anchored and cropped cleanly.
    const hasSlideContainer = await directPage.evaluate(() => !!document.querySelector('.slide-container'));
    if (hasSlideContainer) {
      await directPage.addStyleTag({
        content: `
          html, body {
            margin: 0 !important;
            padding: 0 !important;
            overflow: visible !important;
            background: #ffffff !important;
            width: 100vw !important;
            height: 100vh !important;
          }
          body {
            display: block !important;
            min-height: 100vh !important;
          }
          .slide-container {
            margin: 0 !important;
            width: 100vw !important;
            height: 100vh !important;
            max-width: none !important;
            border-radius: 0 !important;
            box-shadow: none !important;
          }
        `
      });
    }

    const pageSize = await directPage.evaluate(() => {
      const container = document.querySelector('.slide-container');
      if (container) {
        const rect = container.getBoundingClientRect();
        return {
          width: Math.ceil(rect.width),
          height: Math.ceil(rect.height)
        };
      }

      const doc = document.documentElement;
      return {
        width: Math.ceil(Math.max(doc.scrollWidth, doc.clientWidth)),
        height: Math.ceil(Math.max(doc.scrollHeight, doc.clientHeight))
      };
    });

    await directPage.pdf({
      path: outputPath,
      width: `${pageSize.width}px`,
      height: `${pageSize.height}px`,
      printBackground: true,
      margin: { top: 0, right: 0, bottom: 0, left: 0 },
      preferCSSPageSize: false
    });

    await directPage.close();
    console.log(`PDF saved to ${outputPath}`);
    await browser.close();
    return;
  }

  const printHtml = `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>VMware to OpenShift Workshop Deck</title>
  ${styleTag}
  <style>
    html, body {
      background: #ffffff !important;
      overflow: visible !important;
      width: auto !important;
      height: auto !important;
    }

    #toolbar,
    #dotbar,
    #hint,
    #viewport,
    #scaler {
      display: none !important;
    }

    body {
      margin: 0 !important;
      padding: 0 !important;
    }

    .print-deck {
      display: block;
      width: 1600px;
      margin: 0;
      padding: 0;
    }

    .print-deck .slide {
      position: relative !important;
      inset: auto !important;
      opacity: 1 !important;
      transform: none !important;
      pointer-events: all !important;
      display: grid !important;
      width: 1600px !important;
      height: 900px !important;
      margin: 0 !important;
      border-radius: 0 !important;
      box-shadow: none !important;
      page-break-after: always;
      break-after: page;
    }

    .print-deck .slide:last-child {
      page-break-after: auto;
      break-after: auto;
    }
  </style>
</head>
<body>
  <main class="print-deck">
    ${slidesMarkup}
  </main>
</body>
</html>`;

  fs.writeFileSync(tempPath, printHtml, 'utf8');

  const printPage = await browser.newPage();
  await printPage.setViewport({ width: 1600, height: 900, deviceScaleFactor: 1 });
  await printPage.goto(`file://${tempPath}`, { waitUntil: 'networkidle0', timeout: 60000 });
  await printPage.emulateMediaType('screen');
  await enforceMinFontSize(printPage, 12);
  await printPage.pdf({
    path: outputPath,
    width: '1600px',
    height: '900px',
    printBackground: true,
    margin: { top: 0, right: 0, bottom: 0, left: 0 }
  });
  await printPage.close();

  console.log(`PDF saved to ${outputPath}`);

  if (fs.existsSync(tempPath)) {
    fs.unlinkSync(tempPath);
  }

  await browser.close();
})();
