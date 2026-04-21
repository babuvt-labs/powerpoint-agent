import fs from 'fs';
import os from 'os';
import path from 'path';
import { fileURLToPath } from 'url';
import puppeteer from 'puppeteer';
import PptxGenJS from 'pptxgenjs';

const __dirname = path.dirname(fileURLToPath(import.meta.url));

const sourceArg = process.argv[2];
const outputArg = process.argv[3];

const sourcePath = sourceArg
  ? path.resolve(__dirname, sourceArg)
  : path.join(__dirname, 'OCP-SOW-pptx', 'slide.html');
const outputPath = outputArg
  ? path.resolve(__dirname, outputArg)
  : path.join(__dirname, 'OCP-SOW-pptx', 'slide-editable.pptx');

if (!fs.existsSync(sourcePath)) {
  throw new Error(`Source file not found: ${sourcePath}`);
}

function sanitizeBaseName(filePath) {
  return path.basename(filePath).replace(/[^a-zA-Z0-9._-]/g, '_');
}

async function captureSlides(page, sourceFilePath, tempDir) {
  await page.goto(`file://${sourceFilePath}`, {
    waitUntil: 'networkidle0',
    timeout: 120000
  });

  await page.evaluate(async () => {
    await document.fonts.ready;
    const images = Array.from(document.images);
    await Promise.all(
      images.map((img) => {
        if (img.complete) return Promise.resolve();
        return new Promise((resolve) => {
          img.addEventListener('load', resolve, { once: true });
          img.addEventListener('error', resolve, { once: true });
        });
      })
    );
  });

  const meta = await page.evaluate(() => {
    const docEl = document.documentElement;
    const pageWidth = Math.max(docEl.scrollWidth, docEl.clientWidth, window.innerWidth);
    const pageHeight = Math.max(docEl.scrollHeight, docEl.clientHeight, window.innerHeight);

    const normalizeClip = (clip) => {
      let x = Number.isFinite(clip.x) ? clip.x : 0;
      let y = Number.isFinite(clip.y) ? clip.y : 0;
      let width = Number.isFinite(clip.width) ? clip.width : pageWidth;
      let height = Number.isFinite(clip.height) ? clip.height : pageHeight;

      x = Math.max(0, Math.floor(x));
      y = Math.max(0, Math.floor(y));
      width = Math.max(1, Math.ceil(width));
      height = Math.max(1, Math.ceil(height));

      if (x + width > pageWidth) {
        width = Math.max(1, pageWidth - x);
      }
      if (y + height > pageHeight) {
        height = Math.max(1, pageHeight - y);
      }

      return { x, y, width, height };
    };

    const candidates = [
      ...document.querySelectorAll('.slide'),
      ...document.querySelectorAll('.slide-container'),
      ...document.querySelectorAll('[data-slide]')
    ];

    const unique = [];
    for (const el of candidates) {
      if (!unique.includes(el)) unique.push(el);
    }

    if (unique.length === 0) {
      return {
        count: 1,
        slides: [
          normalizeClip({
            x: 0,
            y: 0,
            width: pageWidth,
            height: pageHeight
          })
        ]
      };
    }

    const slides = unique
      .map((el) => {
        const rect = el.getBoundingClientRect();
        return normalizeClip({
          x: Math.max(0, Math.floor(rect.left + window.scrollX)),
          y: Math.max(0, Math.floor(rect.top + window.scrollY)),
          width: Math.max(1, Math.ceil(rect.width)),
          height: Math.max(1, Math.ceil(rect.height))
        });
      })
      .filter(
        (s) =>
          Number.isFinite(s.x) &&
          Number.isFinite(s.y) &&
          Number.isFinite(s.width) &&
          Number.isFinite(s.height) &&
          s.width > 0 &&
          s.height > 0
      );

    return { count: slides.length, slides };
  });

  if (!meta.count) {
    throw new Error('No renderable slide area found in HTML.');
  }

  const images = [];
  for (let i = 0; i < meta.slides.length; i += 1) {
    const clip = meta.slides[i];
    const filePath = path.join(tempDir, `slide-${String(i + 1).padStart(3, '0')}.png`);

    await page.screenshot({
      path: filePath,
      clip,
      type: 'png',
      captureBeyondViewport: true
    });

    images.push({ path: filePath, width: clip.width, height: clip.height });
  }

  return images;
}

function addImagesToPptx(images, outputFilePath, sourceFilePath) {
  const pptx = new PptxGenJS();
  pptx.layout = 'LAYOUT_WIDE';
  pptx.author = 'GitHub Copilot';
  pptx.subject = 'Converted from HTML';
  pptx.company = 'powerpnt-html';
  pptx.title = sanitizeBaseName(sourceFilePath);
  pptx.lang = 'en-US';

  const slideW = 13.333;
  const slideH = 7.5;

  for (const image of images) {
    const slide = pptx.addSlide();
    const ratio = image.width / image.height;
    const slideRatio = slideW / slideH;

    let w;
    let h;
    let x;
    let y;

    if (ratio > slideRatio) {
      w = slideW;
      h = slideW / ratio;
      x = 0;
      y = (slideH - h) / 2;
    } else {
      h = slideH;
      w = slideH * ratio;
      y = 0;
      x = (slideW - w) / 2;
    }

    slide.addImage({ path: image.path, x, y, w, h });
  }

  return pptx.writeFile({ fileName: outputFilePath });
}

const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'html2pptx-'));
let browser;

try {
  browser = await puppeteer.launch({
    headless: 'new',
    args: ['--no-sandbox', '--disable-setuid-sandbox']
  });

  const page = await browser.newPage();
  await page.setViewport({ width: 1920, height: 1080, deviceScaleFactor: 2 });

  const images = await captureSlides(page, sourcePath, tmpDir);
  await addImagesToPptx(images, outputPath, sourcePath);

  await page.close();
  console.log(`PPTX saved to ${outputPath}`);
} finally {
  if (browser) {
    await browser.close();
  }

  if (fs.existsSync(tmpDir)) {
    for (const file of fs.readdirSync(tmpDir)) {
      fs.unlinkSync(path.join(tmpDir, file));
    }
    fs.rmdirSync(tmpDir);
  }
}
