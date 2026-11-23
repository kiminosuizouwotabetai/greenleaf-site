// Импорт каталога из Excel в SQLite.
//
//   node scripts/import-catalog.mjs            — обычный запуск
//   node scripts/import-catalog.mjs --force    — перекачать все картинки заново
//
// Источники:
//   data/greenleaf.xlsx           — товары из таблицы
//   data/products-curated.json    — товары, заведённые вручную (не из Excel)
// Результат:
//   data/catalog.db               — таблица products
//   public/images/catalog/*.webp  — картинки товаров

import fs from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';
import sharp from 'sharp';
import Database from 'better-sqlite3';

const require = createRequire(import.meta.url);
const XLSX = require('xlsx');

const ROOT = process.cwd();
const XLSX_PATH = path.join(ROOT, 'data', 'greenleaf.xlsx');
const CURATED_PATH = path.join(ROOT, 'data', 'products-curated.json');
const DB_PATH = path.join(ROOT, 'data', 'catalog.db');
const IMG_DIR = path.join(ROOT, 'public', 'images', 'catalog');

const FORCE = process.argv.includes('--force');

const SUBCATEGORY_ORDER = [
  'Очищение и демакияж',
  'Основной уход',
  'Комплексный уход',
  'Маски для лица',
];

const fixCase = (s) => {
  const t = String(s || '').trim().replace(/\s+/g, ' ');
  if (!t) return '';
  return t[0].toUpperCase() + t.slice(1).toLowerCase();
};

const escapeHtml = (s) =>
  String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');

// «Артикул SAA069» -> «SAA069»
const cleanArticle = (raw) =>
  String(raw || '')
    .replace(/^\s*артикул\s*:?\s*/i, '')
    .trim();

const articleToFile = (article) => article.replace(/[\\/:*?"<>|]/g, '-');

// Текст из Excel -> HTML: абзацы <p>, строки «- …» -> <ul><li>
const textToHtml = (raw) => {
  const text = String(raw || '').replace(/\r\n/g, '\n').trim();
  if (!text || /^[-—–\s]*$/.test(text)) return '';

  const blocks = text
    .split(/\n\s*\n/)
    .map((b) => b.trim())
    .filter(Boolean);

  return blocks
    .map((block) => {
      const lines = block
        .split('\n')
        .map((l) => l.trim())
        .filter(Boolean);

      if (lines.length > 0 && lines.every((l) => /^[-•*]\s+/.test(l))) {
        const items = lines
          .map((l) => `<li>${escapeHtml(l.replace(/^[-•*]\s+/, ''))}</li>`)
          .join('');
        return `<ul>${items}</ul>`;
      }

      return `<p>${lines.map(escapeHtml).join(' ')}</p>`;
    })
    .join('');
};

const stripTags = (html) =>
  String(html || '')
    .replace(/<[^>]+>/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();

const makeShort = (html) => {
  const plain = stripTags(html);
  if (!plain) return '';
  const sentence = plain.split(/(?<=[.!?])\s/)[0];
  const base = sentence.length > 20 ? sentence : plain;
  return base.length > 120 ? base.slice(0, 117).trimEnd() + '…' : base;
};

const withArticle = (article, html) =>
  `<p><strong>Артикул:</strong> ${escapeHtml(article)}</p>${html || ''}`;

// убирает уже вставленный ранее абзац «Артикул: …», чтобы не дублировать его
const stripArticleParagraph = (html) =>
  String(html || '').replace(/^\s*<p>\s*<strong>\s*Артикул:?\s*<\/strong>[\s\S]*?<\/p>/i, '');

async function downloadImage(url, article) {
  const fileName = articleToFile(article) + '.webp';
  const dest = path.join(IMG_DIR, fileName);

  if (!FORCE && fs.existsSync(dest)) {
    return { file: fileName, cached: true };
  }

  const res = await fetch(url, {
    headers: { 'User-Agent': 'Mozilla/5.0 (greenleaf-site importer)' },
  });
  if (!res.ok) throw new Error(`HTTP ${res.status} для ${url}`);

  const buf = Buffer.from(await res.arrayBuffer());
  await sharp(buf)
    .resize(700, 700, { fit: 'inside', withoutEnlargement: true })
    .webp({ quality: 82 })
    .toFile(dest);

  return { file: fileName, cached: false };
}

function readExcel() {
  if (!fs.existsSync(XLSX_PATH)) {
    throw new Error(`Не найден ${XLSX_PATH}`);
  }

  const wb = XLSX.readFile(XLSX_PATH);
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { defval: null });

  const products = [];
  const skipped = [];

  for (const row of rows) {
    const name = String(row['название'] ?? '').trim();
    const article = cleanArticle(row['артикул']);
    const image = String(row['картинка'] ?? '').trim();
    const category = fixCase(row['категория']);
    const subcategory = fixCase(row['подкатегория']);

    // служебные строки: пустые, повторные заголовки, названия подкатегорий
    if (!name || !article || !image || !category || !subcategory || !/^https?:/i.test(image)) {
      skipped.push(name || article || '(пустая строка)');
      continue;
    }

    products.push({
      article,
      title: name,
      description: textToHtml(row['описание']),
      category,
      subcategory,
      image,
    });
  }

  return { products, skipped };
}

function readCurated() {
  if (!fs.existsSync(CURATED_PATH)) return [];
  const raw = JSON.parse(fs.readFileSync(CURATED_PATH, 'utf8'));
  return Array.isArray(raw) ? raw : [];
}

function createSchema(db) {
  db.exec(`
    DROP TABLE IF EXISTS products;
    CREATE TABLE products (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      article TEXT NOT NULL UNIQUE,
      title TEXT NOT NULL,
      description TEXT NOT NULL DEFAULT '',
      short TEXT NOT NULL DEFAULT '',
      category TEXT NOT NULL,
      subcategory TEXT NOT NULL,
      images TEXT NOT NULL,
      featured INTEGER NOT NULL DEFAULT 0,
      sort_order INTEGER NOT NULL DEFAULT 0
    );
    CREATE INDEX IF NOT EXISTS idx_products_sort ON products (sort_order);
  `);
}

async function main() {
  fs.mkdirSync(IMG_DIR, { recursive: true });

  const { products: fromExcel, skipped } = readExcel();
  const curated = readCurated();

  console.log(`Строк в Excel: товаров ${fromExcel.length}, отброшено служебных ${skipped.length}`);

  const items = [];

  for (const p of curated) {
    items.push({
      article: p.article,
      title: p.title,
      description: withArticle(p.article, stripArticleParagraph(p.description)),
      short: p.short || makeShort(p.description),
      category: fixCase(p.category),
      subcategory: fixCase(p.subcategory),
      image: '/images/catalog/' + articleToFile(p.article) + '.webp',
      local: true,
      curated: true,
    });
  }

  let downloaded = 0;
  let cached = 0;

  for (const p of fromExcel) {
    const { file, cached: isCached } = await downloadImage(p.image, p.article);
    if (isCached) cached += 1;
    else downloaded += 1;

    items.push({
      article: p.article,
      title: p.title,
      description: withArticle(p.article, p.description),
      short: makeShort(p.description),
      category: p.category,
      subcategory: p.subcategory,
      image: '/images/catalog/' + file,
      local: true,
      curated: false,
    });
  }

  console.log(`Картинок: скачано ${downloaded}, из кеша ${cached}`);

  // порядок: группы по подкатегориям, внутри — сначала ручные, потом из Excel
  const rank = (sub) => {
    const i = SUBCATEGORY_ORDER.indexOf(sub);
    return i === -1 ? SUBCATEGORY_ORDER.length : i;
  };

  items.sort((a, b) => {
    const bySub = rank(a.subcategory) - rank(b.subcategory);
    if (bySub !== 0) return bySub;
    const byGroup = a.subcategory.localeCompare(b.subcategory, 'ru');
    if (byGroup !== 0) return byGroup;
    if (a.curated !== b.curated) return a.curated ? -1 : 1;
    return 0;
  });

  const db = new Database(DB_PATH);
  createSchema(db);

  const insert = db.prepare(`
    INSERT INTO products (article, title, description, short, category, subcategory, images, featured, sort_order)
    VALUES (@article, @title, @description, @short, @category, @subcategory, @images, @featured, @sort_order)
  `);

  const insertAll = db.transaction((rows) => {
    rows.forEach((row, index) => {
      insert.run({
        ...row,
        images: JSON.stringify([row.image]),
        featured: 0,
        sort_order: index,
      });
    });
  });

  insertAll(items);

  const total = db.prepare('SELECT COUNT(*) AS n FROM products').get().n;
  const bySub = db
    .prepare('SELECT subcategory, COUNT(*) AS n FROM products GROUP BY subcategory ORDER BY sort_order')
    .all();

  db.close();

  console.log(`\nВ базе товаров: ${total}`);
  for (const r of bySub) console.log(`  ${r.subcategory}: ${r.n}`);
  console.log(`\nБаза: ${DB_PATH}`);
}

main().catch((err) => {
  console.error('Ошибка импорта:', err.message);
  process.exit(1);
});
