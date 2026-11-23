// Синхронизация каталога c маркетплейса greenleaf.online.
//
//   node scripts/sync-catalog.mjs               — полный цикл
//   node scripts/sync-catalog.mjs --use-cache   — без повторного обхода категорий
//   node scripts/sync-catalog.mjs --refresh-desc — перекачать описания
//   node scripts/sync-catalog.mjs --debug       — подробности
//
// Результат:
//   data/catalog.db               — товары
//   data/greenleaf-full.xlsx      — выгрузка для просмотра
//   public/images/catalog/*.webp  — картинки
//   data/catalog.json             — данные для страницы каталога (через эндпоинт)

import fs from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';
import * as cheerio from 'cheerio';
import sharp from 'sharp';
import Database from 'better-sqlite3';

const require = createRequire(import.meta.url);
const XLSX = require('xlsx');

const ROOT = process.cwd();
const MARKET = 'https://greenleaf.online';
const CACHE = path.join(ROOT, 'data', 'marketplace-cache.json');
const DESC_CACHE = path.join(ROOT, 'data', 'descriptions-cache.json');
const CAT_CACHE = path.join(ROOT, 'data', 'categories-cache.json');
const LEGACY_XLSX = path.join(ROOT, 'data', 'greenleaf.xlsx');
const DB_PATH = path.join(ROOT, 'data', 'catalog.db');
const XLSX_OUT = path.join(ROOT, 'data', 'greenleaf-full.xlsx');
const IMG_DIR = path.join(ROOT, 'public', 'images', 'catalog');

const USE_CACHE = process.argv.includes('--use-cache');
const REFRESH_DESC = process.argv.includes('--refresh-desc');
const DEBUG = process.argv.includes('--debug');

const UA =
  'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36';

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

async function pool(items, limit, worker) {
  let index = 0;
  const runners = Array.from({ length: Math.min(limit, items.length) }, async () => {
    while (index < items.length) {
      const i = index;
      index += 1;
      await worker(items[i], i);
    }
  });
  await Promise.all(runners);
}

async function fetchText(url, tries = 3) {
  for (let i = 0; i < tries; i += 1) {
    try {
      const res = await fetch(url, { headers: { 'User-Agent': UA } });
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      return await res.text();
    } catch (err) {
      if (i === tries - 1) throw err;
      await sleep(1000);
    }
  }
  return '';
}

// ---------- утилиты ----------

const BRAND_WORDS = ['SEALUXE', 'iLiFE', 'CARICH', 'NILRICH', 'KARDLI', 'SONPHESKIN', 'GREENLEAF', 'YUE ZHONGGUI', 'MOON FLOWER'];

const brandOf = (name) => {
  const lower = String(name || '').toLowerCase();
  const hit = BRAND_WORDS.find((b) => lower.startsWith(b.toLowerCase()));
  if (!hit) return '';
  return hit === 'ILIFE' ? 'iLiFE' : hit;
};

const volumeOf = (name) => {
  const m = String(name || '').match(/(\d[\d\s.,]*)\s*(мл|гр|г|кг|шт|штук|пар[аы]?|пакет[а-яё]*)(?![а-яё])/i);
  return m ? `${m[1].trim()} ${m[2].toLowerCase()}` : '';
};

const cleanArticle = (raw) =>
  String(raw || '')
    .replace(/^\s*артикул\s*:?\s*/i, '')
    .trim();

const escapeHtml = (s) =>
  String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

const stripTags = (s) => String(s || '').replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim();

function textToHtml(text) {
  const cleaned = String(text || '')
    .replace(/\u00a0/g, ' ')
    .replace(/([.!?])(?=[А-ЯЁA-Z«"(])/g, '$1 ');

  const blocks = cleaned
    .split(/\n\s*\n/)
    .map((b) => b.trim())
    .filter(Boolean);

  return blocks
    .map((block) => {
      const lines = block.split('\n').map((l) => l.trim()).filter(Boolean);
      if (lines.length && lines.every((l) => /^[➤•✔\-–—]\s*/.test(l))) {
        return `<ul>${lines.map((l) => `<li>${escapeHtml(l.replace(/^[➤•✔\-–—]\s*/, ''))}</li>`).join('')}</ul>`;
      }
      return `<p>${escapeHtml(lines.join(' '))}</p>`;
    })
    .join('');
}

const stripEntities = (s) =>
  String(s || '')
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&quot;/g, '"')
    .replace(/&laquo;/g, '«')
    .replace(/&raquo;/g, '»')
    .replace(/&mdash;/g, '—')
    .replace(/&ndash;/g, '–')
    .replace(/&[a-z]+;/g, ' ')
    .replace(/[ \t]+/g, ' ');

// ---------- категории ----------

const ICON_WORDS = new Set(['apps', 'subdirectory_arrow_right', 'chevron_right', 'Все в категории']);

function cleanMenuText(text) {
  const t = String(text || '').replace(/\s+/g, ' ').trim();
  return t
    .replace(/^apps\s+/i, '')
    .replace(/^subdirectory_arrow_right\s+/i, '')
    .replace(/^chevron_right\s+/i, '')
    .replace(/\s*\d+$/, '')
    .trim();
}

async function parseCategories() {
  if (USE_CACHE && fs.existsSync(CAT_CACHE)) {
    const cached = JSON.parse(fs.readFileSync(CAT_CACHE, 'utf8'));
    console.log(`Категории из кеша: ${cached.groups.length} категорий`);
    return cached;
  }

  console.log('Разбираю структуру каталога…');
  const html = await fetchText(MARKET);
  const $ = cheerio.load(html);

  // id верхнеуровневых разделов каталога (видны в меню как «Все в категории»)
  const GROUP_IDS = new Set([
    'oc-1', 'oc-2', 'oc-3', 'oc-4', 'oc-5', 'oc-6', 'oc-7',
    'oc-8', 'oc-9', 'oc-10', 'oc-11', 'oc-12', 'oc-13', 'oc-unassigned',
  ]);

  const links = [];
  $('a[href*="category=oc-"]').each((_, el) => {
    const a = $(el);
    const id = (a.attr('href') || '').match(/category=(oc-[a-z0-9-]+)/)?.[1];
    const text = cleanMenuText(a.text());
    if (id && text) links.push({ id, text });
  });

  const groups = [];
  const seenSubs = new Set();
  let current = null;

  for (const link of links) {
    if (GROUP_IDS.has(link.id)) {
      // «Все в категории» того же раздела возвращает нас к нему
      const existing = groups.find((g) => g.id === link.id);
      if (existing) {
        current = existing;
      } else {
        current = { id: link.id, name: '', subs: [] };
        groups.push(current);
      }
      continue;
    }

    if (!current) continue;
    if (seenSubs.has(link.id)) continue; // вторая копия меню
    seenSubs.add(link.id);
    current.subs.push({ id: link.id, name: link.text });
  }

  // названия разделов — из h1 их страниц
  for (const group of groups) {
    try {
      const page = await fetchText(`${MARKET}/?category=${group.id}`);
      const $page = cheerio.load(page);
      const name = $page('h1').first().text().replace(/\s+/g, ' ').trim().replace(/\s*Greenleaf\s*$/i, '');
      group.name = name || group.id;
      await sleep(120);
    } catch {
      group.name = group.id;
    }
  }

  const result = { groups };
  fs.writeFileSync(CAT_CACHE, JSON.stringify(result, null, 1));

  console.log(`  разделов: ${result.groups.length}, подкатегорий: ${result.groups.reduce((n, g) => n + g.subs.length, 0)}`);
  for (const g of result.groups) console.log(`    ${g.name}: ${g.subs.length} подкатегорий`);

  return result;
}

// ---------- обход товаров ----------

function parseCards(html) {
  const $ = cheerio.load(html);
  const out = [];

  $('article.marketplace-product-card').each((_, el) => {
    const card = $(el);
    const art = card.find('small').first().text().match(/Артикул\s+([A-Z0-9/\-]+)/i);
    if (!art) return;

    out.push({
      article: art[1].trim(),
      name: card.find('.marketplace-product-card__name').text().replace(/\s+/g, ' ').trim(),
      image: card.find('img').attr('src') || '',
      price: card.find('strong').first().text().replace(/\s+/g, ' ').trim(),
      stock: card.find('.marketplace-product-card__stock').text().replace(/\s+/g, ' ').trim(),
      url: card.find('.marketplace-product-card__name').attr('href') || '',
    });
  });

  return out;
}

async function crawlCategory(id, label) {
  const products = [];
  for (let page = 1; page <= 90; page += 1) {
    const url = `${MARKET}/?category=${id}&page=${page}`;
    const html = await fetchText(url);
    const cards = parseCards(html);
    if (!cards.length) break;
    cards.forEach((c) => products.push(c));
    await sleep(180);
  }
  if (DEBUG) console.log(`    ${label}: ${products.length}`);
  return products;
}

async function crawlAll(categories) {
  if (USE_CACHE && fs.existsSync(CACHE)) {
    const cached = JSON.parse(fs.readFileSync(CACHE, 'utf8'));
    cached.forEach((p) => {
      p.brand = brandOf(p.name);
      p.volume = volumeOf(p.name);
    });
    console.log(`Товары из кеша: ${cached.length}`);
    return cached;
  }

  console.log('Обхожу каталог…');
  const collected = [];

  // сначала подкатегории — чтобы товар получил самую точную категорию
  for (const [i, group] of categories.groups.entries()) {
    for (const sub of group.subs) {
      const items = await crawlCategory(sub.id, `${group.name} / ${sub.name}`);
      items.forEach((it) => collected.push({ ...it, parent: group.name, subcategory: sub.name }));
    }
    if ((i + 1) % 3 === 0) console.log(`  ${i + 1}/${categories.groups.length} разделов, товаров ${collected.length}`);
  }

  // затем сами разделы — добираем всё остальное («Прочее»)
  for (const group of categories.groups) {
    const items = await crawlCategory(group.id, group.name);
    items.forEach((it) => collected.push({ ...it, parent: group.name, subcategory: 'Прочее' }));
  }

  // дедуп: приоритет у записи с конкретной подкатегорией
  const byArticle = new Map();
  for (const item of collected) {
    const existing = byArticle.get(item.article);
    if (!existing) {
      byArticle.set(item.article, item);
      continue;
    }
    if (existing.subcategory === 'Прочее' && item.subcategory !== 'Прочее') {
      byArticle.set(item.article, item);
    }
  }

  const result = [...byArticle.values()];
  result.forEach((p) => {
    p.brand = brandOf(p.name);
    p.volume = volumeOf(p.name);
  });
  fs.writeFileSync(CACHE, JSON.stringify(result, null, 1));
  console.log(`  всего товаров: ${result.length}`);

  return result;
}

// ---------- описания ----------

async function fetchDescription(url) {
  const html = await fetchText(url);
  const $ = cheerio.load(html);

  const heading = $('h2').filter((_, el) => $(el).text().trim() === 'О товаре').first();
  if (!heading.length) return '';

  const container = heading.parent().clone();
  container.find('h2').first().remove();

  container.find('style, script, noscript').remove();

  const raw = stripEntities(
    (container.html() || '')
      .replace(/<br\s*\/?>/gi, '\n')
      .replace(/<\/p>/gi, '\n\n')
      .replace(/<\/li>/gi, '\n')
      .replace(/<li[^>]*>/gi, '• ')
      .replace(/<[^>]+>/g, ''),
  );

  const clean = raw
    .split(/\n\s*\n/)
    .filter((block) => !/var\(--|calc\(|pointer-events|scroll-mb|tailwind|@media|\{\s*\}|theme\(/i.test(block))
    .join('\n\n');

  return textToHtml(clean);
}

// ---------- исходная таблица ----------

function readLegacyTable() {
  if (!fs.existsSync(LEGACY_XLSX)) return [];

  const wb = XLSX.readFile(LEGACY_XLSX);
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: null });
  const out = [];

  for (const r of rows) {
    const name = String(r['название'] ?? '').trim();
    const article = cleanArticle(r['артикул']);
    const image = String(r['картинка'] ?? '').trim();
    if (!name || !article || !/^https?:/i.test(image)) continue;
    out.push({
      name,
      article,
      description: String(r['описание'] ?? ''),
      image,
      category: String(r['категория'] ?? '').trim(),
      subcategory: String(r['подкатегория'] ?? '').trim(),
    });
  }

  return out;
}

// ---------- картинки ----------

async function downloadImage(url, article) {
  const fileName = article.replace(/[\\/:*?"<>|]/g, '-') + '.webp';
  const dest = path.join(IMG_DIR, fileName);

  if (fs.existsSync(dest) && fs.statSync(dest).size > 500) return fileName;

  const large = url.replace('/medium/', '/large/');
  const res = await fetch(large, { headers: { 'User-Agent': UA } });
  if (!res.ok) throw new Error(`HTTP ${res.status}`);

  const buf = Buffer.from(await res.arrayBuffer());
  await sharp(buf).resize(700, 700, { fit: 'inside', withoutEnlargement: true }).webp({ quality: 82 }).toFile(dest);

  return fileName;
}

// ---------- база ----------

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
      volume TEXT NOT NULL DEFAULT '',
      brand TEXT NOT NULL DEFAULT '',
      price_client INTEGER,
      price_partner INTEGER,
      coupon TEXT NOT NULL DEFAULT '',
      price_market TEXT NOT NULL DEFAULT '',
      stock TEXT NOT NULL DEFAULT '',
      url TEXT NOT NULL DEFAULT '',
      featured INTEGER NOT NULL DEFAULT 0,
      sort_order INTEGER NOT NULL DEFAULT 0
    );
  `);
}

function makeShort(html, title) {
  const plain = stripTags(html);
  if (!plain) return '';
  const sentence = plain.split(/(?<=[.!?])\s/)[0];
  const base = sentence.length > 20 ? sentence : plain;
  return base.length > 140 ? base.slice(0, 137).trimEnd() + '…' : base;
}

// ---------- запуск ----------

async function main() {
  fs.mkdirSync(IMG_DIR, { recursive: true });

  const categories = await parseCategories();
  const market = await crawlAll(categories);

  console.log(`\nТоваров на маркетплейсе: ${market.length}`);

  // ---- описания ----
  let descCache = {};
  if (fs.existsSync(DESC_CACHE)) {
    try {
      descCache = JSON.parse(fs.readFileSync(DESC_CACHE, 'utf8'));
    } catch {
      descCache = {};
    }
  }

  console.log('\nОписания с карточек…');
  let fetchedDesc = 0;
  let cachedDesc = 0;
  let doneDesc = 0;

  const needDesc = market.filter((p) => {
    const cached = !REFRESH_DESC ? descCache[p.article] : '';
    if (cached) {
      p.description = cached;
      cachedDesc += 1;
      return false;
    }
    return true;
  });

  console.log(`  в кеше: ${cachedDesc}, надо скачать: ${needDesc.length}`);

  await pool(needDesc, 8, async (p) => {
    try {
      p.description = await fetchDescription(p.url);
    } catch {
      p.description = '';
    }
    descCache[p.article] = p.description;
    fetchedDesc += 1;
    doneDesc += 1;
    if (doneDesc % 100 === 0) {
      console.log(`  скачано ${doneDesc}/${needDesc.length}`);
      fs.writeFileSync(DESC_CACHE, JSON.stringify(descCache, null, 1));
    }
  });

  for (const p of market) if (p.description === undefined) p.description = '';

  fs.writeFileSync(DESC_CACHE, JSON.stringify(descCache, null, 1));
  console.log(`  скачано: ${fetchedDesc}, из кеша: ${cachedDesc}`);

  // ---- товары из исходной таблицы ----
  const legacy = readLegacyTable();
  const used = new Set(market.map((p) => p.article));
  let addedLegacy = 0;

  for (const item of legacy) {
    if (used.has(item.article)) continue;
    market.push({
      article: item.article,
      name: item.name,
      image: item.image,
      price: '',
      stock: 'legacy',
      url: '',
      parent: item.category || 'Ассортимент',
      subcategory: item.subcategory || 'Разное',
      description: textToHtml(item.description),
      legacy: true,
    });
    used.add(item.article);
    addedLegacy += 1;
  }
  console.log(`  добавлено из исходной таблицы: ${addedLegacy}`);

  // ---- порядок категорий ----
  const parentOrder = categories.groups.map((p) => p.name);
  const rank = (name) => {
    const i = parentOrder.indexOf(name);
    return i === -1 ? parentOrder.length : i;
  };

  market.sort((a, b) => {
    const p = rank(a.parent) - rank(b.parent);
    if (p) return p;
    if (a.parent === b.parent) {
      const s = String(a.subcategory).localeCompare(String(b.subcategory), 'ru');
      if (s) return s;
    }
    return a.name.localeCompare(b.name, 'ru');
  });

  // ---- картинки ----
  console.log('\nКартинки…');
  let downloaded = 0;
  let failed = 0;
  let doneImg = 0;

  await pool(market, 8, async (p) => {
    try {
      const file = await downloadImage(p.image, p.article);
      p.imageFile = '/images/catalog/' + file;
      downloaded += 1;
    } catch {
      p.imageFile = '/logo.svg';
      failed += 1;
    }
    doneImg += 1;
    if (doneImg % 200 === 0) console.log(`  ${doneImg}/${market.length}`);
  });

  console.log(`  обработано: ${downloaded}, ошибок: ${failed}`);

  // ---- Excel ----
  const rows = market.map((p, i) => ({
    '№': i + 1,
    'Категория': p.parent,
    'Подкатегория': p.subcategory,
    'Артикул': p.article,
    'Название': p.name,
    'Описание': stripTags(p.description),
    'Наличие': p.stock,
    'Цена маркетплейса': p.price,
    'Ссылка на товар': p.url,
    'Файл картинки': p.imageFile,
  }));

  const ws = XLSX.utils.json_to_sheet(rows);
  ws['!cols'] = [
    { wch: 4 }, { wch: 22 }, { wch: 22 }, { wch: 12 }, { wch: 55 }, { wch: 80 },
    { wch: 26 }, { wch: 16 }, { wch: 50 }, { wch: 30 },
  ];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Товары');
  XLSX.writeFile(wb, XLSX_OUT);
  console.log(`\nExcel: ${rows.length} строк`);

  // ---- база ----
  const db = new Database(DB_PATH);
  createSchema(db);

  const insert = db.prepare(`
    INSERT INTO products
      (article, title, description, short, category, subcategory, images,
       volume, brand, price_client, price_partner, coupon, price_market, stock, url, featured, sort_order)
    VALUES
      (@article, @title, @description, @short, @category, @subcategory, @images,
       @volume, @brand, NULL, NULL, '', @price_market, @stock, @url, 0, @sort_order)
  `);

  db.transaction((list) => {
    list.forEach((p, i) => {
      insert.run({
        article: p.article,
        title: p.name,
        description: p.description || '',
        short: makeShort(p.description, p.name),
        category: p.parent,
        subcategory: p.subcategory,
        images: JSON.stringify([p.imageFile || '/logo.svg']),
        volume: p.volume || '',
        brand: p.brand || '',
        price_market: p.price || '',
        stock: p.stock || '',
        url: p.url || '',
        sort_order: i,
      });
    });
  })(market);

  const total = db.prepare('SELECT COUNT(*) AS n FROM products').get().n;
  const groups = db
    .prepare('SELECT category, COUNT(*) AS n FROM products GROUP BY category ORDER BY MIN(sort_order)')
    .all();
  db.close();

  // ---- чистка неиспользуемых картинок ----
  const usedFiles = new Set(market.map((p) => path.basename(p.imageFile || '')));
  let removed = 0;
  for (const file of fs.readdirSync(IMG_DIR)) {
    if (!usedFiles.has(file)) {
      fs.unlinkSync(path.join(IMG_DIR, file));
      removed += 1;
    }
  }

  console.log(`\nВ базе: ${total}`);
  for (const g of groups) console.log(`  ${g.category}: ${g.n}`);
  console.log(`Удалено неиспользуемых картинок: ${removed}`);
}

main().catch((err) => {
  console.error('Ошибка синхронизации:', err);
  process.exit(1);
});
