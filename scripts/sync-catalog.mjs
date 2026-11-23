// Синхронизация каталога: партнёрская страница + маркетплейс → база + Excel.
//
//   node scripts/sync-catalog.mjs              — полный цикл (обходит маркетплейс)
//   node scripts/sync-catalog.mjs --use-cache  — не обходить маркетплейс, взять прошлый обход
//   node scripts/sync-catalog.mjs --debug      — подробный лог отбракованных товаров
//
// Источники:
//   страница партнёра   — название, артикул (из имени картинки), цены, описание, категории
//   greenleaf.online    — наличие, цена, ссылка, картинка
// Результат:
//   data/greenleaf-full.xlsx      — все товары с партнёрской страницы (для просмотра)
//   data/catalog.db               — только товары «Есть у подключённых партнёров»
//   public/images/catalog/*.webp  — картинки с маркетплейса

import fs from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';
import * as cheerio from 'cheerio';
import sharp from 'sharp';
import Database from 'better-sqlite3';

const require = createRequire(import.meta.url);
const XLSX = require('xlsx');

const ROOT = process.cwd();
const PARTNER_URL =
  'https://xn--80aaajcmapeahqleykmtil0bk6a7d8m.xn--p1ai/' +
  encodeURIComponent('наталья-марьясова-г-геленджик') +
  '/';
const MARKET_BASE = 'https://greenleaf.online';
const CACHE = path.join(ROOT, 'data', 'marketplace-cache.json');
const XLSX_OUT = path.join(ROOT, 'data', 'greenleaf-full.xlsx');
const DB_PATH = path.join(ROOT, 'data', 'catalog.db');
const IMG_DIR = path.join(ROOT, 'public', 'images', 'catalog');
const DESC_CACHE = path.join(ROOT, 'data', 'descriptions-cache.json');

const USE_CACHE = process.argv.includes('--use-cache');
const DEBUG = process.argv.includes('--debug');
const DRY = process.argv.includes('--dry');
const REFRESH_IMAGES = process.argv.includes('--refresh-images');
const REFRESH_DESC = process.argv.includes('--refresh-desc');

const UA =
  'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36';

const AVAILABLE = 'Есть у подключённых партнёров';

// категория макияжа на сайте встречается с опечаткой
const CATEGORY_FIX = {
  'макияж и парюм': 'Макияж и парфюм',
  'макияж и парфюм': 'Макияж и парфюм',
};

const SUBCATEGORY_ORDER = [
  'Очищение и демакияж',
  'Комплексный уход',
  'Основной уход',
  'Маски',
  'Базы и ровный тон',
  'Декоративная косметика',
  'Коллекция ароматов',
  'Шампуни и бальзамы',
  'Маски, масла, укладка',
  'Гели для душа',
  'Уходовые средства',
  'Мыло для лица и тела',
  'Зубные пасты',
  'Ополаскиватели и щетки',
  'Женская гигиена',
  'Предметы личной гигиены',
  'Детская гигиена',
  'БАДы и витамины',
  'Полезные напитки',
  'Чай, кофе',
  'Здоровое тело',
  'Средства для стирки',
  'Безопасная уборка дома и кухни',
  'Разное',
];

const CATEGORY_ORDER = [
  'Уход за кожей лица',
  'Макияж и парфюм',
  'Уход для волос',
  'Уход за кожей тела',
  'Уход за полостью рта',
  'Гигиена',
  'Товары для детей',
  'Здоровье',
  'Товары для дома и быта',
  'Ассортимент',
];

const normCase = (s) => {
  const t = String(s || '').replace(/\s+/g, ' ').trim();
  if (!t) return '';
  return t[0].toUpperCase() + t.slice(1).toLowerCase();
};

const fixCategory = (name) => {
  const n = normCase(name);
  return CATEGORY_FIX[n.toLowerCase()] || n;
};

const fixSubcategory = (name) => {
  const n = normCase(name);
  return n === 'Бады и витамины' ? 'БАДы и витамины' : n;
};

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

const STOP_WORDS = new Set(['для', 'лица', 'кожи', 'уход', 'серия', 'средство', 'товар']);

const normTokens = (value) => {
  const words = String(value || '')
    .toLowerCase()
    .replace(/ё/g, 'е')
    .replace(/[^a-zа-я0-9]+/g, ' ')
    .split(' ')
    .filter((w) => w.length > 2 && !STOP_WORDS.has(w));
  return new Set(words);
};

// бренд из начала названия маркетплейса
const BRAND_WORDS = ['SEALUXE', 'iLiFE', 'ILIFE', 'CARICH', 'NILRICH', 'KARDLI', 'SONPHESKIN', 'GREENLEAF', 'YUE ZHONGGUI', 'MOON FLOWER'];
const marketBrand = (name) => {
  const lower = String(name || '').toLowerCase();
  const hit = BRAND_WORDS.find((b) => lower.startsWith(b.toLowerCase()));
  return hit ? hit.toLowerCase().replace('ilife', 'ilife') : '';
};

// лучший товар маркетплейса: сначала по артикулу с картинки, потом по названию
const findMatch = (product, marketList, marketByArticle) => {
  const A = normTokens(product.name);
  if (!A.size) return null;

  const brand = String(product.brand || '').toLowerCase();

  const scoreOf = (m) => {
    let shared = 0;
    for (const token of A) if (m.tokens.has(token)) shared += 1;
    return { shared, score: shared / Math.max(1, Math.min(A.size, m.tokens.size)) };
  };

  const brandOk = (m) => {
    const mb = marketBrand(m.name);
    if (!brand || !mb) return true;
    return mb === brand;
  };

  // 1. артикул, вычисленный из картинки партнёрской страницы
  if (product.article) {
    const direct = marketByArticle[product.article];
    if (direct) {
      const { shared, score } = scoreOf({ tokens: normTokens(direct.name) });
      if (shared >= 2 && score >= 0.4 && brandOk({ name: direct.name })) {
        return { ...direct, score: Number(score.toFixed(2)) };
      }
    }
  }

  // 2. поиск по названию, со строгим порогом и проверкой бренда
  let best = null;
  let bestScore = 0;
  let bestShared = 0;

  for (const m of marketList) {
    if (!brandOk(m)) continue;
    const { shared, score } = scoreOf(m);
    if (shared < 2) continue;
    if (score < 0.75) continue;
    if (shared < 3 && score < 0.85) continue;

    if (score > bestScore || (score === bestScore && shared > bestShared)) {
      best = m;
      bestScore = score;
      bestShared = shared;
    }
  }

  return best ? { ...best, score: Number(bestScore.toFixed(2)) } : null;
};

async function fetchText(url, tries = 3) {
  for (let i = 0; i < tries; i += 1) {
    try {
      const res = await fetch(url, { headers: { 'User-Agent': UA } });
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      return await res.text();
    } catch (err) {
      if (i === tries - 1) throw err;
      await sleep(1200);
    }
  }
  return '';
}

// ---------- 1. партнёрская страница ----------

function parseVolumeAndBrand(html) {
  const text = String(html || '').replace(/<br\s*\/?>/gi, '\n');
  const lines = text
    .split('\n')
    .map((l) => l.replace(/<[^>]+>/g, '').replace(/&nbsp;/g, ' ').trim())
    .filter(Boolean);

  if (!lines.length) return { volume: '', brand: '' };

  const first = lines[0];
  const volMatch = first.match(/^(?:ОБЪЕМ|ОБЪЁМ|VOLUME)?\s*([\d.,]+\s*(?:мл|г|гр|шт|ml|g))\b/i);
  const volume = volMatch ? volMatch[1].replace(/\s+/g, ' ') : '';
  const rest = volMatch ? first.slice(volMatch[0].length).trim() : first;

  const brand = lines[1] || rest || '';
  return { volume, brand: brand.replace(/&nbsp;/g, ' ').trim() };
}

function splitMeta(text) {
  let rest = String(text || '').replace(/\u00a0/g, ' ').replace(/[ \t]+/g, ' ').trim();
  rest = rest.replace(/^(?:ОБЪЕМ|ОБЪЁМ|VOLUME)\s*:?\s*/i, '');

  const volMatch = rest.match(
    /^(\d[\d\s.,]*\s*(?:мл|гр|г|кг|шт|штук|пара|пар|пакет[а-яё]*|таблет[а-яё]*|саше|уп)(?![а-яё])(?:\s*[хx×*]\s*\d+\s*[А-Яа-яё]{0,12})?)/i,
  );
  const volume = volMatch ? volMatch[1].replace(/\s+/g, ' ').trim() : '';
  if (volMatch) rest = rest.slice(volMatch[0].length).trim();

  const brandMatch = rest.match(/(SEALUXE|iLiFE|ILIFE|CARICH|NILRICH|KARDLI|SONPHESKIN|GREENLEAF|YUE ZHONGGUI|MOON FLOWER)/i);
  const BRAND_CANON = {
    sealuxe: 'SEALUXE',
    ilife: 'iLiFE',
    carich: 'CARICH',
    nilrich: 'NILRICH',
    kardli: 'KARDLI',
    sonpheskin: 'SONPHESKIN',
    greenleaf: 'GREENLEAF',
    'yue zhonggui': 'YUE ZHONGGUI',
    'moon flower': 'MOON FLOWER',
  };
  const brand = brandMatch ? (BRAND_CANON[brandMatch[1].toLowerCase()] || brandMatch[1]) : '';
  if (brandMatch) {
    rest = rest.slice(brandMatch.index + brandMatch[0].length);
    rest = rest.replace(/^\s*\/\s*(SEALUXE|iLiFE|ILIFE|CARICH|NILRICH|KARDLI|SONPHESKIN)\b\s*/i, '');
  }

  return { volume, brand, rest: rest.replace(/^[\s\-–—:]+/, '').trim() };
}

function textToHtml(text) {
  const escape = (s) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const cleaned = String(text || '').replace(/([.!?])(?=[А-ЯЁA-Z«"(])/g, '$1 ');
  const blocks = cleaned
    .split(/\n\s*\n/)
    .map((b) => b.trim())
    .filter(Boolean);

  return blocks
    .map((block) => {
      const lines = block.split('\n').map((l) => l.trim()).filter(Boolean);
      if (lines.length && lines.every((l) => /^[➤•\-–—]\s*/.test(l))) {
        const items = lines.map((l) => `<li>${escape(l.replace(/^[➤•\-–—]\s*/, ''))}</li>`).join('');
        return `<ul>${items}</ul>`;
      }
      return `<p>${escape(lines.join(' '))}</p>`;
    })
    .join('');
}

async function parsePartnerPage() {
  console.log('Читаю партнёрскую страницу…');
  const html = await fetchText(PARTNER_URL);
  console.log(`  получено ${(html.length / 1048576).toFixed(2)} МБ`);
  const $ = cheerio.load(html);

  const nodes = $('h3, h4, h5, h6, p, img').toArray();
  const products = [];

  // названия категорий (h3) — чтобы не путать их с подкатегориями
  const categoryNames = new Set(
    nodes
      .filter((el) => el.tagName.toLowerCase() === 'h3')
      .map((el) => $(el).text().replace(/\s+/g, ' ').trim().toUpperCase()),
  );

  let category = '';
  let subcategory = '';
  let lastImg = '';
  let current = null;

  for (const el of nodes) {
    const tag = el.tagName.toLowerCase();

    if (tag === 'img') {
      lastImg = $(el).attr('src') || '';
      continue;
    }

    const text = $(el).text().replace(/\s+/g, ' ').trim();
    if (!text) continue;
    const upper = text.toUpperCase();

    if (tag === 'h3') {
      if (upper !== category) {
        category = upper;
        subcategory = '';
      }
      current = null;
      continue;
    }

    if ((tag === 'h4' || tag === 'h5') && upper !== 'ОТКРЫТЬ ОПИСАНИЕ') {
      if (!categoryNames.has(upper)) {
        subcategory = upper;
        current = null;
      }
      continue;
    }

    if (tag === 'h6' && /КЛИЕНТ/.test(text) && /ПАРТНЕР/.test(text)) {
      const m = text.match(/^(.*?)₽\s*КЛИЕНТ\s*([\d\s]+)₽\s*ПАРТНЕР\s*([\d\s]*)(.*)$/);
      const name = (m ? m[1] : text).replace(/\s+/g, ' ').trim();
      const client = m ? Number(m[2].replace(/\s/g, '')) : null;
      const partner = m && m[3].trim() ? Number(m[3].replace(/\s/g, '')) : null;
      const tail = m ? m[4].trim() : '';

      const fileName = decodeURIComponent(lastImg.split('/').pop() || '');
      const artMatch = fileName.toUpperCase().match(/([A-Z]{2,4}\d{3,4}[A-Z]?)[-_]/);

      current = {
        name,
        article: artMatch ? artMatch[1] : '',
        client,
        partner,
        coupon: /купон/i.test(tail) ? tail.replace(/\s+/g, ' ').trim() : '',
        image: lastImg,
        category: fixCategory(category),
        subcategory: fixSubcategory(subcategory),
        volume: '',
        brand: '',
        description: '',
        paragraphs: [],
      };
      products.push(current);
      continue;
    }

    if (tag === 'p' && current) {
      const cls = $(el).attr('class') || '';
      if (!cls.includes('stk-block-text__text')) continue;
      current.paragraphs.push($(el).html() || '');
    }
  }

  for (const p of products) {
    const text = p.paragraphs
      .map((h) =>
        h
          .replace(/<br\s*\/?>/gi, '\n')
          .replace(/<\/p>/gi, '\n\n')
          .replace(/<\/li>/gi, '\n')
          .replace(/<li[^>]*>/gi, '• ')
          .replace(/<\/(div|h[1-6]|ul|ol)>/gi, '\n'),
      )
      .join('\n\n')
      .replace(/<[^>]+>/g, '')
      .replace(/&nbsp;/g, ' ')
      .replace(/&amp;/g, '&')
      .replace(/&quot;/g, '"')
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>');
    const meta = splitMeta(text);
    p.volume = meta.volume;
    p.brand = meta.brand;
    p.description = textToHtml(meta.rest);
    delete p.paragraphs;
  }

  console.log(`  товаров: ${products.length}`);
  return products;
}

// ---------- 2. маркетплейс ----------

function parseMarketCards(html) {
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

async function crawlMarketplace() {
  console.log('Обхожу маркетплейс…');
  const sitemap = await fetchText(`${MARKET_BASE}/sitemap.xml`);
  const categories = [...new Set([...sitemap.matchAll(/greenleaf\.online\?category=([a-z0-9\-]+)/g)].map((m) => m[1]))];
  console.log(`  категорий в sitemap: ${categories.length}`);

  const products = new Map();
  let pages = 0;

  for (const [i, cat] of categories.entries()) {
    for (let page = 1; page <= 60; page += 1) {
      const html = await fetchText(`${MARKET_BASE}/?category=${cat}&page=${page}`);
      const cards = parseMarketCards(html);
      pages += 1;
      if (!cards.length) break;

      let added = 0;
      for (const c of cards) {
        if (!products.has(c.article)) {
          products.set(c.article, c);
          added += 1;
        }
      }
      if (!added) break;
      await sleep(200);
    }
    if ((i + 1) % 10 === 0) {
      console.log(`  ${i + 1}/${categories.length} категорий, товаров ${products.size}`);
    }
  }

  console.log(`  страниц: ${pages}, уникальных артикулов: ${products.size}`);
  return Object.fromEntries(products);
}

// ---------- 3. сборка ----------

async function downloadImage(url, article) {
  const fileName = article.replace(/[\\/:*?"<>|]/g, '-') + '.webp';
  const dest = path.join(IMG_DIR, fileName);

  if (!REFRESH_IMAGES && fs.existsSync(dest) && fs.statSync(dest).size > 500) {
    return fileName;
  }

  const large = url.replace('/medium/', '/large/');
  const res = await fetch(large, { headers: { 'User-Agent': UA } });
  if (!res.ok) throw new Error(`HTTP ${res.status} для ${large}`);

  const buf = Buffer.from(await res.arrayBuffer());
  await sharp(buf)
    .resize(700, 700, { fit: 'inside', withoutEnlargement: true })
    .webp({ quality: 82 })
    .toFile(dest);

  return fileName;
}

async function fetchMarketDescription(url) {
  const html = await fetchText(url);
  const $ = cheerio.load(html);

  const heading = $('h2').filter((_, el) => $(el).text().trim() === 'О товаре').first();
  if (!heading.length) return '';

  const container = heading.parent().clone();
  container.find('h2').first().remove();

  const raw = (container.html() || '')
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/p>/gi, '\n\n')
    .replace(/<\/li>/gi, '\n')
    .replace(/<li[^>]*>/gi, '• ')
    .replace(/<[^>]+>/g, '')
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&quot;/g, '"')
    .replace(/&laquo;/g, '«')
    .replace(/&raquo;/g, '»')
    .replace(/&mdash;/g, '—')
    .replace(/&ndash;/g, '–')
    .replace(/&[a-z]+;/g, ' ')
    .replace(/[ \t]+/g, ' ');

  return textToHtml(raw);
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

const stripTags = (s) => String(s || '').replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim();

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

  const partnerProducts = await parsePartnerPage();

  if (DRY) {
    console.log('--- примеры разбора ---');
    for (const p of partnerProducts.slice(0, 5)) {
      console.log(JSON.stringify({ article: p.article, volume: p.volume, brand: p.brand, cat: p.category, sub: p.subcategory, desc: p.description.slice(0, 200) }, null, 1));
    }
    console.log('без объёма:', partnerProducts.filter((x) => !x.volume).length, 'из', partnerProducts.length);
    console.log('без описания:', partnerProducts.filter((x) => !x.description).length);
    console.log('без подкатегории:', partnerProducts.filter((x) => !x.subcategory).length);
    return;
  }

  let market;
  if (USE_CACHE && fs.existsSync(CACHE)) {
    market = JSON.parse(fs.readFileSync(CACHE, 'utf8'));
    console.log(`Маркетплейс из кеша: ${Object.keys(market).length} артикулов`);
  } else {
    market = await crawlMarketplace();
    fs.writeFileSync(CACHE, JSON.stringify(market, null, 1));
    console.log(`Кеш маркетплейса: ${CACHE}`);
  }

  // ---- отбор: сопоставление по названию ----
  const marketList = Object.values(market).map((m) => ({ ...m, tokens: normTokens(m.name) }));
  const marketByArticle = market;

  const seen = new Set();
  const selected = [];
  const rejected = { notFound: [], notAvailable: [], duplicate: [] };
  const matches = new Map();

  for (const p of partnerProducts) {
    const m = findMatch(p, marketList, marketByArticle);
    if (!m) {
      rejected.notFound.push(p);
      continue;
    }

    p.market = m;
    p.marketArticle = m.article;
    matches.set(p, m);

    if (seen.has(m.article)) {
      rejected.duplicate.push(p);
      continue;
    }
    seen.add(m.article);

    if (m.stock !== AVAILABLE) {
      rejected.notAvailable.push({ ...p, stock: m.stock });
      continue;
    }
    selected.push(p);
  }

  console.log('');
  console.log('=== ОТБОР ===');
  console.log(`  всего на странице: ${partnerProducts.length}`);
  console.log(`  не найдено в маркетплейсе: ${rejected.notFound.length}`);
  console.log(`  не «у партнёров»: ${rejected.notAvailable.length}`);
  console.log(`  дубли: ${rejected.duplicate.length}`);
  console.log(`  В КАТАЛОГ: ${selected.length}`);

  if (DEBUG) {
    console.log('\nне найдено (примеры):', rejected.notFound.slice(0, 15).map((p) => p.name.slice(0, 45)));
  }

  // ---- сортировка ----
  const catRank = (c) => {
    const i = CATEGORY_ORDER.indexOf(c);
    return i === -1 ? CATEGORY_ORDER.length : i;
  };
  const subRank = (s) => {
    const i = SUBCATEGORY_ORDER.indexOf(s);
    return i === -1 ? SUBCATEGORY_ORDER.length : i;
  };
  selected.sort((a, b) => {
    const c = catRank(a.category) - catRank(b.category);
    if (c) return c;
    const s = subRank(a.subcategory) - subRank(b.subcategory);
    if (s) return s;
    return a.name.localeCompare(b.name, 'ru');
  });

  // ---- описания с маркетплейса ----
  let descCache = {};
  if (fs.existsSync(DESC_CACHE)) {
    try {
      descCache = JSON.parse(fs.readFileSync(DESC_CACHE, 'utf8'));
    } catch {
      descCache = {};
    }
  }

  console.log('\nТяну описания с маркетплейса…');
  let fetched = 0;
  let fromCache = 0;
  let emptyDesc = 0;

  for (const p of selected) {
    const key = p.marketArticle;
    let desc = !REFRESH_DESC ? descCache[key] : '';

    if (desc === undefined || desc === null) desc = '';

    if (!desc) {
      try {
        desc = await fetchMarketDescription(p.market.url);
        fetched += 1;
        if (!desc) emptyDesc += 1;
      } catch (err) {
        desc = '';
        emptyDesc += 1;
      }
      descCache[key] = desc;
      await sleep(120);
      if (fetched % 25 === 0) console.log(`  ${fetched}/${selected.length}`);
    } else {
      fromCache += 1;
    }

    p.marketDescription = desc;
  }

  fs.writeFileSync(DESC_CACHE, JSON.stringify(descCache, null, 1));
  console.log(`  скачано: ${fetched}, из кеша: ${fromCache}, без описания: ${emptyDesc}`);

  // ---- картинки: имя файла = артикул маркетплейса ----
  console.log('\nКачаю картинки…');
  let downloaded = 0;
  let failed = 0;
  for (const p of selected) {
    try {
      const file = await downloadImage(p.market.image, p.marketArticle);
      p.imageFile = '/images/catalog/' + file;
      downloaded += 1;
      if (downloaded % 25 === 0) console.log(`  ${downloaded}/${selected.length}`);
    } catch (err) {
      p.imageFile = '/logo.svg';
      failed += 1;
      console.error(`  ! ${p.marketArticle}: ${err.message}`);
    }
    await sleep(90);
  }
  console.log(`  скачано: ${downloaded}, ошибок: ${failed}`);

  // ---- Excel: все товары партнёрской страницы ----
  const selectedSet = new Set(selected);

  const rows = partnerProducts.map((p, i) => {
    const m = p.market || null;
    const inCatalog = selectedSet.has(p);
    const file = m && inCatalog ? `/images/catalog/${m.article.replace(/[\\/:*?"<>|]/g, '-')}.webp` : '';

    return {
      '№': i + 1,
      'Категория': p.category,
      'Подкатегория': p.subcategory,
      'Артикул (сайт)': p.article || '',
      'Артикул маркетплейса': m ? m.article : '',
      'Совпадение': m ? m.score : '',
      'Название': p.name,
      'Название на маркетплейсе': m ? m.name : '',
      'Объём': p.volume,
      'Бренд': p.brand,
      'Описание': stripTags(p.market && p.marketDescription ? p.marketDescription : p.description),
      'Источник описания': p.market && p.marketDescription ? 'маркетплейс' : p.description ? 'сайт партнёра' : 'нет',
      'Цена клиента': p.client ?? '',
      'Цена партнёра': p.partner ?? '',
      'Купон': p.coupon,
      'Цена маркетплейса': m ? m.price : '',
      'Наличие': m ? m.stock : 'Не найден в маркетплейсе',
      'В каталоге': inCatalog ? 'да' : 'нет',
      'Ссылка на товар': m ? m.url : '',
      'Файл картинки': file,
    };
  });

  const ws = XLSX.utils.json_to_sheet(rows);
  ws['!cols'] = [
    { wch: 4 }, { wch: 22 }, { wch: 22 }, { wch: 12 }, { wch: 16 }, { wch: 9 },
    { wch: 45 }, { wch: 45 }, { wch: 10 }, { wch: 12 }, { wch: 80 }, { wch: 18 },
    { wch: 9 }, { wch: 10 }, { wch: 12 }, { wch: 16 }, { wch: 22 },
    { wch: 8 }, { wch: 45 }, { wch: 30 },
  ];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Товары');
  XLSX.writeFile(wb, XLSX_OUT);
  console.log(`\nExcel: ${XLSX_OUT} (${rows.length} строк)`);

  // ---- база ----
  const db = new Database(DB_PATH);
  createSchema(db);

  const insert = db.prepare(`
    INSERT INTO products
      (article, title, description, short, category, subcategory, images,
       volume, brand, price_client, price_partner, coupon, price_market, stock, url, featured, sort_order)
    VALUES
      (@article, @title, @description, @short, @category, @subcategory, @images,
       @volume, @brand, @price_client, @price_partner, @coupon, @price_market, @stock, @url, 0, @sort_order)
  `);

  db.transaction((list) => {
    list.forEach((p, i) => {
      insert.run({
        article: p.marketArticle,
        title: p.name,
        description: p.marketDescription || p.description,
        short: makeShort(p.marketDescription || p.description, p.name),
        category: p.category,
        subcategory: p.subcategory,
        images: JSON.stringify([p.imageFile]),
        volume: p.volume,
        brand: p.brand,
        price_client: p.client,
        price_partner: p.partner,
        coupon: p.coupon,
        price_market: p.market.price,
        stock: p.market.stock,
        url: p.market.url,
        sort_order: i,
      });
    });
  })(selected);

  const total = db.prepare('SELECT COUNT(*) AS n FROM products').get().n;
  const groups = db
    .prepare('SELECT category, subcategory, COUNT(*) AS n FROM products GROUP BY category, subcategory ORDER BY MIN(sort_order)')
    .all();

  db.close();

  // ---- чистка картинок, на которые больше нет ссылок ----
  const used = new Set(selected.map((p) => path.basename(p.imageFile || '')));
  let removed = 0;
  for (const file of fs.readdirSync(IMG_DIR)) {
    if (!used.has(file)) {
      fs.unlinkSync(path.join(IMG_DIR, file));
      removed += 1;
    }
  }

  console.log(`\nВ базе: ${total}`);
  for (const g of groups) console.log(`  ${g.category} / ${g.subcategory}: ${g.n}`);
  console.log(`\nУдалено неиспользуемых картинок: ${removed}`);
}

main().catch((err) => {
  console.error('Ошибка синхронизации:', err);
  process.exit(1);
});
