#!/usr/bin/env node
// Создание страницы нового партнёра одной командой.
//
// Пример:
//   node scripts/new-partner.mjs ^
//     --name "Иванова Анна Петровна" ^
//     --photo "C:\Users\Admin\Downloads\anna.jpg" ^
//     --short "Анна" ^
//     --phone "+79160000000" ^
//     --tg anna_greenleaf ^
//     --vk "https://vk.ru/anna" ^
//     --max "https://max.ru/u/xxxx" ^
//     --email "anna@mail.ru"
//
// Что делает:
//   1. сжимает фото в public/<slug>.webp (квадрат 800×800) и public/og/<slug>.jpg (1200×630)
//   2. добавляет запись в src/data/partners.js
//   3. пишет, что делать дальше (npm run build / npm run deploy)

import fs from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';

const require = createRequire(new URL('../package.json', import.meta.url));
const sharp = require('sharp');

const ROOT = path.resolve(new URL('..', import.meta.url).pathname.replace(/^\/([A-Za-z]:)/, '$1'));
const PARTNERS_FILE = path.join(ROOT, 'src', 'data', 'partners.js');

// ---------- аргументы ----------
const args = {};
for (let i = 2; i < process.argv.length; i += 2) {
  const key = process.argv[i].replace(/^--/, '');
  args[key] = process.argv[i + 1];
}

const required = ['name', 'photo'];
for (const r of required) {
  if (!args[r]) {
    console.error(`Не хватает аргумента --${r}`);
    process.exit(1);
  }
}

// ---------- транслитерация ----------
const TRANSLIT = {
  а: 'a', б: 'b', в: 'v', г: 'g', д: 'd', е: 'e', ё: 'e', ж: 'zh', з: 'z', и: 'i',
  й: 'y', к: 'k', л: 'l', м: 'm', н: 'n', о: 'o', п: 'p', р: 'r', с: 's', т: 't',
  у: 'u', ф: 'f', х: 'h', ц: 'ts', ч: 'ch', ш: 'sh', щ: 'sch', ъ: '', ы: 'y', ь: '',
  э: 'e', ю: 'yu', я: 'ya',
};

const translit = (str) =>
  String(str)
    .toLowerCase()
    .split('')
    .map((ch) => TRANSLIT[ch] ?? ch)
    .join('')
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '');

const slug = args.slug || translit(args.name);
const short = args.short || args.name.split(' ')[1] || args.name.split(' ')[0];
const role = args.role || 'VIP-партнёр';
const city = args.city || 'Геленджик';

// ---------- проверки ----------
if (!fs.existsSync(PARTNERS_FILE)) {
  console.error('Не найден src/data/partners.js — запускай скрипт из папки проекта');
  process.exit(1);
}

const source = fs.readFileSync(PARTNERS_FILE, 'utf8');
if (source.includes(`'${slug}':`)) {
  console.error(`Партнёр с ключом "${slug}" уже есть в partners.js`);
  process.exit(1);
}

if (!fs.existsSync(args.photo)) {
  console.error(`Не найден файл фото: ${args.photo}`);
  process.exit(1);
}

// ---------- фото ----------
const publicDir = path.join(ROOT, 'public');
const ogDir = path.join(publicDir, 'og');
fs.mkdirSync(ogDir, { recursive: true });

const photoOut = path.join(publicDir, `${slug}.webp`);
const ogOut = path.join(ogDir, `${slug}.jpg`);

await sharp(args.photo).resize(800, 800, { fit: 'cover', position: 'attention' }).webp({ quality: 84 }).toFile(photoOut);
await sharp(args.photo).resize(1200, 630, { fit: 'cover', position: 'attention' }).jpeg({ quality: 82 }).toFile(ogOut);

console.log(`фото: public/${slug}.webp (${Math.round(fs.statSync(photoOut).size / 1024)} КБ)`);
console.log(`og:   public/og/${slug}.jpg (${Math.round(fs.statSync(ogOut).size / 1024)} КБ)`);

// ---------- запись в partners.js ----------
const phones = [];
if (args.phone) phones.push(...String(args.phone).split(',').map((p) => p.trim()).filter(Boolean));

const contactsLines = [];
contactsLines.push(`      phones: [${phones.map((p) => `'${p}'`).join(', ')}],`);
if (args.max) contactsLines.push(`      max: '${args.max}',`);
if (args.vk) contactsLines.push(`      vk: '${args.vk}',`);
if (args.telegram || args.tg) contactsLines.push(`      telegram: '${args.telegram || args.tg}',`);
if (args.email) contactsLines.push(`      email: '${args.email}',`);

const entry = `
  '${slug}': {
    slug: '${slug}',
    fullName: '${args.name}',
    shortName: '${short}',
    role: '${role}',
    city: '${city}',
    photo: '/${slug}.webp',
    ogImage: '/og/${slug}.jpg',
    title: '${args.name.split(' ').slice(0, 2).join(' ')} — ${role} Greenleaf в ${city}е',
    description:
      '${args.name} — ${role} Greenleaf в ${city}е. Премиальный уход, витамины, товары для дома и одежды. Помощь с выбором и оформлением заказа.',
    contacts: {
${contactsLines.join('\n')}
    },
  },
`;

const objStart = source.indexOf('export const partners = {');
const insertAt = source.indexOf('\n};', objStart) + 1;
if (objStart === -1 || insertAt === 0) {
  console.error('Не смог найти объект partners в partners.js');
  process.exit(1);
}
const updated = source.slice(0, insertAt) + entry.trimEnd() + '\n' + source.slice(insertAt);
fs.writeFileSync(PARTNERS_FILE, updated, 'utf8');

console.log(`\nдобавлено в partners.js: ${slug}`);
console.log(`страница будет тут: /partners/${slug}/`);
console.log('\nДальше:');
console.log('  npm run build');
console.log('  npm run deploy');
