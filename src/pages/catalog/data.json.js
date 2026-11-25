import Database from 'better-sqlite3';
import path from 'node:path';

// Компактный формат: массив массивов вместо массива объектов.
// Поля в порядке KEYS — так JSON получается почти вдвое легче.
export const KEYS = ['id', 'article', 'title', 'short', 'cat', 'sub', 'img', 'brand', 'vol'];

const shortText = (value, max = 110) => {
  const text = String(value || '').trim();
  if (text.length <= max) return text;
  return text.slice(0, max).replace(/\s+\S*$/, '') + '…';
};

export async function GET() {
  const db = new Database(path.join(process.cwd(), 'data', 'catalog.db'), { readonly: true });

  const rows = db
    .prepare(
      `SELECT id, article, title, short, category, subcategory, images, brand, volume
       FROM products ORDER BY sort_order, id`,
    )
    .all();

  db.close();

  const products = rows.map((r) => [
    r.id,
    r.article,
    r.title,
    shortText(r.short),
    r.category,
    r.subcategory,
    JSON.parse(r.images)[0] || '/logo.svg',
    r.brand || '',
    r.volume || '',
  ]);

  const categories = [];
  const seen = new Set();
  for (const p of products) {
    const cat = p[4];
    if (seen.has(cat)) continue;
    seen.add(cat);
    const subs = [...new Set(products.filter((x) => x[4] === cat).map((x) => x[5]))];
    categories.push({ name: cat, subs });
  }

  const brands = [...new Set(products.map((p) => p[7]).filter(Boolean))].sort((a, b) =>
    a.localeCompare(b, 'ru'),
  );

  const payload = JSON.stringify({ keys: KEYS, products, categories, brands });

  return new Response(payload, {
    headers: { 'Content-Type': 'application/json; charset=utf-8' },
  });
}
