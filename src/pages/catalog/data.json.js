import Database from 'better-sqlite3';
import path from 'node:path';

export async function GET() {
  const db = new Database(path.join(process.cwd(), 'data', 'catalog.db'), { readonly: true });

  const rows = db
    .prepare(
      `SELECT id, article, title, short, category, subcategory, images, brand, volume, price_market
       FROM products ORDER BY sort_order, id`,
    )
    .all();

  db.close();

  const products = rows.map((r) => ({
    id: r.id,
    article: r.article,
    title: r.title,
    short: r.short || '',
    cat: r.category,
    sub: r.subcategory,
    img: JSON.parse(r.images)[0] || '/logo.svg',
    brand: r.brand || '',
    vol: r.volume || '',
  }));

  const categories = [];
  const seen = new Set();
  for (const p of products) {
    if (seen.has(p.cat)) continue;
    seen.add(p.cat);
    categories.push({
      name: p.cat,
      subs: [...new Set(products.filter((x) => x.cat === p.cat).map((x) => x.sub))],
    });
  }

  const brands = [...new Set(products.map((p) => p.brand).filter(Boolean))].sort((a, b) =>
    a.localeCompare(b, 'ru'),
  );

  const payload = JSON.stringify({ products, categories, brands });

  return new Response(payload, {
    headers: { 'Content-Type': 'application/json; charset=utf-8' },
  });
}
