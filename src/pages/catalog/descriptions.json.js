import Database from 'better-sqlite3';
import path from 'node:path';

export async function GET() {
  const db = new Database(path.join(process.cwd(), 'data', 'catalog.db'), { readonly: true });

  const rows = db.prepare('SELECT id, description FROM products ORDER BY id').all();

  db.close();

  const payload = JSON.stringify(Object.fromEntries(rows.map((r) => [String(r.id), r.description])));

  return new Response(payload, {
    headers: { 'Content-Type': 'application/json; charset=utf-8' },
  });
}
