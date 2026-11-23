import Database from 'better-sqlite3';
import fs from 'fs';
import path from 'path';

const dbDir = path.join(process.cwd(), 'data');
if (!fs.existsSync(dbDir)) {
  fs.mkdirSync(dbDir);
}

const db = new Database(path.join(dbDir, 'catalog.db'));

db.exec(`
  DROP TABLE IF EXISTS products;
  CREATE TABLE products (
    id INTEGER PRIMARY KEY,
    title TEXT,
    description TEXT,
    short TEXT,
    category TEXT,
    subcategory TEXT,
    images TEXT,
    featured INTEGER
  )
`);

const jsonPath = path.join(process.cwd(), 'scripts', 'products.json');
const rawData = fs.readFileSync(jsonPath, 'utf8');
const products = JSON.parse(rawData);

const insert = db.prepare('INSERT INTO products (id, title, description, short, category, subcategory, images, featured) VALUES (@id, @title, @description, @short, @category, @subcategory, @images, @featured)');

const insertMany = db.transaction((items) => {
  for (const item of items) {
    insert.run({
      ...item,
      images: JSON.stringify(item.images),
      featured: item.featured ? 1 : 0
    });
  }
});

insertMany(products);
console.log(`База данных обновлена! Загружено товаров: ${products.length}`);