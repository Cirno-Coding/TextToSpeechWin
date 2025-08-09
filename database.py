import os
import sqlite3
from datetime import datetime


class DatabaseManager:
    def __init__(self, db_name='texts.db'):
        self.db_name = db_name
        self.conn = None
        self._initialize_database()

    def _initialize_database(self):
        """Инициализация БД с проверкой структуры"""
        is_new_db = not os.path.exists(self.db_name)

        try:
            self.conn = sqlite3.connect(self.db_name)
            if is_new_db:
                self._create_tables()
            else:
                if not self._check_tables_structure():
                    self._handle_invalid_database()
        except Exception as e:
            raise RuntimeError(f"Ошибка инициализации БД: {str(e)}")

    def _create_tables(self):
        """Создание таблиц в новой БД"""
        cursor = self.conn.cursor()

        cursor.execute('''
            CREATE TABLE categories (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                created_at DATETIME DEFAULT (datetime('now', 'localtime'))
            )
        ''')

        cursor.execute('''
            CREATE TABLE texts (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                category_id INTEGER,
                title TEXT NOT NULL,
                content TEXT NOT NULL,
                sort_index INTEGER DEFAULT 0,
                created_at DATETIME DEFAULT (datetime('now', 'localtime')),
                updated_at DATETIME DEFAULT (datetime('now', 'localtime')),
                FOREIGN KEY(category_id) REFERENCES categories(id)
            )
        ''')
        self.conn.commit()

    def _check_tables_structure(self):
        """Проверка соответствия структуры таблиц"""
        try:
            cursor = self.conn.cursor()

            # Проверяем существование и структуру таблицы categories
            cursor.execute("PRAGMA table_info(categories)")
            categories_columns = {row[1] for row in cursor.fetchall()}
            required_categories = {'id', 'name', 'created_at'}
            if categories_columns != required_categories:
                return False

            # Проверяем существование и структуру таблицы texts
            cursor.execute("PRAGMA table_info(texts)")
            texts_columns = {row[1] for row in cursor.fetchall()}
            required_texts = {'id', 'category_id', 'title', 'content',
                              'sort_index', 'created_at', 'updated_at'}
            if texts_columns != required_texts:
                return False

            return True
        except sqlite3.DatabaseError:
            return False

    def _handle_invalid_database(self):
        """Обработка невалидной БД: переименование и создание новой"""
        self.conn.close()
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_name = f"{self.db_name}.invalid_{timestamp}"

        os.makedirs("dump_files", exist_ok=True)
        dest_path = os.path.join("dump_files", backup_name)
        os.rename(self.db_name, dest_path)

        self.conn = sqlite3.connect(self.db_name)
        self._create_tables()

    # Методы для работы с категориями
    def get_all_categories(self):
        cursor = self.conn.cursor()
        cursor.execute('SELECT id, name FROM categories ORDER BY created_at DESC')
        return cursor.fetchall()

    def add_category(self, name):
        cursor = self.conn.cursor()
        cursor.execute('INSERT INTO categories (name) VALUES (?)', (name,))
        self.conn.commit()
        return cursor.lastrowid

    # Методы для работы с текстами
    def get_texts_by_category(self, category_id):
        cursor = self.conn.cursor()
        cursor.execute('''
            SELECT id, category_id, title, content
            FROM texts
            WHERE category_id = ?
            ORDER BY sort_index, created_at DESC
        ''', (category_id,))
        return cursor.fetchall()

    def get_text_content(self, text_id):
        cursor = self.conn.cursor()
        cursor.execute('SELECT content FROM texts WHERE id = ?', (text_id,))
        result = cursor.fetchall()
        return result[0] if result else ""

    def save_text(self, category_id, title, content):
        cursor = self.conn.cursor()
        cursor.execute('''
            INSERT INTO texts (category_id, title, content)
            VALUES (?, ?, ?)
        ''', (category_id, title, content))
        self.conn.commit()
        return cursor.lastrowid

    def update_text(self, text_id, title, content):
        cursor = self.conn.cursor()
        cursor.execute('''
            UPDATE texts
            SET title = ?, content = ?, updated_at = (datetime('now', 'localtime'))
            WHERE id = ?
        ''', (title, content, text_id))
        self.conn.commit()

    def update_sort_indexes(self, indexes):
        cursor = self.conn.cursor()
        cursor.executemany('''
            UPDATE texts
            SET sort_index = ?
            WHERE id = ?
        ''', indexes)
        self.conn.commit()

    def close(self):
        self.conn.close()


class Category:
    def __init__(self, id, name):
        self.id = id
        self.name = name


class Text:
    def __init__(self, id, title, content, sort_index):
        self.id = id
        self.title = title
        self.content = content
        self.sort_index = sort_index
