import pandas as pd
import psycopg2
from psycopg2.extensions import ISOLATION_LEVEL_AUTOCOMMIT
import numpy as np
import os

class DatabaseSetup:
    def __init__(self, host="localhost", user="postgres", password="1234", database="codd"):
        self.host = host
        self.user = user
        self.password = password
        self.database = database
        self.connection = None
        
    def create_database(self):
        """Создание базы данных если не существует"""
        try:
            # Подключение к серверу PostgreSQL для создания БД
            conn = psycopg2.connect(
                host=self.host,
                user=self.user,
                password=self.password,
                database="postgres"
            )
            conn.set_isolation_level(ISOLATION_LEVEL_AUTOCOMMIT)
            cursor = conn.cursor()
            
            # Проверка существования БД
            cursor.execute("SELECT 1 FROM pg_catalog.pg_database WHERE datname = %s", (self.database,))
            exists = cursor.fetchone()
            
            if not exists:
                cursor.execute(f"CREATE DATABASE {self.database} WITH OWNER {self.user} ENCODING 'UTF8'")
                print(f"База данных '{self.database}' создана успешно")
            else:
                print(f"База данных '{self.database}' уже существует")
            
            cursor.close()
            conn.close()
            return True
            
        except Exception as e:
            print(f"Ошибка при создании БД: {e}")
            return False
    
    def connect(self):
        """Подключение к конкретной БД"""
        try:
            self.connection = psycopg2.connect(
                host=self.host,
                database=self.database,
                user=self.user,
                password=self.password
            )
            return True
        except Exception as e:
            print(f"Ошибка подключения: {e}")
            return False
    
    def create_tables(self):
        """Создание таблиц из DB.py"""
        try:
            cursor = self.connection.cursor()
            
            # Создание типов ENUM
            cursor.execute("""
                DO $$ BEGIN
                    CREATE TYPE req_status AS ENUM('new','in_progress','completed');
                EXCEPTION
                    WHEN duplicate_object THEN null;
                END $$;
            """)
            
            cursor.execute("""
                DO $$ BEGIN
                    CREATE TYPE req_role AS ENUM('editor','admin');
                EXCEPTION
                    WHEN duplicate_object THEN null;
                END $$;
            """)
            
            # Таблица для штрафов
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS fines(
                    id SERIAL PRIMARY KEY,
                    date DATE NOT NULL,
                    violations_count INT,
                    decrees_count INT,
                    fines_issued_sum DECIMAL(15,2),
                    fines_collected_sum DECIMAL(15,2)
                )
            """)
            
            # Эвакуация
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS evacuations(
                    id SERIAL PRIMARY KEY,
                    date DATE NOT NULL,
                    evacuators_count INT,
                    trips_count INT,
                    evacuations_count INT,
                    parking_fine_sum DECIMAL(10,2)
                )
            """)
            
            # Реестр светофоров (исправлена опечатка)
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS traffic_lights(
                    id SERIAL PRIMARY KEY,
                    address VARCHAR(255),
                    type VARCHAR(10),
                    install_year INT
                )
            """)
            
            # Таблица эвакуации маршрута
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS evacuations_routes(
                    id SERIAL PRIMARY KEY,
                    year INT,
                    month VARCHAR(20),
                    route TEXT
                ) 
            """)
            
            # Таблица дорожно-транспортных происшествий
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS accidents(
                    id SERIAL PRIMARY KEY,
                    date DATE,
                    incidents_count INT,
                    injured_count INT,
                    fatalities_count INT
                )
            """)
            
            # Таблица для новостей
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS news(
                    id SERIAL PRIMARY KEY,
                    title VARCHAR(255),
                    content TEXT,
                    publish_date DATE,
                    image_url VARCHAR(500),
                    is_published BOOLEAN
                )
            """)
            
            # Таблица Услуги (исправлена опечатка)
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS services(
                    id SERIAL PRIMARY KEY,
                    name VARCHAR(255),
                    description TEXT,
                    price DECIMAL(10,2),
                    is_free BOOLEAN
                )
            """)
            
            # Таблица оставленные заявки
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS service_requests(
                    id SERIAL PRIMARY KEY,
                    service_id INT REFERENCES services(id),
                    customer_name VARCHAR(255),
                    phone VARCHAR(20),
                    address VARCHAR(500),
                    car_type VARCHAR(100),
                    created_at DATE,
                    status req_status
                )
            """)
            
            # Таблица пользователей
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS users(
                    id SERIAL PRIMARY KEY,
                    name VARCHAR(100),
                    password VARCHAR(128),
                    role req_role
                )
            """)
            
            self.connection.commit()
            print("Все таблицы созданы успешно")
            return True
            
        except Exception as e:
            print(f"Ошибка при создании таблиц: {e}")
            self.connection.rollback()
            return False
    
    def import_excel_data(self, excel_file):
        """Импорт данных из Excel файла"""
        try:
            # Чтение Excel файла
            xl_file = pd.ExcelFile(excel_file)
            cursor = self.connection.cursor()
            
            print("Начинаем импорт данных из Excel...")
            
            # Обработка каждого листа
            for sheet_name in xl_file.sheet_names:
                print(f"Обрабатываем лист: {sheet_name}")
                
                # Для разных листов разные заголовки
                if 'Реестр светофоров' in sheet_name:
                    df = pd.read_excel(excel_file, sheet_name=sheet_name, header=2)
                elif 'Эвакуация маршрут' in sheet_name:
                    df = pd.read_excel(excel_file, sheet_name=sheet_name, header=5)
                else:
                    df = pd.read_excel(excel_file, sheet_name=sheet_name, header=2)
                
                # Удаляем строки с NaN в первых столбцах
                if not df.empty:
                    df = df.dropna(subset=[df.columns[0]])
                
                if df.empty:
                    print(f"  Нет данных в листе {sheet_name}")
                    continue
                
                # Определяем тип таблицы по названию листа
                if 'Штрафы' in sheet_name:
                    self.import_fines_data(df, cursor)
                elif 'Эвакуация' in sheet_name and 'маршрут' not in sheet_name:
                    self.import_evacuations_data(df, cursor)
                elif 'Эвакуация маршрут' in sheet_name:
                    self.import_evacuation_routes_data(df, cursor)
                elif 'Реестр светофоров' in sheet_name:
                    self.import_traffic_lights_data(df, cursor)
                else:
                    print(f"  Неизвестный тип листа: {sheet_name}")
            
            self.connection.commit()
            print("Импорт данных завершен успешно")
            return True
            
        except Exception as e:
            print(f"Ошибка при импорте данных: {e}")
            self.connection.rollback()
            return False
    
    def import_fines_data(self, df, cursor):
        """Импорт данных о штрафах"""
        for index, row in df.iterrows():
            try:
                # Преобразуем дату
                date_value = row.iloc[0]
                if pd.isna(date_value):
                    continue
                    
                if isinstance(date_value, str):
                    if ' ' in date_value:
                        date_str = date_value.split(' ')[0]
                    else:
                        date_str = date_value
                else:
                    date_str = date_value.strftime('%Y-%m-%d') if hasattr(date_value, 'strftime') else str(date_value)
                
                # Преобразуем числовые значения
                violations = int(row.iloc[1]) if pd.notna(row.iloc[1]) else 0
                decrees = int(row.iloc[2]) if pd.notna(row.iloc[2]) else 0
                issued = float(row.iloc[3]) if pd.notna(row.iloc[3]) else 0.0
                collected = float(row.iloc[4]) if pd.notna(row.iloc[4]) else 0.0
                
                cursor.execute("""
                    INSERT INTO fines (date, violations_count, decrees_count, fines_issued_sum, fines_collected_sum)
                    VALUES (%s, %s, %s, %s, %s)
                """, (date_str, violations, decrees, issued, collected))
                print(f"  Добавлен штраф за {date_str}")
                
            except Exception as e:
                print(f"  Ошибка при добавлении штрафа (строка {index}): {e}")
    
    def import_evacuations_data(self, df, cursor):
        """Импорт данных об эвакуациях"""
        for index, row in df.iterrows():
            try:
                # Преобразуем дату
                date_value = row.iloc[0]
                if pd.isna(date_value):
                    continue
                    
                if isinstance(date_value, str):
                    if ' ' in date_value:
                        date_str = date_value.split(' ')[0]
                    else:
                        date_str = date_value
                else:
                    date_str = date_value.strftime('%Y-%m-%d') if hasattr(date_value, 'strftime') else str(date_value)
                
                # Преобразуем числовые значения
                evacuators = int(row.iloc[1]) if pd.notna(row.iloc[1]) else 0
                trips = int(row.iloc[2]) if pd.notna(row.iloc[2]) else 0
                evacuations = int(row.iloc[3]) if pd.notna(row.iloc[3]) else 0
                parking_fine = float(row.iloc[4]) if pd.notna(row.iloc[4]) else 0.0
                
                cursor.execute("""
                    INSERT INTO evacuations (date, evacuators_count, trips_count, evacuations_count, parking_fine_sum)
                    VALUES (%s, %s, %s, %s, %s)
                """, (date_str, evacuators, trips, evacuations, parking_fine))
                print(f"  Добавлена эвакуация за {date_str}")
                
            except Exception as e:
                print(f"  Ошибка при добавлении эвакуации (строка {index}): {e}")
    
    def import_evacuation_routes_data(self, df, cursor):
        """Импорт данных о маршрутах эвакуации"""
        for index, row in df.iterrows():
            try:
                if pd.isna(row.iloc[0]) or pd.isna(row.iloc[1]) or pd.isna(row.iloc[2]):
                    continue
                    
                year = int(row.iloc[0]) if pd.notna(row.iloc[0]) else 2024
                month = str(row.iloc[1]) if pd.notna(row.iloc[1]) else ''
                route = str(row.iloc[2]) if pd.notna(row.iloc[2]) else ''
                
                cursor.execute("""
                    INSERT INTO evacuations_routes (year, month, route)
                    VALUES (%s, %s, %s)
                """, (year, month, route))
                print(f"  Добавлен маршрут за {month} {year}")
                
            except Exception as e:
                print(f"  Ошибка при добавлении маршрута (строка {index}): {e}")
    
    def import_traffic_lights_data(self, df, cursor):
        """Импорт данных о светофорах"""
        for index, row in df.iterrows():
            try:
                if pd.isna(row.iloc[1]):  # Проверяем адрес
                    continue
                    
                address = str(row.iloc[1]) if pd.notna(row.iloc[1]) else ''
                light_type = str(row.iloc[2]) if pd.notna(row.iloc[2]) else ''
                install_year = int(row.iloc[3]) if pd.notna(row.iloc[3]) else 2020
                
                cursor.execute("""
                    INSERT INTO traffic_lights (address, type, install_year)
                    VALUES (%s, %s, %s)
                """, (address, light_type, install_year))
                print(f"  Добавлен светофор: {address}")
                
            except Exception as e:
                print(f"  Ошибка при добавлении светофора (строка {index}): {e}")
    
    def add_sample_data(self):
        """Добавление примерных данных для остальных таблиц"""
        try:
            cursor = self.connection.cursor()
            
            # Добавляем примерные услуги
            services = [
                ('Эвакуация автомобиля', 'Услуга по эвакуации транспортного средства', 5000.00, False),
                ('Консультация по ПДД', 'Бесплатная консультация по правилам дорожного движения', 0.00, True),
                ('Оформление ДТП', 'Помощь в оформлении дорожно-транспортного происшествия', 3000.00, False)
            ]
            
            for service in services:
                cursor.execute("""
                    INSERT INTO services (name, description, price, is_free)
                    VALUES (%s, %s, %s, %s)
                """, service)
            
            # Добавляем примерного пользователя
            cursor.execute("""
                INSERT INTO users (name, password, role)
                VALUES (%s, %s, %s)
            """, ('admin', 'admin123', 'admin'))
            
            self.connection.commit()
            print("Примерные данные добавлены успешно")
            
        except Exception as e:
            print(f"Ошибка при добавлении примерных данных: {e}")
    
    def setup_complete_database(self, excel_file):
        """Полная настройка базы данных"""
        print("=== НАЧАЛО НАСТРОЙКИ БАЗЫ ДАННЫХ ===")
        
        # Создание БД
        if not self.create_database():
            return False
        
        # Подключение к БД
        if not self.connect():
            return False
        
        # Создание таблиц
        if not self.create_tables():
            return False
        
        # Импорт данных из Excel
        if os.path.exists(excel_file):
            if not self.import_excel_data(excel_file):
                print("Предупреждение: Возникли ошибки при импорте данных")
        else:
            print(f"Файл {excel_file} не найден, пропускаем импорт данных")
        
        # Добавление примерных данных
        self.add_sample_data()
        
        print("=== НАСТРОЙКА БАЗЫ ДАННЫХ ЗАВЕРШЕНА ===")
        return True

def main():
    """Основная функция"""
    # Параметры подключения
    db_setup = DatabaseSetup(
        host="localhost",
        user="postgres", 
        password="1234",
        database="codd"
    )
    
    # Файл с данными
    excel_file = "CODD.xlsx"
    
    # Запуск полной настройки
    success = db_setup.setup_complete_database(excel_file)
    
    if success:
        print("\nБаза данных успешно создана и заполнена!")
        print("Таблицы: fines, evacuations, traffic_lights, evacuations_routes")
        print("Данные импортированы из Excel файла")
    else:
        print("\nВозникли ошибки при настройке базы данных")
    
    # Закрытие соединения
    if db_setup.connection:
        db_setup.connection.close()

if __name__ == "__main__":
    main()