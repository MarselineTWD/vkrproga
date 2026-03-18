# Инструкция по запуску программы на другом ПК

## Что нужно для запуска

На другом компьютере должны быть установлены:

- Python 3.11+  
- PostgreSQL  
- зависимости Python из файла `requirements.txt`

Также в папке проекта должен быть настроен файл `.env` с параметрами подключения к базе данных.

## Как перенести проект

1. Скопируйте всю папку проекта на другой компьютер.
2. Убедитесь, что на другом компьютере установлен Python.
3. Убедитесь, что установлен PostgreSQL.

## Как установить зависимости

Откройте терминал в папке проекта и выполните:

```powershell
pip install -r requirements.txt
```

## Что такое `DATABASE_URL`

`DATABASE_URL` — это строка подключения программы к базе данных PostgreSQL.

По ней приложение понимает:

- где находится база данных
- как называется база
- какой пользователь подключается
- какой пароль использовать

Пример:

```env
DATABASE_URL=postgresql://postgres:postgres@localhost:5432/rentability_db
```

Расшифровка:

- `postgresql://` — тип базы данных
- `postgres:postgres` — логин и пароль
- `localhost` — база находится на этом же компьютере
- `5432` — порт PostgreSQL
- `rentability_db` — имя базы данных

## Где нужно указать `DATABASE_URL`

Есть два варианта:

### Вариант 1. Через файл `.env`

Это самый простой вариант.

1. Создайте в папке проекта файл `.env`
2. Добавьте в него строки:

```env
POSTGRES_ADMIN_URL=postgresql://postgres:postgres@localhost:5432/postgres
DATABASE_URL=postgresql://postgres:postgres@localhost:5432/rentability_db
```

Если логин, пароль, порт или имя базы у вас другие, замените их на свои.

### Вариант 2. Через переменные окружения Windows

Можно не создавать `.env`, а задать переменные в системе Windows.

Нужно создать переменную среды:

- `DATABASE_URL`

и присвоить ей значение вида:

```env
postgresql://postgres:postgres@localhost:5432/rentability_db
```

Но для обычного запуска удобнее использовать `.env`.

## Как создать базу данных

Если PostgreSQL уже установлен, создайте базу данных `rentability_db`.

Это можно сделать, например, через pgAdmin или через SQL-команду:

```sql
CREATE DATABASE rentability_db;
```

## Как запустить программу

В папке проекта выполните:

```powershell
python диплом.py
```

## Если программа не запускается

Проверьте:

- установлен ли Python
- установлен ли PostgreSQL
- существует ли база `rentability_db`
- правильно ли указан `DATABASE_URL`
- установлены ли библиотеки из `requirements.txt`

## Быстрый пример готового `.env`

```env
POSTGRES_ADMIN_URL=postgresql://postgres:postgres@localhost:5432/postgres
DATABASE_URL=postgresql://postgres:postgres@localhost:5432/rentability_db
```

