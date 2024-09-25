import pandas as pd
import requests
import json
import idna
from urllib.parse import urlparse, urlunparse, unquote

def fetch_and_search_keywords(file_path, output_json):
    keywords = ["стажировка", "стажёр", "стажер", "практика", "стажировки"]

    try:
        # Загружаем Excel файл, ограничиваем количество строк до 12483
        df = pd.read_excel(file_path, nrows=12483)
        # Проверка на пустые названия столбцов и корректировка
        df.columns = [str(col) if str(col).strip() else f"Unnamed_{i}" for i, col in enumerate(df.columns)]
    except Exception as e:
        print(f"Ошибка при загрузке файла: {e}")
        return

    with open(output_json, 'w', encoding='utf-8') as f:
        results = []

        # Проверка на существование столбца с ссылками (например, 3-й столбец)
        if len(df.columns) < 4:
            print("Файл не содержит необходимого столбца с ссылками.")
            return

        # Извлекаем ссылки из 3-го столбца
        for link in df.iloc[:, 3].dropna().tolist():
            if not isinstance(link, str) or len(link.strip()) == 0:
                continue

            # Добавляем протокол к ссылке, если его нет
            if not link.startswith(('http://', 'https://')):
                link = 'http://' + link

            try:
                # Декодируем только путь, а не домен (чтобы избежать проблем с IDNA)
                parsed_url = urlparse(link)
                decoded_path = unquote(parsed_url.path)

                # Преобразуем домен в Punycode, если он содержит нелатинские символы
                if parsed_url.hostname:
                    try:
                        hostname_idna = idna.encode(parsed_url.hostname).decode('utf-8')
                    except idna.IDNAError as e:
                        print(f"Ошибка преобразования домена: {parsed_url.hostname}. Пропускаем URL.")
                        continue

                    # Пересобираем URL с Punycode доменом и декодированным путем
                    new_url = urlunparse(parsed_url._replace(netloc=hostname_idna, path=decoded_path))

                    # Проверка корректности нового URL
                    if not new_url:
                        print(f"Некорректный URL: {link}. Пропускаем.")
                        continue

                    print(f"Запрос к {new_url}...")
                    response = requests.get(new_url)
                    response.raise_for_status()

                    page_content = response.text

                    # Поиск ключевых слов на странице
                    if any(keyword.lower() in page_content.lower() for keyword in keywords):
                        result = {"url": new_url, "keywords_found": True}
                        results.append(result)
                        json.dump(result, f, ensure_ascii=False)
                        f.write('\n')

            except requests.RequestException as e:
                print(f"Ошибка при запросе {link}: {e}")
            except UnicodeError as e:
                print(f"Ошибка при обработке ссылки {link}: {e}")

    if len(results) == 0:
        print("Не удалось подключиться ни к одной ссылке или не найдено ключевых слов.")
    else:
        print(f"Найдено совпадений: {len(results)}")

# Вызов функции
fetch_and_search_keywords('DB.xlsx', 'results.json')
