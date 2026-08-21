import pandas as pd
import sys
from pathlib import Path


def compare_files_by_key(file1: Path, file2: Path, key_column: str, sheet_name: str = None, show_count=10):
    print(f"\n--- Сравнение: {file1.name} vs {file2.name}" + (f" (лист: {sheet_name})" if sheet_name else "") + " ---")

    read_kwargs = {"engine": "openpyxl", "header": 1}
    if sheet_name:
        read_kwargs["sheet_name"] = sheet_name

    try:
        df1 = pd.read_excel(file1, **read_kwargs)
        df2 = pd.read_excel(file2, **read_kwargs)
    except Exception as e:
        print(f"Ошибка при чтении файлов: {e}")
        return

    # Проверка наличия ключа
    if key_column not in df1.columns or key_column not in df2.columns:
        cols1 = ", ".join(df1.columns[:10]) + ("..." if len(df1.columns) > 10 else "")
        cols2 = ", ".join(df2.columns[:10]) + ("..." if len(df2.columns) > 10 else "")
        print(f"Колонка '{key_column}' не найдена ни в одном из файлов.")
        print(f"Доступные колонки в файле 1: {cols1}")
        print(f"Доступные колонки в файле 2: {cols2}")
        return

    # Нормализуем ключ.
    df1[key_column] = df1[key_column].astype(str).str.strip()
    df2[key_column] = df2[key_column].astype(str).str.strip()

    # Неуникальный ключ — нормальный случай.
    # Для каждой группы одинаковых ключей нумеруем строки по порядку появления.
    # Затем сравниваем пары (ключ, номер строки внутри группы).
    df1["__occurrence__"] = df1.groupby(key_column, dropna=False).cumcount()
    df2["__occurrence__"] = df2.groupby(key_column, dropna=False).cumcount()

    df1 = df1.set_index([key_column, "__occurrence__"])
    df2 = df2.set_index([key_column, "__occurrence__"])

    # Общие составные ключи: (значение ключа, номер его появления).
    common_keys = df1.index.intersection(df2.index)
    only_in_1 = df1.index.difference(df2.index)
    only_in_2 = df2.index.difference(df1.index)

    def values_equal(v1, v2):
        # Две пустые Excel-ячейки (NaN/None/NaT) считаем одинаковыми.
        if pd.isna(v1) and pd.isna(v2):
            return True
        if pd.isna(v1) or pd.isna(v2):
            return False

        # Для строк игнорируем случайные пробелы по краям.
        if isinstance(v1, str) and isinstance(v2, str):
            return v1.strip() == v2.strip()

        return v1 == v2

    def get_diff_columns(row1: pd.Series, row2: pd.Series):
        # Сравниваем только колонки, присутствующие в обоих файлах.
        return [
            col for col in row1.index
            if col in row2.index and not values_equal(row1[col], row2[col])
        ]

    diff_rows = []

    # Ищем различия по общим ключам и сразу запоминаем конкретные поля.
    for composite_key in common_keys:
        row1 = df1.loc[composite_key]
        row2 = df2.loc[composite_key]
        changed_cols = get_diff_columns(row1, row2)
        if changed_cols:
            key_value, occurrence = composite_key
            diff_rows.append((key_value, occurrence, row1, row2, changed_cols))

    total_diffs = len(diff_rows) + len(only_in_1) + len(only_in_2)

    print(f"Строк в файле 1: {len(df1)}")
    print(f"Строк в файле 2: {len(df2)}")
    print(f"Общие строки по ключу: {len(common_keys)}")
    print(f"Только в файле 1: {len(only_in_1)}")
    print(f"Только в файле 2: {len(only_in_2)}")
    print(f"Найдено различий по совпадающим ключам: {len(diff_rows)}")
    print(f"Всего различий (включая отсутствующие строки): {total_diffs}")

    if total_diffs == 0:
        print("Различий не найдено!")
        return

    # Вывод примеров
    sample_diff = diff_rows[:show_count]
    sample_only_1 = list(only_in_1[:show_count - len(sample_diff)])
    sample_only_2 = list(only_in_2[:show_count - len(sample_diff) - len(sample_only_1)])

    def shorten(value, max_len=30):
        value = str(value)
        return value[:max_len] + ("..." if len(value) > max_len else "")

    print(f"\nПримеры первых {len(sample_diff)} отличающихся строк (по ключу '{key_column}'):\n")

    if sample_diff:
        # Собираем все колонки, в которых есть различия хотя бы в одной
        # из выводимых строк.
        diff_columns = []
        for _, _, _, _, changed_cols in sample_diff:
            for col in changed_cols:
                if col not in diff_columns:
                    diff_columns.append(col)

        # Заголовок: ключевой столбец + только различающиеся поля.
        headers = [key_column] + diff_columns
        print(" | ".join(headers))
        print("-" * max(80, len(" | ".join(headers))))

        for k, occurrence, r1, r2, changed_cols in sample_diff:
            # Номер строки внутри группы показываем только для реально
            # повторяющегося конкретного ключа.
            count1 = int((df1.index.get_level_values(0) == k).sum())
            count2 = int((df2.index.get_level_values(0) == k).sum())
            key_label = f"{k} [#{occurrence + 1}]" if max(count1, count2) > 1 else k

            values1 = [shorten(key_label)]
            values2 = [shorten(key_label)]

            for col in diff_columns:
                # Показываем значение только если эта колонка реально
                # отличается именно в данной строке.
                if col in changed_cols:
                    values1.append(shorten(r1[col]))
                    values2.append(shorten(r2[col]))
                else:
                    values1.append("")
                    values2.append("")

            print("Файл 1: " + " | ".join(values1))
            print("Файл 2: " + " | ".join(values2))
            print("-" * 80)

    def row_to_str(row: pd.Series):
        cols = row.index[:6]
        vals = [shorten(row[c]) for c in cols]
        return " | ".join(vals)

    if sample_only_1:
        print(f"\nПримеры строк, которые есть только в файле 1 (первые {len(sample_only_1)}):\n")
        for composite_key in sample_only_1:
            key_value, occurrence = composite_key
            count1 = int((df1.index.get_level_values(0) == key_value).sum())
            count2 = int((df2.index.get_level_values(0) == key_value).sum())
            label = f"{key_value} [#{occurrence + 1}]" if max(count1, count2) > 1 else key_value
            print(f"Ключ: {label} | {row_to_str(df1.loc[composite_key])}")
        print("-" * 80)

    if sample_only_2:
        print(f"\nПримеры строк, которые есть только в файле 2 (первые {len(sample_only_2)}):\n")
        for composite_key in sample_only_2:
            key_value, occurrence = composite_key
            count1 = int((df1.index.get_level_values(0) == key_value).sum())
            count2 = int((df2.index.get_level_values(0) == key_value).sum())
            label = f"{key_value} [#{occurrence + 1}]" if max(count1, count2) > 1 else key_value
            print(f"Ключ: {label} | {row_to_str(df2.loc[composite_key])}")
        print("-" * 80)


def main():
    script_dir = Path(__file__).resolve().parent

    # Аргументы: file1 file2 [key_column] [sheet_name]
    if len(sys.argv) >= 3:
        f1 = Path(sys.argv[1])
        f2 = Path(sys.argv[2])
        key_col = sys.argv[3] if len(sys.argv) > 3 else "Артикул"
        sheet = sys.argv[4] if len(sys.argv) > 4 else None
        if not f1.is_file() or not f2.is_file():
            print("Ошибка: один или оба файла не найдены.")
            sys.exit(1)
    else:
        # Автовыбор двух первых .xlsx/.xlsm
        xlsx_files = sorted(
            [f for f in script_dir.iterdir() if f.suffix.lower() in [".xlsx", ".xlsm"]]
        )
        if len(xlsx_files) < 2:
            print(f"Недостаточно файлов .xlsx/.xlsm в папке {script_dir}")
            print("Передай файлы явно: python compare_xlsx.py file1.xlsx file2.xlsm sku Расчет")
            sys.exit(1)
        f1, f2 = xlsx_files[0], xlsx_files[1]
        key_col = "Артикул"  # можно поменять на order_id, артикул и т.п.
        sheet = None
        print(f"Аргументов нет — выбраны файлы:\n  1) {f1.name}\n  2) {f2.name}")
        print(f"Используется ключ: '{key_col}'. Если нужен другой — передай его третьим аргументом.")

    compare_files_by_key(f1, f2, key_col, sheet)


if __name__ == "__main__":
    main()
