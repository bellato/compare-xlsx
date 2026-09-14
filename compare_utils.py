from pathlib import Path

import pandas as pd


def read_excel_file(file: Path, sheet_name: str | None = None, header_row: int = 0) -> pd.DataFrame:
    """Читает Excel-файл в DataFrame.

    Args:
        file: Путь к Excel-файлу.
        sheet_name: Имя листа. Если не указано, используется первый лист.
        header_row: Индекс строки с заголовками.
            0 — первая строка,
            1 — вторая строка и т. д.

    Returns:
        DataFrame с данными из Excel-файла.
    """
    return pd.read_excel(
        file,
        engine="openpyxl",
        header=header_row,
        sheet_name=sheet_name or 0,
    )


def prepare_dataframe(df: pd.DataFrame, key_column: str) -> pd.DataFrame:
    """Подготавливает DataFrame к сравнению.

    Нормализует значения ключевого столбца и добавляет порядковый
    номер строки для повторяющихся ключей.

    Args:
        df: Исходный DataFrame.
        key_column: Название ключевого столбца.

    Returns:
        DataFrame с составным индексом из ключа и номера его появления.
    """
    df = df.copy()

    df[key_column] = df[key_column].astype(str).str.strip()

    df["__occurrence__"] = df.groupby(key_column, dropna=False).cumcount()

    return df.set_index([key_column, "__occurrence__"])


def values_equal(v1, v2) -> bool:
    """Проверяет два значения на равенство с учётом пустых значений."""
    if pd.isna(v1) and pd.isna(v2):
        return True

    if pd.isna(v1) or pd.isna(v2):
        return False

    if isinstance(v1, str) and isinstance(v2, str):
        return v1.strip() == v2.strip()

    return v1 == v2


def get_diff_columns(row1: pd.Series, row2: pd.Series) -> list[str]:
    """Возвращает названия колонок с различающимися значениями."""
    return [col for col in row1.index if col in row2.index and not values_equal(row1[col], row2[col])]


def find_differences(df1: pd.DataFrame, df2: pd.DataFrame) -> tuple[list, pd.Index, pd.Index, pd.Index]:
    """Находит различия между двумя подготовленными DataFrame.

    Args:
        df1: Первый DataFrame с составным индексом
            (ключ, номер появления ключа).
        df2: Второй DataFrame с таким же составным индексом.

    Returns:
        Кортеж из:
        - списка различающихся строк;
        - индексов строк, которые есть только в первом DataFrame;
        - индексов строк, которые есть только во втором DataFrame;
        - общих индексов обоих DataFrame.
    """
    common_keys = df1.index.intersection(df2.index)
    only_in_1 = df1.index.difference(df2.index)
    only_in_2 = df2.index.difference(df1.index)

    diff_rows = []

    for composite_key in common_keys:
        row1 = df1.loc[composite_key]
        row2 = df2.loc[composite_key]

        changed_cols = get_diff_columns(row1, row2)

        if changed_cols:
            key_value, occurrence = composite_key
            diff_rows.append((key_value, occurrence, row1, row2, changed_cols))

    return diff_rows, only_in_1, only_in_2, common_keys


def print_summary(
    df1: pd.DataFrame,
    df2: pd.DataFrame,
    common_keys: pd.Index,
    diff_rows: list,
    only_in_1: pd.Index,
    only_in_2: pd.Index,
) -> int:
    """Выводит сводную информацию о результатах сравнения.

    Args:
        df1: Первый сравниваемый DataFrame.
        df2: Второй сравниваемый DataFrame.
        common_keys: Общие индексы обоих DataFrame.
        diff_rows: Список различающихся строк.
        only_in_1: Индексы строк, которые есть только в первом DataFrame.
        only_in_2: Индексы строк, которые есть только во втором DataFrame.

    Returns:
        Общее количество найденных различий.
    """
    total_diffs = len(diff_rows) + len(only_in_1) + len(only_in_2)

    print(f"Строк в файле 1: {len(df1)}")
    print(f"Строк в файле 2: {len(df2)}")
    print(f"Общие строки по ключу: {len(common_keys)}")
    print(f"Только в файле 1: {len(only_in_1)}")
    print(f"Только в файле 2: {len(only_in_2)}")
    print(f"Найдено различий по совпадающим ключам: {len(diff_rows)}")
    print(f"Всего различий (включая отсутствующие строки): {total_diffs}")

    return total_diffs


def shorten(value, max_len: int = 30) -> str:
    """Сокращает строковое представление значения до заданной длины."""
    value = str(value)
    return value[:max_len] + ("..." if len(value) > max_len else "")


def row_to_str(row: pd.Series) -> str:
    """Преобразует первые шесть значений строки в строку для вывода.

    Args:
        row: Строка DataFrame.

    Returns:
        Строковое представление первых шести значений.
    """
    cols = row.index[:6]
    vals = [shorten(row[col]) for col in cols]

    return " | ".join(vals)


def print_diff_examples(df1: pd.DataFrame, df2: pd.DataFrame, sample_diff: list, key_column: str) -> None:
    """Выводит примеры строк с различающимися значениями.

    Показывает ключевой столбец и только те поля,
    в которых найдены различия.

    Args:
        df1: Первый сравниваемый DataFrame.
        df2: Второй сравниваемый DataFrame.
        sample_diff: Список различающихся строк для вывода.
        key_column: Название ключевого столбца.
    """

    print(f"\nПримеры первых {len(sample_diff)} отличающихся строк (по ключу '{key_column}'):\n")

    if not sample_diff:
        return

    diff_columns = []

    for _, _, _, _, changed_cols in sample_diff:
        for col in changed_cols:
            if col not in diff_columns:
                diff_columns.append(col)

    headers = [key_column] + diff_columns

    print(" | ".join(headers))
    print("-" * max(80, len(" | ".join(headers))))

    for key, occurrence, row1, row2, changed_cols in sample_diff:
        count1 = int((df1.index.get_level_values(0) == key).sum())
        count2 = int((df2.index.get_level_values(0) == key).sum())

        key_label = f"{key} [#{occurrence + 1}]" if max(count1, count2) > 1 else key

        values1 = [shorten(key_label)]
        values2 = [shorten(key_label)]

        for col in diff_columns:
            if col in changed_cols:
                values1.append(shorten(row1[col]))
                values2.append(shorten(row2[col]))
            else:
                values1.append("")
                values2.append("")

        print("Файл 1: " + " | ".join(values1))
        print("Файл 2: " + " | ".join(values2))
        print("-" * 80)


def print_missing_rows(df1: pd.DataFrame, df2: pd.DataFrame, sample_only_1: list, sample_only_2: list) -> None:
    """Выводит примеры строк, которые есть только в одном из файлов.

    Args:
        df1: Первый сравниваемый DataFrame.
        df2: Второй сравниваемый DataFrame.
        sample_only_1: Индексы строк, которые есть только в первом DataFrame.
        sample_only_2: Индексы строк, которые есть только во втором DataFrame.
    """
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
