import sys
from pathlib import Path

from compare_utils import (
    find_differences,
    prepare_dataframe,
    print_diff_examples,
    print_missing_rows,
    print_summary,
    read_excel_file,
)
from config import COMPARE_PROFILE, COMPARE_PROFILES


def get_compare_settings() -> tuple[int, str | None]:
    """Возвращает настройки выбранного профиля сравнения.

    Returns:
        Номер строки заголовков и название ключевого столбца.
        Если профиль не выбран, используются стандартные настройки.
    """
    if COMPARE_PROFILE not in COMPARE_PROFILES:
        raise ValueError(f"Неизвестный профиль сравнения: {COMPARE_PROFILE}")

    profile = COMPARE_PROFILES[COMPARE_PROFILE]

    return (
        profile["header_row"],
        profile["key_column"],
    )


def compare_files_by_key(
    file1: Path,
    file2: Path,
    sheet_name: str | None = None,
    show_count: int = 10,
) -> None:
    """Сравнивает два Excel-файла по ключевому столбцу.

    Args:
        file1: Путь к первому Excel-файлу.
        file2: Путь ко второму Excel-файлу.
        sheet_name: Имя листа. Если не указано, используется первый лист.
        show_count: Максимальное количество примеров различий для вывода.
    """
    print(f"\n--- Сравнение: {file1.name} vs {file2.name}" + (f" (лист: {sheet_name})" if sheet_name else "") + " ---")

    try:
        header_row, key_column = get_compare_settings()
        print(f"Используется ключ: '{key_column}'. Если нужен другой — выбери другой провфиль в config.py.")

        df1 = read_excel_file(file1, sheet_name, header_row)
        df2 = read_excel_file(file2, sheet_name, header_row)
    except Exception as e:
        print(f"Ошибка при чтении файлов: {e}")
        return

    if key_column is None:
        key_column = df1.columns[0]

    # Проверка наличия ключа
    if key_column not in df1.columns or key_column not in df2.columns:
        cols1 = ", ".join(df1.columns[:10]) + ("..." if len(df1.columns) > 10 else "")
        cols2 = ", ".join(df2.columns[:10]) + ("..." if len(df2.columns) > 10 else "")

        print(f"Колонка '{key_column}' не найдена ни в одном из файлов.")
        print(f"Доступные колонки в файле 1: {cols1}")
        print(f"Доступные колонки в файле 2: {cols2}")
        return

    df1 = prepare_dataframe(df1, key_column)
    df2 = prepare_dataframe(df2, key_column)

    diff_rows, only_in_1, only_in_2, common_keys = find_differences(df1, df2)

    total_diffs = print_summary(df1, df2, common_keys, diff_rows, only_in_1, only_in_2)

    if total_diffs == 0:
        print("Различий не найдено!")
        return

    # Вывод примеров
    sample_diff = diff_rows[:show_count]

    remaining_count = show_count - len(sample_diff)

    sample_only_1 = list(only_in_1[:remaining_count])

    remaining_count -= len(sample_only_1)

    sample_only_2 = list(only_in_2[:remaining_count])
    print_diff_examples(df1, df2, sample_diff, key_column)
    print_missing_rows(df1, df2, sample_only_1, sample_only_2)


def main():
    script_dir = Path(__file__).resolve().parent

    # Аргументы: file1 file2 [key_column] [sheet_name]
    if len(sys.argv) >= 3:
        f1 = Path(sys.argv[1])
        f2 = Path(sys.argv[2])
        sheet = sys.argv[3] if len(sys.argv) > 3 else None
        if not f1.is_file() or not f2.is_file():
            print("Ошибка: один или оба файла не найдены.")
            sys.exit(1)
    else:
        # Автовыбор двух первых .xlsx/.xlsm
        xlsx_files = sorted([f for f in script_dir.iterdir() if f.suffix.lower() in [".xlsx", ".xlsm"]])
        if len(xlsx_files) < 2:
            print(f"Недостаточно файлов .xlsx/.xlsm в папке {script_dir}")
            print("Передай файлы явно: python compare_xlsx.py file1.xlsx file2.xlsm sku Расчет")
            sys.exit(1)
        f1, f2 = xlsx_files[0], xlsx_files[1]
        sheet = None
        print(f"Аргументов нет — выбраны файлы:\n  1) {f1.name}\n  2) {f2.name}")

    compare_files_by_key(f1, f2, sheet)


if __name__ == "__main__":
    main()
