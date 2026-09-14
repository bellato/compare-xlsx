# Какой профиль сравнения использовать.
# None — стандартное сравнение:
#   заголовки из первой строки,
#   ключ — первая колонка.
COMPARE_PROFILE = "calculator"
# COMPARE_PROFILE = "default"


COMPARE_PROFILES = {
    "default": {
        "header_row": 0,
        "key_column": None,
    },
    "calculator": {
        "header_row": 1,
        "key_column": "Артикул",
    },
    # В будущем можно добавлять другие профили:
    # "podsort": {
    #     "header_row": 2,
    #     "key_column": "key",
    # },
}
