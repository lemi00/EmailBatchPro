"""
Funções utilitárias do EmailBatchPro
"""

from datetime import datetime


def column_number_to_name(index):
    """
    0 -> A
    1 -> B
    ...
    25 -> Z
    26 -> AA
    """

    result = ""

    while index >= 0:
        result = chr(index % 26 + 65) + result
        index = index // 26 - 1

    return result


def column_name_to_number(name):
    """
    A -> 0
    B -> 1
    Z -> 25
    AA -> 26
    """

    result = 0

    name = name.upper().strip()

    for char in name:
        result = result * 26 + ord(char) - ord("A") + 1

    return result - 1


def current_date():
    return datetime.now().strftime("%d-%m-%Y")


def current_time():
    return datetime.now().strftime("%H:%M:%S")


def current_timestamp():
    return datetime.now().strftime("%Y%m%d_%H%M%S")