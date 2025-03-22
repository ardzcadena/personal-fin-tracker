import openpyxl

from utils.excel import read_excel


if __name__ == "__main__":
    file_path = "data/transactions.xlsx"
    data = read_excel(file_path)
    print(data)