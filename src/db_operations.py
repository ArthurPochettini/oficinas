import re, os, time
from io import StringIO
from datetime import datetime
import warnings

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

import pandas as pd

PROG = re.compile("[A-Z]{3} [0-9]{4}")
DEBUG: bool = False

last_modified: datetime = None
digested_mails = list()

def parse_table(table):
    df:pd.DataFrame = pd.read_html(StringIO(str(table)))[0]
    df = df.dropna(axis=0, how="all")
    df = df.dropna(axis=1, how="all")
    df = df.drop([0,1])
    df.index = range(len(df.index))
    df.columns = ["Informação", "Dado"]

    return df

def get_name(df:pd.DataFrame) -> str:
    raw_name: str = df.at[2, "Dado"]
    return PROG.search(raw_name).group(0)

def open_warranty_db():
    global last_modified
    if DEBUG:
        warranty_path = r"C:\Users\arthur.pseixas\OneDrive - Grupo Aguas do Brasil\Área de Trabalho\OFICINA\resources\OFICINA_TERCEIRA 2.xlsx"
    else:
        warranty_path:str = r"\\prd-fsgab-01\INSTITUCIONAL_ETE$\ELETROMECANICA\Pastas de Funcionários\Felipe Rizzetto\OFICINA_TERCEIRA.xlsx"
    error_count: int = 0
    while True:
        try:
            df: pd.DataFrame = pd.read_excel(warranty_path, engine="openpyxl")
            break
        except FileNotFoundError:
            if error_count >= 10:
                raise Exception
            error_count += 1
            time.sleep(10)
            continue
    return df

def get_last_maintenance(equipment_filter:pd.DataFrame) -> tuple[datetime, str]:
    last_maintenance = equipment_filter["DATA "].max()
    third_party:pd.DataFrame = equipment_filter[equipment_filter["DATA "] == last_maintenance]
    if type(last_maintenance) == float: return (None, None)
    else: return last_maintenance, (third_party["terceira"].values)[0]
    