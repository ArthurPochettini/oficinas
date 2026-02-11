import time, math, logging, os
from datetime import timedelta
from enum import StrEnum

import numpy as np
import win32api as win32
from wakepy import keep

from mail_operations import find_mail, parse_mail, digest_mail, to_digest_mails, TODAY
from db_operations import parse_table, get_name, open_warranty_db, get_last_maintenance

console_logger = logging.getLogger("CONSOLE")
console_logger.setLevel(logging.INFO)
console_handler = logging.StreamHandler()
console_format = logging.Formatter('%(asctime)s - %(message)s')
console_handler.setFormatter(console_format)
console_logger.addHandler(console_handler)


text_logger = logging.getLogger("TEXT")
text_logger.setLevel(logging.INFO)
text_handler = logging.FileHandler("application.log")
text_format = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
text_handler.setFormatter(text_format)
text_logger.addHandler(text_handler)

class color(StrEnum):
   PURPLE = '\033[95m'
   CYAN = '\033[96m'
   DARKCYAN = '\033[36m'
   BLUE = '\033[94m'
   GREEN = '\033[92m'
   YELLOW = '\033[93m'
   RED = '\033[91m'
   BOLD = '\033[1m'
   UNDERLINE = '\033[4m'
   END = '\033[0m'

def read_loop(cooldown:int):
    while True:
        find_mail()
        warranty_db = open_warranty_db()
        for mail in to_digest_mails:
            if digest_mail(mail):
                table = parse_mail(mail)
                df = parse_table(table)
                equipment_name = get_name(df)
                if not ("BCS" in equipment_name or "MOE" in equipment_name): continue
                equipment_filter = warranty_db[warranty_db["EQUIPAMENTO"] == equipment_name]
                last_maintenance, third_party = get_last_maintenance(equipment_filter)
                if last_maintenance == None:
                    win32.MessageBeep(0x00000030)
                    text_logger.info(f"{equipment_name} não está na garantia")
                    console_logger.info(color.RED + f"{equipment_name} não está na garantia - {third_party}" + color.END)
                if (TODAY - last_maintenance).days <= 90:
                    win32.MessageBeep(0x00000030)
                    text_logger.info(f"{equipment_name} está na garantia")
                    console_logger.info(color.GREEN + f"{equipment_name} está na garantia - {third_party}" + color.END)
                else:
                    win32.MessageBeep(0x00000030)
                    text_logger.info(f"{equipment_name} não está na garantia")
                    console_logger.info(color.RED + f"{equipment_name} não está na garantia - {third_party}" + color.END)
        text_logger.info("Cooling down")
        time.sleep(cooldown)

if __name__ == "__main__":
    with keep.presenting():
        os.system("cls")
        read_loop(15)