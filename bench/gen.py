import sys, random, datetime
from openpyxl import Workbook
rows, cols, out = int(sys.argv[1]), int(sys.argv[2]), sys.argv[3]
random.seed(42)
wb = Workbook(write_only=True)
ws = wb.create_sheet("Daten")
ws.append([f"Spalte_{i}" for i in range(cols)])
words = ["Nord","Süd","Ost","West","Alpha","Beta","Gamma","Delta","offen","erledigt"]
base = datetime.date(2020,1,1)
for r in range(rows):
    row=[]
    for c in range(cols):
        m = c % 5
        if m==0: row.append(f"ID-{r:07d}-{c}")
        elif m==1: row.append(random.randint(1,999999))
        elif m==2: row.append(round(random.uniform(0,10000),2))
        elif m==3: row.append(random.choice(words))
        else: row.append(base + datetime.timedelta(days=random.randint(0,2000)))
    ws.append(row)
wb.save(out)
