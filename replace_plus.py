# A enhanced replace tool for AutoCAD drawing using regrex match
from caddoc import CADDoc
import time

# example
drawing = CADDoc(r'D:\Work\Project\FRONTEND\XY2020FZ005-Xiangyan.P2\input\05-PID-2020.1103.dwg')
start = time.time()
drawing.replace_text(r'(\D+|^)1(\d{4})', r'\g<1>2\g<2>')
print(f'{(time.time() - start):.2f}s spent.')