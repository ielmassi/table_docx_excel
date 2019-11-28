import pandas as pd
from docx.api import Document
document = Document('doc1.docx')
table = document.tables

for j in range (0,len(table)):
	data = []
	keys = None
	for i, row in enumerate(table[j].rows):
		text = (cell.text for cell in row.cells)
		if i == 0:
			keys = tuple(text)
			continue
		row_data = dict(zip(keys, text))
		data.append(row_data)
	df = pd.DataFrame(data)
	df.to_excel(str(j)+'.xlsx', index=False)
