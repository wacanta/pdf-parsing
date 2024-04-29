# pdf-parsing

## Requires PyMuPDF version 1.24.0 or later

Install
```
pip install PyMuPDF
```

Run
```
python parse.py <name_of_your_pdf>.pdf
```

Optionally you can define which pages should get parsed. "N" is the document's last page number. 

Example: 
```
python parse.py <name_of_your_pdf>.pdf -pages 2-4,12,14-N
```

Output will be a markdown file: <name_of_your_pdf>.md in the same directory as your PDF. 
