import camelot.io as  camelot
tables=camelot.read_pdf('./test.pdf' , pages=('1'), flavor='stream')
tables.export('./test.csv',f='csv')