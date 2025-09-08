import sys
try:
    import pandas
    import matplotlib
    import ttkbootstrap
    import openpyxl
    print('imports ok')
except Exception as e:
    print('import error:', e)
    sys.exit(1)
