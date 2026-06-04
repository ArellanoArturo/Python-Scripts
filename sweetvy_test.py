import pandas as pd
import sweetviz as sv

 # Load one specific sheet
df = pd.read_excel(r"C:\Users\zomav\Downloads\CSO_KPI_Case.xlsx", sheet_name="Daily Log")
df = df.loc[:, ~df.columns.str.startswith('Unnamed')]
report = sv.analyze(df)
report.show_html("my_report.html")  # opens automatically in your browser
