
import pandas

grim_all_causes_file = pandas.ExcelFile('grim-all-causes-combined-2017.xlsx')
grim_all_causes = {'male': grim_all_causes_file.parse('Deaths', header=5, usecols=range(0, 21))}

print(grim_all_causes['male'].head())

print pandas.__version__

