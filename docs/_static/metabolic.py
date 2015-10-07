
"""
This example demonstrates log plots using body_mass vs metabolic_rate
Data Taken From: http://www.datacarpentry.org/semester-biology/assignments/r-4/
Data Carpentry for Biologists

Used XY_Math to fit the data 

Fitting Non-Linear Equation to dataset with Percent Error.

y = A*x**c
    A = 0.022110831039
    c = 0.745374452331
    x = Body Mass (g)
    y = Metabolic Rate (mLO2/hr)
    Correlation Coefficient = 0.935357443771
    Standard Deviation = 31.8545944339
    Percent Standard Deviation = 20.5192091584%
y = 0.022110831039*x**0.745374452331

"""

from odscharts.spreadsheet import SpreadSheet

# Grams
body_mass = [32000, 37800, 347000, 4200, 196500, 100000, 4290, 
    32000, 65000, 69125, 9600, 133300, 150000, 407000, 115000, 67000, 
    325000, 21500, 58588, 65320, 85000, 135000, 20500, 1613, 1618]

# kcal/hr
metabolic_rate = [49.984, 51.981, 306.770, 10.075, 230.073, 
    148.949, 11.966, 46.414, 123.287, 106.663, 20.619, 180.150, 
    200.830, 224.779, 148.940, 112.430, 286.847, 46.347, 142.863, 
    106.670, 119.660, 104.150, 33.165, 4.900, 4.865]

dataL = []
for bm, mr in zip(body_mass, metabolic_rate):
    dataL.append( [bm,mr] )
dataL.sort()

# create a list of curve fit data
for row in dataL:
    bm, mr = row
    row.append( 0.022110831039*bm**0.745374452331 )

list_of_rows = [['Body Mass','Rate Data','Fitted Rate'], ['g','mLO2/hr','mLO2/hr']]
list_of_rows.extend( dataL )
    
mySprSht = SpreadSheet()
mySprSht.add_sheet('MassRate_Data', list_of_rows)

mySprSht.add_scatter( 'MassRate_Plot', 'MassRate_Data',
                          title='Metabolic Rate vs Body Mass', 
                          xlabel='Body Mass', 
                          ylabel='Metabolic Rate', 
                          xcol=1, logx=1, logy=1, 
                          showLineL=[0,1], showMarkerL=[1,0],
                          ycolL=[2,3])
mySprSht.setXrange(xmin=1000)
mySprSht.save( filename='mass_vs_rate', launch=True )

