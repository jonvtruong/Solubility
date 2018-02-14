import xlwt, os, time
from Solubility import app

fileName = 'test.xls'
ssDict = {  'ssShakerDelayTimeHI': 1080, 
                'ssShakerSpeedHI': 700, 
                'ssNumTempsH': 3, 
                'ssNumTempsI':0, 
                'ssTemp1H': 25,
                'ssTemp2H': 25,
                'ssTemp3H': 25,
                'ssTemp4H': 0,
                'ssTemp1I': 0,
                'ssTemp2I': 0,
                'ssTemp3I': 0,
                'ssTemp4I': 0,
                'ssSerialDilAddOn': 'FALSE',
                'ssPolyScreenAddOn': 'FALSE'
            }

gblDict = {     'gblNumVialsHuber': 24,
                'gblNumVialsInheco': 0,
                'gblPlateType': 'DWP',
                'gblEmail': 'jonathan.truong@Takeda.com',
                'gblSysLiqForDil': 'MeCN',
                'gblMultiSolventDilution': 'FALSE',
                'gblMultiSolvDilVol1': 0,
                'gblMultiSolvDilVol2': 0,
                'gblSysLiqForDil2': 0
        }

solventRDict = [
    'MTBE',
    'Ethyl Acetate',
    '1,4-Dioxane',
    'Dichloromethane',
    'Acetonitrile',
    'Water',
    'Ethanol',
    'Methyl Ethyl Ketone',
    'THF',
    'Anisole',
    'Toluene',
    'Methyl Isobutyl Ketone',	
    'Pyridine',
    'Tri-Fluoro Toluene',
    'Methanol',
    'Heptane',
    'Acetone',
    '2-MethylTHF',
    'IPAc',
    'Nitromethane',
    'IPA',
    '2-Butanol',
    'Dimethyl Sulfoxide',	
    'NMP'
]

solventSDict = [
    'Amyl Acetate',	
    'Methyl Acetate',	
    'Butyl Acetate',	
    'Dimethyl Acetamide',	
    'DMF',	
    'Benzyl Alcohol',	
    '1-Butanol',	
    'Dimethoxyethane',	
    'Cyclohexane',	
    '1-Propanol',	
    'Acetic Acid',	
    'Butyronitrile',	
    'CPME',	
    'Diethyl Carbonate',	
    'Ethyl Lactate',	
    'Tetramethylurea',	
    'S17',	
    'S18',	
    'S19',	
    'S20',	
    'S21',	
    'S22',	
    'S23',	
    'S24'	
]

def saveFile():
    try:
        wb = xlwt.Workbook()
        writeVarList(wb)
        writeHPipList(wb)
        wb.save(fileName)

    except PermissionError:
        #close the file wait a second after closing before trying to save to avoid error
        os.system("taskkill /f /im EXCEL.exe")
        time.sleep(1) 
        wb.save(fileName)

def writeHPipList(wb):
    hPipList = wb.add_sheet('HUBER_PIPLIST')

def writeVarList(wb):
    varList = wb.add_sheet('VARIABLESLIST')
    varList.write(0, 0, "VariablesName", xlwt.easyxf('font: bold on; align: vert center, horiz left'))
    varList.write(0, 1,"VariablesValue", xlwt.easyxf('font: bold on; align: vert center, horiz left'))

    row = 1
    for k,v in sorted(ssDict.items()): #write ss variables in variableslist
        varList.write(row, 0, k, xlwt.easyxf('pattern: pattern solid, fore_color pale_blue; align: vert center, horiz left'))
        varList.write(row, 1, v, xlwt.easyxf('pattern: pattern solid, fore_color pale_blue; align: vert center, horiz center'))
        row+=1

    for k,v in sorted(gblDict.items()): #write gbl variable in variables list
        varList.write(row, 0, k, xlwt.easyxf('pattern: pattern solid, fore_color light_yellow; align: vert center, horiz left'))
        varList.write(row, 1, v, xlwt.easyxf('pattern: pattern solid, fore_color light_yellow; align: vert center, horiz center'))
        row+=1