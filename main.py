import xlwings as xw
import pandas as pd


wb = xw.Book()

#L2L
wb.sheets.add(name='L2L')
sheet = wb.sheets['L2L']
sheet.range('A1','D1').value = "Cell Name","Neighbouring Cell Name", "EARFCN","Relation Type"
sheet.range('A1','D1').color = (253, 233, 217)
wb_new = xw.Book(r'L2L.xlsx')
wb_new.activate(steal_focus=False)
Cell_name = wb_new.sheets['L2L'].range('A1:A10000').options(ndim=2).value
Neighbouring_cell_name = wb_new.sheets['L2L'].range('B1:B10000').options(ndim=2).value
RType = wb_new.sheets['L2L'].range('F1:F10000').options(ndim=2).value
wb.activate(steal_focus=False)
sheet = wb.sheets['L2L']
wb.sheets['L2L'].range('A1:A10000').value = Cell_name
wb.sheets['L2L'].range('D1:D10000').value = RType
wb.sheets['L2L'].range('B1:B10000').value = Neighbouring_cell_name
wb.save("ЧТП.xlsx")
excel_file_path = "4g_cells_all.xlsx"
excel_file_path = "ЧТП.xlsx"
EARFCN = pd.read_excel("ЧТП.xlsx", sheet_name="L2L")
EARFCN1 = pd.read_excel("4g_cells_all.xlsx", sheet_name="4g_cells_all")
df1 = EARFCN.merge(EARFCN1, left_on="Neighbour", right_on="Transmitter", how="left")
sheet.range("C1").options(index=False).value = df1["Channel Number"]

#L2U
wb.sheets.add(name='L2U')
sheet = wb.sheets['L2U']
sheet.range('A1','G1').value = "Cell Name","Neigh Cell Name", "Neigh Cell ID","UARFCN","LAC","RNC","SC"
sheet.range('A1','G1').color = (253, 233, 217)
wb_l2u = xw.Book(r'L2U.xlsx')
wb_l2u.activate(steal_focus=False)
Cell_name1 = wb_l2u.sheets['L2U'].range('A1:A10000').options(ndim=2).value
Neighbouring_cell_ID = wb_l2u.sheets['L2U'].range('B1:B10000').options(ndim=2).value
wb.activate(steal_focus=False)
sheet = wb.sheets['L2U']
wb.sheets['L2U'].range('A1:A10000').value = Cell_name1
wb.sheets['L2U'].range('C1:C10000').value = Neighbouring_cell_ID
wb_l2uu = xw.Book(r'3g_transmitters_all.xlsx')
wb_l2uu.activate(steal_focus=False)
LAC = wb_l2uu.sheets['3g_transmitters_all'].range('AQ1:AQ10000').options(ndim=2).value
RNC = wb_l2uu.sheets['3g_transmitters_all'].range('AR1:AR10000').options(ndim=2).value
wb.activate(steal_focus=False)
sheet = wb.sheets['L2U']
wb.sheets['L2U'].range('F1:F10000').value = RNC
wb.sheets['L2U'].range('E1:E10000').value = LAC
wb.save("ЧТП.xlsx")
excel_file_path = "3g_transmitters_all.xlsx"
excel_file_path = "ЧТП.xlsx"
excel_file_path = "3g_cells_all.xlsx"
Cell_nameu = pd.read_excel("ЧТП.xlsx", sheet_name="L2U")
Cell_nameu1 = pd.read_excel("3g_transmitters_all.xlsx", sheet_name="3g_transmitters_all")
SC = pd.read_excel("ЧТП.xlsx", sheet_name="L2U")
SC1 = pd.read_excel("3g_cells_all.xlsx", sheet_name="3g_cells_all")
df3 = Cell_nameu.merge(Cell_nameu1, left_on="Neighbour", right_on="Transmitter", how="left")
sheet.range("B1").options(index=False).value = df3["Cell_name"]
df4 = SC.merge(SC1, left_on="Neighbour", right_on="Transmitter", how="left")
sheet.range("G1").options(index=False).value = df4["Primary scrambling code"]
UARFCN = pd.read_excel("ЧТП.xlsx", sheet_name="L2U")
UARFCN1 = pd.read_excel("3g_transmitters_all.xlsx", sheet_name="3g_transmitters_all")
df5 = UARFCN.merge(UARFCN1, left_on="Neighbour", right_on="Transmitter", how="left")
sheet.range("D1").options(index=False).value = df5["Downlink_UARFCN"]

#L2G
wb.sheets.add(name='L2G')
sheet = wb.sheets['L2G']
sheet.range('A1','F1').value = "Cell Name","Neighbourng Cell Name", "Neighbourng Cell ID","BCCH1","BSC","LAC"
sheet.range('A1','F1').color = (253, 233, 217)
wb_l2g = xw.Book(r'L2G.xlsx')
wb_l2g.activate(steal_focus=False)
Cell_name12 = wb_l2g.sheets['L2G'].range('A1:A10000').options(ndim=2).value
Neighbouring_cell_ID1 = wb_l2g.sheets['L2G'].range('B1:B10000').options(ndim=2).value
wb.activate(steal_focus=False)
sheet = wb.sheets['L2G']
wb.sheets['L2G'].range('A1:A10000').value = Cell_name12
wb.sheets['L2G'].range('C1:C10000').value = Neighbouring_cell_ID1
wb_l2g = xw.Book(r'2g_transmitters_all.xlsx')
wb_l2g.activate(steal_focus=False)
BSC = wb_l2g.sheets['2g_transmitters_all'].range('BQ1:BQ10000').options(ndim=2).value
LAC1 = wb_l2g.sheets['2g_transmitters_all'].range('BP1:BP10000').options(ndim=2).value
wb.activate(steal_focus=False)
sheet = wb.sheets['L2G']
wb.sheets['L2G'].range('E1:E10000').value = BSC
wb.sheets['L2G'].range('F1:F10000').value = LAC1
wb.save("ЧТП.xlsx")
excel_file_path = "2g_transmitters_all.xlsx"
excel_file_path = "ЧТП.xlsx"
BCCH1 = pd.read_excel("ЧТП.xlsx", sheet_name="L2G")
BCCH2 = pd.read_excel("2g_transmitters_all.xlsx", sheet_name="2g_transmitters_all")
Neighbouring_cell_name1 = pd.read_excel("ЧТП.xlsx", sheet_name="L2G")
Neighbouring_cell_name12 = pd.read_excel("2g_transmitters_all.xlsx", sheet_name="2g_transmitters_all")
df6 = BCCH1.merge(BCCH2, left_on="Neighbour", right_on="Transmitter", how="left")
sheet.range("D1").options(index=False).value = df6["BCCH"]
df7 = Neighbouring_cell_name1.merge(Neighbouring_cell_name12, left_on="Neighbour", right_on="Transmitter", how="left")
sheet.range("B1").options(index=False).value = df7["Cell_name"]

#LTE
wb.sheets.add(name='LTE')
sheet = wb.sheets['LTE']
sheet.range('A1','K1').value = "Site ID","BS ID", "Cell Name","EnodeB_ID(20000+BS_ID)","Cell ID","Band","EARFCN","PCI","Root Seq","TAC","Power, dbm"
sheet.range('A1','K1').color = (253, 233, 217)
wb_lte = xw.Book(r'4g_transmitters_new.xlsx')
wb_lte.activate(steal_focus=False)
Cell_name123 = wb_lte.sheets['4g_transmitters_new'].range('A2:A20000').options(ndim=2).value
СellID1 = wb_lte.sheets['4g_transmitters_new'].range('B2:B20000').options(ndim=2).value
EnodeB_ID = wb_lte.sheets['4g_transmitters_new'].range('AM2:AM10000').options(ndim=2).value
Сell_ID = wb_lte.sheets['4g_transmitters_new'].range('AL2:AL10000').options(ndim=2).value
wb_lte1 = xw.Book(r'4g_cells_new.xlsx')
wb_lte1.activate(steal_focus=False)
Power = wb_lte1.sheets['4g_cells_new'].range('C2:C20000').options(ndim=2).value
FBand = wb_lte1.sheets['4g_cells_new'].range('D2:D20000').options(ndim=2).value
CHNamber = wb_lte1.sheets['4g_cells_new'].range('E2:E20000').options(ndim=2).value
PCI = wb_lte1.sheets['4g_cells_new'].range('S2:S20000').options(ndim=2).value
ROOT = wb_lte1.sheets['4g_cells_new'].range('BD2:BD20000').options(ndim=2).value

wb.activate(steal_focus=False)
sheet = wb.sheets['LTE']
wb.sheets['LTE'].range('A2:A10000').value = Cell_name123
wb.sheets['LTE'].range('C2:C10000').value = СellID1
wb.sheets['LTE'].range('D2:D10000').value = EnodeB_ID
wb.sheets['LTE'].range('E2:E10000').value = Сell_ID
wb.sheets['LTE'].range('K2:K10000').value = Power
wb.sheets['LTE'].range('F2:F10000').value = FBand
wb.sheets['LTE'].range('G2:G10000').value = CHNamber
wb.sheets['LTE'].range('H2:H10000').value = PCI
wb.sheets['LTE'].range('I2:I10000').value = ROOT
sheet.range('J2:J1000').value = 49100
sheet.range('Z2:Z1000').value = 20000
sheet.range('Z1').value = "numb"
wb.save("ЧТП.xlsx")
df8 = pd.read_excel("ЧТП.xlsx", sheet_name="LTE")
df9= df8['EnodeB_ID(20000+BS_ID)']-df8['numb']
sheet.range("B2").options(index=False).value = df9
#U2L
wb.sheets.add(name='U2L')
sheet = wb.sheets['U2L']
sheet.range('A1','G1').value = "Cell Name","Cell ID", "Neigh Cell Name","Band","Cell ID","Band","TAC"
sheet.range('A1','G1').color = (253, 233, 217)
wb_new = xw.Book(r'U2L.xlsx')
wb_new.activate(steal_focus=False)
Cell_ID1234 = wb_new.sheets['U2L'].range('A2:A10000').options(ndim=2).value
Neighbouring_cell_name123 = wb_new.sheets['U2L'].range('B2:B10000').options(ndim=2).value
wb.activate(steal_focus=False)
sheet = wb.sheets['U2L']
wb.sheets['U2L'].range('B2:B10000').value = Cell_ID1234
wb.sheets['U2L'].range('C2:C10000').value = Neighbouring_cell_name123
wb.save("ЧТП.xlsx")
excel_file_path = "4g_cells_all.xlsx"
excel_file_path = "ЧТП.xlsx"
excel_file_path = "3g_transmitters_all.xlsx"
Cell_name1234 = pd.read_excel("ЧТП.xlsx", sheet_name="U2L")
Cell_name12345 = pd.read_excel("3g_transmitters_all.xlsx", sheet_name="3g_transmitters_all")
df9 = Cell_name1234.merge(Cell_name12345, left_on="Cell ID", right_on="Transmitter", how="left")
sheet.range("A1").options(index=False).value = df9["Cell_name"]
sheet.range('G2:G1000').value = 49100
EARFCN12 = pd.read_excel("ЧТП.xlsx", sheet_name="U2L")
EARFCN123 = pd.read_excel("4g_cells_all.xlsx", sheet_name="4g_cells_all")
df10 = EARFCN12.merge(EARFCN123, left_on="Neigh Cell Name", right_on="Transmitter", how="left")
sheet.range("E1").options(index=False).value = df10["Channel Number"]
Band2 = pd.read_excel("ЧТП.xlsx", sheet_name="U2L")
Band12 = pd.read_excel("4g_cells_all.xlsx", sheet_name="4g_cells_all")
df11 = Band2.merge(Band12, left_on="Neigh Cell Name", right_on="Transmitter", how="left")
sheet.range("D1").options(index=False).value = df11["Frequency Band"]
PCI12 = pd.read_excel("ЧТП.xlsx", sheet_name="U2L")
PCI123 = pd.read_excel("4g_cells_all.xlsx", sheet_name="4g_cells_all")
df12 = PCI12.merge(PCI123, left_on="Neigh Cell Name", right_on="Transmitter", how="left")
sheet.range("F1").options(index=False).value = df12["Physical Cell ID"]

#U2U
wb.sheets.add(name='U2U')
sheet = wb.sheets['U2U']
sheet.range('A1','D1').value = "Cell ID","Cell Name", "Neighboring Cell ID","Neighboring Cell Name"
sheet.range('A1','D1').color = (253, 233, 217)
wb_u2u = xw.Book(r'U2U.xlsx')
wb_u2u.activate(steal_focus=False)
Cell_name1u = wb_u2u.sheets['U2U'].range('A2:A10000').options(ndim=2).value
Cell_name_Nei = wb_u2u.sheets['U2U'].range('B2:B10000').options(ndim=2).value
wb.activate(steal_focus=False)
sheet = wb.sheets['U2U']
wb.sheets['U2U'].range('A2:A10000').value = Cell_name1u
wb.sheets['U2U'].range('C2:C10000').value = Cell_name_Nei
wb.save("ЧТП.xlsx")
excel_file_path = "3g_transmitters_all.xlsx"
excel_file_path = "ЧТП.xlsx"
excel_file_path = "3g_cells_all.xlsx"
Cell_nameu = pd.read_excel("ЧТП.xlsx", sheet_name="U2U")
Cell_nameu1 = pd.read_excel("3g_transmitters_all.xlsx", sheet_name="3g_transmitters_all")
CL = pd.read_excel("ЧТП.xlsx", sheet_name="U2U")
CL1 = pd.read_excel("3g_transmitters_all.xlsx", sheet_name="3g_transmitters_all")
df12 = Cell_nameu.merge(Cell_nameu1, left_on="Cell ID", right_on="Transmitter", how="left")
sheet.range("B1").options(index=False).value = df12["Cell_name"]
df13 = CL.merge(CL1, left_on="Neighboring Cell ID", right_on="Transmitter", how="left")
sheet.range("D1").options(index=False).value = df13["Cell_name"]


#U2G
wb.sheets.add(name='U2G')
sheet = wb.sheets['U2G']
sheet.range('A1','G1').value = "3G Cell ID","3G Cell Name", "Neighbouring 2G Cell ID","Neighboring Cell Name","BCCH1","BSC","LAC"
sheet.range('A1','G1').color = (253, 233, 217)
wb_u2g = xw.Book(r'U2G.xlsx')
wb_u2g.activate(steal_focus=False)
Cell_nam12 = wb_u2g.sheets['U2G'].range('A2:A10000').options(ndim=2).value
Neighbouring_cel_ID1 = wb_u2g.sheets['U2G'].range('B2:B10000').options(ndim=2).value
wb.activate(steal_focus=False)
sheet = wb.sheets['U2G']
wb.sheets['U2G'].range('A2:A10000').value = Cell_nam12
wb.sheets['U2G'].range('C2:C10000').value = Neighbouring_cel_ID1
wb_l2gg = xw.Book(r'2g_transmitters_all.xlsx')
wb_l2gg.activate(steal_focus=False)
BSC1 = wb_l2gg.sheets['2g_transmitters_all'].range('BQ2:BQ10000').options(ndim=2).value
LAC2 = wb_l2gg.sheets['2g_transmitters_all'].range('BP2:BP10000').options(ndim=2).value
wb.activate(steal_focus=False)
sheet = wb.sheets['U2G']
wb.sheets['U2G'].range('F2:F10000').value = BSC1
wb.sheets['U2G'].range('G2:G10000').value = LAC2
wb.save("ЧТП.xlsx")
excel_file_path = "2g_transmitters_all.xlsx"
excel_file_path = "ЧТП.xlsx"
excel_file_path = "3g_transmitters_all.xlsx"
BCCH2 = pd.read_excel("ЧТП.xlsx", sheet_name="U2G")
BCCH3 = pd.read_excel("2g_transmitters_all.xlsx", sheet_name="2g_transmitters_all")
Neighbouring_cell_name34 = pd.read_excel("ЧТП.xlsx", sheet_name="U2G")
Neighbouring_cell_name345 = pd.read_excel("2g_transmitters_all.xlsx", sheet_name="2g_transmitters_all")
df14 = BCCH2.merge(BCCH3, left_on="Neighbouring 2G Cell ID", right_on="Transmitter", how="left")
sheet.range("E1").options(index=False).value = df14["BCCH"]
df15 = Neighbouring_cell_name34.merge(Neighbouring_cell_name345, left_on="Neighbouring 2G Cell ID", right_on="Transmitter", how="left")
sheet.range("D1").options(index=False).value = df15["Cell_name"]
cell_name34 = pd.read_excel("ЧТП.xlsx", sheet_name="U2G")
cell_name345 = pd.read_excel("3g_transmitters_all.xlsx", sheet_name="3g_transmitters_all")
df16 = cell_name34.merge(cell_name345, left_on="3G Cell ID", right_on="Transmitter", how="left")
sheet.range("B1").options(index=False).value = df16["Cell_name"]

#UMTS
wb.sheets.add(name='UMTS')
sheet = wb.sheets['UMTS']
sheet.range('A1','K1').value = "Site ID","Cell ID", "Cell Name","LAC","RNC","UARFCN","MaxPwr","CPICH Pwr","Time_Offset","DL PSC","Band Indicator"
sheet.range('A1','K1').color = (253, 233, 217)
wb_umts = xw.Book(r'3g_transmitters_new.xlsx')
wb_umts.activate(steal_focus=False)
Cell_name12367 = wb_umts.sheets['3g_transmitters_new'].range('A2:A20000').options(ndim=2).value
СellID145 = wb_umts.sheets['3g_transmitters_new'].range('B2:B20000').options(ndim=2).value
cell_nam345 = wb_umts.sheets['3g_transmitters_new'].range('AP2:AP10000').options(ndim=2).value
Downlink_UARFCN = wb_umts.sheets['3g_transmitters_new'].range('AU2:AU10000').options(ndim=2).value
Frequency_Band = wb_umts.sheets['3g_transmitters_new'].range('F2:F10000').options(ndim=2).value
Time_Offset = wb_umts.sheets['3g_transmitters_new'].range('AT2:AT10000').options(ndim=2).value
wb_umts1 = xw.Book(r'3g_cells_new.xlsx')
wb_umts1.activate(steal_focus=False)
Power1 = wb_umts1.sheets['3g_cells_new'].range('G2:G20000').options(ndim=2).value
PowerPILOT = wb_umts1.sheets['3g_cells_new'].range('H2:H20000').options(ndim=2).value
wb.activate(steal_focus=False)
sheet = wb.sheets['UMTS']
wb.sheets['UMTS'].range('A2:A10000').value = Cell_name12367
wb.sheets['UMTS'].range('B2:B10000').value = СellID145
wb.sheets['UMTS'].range('C2:C10000').value = cell_nam345
wb.sheets['UMTS'].range('F2:F10000').value = Downlink_UARFCN
wb.sheets['UMTS'].range('I2:I10000').value = Time_Offset
wb.sheets['UMTS'].range('G2:G10000').value = Power1
wb.sheets['UMTS'].range('H2:H10000').value = PowerPILOT
wb.sheets['UMTS'].range('K2:K10000').value = Frequency_Band
sheet.range('D2:D1000').value = 49100
sheet.range('E2:E1000').value = '301_ZTE'
wb.save("ЧТП.xlsx")
excel_file_path = "3g_cells_new.xlsx"
excel_file_path = "ЧТП.xlsx"
PSC = pd.read_excel("ЧТП.xlsx", sheet_name="UMTS")
PSC1 = pd.read_excel("3g_cells_new.xlsx", sheet_name="3g_cells_new")
df18 = PSC.merge(PSC1, left_on="Cell ID", right_on="Transmitter", how="left")
sheet.range("J1").options(index=False).value = df18["Primary scrambling code"]

#G2L
wb.sheets.add(name='G2L')
sheet = wb.sheets['G2L']
sheet.range('A1','H1').value = "Cell Name","Cell ID", "Neighbouring Cell Name","PCI","TAC","Band","UARFCN","PRB"
sheet.range('A1','H1').color = (253, 233, 217)
wb_g2l = xw.Book(r'G2L.xlsx')
wb_g2l.activate(steal_focus=False)
Cell_ID12345 = wb_g2l.sheets['G2L'].range('A2:A10000').options(ndim=2).value
Neighbouring_cell_name1234 = wb_g2l.sheets['G2L'].range('B2:B10000').options(ndim=2).value
wb.activate(steal_focus=False)
sheet = wb.sheets['G2L']
wb.sheets['G2L'].range('B2:B10000').value = Cell_ID12345
wb.sheets['G2L'].range('C2:C10000').value = Neighbouring_cell_name1234
wb.save("ЧТП.xlsx")
excel_file_path = "4g_cells_all.xlsx"
excel_file_path = "ЧТП.xlsx"
excel_file_path = "2g_transmitters_all.xlsx"
Cell_name12349 = pd.read_excel("ЧТП.xlsx", sheet_name="G2L")
Cell_name1234591 = pd.read_excel("2g_transmitters_all.xlsx", sheet_name="2g_transmitters_all")
df19 = Cell_name12349.merge(Cell_name1234591, left_on="Cell ID", right_on="Transmitter", how="left")
sheet.range("A1").options(index=False).value = df19["Cell_name"]
sheet.range('E2:E1000').value = 49100
PCI4 = pd.read_excel("ЧТП.xlsx", sheet_name="G2L")
PCI5 = pd.read_excel("4g_cells_all.xlsx", sheet_name="4g_cells_all")
df20 = PCI4.merge(PCI5, left_on="Neighbouring Cell Name", right_on="Transmitter", how="left")
sheet.range("D1").options(index=False).value = df20["Physical Cell ID"]
Band4 = pd.read_excel("ЧТП.xlsx", sheet_name="G2L")
Band5 = pd.read_excel("4g_cells_all.xlsx", sheet_name="4g_cells_all")
df21 = Band4.merge(Band5, left_on="Neighbouring Cell Name", right_on="Transmitter", how="left")
sheet.range("F1").options(index=False).value = df21["Frequency Band"]
EARFCN4 = pd.read_excel("ЧТП.xlsx", sheet_name="G2L")
EARFCN5 = pd.read_excel("4g_cells_all.xlsx", sheet_name="4g_cells_all")
df22 = EARFCN4.merge(EARFCN5, left_on="Neighbouring Cell Name", right_on="Transmitter", how="left")
sheet.range("G1").options(index=False).value = df22["Channel Number"]




#G2U
wb.sheets.add(name='G2U')
sheet = wb.sheets['G2U']
sheet.range('A1','H1').value = "2G Cell Name","2G Cell ID", "3G Cell Name","LAC","3G Cell ID","RNC","UARFCN","SC"
sheet.range('A1','H1').color = (253, 233, 217)
wb_g2u = xw.Book(r'G2U.xlsx')
wb_g2u.activate(steal_focus=False)
Cell_name176 = wb_g2u.sheets['G2U'].range('A2:A10000').options(ndim=2).value
Neigring_cell_ID = wb_g2u.sheets['G2U'].range('B2:B10000').options(ndim=2).value
wb.activate(steal_focus=False)
sheet = wb.sheets['L2U']
wb.sheets['G2U'].range('B2:B10000').value = Cell_name176
wb.sheets['G2U'].range('E2:E10000').value = Neigring_cell_ID
wb.activate(steal_focus=False)
sheet = wb.sheets['G2U']
wb.save("ЧТП.xlsx")
excel_file_path = "3g_transmitters_all.xlsx"
excel_file_path = "ЧТП.xlsx"
excel_file_path = "3g_cells_all.xlsx"
excel_file_path = "2g_transmitters_all.xlsx"
Cell_name9 = pd.read_excel("ЧТП.xlsx", sheet_name="G2U")
Cell_nameu9 = pd.read_excel("2g_transmitters_all.xlsx", sheet_name="2g_transmitters_all")
df23 = Cell_name9.merge(Cell_nameu9, left_on="2G Cell ID", right_on="Transmitter", how="left")
sheet.range("A1").options(index=False).value = df23["Cell_name"]
Cell_name91 = pd.read_excel("ЧТП.xlsx", sheet_name="G2U")
Cell_name912 = pd.read_excel("3g_transmitters_all.xlsx", sheet_name="3g_transmitters_all")
df24 = Cell_name91.merge(Cell_name912, left_on="3G Cell ID", right_on="Transmitter", how="left")
sheet.range("C1").options(index=False).value = df24["Cell_name"]
SC9 = pd.read_excel("ЧТП.xlsx", sheet_name="G2U")
SC91 = pd.read_excel("3g_cells_all.xlsx", sheet_name="3g_cells_all")
df25 = SC9.merge(SC91, left_on="3G Cell ID", right_on="Transmitter", how="left")
sheet.range("H1").options(index=False).value = df25["Primary scrambling code"]
UARFCND1 = pd.read_excel("ЧТП.xlsx", sheet_name="G2U")
UARFCND12 = pd.read_excel("3g_transmitters_all.xlsx", sheet_name="3g_transmitters_all")
df26 = UARFCND1.merge(UARFCND12, left_on="3G Cell ID", right_on="Transmitter", how="left")
sheet.range("G1").options(index=False).value = df26["Downlink_UARFCN"]
sheet.range('D2:D10000').value = 49100
sheet.range('F2:F10000').value = '301_ZTE'


#G2G
wb.sheets.add(name='G2G')
sheet = wb.sheets['G2G']
sheet.range('A1','J1').value = "Cell Name","Cell ID", "Neighbouring Cell Name","BCCH1","Neighbouring Cell ID","BSIC1","LAC","Standart","BSC","BSCR"
sheet.range('A1','J1').color = (253, 233, 217)
wb_g2g = xw.Book(r'G2G.xlsx')
wb_g2g.activate(steal_focus=False)
Cell_am12 = wb_g2g.sheets['G2G'].range('A2:A10000').options(ndim=2).value
Cell_am123  = wb_g2g.sheets['G2G'].range('B2:B10000').options(ndim=2).value
wb.activate(steal_focus=False)
sheet = wb.sheets['G2G']
wb.sheets['G2G'].range('B2:B10000').value = Cell_am12
wb.sheets['G2G'].range('E2:E10000').value = Cell_am123
wb.save("ЧТП.xlsx")
excel_file_path = "2g_transmitters_all.xlsx"
excel_file_path = "ЧТП.xlsx"
CL9 = pd.read_excel("ЧТП.xlsx", sheet_name="G2G")
CL91 = pd.read_excel("2g_transmitters_all.xlsx", sheet_name="2g_transmitters_all")
NCL9 = pd.read_excel("ЧТП.xlsx", sheet_name="G2G")
NCL91 = pd.read_excel("2g_transmitters_all.xlsx", sheet_name="2g_transmitters_all")
df30 = CL9.merge(CL91, left_on="Cell ID", right_on="Transmitter", how="left")
sheet.range("A1").options(index=False).value = df30["Cell_name"]
df31 = NCL9.merge(NCL91, left_on="Neighbouring Cell ID", right_on="Transmitter", how="left")
sheet.range("C1").options(index=False).value = df31["Cell_name"]
BCCH9 = pd.read_excel("ЧТП.xlsx", sheet_name="G2G")
BCCH91 = pd.read_excel("2g_transmitters_all.xlsx", sheet_name="2g_transmitters_all")
df32 = BCCH9.merge(BCCH91, left_on="Neighbouring Cell ID", right_on="Transmitter", how="left")
sheet.range("D1").options(index=False).value = df32["BCCH"]
BSIC9 = pd.read_excel("ЧТП.xlsx", sheet_name="G2G")
BSIC91 = pd.read_excel("2g_transmitters_all.xlsx", sheet_name="2g_transmitters_all")
df33 = BSIC9.merge(BSIC91, left_on="Neighbouring Cell ID", right_on="Transmitter", how="left")
sheet.range("F1").options(index=False).value = df33["BSIC"]
STANDART = pd.read_excel("ЧТП.xlsx", sheet_name="G2G")
STANDART1 = pd.read_excel("2g_transmitters_all.xlsx", sheet_name="2g_transmitters_all")
df34 = STANDART.merge(STANDART1, left_on="Neighbouring Cell ID", right_on="Transmitter", how="left")
sheet.range("H1").options(index=False).value = df34["Frequency Band"]
sheet.range('G2:G10000').value = 50100
sheet.range('I2:I10000').value = '201_ZTE'
sheet.range('J2:J10000').value = '201_ZTE'

#GSM
wb.sheets.add(name='GSM')
sheet = wb.sheets['GSM']
sheet.range('A1','I1').value = "Site ID","Cell ID", "BCCH","MaxPwr","TCH","BSIC","LAC","BSC","Cell Name"
sheet.range('A1','I1').color = (253, 233, 217)
wb_gsm = xw.Book(r'2g_transmitters_new.xlsx')
wb_gsm.activate(steal_focus=False)
Siteid9 = wb_gsm.sheets['2g_transmitters_new'].range('A2:A20000').options(ndim=2).value
cellid9 = wb_gsm.sheets['2g_transmitters_new'].range('B2:B20000').options(ndim=2).value
BCCH7 = wb_gsm.sheets['2g_transmitters_new'].range('U2:U20000').options(ndim=2).value
Cell_nme = wb_gsm.sheets['2g_transmitters_new'].range('BO2:BO20000').options(ndim=2).value
BSIC8 = wb_gsm.sheets['2g_transmitters_new'].range('AS2:AS20000').options(ndim=2).value
wb.activate(steal_focus=False)
sheet = wb.sheets['GSM']
wb.sheets['GSM'].range('A2:A10000').value = Siteid9
wb.sheets['GSM'].range('B2:B10000').value = cellid9
wb.sheets['GSM'].range('C2:C10000').value = BCCH7
wb.sheets['GSM'].range('I2:I10000').value = Cell_nme
wb.sheets['GSM'].range('F2:F10000').value = BSIC8
wb.save("ЧТП.xlsx")
sheet.range('G2:G1000').value = 50100
sheet.range('H2:H1000').value = '201_ZTE'
sheet.range('D2:D1000').value = 43
excel_file_path = "2g_trx_new.xlsx"
excel_file_path = "ЧТП.xlsx"
TRX = pd.read_excel("ЧТП.xlsx", sheet_name="GSM")
TRX1 = pd.read_excel("2g_trx_new.xlsx", sheet_name="2g_trx_new")
df32 = TRX.merge(TRX1, left_on="Cell ID", right_on="Transmitter", how="left")
sheet.range("E1").options(index=False).value = df32["Channels"]

for sheet in wb.sheets:
    if 'Лист1' in sheet.name:
        sheet.delete()
wb.save("ЧТП.xlsx")

