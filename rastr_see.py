import win32com.client

rastr = win32com.client.Dispatch("Astra.Rastr")

rastr.Load(1, r'I:\rastr_add\rastr example\rastr file.rg2', r'I:\rastr_add\rastr example\pattern\режим.rg2')  # загрузить режим
rastr.Load(1, r'I:\rastr_add\rastr example\graphics.grf', r'I:\rastr_add\rastr example\pattern\графика.grf')  # загрузить режим

tables = rastr.tables
tables_find = rastr.Tables.Find("node")

node = rastr.Tables("node")
node_find = node.cols.Find("uhom")
node_name = node.cols.item("name").ZS(0)
node_name = node.Cols.Item("ny").ZN(0)
node.cols.Item("ny").SetZ(0, 10)  # ndx=0 val=10

SelString = node.SelString(0)
node_size = node.Size
vetv_size = rastr.Tables("vetv").size
node.setsel("ny<100")
node.SetSel("ny<1000")
kod_calc = rastr.rgm ("")
rastr.printp ("txt")

rastr.SendChangeData (0,"","",0 )
print("конец")



