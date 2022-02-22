import win32com.client
import sys

rastr = win32com.client.Dispatch("Astra.Rastr")

rastr.Load(1, r'I:\rastr_add\test\2026 зим макс (0°C).rg2',
           r'I:\rastr_add\rastr example\pattern\режим.rg2')  # загрузить режим
# rastr.Load(1, r'I:\rastr_add\rastr example\graphics.grf',
#            r'I:\rastr_add\rastr example\pattern\графика.grf')  # загрузить режим

tables = rastr.tables
tables_find = rastr.Tables.Find("node")

node = rastr.Tables("node")
node_find = node.cols.Find("uhom")
node_name = node.cols.item("name").ZS(0)
node_name = node.Cols.item("ny").ZN(0)
node.cols.Item("ny").SetZ(0, 10)  # ndx=0 val=10
node.AddRow()  # добавить строку в конце таблицы
SelString = node.SelString(0)
node_size = node.size
vetv_size = rastr.Tables("vetv").size
BDO = node.writesafearray("ny,name", "000")
node.setsel("ny<100")
node.SetSel("ny=20001")
kod_calc = rastr.rgm("")
kod_calc = rastr.rgm("")
kod_calc = rastr.rgm("p")
kod_calc = rastr.rgm("p")
kod_calc = rastr.rgm("p")
kod_calc = rastr.rgm("p")
rastr.printp("txt")

rastr.SendChangeData(0, "", "", 0)

print("конец")

node.cols('name').Prop(1)  # тип поля 0 целый, 1 вещ, 2 строка, 3 переключатель, 4 перечисление, 5 рис, 6 цвет и тд
