import win32com.client
# Взять из исходной модели поля ny и arm_num
# Изменить в корректируемой модели ny на ny исходной модели c тем же arm_num
# Задание
pattern = r'I:\rastr_add\test\pattern\динамика.rst'
path_id = r'I:\ОЭС Урала ТЭ\!КПР ХМАО ЯНАО ТО\ткз\модели арм\перспектива 2022-2027 v4\temp\2022 ТЭ v21 заполярье ид ny.rst'
path_modify = r'I:\ОЭС Урала ТЭ\!КПР ХМАО ЯНАО ТО\ткз\модели арм\перспектива 2022-2027 v4\модель 2022 исходная.rst'
path_modify_save = r'I:\ОЭС Урала ТЭ\!КПР ХМАО ЯНАО ТО\ткз\модели арм\перспектива 2022-2027 v4\модель 2022 исходная v2.rst'

# Программа
rastr_id = win32com.client.Dispatch("Astra.Rastr")
rastr_id.Load(1, path_id, pattern) 
rastr_modify = win32com.client.Dispatch("Astra.Rastr")
rastr_modify.Load(1, path_modify, pattern)  

node_id = rastr_id.Tables("node")
# Словарь {arm_num: ny}
ny_id = {x[1].strip(): x[0] for x in node_id.writesafearray("ny,arm_num", "000")}

node_modify = rastr_modify.Tables("node")
rastr_modify.RenumWP = True  # включить ссылки, отключить
ny_modify_available = set((i[0] for i in node_modify.writesafearray("ny", "000")))

node_modify.setsel("")
i = node_modify.FindNextSel(-1)
while i > -1:
    arm_num = node_modify.cols.item('arm_num').Z(i).strip()
    if arm_num:
        if arm_num in ny_id:
            if ny_id[arm_num] not in ny_modify_available:
                ny_modify = node_modify.cols.item('ny').Z(i)
                print(f'Узел {ny_modify} изменен на {ny_id[arm_num]} {arm_num=}')
                node_modify.cols.item('ny').SetZ(i, ny_id[arm_num])
    i = node_modify.FindNextSel(i)

rastr_modify.save(path_modify_save, pattern)




