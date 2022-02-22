import array

# массив с элементами одного типа
arr2 = array.array('i', (10, 21))  # https://docs.python.org/3/library/array.html?highlight=array%20array
print(arr2)

# неизменяемый от 0 до 255 включительно
arr = bytes((0, 1, 255))
print(arr)  # b'\x00\x01\xff'
print(arr[2])  # b'\x00\x01\xff'

# изменяемый
arr = bytearray((0, 1, 255))
print(arr)  # b'\x00\x01\xff'
arr[2] = 3
print(arr)  # b'\x00\x01\xff'
arr = bytes(arr)
print(type(arr))  # преобразовать в bytes


