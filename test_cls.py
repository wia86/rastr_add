class root:
    r=1
    def __init__(self, rt):
        self.rt = rt

class one(root):
    o=1
    def __init__(self, col):
        super(one, self).__init__(col)
        self.col = col

r1= root(10)
o1= one(10)
p=1