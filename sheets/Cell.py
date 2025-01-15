class Cell:
    def __init__(self, contents=''):
        self.cell_type = None
        self.contents = ''
        self.value = ''
        self.outgoing = []
        self.ingoing = []
