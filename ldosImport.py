class ldosSummary:
    def __init__(self, BU=''):
        self.BU = BU
        self.sItems = []

    def appendItems(self, sItem):
        self.sItems.append(sItem)


class summaryItem:
    def __init__(self, name='', qty=0, ldosDatesDict={}):
        self.name = name
        self.qty = qty
        self.dates = ldosDatesDict
