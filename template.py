
class templateByBU:
    def __init__(self, BU='', templates=[]):
        self.BU = BU
        self.templates = templates


class template:
    def __init__(self):
        self.BU = ''
        self.name = ''
        self.items = []
        self.subtotal = 0
        self.dates = {}

    def calcSubtotal(self):
        total = sum(i.ExtendedNetPrice for i in self.items)
        self.subtotal = total
        return total

    def appendItem(self,lineNum, partNum, desc, unitList, qty, disc):
        item = templateItem(lineNum, partNum, desc, unitList, qty, disc)

        self.items.append(item)


class templateItem:
    def __init__(self, lineNum, partNum, desc, unitList, qty, disc):
        self.LineNumber = lineNum
        self.PartNumber = partNum
        self.Description = desc
        self.UnitListPrice = unitList
        self.Qty = qty
        self.Disc = disc
        self.UnitNetPrice = self.calcUnitNetPrice()
        self.ExtendedNetPrice = self.calcExtNetPrice()

    def calcUnitNetPrice(self):
        unitNet = 0
        if self.UnitListPrice > 0:
            unitNet = float(self.UnitListPrice) * (1 - float(self.Disc)/100)

        return round(unitNet,2)

    def calcExtNetPrice(self):
        extNet = self.calcUnitNetPrice() * self.Qty

        return round(extNet, 2)


