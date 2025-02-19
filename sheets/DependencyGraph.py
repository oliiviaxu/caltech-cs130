
class DependencyGraph:

    def __init__(self):
        self.outgoing = {}
        self.ingoing = {}

    def add_sheet(self, sheet_name):
        sheet_name = sheet_name.lower()
        if (sheet_name not in self.outgoing):
            self.outgoing[sheet_name] = {}
        if (sheet_name not in self.ingoing):
            self.ingoing[sheet_name] = {}
    
    def outgoing_get(self, sheet_name, location):
        sheet_name = sheet_name.lower()
        location = location.lower()

        if (sheet_name not in self.outgoing or location not in self.outgoing[sheet_name]):
            return []
        return self.outgoing[sheet_name][location]

    def ingoing_get(self, sheet_name, location):
        sheet_name = sheet_name.lower()
        location = location.lower()

        if (sheet_name not in self.ingoing or location not in self.ingoing[sheet_name]):
            return []
        return self.ingoing[sheet_name][location]

    def outgoing_set(self, sheet_name, location, outgoing_arr):
        sheet_name = sheet_name.lower()
        location = location.lower()
        if sheet_name not in self.outgoing:
            self.outgoing[sheet_name] = {}
        sheet_outgoing = self.outgoing[sheet_name]
        sheet_outgoing[location] = outgoing_arr

    def outgoing_add(self, sheet_name_1, loc_1, sheet_name_2, loc_2):
        sheet_name_1 = sheet_name_1.lower()
        loc_1 = loc_1.lower()
        sheet_name_2 = sheet_name_2.lower()
        loc_2 = loc_2.lower()

        if sheet_name_1 not in self.outgoing:
            self.outgoing[sheet_name_1] = {}
                
        sheet_outgoing = self.outgoing[sheet_name_1]
        if loc_1 not in sheet_outgoing:
            sheet_outgoing[loc_1] = []
        
        sheet_outgoing[loc_1].append((sheet_name_2, loc_2))

    def ingoing_add(self, sheet_name_1, loc_1, sheet_name_2, loc_2):
        sheet_name_1 = sheet_name_1.lower()
        loc_1 = loc_1.lower()
        sheet_name_2 = sheet_name_2.lower()
        loc_2 = loc_2.lower()

        if sheet_name_1 not in self.ingoing:
            self.ingoing[sheet_name_1] = {}
                
        sheet_ingoing = self.ingoing[sheet_name_1]
        if loc_1 not in sheet_ingoing:
            sheet_ingoing[loc_1] = []
        
        sheet_ingoing[loc_1].append((sheet_name_2, loc_2))

    def outgoing_remove(self, sheet_name_1, loc_1, sheet_name_2, loc_2):
        sheet_name_1 = sheet_name_1.lower()
        loc_1 = loc_1.lower()
        sheet_name_2 = sheet_name_2.lower()
        loc_2 = loc_2.lower()

        self.outgoing[sheet_name_1][loc_1].remove((sheet_name_2, loc_2))

    def ingoing_remove(self, sheet_name_1, loc_1, sheet_name_2, loc_2):
        sheet_name_1 = sheet_name_1.lower()
        loc_1 = loc_1.lower()
        sheet_name_2 = sheet_name_2.lower()
        loc_2 = loc_2.lower()

        self.ingoing[sheet_name_1][loc_1].remove((sheet_name_2, loc_2))