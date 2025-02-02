from typing import List, Optional, Tuple, Any, Set, Callable, Iterable

class DependencyGraph:

    def __init__(self):
        self.outgoing = {}
        self.ingoing = {}
    
    def outgoing_get(self, sheet_name, location):
        return self.outgoing[sheet_name.lower()][location.lower()]

    def ingoing_get(self, sheet_name, location):
        return self.ingoing[sheet_name][location]
    
    # TODO: updating dep graph when setting cell contents, calls both 
    # outgoing and ingoing

    def outgoing_set(self, sheet_name, location, outgoing_arr):
        if sheet_name not in self.outgoing:
            self.outgoing[sheet_name] = {}
        sheet_outgoing = self.outgoing[sheet_name]
        sheet_outgoing[location] = outgoing_arr
    
    def outgoing_add(self, sheet_name_1, loc_1, sheet_name_2, loc_2):
        if sheet_name_1 not in self.outgoing:
            self.outgoing[sheet_name_1] = {}
                
        sheet_outgoing = self.outgoing[sheet_name_1]
        if loc_1 not in sheet_outgoing:
            sheet_outgoing[loc_1] = []
        
        sheet_outgoing[loc_1].append((sheet_name_2, loc_2))

    def ingoing_add(self, sheet_name_1, loc_1, sheet_name_2, loc_2):
        if sheet_name_1 not in self.ingoing:
            self.ingoing[sheet_name_1] = {}
                
        sheet_ingoing = self.ingoing[sheet_name_1]
        if loc_1 not in sheet_ingoing:
            sheet_ingoing[loc_1] = []
        
        sheet_ingoing[loc_1].append((sheet_name_2, loc_2))

    def outgoing_remove(self, sheet_name_1, loc_1, sheet_name_2, loc_2):
        pass

    def ingoing_remove(self, sheet_name_1, loc_1, sheet_name_2, loc_2):
        
        pass
