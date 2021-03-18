class Smelter(object):
    def __init__(self, id_number, metal):
        self.id_number = id_number
        self.metal = metal

    def __repr__(self):
        return "ID: " + self.id_number + ", metal: " + self.metal
