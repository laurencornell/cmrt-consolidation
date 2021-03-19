class Smelter(object):
    def __init__(self, id_number, metal, lookup, name, country, id_source, city, facility, contact_name, contact_email, next_steps, mine_names, mine_location, recycled_sources, comments):
        self.id_number = id_number
        self.metal = metal
        self.lookup = lookup
        self.name = name
        self.country = country
        self.id_source = id_source
        self.city = city
        self.facility = facility
        self.contact_name = contact_name
        self.contact_email = contact_email
        self.next_steps

    def __repr__(self):
        return "ID: " + self.id_number + ", metal: " + self.metal
