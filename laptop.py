"""
Unlike regular methods, class methods don’t take the current instance, self, as an argument. Instead, they take the class itself, which is commonly passed in as the cls argument. Using cls to name this argument is a popular convention in the Python community.
"""


class Laptop:
    # Initialization of th constructor method.
    def __init__(self, serial_num, pid='', country="#AK8", title='', status='INFOSYSEON_SE_STAGING', qty=1, location='', price='0,01') -> object:
        self.Seriennummer = serial_num
        self.Artikel = pid + country
        self.Bezeichnung = title
        self.Menge = qty
        self.Ziellager = status  # INFOSYSEON_SE_STAGING
        self.Lagerot = location
        self.Preis = price

    @classmethod
    def customize_laptop(cls, pid, country, title, serial_num, status):
        return cls(serial_num, pid, country, title, status)

    def display_laptop(self):
        laptop_details = [self.Artikel, self.Bezeichnung, self.Menge, self.Seriennummer, self.Ziellager, self.Lagerot, self.Preis]
        return laptop_details