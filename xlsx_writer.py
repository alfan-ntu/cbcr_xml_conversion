import xlsxwriter
import pdb


class XML_To_EXCEL():
    def __init__(self):
        print("Initializing XML_To_Excel instance!")
        self.tax_jurisdiction = []
        self.no_of_employee = []
        self.schedule_count = 0

    def set_table1(self, tax_jurisdiction, no_of_employee):
        self.tax_jurisdiction = tax_jurisdiction
        self.no_of_employee = no_of_employee

    def set_schedule_count(self, schedule_count):
        self.schedule_count = schedule_count
        print("Number of schedule count in set_schedule_count:", self.schedule_count)

    def display_table1(self):
        print("Number of schedule count:", self.schedule_count)
        i = 0
        for i in range(0, self.schedule_count):
            print(i, "=>tax_jurisdiction:", self.tax_jurisdiction[i],
            "\t=>constituent_entity", self.no_of_employee[i])
