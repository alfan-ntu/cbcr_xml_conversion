import xlsxwriter
import pdb


class XML_To_EXCEL():
    def __init__(self):
        self.tax_jurisdiction = []
        self.no_of_employee = []
        self.schedule_count = 0
        self.workbook = xlsxwriter.Workbook("cbcr_excel.xlsx")
        self.worksheet1 = self.workbook.add_worksheet("表1")
        self.worksheet2 = self.workbook.add_worksheet("表2")

    def set_table1(self, \
                    tax_jurisdiction, \
                    unrelated_party_revenue, \
                    related_party_revenue, \
                    total_revenue, \
                    profit_before_tax, \
                    income_tax_paid, \
                    income_tax_accrued, \
                    stated_capital, \
                    accumulated_earning, \
                    no_of_employee, \
                    tangible_assets):
        self.tax_jurisdiction = tax_jurisdiction
        self.unrelated_party_revenue = unrelated_party_revenue
        self.related_party_revenue = related_party_revenue
        self.total_revenue = total_revenue
        self.profit_before_tax = profit_before_tax
        self.income_tax_paid = income_tax_paid
        self.income_tax_accrued = income_tax_accrued
        self.stated_capital = stated_capital
        self.accumulated_earning = accumulated_earning
        self.no_of_employee = no_of_employee
        self.tangible_assets = tangible_assets

    def set_schedule_count(self, schedule_count):
        self.schedule_count = schedule_count
#        print("Number of schedule count in set_schedule_count:", self.schedule_count)

    def display_table1(self):
        i = 0
        for i in range(0, self.schedule_count):
            print("Sechedule:", i,
                    "\n\t=>tax_jurisdiction:", self.tax_jurisdiction[i],
                    "\n\t=>unrelated_party_revenue:", self.unrelated_party_revenue[i],
                    "\n\t=>related_party_revenue:", self.related_party_revenue[i],
                    "\n\t=>total_revenue:", self.total_revenue[i],
                    "\n\t=>profit_before_tax:", self.profit_before_tax[i],
                    "\n\t=>income_tax_paid:", self.income_tax_paid[i],
                    "\n\t=>income_tax_accrued:", self.income_tax_accrued[i],
                    "\n\t=>stated_capital:", self.stated_capital[i],
                    "\n\t=>accumulated_earning:", self.accumulated_earning[i],
                    "\n\t=>no_of_employee:", self.no_of_employee[i],
                    "\n\t=>tangible_assets:", self.tangible_assets[i])

    def excel_output_table1(self):
        row = 0; col = 0
        for row in range(0, self.schedule_count):
            self.worksheet1.write(row, col, self.tax_jurisdiction[row])
            self.worksheet1.write(row, col+1, self.unrelated_party_revenue[row])
            self.worksheet1.write(row, col+2, self.related_party_revenue[row])
            self.worksheet1.write(row, col+3, self.total_revenue[row])
            self.worksheet1.write(row, col+4, self.profit_before_tax[row])
            self.worksheet1.write(row, col+5, self.income_tax_paid[row])
            self.worksheet1.write(row, col+6, self.income_tax_accrued[row])
            self.worksheet1.write(row, col+7, self.stated_capital[row])
            self.worksheet1.write(row, col+8, self.accumulated_earning[row])
            self.worksheet1.write(row, col+9, self.no_of_employee[row])
            self.worksheet1.write(row, col+10, self.tangible_assets[row])
            row += 1

    def set_table2(self, \
                    tax_jurisdiction, \
                    constituent_entity, \
                    org_jurisdiction, \
                    cbc501_activity, \
                    cbc502_activity, \
                    cbc503_activity, \
                    cbc504_activity, \
                    cbc505_activity, \
                    cbc506_activity, \
                    cbc507_activity, \
                    cbc508_activity, \
                    cbc509_activity, \
                    cbc510_activity, \
                    cbc511_activity, \
                    cbc512_activity, \
                    cbc513_activity):
        self.tax_jurisdiction = tax_jurisdiction
        self.constituent_entity = constituent_entity
        self.org_jurisdiction = org_jurisdiction
        self.cbc501_activity = cbc501_activity
        self.cbc502_activity = cbc502_activity
        self.cbc503_activity = cbc503_activity
        self.cbc504_activity = cbc504_activity
        self.cbc505_activity = cbc505_activity
        self.cbc506_activity = cbc506_activity
        self.cbc507_activity = cbc507_activity
        self.cbc508_activity = cbc508_activity
        self.cbc509_activity = cbc509_activity
        self.cbc510_activity = cbc510_activity
        self.cbc511_activity = cbc511_activity
        self.cbc512_activity = cbc512_activity
        self.cbc513_activity = cbc513_activity

    def set_entity_count(self, entity_count):
        self.entity_count = entity_count

    def close_excel_workbook(self):
        self.workbook.close()
