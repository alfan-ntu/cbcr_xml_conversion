import xlsxwriter
import pdb
import constant


class XML_To_EXCEL():
    def __init__(self):
        self.tax_jurisdiction = []
        self.no_of_employee = []
        self.schedule_count = 0
        self.workbook = xlsxwriter.Workbook("cbcr_excel.xlsx")
        # initialized worksheet of table 1
        self.worksheet1 = self.workbook.add_worksheet("所得稅負營運表")
        self.worksheet1.set_column(0, \
                        constant.TABLE1_NUMBER_OF_COLUMN-1, \
                        constant.TABLE1_COLUMN_WIDTH)
        self.header_format = self.workbook.add_format()
        self.header_format.set_align("center")
        self.header_format.set_align("vcenter")
        self.header_format.set_text_wrap()
        # initialized worksheet of table 2
        self.worksheet2 = self.workbook.add_worksheet("成員集合名單")
        self.worksheet2.set_column(0, \
                        constant.TABLE2_NUMBER_OF_COLUMN-1, \
                        constant.TABLE2_COLUMN_WIDTH)
        self.worksheet2.set_column(1, 1, constant.TABLE2_COLUMN_WIDTH+35)
        self.corp_name_format = self.workbook.add_format()
        self.corp_name_format.set_align("left")
        self.corp_name_format.set_text_wrap()
        self.tax_jurisdiction_format = self.workbook.add_format()
        self.tax_jurisdiction_format.set_align("vcenter")
        self.bz_activity_format = self.workbook.add_format()
        self.bz_activity_format.set_align("center")
        self.bz_activity_format.set_align("vcenter")

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
        self.worksheet1.write(row, col,    "租稅管轄區", self.header_format)
        self.worksheet1.write(row, col+1,  "非關係人收入", self.header_format)
        self.worksheet1.write(row, col+2,  "關係人收入", self.header_format)
        self.worksheet1.write(row, col+3,  "收入合計", self.header_format)
        self.worksheet1.write(row, col+4,  "所得稅前損益", self.header_format)
        self.worksheet1.write(row, col+5,  "已納所得稅", self.header_format)
        self.worksheet1.write(row, col+6,  "當期應付所得稅", self.header_format)
        self.worksheet1.write(row, col+7,  "實收資本額", self.header_format)
        self.worksheet1.write(row, col+8,  "累積盈餘", self.header_format)
        self.worksheet1.write(row, col+9,  "員工人數", self.header_format)
        self.worksheet1.write(row, col+10, "有形資產", self.header_format)
        for row in range(0, self.schedule_count):
            self.worksheet1.write(row+1, col, self.tax_jurisdiction[row])
            self.worksheet1.write_number(row+1, col+1, int(self.unrelated_party_revenue[row]))
            self.worksheet1.write_number(row+1, col+2, int(self.related_party_revenue[row]))
            self.worksheet1.write_number(row+1, col+3, int(self.total_revenue[row]))
            self.worksheet1.write_number(row+1, col+4, int(self.profit_before_tax[row]))
            self.worksheet1.write_number(row+1, col+5, int(self.income_tax_paid[row]))
            self.worksheet1.write_number(row+1, col+6, int(self.income_tax_accrued[row]))
            self.worksheet1.write_number(row+1, col+7, int(self.stated_capital[row]))
            self.worksheet1.write_number(row+1, col+8, int(self.accumulated_earning[row]))
            self.worksheet1.write_number(row+1, col+9, int(self.no_of_employee[row]))
            self.worksheet1.write_number(row+1, col+10, int(self.tangible_assets[row]))
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
        print("Number of entity:", self.entity_count)

    def display_table2(self):
            i = 0
            for i in range(0, self.entity_count):
                print("Entity:", i,
                    "\n\t=>tax_jurisdiction:", self.tax_jurisdiction[i], \
                    "\n\t=>constituent_entity:", self.constituent_entity[i], \
                    "\n\t=>org_jurisdiction:", self.org_jurisdiction[i], \
                    "\n\t=>cbc501_activity:", self.cbc501_activity[i], \
                    "\n\t=>cbc502_activity:", self.cbc502_activity[i], \
                    "\n\t=>cbc503_activity:", self.cbc503_activity[i], \
                    "\n\t=>cbc504_activity:", self.cbc504_activity[i], \
                    "\n\t=>cbc505_activity:", self.cbc505_activity[i], \
                    "\n\t=>cbc506_activity:", self.cbc506_activity[i], \
                    "\n\t=>cbc507_activity:", self.cbc507_activity[i], \
                    "\n\t=>cbc508_activity:", self.cbc508_activity[i], \
                    "\n\t=>cbc509_activity:", self.cbc509_activity[i], \
                    "\n\t=>cbc510_activity:", self.cbc510_activity[i], \
                    "\n\t=>cbc511_activity:", self.cbc511_activity[i], \
                    "\n\t=>cbc512_activity:", self.cbc512_activity[i], \
                    "\n\t=>cbc513_activity:", self.cbc513_activity[i])

    def excel_output_table2(self):
        row = 0; col = 0
        self.worksheet2.write(row, col,    "租稅管轄區", self.header_format)
        self.worksheet2.write(row, col+1,  "國別報告成員", self.header_format)
        self.worksheet2.write(row, col+2,  "成員設立地", self.header_format)
        self.worksheet2.write(row, col+3,  "研究發展", self.header_format)
        self.worksheet2.write(row, col+4,  "智產權管理", self.header_format)
        self.worksheet2.write(row, col+5,  "採購", self.header_format)
        self.worksheet2.write(row, col+6,  "製造或生產", self.header_format)
        self.worksheet2.write(row, col+7,  "銷售行銷", self.header_format)
        self.worksheet2.write(row, col+8,  "行政管理或支援", self.header_format)
        self.worksheet2.write(row, col+9,  "對非關係人服務", self.header_format)
        self.worksheet2.write(row, col+10, "內部融資", self.header_format)
        self.worksheet2.write(row, col+11, "金融服務", self.header_format)
        self.worksheet2.write(row, col+12, "保險", self.header_format)
        self.worksheet2.write(row, col+13, "股份持有", self.header_format)
        self.worksheet2.write(row, col+14, "停業", self.header_format)
        self.worksheet2.write(row, col+15, "其他", self.header_format)
        for row in range(0, self.entity_count):
            self.worksheet2.write(row+1, col, self.tax_jurisdiction[row], self.tax_jurisdiction_format)
            self.worksheet2.write(row+1, col+1, self.constituent_entity[row], self.corp_name_format)
            self.worksheet2.write(row+1, col+2, self.org_jurisdiction[row])
            self.worksheet2.write(row+1, col+3, self.cbc501_activity[row], self.bz_activity_format)
            self.worksheet2.write(row+1, col+4, self.cbc502_activity[row], self.bz_activity_format)
            self.worksheet2.write(row+1, col+5, self.cbc503_activity[row], self.bz_activity_format)
            self.worksheet2.write(row+1, col+6, self.cbc504_activity[row], self.bz_activity_format)
            self.worksheet2.write(row+1, col+7, self.cbc505_activity[row], self.bz_activity_format)
            self.worksheet2.write(row+1, col+8, self.cbc506_activity[row], self.bz_activity_format)
            self.worksheet2.write(row+1, col+9, self.cbc507_activity[row], self.bz_activity_format)
            self.worksheet2.write(row+1, col+10, self.cbc508_activity[row], self.bz_activity_format)
            self.worksheet2.write(row+1, col+11, self.cbc509_activity[row], self.bz_activity_format)
            self.worksheet2.write(row+1, col+12, self.cbc510_activity[row], self.bz_activity_format)
            self.worksheet2.write(row+1, col+13, self.cbc511_activity[row], self.bz_activity_format)
            self.worksheet2.write(row+1, col+14, self.cbc512_activity[row], self.bz_activity_format)
            self.worksheet2.write(row+1, col+15, self.cbc513_activity[row], self.bz_activity_format)

    def close_excel_workbook(self):
        self.workbook.close()
