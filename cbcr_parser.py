import biz_activity_code
import xml.etree.ElementTree as ET
import sys
import pdb
import xlsx_writer


class cbcr_parser():
    def __init__(self):

        tree = ET.parse('cbcr_gcp.xml')
        self.root = tree.getroot()
        print(self.root.tag, self.root.attrib)
        self.initialize_table1()
        self.parser_for_table1()

#        self.initialize_table2()
#        self.parser_for_table2()

    def initialize_table1(self):
        # initializing column lists of table1
        self.spread_sheet_table1 = []
        self.tax_jurisdiction = []
        self.unrelated_party_revenue = []
        self.related_party_revenue = []
        self.total_revenue = []
        self.profit_before_tax = []
        self.income_tax_paid = []
        self.income_tax_accrued = []
        self.stated_capital = []
        self.accumulated_earning = []
        self.no_of_employee = []
        self.tangible_assets = []

    def initialize_table2(self):
        # initializing column lists of table2
        self.spread_sheet_table2 = []
        self.tax_jurisdiction = []
        self.constituent_entity = []
        self.org_jurisdiction = []          # jurisdiction other than residence
        self.cbc501_activity = []           # research and development
        self.cbc502_activity = []           # holding or managing intellectual property
        self.cbc503_activity = []           # purchasing or procurement
        self.cbc504_activity = []           # manufacturing or production
        self.cbc505_activity = []           # Sales, Marketing or Distribution
        self.cbc506_activity = []           # Administrative, Management or Support Services
        self.cbc507_activity = []           # Provision of Services to unrelated parties
        self.cbc508_activity = []           # Internal Group Finance
        self.cbc509_activity = []           # Regulated Financial Services
        self.cbc510_activity = []           # Insurance
        self.cbc511_activity = []           # Holding shares or other equity instruments
        self.cbc512_activity = []           # Dormant
        self.cbc513_activity = []           # Other

    def parser_for_table1(self):
        self.schedule_count = 0
        for child in self.root:
            for schedule in child.findall('{http://www.irs.gov/efile}IRS8975ScheduleA'):
                self.schedule_count += 1
                irs8975attrib = schedule.attrib
                print("Schedule:", irs8975attrib)
                for jurisdiction in schedule.findall('{http://www.irs.gov/efile}TaxJurisdictionCountryCd'):
                    print("\tcntryCode:", jurisdiction.text)
                    self.tax_jurisdiction.append(jurisdiction.text)
                for unrelatedRevenue in schedule.findall('{http://www.irs.gov/efile}UnrelatedRevenueAmt'):
                    print("\tunrelatedRevenueAmt:", unrelatedRevenue.text)
                    self.unrelated_party_revenue.append(unrelatedRevenue.text)
                for relatedRevenue in schedule.findall('{http://www.irs.gov/efile}RelatedRevenueAmt'):
                    print("\trelatedRevenueAmt:", relatedRevenue.text)
                    self.related_party_revenue.append(relatedRevenue.text)
                for totalRevenue in schedule.findall('{http://www.irs.gov/efile}TotalRevenueAmt'):
                    print("\ttotalRevenueAmt:", totalRevenue.text)
                    self.total_revenue.append(totalRevenue.text)
                for profitBeforeTax in schedule.findall('{http://www.irs.gov/efile}ProfitOrLossAmt'):
                    print("\tprofitLossAmt:", profitBeforeTax.text)
                    self.profit_before_tax.append(profitBeforeTax.text)
                for incomeTaxPaid in schedule.findall('{http://www.irs.gov/efile}TotalTaxesPaidAmt'):
                    print("\tincomeTaxPaid:", incomeTaxPaid.text)
                    self.income_tax_paid.append(incomeTaxPaid.text)
                for incomeTaxAccrued in schedule.findall('{http://www.irs.gov/efile}TaxesAccruedAmt'):
                    print("\tincomeTaxAccrued:", incomeTaxAccrued.text)
                    self.income_tax_accrued.append(incomeTaxAccrued.text)
                for statedCapital in schedule.findall('{http://www.irs.gov/efile}CapitalAmt'):
                    print("\tstatedCapital:", statedCapital.text)
                    self.income_tax_accrued.append(statedCapital.text)
                for earningAmt in schedule.findall('{http://www.irs.gov/efile}EarningsAmt'):
                    print("\tearningAmt:", earningAmt.text)
                    self.accumulated_earning.append(earningAmt.text)
                for employeeCnt in schedule.findall('{http://www.irs.gov/efile}EmployeeCnt'):
                    print("\temployeeCnt:", employeeCnt.text)
                    self.no_of_employee.append(employeeCnt.text)
                for assetAmt in schedule.findall('{http://www.irs.gov/efile}AssetAmt'):
                    print("\tassetAmt:", assetAmt.text)
                    self.tangible_assets.append(assetAmt.text)

        print("Schedule count:", self.schedule_count)
        self.export_table1()

    def export_table1(self):
        xls2excel = xlsx_writer.XML_To_EXCEL()
        # pdb.set_trace()
        xls2excel.set_table1(self.tax_jurisdiction, self.no_of_employee)
        xls2excel.set_schedule_count(self.schedule_count)
        xls2excel.display_table1()

    def parser_for_table2(self):
        schedule_count = 0
        entity_count = 0
        row_count = 0
        for child in self.root:
            for schedule in child.findall('{http://www.irs.gov/efile}IRS8975ScheduleA'):
                schedule_count += 1
                for jurisdiction in schedule.findall('{http://www.irs.gov/efile}TaxJurisdictionCountryCd'):
                    print("\tcntryCode:", jurisdiction.text)
                    for const_entity_info in schedule.\
                        iter('{http://www.irs.gov/efile}ConstituentEntityInfoGrp'):
                        bzn = \
                            const_entity_info.find('{http://www.irs.gov/efile}BusinessName')
                        bzn_text = \
                            bzn.find('{http://www.irs.gov/efile}BusinessNameLine1Txt').text
                        print("\t\tbzName:", bzn_text)
#                        self.tax_jurisdiction.append(jurisdiction.text)
#                        self.constituent_entity.append(bzn_text)
                        entity_count += 1
                        # definition of CbcBizActivityType is listed in CbcXML_v1.0.1.xsd
                        # e.g. CBC501 is "Research and Development"
                        for bz_activity in const_entity_info.findall('{http://www.irs.gov/efile}CBCBusinessActivityCd'):
                            # print("\t\t\tactivityCode:", bz_activity.text)
                            # create a new row
                            self.init_bz_activity_lists()
                            self.tax_jurisdiction.append(jurisdiction.text)
                            self.constituent_entity.append(bzn_text)
                            self.org_jurisdiction.append("")
                            self.activity_switch(bz_activity.text)

#        pdb.set_trace()
        print("Entity count:", entity_count)

    def init_bz_activity_lists(self):
        self.cbc501_activity.append("")
        self.cbc502_activity.append("")
        self.cbc503_activity.append("")
        self.cbc504_activity.append("")
        self.cbc505_activity.append("")
        self.cbc506_activity.append("")
        self.cbc507_activity.append("")
        self.cbc508_activity.append("")
        self.cbc509_activity.append("")
        self.cbc510_activity.append("")
        self.cbc511_activity.append("")
        self.cbc512_activity.append("")
        self.cbc513_activity.append("")

    def activity_switch(self, argument):
        switcher = {
            "CBC501": self.cbc501,
            "CBC502": self.cbc502,
            "CBC503": self.cbc503,
            "CBC504": self.cbc504,
            "CBC505": self.cbc505,
            "CBC506": self.cbc506,
            "CBC507": self.cbc507,
            "CBC508": self.cbc508,
            "CBC509": self.cbc509,
            "CBC510": self.cbc510,
            "CBC511": self.cbc511,
            "CBC512": self.cbc512,
            "CBC513": self.cbc513
        }
        # print(switcher.get(argument, "Invalid month"))
        func = switcher.get(argument, lambda: "Invalid activity code")
        func()

    def cbc501(self):
        print("\t\tcbc501 of lenght:", len(self.cbc501_activity))
        self.cbc501_activity[len(self.cbc501_activity)-1]="V"

    def cbc502(self):
        print("\t\tcbc502 of lenght:", len(self.cbc502_activity))
        self.cbc502_activity[len(self.cbc502_activity)-1]="V"

    def cbc503(self):
        print("\t\tcbc503 of lenght:", len(self.cbc503_activity))
        self.cbc503_activity[len(self.cbc503_activity)-1]="V"

    def cbc504(self):
        print("\t\tcbc504 of lenght:", len(self.cbc504_activity))
        self.cbc504_activity[len(self.cbc504_activity)-1]="V"

    def cbc505(self):
        print("\t\tcbc505 of lenght:", len(self.cbc505_activity))
        self.cbc505_activity[len(self.cbc505_activity)-1]="V"

    def cbc506(self):
        print("\t\tcbc506 of lenght:", len(self.cbc506_activity))
        self.cbc506_activity[len(self.cbc506_activity)-1]="V"

    def cbc507(self):
        print("\t\tcbc507 of lenght:", len(self.cbc507_activity))
        self.cbc507_activity[len(self.cbc507_activity)-1]="V"

    def cbc508(self):
        print("\t\tcbc508 of lenght:", len(self.cbc508_activity))
        self.cbc508_activity[len(self.cbc508_activity)-1]="V"

    def cbc509(self):
        print("\t\tcbc509 of lenght:", len(self.cbc509_activity))
        self.cbc509_activity[len(self.cbc509_activity)-1]="V"

    def cbc510(self):
        print("\t\tcbc510 of lenght:", len(self.cbc510_activity))
        self.cbc510_activity[len(self.cbc510_activity)-1]="V"

    def cbc511(self):
        print("\t\tcbc511 of lenght:", len(self.cbc511_activity))
        self.cbc511_activity[len(self.cbc511_activity)-1]="V"

    def cbc512(self):
        print("\t\tcbc512 of lenght:", len(self.cbc512_activity))
        self.cbc512_activity[len(self.cbc512_activity)-1]="V"

    def cbc513(self):
        print("\t\tcbc513 of lenght:", len(self.cbc513_activity))
        self.cbc513_activity[len(self.cbc513_activity)-1]="V"


if __name__ == "__main__":
    main = cbcr_parser()
