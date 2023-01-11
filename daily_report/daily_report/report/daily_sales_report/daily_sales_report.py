# Copyright (c) 2022, abayomi.awosusi@sgatechsolutions.com and contributors
# For license information, please see license.txt

import calendar
import datetime
import json
import frappe
from frappe import _
from .gross_profit_generators import *

def execute(filters=None):
    return [], [], None, []

def validate_filters(filters):
    cmonth = filters.get("month")

    if not cmonth:
        frappe.throw(_("Please select a month to generate report."))
    
def get_conditions(filters):
    conditions = ""
    
    if filters.get("month"):
        long_month_name = filters.get("month")
        datetime_object = datetime.datetime.strptime(long_month_name, "%B")
        month_number = datetime_object.month
        conditions += " MONTH(si.posting_date) = " + str(month_number)

    if filters.get("year"):
        conditions += " and YEAR(si.posting_date) = %(year)s"	

    return conditions

def get_merged_dataongrossprofls(filters,si_list):    
    workdaysinmth = []
    long_month_name = filters.get("month")
    datetime_object = datetime.datetime.strptime(long_month_name, "%B")
    month_number = datetime_object.month
    cal= calendar.Calendar()

    lastwkdayincurrper = 0
    for x in cal.itermonthdays2(int(filters.get("year")), month_number):
        
        if (x[0] != 0):
            temparr = list((x[0],x[1],))
            workdaysinmth.append(temparr)
            lastwkdayincurrper = x[0]
    
    datamg = get_corrdataingplistwithcstcnt(si_list,workdaysinmth,lastwkdayincurrper,filters)

    return datamg

def get_corrdataingplistwithcstcnt(lstdata,workdaysinmth,lastwkdayincurrper,filters):
    data5 = []
    data6 = []
    cstcnt = []

    cstcnt0 = frappe.db.get_list("Cost Center",pluck='name',filters={'company': filters.get("company"),'is_group':0})
    # change the order of cost center this is customized for this client
    #specify order here 02, 03, 01, 06
    if filters.cost_center:
        cstcnt0.clear()
        ccc = filters.cost_center
        for cc in ccc:
            cstcnt0.append(cc)

    cstorder = ['02', '03', '06', '01']
    
    i = 0
    while(i<len(cstorder)):
        for cstr in cstcnt0:
            if (cstr.startswith(cstorder[i])):
                cstcnt.append(cstr)
        i+=1
        
    # if created cost centers increase
    if ((len(cstorder)<len(cstcnt0)) and (len(cstcnt)>0) ):
        for cstr2 in cstcnt0:
            cstfound = False
            for m in cstcnt:
                if (m==cstr2):
                    cstfound = True
            if (cstfound == False):
                 cstcnt.append(cstr2)         

    if (len(cstcnt)==0):
       cstcnt = cstcnt0 

    cclength = len(cstcnt)
    month_number = 0
    curryr = 0
    prevyr = 0
    cumsalesmtd1 = 0
    cumsalesmtd2 = 0
    grossprfmtd1 = 0.0
    grossprfmtd2 = 0.0
    grossprfmtdcum1 = 0.0
    grossprfmtdcum2 = 0.0
    
    grossprfamt1 = 0.0
    grossprfamt2 = 0.0

    cumsales3 = []
    cumnoofinv3 = []
    cumsalesmtd3 = []
    dategrossprf3 = []
    cumgrossavg3 = []

    dategrossprfamt3 = []

    cumgrossmtd3 = []
    daysoftxns3 = []
    cumgrossprfmtd3 = []
    grossprfmtd3 = [] 

    grossprfamt3 = []
    grossprfmtdcum3 = []

    i = 0
    while(i<cclength):
        cumsales3.append(0)
        cumnoofinv3.append(0)
        cumsalesmtd3.append(0)
        dategrossprf3.append(0.0)
        cumgrossavg3.append(0)

        cumgrossmtd3.append(0)
        daysoftxns3.append(0)
        cumgrossprfmtd3.append(0)
        grossprfmtd3.append(0.0)

        grossprfamt3.append(0)

        dategrossprfamt3.append(0)
        grossprfmtdcum3.append(0.0)

        i+= 1
     
    if filters.get("month"):
        long_month_name = filters.get("month")
        datetime_object = datetime.datetime.strptime(long_month_name, "%B")
        month_number = datetime_object.month
        curryr = int(filters.get("year"))
        prevyr = curryr -1

    for y in workdaysinmth:
        salesday = y[0]
        salesdate = datetime.datetime(curryr, month_number, salesday)
        salesdate2 = datetime.datetime(prevyr, month_number, salesday)
        salesdayinwords = salesdate.strftime("%A")
        salesdayinwords2 = salesdate2.strftime("%A")

        cumnoofinv1 = 0
        cumsales1 = 0
        cumgrossavg1 = 0
        dategrossprf = 0.0
        dategrossprfamt = 0.0
        cumnoofinv2 = 0
        cumsales2 = 0
        cumgrossavg2 = 0
        dategrossprf2 = 0.0
        dategrossprfamt2 = 0.0

        for i in range(cclength):
            cumsales3[i] = 0
            cumnoofinv3[i] = 0
            cumgrossavg3[i] = 0
            dategrossprf3[i]= 0.0
            dategrossprfamt3[i]= 0.0

        for row in lstdata:
            dayno1 = 1
            if ( (row["indent"]==0.0) and ((row["posting_date"]).strftime("%Y-%m-%d")==salesdate.strftime("%Y-%m-%d"))):
                cumnoofinv1 += 1
                cumsales1 += row["base_net_amount"]
                cumsalesmtd1 += row["base_net_amount"] 
                dategrossprf += row["gross_profit_percent"]

                dategrossprfamt += row["gross_profit"]
                grossprfmtdcum1 += row["gross_profit"]
                
                dayno_object1 = datetime.datetime.strptime(str(row["posting_date"]), "%Y-%m-%d")
                dayno_object = dayno_object1.strftime("%d")
                dayno1 = int(str.lstrip(dayno_object))

                for i in range(cclength):
                    if(cstcnt[i] == row["cost_center"]): 
                        cumnoofinv3[i] += 1
                        cumsales3[i] += row["base_net_amount"]
                        cumsalesmtd3[i] += row["base_net_amount"]
                        dategrossprf3[i] += row["gross_profit_percent"]

                        dategrossprfamt3[i] += row["gross_profit"]
                        grossprfmtdcum3[i] += row["gross_profit"]
                            

            dayno_object2 = datetime.datetime.strptime(str(row["posting_date"]), "%Y-%m-%d")
            dayno_object2 = dayno_object2.strftime("%d")
            dayno2 = int(str.lstrip(dayno_object2))
            prevdateyr = (row["posting_date"]).strftime("%Y")

            if ((row["indent"]==0.0) and (int(prevdateyr) == prevyr) and (dayno2 == salesday)):
                cumnoofinv2 += 1
                cumsales2 += row["base_net_amount"]
                cumsalesmtd2 += row["base_net_amount"]  
                dategrossprf2 += row["gross_profit_percent"]
            
                dategrossprfamt2 += row["gross_profit"]
                grossprfmtdcum2 += row["gross_profit"]

            #check for last lines that maybe left after loop and sum up
            elif ( (row["indent"]==0.0) and (int(prevdateyr) == prevyr) and (dayno1 == lastwkdayincurrper) and (dayno2 > dayno1)):
                cumnoofinv2 += 1
                cumsales2 += row["base_net_amount"]
                cumsalesmtd2 += row["base_net_amount"]  
                dategrossprf2 += row["gross_profit_percent"]

                dategrossprfamt2 += row["gross_profit"]
                grossprfmtdcum2 += row["gross_profit"]
    
        try:        
            cumgrossavg1 = (dategrossprfamt/cumsales1)*100 
        except ZeroDivisionError:
            cumgrossavg1 = 0

        try:        
            cumgrossavg2 = (dategrossprfamt2/cumsales2)*100 
        except ZeroDivisionError:
            cumgrossavg2 = 0     

        for i in range(cclength):
            cumgrossavg3[i] = 0
            if (dategrossprfamt3[i]!=0):
                cumgrossavg3[i] = (dategrossprfamt3[i]/cumsales3[i])*100
            
        try:    
            grossprfmtd1 = (grossprfmtdcum1/cumsalesmtd1)*100
        except ZeroDivisionError:
            grossprfmtd1 = grossprfmtd1
        try:
            grossprfmtd2 = (grossprfmtdcum2/cumsalesmtd2)*100
        except ZeroDivisionError:
            grossprfmtd2 = grossprfmtd2     

        data5.append({"date":salesdate,"day":salesdayinwords,
                      "noofinv":cumnoofinv1,"sales":cumsales1,"salesmtd":cumsalesmtd1,
                      "gross":cumgrossavg1,"grossmtd":grossprfmtd1,
                      "date2":salesdate2,"day2":salesdayinwords2, 
                      "noofinv2":cumnoofinv2,"sales2":cumsales2,"salesmtd2":cumsalesmtd2,
                      "gross2":cumgrossavg2,"grossmtd2":grossprfmtd2})
        
        cstdict = {"date":salesdate,"day":salesdayinwords} 
        for i in range(cclength):
            cstdict["noofinvcstcnt" + str(i)] = cumnoofinv3[i]
            cstdict["salescstcnt" + str(i)] = cumsales3[i]
            cstdict["salesmtdcstcnt" + str(i)] = cumsalesmtd3[i]
            cstdict["grosscstcnt" + str(i)] = cumgrossavg3[i]
            try:
                grossprfmtd3[i] = (grossprfmtdcum3[i]/cumsalesmtd3[i])*100
            except ZeroDivisionError:
                grossprfmtd3[i] = grossprfmtd3[i]
                
            if (cumgrossavg3[i]!=0.0) :
                cumgrossprfmtd3[i] += cumgrossavg3[i]
                daysoftxns3[i] += 1

            cstdict["grossmtdcstcnt" + str(i)] = grossprfmtd3[i]

        data6.append(cstdict)  
                
    datall=[data5,data6,cstcnt]

    return datall 


def get_columns(filters):
    long_month_name = filters.get("month")
    curyr = filters.get("year")
    prevyr = int(curyr) - 1
    datetime_object = datetime.datetime.strptime(long_month_name, "%B")
    month_short = datetime_object.strftime("%b")
    columns = [
        {"label": _("Date"), "fieldname": "date", "fieldtype": "Date", "width": 90},
        {"label": _("Day"), "fieldname": "day", "fieldtype": "String", "width": 90},
        {
            "label": _("# of Invoices"),
            "fieldname": "noofinv",
            "fieldtype": "Integer",
            "width": 80,
            "convertible": "qty",
        },
        {
            "label": _("Sales"),
            "fieldname": "sales",
            "fieldtype": "Currency",
            "options": "Company:company:default_currency",
            "convertible": "rate",
            "width": 160,
        },
        {
            "label": _("Sales MTD"),
            "fieldname": "salesmtd",
            "fieldtype": "Currency",
            "options": "Company:company:default_currency",
            "width": 160,
        },
        {
            "label": _("Gross %"),
            "fieldname": "gross",
            "fieldtype": "Float",
            "convertible": "qty",
            "width": 90,
        },
        {
            "label": _("Gross % MTD"),
            "fieldname": "grossmtd",
            "fieldtype": "Float",
            "convertible": "qty",
            "width": 90,
        },
        {"label": _("Date"), "fieldname": "date2", "fieldtype": "Date", "width": 90},
        {"label": _("Day"), "fieldname": "day2", "fieldtype": "String", "width": 90},
        {
            "label": _("# of Invoices(" + month_short + ", " + str(prevyr) + ")"),
            "fieldname": "noofinv2",
            "fieldtype": "Integer",
            "width": 80,
            "convertible": "qty",
        },
        {
            "label": _("Sales( " + month_short + ", " + str(prevyr) + ")"),
            "fieldname": "sales2",
            "fieldtype": "Currency",
            "options": "Company:company:default_currency",
            "convertible": "rate",
            "width": 160,
        },
        {
            "label": _("Sales MTD( "+ month_short + ", " + str(prevyr) + ")"),
            "fieldname": "salesmtd2",
            "fieldtype": "Currency",
            "options": "Company:company:default_currency",
            "convertible": "rate",
            "width": 160,
        },
        {
            "label": _("Gross % (" + month_short + ", " + str(prevyr) + ")"),
            "fieldname": "gross2",
            "fieldtype": "Float",
            "convertible": "qty",
            "width": 90,
        },
        {
            "label": _("Gross % (" + month_short + ", " + str(prevyr) + ")"),
            "fieldname": "grossmtd2",
            "fieldtype": "Float",
            "convertible": "qty",
            "width": 90,
        },
    ]

    return columns	

@frappe.whitelist()
def get_daily_report_record(report_name,filters):	
    # Skipping total row for tree-view reports
    skip_total_row = 0
    filterDt= json.loads(filters)	
    filters = frappe._dict(filterDt or {})	
    
        
    if not filters:
        return [], [], None, []

    validate_filters(filters)

    gross_profit_data1 = GrossProfitGenerator(filters)
    #I had to create a replica of the class - if same class is used and the sytem merges the list 15 percent of gross 
    #margin calculation gets wrongly merged - so i had to create a seperate class for the prev yr and have 
    # the result merged at the end of gross nmargin calculation - reports take longer but results are accurate
    gross_profit_data2 = GrossProfitGenerator2(filters)
    gross_profit_data = []
    gross_profit_data.extend(gross_profit_data1.si_list)
    gross_profit_data.extend(gross_profit_data2.si_list)
    columns = get_columns(filters)
    data = get_merged_dataongrossprofls(filters,gross_profit_data)
    if not data:
        return [], [], None, []
    
    return columns, data        