// Copyright (c) 2022, abayomi.awosusi@sgatechsolutions.com and contributors
// For license information, please see license.txt
/* eslint-disable */

// Determines whether or not the gross profit MTD should be displayed. Adjusts the spreadsheet to account for the columns absence/presence.
_use_gross_profit_mtd = true;

const _months = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];

frappe.query_reports["Daily Sales Report"] = {
	"filters": [
		{
            fieldname: 'company',
            label: __('Company'),
            fieldtype: 'Link',
            options: 'Company',
            default: frappe.defaults.get_user_default('company')
        },
		{
			"fieldname": "month",
			"fieldtype": "Select",
			"label": "Month",
			"options": _months,
			"default": _months[new Date().getMonth()],
			"mandatory": 0,
			"wildcard_filter": 0
		   },
		   {
			"fieldname": "year",
			"fieldtype": "Select",
			"label": "Year",
			"options": getyears(),
			"default": new Date().getFullYear(),
			"mandatory": 0,
			"wildcard_filter": 0
		   },
		   {
			   fieldname: "cost_center",
			   label: __("Cost Center"),
			   fieldtype: "MultiSelectList",
			   options: "Cost Center",
			   reqd:0,				
			   get_data: function(txt) {				
				   return frappe.db.get_link_options("Cost Center", txt);
			   }
		   }
	],
	onload: function(report) {		
		report.page.add_inner_button(__("Export Report"), function() {			
			debugger
			let filters = report.get_values();	
			
			var start_time = new Date();
			frappe.show_progress('Generating Report...', 0, 1, `Gathering data...Please wait.`);

			frappe.call({
				method: 'daily_report.daily_report.report.daily_sales_report.daily_sales_report.get_daily_report_record',			
				args: {					
					report_name: report.report_name,
					filters: filters
				},
				callback: function(r) {		
					$(".report-wrapper").html("");					
					$(".justify-center").remove()
					if(r.message[1] != ""){
						dynamic_exportcontent(r.message,filters.company,filters.month,filters.year)
						
						// Total execution time for display
						var total_time = ((new Date()).getTime() - start_time.getTime()) / 1000;
						var minutes = Math.floor(total_time / 60);
						var seconds = Math.round(total_time - (minutes * 60));
						var display_time = (minutes > 0 ? `${minutes} minute${minutes > 1 ? 's' : ''} and ` : '') + `${seconds} second${seconds != 1 ? 's' : ''}`;

						frappe.show_progress('Generating Report...', 1, 1, `Completed in ${display_time}`);
					} else {
						alert("No record found.")
					}			
				}
			});				
		});	
	},	
};

function dynamic_exportcontent(cnt_list,company,fmonth,fyear){	
	var dynhtml="";
	dynhtml='<div id="dvData">';
	var totlcnt=[];
	var $crntid="exprtid_1";
	totlcnt[0]="#"+$crntid;

	var titlelist = cnt_list[0]
	var datalist = cnt_list[1][0]
	var datalist2 = cnt_list[1][1]
	var datacostcntlist3 = cnt_list[1][2]

	//==================================================
	// TABLE - CONSOLIDATED DAILY SALES
	//==================================================

	// Generate the table title
	dynhtml+='<table style= "font-family: Arial; font-size: 10pt;" id='+$crntid+'>';
	dynhtml+='<caption style="text-align: left;"><span style="font-weight: bold;text-align: left;font-family: Arial; font-size: 10pt;">Sales Statistics For '+company+'</br></span><span style="font-family: Arial; font-size: 10pt; font-weight: normal;text-align: left;">'+ 'Month: &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;' + fmonth + '-' + fyear +'</span><caption>';	
	dynhtml+='<tr><td>&nbsp;</td></tr>';
    dynhtml+='<tr>';
	dynhtml+=`<th style="text-align: center;border: 1px solid #89898d;font-weight: bold;font-style: italic;font-family: Arial; font-size: 10pt;" colspan="${_use_gross_profit_mtd ? '7' : '6'}"> ` + "&nbsp;" + fmonth + ' ' + fyear + '</span></th>';
    dynhtml+=`<th style="text-align: center;border: 1px solid #89898d;font-weight: bold;font-style: italic;font-family: Arial; font-size: 10pt;" colspan="${_use_gross_profit_mtd ? '5' : '4'}"> `+ "&nbsp;" + fmonth + ' ' + (parseInt(fyear) - 1).toString()+'</th>';
	dynhtml+='</tr>';

	// Generate the table headers for each cost center
	dynhtml+='<tr>';
	for(var cnt=0; cnt < titlelist.length; cnt++) 
	{					
		var colmnth= titlelist[cnt].label.toString();
		var colfldname = titlelist[cnt].fieldname.toString();
		if (colfldname=='noofinv2') {
			colmnth= "# of Inv.'s"
			dynhtml+='<td width="110" style="text-align: center;border: 1px solid #89898d;font-weight: bold;font-family: Arial; font-size: 10pt;">'+(colmnth).toString()+'</td>';
		}
		else if (colfldname=='sales2') {
			colmnth= 'Sales'
			dynhtml+='<td width="150" style="text-align: center;border: 1px solid #89898d;font-weight: bold;font-family: Arial; font-size: 10pt;">'+(colmnth).toString()+'</td>';
		}
		else if (colfldname=='salesmtd2') {
			colmnth= 'Sales MTD'
			dynhtml+='<td width="150" style="text-align: center;border: 1px solid #89898d;font-weight: bold;font-family: Arial; font-size: 10pt;">'+(colmnth).toString()+'</td>';
		}
		else if (colfldname=='gross2') {
			colmnth= 'Gross %'
			dynhtml+='<td width="100" style="text-align: center;border: 1px solid #89898d;font-weight: bold;font-family: Arial; font-size: 10pt;">'+(colmnth).toString()+'</td>';
		}
		else if (colfldname=='grossmtd2' && _use_gross_profit_mtd) {
			colmnth= 'Gross % MTD'
			dynhtml+='<td width="150" style="text-align: center;border: 1px solid #89898d;font-weight: bold;font-family: Arial; font-size: 10pt;">'+(colmnth).toString()+'</td>';
		}
		else if (colfldname=='sales') {
			dynhtml+='<td width="150" style="text-align: center;border: 1px solid #89898d;font-weight: bold;font-family: Arial; font-size: 10pt;">'+(colmnth).toString()+'</td>';
		}
		else if (colfldname=='salesmtd') {
			dynhtml+='<td width="150" style="text-align: center;border: 1px solid #89898d;font-weight: bold;font-family: Arial; font-size: 10pt;">'+(colmnth).toString()+'</td>';
		}
		else if (colfldname=='noofinv') {
			colmnth= "# of Inv.'s"
			dynhtml+='<td width="110" style="text-align: center;border: 1px solid #89898d;font-weight: bold;font-family: Arial; font-size: 10pt;">'+(colmnth).toString()+'</td>';
		}
		else if (colfldname=='gross') {
			dynhtml+='<td width="100" style="text-align: center;border: 1px solid #89898d;font-weight: bold;font-family: Arial; font-size: 10pt;">'+(colmnth).toString()+'</td>';
		}
		else if (colfldname=='grossmtd' && _use_gross_profit_mtd) {
			dynhtml+='<td width="150" style="text-align: center;border: 1px solid #89898d;font-weight: bold;font-family: Arial; font-size: 10pt;">'+(colmnth).toString()+'</td>';
		}
		else if (colfldname=='date') {
			dynhtml+='<td width="110" style="text-align: center;border: 1px solid #89898d;font-weight: bold;font-family: Arial; font-size: 10pt;font-style: italic;">'+(colmnth).toString()+'</td>';
		}
		else if (colfldname=='day')
		{
			dynhtml+='<td width="120" style="text-align: center;border: 1px solid #89898d;font-weight: bold;font-family: Arial; font-size: 10pt;">'+(colmnth).toString()+'</td>';
		} 
	}
	dynhtml+='</tr>';

	// Populate the table rows
	dynhtml+=row_celldynFunc(datalist);

	// Calculate previous year cumulative values
	var prevyrsalessum = 0
	var grossmarlast = 0.0

	for(var index = 0; index < datalist.length; index++) {	
		prevyrsalessum += datalist[index].sales2
		grossmarlast = 	datalist[index].grossmtd2
	}

	let dollarCAD = Intl.NumberFormat("en-CA", {
		style: "currency",
		currency: "CAD",
		useGrouping: true,
	});

	// Generate year totals at the bottom of the table
	var grossmar = formatAsPercent(grossmarlast)
	dynhtml+=`<tr><td /><td /><td /><td /><td />${_use_gross_profit_mtd ? '<td />' : ''}<td style="border-right: 1px solid #89898d;" /><td /><td /><td />${_use_gross_profit_mtd ? '<td />' : ''}<td style="border-right: 1px solid #89898d;"></td></tr>`;
	dynhtml+=`<tr><td style="border-bottom: 1px solid #89898d;" /><td style="border-bottom: 1px solid #89898d;" /><td style="border-bottom: 1px solid #89898d;" /><td style="border-bottom: 1px solid #89898d;" /><td style="border-bottom: 1px solid #89898d;" />${_use_gross_profit_mtd ? '<td style="border-bottom: 1px solid #89898d;" />' : ''}<td style="border-bottom: 1px solid #89898d;border-right: 1px solid #89898d;" /><td style="border-bottom: 1px solid #89898d;" /><td style="border-bottom: 1px solid #89898d;" /><td style="border-bottom: 1px solid #89898d;" />${_use_gross_profit_mtd ? '<td style="border-bottom: 1px solid #89898d;" />' : ''}<td style="border-bottom: 1px solid #89898d;border-right: 1px solid #89898d;" /></tr>`;
	dynhtml+='<tr />';
	dynhtml+='<tr>';
	dynhtml+='<td style="text-align: left;border: 0px solid #89898d;font-weight: bold;font-family: Arial; font-size: 10pt;" colspan="3"> ' + "&nbsp;" + 'Last Year Actual Sales   ' + '</td>';
    dynhtml+='<td style="text-align: right;border: 0px solid #89898d;font-weight: bold;font-family: Arial; font-size: 10pt;" colspan="1"> ' + "&nbsp;" + dollarCAD.format(prevyrsalessum) +  '</td>';
    dynhtml+='</tr>';
	dynhtml+='<tr>';
	dynhtml+='<td style="text-align: left;border: 0px solid #89898d;font-weight: bold;font-family: Arial; font-size: 10pt;" colspan="3"> ' + "&nbsp;" + 'Last Year Actual Margin  ' + '</td>';
    dynhtml+='<td style="text-align: right;border: 0px solid #89898d;font-weight: bold;font-family: Arial; font-size: 10pt;" colspan="1"> ' + "&nbsp;" + grossmar +  '</td>';
	dynhtml+='</tr>';
	dynhtml+='</table>';

	//==================================================
	// TABLE - DAILY SALES BY COST CENTER
	//==================================================
	var $crntid2="exprtid_2";
	totlcnt[1]="#"+$crntid2;

	// Generate title
	dynhtml+='<table style= "font-family: Arial; font-size: 10pt;" id='+$crntid2+'>';
	dynhtml+='<caption style="text-align: left;"><span style="font-weight: bold;text-align: left;font-family: Arial; font-size: 10pt;">Sales Statistics For '+company+' by Cost Center</br></span><span style="font-family: Arial; font-size: 10pt; font-weight: normal;text-align: left;">'+ 'Month: &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;' + fmonth + '-' + fyear +'</span><caption>';	
	dynhtml+='<tr><td>&nbsp;</td></tr>';

	// Generate table headers
	dynhtml+='<tr>';
	dynhtml+='<th style="text-align: center;border: 1px solid #89898d;font-weight: bold;font-style: italic;font-family: Arial; font-size: 10pt;" colspan="2"> ' + "&nbsp;Cost Center : " + '</span></th>';
	for(var cnt=0; cnt < datacostcntlist3.length; cnt++) 
	{
		var costcntcoltitle= datacostcntlist3[cnt];
		dynhtml+=`<th style="text-align: center;border: 1px solid #89898d;font-weight: bold;font-style: italic;font-family: Arial; font-size: 10pt;" colspan="${_use_gross_profit_mtd ? '5' : '4'}"> ` + "&nbsp;" + costcntcoltitle + '</span></th>';
	}
	dynhtml+='</tr>';

    dynhtml+='<tr>';
	dynhtml+='<th style="text-align: center;border: 1px solid #89898d;font-weight: bold;font-style: italic;font-family: Arial; font-size: 10pt;" colspan="2"> ' + "&nbsp;" + '</span></th>';
	for(var cnt=0; cnt < datacostcntlist3.length; cnt++) 
	{
		dynhtml+=`<th style="text-align: center;border: 1px solid #89898d;font-weight: bold;font-style: italic;font-family: Arial; font-size: 10pt;" colspan="${_use_gross_profit_mtd ? '5' : '4'}"> ` + "&nbsp;" + fmonth + ' ' + fyear + '</span></th>';
	}
	dynhtml+='</tr>';

	dynhtml+='<tr>';
	dynhtml+='<td width="110" style="text-align: center;border: 1px solid #89898d;font-weight: bold;font-family: Arial; font-size: 10pt;font-style: italic;">Date</td>';
	dynhtml+='<td width="120" style="text-align: center;border: 1px solid #89898d;font-weight: bold;font-family: Arial; font-size: 10pt;">Day</td>';
	for(var cnt0=0; cnt0 < datacostcntlist3.length; cnt0++) 
	{
		for(var cnt=2; cnt < titlelist.length; cnt++) 
	 	{					
			var colmnth= titlelist[cnt].label.toString();
			var colfldname = titlelist[cnt].fieldname.toString();

			if (colfldname=='sales') {
				dynhtml+='<td width="150" style="text-align: center;border: 1px solid #89898d;font-weight: bold;font-family: Arial; font-size: 10pt;">'+(colmnth).toString()+'</td>';
			}
			else if (colfldname=='salesmtd') {
				dynhtml+='<td width="150" style="text-align: center;border: 1px solid #89898d;font-weight: bold;font-family: Arial; font-size: 10pt;">'+(colmnth).toString()+'</td>';
			}
			else if (colfldname=='noofinv') {
				colmnth= "# of Inv.'s"
				dynhtml+='<td width="110" style="text-align: center;border: 1px solid #89898d;font-weight: bold;font-family: Arial; font-size: 10pt;">'+(colmnth).toString()+'</td>';
			}
			else if (colfldname=='gross') {
				dynhtml+='<td width="100" style="text-align: center;border: 1px solid #89898d;font-weight: bold;font-family: Arial; font-size: 10pt;">'+(colmnth).toString()+'</td>';
			}
			else if (colfldname=='grossmtd' && _use_gross_profit_mtd) {
				dynhtml+='<td width="150" style="text-align: center;border: 1px solid #89898d;font-weight: bold;font-family: Arial; font-size: 10pt;">'+(colmnth).toString()+'</td>';
			}
			else if (colfldname=='date') {
				dynhtml+='<td width="110" style="text-align: center;border: 1px solid #89898d;font-weight: bold;font-family: Arial; font-size: 10pt;font-style: italic;">'+(colmnth).toString()+'</td>';
			}
			else if (colfldname=='day') {
					dynhtml+='<td width="120" style="text-align: center;border: 1px solid #89898d;font-weight: bold;font-family: Arial; font-size: 10pt;">'+(colmnth).toString()+'</td>';
			} 
		}		
	}
	dynhtml+='</tr>';

	// Populate table data
	dynhtml += row_celldynFunc2(datalist2,datacostcntlist3)

	// Close off the bottom of the table.
	dynhtml+='<tr>'
	dynhtml+='<td style=""></td><td style="">'
	for(var cnt=0; cnt < datacostcntlist3.length; cnt++) 
	{
	    dynhtml+=`<td /><td /><td />${_use_gross_profit_mtd ? '<td />' : ''}<td style="border-right: 1px solid #89898d;" />`;
	}
	dynhtml+='</tr>'
	dynhtml+='<tr>'
	dynhtml+='<td style="border-bottom: 1px solid #89898d;"></td><td style="border-bottom: 1px solid #89898d;"></td>'
	for(var cnt=0; cnt < datacostcntlist3.length; cnt++) 
	{
		dynhtml+=`<td style="border-bottom: 1px solid #89898d;" />
				  <td style="border-bottom: 1px solid #89898d;" />
				  <td style="border-bottom: 1px solid #89898d;" />
				  ${_use_gross_profit_mtd ? '<td style="border-bottom: 1px solid #89898d;" />' : ''}
				  <td style="border-bottom: 1px solid #89898d;border-right: 1px solid #89898d;"></td>`
	}
	dynhtml+='</tr>'
	dynhtml+='<tr></tr><tr></tr><tr></tr><tr></tr>';
	dynhtml+='</table></div>';

	$(".report-wrapper").hide();
	$(".report-wrapper").append(dynhtml);	
	tablesToExcel(totlcnt, 'SalesDailyReport.xls')
}

// Generates an html data row using consolidated cost center data.
function row_celldynFunc(datalist){	
	celldynhtml=``;

	curr_year_html = [];
	prev_year_html = [];

	const options = { year: 'numeric', month: 'short', day: 'numeric' }
	let dollarCAD = Intl.NumberFormat("en-CA", {
		style: "currency",
		currency: "CAD",
		useGrouping: true,
	});
	let amountFormatter = Intl.NumberFormat("en-CA", {
		style: "decimal",
		useGrouping: true,
		minimumFractionDigits: 2,
	    maximumFractionDigits: 2,
	});

	var right_border = 'border-right: 1px solid #89898d;';
	var curr_year_blank = `<td /><td /><td /><td /><td />${_use_gross_profit_mtd ? '<td />' : ''}<td style="${right_border}" />`;
	var prev_year_blank = `<td /><td /><td />${_use_gross_profit_mtd ? '<td />' : ''}<td style="${right_border}" />`;

	// Iterate through the data list, adding seperate year data to separate lists. Only include data
	// from weekdays in the lists.
	for(var index = 0; index < datalist.length; index++) {	
		var row_data = datalist[index];
		html = ``;

		var date = new Date(row_data.date);
		if (date <= new Date()) {
			html+=`<td style="font-family: Arial; font-size: 10pt;">`+date.toLocaleDateString('en-CA', options)+`</td>`;
			html+=`<td style="font-family: Arial; font-size: 10pt;">`+row_data.day+`</td>`;
			html+=`<td style="text-align: center;font-family: Arial; font-size: 10pt;">`+row_data.noofinv+`</td>`;
			html+=`<td style="font-family: Arial; font-size: 10pt;">`+amountFormatter.format(row_data.sales)+`</td>`;
			html+=`<td style="font-family: Arial; font-size: 10pt;">`+dollarCAD.format(row_data.salesmtd)+`</td>`;
			html+=`<td style="font-family: Arial; font-size: 10pt; ${!_use_gross_profit_mtd ? right_border : ''}">`+formatAsPercent(row_data.gross)+`</td>`;
			html+=_use_gross_profit_mtd ? `<td style="font-family: Arial; font-size: 10pt; ${right_border}">`+formatAsPercent(row_data.grossmtd)+`</td>` : ``;
		} else {
			html+= curr_year_blank;
		}

		curr_year_html.push(html);

		html = ``;
		html+=`<td style="font-family: Arial; font-size: 10pt; text-align: center;">`+row_data.noofinv2+`</td>`;
		html+=`<td style="font-family: Arial; font-size: 10pt;">`+amountFormatter.format(row_data.sales2)+`</td>`;
		html+=`<td style="font-family: Arial; font-size: 10pt;">`+dollarCAD.format(row_data.salesmtd2)+`</td>`;
		html+=`<td style="font-family: Arial; font-size: 10pt; ${!_use_gross_profit_mtd ? right_border : ''}">`+formatAsPercent(row_data.gross2)+`</td>`;
		html+=_use_gross_profit_mtd ? `<td style="font-family: Arial; font-size: 10pt; ${right_border}">`+formatAsPercent(row_data.grossmtd2)+`</td>` : ``;

		prev_year_html.push(html);
	}

	// Pair rows by their list index instead of business day, so there are no gaps in the list.
	for (var index = 0; index < Math.max(curr_year_html.length, prev_year_html.length); index++) {
		celldynhtml += `<tr>`;
		celldynhtml += (index < curr_year_html.length) ? curr_year_html[index] : curr_year_blank;
		celldynhtml += (index < prev_year_html.length) ? prev_year_html[index] : prev_year_blank;
		celldynhtml += `</tr>`;
	}

	return celldynhtml;
}

// Generates an html data row using data for each cost center.
function row_celldynFunc2(datalist, costcentlst){	
	celldynhtml="";

	const options = { year: 'numeric', month: 'short', day: 'numeric' }
	let dollarCAD = Intl.NumberFormat("en-CA", {
		style: "currency",
		currency: "CAD",
		useGrouping: true,
	});
	let amountFormatter = Intl.NumberFormat("en-CA", {
		style: "decimal",
		useGrouping: true,
		minimumFractionDigits: 2,
		maximumFractionDigits: 2,
	});

	// Iterate through each data row
	for(var index = 0; index < datalist.length; index++) {
		var row_data = datalist[index];	

		var right_border = 'border-right: 1px solid #89898d;';

		celldynhtml+='<tr>';
		celldynhtml+='<td style="font-family: Arial; font-size: 10pt;">'+new Date(row_data.date).toLocaleDateString('en-CA', options)+'</td>';
		celldynhtml+='<td style="font-family: Arial; font-size: 10pt;">'+row_data.day+'</td>';
		for(var cnt=0; cnt < costcentlst.length; cnt++) 
		{
			var col1 = 'noofinvcstcnt' + cnt
			var col2 = 'salescstcnt' + cnt
			var col3 = 'salesmtdcstcnt' + cnt
			var col4 = 'grosscstcnt' + cnt
			var col5 = 'grossmtdcstcnt' + cnt
			
			celldynhtml+='<td style="text-align: center;font-family: Arial; font-size: 10pt;">'+row_data[col1]+'</td>';
			celldynhtml+='<td style="font-family: Arial; font-size: 10pt;">'+amountFormatter.format(row_data[col2])+'</td>';
			celldynhtml+='<td style="font-family: Arial; font-size: 10pt;">'+dollarCAD.format(row_data[col3])+'</td>';
			celldynhtml+=`<td style="font-family: Arial; font-size: 10pt; ${!_use_gross_profit_mtd ? right_border : ''}">`+formatAsPercent(row_data[col4])+'</td>';
			celldynhtml+=_use_gross_profit_mtd ? `<td style="font-family: Arial; font-size: 10pt; ${right_border}">`+formatAsPercent(row_data[col5])+'</td>' : '';
		}
		celldynhtml+='</tr>';
	}

	return celldynhtml;
}

function getyears(){
	let yrarr = [];
	curyr = new Date().getFullYear();
	for (let i = 0; i < 10; i++) {
	   yrarr.push(curyr);
	   curyr = curyr -1;
	 }
	return yrarr;
}

function formatAsPercent(num) {
	return new Intl.NumberFormat('default', {
	  style: 'percent',
	  minimumFractionDigits: 2,
	  maximumFractionDigits: 2,
	}).format(num / 100);
  }


var tablesToExcel = (function () {
	var uri = 'data:application/vnd.ms-excel;base64,'
		, html_start = `<html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">`
		, template_ExcelWorksheet = `<x:ExcelWorksheet><x:Name>{SheetName}</x:Name><x:WorksheetSource HRef="sheet{SheetIndex}.htm"/></x:ExcelWorksheet>`
		, template_ListWorksheet = `<o:File HRef="sheet{SheetIndex}.htm"/>`
		, template_HTMLWorksheet = `
------=_NextPart_dummy
Content-Location: sheet{SheetIndex}.htm
Content-Type: text/html; charset=windows-1252

` + html_start + `
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<link id="Main-File" rel="Main-File" href="../WorkBook.htm">
<link rel="File-List" href="filelist.xml">
<style>
	@page {
		margin:.25in .25in .25in .25in;
		mso-header-margin:.025in;
		mso-footer-margin:.025in;
		mso-page-orientation:landscape;
	}
</style>
</head>
<body><table>{SheetContent}</table></body>
</html>`
		, template_WorkBook = `MIME-Version: 1.0
X-Document-Type: Workbook
Content-Type: multipart/related; boundary="----=_NextPart_dummy"

------=_NextPart_dummy
Content-Location: WorkBook.htm
Content-Type: text/html; charset=windows-1252

` + html_start + `
<head>
<meta name="Excel Workbook Frameset">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<link rel="File-List" href="filelist.xml">
<!--[if gte mso 9]><xml>
<x:ExcelWorkbook>
<x:ExcelWorksheets>{ExcelWorksheets}</x:ExcelWorksheets>
<x:ActiveSheet>0</x:ActiveSheet>
</x:ExcelWorkbook>
</xml><![endif]-->
</head>
<frameset>
<frame src="sheet0.htm" name="frSheet">
<noframes><body><p>This page uses frames, but your browser does not support them.</p></body></noframes>
</frameset>
</html>
{HTMLWorksheets}
Content-Location: filelist.xml
Content-Type: text/xml; charset="utf-8"

<xml xmlns:o="urn:schemas-microsoft-com:office:office">
<o:MainFile HRef="../WorkBook.htm"/>
{ListWorksheets}
<o:File HRef="filelist.xml"/>
</xml>
------=_NextPart_dummy--
`
		, base64 = function (s) { return window.btoa(unescape(encodeURIComponent(s))) }
		, format = function (s, c) { return s.replace(/{(\w+)}/g, function (m, p) { return c[p]; }) }
	return function (tables, filename) {
		var context_WorkBook = {
			ExcelWorksheets: ''
			, HTMLWorksheets: ''
			, ListWorksheets: ''
		};		
		var tables = jQuery(tables);
		var tbk = 0
		//SheetIndex =1;
		$.each(tables, function (SheetIndex,val) {			
			var $table = $(val);
			var SheetName = "";
			if (SheetIndex == 0) {
				SheetName = 'Consolidated Daily Sales' ;
			}
			else if(SheetIndex == 1) {
				SheetName = 'Daily Sales by Cost Center' ;
			}
			context_WorkBook.ExcelWorksheets += format(template_ExcelWorksheet, {
				SheetIndex: SheetIndex
				, SheetName: SheetName
			});
			context_WorkBook.HTMLWorksheets += format(template_HTMLWorksheet, {
				SheetIndex: SheetIndex
				, SheetContent: $table.html()
			});
			context_WorkBook.ListWorksheets += format(template_ListWorksheet, {
				SheetIndex: SheetIndex
			});
			tbk += 1
		});

		var link = document.createElement("A");
		link.href = uri + base64(format(template_WorkBook, context_WorkBook));
		link.download = filename || 'Workbook.xls';
		link.target = '_blank';
		document.body.appendChild(link);
		link.click();
		document.body.removeChild(link);
	}
})();

