// Copyright (c) 2022, abayomi.awosusi@sgatechsolutions.com and contributors
// For license information, please see license.txt
/* eslint-disable */

// Determines whether or not the gross profit MTD should be displayed. Adjusts the spreadsheet to account for the columns absence/presence.
var _use_gross_profit_mtd = true;
var _is_first_run = true;

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
			let filters = report.get_values();	

			// Avoid the issue of the report not using the correct date on subsequent runs.
			if (!_is_first_run) {
				frappe.msgprint("Please refresh the page before running this report again.");
				return;
			}
			_is_first_run = false;
			
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

function dynamic_exportcontent(cnt_list, company, fmonth, fyear){	
	var dynhtml="";
	var totlcnt=[];
	
	var titlelist = cnt_list[0]
	var datalist = cnt_list[1][0]
	var datalist2 = cnt_list[1][1]
	var datacostcntlist3 = cnt_list[1][2]

	//==================================================
	// TABLE - CONSOLIDATED DAILY SALES
	//==================================================
	var $crntid="exprtid_1";
	totlcnt[0]="#"+$crntid;

	dynhtml += `<div id="dvData">`;
	dynhtml += `<table id=${$crntid}>`;
	dynhtml += generate_headers(titlelist, company, fmonth, fyear);
	dynhtml += populate_consolidated(datalist);
	dynhtml += generate_table_end(titlelist);
	dynhtml += '</table>';

	//==================================================
	// TABLE - DAILY SALES BY COST CENTER
	//==================================================
	var $crntid2="exprtid_2";
	totlcnt[1]="#"+$crntid2;

	// Generate headers
	dynhtml += `<table id=${$crntid2}>`;
	dynhtml += generate_headers(titlelist, company, fmonth, fyear, datacostcntlist3)
	dynhtml += populate_cost_center(datalist2, datacostcntlist3)
	dynhtml += generate_table_end(titlelist, datacostcntlist3)
	dynhtml += '</table>';
	dynhtml += '</div>';

	$(".report-wrapper").hide();
	$(".report-wrapper").append(dynhtml);	
	tablesToExcel(totlcnt, 'SalesDailyReport.xls')
}

// Generates the header and title rows for both the consolidated and cost center tabs
function generate_headers(titlelist, company, fmonth, fyear, cost_centers = null) {
	var html = '';

	// This is a really hacky way of getting the tables to print nicely.
	// This gets the cost center tables to be very slightly wider to account
	// for an extra column on the consolidated tab.
	var width_factor = (cost_centers != null ? 1.05 : 1)

	// Generate the table title
	// On the cost center tab, this will appear for every 2 data sets.
	var col_count = titlelist.length - (_use_gross_profit_mtd ? 0 : 1);
	var col_span = (col_count * 2) - 2;
	var company_title = '';
	var month_tile = '';
	for (var i = 0; i < (cost_centers == null ? 2 : cost_centers.length); i++) {
		if (i % 2 == 0) {
			company_title += `<td style="text-align: left; font-weight: bold;" colspan="${col_span}">Sales Statistics For ${company}</td>`;
			month_tile += `<td style="text-align: left; font-weight: normal;" colspan="${col_span}">Month: ${fmonth}&nbsp${fyear}</td>`;	
		}
	}

	html += `<tr>${company_title}</tr><tr>${month_tile}</tr><tr />`;
	
	// Generate primary headers for each cost center
	if (cost_centers != null) {
		var col_count = titlelist.length - 2 - (_use_gross_profit_mtd ? 0 : 1)
		var cc_row = '';
		var month_row = '';

		for(var i = 0; i < cost_centers.length; i++) {
			if (i % 2 == 0) {
				cc_row += '<th style="text-align: center; border: 1px solid #89898d; font-weight: bold; font-style: italic;" colspan="2">Cost Center :</th>';
				month_row += '<th style="border: 1px solid #89898d;" colspan="2"></th>';
			}

			cc_row +=`<th style="text-align: center; border: 1px solid #89898d; font-weight: bold; font-style: italic;" colspan="${col_count}">${cost_centers[i]}</th>`;
			month_row +=`<th style="text-align: center; border: 1px solid #89898d; font-weight: bold; font-style: italic;" colspan="${col_count}">${fmonth}&nbsp${fyear}</th>`;
		}

		html += `<tr>${cc_row}</tr><tr>${month_row}</tr>`;
	} else {
		html += '<tr>';
		html += `<th style="text-align: center; border: 1px solid #89898d; font-weight: bold; font-style: italic;" colspan="${col_count}">${fmonth}&nbsp${fyear}</th>`;
		html += `<th style="text-align: center; border: 1px solid #89898d; font-weight: bold; font-style: italic;" colspan="${col_count - 1}">${fmonth}&nbsp${parseInt(fyear) - 1}</th>`;
		html += '</tr>';
	}
	
	// Generate the sub-headers for each cost center
	html += '<tr>';
	for (var i = 0; i < (cost_centers == null ? 2 : cost_centers.length); i++) {
		
		// For every 2 data sets, add the date/day columns
		if (i % 2 == 0) {
			for(var j = 0; j < 2; j++) {
				html += `<td width="${titlelist[j].width * width_factor}" style="text-align: ${titlelist[j].alignment}; 
					border: 1px solid #89898d; font-weight: bold;">${titlelist[j].label}</td>`;
			}
		} else if (cost_centers == null) {
			// For the consolidated report, repeat the "Day" column for the second data
			html += `<td width="${titlelist[1].width * width_factor}" style="text-align: ${titlelist[1].alignment}; 
				border: 1px solid #89898d; font-weight: bold;">${titlelist[1].label}</td>`;
		}

		// Add the rest of the columns
		for(var j = 2; j < titlelist.length; j++) {
			if (titlelist[j].fieldname != "grossmtd" || _use_gross_profit_mtd) {
				html += `<td width="${titlelist[j].width * width_factor}" style="text-align: ${titlelist[j].alignment}; 
					border: 1px solid #89898d; font-weight: bold;">${titlelist[j].label}</td>`;
			}
		}
	}

	html += '</tr>';
	return html;
}

// Generates the end rows for the tables
function generate_table_end(titlelist, cost_centers = null) {
	html = '';

	// Generate the table ends for each cost center
	html += '<tr>';
	for (var i = 0; i < (cost_centers == null ? 2 : cost_centers.length); i++) {
		
		// For every 2 data sets, add the date/day column ends
		if (i % 2 == 0) {
			html += `<td style="border-left: 1px solid #89898d; border-bottom: 1px solid #89898d;" />
						<td style="border-bottom: 1px solid #89898d; ${cost_centers != null ? "border-right: 1px solid #89898d;" : ""}" />`;
		} else if (cost_centers == null) {
			// For the consolidated report, repeat the "Day" column end for the second data
			html += `<td style="border-left: 1px solid #89898d; border-bottom: 1px solid #89898d;" />`;
		}

		// Add the rest of the columns
		var col_count = titlelist.length - 2 - (_use_gross_profit_mtd ? 0 : 1)
		for(var j = 0; j < col_count; j++) {
			html += `<td style="border-bottom: 1px solid #89898d; ${j == 0 && cost_centers != null ? "border-left: 1px solid #89898d;" : ""} 
				${j == col_count - 1 ? "border-right: 1px solid #89898d;" : ""}" />`;
		}
	}

	html += '</tr>';
	return html;
}

// Generates an html data row using consolidated cost center data.
function populate_consolidated(datalist){	
	celldynhtml=``;

	curr_year_html = [];
	prev_year_html = [];

	const options = { year: 'numeric', month: 'short', day: 'numeric' }
	let dollarCAD = Intl.NumberFormat("en-CA", {
		style: "currency",
		currency: "CAD",
		useGrouping: true,
		minimumFractionDigits: 0,
		maximumFractionDigits: 0,
	});

	var left_border = 'border-left: 1px solid #89898d;';
	var right_border = 'border-right: 1px solid #89898d;';
	var curr_year_blank = `<td style="${left_border}" /><td /><td /><td /><td />${_use_gross_profit_mtd ? '<td />' : ''}<td style="${right_border}" />`;
	var prev_year_blank = `<td style="${left_border}" /><td /><td />${_use_gross_profit_mtd ? '<td />' : ''}<td style="${right_border}" />`;

	// Iterate through the data list, adding seperate year data to separate lists.
	for(var index = 0; index < datalist.length; index++) {	
		var row_data = datalist[index];
		html = ``;

		var date = new Date(row_data.date);
		if (date <= new Date()) {
			html+=`<td style="text-align: center;   ${left_border}">`+date.toLocaleDateString('en-CA', options)+`</td>`;
			html+=`<td style="text-align: left;">`+row_data.day+`</td>`;
			html+=`<td style="text-align: center;">`+row_data.noofinv+`</td>`;
			html+=`<td style="text-align: left;">`+dollarCAD.format(row_data.sales)+`</td>`;
			html+=`<td style="text-align: left;">`+dollarCAD.format(row_data.salesmtd)+`</td>`;
			html+=`<td style="text-align: center; ${!_use_gross_profit_mtd ? right_border : ''}">`+formatAsPercent(row_data.gross)+`</td>`;
			html+=_use_gross_profit_mtd ? `<td style="text-align: center; ${right_border}">`+formatAsPercent(row_data.grossmtd)+`</td>` : ``;
		} else {
			html+= curr_year_blank;
		}

		curr_year_html.push(html);

		html = ``;
		html+=`<td style="text-align: left;">`+row_data.day2+`</td>`;
		html+=`<td style="text-align: center;">`+row_data.noofinv2+`</td>`;
		html+=`<td style="text-align: left;">`+dollarCAD.format(row_data.sales2)+`</td>`;
		html+=`<td style="text-align: left;">`+dollarCAD.format(row_data.salesmtd2)+`</td>`;
		html+=`<td style="text-align: center; ${!_use_gross_profit_mtd ? right_border : ''}">`+formatAsPercent(row_data.gross2)+`</td>`;
		html+=_use_gross_profit_mtd ? `<td style="text-align: center; ${right_border}">`+formatAsPercent(row_data.grossmtd2)+`</td>` : ``;

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
function populate_cost_center(datalist, costcentlst){	
	celldynhtml="";

	const options = { year: 'numeric', month: 'short', day: 'numeric' }
	let dollarCAD = Intl.NumberFormat("en-CA", {
		style: "currency",
		currency: "CAD",
		useGrouping: true,
		minimumFractionDigits: 0,
		maximumFractionDigits: 0,
	});

	// Iterate through each data row
	for(var index = 0; index < datalist.length; index++) {
		var row_data = datalist[index];	

		var left_border = 'border-left: 1px solid #89898d;';
		var right_border = 'border-right: 1px solid #89898d;';

		var date = new Date(row_data.date);
		if (date <= new Date()) {
			celldynhtml+='<tr>';
			for(var cnt=0; cnt < costcentlst.length; cnt++) 
			{
				// Add the dates for every 2 data sets
				if (cnt % 2 == 0) {
					celldynhtml+=`<td style="text-align: center; ${left_border}">`+new Date(row_data.date).toLocaleDateString('en-CA', options)+'</td>';
					celldynhtml+=`<td style="${right_border}">`+row_data.day+'</td>';
				}

				var col1 = 'noofinvcstcnt' + cnt
				var col2 = 'salescstcnt' + cnt
				var col3 = 'salesmtdcstcnt' + cnt
				var col4 = 'grosscstcnt' + cnt
				var col5 = 'grossmtdcstcnt' + cnt
				
				celldynhtml+=`<td style="text-align: center; ${left_border}">`+row_data[col1]+'</td>';
				celldynhtml+='<td style="text-align: center;">'+dollarCAD.format(row_data[col2])+'</td>';
				celldynhtml+='<td style="text-align: center;">'+dollarCAD.format(row_data[col3])+'</td>';
				celldynhtml+=`<td style="text-align: center; ${!_use_gross_profit_mtd ? right_border : ''}">`+formatAsPercent(row_data[col4])+'</td>';
				celldynhtml+=_use_gross_profit_mtd ? `<td style="text-align: center; ${right_border}">`+formatAsPercent(row_data[col5])+'</td>' : '';
			}
			celldynhtml+='</tr>';
		} else {
			celldynhtml+='<tr>';
			for(var cnt=0; cnt < costcentlst.length; cnt++) {
				// Add blank spots for the dates for every 2 data sets
				if (cnt % 2 == 0) {
					celldynhtml+=`<td style="${left_border}" /><td style="${right_border}" />`;
				}
				celldynhtml += `<td style="${left_border}" /><td /><td />${_use_gross_profit_mtd ? '<td />' : ''}<td style="${right_border}" />`;
			}
			celldynhtml+='</tr>';
		}
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
		margin-top:.5in;
		margin-left:.25in;
		margin-bottom:.25in;
		margin-right:.025in;
		mso-header-margin:.025in;
		mso-footer-margin:.025in;
		mso-page-orientation:landscape;
	}
	td {
		font-family:Calibri; 
		font-size:10pt;
		vertical-align: middle;
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

		// Reload the page automatically after generating the report.
		window.location.reload();
	}
})();

