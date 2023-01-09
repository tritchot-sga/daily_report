import frappe
import datetime
from frappe import _, qb, scrub
from frappe.query_builder import CustomFunction
from frappe.query_builder.functions import Max
from frappe.utils import date_diff, flt, getdate, cint
from erpnext.controllers.queries import get_match_cond
from erpnext.stock.utils import get_incoming_rate

class GrossProfitGenerator(object):
    def __init__(self, filters=None):
        self.data = []
        self.average_buying_rate = {}
        self.filters = frappe._dict(filters)
        self.load_invoice_items()

        self.group_items_by_invoice()

        self.load_stock_ledger_entries()
        self.load_product_bundle()
        self.load_non_stock_items()
        self.get_returned_invoice_items()
        self.process()

    def process(self):
        self.grouped = {}
        self.grouped_data = []

        self.currency_precision = cint(frappe.db.get_default("currency_precision")) or 3
        self.float_precision = cint(frappe.db.get_default("float_precision")) or 2
        self.filters.group_by = "Invoice"
        grouped_by_invoice = True if self.filters.get("group_by") == "Invoice" else False

        buying_amount = 0

        for row in reversed(self.si_list):
            if self.skip_row(row):
                continue

            row.base_amount = flt(row.base_net_amount, self.currency_precision)

            product_bundles = []
            if row.update_stock:
                product_bundles = self.product_bundles.get(row.parenttype, {}).get(row.parent, frappe._dict())
            elif row.dn_detail:
                product_bundles = self.product_bundles.get("Delivery Note", {}).get(
                    row.delivery_note, frappe._dict()
                )
                row.item_row = row.dn_detail

            # get buying amount
            if row.item_code in product_bundles:
                row.buying_amount = flt(
                    self.get_buying_amount_from_product_bundle(row, product_bundles[row.item_code]),
                    self.currency_precision,
                )
            else:
                row.buying_amount = flt(self.get_buying_amount(row, row.item_code), self.currency_precision)

            if grouped_by_invoice:
                if row.indent == 1.0:
                    buying_amount += row.buying_amount
                elif row.indent == 0.0:
                    row.buying_amount = buying_amount
                    buying_amount = 0

            # get buying rate
            if flt(row.qty):
                row.buying_rate = flt(row.buying_amount / flt(row.qty), self.float_precision)
                row.base_rate = flt(row.base_amount / flt(row.qty), self.float_precision)
            else:
                if self.is_not_invoice_row(row):
                    row.buying_rate, row.base_rate = 0.0, 0.0

            # calculate gross profit
            row.gross_profit = flt(row.base_amount - row.buying_amount, self.currency_precision)
            if row.base_amount:
                row.gross_profit_percent = flt(
                    (row.gross_profit / row.base_amount) * 100.0, self.currency_precision
                )
            else:
                row.gross_profit_percent = 0.0

            # add to grouped
            self.grouped.setdefault(row.get(scrub(self.filters.group_by)), []).append(row)

        if self.grouped:
            self.get_average_rate_based_on_group_by()
        #print(self.si_list)    

    def get_average_rate_based_on_group_by(self):
        for key in list(self.grouped):
            if self.filters.get("group_by") != "Invoice":
                for i, row in enumerate(self.grouped[key]):
                    if i == 0:
                        new_row = row
                    else:
                        new_row.qty += flt(row.qty)
                        new_row.buying_amount += flt(row.buying_amount, self.currency_precision)
                        new_row.base_amount += flt(row.base_amount, self.currency_precision)
                new_row = self.set_average_rate(new_row)
                self.grouped_data.append(new_row)
            else:
                for i, row in enumerate(self.grouped[key]):
                    if row.indent == 1.0:
                        if (
                            row.parent in self.returned_invoices and row.item_code in self.returned_invoices[row.parent]
                        ):
                            returned_item_rows = self.returned_invoices[row.parent][row.item_code]
                            for returned_item_row in returned_item_rows:
                                row.qty += flt(returned_item_row.qty)
                                row.base_amount += flt(returned_item_row.base_amount, self.currency_precision)
                            row.buying_amount = flt(flt(row.qty) * flt(row.buying_rate), self.currency_precision)
                        if flt(row.qty) or row.base_amount:
                            row = self.set_average_rate(row)
                            self.grouped_data.append(row)

    def is_not_invoice_row(self, row):
        return (self.filters.get("group_by") == "Invoice" and row.indent != 0.0) or self.filters.get(
            "group_by"
        ) != "Invoice"

    def set_average_rate(self, new_row):
        self.set_average_gross_profit(new_row)
        new_row.buying_rate = (
            flt(new_row.buying_amount / new_row.qty, self.float_precision) if new_row.qty else 0
        )
        new_row.base_rate = (
            flt(new_row.base_amount / new_row.qty, self.float_precision) if new_row.qty else 0
        )
        return new_row

    def set_average_gross_profit(self, new_row):
        new_row.gross_profit = flt(new_row.base_amount - new_row.buying_amount, self.currency_precision)
        new_row.gross_profit_percent = (
            flt(((new_row.gross_profit / new_row.base_amount) * 100.0), self.currency_precision)
            if new_row.base_amount
            else 0
        )
        new_row.buying_rate = (
            flt(new_row.buying_amount / flt(new_row.qty), self.float_precision) if flt(new_row.qty) else 0
        )
        new_row.base_rate = (
            flt(new_row.base_amount / flt(new_row.qty), self.float_precision) if flt(new_row.qty) else 0
        )

    def get_returned_invoice_items(self):
        returned_invoices = frappe.db.sql(
            """
            select
                si.name, si_item.item_code, si_item.stock_qty as qty, si_item.base_net_amount as base_amount, si.return_against
            from
                `tabSales Invoice` si, `tabSales Invoice Item` si_item
            where
                si.name = si_item.parent
                and si.docstatus = 1
                and si.is_return = 1
        """,
            as_dict=1,
        )

        self.returned_invoices = frappe._dict()
        for inv in returned_invoices:
            self.returned_invoices.setdefault(inv.return_against, frappe._dict()).setdefault(
                inv.item_code, []
            ).append(inv)

    def skip_row(self, row):
        if self.filters.get("group_by") != "Invoice":
            if not row.get(scrub(self.filters.get("group_by", ""))):
                return True

        return False

    def get_buying_amount_from_product_bundle(self, row, product_bundle):
        buying_amount = 0.0
        for packed_item in product_bundle:
            if packed_item.get("parent_detail_docname") == row.item_row:
                buying_amount += self.get_buying_amount(row, packed_item.item_code)

        return flt(buying_amount, self.currency_precision)

    def get_buying_amount(self, row, item_code):
        # IMP NOTE
        # stock_ledger_entries should already be filtered by item_code and warehouse and
        # sorted by posting_date desc, posting_time desc
        if item_code in self.non_stock_items and (row.project or row.cost_center):
            # Issue 6089-Get last purchasing rate for non-stock item
            item_rate = self.get_last_purchase_rate(item_code, row)
            return flt(row.qty) * item_rate

        else:
            my_sle = self.sle.get((item_code, row.warehouse))
            if (row.update_stock or row.dn_detail) and my_sle:
                parenttype, parent = row.parenttype, row.parent
                if row.dn_detail:
                    parenttype, parent = "Delivery Note", row.delivery_note

                for i, sle in enumerate(my_sle):
                    # find the stock valution rate from stock ledger entry
                    if (
                        sle.voucher_type == parenttype
                        and parent == sle.voucher_no
                        and sle.voucher_detail_no == row.item_row
                    ):
                        previous_stock_value = len(my_sle) > i + 1 and flt(my_sle[i + 1].stock_value) or 0.0

                        if previous_stock_value:
                            return (previous_stock_value - flt(sle.stock_value)) * flt(row.qty) / abs(flt(sle.qty))
                        else:
                            return flt(row.qty) * self.get_average_buying_rate(row, item_code)
            else:
                return flt(row.qty) * self.get_average_buying_rate(row, item_code)

        return 0.0

    def get_average_buying_rate(self, row, item_code):
        args = row
        if not item_code in self.average_buying_rate:
            args.update(
                {
                    "voucher_type": row.parenttype,
                    "voucher_no": row.parent,
                    "allow_zero_valuation": True,
                    "company": self.filters.company,
                }
            )

            average_buying_rate = get_incoming_rate(args)
            self.average_buying_rate[item_code] = flt(average_buying_rate)

        return self.average_buying_rate[item_code]

    def get_last_purchase_rate(self, item_code, row):
        purchase_invoice = frappe.qb.DocType("Purchase Invoice")
        purchase_invoice_item = frappe.qb.DocType("Purchase Invoice Item")

        query = (
            frappe.qb.from_(purchase_invoice_item)
            .inner_join(purchase_invoice)
            .on(purchase_invoice.name == purchase_invoice_item.parent)
            .select(purchase_invoice_item.base_rate / purchase_invoice_item.conversion_factor)
            .where(purchase_invoice.docstatus == 1)
            .where(purchase_invoice.posting_date <= self.filters.to_date)
            .where(purchase_invoice_item.item_code == item_code)
        )

        if row.project:
            query.where(purchase_invoice_item.project == row.project)

        if row.cost_center:
            query.where(purchase_invoice_item.cost_center == row.cost_center)

        query.orderby(purchase_invoice.posting_date, order=frappe.qb.desc)
        query.limit(1)
        last_purchase_rate = query.run()

        return flt(last_purchase_rate[0][0]) if last_purchase_rate else 0

    def load_invoice_items(self):
        conditions = ""
        if self.filters.company:
            conditions += " and company = %(company)s"
        if self.filters.get("cost_center"):
            conditions += " and `tabSales Invoice Item`.cost_center IN %(cost_center)s"
        if self.filters.month:
            long_month_name = self.filters.get("month")
            datetime_object = datetime.datetime.strptime(long_month_name, "%B")
            month_number = datetime_object.month
            conditions += " and ((MONTH(posting_date) = " + str(month_number) + " and YEAR(posting_date) = " + str(self.filters.get("year")) + "))"
            
        self.si_list = frappe.db.sql(
            """
            select
                `tabSales Invoice Item`.parenttype, `tabSales Invoice Item`.parent,
                `tabSales Invoice`.posting_date, `tabSales Invoice`.posting_time,
                `tabSales Invoice`.project, `tabSales Invoice`.update_stock,
                `tabSales Invoice`.customer, `tabSales Invoice`.customer_group,
                `tabSales Invoice`.territory, `tabSales Invoice Item`.item_code,
                `tabSales Invoice Item`.item_name, `tabSales Invoice Item`.description,
                `tabSales Invoice Item`.warehouse, `tabSales Invoice Item`.item_group,
                `tabSales Invoice Item`.brand, `tabSales Invoice Item`.dn_detail,
                `tabSales Invoice Item`.delivery_note, `tabSales Invoice Item`.stock_qty as qty,
                `tabSales Invoice Item`.base_net_rate, `tabSales Invoice Item`.base_net_amount,
                `tabSales Invoice Item`.name as "item_row", `tabSales Invoice`.is_return,
                `tabSales Invoice Item`.cost_center
            
            from
                `tabSales Invoice` inner join `tabSales Invoice Item`
                    on `tabSales Invoice Item`.parent = `tabSales Invoice`.name
                
            where
                `tabSales Invoice`.docstatus=1 and `tabSales Invoice`.is_opening!='Yes' {conditions} {match_cond}
            order by
                `tabSales Invoice`.posting_date desc, `tabSales Invoice`.posting_time desc""".format(
                conditions=conditions,
                match_cond=get_match_cond("Sales Invoice"),
            ),
            self.filters,
            as_dict=1,
            debug=0,
        )

    def group_items_by_invoice(self):
        """
        Turns list of Sales Invoice Items to a tree of Sales Invoices with their Items as children.
        """

        parents = []

        for row in self.si_list:
            if row.parent not in parents:
                parents.append(row.parent)

        parents_index = 0
        for index, row in enumerate(self.si_list):
            if parents_index < len(parents) and row.parent == parents[parents_index]:
                invoice = self.get_invoice_row(row)
                self.si_list.insert(index, invoice)
                parents_index += 1

            else:
                # skipping the bundle items rows
                if not row.indent:
                    row.indent = 1.0
                    row.parent_invoice = row.parent
                    row.invoice_or_item = row.item_code

                    if frappe.db.exists("Product Bundle", row.item_code):
                        self.add_bundle_items(row, index)

    def get_invoice_row(self, row):
        return frappe._dict(
            {
                "parent_invoice": "",
                "indent": 0.0,
                "invoice_or_item": row.parent,
                "parent": None,
                "posting_date": row.posting_date,
                "posting_time": row.posting_time,
                "project": row.project,
                "update_stock": row.update_stock,
                "customer": row.customer,
                "customer_group": row.customer_group,
                "item_code": None,
                "item_name": None,
                "description": None,
                "warehouse": None,
                "item_group": None,
                "brand": None,
                "dn_detail": None,
                "delivery_note": None,
                "qty": None,
                "item_row": None,
                "is_return": row.is_return,
                "cost_center": row.cost_center,
                "base_net_amount": frappe.db.get_value("Sales Invoice", row.parent, "base_net_total"),
            }
        )

    def add_bundle_items(self, product_bundle, index):
        bundle_items = self.get_bundle_items(product_bundle)

        for i, item in enumerate(bundle_items):
            bundle_item = self.get_bundle_item_row(product_bundle, item)
            self.si_list.insert((index + i + 1), bundle_item)

    def get_bundle_items(self, product_bundle):
        return frappe.get_all(
            "Product Bundle Item", filters={"parent": product_bundle.item_code}, fields=["item_code", "qty"]
        )

    def get_bundle_item_row(self, product_bundle, item):
        item_name, description, item_group, brand = self.get_bundle_item_details(item.item_code)

        return frappe._dict(
            {
                "parent_invoice": product_bundle.item_code,
                "indent": product_bundle.indent + 1,
                "parent": None,
                "invoice_or_item": item.item_code,
                "posting_date": product_bundle.posting_date,
                "posting_time": product_bundle.posting_time,
                "project": product_bundle.project,
                "customer": product_bundle.customer,
                "customer_group": product_bundle.customer_group,
                "item_code": item.item_code,
                "item_name": item_name,
                "description": description,
                "warehouse": product_bundle.warehouse,
                "item_group": item_group,
                "brand": brand,
                "dn_detail": product_bundle.dn_detail,
                "delivery_note": product_bundle.delivery_note,
                "qty": (flt(product_bundle.qty) * flt(item.qty)),
                "item_row": None,
                "is_return": product_bundle.is_return,
                "cost_center": product_bundle.cost_center,
            }
        )

    def get_bundle_item_details(self, item_code):
        return frappe.db.get_value(
            "Item", item_code, ["item_name", "description", "item_group", "brand"]
        )

    def load_stock_ledger_entries(self):
        res = frappe.db.sql(
            """select item_code, voucher_type, voucher_no,
                voucher_detail_no, stock_value, warehouse, actual_qty as qty
            from `tabStock Ledger Entry`
            where company=%(company)s and is_cancelled = 0
            order by
                item_code desc, warehouse desc, posting_date desc,
                posting_time desc, creation desc""",
            self.filters,
            as_dict=True,
        )
        self.sle = {}
        for r in res:
            if (r.item_code, r.warehouse) not in self.sle:
                self.sle[(r.item_code, r.warehouse)] = []

            self.sle[(r.item_code, r.warehouse)].append(r)

    def load_product_bundle(self):
        self.product_bundles = {}

        for d in frappe.db.sql(
            """select parenttype, parent, parent_item,
            item_code, warehouse, -1*qty as total_qty, parent_detail_docname
            from `tabPacked Item` where docstatus=1""",
            as_dict=True,
        ):
            self.product_bundles.setdefault(d.parenttype, frappe._dict()).setdefault(
                d.parent, frappe._dict()
            ).setdefault(d.parent_item, []).append(d)

    def load_non_stock_items(self):
        self.non_stock_items = frappe.db.sql_list(
            """select name from tabItem
            where is_stock_item=0"""
        )


class GrossProfitGenerator2(object):
    def __init__(self, filters=None):
        self.data = []
        self.average_buying_rate = {}
        self.filters = frappe._dict(filters)
        self.load_invoice_items()

        self.group_items_by_invoice()

        self.load_stock_ledger_entries()
        self.load_product_bundle()
        self.load_non_stock_items()
        self.get_returned_invoice_items()
        self.process()

    def process(self):
        self.grouped = {}
        self.grouped_data = []

        self.currency_precision = cint(frappe.db.get_default("currency_precision")) or 3
        self.float_precision = cint(frappe.db.get_default("float_precision")) or 2
        self.filters.group_by = "Invoice"
        grouped_by_invoice = True if self.filters.get("group_by") == "Invoice" else False

        buying_amount = 0

        for row in reversed(self.si_list):
            if self.skip_row(row):
                continue

            row.base_amount = flt(row.base_net_amount, self.currency_precision)

            product_bundles = []
            if row.update_stock:
                product_bundles = self.product_bundles.get(row.parenttype, {}).get(row.parent, frappe._dict())
            elif row.dn_detail:
                product_bundles = self.product_bundles.get("Delivery Note", {}).get(
                    row.delivery_note, frappe._dict()
                )
                row.item_row = row.dn_detail

            # get buying amount
            if row.item_code in product_bundles:
                row.buying_amount = flt(
                    self.get_buying_amount_from_product_bundle(row, product_bundles[row.item_code]),
                    self.currency_precision,
                )
            else:
                row.buying_amount = flt(self.get_buying_amount(row, row.item_code), self.currency_precision)

            if grouped_by_invoice:
                if row.indent == 1.0:
                    buying_amount += row.buying_amount
                elif row.indent == 0.0:
                    row.buying_amount = buying_amount
                    buying_amount = 0

            # get buying rate
            if flt(row.qty):
                row.buying_rate = flt(row.buying_amount / flt(row.qty), self.float_precision)
                row.base_rate = flt(row.base_amount / flt(row.qty), self.float_precision)
            else:
                if self.is_not_invoice_row(row):
                    row.buying_rate, row.base_rate = 0.0, 0.0

            # calculate gross profit
            row.gross_profit = flt(row.base_amount - row.buying_amount, self.currency_precision)
            if row.base_amount:
                row.gross_profit_percent = flt(
                    (row.gross_profit / row.base_amount) * 100.0, self.currency_precision
                )
            else:
                row.gross_profit_percent = 0.0

            # add to grouped
            self.grouped.setdefault(row.get(scrub(self.filters.group_by)), []).append(row)

        if self.grouped:
            self.get_average_rate_based_on_group_by()
        #print(self.si_list)    

    def get_average_rate_based_on_group_by(self):
        for key in list(self.grouped):
            if self.filters.get("group_by") != "Invoice":
                for i, row in enumerate(self.grouped[key]):
                    if i == 0:
                        new_row = row
                    else:
                        new_row.qty += flt(row.qty)
                        new_row.buying_amount += flt(row.buying_amount, self.currency_precision)
                        new_row.base_amount += flt(row.base_amount, self.currency_precision)
                new_row = self.set_average_rate(new_row)
                self.grouped_data.append(new_row)
            else:
                for i, row in enumerate(self.grouped[key]):
                    if row.indent == 1.0:
                        if (
                            row.parent in self.returned_invoices and row.item_code in self.returned_invoices[row.parent]
                        ):
                            returned_item_rows = self.returned_invoices[row.parent][row.item_code]
                            for returned_item_row in returned_item_rows:
                                row.qty += flt(returned_item_row.qty)
                                row.base_amount += flt(returned_item_row.base_amount, self.currency_precision)
                            row.buying_amount = flt(flt(row.qty) * flt(row.buying_rate), self.currency_precision)
                        if flt(row.qty) or row.base_amount:
                            row = self.set_average_rate(row)
                            self.grouped_data.append(row)

    def is_not_invoice_row(self, row):
        return (self.filters.get("group_by") == "Invoice" and row.indent != 0.0) or self.filters.get(
            "group_by"
        ) != "Invoice"

    def set_average_rate(self, new_row):
        self.set_average_gross_profit(new_row)
        new_row.buying_rate = (
            flt(new_row.buying_amount / new_row.qty, self.float_precision) if new_row.qty else 0
        )
        new_row.base_rate = (
            flt(new_row.base_amount / new_row.qty, self.float_precision) if new_row.qty else 0
        )
        return new_row

    def set_average_gross_profit(self, new_row):
        new_row.gross_profit = flt(new_row.base_amount - new_row.buying_amount, self.currency_precision)
        new_row.gross_profit_percent = (
            flt(((new_row.gross_profit / new_row.base_amount) * 100.0), self.currency_precision)
            if new_row.base_amount
            else 0
        )
        new_row.buying_rate = (
            flt(new_row.buying_amount / flt(new_row.qty), self.float_precision) if flt(new_row.qty) else 0
        )
        new_row.base_rate = (
            flt(new_row.base_amount / flt(new_row.qty), self.float_precision) if flt(new_row.qty) else 0
        )

    def get_returned_invoice_items(self):
        returned_invoices = frappe.db.sql(
            """
            select
                si.name, si_item.item_code, si_item.stock_qty as qty, si_item.base_net_amount as base_amount, si.return_against
            from
                `tabSales Invoice` si, `tabSales Invoice Item` si_item
            where
                si.name = si_item.parent
                and si.docstatus = 1
                and si.is_return = 1
        """,
            as_dict=1,
        )

        self.returned_invoices = frappe._dict()
        for inv in returned_invoices:
            self.returned_invoices.setdefault(inv.return_against, frappe._dict()).setdefault(
                inv.item_code, []
            ).append(inv)

    def skip_row(self, row):
        if self.filters.get("group_by") != "Invoice":
            if not row.get(scrub(self.filters.get("group_by", ""))):
                return True

        return False

    def get_buying_amount_from_product_bundle(self, row, product_bundle):
        buying_amount = 0.0
        for packed_item in product_bundle:
            if packed_item.get("parent_detail_docname") == row.item_row:
                buying_amount += self.get_buying_amount(row, packed_item.item_code)

        return flt(buying_amount, self.currency_precision)

    def get_buying_amount(self, row, item_code):
        # IMP NOTE
        # stock_ledger_entries should already be filtered by item_code and warehouse and
        # sorted by posting_date desc, posting_time desc
        if item_code in self.non_stock_items and (row.project or row.cost_center):
            # Issue 6089-Get last purchasing rate for non-stock item
            item_rate = self.get_last_purchase_rate(item_code, row)
            return flt(row.qty) * item_rate

        else:
            my_sle = self.sle.get((item_code, row.warehouse))
            if (row.update_stock or row.dn_detail) and my_sle:
                parenttype, parent = row.parenttype, row.parent
                if row.dn_detail:
                    parenttype, parent = "Delivery Note", row.delivery_note

                for i, sle in enumerate(my_sle):
                    # find the stock valution rate from stock ledger entry
                    if (
                        sle.voucher_type == parenttype
                        and parent == sle.voucher_no
                        and sle.voucher_detail_no == row.item_row
                    ):
                        previous_stock_value = len(my_sle) > i + 1 and flt(my_sle[i + 1].stock_value) or 0.0

                        if previous_stock_value:
                            return (previous_stock_value - flt(sle.stock_value)) * flt(row.qty) / abs(flt(sle.qty))
                        else:
                            return flt(row.qty) * self.get_average_buying_rate(row, item_code)
            else:
                return flt(row.qty) * self.get_average_buying_rate(row, item_code)

        return 0.0

    def get_average_buying_rate(self, row, item_code):
        args = row
        if not item_code in self.average_buying_rate:
            args.update(
                {
                    "voucher_type": row.parenttype,
                    "voucher_no": row.parent,
                    "allow_zero_valuation": True,
                    "company": self.filters.company,
                }
            )

            average_buying_rate = get_incoming_rate(args)
            self.average_buying_rate[item_code] = flt(average_buying_rate)

        return self.average_buying_rate[item_code]

    def get_last_purchase_rate(self, item_code, row):
        purchase_invoice = frappe.qb.DocType("Purchase Invoice")
        purchase_invoice_item = frappe.qb.DocType("Purchase Invoice Item")

        query = (
            frappe.qb.from_(purchase_invoice_item)
            .inner_join(purchase_invoice)
            .on(purchase_invoice.name == purchase_invoice_item.parent)
            .select(purchase_invoice_item.base_rate / purchase_invoice_item.conversion_factor)
            .where(purchase_invoice.docstatus == 1)
            .where(purchase_invoice.posting_date <= self.filters.to_date)
            .where(purchase_invoice_item.item_code == item_code)
        )

        if row.project:
            query.where(purchase_invoice_item.project == row.project)

        if row.cost_center:
            query.where(purchase_invoice_item.cost_center == row.cost_center)

        query.orderby(purchase_invoice.posting_date, order=frappe.qb.desc)
        query.limit(1)
        last_purchase_rate = query.run()

        return flt(last_purchase_rate[0][0]) if last_purchase_rate else 0

    def load_invoice_items(self):
        conditions2 = ""
        if self.filters.company:
            conditions2 += " and company = %(company)s"
        if self.filters.get("cost_center"):
            conditions2 += " and `tabSales Invoice Item`.cost_center IN %(cost_center)s"    
        if self.filters.month:
            long_month_name = self.filters.get("month")
            datetime_object = datetime.datetime.strptime(long_month_name, "%B")
            month_number = datetime_object.month
            conditions2 += " and ((MONTH(posting_date) = " + str(month_number) + " and YEAR(posting_date) = " + str(int(self.filters.get("year"))-1) + "))"
            
        
        self.si_list = frappe.db.sql(
            """
            select
                `tabSales Invoice Item`.parenttype, `tabSales Invoice Item`.parent,
                `tabSales Invoice`.posting_date, `tabSales Invoice`.posting_time,
                `tabSales Invoice`.project, `tabSales Invoice`.update_stock,
                `tabSales Invoice`.customer, `tabSales Invoice`.customer_group,
                `tabSales Invoice`.territory, `tabSales Invoice Item`.item_code,
                `tabSales Invoice Item`.item_name, `tabSales Invoice Item`.description,
                `tabSales Invoice Item`.warehouse, `tabSales Invoice Item`.item_group,
                `tabSales Invoice Item`.brand, `tabSales Invoice Item`.dn_detail,
                `tabSales Invoice Item`.delivery_note, `tabSales Invoice Item`.stock_qty as qty,
                `tabSales Invoice Item`.base_net_rate, `tabSales Invoice Item`.base_net_amount,
                `tabSales Invoice Item`.name as "item_row", `tabSales Invoice`.is_return,
                `tabSales Invoice Item`.cost_center
            
            from
                `tabSales Invoice` inner join `tabSales Invoice Item`
                    on `tabSales Invoice Item`.parent = `tabSales Invoice`.name
                
            where
                `tabSales Invoice`.docstatus=1 and `tabSales Invoice`.is_opening!='Yes' {conditions} {match_cond}
            order by
                `tabSales Invoice`.posting_date desc, `tabSales Invoice`.posting_time desc""".format(
                conditions=conditions2,
                match_cond=get_match_cond("Sales Invoice"),
            ),
            self.filters,
            as_dict=1,
            debug=0,
        )



    def group_items_by_invoice(self):
        """
        Turns list of Sales Invoice Items to a tree of Sales Invoices with their Items as children.
        """

        parents = []

        for row in self.si_list:
            if row.parent not in parents:
                parents.append(row.parent)

        parents_index = 0
        for index, row in enumerate(self.si_list):
            if parents_index < len(parents) and row.parent == parents[parents_index]:
                invoice = self.get_invoice_row(row)
                self.si_list.insert(index, invoice)
                parents_index += 1

            else:
                # skipping the bundle items rows
                if not row.indent:
                    row.indent = 1.0
                    row.parent_invoice = row.parent
                    row.invoice_or_item = row.item_code

                    if frappe.db.exists("Product Bundle", row.item_code):
                        self.add_bundle_items(row, index)

    def get_invoice_row(self, row):
        return frappe._dict(
            {
                "parent_invoice": "",
                "indent": 0.0,
                "invoice_or_item": row.parent,
                "parent": None,
                "posting_date": row.posting_date,
                "posting_time": row.posting_time,
                "project": row.project,
                "update_stock": row.update_stock,
                "customer": row.customer,
                "customer_group": row.customer_group,
                "item_code": None,
                "item_name": None,
                "description": None,
                "warehouse": None,
                "item_group": None,
                "brand": None,
                "dn_detail": None,
                "delivery_note": None,
                "qty": None,
                "item_row": None,
                "is_return": row.is_return,
                "cost_center": row.cost_center,
                "base_net_amount": frappe.db.get_value("Sales Invoice", row.parent, "base_net_total"),
            }
        )

    def add_bundle_items(self, product_bundle, index):
        bundle_items = self.get_bundle_items(product_bundle)

        for i, item in enumerate(bundle_items):
            bundle_item = self.get_bundle_item_row(product_bundle, item)
            self.si_list.insert((index + i + 1), bundle_item)

    def get_bundle_items(self, product_bundle):
        return frappe.get_all(
            "Product Bundle Item", filters={"parent": product_bundle.item_code}, fields=["item_code", "qty"]
        )

    def get_bundle_item_row(self, product_bundle, item):
        item_name, description, item_group, brand = self.get_bundle_item_details(item.item_code)

        return frappe._dict(
            {
                "parent_invoice": product_bundle.item_code,
                "indent": product_bundle.indent + 1,
                "parent": None,
                "invoice_or_item": item.item_code,
                "posting_date": product_bundle.posting_date,
                "posting_time": product_bundle.posting_time,
                "project": product_bundle.project,
                "customer": product_bundle.customer,
                "customer_group": product_bundle.customer_group,
                "item_code": item.item_code,
                "item_name": item_name,
                "description": description,
                "warehouse": product_bundle.warehouse,
                "item_group": item_group,
                "brand": brand,
                "dn_detail": product_bundle.dn_detail,
                "delivery_note": product_bundle.delivery_note,
                "qty": (flt(product_bundle.qty) * flt(item.qty)),
                "item_row": None,
                "is_return": product_bundle.is_return,
                "cost_center": product_bundle.cost_center,
            }
        )

    def get_bundle_item_details(self, item_code):
        return frappe.db.get_value(
            "Item", item_code, ["item_name", "description", "item_group", "brand"]
        )

    def load_stock_ledger_entries(self):
        res = frappe.db.sql(
            """select item_code, voucher_type, voucher_no,
                voucher_detail_no, stock_value, warehouse, actual_qty as qty
            from `tabStock Ledger Entry`
            where company=%(company)s and is_cancelled = 0
            order by
                item_code desc, warehouse desc, posting_date desc,
                posting_time desc, creation desc""",
            self.filters,
            as_dict=True,
        )
        self.sle = {}
        for r in res:
            if (r.item_code, r.warehouse) not in self.sle:
                self.sle[(r.item_code, r.warehouse)] = []

            self.sle[(r.item_code, r.warehouse)].append(r)

    def load_product_bundle(self):
        self.product_bundles = {}

        for d in frappe.db.sql(
            """select parenttype, parent, parent_item,
            item_code, warehouse, -1*qty as total_qty, parent_detail_docname
            from `tabPacked Item` where docstatus=1""",
            as_dict=True,
        ):
            self.product_bundles.setdefault(d.parenttype, frappe._dict()).setdefault(
                d.parent, frappe._dict()
            ).setdefault(d.parent_item, []).append(d)

    def load_non_stock_items(self):
        self.non_stock_items = frappe.db.sql_list(
            """select name from tabItem
            where is_stock_item=0"""
        )
