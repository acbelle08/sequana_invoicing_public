"""
Class represenation of the invoices table
"""

import sqlite3
import sys

class Invoice:
    def __init__(self, invoice_id, con=None):
        self.con: sqlite3.Connection
        if con:
            self.con = con
        else:
            self.con = sqlite3.connect('invoicing.db')
        self.cur = self.con.cursor()
        self._invoice_id = invoice_id
        self.cur.execute(
            f"select invoice_timestamp, first_month, last_month, \
            chargeable_account, user_id, amount_payable, reference_text, sent, paid \
            from invoices where invoice_id={self._invoice_id}"
            )
        results = self.cur.fetchall()
        assert(len(results) == 1)
        (
            self._invoice_timestamp, self._first_month, self._last_month,
            self._chargeable_account, self._user_id,
            self._amount_payable, self._reference_text, self._sent, self._paid
            ) = results[0]
        
        # Properties used when generating an invoice
        self.balance = 0.0
        self.total_staff_hours = 0.0
        self.total_staff_cost = 0.0
        self.total_staff_subsidy_amount = 0.0
        self.charges_count = 0
        self.total_consumable_cost = 0.0
        self.total_comsumables_subsidy_amount = 0.0
        self.total_consumables_amount_payable = 0.0
        self.credit_used_against_this_invoice = 0.0

    @property
    def invoice_timestamp(self):
        return self._invoice_timestamp

    @property
    def paid(self):
        return self._paid
        
    @property
    def sent(self):
        return self._sent

    @property
    def reference_text(self):
        return self._reference_text

    @property
    def invoice_id(self):
        return self._invoice_id

    @property
    def amount_payable(self):
        self.cur.execute(
            f"select amount_payable \
            from invoices where invoice_id={self._invoice_id}"
            )
        return self.cur.fetchone()[0]

    @amount_payable.setter
    def amount_payable(self, amount):
        assert(amount >= 0)
        self.cur.execute(f"update invoices set amount_payable={amount} where invoice_id={self._invoice_id}")
        self.con.commit()

    @property
    def user_id(self):
        return self._user_id

    def _get_consumables_charges(self):
        # Get the consumable charges for the invoice
        self.cur.execute(f"select charge_id, unit_cost, quantity, subsidy_percent from consumable_charges where invoice_id = {self.invoice_id}")
        # We want a total value for the charges
        results = self.cur.fetchall()
        total_consumable_charges = 0
        for charge_id, unit_cost, quantity, subsidy_percent in results:
            total_consumable_charges += (quantity - (quantity * subsidy_percent/100)) * unit_cost
        
        return total_consumable_charges