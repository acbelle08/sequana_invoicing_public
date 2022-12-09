"""
Class representation of te users table from the dadtabase.
"""

import sqlite3
from invoices import Invoice
import sys

class User:
    def __init__(self, user_id, con):
        self.con: sqlite3.Connection
        self.con = con

        self.cur = self.con.cursor()
        self._user_id = user_id
        self.cur.execute(
            f"select email, first_name, last_name, \
            staff_subsidy_percent, consumable_subsidy_percent \
            from users where user_id={self._user_id}"
            )
        results = self.cur.fetchall()
        assert(len(results) == 1)
        (
            self._email, self._first_name, self._last_name,
            self._staff_subsidy_percent, self._consumable_subsidy_percent
            ) = results[0]

        self._invoices = self._get_ordered_list_of_invoices()

    @property
    def available_credit(self):
        """
        The credit available to a user is simply the sum of their paid credit invoices
        with the sum of the credit_debit objects subtracted.
        """
        self.cur.execute("select sum(amount_payable) from invoices where user_id=:user_id and invoice_type='credit' and sent=:sent and paid=:paid", {"user_id":self._user_id, "sent":True, "paid":True})
        sum_result = self.cur.fetchone()[0]
        if sum_result is None:
            return 0.00
        else:
            total_credit = float(sum_result)
        
        self.cur.execute("select sum(amount) from credit_debit where user_id=:user_id", {"user_id":self._user_id})
        sum_result = self.cur.fetchone()[0]
        if sum_result is None:
            return total_credit
        else:
            total_debit = float(sum_result)
            return total_credit - total_debit

    @property
    def user_id(self):
        return self._user_id

    @property
    def invoices(self):
        return self._get_ordered_list_of_invoices()

    @property
    def sent_invoices(self):
        return self._get_ordered_list_of_sent_invoices()

    @property
    def staff_subsidy_percent(self):
        return self._staff_subsidy_percent

    @property
    def consumable_subsidy_percent(self):
        return self._consumable_subsidy_percent

    @property
    def last_name(self):
        return self._last_name

    @property
    def first_name(self):
        return self._first_name

    @property
    def email(self):
        return self._email

    def _get_ordered_list_of_invoices(self):
        self.cur.execute(f"select invoice_id from invoices where user_id = {self._user_id} order by invoice_id asc")
        results = self.cur.fetchall()
        if len(results) == 0:
            return []
        else:
            return [Invoice(int(_[0]), self.con) for _ in results]

    def _get_ordered_list_of_sent_invoices(self):
        self.cur.execute(f"select invoice_id from invoices where user_id = {self._user_id} and sent = 1 order by invoice_id asc")
        results = self.cur.fetchall()
        if len(results) == 0:
            return []
        else:
            return [Invoice(int(_[0]), self.con) for _ in results]

    