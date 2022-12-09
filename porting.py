"""
This script will be responsible for porting the data from the old database structure
to the new database structure. It will essentially be a one of use
so will be minimally commented.

I have made csv outputs of the relevant datatables:

users
user_balances
projects
invoices
charges
consumable_charges
credit_invoices
"""

import sqlite3
import pandas as pd

class DBPorting:
    def __init__(self):
        self.con = sqlite3.connect('new_invoicing.db')
        self.cur = self.con.cursor()
        self.users_df = pd.read_csv("users.csv")
        self.user_balances_df = pd.read_csv("user_balances.csv")
        self.user_balances_df.set_index("user_id", inplace=True)
        self.projects_df = pd.read_csv("projects.csv")
        self.invoices_df = pd.read_csv("invoices.csv")
        self.staff_time_charges_df = pd.read_csv("staff_time_charges.csv")
        self.consumable_charges_df = pd.read_csv("consumable_charges.csv")

        # Users including balance
        for ind, ser in self.users_df.iterrows():
            # Get the balance from the user_balances
            # NB the insert statements don't work using f-strings.
            current_balance = self.user_balances_df.at[int(ser["user_id"]), "current_balance_euro"]
            self.cur.execute("insert into users (user_id, email, first_name, last_name, staff_subsidy_percent, consumable_subsidy_percent, current_balance_euro) values (:user_id, :email, :first_name, :last_name, :staff_subsidy_percent, :consumable_subsidy_percent, :current_balance)", 
            {"user_id": ser['user_id'], "email": ser['email'], "first_name": ser['first_name'], "last_name": ser['last_name'], "staff_subsidy_percent": ser['staff_subsidy_percent'], "consumable_subsidy_percent": ser['consumable_subsidy_percent'], "current_balance": current_balance})

        # Projects
        for ind, ser in self.projects_df.iterrows():
            self.cur.execute(f"insert into projects (project_id, project_type, project_title, user_id) values (:project_id, :project_type, :project_title, :user_id)",
            {"project_id": ser['project_id'], "project_type": ser['project_type'], "project_title": ser['project_title'], "user_id": ser['user_id']})
        
        # Invoices
        for ind, ser in self.invoices_df.iterrows():
            self.cur.execute(f"insert into invoices (invoice_id, invoice_timestamp, first_month, last_month, chargeable_account, user_id, amount_payable, sent, paid, invoice_type) values \
                (:invoice_id, :invoice_timestamp, :first_month, :last_month, :chargeable_account, :user_id, :amount_payable, :sent, :paid, 'debit')", 
                {"invoice_id": ser['invoice_id'], "invoice_timestamp": ser['invoice_timestamp'], 
                "first_month": ser['first_month'], "last_month": ser['last_month'], "chargeable_account": ser['chargeable_account'],
                "user_id": ser['user_id'], "amount_payable": ser['amount_payable'], "sent": ser['sent'], "paid": ser['paid']}
                )

        # Staff time charges
        for ind, ser in self.staff_time_charges_df.iterrows():
            self.cur.execute(f"insert into staff_time_charges (charge_id, staff_hours, staff_hourly_rate_eur, subsidy_percent, invoice_id, project_id) values \
                (:charge_id, :staff_hours, :staff_hourly_rate_eur, :subsidy_percent, :invoice_id, :project_id)",
                {"charge_id": ser['charge_id'], "staff_hours": ser['staff_hours'], "staff_hourly_rate_eur": ser['staff_hourly_rate_eur'], "subsidy_percent": ser['subsidy_percent'], "invoice_id": ser['invoice_id'], "project_id": ser['project_id']})

        # Consumable charges
        for ind, ser in self.consumable_charges_df.iterrows():
            self.cur.execute(f"insert into consumable_charges (charge_id, name, unit_cost, subsidy_percent, quantity, subsidy_percent, date, invoice_id, project_id, PPMS_reference) values \
                (:charge_id, :name, :unit_cost, :subsidy_percent, :quantity, :subsidy_percent, :date, :invoice_id, :project_id, :PPMS_reference)", 
                {"charge_id": ser['charge_id'], "name": ser['name'], "unit_cost": ser['unit_cost'], "subsidy_percent": ser['subsidy_percent'],
                "quantity": ser['quantity'], "subsidy_percent": ser['subsidy_percent'], "date": ser['date'],
                "invoice_id": ser['invoice_id'], "project_id": ser['project_id'], "PPMS_reference": ser['PPMS_reference']})

        self.con.commit()

        

DBPorting()