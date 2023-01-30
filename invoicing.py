#!/usr/bin/env python3
"""
Script for making the SequAna invoices.

TODO write output of database to csv file.

"""

from mailbox import FormatError
from docxtpl import DocxTemplate
import argparse
import sqlite3
import os
import pandas as pd
from difflib import get_close_matches
import sys
import datetime
import shutil
from datetime import timezone
from invoices import Invoice
from users import User
from projects import Project


class Invoicing:
    def __init__(self):
        self.backup_date_time_str = str(datetime.datetime.now(timezone.utc)).replace(" ", "T").replace("-","").replace(":","").split(".")[0]
        self.args = self._parse_args()
        
    def _init_db(self):
        self.con = sqlite3.connect('invoicing.db')
        self.cur = self.con.cursor()

    def _output_xlsx_of_database(self):
        """
        As a fail safe we will output the database to csv and save this in the db_backup
        directory
        """
        # Get a list of all of the tables in the database
        self.cur.execute("SELECT name FROM sqlite_master WHERE type='table'")
        tables = [_[0] for _ in self.cur.fetchall()]

        # Create a writer so that we can write the pandas dataframes to the excel
        # on separate sheets
        try:
            db_csv_path = os.path.join(self.db_backup_dir, f'{self.backup_date_time_str}_db_backup.xlsx')
        except AttributeError:
            if not os.path.exists("db_backup"):
                os.makedirs("db_backup")
            db_csv_path = os.path.join("db_backup", f'{self.backup_date_time_str}_db_backup.xlsx')
        
        with pd.ExcelWriter(db_csv_path, engine='openpyxl') as writer:

            # Make a dataframe from each of the tables in the sqlite database
            for table in tables:
                df = pd.read_sql_query(f"SELECT * FROM {table}", self.con)
                df.to_excel(writer, sheet_name=f'{table}')

    def _make_new_user(self):
        self._init_db()
        first_name, last_name, email, staff_subsidy, consumable_subsidy = self._get_first_name_email_subsidy_of_user()
        self.cur.execute(
                "INSERT INTO users (first_name, last_name, email, staff_subsidy_percent, consumable_subsidy_percent) VALUES (:first_name, :user_last_name, :email, :staff_subsidy, :consumable_subsidy)",
                {"first_name":first_name, "user_last_name":last_name, "email":email, "staff_subsidy":staff_subsidy, "consumable_subsidy":consumable_subsidy}
                )
        self.con.commit()
        print(f"User {first_name} {last_name} successfully added to database.") 

    def _set_invoices_paid(self):
        """
        Set the paid status of one or more invoices to True
        """
        self._init_db()
    
        # Get a dataframe where each row is a credit invoice to be created.
        self.credit_df = self._do_credit_invoice_csv_qc()

        # Check that all credit invoices listed in the df exist
        self._check_that_all_invoices_in_credit_input_df_exist_sent()

        # Set the invoices to paid
        self._set_invoice_to_paid()

        self._output_xlsx_of_database()

    def _set_invoice_to_paid(self):
        for ind, ser in self.credit_df.iterrows():
            self.cur.execute(
                    "select invoice_id from invoices inner join users on \
                    invoices.user_id=users.user_id where \
                    sent=:sent and \
                    paid=:paid and users.email=:email and \
                    invoices.amount_payable=:amount_payable", 
                    {"sent":True, "paid":False, "email":ser["user_email"], "amount_payable": ser["amount_payable"]}
                    )
            invoice_id = self.cur.fetchone()[0]
            invoice = Invoice(invoice_id=invoice_id, con=self.con)
            self.cur.execute("update invoices set paid=:paid where invoice_id=:invoice_id", {"paid": True, "invoice_id": invoice_id})
            self.con.commit()
            print(f"Invoice {invoice.invoice_id} ({invoice.reference_text}) set to paid")

    def _set_invoice_to_sent(self):
        """
        Set the sent status of one or more invoices to True
        """

        self._init_db()

        # Get a dataframe where each row is a credit invoice to be created.
        self.credit_df = self._do_credit_invoice_csv_qc()

        # Check that all credit invoices listed in the df exist
        self._check_that_all_invoices_in_credit_input_df_exist_not_sent()

        # Set the invoices to sent
        self._set_invoices_to_sent()

        self._output_xlsx_of_database()
        
    def _set_invoices_to_sent(self):
        for ind, ser in self.credit_df.iterrows():
            self.cur.execute(
                    "select invoice_id from invoices inner join users on \
                    invoices.user_id=users.user_id where \
                    sent=:sent and \
                    paid=:paid and users.email=:email and \
                    invoices.amount_payable=:amount_payable", 
                    {"sent":False, "paid":False, "email":ser["user_email"], "amount_payable": ser["amount_payable"]}
                    )
            invoice_id = self.cur.fetchone()[0]
            invoice = Invoice(invoice_id=invoice_id, con=self.con)
            self.cur.execute("update invoices set sent=:sent where invoice_id=:invoice_id", {"sent": True, "invoice_id": invoice_id})
            self.con.commit()
            print(f"Invoice {invoice.invoice_id} ({invoice.reference_text}) set to sent")
        
    def _check_that_all_invoices_in_credit_input_df_exist_sent(self):
        for ind, ser in self.credit_df.iterrows():
            self.cur.execute(
                    "select invoice_id from invoices inner join users on \
                    invoices.user_id=users.user_id where \
                    sent=:sent and \
                    paid=:paid and users.email=:email and \
                    invoices.amount_payable=:amount_payable", 
                    {"sent":True, "paid":False, "email":ser["user_email"], "amount_payable": ser["amount_payable"]}
                    )
            invoice_ids = [_[0] for _ in self.cur.fetchall()]
            if len(invoice_ids) > 1:
                print(f"More than 1 invoice was found matching email {ser['user_email']} and amount {ser['amount_payable']} that is sent but not paid.")
                sys.exit("Exiting")
            if len(invoice_ids) == 0:
                print(f"Cannot find an invoice matching email {ser['user_email']} and amount {ser['amount_payable']} that is sent but not paid.")
                sys.exit

    def _check_that_all_invoices_in_credit_input_df_exist_not_sent(self):
        for ind, ser in self.credit_df.iterrows():
            self.cur.execute(
                    "select invoice_id from invoices inner join users on \
                    invoices.user_id=users.user_id where \
                    sent=:sent and \
                    paid=:paid and users.email=:email and \
                    invoices.amount_payable=:amount_payable", 
                    {"sent":False, "paid":False, "email":ser["user_email"], "amount_payable": ser["amount_payable"]}
                    )
            invoice_ids = [_[0] for _ in self.cur.fetchall()]
            if len(invoice_ids) > 1:
                print(f"More than 1 invoice was found matching email {ser['user_email']} and amount {ser['amount_payable']} that is not sent nor paid.")
                sys.exit("Exiting")
            if len(invoice_ids) == 0:
                print(f"Cannot find an invoice matching email {ser['user_email']} and amount {ser['amount_payable']} that is not sent nor paid.")
                sys.exit

    def _init_create_credit_invoices(self):
        """
        Make one or more credit invoices based on user input details from csv
        based on template docx.
        """

        self._init_db()

        # Get a dataframe where each row is a credit invoice to be created.
        self.credit_df = self._do_credit_invoice_csv_qc()

        # Make the credit invoices if they don't already exist
        self._make_credit_invoices()

        self._output_xlsx_of_database()

    def _do_credit_invoice_csv_qc(self):
        c_inv_df = pd.read_csv(self.args.input)
        return c_inv_df

    def _check_all_users_in_input_csv_exist(self):
        # check that all of the user emails are found in the database
        # that are included in users input csv.
        self.cur.execute("SELECT email from users;")
        db_user_emails = [_[0] for _ in self.cur.fetchall()]
        input_users = self.credit_df["user_email"].to_list()
        missing_users = [_ for _ in input_users if _ not in db_user_emails]
        if missing_users:
            print(
                    ("WARNING: The following emails are related to a user in the database.\n"
                "Check the email is correct.\n"
                "If you need to add a new user you can do so using the adduser sub command.\n")
                )
            for missing in missing_users:
                print(f"\t{missing}")
        
            print("The following user emails are currently in the database:")
            for u_email in db_user_emails:
                print(f"\t{u_email}")

            sys.exit("Exiting")
    
    def _check_credit_invoices_dont_already_exist(self):
        # Check to see that such an invoice doesn't already exist
        # if not create the invoice in the db.
        for ind, ser in self.credit_df.iterrows():
            user_email = ser["user_email"]
            amount = ser["amount_payable"]
            self.cur.execute(
                    "SELECT invoices.invoice_id, users.email, users.user_id from \
                        invoices inner join users on invoices.user_id=users.user_id where \
                            invoices.amount_payable=:amount and users.email=:user_email and invoices.sent=:sent and invoices.invoice_type=:invoice_type", 
                            {"amount": amount, "user_email": user_email, "sent":False, "invoice_type": "credit"}
                            )
            results = self.cur.fetchall()
            if results:
                print(("One or more invoices already exist matching your inputs. "
                "If you really want to make another credit invoice that already "
                "matches an existing credit invoice, "
                "then make sure that the existing invoice has been marked "
                "as sent using the subcommand mark_invoice_sent."))
                print("invoice_id\temail\tuser_id")
                for result in results:
                    print("\t".join([str(_) for _ in result]))
                sys.exit("Exiting")

    def _make_credit_invoices(self):
        # Check that the credit template exists
        if not os.path.exists(self.args.template):
            sys.exit(f"Cannot locate {self.args.template}. Exiting.")
        
        self._check_all_users_in_input_csv_exist()

        self._check_credit_invoices_dont_already_exist()
        
        # Create the credit invoice and print confirmation out to the terminal
        for ind, ser in self.credit_df.iterrows():
            user_email = ser["user_email"]
            self.cur.execute(f"select user_id from users where email=:user_email", {"user_email": user_email})
            user_id = self.cur.fetchone()[0]
            user = User(user_id, con=self.con)
            amount = ser["amount_payable"]
            invoice_timestamp = str(datetime.datetime.now())
            first_month = last_month = invoice_timestamp.split(" ")[0].split("-")[0] + invoice_timestamp.split(" ")[0].split("-")[1]
            chargeable_account = self.args.chargeable_account
            
            # First insert the record
            self.cur.execute("insert into invoices \
            (invoice_timestamp, first_month, last_month, chargeable_account, user_id, invoice_type, amount_payable, sent, paid) \
            values(:invoice_timestamp, :first_month, :last_month, :chargeable_account, :user_id, :invoice_type, :amount_payable, :sent, :paid)", 
                {
                    "invoice_timestamp": invoice_timestamp, "first_month": first_month, "last_month": last_month, 
                    "chargeable_account": chargeable_account, "user_id": user.user_id, "invoice_type": "credit", 
                    "amount_payable": amount, "sent": False, "paid": False
                }
            )

            self.con.commit()
            # Then get the id of the record that we just made and use it to make the reference text
            # Example: SequAna credit; invoice C2; Yamada, Norico
            self.cur.execute(f"select invoice_id from invoices where invoice_timestamp=:invoice_timestamp", {"invoice_timestamp": invoice_timestamp})
            invoice_id = self.cur.fetchone()[0]
            reference = f"SequAna credit; invoice C{invoice_id}; {user.last_name}, {user.first_name}"
            # update the reference text
            self.cur.execute(
                    "update invoices set reference_text=:reference where invoice_id=:invoice_id", 
                    {"reference": reference, "invoice_id": invoice_id}
                    )
            self.con.commit()
            
            # Now we need to populate the credit invoice template
            self.credit_doc = DocxTemplate(self.args.template)
            self.credit_context = {}

            invoice_date = invoice_timestamp.split(" ")[0].replace("-", "")
            self.credit_context["invoice_date"] = invoice_date
            self.credit_context["user_name"] = f"{user.last_name}, {user.first_name}"
            self.credit_context["user_email"] = user.email
            self.credit_context["invoice_id"] = f"C{invoice_id}"
            self.credit_context["chargeable_account"] = chargeable_account
            self.credit_context["amount"] = amount
            self.credit_doc.render(self.credit_context)
            outpath = os.path.join(self.args.output_dir, f"{invoice_date}_{user.last_name.replace(' ', '_')}_SequAna_Credit_Invoice_C{invoice_id}.docx")
            
            if os.path.exists(outpath):
                if self._get_n_y_user_response(question_text=f"\n\n{outpath} already exists.\nOverwrite? [y/n]: ") == "y":
                    self.credit_doc.save(outpath)
                    print(f"\nWriting {outpath}.")
                else:
                    print("Skipping credit invoice output")
                    continue
            else:
                self.credit_doc.save(outpath)
                print(f"\nWriting {outpath}.")

    def _init_create_invoices(self):
        """
        Make standard charge invoives according to the user provided inputs
        """

        self._init_db()
        

        (
            self.month_range, self.hours_charged_df, self.user_last_names_to_invoice,
            self.chargeable_account, self.staff_hourly_rate_eur, self.output_dir
        ) = self._do_argument_qc()

        # We need to make the consumable charges for the period
        # so that they are available when we work out the balances
        # on the per user basis.
        if self.args.PPMS_input_consumables_csv:
            self._make_consumable_charges()
        
        self._make_user_invoices()

        self._back_up_db()

        self._output_xlsx_of_database()

    def _back_up_db(self):
        """
        Back up the database by simply copying the current invoices.db file into the specified db_backup directory
        """
        backup_db_path = os.path.join(self.db_backup_dir, f'{self.backup_date_time_str}_db_backup_invoicing.db')
        print(f"\n\nBacking up invoicing.db to {backup_db_path}")
        shutil.copyfile('invoicing.db', backup_db_path)

    def _make_consumable_charges(self):
        """
        Create new consumable charges for the rows of the consumables input if they don't already exist
        This means creating new projects and invoices if they don't exist that the consumable
        charges can be associated to.
        """
        print("\nChecking whether consumable charges in input already exist and creating if not.\n")
        # Now we need to check whether a consumables object exists for each of the rows of the df
        for ind, ser in self.consumables_df.iterrows():
            if self._project_exists(project_name=ser["Project name"]):
                project_id, user_id = self._get_project_id_user_id_from_project(project_name=ser["Project name"])

            # First check to see if there is an index matching the Project name
                if self._invoice_exists(user_id=user_id):
                    invoice_id = self._get_invoice_id(user_id=user_id)
                    # Then an invoice exists and it is possible that a consumable charge already exists that we can check for
                    if self._consumable_charge_exits(project_id=project_id, invoice_id=invoice_id, name=ser["Consumable name"], date=ser["month"], unit_cost=ser["Unit price"], quantity=ser["Quantity"], ref=ser["Ref."]):
                        pass
                    else:
                        # create consumable charge with invoice and project ids
                        self._create_consumable_charge(ser)
                else:
                    self._create_consumable_charge(ser)
            else:
                self._create_consumable_charge(ser)
        print("\nFinished checking consumable charges from input.")

    def _make_user_invoices(self):
        for user_last_name in self.user_last_names_to_invoice:
            
            self.current_user = self._get_or_make_user_for_invoicing(user_last_name)
            
            self.current_invoice = self._get_or_make_invoice()

            if self.current_invoice.sent:
                # Then this invoice has already been sent and we should not be modifying it so we will skip
                print("An invoice already exists that has been sent.\nSkipping this user and moving to next.\n")
                continue
            # Find projects that are of the user in the PPMS input and check to see if they match projects in the database
            # If they do not match then create the project in a similar way to users above
            # For each project, make sure that a charge exists or create if it does not.
            # Delete any existing charges from the db for the user for this period that aren't included
            # in the current PPMS input.
            
            # Keep track of the db charge objects that are related to the PPMS input
            # NB it may be that there were other charges that were already in the db that were
            # related to the invoice that may have been wrong and we don't want to include
            # these. As such we should delete any charges that belong to the current invoice
            # that aren't in the self.charge_ids_related_to_PPMS_invoice
            self.charge_ids_related_to_PPMS_invoice = []
            for proj_of_user in [_ for _ in self.hours_charged_df.index if user_last_name.replace(" ", "_") in _]:
                self.current_project = self._get_or_make_project(proj_of_user)
                
                self._get_or_make_charge(proj_of_user)
            
            self._check_for_and_delete_unused_or_old_charges()

            self._populate_and_write_template()

        print("\nOutput of invoices complete.")

        print("Invoicing Complete.")

    def _apply_credits_for_user(self):
        """
        Credits are applied from users pre-paid balances
        users pre-paid balances. We create a credit_debit object to show that
        balance has been applied.
        """
        if self.current_invoice.balance > 0:
            # First check to see if a credit_debit has already been associated to this
            # invoice for this user. 
            self.cur.execute("select credit_id, amount from credit_debit where debit_invoice_id=:current_invoice_id and user_id=:current_user_id", {"current_invoice_id": self.current_invoice.invoice_id, "current_user_id": self.current_user.user_id})
            results = self.cur.fetchall()
            credit_used = 0

            if len(results) == 0:
                credit_used = self._make_credit_debit_object_for_invoices_if_available_credit()
            else:
                # Delete any credit_debit objects that already exist.
                self.cur.execute(
                        "delete from credit_debit where \
                            debit_invoice_id=:current_invoice_id and user_id=:current_user_id",
                            {
                                "current_invoice_id": self.current_invoice.invoice_id, 
                                "current_user_id": self.current_user.user_id
                                }
                                )
                self.con.commit()
                # Then make credit_debit object if available credit
                credit_used = self._make_credit_debit_object_for_invoices_if_available_credit()
                    
            # Then we need to make / apply credits to the invoice
            self.context["user_balance"] = {}
            self.context["user_balance"]["starting_available_credit"] = f"{self.current_user.available_credit + credit_used:.2f}"
            self.context["user_balance"]["applied_available_credit"] = f"{credit_used:.2f}"
            self.context["user_balance"]["closing_available_credit"] = f"{self.current_user.available_credit:.2f}"
            # update the balance to reflect the credit being added.
            self.context["balance"] = f"{self.current_invoice.balance:.2f}"
        else:
            # There is no need for credit_debit objects to be applied.
            self.context["user_balance"] = {}
            self.context["user_balance"]["starting_available_credit"] = f"{self.current_user.available_credit:.2f}"
            self.context["user_balance"]["applied_available_credit"] = "0.00"
            self.context["user_balance"]["closing_available_credit"] = f"{self.current_user.available_credit:.2f}"

    def _make_credit_debit_object_for_invoices_if_available_credit(self):
        # If a credit debit does not exist, then check to see if there is avialable credit
        if self.current_user.available_credit > 0:
            # If so then create a credit debit object and associate it to the current invoice.
            if self.current_user.available_credit > self.current_invoice.balance:
                credit_used = self.current_invoice.balance
                # If the amount of credit available is greater than the balance of the invoice
                # then make a credit_debit for the amount of the invoice balance
                self.cur.execute(
                        "insert into credit_debit (amount, debit_invoice_id, user_id) \
                            values(:current_balance, :current_invoice_id, :current_user_id)", 
                            {
                                "current_balance": self.current_invoice.balance, 
                                "current_invoice_id": self.current_invoice.invoice_id, 
                                "current_user_id":self.current_user.user_id
                                }
                                )
                self.con.commit()
            else:
                credit_used = self.current_user.available_credit
                # If the balance of the invoice is greater than the available credit
                # then make a credit_debit for the amount of the available credit
                self.cur.execute(
                        "insert into credit_debit (amount, debit_invoice_id, user_id) \
                            values(:available_credit, :current_invoice_id, :current_user_id)", 
                            {
                                "available_credit": self.current_user.available_credit, 
                                "current_invoice_id": self.current_invoice.invoice_id, 
                                "current_user_id":self.current_user.user_id
                                }
                                )
                self.con.commit()
            self.current_invoice.credit_used_against_this_invoice + credit_used
            self.current_invoice.balance -= credit_used
            return credit_used
        else:
            return 0.00

    def _check_for_and_delete_unused_or_old_charges(self):
        # NB for some reason we cannot get the 'IN' keyword to work with parameter substitution
        # 202208 (We did get the IN working, look else where in the code to see how (using join statement))
        # The statement otherwise works fine hard coded.
        # Because of this we will manually parse the results to select
        # those that are not in the self.charge_ids_related_to_PPMS_invoice
        self.cur.execute(
            "SELECT charge_id, project_title, staff_hours, staff_hourly_rate_eur, subsidy_percent FROM staff_time_charges INNER JOIN projects ON projects.project_id = staff_time_charges.project_id WHERE invoice_id=:invoice_id",
            {"invoice_id":self.current_invoice.invoice_id}
            )
        results = self.cur.fetchall()
        results = [_ for _ in results if _[0] not in self.charge_ids_related_to_PPMS_invoice]
        if results:
            print(f"Deleting the following old charges for {self.current_user.last_name}")
            print("charge_id\tproject_title\tstaff_hours\tstaff_hourly_rate_eur\tsubsidy_percent")
            for result in results:
                print("\t".join([str(_) for _ in result]))
        delete_ids = [_[0] for _ in results]
        for delete_id in delete_ids:
            self.cur.execute("DELETE FROM staff_time_charges WHERE charge_id = :delete_id", {"delete_id":delete_id})
            self.con.commit()

    def _populate_and_write_template(self):
        # Here populate the template
        # Populate the user and invoice data
        self.doc = DocxTemplate(self.template_path)
        self._populate_context()
        
        # Apply credits to the
        # invoice if there is the user balance to do so.
        self._apply_credits_for_user()
        
        self._write_template()

        # At this point we have all of the projects and charges created and associated to the users invoice.
        # We can now commit
        self.con.commit()

        # The final thing to do is to update the invoice amount_payable to self.current_invoice.balance
        # And if this is > 0 then to create an outstanding payment object.
        self.current_invoice.amount_payable = self.current_invoice.balance

    def _write_template(self):
        self.doc.render(self.context)
        outpath = os.path.join(self.output_dir, f"{self.first_month}_{self.last_month}_{self.current_user.last_name.replace(' ', '_')}_SequAna_Invoice.docx")
        if os.path.exists(outpath):
            if self._get_n_y_user_response(question_text=f"\n\n{outpath} already exists.\nOverwrite? [y/n]: ") == "y":
                self.doc.save(outpath)
                print(f"\nWriting {outpath}.")
            else:
                print("Commiting db objects and moving to next invoice.")
        else:
            self.doc.save(outpath)
            print(f"\nWriting {outpath}.")

    def _populate_context(self):
        self.context = {}
        self.context["invoice_id"] = self.current_invoice.invoice_id
        self.context["invoice_date"] = self.current_invoice.invoice_timestamp.split(" ")[0].replace("-", "")
        self.context["invoice_period"] = f"{self.first_month}-{self.last_month} inc."
        self.context["chargeable_account"] = self.chargeable_account
        self.context["user_name"] = f"{self.current_user.last_name}, {self.current_user.first_name}"
        self.context["user_email"] = self.current_user.email

        # Charge data for wetlab projects
        # Get wetlab projects for the user
        self._populate_context_with_projects_for_project_type("wetlab")
        self._populate_context_with_projects_for_project_type("bioinf")
        self._populate_context_with_projects_for_project_type("training")
        self._populate_context_with_consumable_costs()

        self.context["total_consumables_cost"] = f"{self.current_invoice.total_consumable_cost:.2f}"
        self.context["total_consumables_subsidy"] = f"{self.current_invoice.total_comsumables_subsidy_amount:.2f}"
        self.context["total_consumables_amount_payable"] = f"{self.current_invoice.total_consumables_amount_payable:.2f}"

        if self.current_invoice.charges_count == 0 and self.current_invoice.total_consumables_amount_payable:
            raise RuntimeError(f"No charges for user {self.current_user.last_name}")

        self.context["total_staff_hours"] = f"{self.current_invoice.total_staff_hours:.2f}"
        self.context["total_staff_cost"] = f"{self.current_invoice.total_staff_cost:.2f}"
        self.context["total_subsidy_amount"] = f"{self.current_invoice.total_staff_subsidy_amount:.2f}"
        self.context["amount_payable_staff"] = f"{self.current_invoice.total_staff_cost - self.current_invoice.total_staff_subsidy_amount:.2f}"
        self.context["balance"] = f"{self.current_invoice.balance:.2f}"

    def _populate_context_with_consumable_costs(self):
        self.cur.execute(
            "SELECT name, unit_cost, quantity, subsidy_percent \
            FROM consumable_charges \
            WHERE invoice_id=:invoice_id ORDER BY name ASC", 
            {"invoice_id": self.current_invoice.invoice_id}
            )
        results = self.cur.fetchall()
        if results:
            self.context["consumables"] = []
        else:
            # If there are no consumables then no need to proceed.
            return
        
        for con_charge in results:
            name, unit_cost, quantity, subsidy_percent = con_charge
            con_cost = unit_cost * quantity
            subsidy_dec = subsidy_percent/100
            subtotal = con_cost - (con_cost * subsidy_dec)
            
            self.current_invoice.total_consumable_cost += con_cost
            self.current_invoice.total_comsumables_subsidy_amount += con_cost * subsidy_dec
            self.current_invoice.total_consumables_amount_payable += subtotal

            self.current_invoice.balance += subtotal
            self.context[f"consumables"].append(
                {
                    "name": name,
                    "quantity": f"{quantity}",
                    "unit_cost": f"{unit_cost:.2f}",
                    "cost": f"{con_cost:.2f}",
                    "subsidy": f"{subsidy_percent:.2f}",
                    "subtotal": f"{subtotal:.2f}"
                }
            )
        

    def _get_or_make_charge(self, proj_of_user):
        # Now make a staff charge for the project
        staff_hours = self.hours_charged_df.loc[proj_of_user,:].sum()
        self.cur.execute(
                    "SELECT charge_id, staff_hours, staff_hourly_rate_eur, subsidy_percent FROM staff_time_charges WHERE invoice_id=:invoice_id AND project_id=:project_id",
                    {"invoice_id":self.current_invoice.invoice_id, "project_id":self.current_project.project_id}
                    )
        result = self.cur.fetchall()
        if result:
            assert(len(result) == 1)
            print(f"\nCharge for the {self.current_project.project_type} project {self.current_project.project_id}: {self.current_project.project_title} for this invoice already exists.")
            obj_charge_id, obj_staff_hours, obj_staff_hourly_rate_eur, obj_staff_subsidy_percent = result[0]
            # Check that each of the values in the db for the given charge object match and if they
            # don't alter the table entry and log out to user
            if obj_staff_hours == staff_hours:
                pass
            else:
                self.cur.execute("UPDATE staff_time_charges SET staff_hours = :staff_hours WHERE charge_id=:obj_charge_id", {"staff_hours":staff_hours, "obj_charge_id":obj_charge_id})
                print(f"Modifying staff_hours for staff_time_charges object already in database: {obj_staff_hours} --> {staff_hours}")
            if obj_staff_hourly_rate_eur == self.staff_hourly_rate_eur:
                pass
            else:
                self.cur.execute("UPDATE staff_time_charges SET staff_hourly_rate_eur = :staff_hourly_rate_eur WHERE charge_id=:obj_charge_id", {"staff_hourly_rate_eur":self.staff_hourly_rate_eur, "obj_charge_id":obj_charge_id})
                print(f"Modifying staff_hourly_rate_eur for staff_time_charges object already in database: {obj_staff_hourly_rate_eur} --> {self.staff_hourly_rate_eur}")
                
            if obj_staff_subsidy_percent == self.current_user.staff_subsidy_percent:
                pass
            else:
                self.cur.execute("UPDATE staff_time_charges SET subsidy_percent = :subsidy_percent WHERE charge_id=:obj_charge_id", {"subsidy_percent":self.current_user.staff_subsidy_percent, "obj_charge_id":obj_charge_id})
                print(f"Modifying subsidy_percent for staff_time_charges object already in database: {obj_staff_subsidy_percent} --> {self.current_user.staff_subsidy_percent}")
            self.charge_ids_related_to_PPMS_invoice.append(obj_charge_id)
            
        else:
            print(f"\nMaking a staff_time_charges for the {self.current_project.project_type} project {self.current_project.project_id}: {self.current_project.project_title}")
            self.cur.execute(
                        "INSERT INTO staff_time_charges (staff_hours, staff_hourly_rate_eur, subsidy_percent, invoice_id, project_id) \
                            VALUES (:staff_hours, :staff_hourly_rate_eur, :subsidy_percent, :invoice_id, :project_id)",
                            {
                                "staff_hours": staff_hours, "staff_hourly_rate_eur": self.staff_hourly_rate_eur,
                                "subsidy_percent": self.current_user.staff_subsidy_percent, "invoice_id":self.current_invoice.invoice_id,
                                "project_id":self.current_project.project_id
                                }
                            )
            self.charge_ids_related_to_PPMS_invoice.append(self.cur.lastrowid)

    def _get_or_make_project(self, proj_of_user):
        proj_title = ":".join(proj_of_user.split(":")[1:]).strip()
        proj_type  = proj_of_user.split(':')[0].split("_")[-1]
        assert(proj_type in ["bioinf", "wetlab", "training"])
        # Check to see if the project has a match in the database
        self.cur.execute("SELECT project_id, user_id FROM projects WHERE project_title=:proj_title AND project_type=:proj_type AND user_id=:user_id", {"proj_title": proj_title, "proj_type":proj_type, "user_id": self.current_user.user_id})
        results = self.cur.fetchall()
        if not results:
            # The project does not already exist and we need to create the project
            if self._get_n_y_user_response(question_text=f"\n\nProject with title: {proj_title} does not exist in the database. \n\nWould you like to create this project now?\nEntering n will skip this project.\n[y/n]:") == "y":
                        
                if self._get_n_y_user_response(question_text=f"\n\nProject details are:\n\ttitle: {proj_title}\n\tproject_type: {proj_type}\n\tuser: {self.current_user.first_name} {self.current_user.last_name} {self.current_user.email}\n\tIs this correct? Entering n will exit the program so that you can correct the project information in PPMS input.\n[y/n]: ") == "y":
                    # Create the project
                    self.cur.execute(
                                "INSERT INTO projects (project_title, project_type, user_id) VALUES (:title, :proj_type, :user_id)",
                                {"title": proj_title, "proj_type": proj_type, "user_id": self.current_user.user_id}
                                )
                    self.con.commit()
                else:
                    sys.exit("\nExiting at users request.")
        else:
            assert(len(results) == 1)
            # Then the project already exits
        self.cur.execute("SELECT project_id, project_title, project_type FROM projects WHERE project_title=:proj_title AND project_type=:proj_type", {"proj_title": proj_title, "proj_type":proj_type})
        results = self.cur.fetchall()
        assert(len(results) == 1)
        return Project(results[0][0], con=self.con)

    def _get_or_make_invoice(self):
        # Now check to see if there are already invoices linked to this user for the given period
        # We are working under the logic that for a given user, separate invoices cannot overlap
        # If we find overlapping invoices then we raise an error
        # If not, then we make an invoice object that we will attach the charges to
        # when we make charges for each one of the projects
        # and build these into an invoice for the user
        # It will be important that we commit all of the charges in parallel with the new invoice and not individually
        # If we are remaking an invoice, then update the date to the current date.
        self.cur.execute(
            "SELECT invoice_id, first_month, last_month, invoice_timestamp, chargeable_account, sent FROM invoices WHERE user_id=:user_id AND invoice_type='debit' AND first_month >=:user_defined_first_month AND first_month <= :user_defined_last_month",
            {"user_id":self.current_user.user_id, "user_defined_first_month":self.first_month, "user_defined_last_month":self.last_month}
            )

        result = self.cur.fetchall()
        if result:
            assert(len(result) == 1)
            invoice_id, first_month, last_month, invoice_timestamp, chargeable_account, sent = result[0]
            if first_month == self.first_month and last_month == self.last_month:
                if sent:
                    # The invoice that matches the first and last month for this user already exits and has
                    # already been charged to that user
                    # We will not allow the script to overwrite this invoice
                    # Then pull the results back out of the database
                    print(
                        "An invoice object matches the first month, last month and current user exactly and has already been sent\n \
                        This script will not allow you to modify that invoice. You will need to do so manually in the db.\n\
                        Returning the invoice object unchanged.\n")
                    self.cur.execute(
                        "SELECT invoice_id, first_month, last_month, invoice_timestamp, chargeable_account, sent FROM invoices WHERE invoice_id=:invoice_id",
                        {"invoice_id": invoice_id}
                    )
                    result = self.cur.fetchall()
                else:
                    print(
                        "An invoice object matches the first month, last month and current user exactly but has not been sent\n \
                        This script will update the timestamp and chargeable account of the invoice.\n")
                    # Update the invoice object so that it get's a current timestamp and the chargeable_account
                    # matches the user supplied chargeable account.
                    self.cur.execute("UPDATE invoices SET invoice_timestamp=:timestamp, chargeable_account=:chargeable_account WHERE invoice_id=:invoice_id",
                    {"timestamp": datetime.datetime.now(), "chargeable_account":self.chargeable_account, "invoice_id":invoice_id})
                    self.con.commit()
                    # Then pull the results back out of the database
                    self.cur.execute(
                        "SELECT invoice_id, first_month, last_month, invoice_timestamp, chargeable_account, sent FROM invoices WHERE invoice_id=:invoice_id",
                        {"invoice_id": invoice_id}
                    )

                    result = self.cur.fetchall()
            else:
                raise NotImplementedError(
                        f"One or more invoices exist for user {self.current_user_last_name} \
                            which have a charging period that over lap with the user specified charging period.\n{result}"
                            )
        else:
            # Invoice does not exist.
            # Make the invoice.
            self.cur.execute(
                    "INSERT INTO invoices(invoice_timestamp, first_month, last_month, chargeable_account, user_id, amount_payable) VALUES (:timestamp, :first_month, :last_month, :chargeable_account, :user_id, :amount_payable)", 
                    {"timestamp": datetime.datetime.now(), "first_month": self.first_month, "last_month":self.last_month, "chargeable_account":self.chargeable_account, "user_id":self.current_user.user_id, "amount_payable":99999.99}
                    )
            
            self.cur.execute("SELECT invoice_id, first_month, last_month, invoice_timestamp, chargeable_account, sent FROM invoices WHERE invoice_id=:invoice_id", {"invoice_id": self.cur.lastrowid})
            self.con.commit()
            result = self.cur.fetchall()
            assert(len(result) == 1)
        return Invoice(result[0][0], self.con)

    def _get_or_make_user_for_invoicing(self, user_last_name):
        self.cur.execute("SELECT last_name FROM users where last_name=:user_last_name", {"user_last_name": user_last_name})
        results = self.cur.fetchall()
        if not results:
                # Then the user is not already in the database and it needs to be added or skipped
            if self._get_n_y_user_response(question_text=f"\n\nUser with last name: {user_last_name} does not exist in the database. Would you like to create this user now?\nEntering n will exit the script.\n[y/n]:") == "y":
                    # Create the user
                first_name, user_last_name, email, staff_subsidy, consumable_subsidy = self._get_first_name_email_subsidy_of_user(user_last_name)
                self.cur.execute(
                        "INSERT INTO users (first_name, last_name, email, staff_subsidy_percent, consumable_subsidy_percent) VALUES (:first_name, :user_last_name, :email, :staff_subsidy, :consumable_subsidy)",
                        {"first_name":first_name, "user_last_name":user_last_name, "email":email, "staff_subsidy":staff_subsidy, "consumable_subsidy":consumable_subsidy}
                        )
                self.con.commit()
                print(f"User {first_name} {user_last_name} successfully added to database.")
            else:
                sys.exit(f"Exiting script at users request.")
        else:
            assert(len(results) == 1)
        
        # Now get the user_id for use in creating the invoice, projects and charges
        # If user chose to skip then the user will not exist and we will move on to the next user.
        self.cur.execute("SELECT user_id FROM users WHERE last_name=:user_last_name", {"user_last_name": user_last_name})
        return User(self.cur.fetchone()[0], con=self.con)

    def _populate_context_with_projects_for_project_type(self, project_type):
        self.cur.execute(
            "SELECT staff_hours, staff_hourly_rate_eur, subsidy_percent, project_type, project_title \
                FROM staff_time_charges INNER JOIN projects ON projects.project_id = staff_time_charges.project_id \
                    WHERE invoice_id=:invoice_id AND project_type=:project_type ORDER BY staff_hours DESC", 
                    {"invoice_id": self.current_invoice.invoice_id, "project_type":project_type}
                    )
        results = self.cur.fetchall()
        self.context[f"{project_type}_projects"] = []
        for user_charge in results:
            staff_hours, staff_hourly_rate_eur, subsidy_percent, project_type, project_title = user_charge
            staff_cost = staff_hours * staff_hourly_rate_eur
            subsidy_dec = subsidy_percent/100
            subtotal = staff_cost - (staff_cost * subsidy_dec)
            self.current_invoice.balance += subtotal
            self.current_invoice.total_staff_hours += staff_hours
            self.current_invoice.total_staff_cost += staff_cost
            self.current_invoice.total_staff_subsidy_amount += staff_cost * subsidy_dec
            self.context[f"{project_type}_projects"].append(
                {
                    "project_title": project_title,
                    "staff_hours": f"{staff_hours:.2f}",
                    "staff_hourly_rate_eur": f"{staff_hourly_rate_eur:.2f}",
                    "staff_cost": f"{staff_cost:.2f}",
                    "subsidy": f"{subsidy_percent:.2f}",
                    "subtotal": f"{subtotal:.2f}"
                }
            )
            self.current_invoice.charges_count += 1

    def _get_first_name_email_subsidy_of_user(self, last_name=None):
        if last_name is None:
            last_name = input("Please provide the last name of the user: ")
        first_name = input("Please provide the first name of the user: ")
        email = input("Please provide the email of the user: ")
        staff_subsidy_percent = float(input("Percentage subsidy on staff hour charges (please enter a value between 0 and 1; 0.5 = 50%): "))
        assert((staff_subsidy_percent >= 0) and (staff_subsidy_percent <= 1))
        consumable_subsidy_percent = float(input("Percentage subsidy on consumable charges (please enter a value between 0 and 1; 0.5 = 50%): "))
        assert((consumable_subsidy_percent >= 0) and (consumable_subsidy_percent <= 1))
        staff_subsidy_percent = staff_subsidy_percent*100
        consumable_subsidy_percent = consumable_subsidy_percent*100
        if self._get_n_y_user_response(question_text=f"\n\nFirst name: {first_name}\nLast name: {last_name}\nEmail: {email}\nStaff cost subsidy: {staff_subsidy_percent}\nConsumable cost subsidy: {consumable_subsidy_percent}\nIs this correct? [y/n]: ") == "y":
            return first_name, last_name, email, staff_subsidy_percent, consumable_subsidy_percent
        else:
            self._get_first_name_email_subsidy_of_user(last_name)

    def _do_argument_qc(self):
        """
        Do a range of QC on the arguments provided on the command line and explicitly return the
        self variables that will be used in the remainder of the program for readability purposes
        """
        self.skip_user_input = self.args.answer_yes

        self._do_first_last_month_qc()
        
        if self.args.PPMS_input_consumables_csv:
            self._do_ppms_input_consumables_csv_qc()

        self.month_range, self.hours_charged_df = self._do_ppms_input_staff_hours_csv_qc()
                
        self.user_last_names_to_invoice = self._do_user_input_qc()
        
        self._do_template_qc()

        self.chargeable_account = self.args.chargeable_account

        self.staff_hourly_rate_eur = self.args.staff_hourly_rate_eur

        self.output_dir = os.path.abspath(self.args.output_dir)
        if not os.path.exists(self.output_dir):
            os.makedirs(self.output_dir)

        self.db_backup_dir = self.args.db_backup_dir
        if not os.path.exists(self.db_backup_dir):
            os.makedirs(self.db_backup_dir, exist_ok=True)

        return self.month_range, self.hours_charged_df, self.user_last_names_to_invoice, self.chargeable_account, self.staff_hourly_rate_eur, self.output_dir


    def _do_template_qc(self):
        if self.args.template:
            self.template_path = os.path.abspath(self.args.template)
            if not os.path.exists(self.template_path):
                raise FileNotFoundError(f"{self.template_path} not found.")
        else:
            self.template_path = None
            # Then search for a file in the current directory that ends with sequana_invoice_template.docx
            for file in os.listdir('.'):
                if file.endswith("sequana_invoice_template.docx"):
                    self.template_path = os.path.abspath(os.path.join(".", "file"))
                    break

            if self.template_path is None:
                raise FileNotFoundError("A suitable template file could not be found. Please specify the path to the .docx template file using the --template argument")

    def _do_user_input_qc(self):
        if self.args.users:
            self.users = [_.replace("_", " ") for _ in self.args.users.split(",")]
        else:
            self.users = None

        # Generate a list of the users in the PPMS input
        # Incorporate checking to see if there are very similar last_names
        # We will need to do 2 run throughs the index. One to get the initial last name values
        # And the second where we check to see if there are any very similar matches
        
        # run one: collect the last name values
        self.user_last_names_to_invoice = set()
        for ind in self.hours_charged_df.index:
            last_name = "_".join(ind.split("_")[:-1]).replace("_", " ")
            if "sequana" not in last_name:
                self.user_last_names_to_invoice.add(last_name)
        # Include the names in the consumables input sheet to cover the
        # case where we have consumables but no staff costs.    
        for ind, ser in self.consumables_df.iterrows():
            last_name = "_".join(ser["Project name"].split(":")[0].split("_")[:-1]).replace("_", " ")
            if "sequana" not in last_name:
                self.user_last_names_to_invoice.add(last_name)
        
        # Check to see that any user specified users are found in the PPMS input
        if self.users:
            users_not_found = [_ for _ in self.users if _ not in self.user_last_names_to_invoice]
            if users_not_found:
                raise RuntimeError(f"The following users were not found in the PPMS input you have provided {users_not_found}")

        # Trim down the self.user_last_names_to_invoice to only those listed by the user
        if self.users:
            self.user_last_names_to_invoice = [_ for _ in self.user_last_names_to_invoice if _ in self.users]
            if not self.user_last_names_to_invoice:
                raise RuntimeError("user_last_names_to_invoice list is empty")

        # run two: check for similarity in values
        for last_name in self.user_last_names_to_invoice:
            close_matches = get_close_matches(last_name, [_ for _ in self.user_last_names_to_invoice if _ != last_name], n=99, cutoff=0.8)
            if close_matches:
                if self._get_n_y_user_response(question_text=f"One of the last names in the PPMS output is very similar to one or more other last names:\n{last_name} is similar to {close_matches}.\nDo you want to continue anyway? [y/n]. Enter n to exit the program and fix the PPMS input.") != 'y':
                    sys.exit("Exiting at users request")

        # Trim down the input to only those of the requested users
        if self.user_last_names_to_invoice:
            inds_to_keep = []
            for ind in self.hours_charged_df.index:
                last_name = "_".join(ind.split("_")[:-1]).replace("_", " ")
                if last_name in self.user_last_names_to_invoice:
                    inds_to_keep.append(ind)
        self.hours_charged_df = self.hours_charged_df.loc[inds_to_keep,:]

        return self.user_last_names_to_invoice

    def _get_n_y_user_response(self, question_text, allow_skip=True):
        try:
            if self.skip_user_input and allow_skip:
                print(f"Answering y to: {question_text}")
                return "y"
            else:
                return input(question_text)
        except AttributeError:
            return input(question_text)

    def _format_completed_date(self, bad_date):
        yyyy = int(bad_date.split(" ")[0].split("/")[2])*100
        mm = int(bad_date.split(" ")[0].split("/")[1])

        return yyyy + mm

    def _do_ppms_input_consumables_csv_qc(self):
        # QC of consumables inputs
        # There should be only one such file
        self.ppms_consumables_input_csv_path = self.args.PPMS_input_consumables_csv
        if not os.path.exists(self.ppms_consumables_input_csv_path):
            raise FileNotFoundError(f"{self.ppms_consumables_input_csv_path} not found.")
        
        self.consumables_df = df = pd.read_csv(self.ppms_consumables_input_csv_path)
        self.consumables_df.drop("Group", axis=1, inplace=True)
        self.consumables_df.drop("User", axis=1, inplace=True)
        
        # Convert the Completed date from DD/MM/YYY TT:TT to YYYYMM
        self.consumables_df["month"] = [self._format_completed_date(_) for _ in self.consumables_df["Completed date"]]
        self.consumables_df.drop("Completed date", axis=1, inplace=True)
        
        # Filter down to only those months that fall within the first and last month
        self.consumables_df = self.consumables_df.loc[(self.consumables_df["month"] >= int(self.first_month)) & (self.consumables_df["month"] <= int(self.last_month)),:]

    def _create_consumable_charge(self, ser):
        # We should be able to get the user, project and invoice  or make them  using the project_info only.
        last_name = "_".join(ser["Project name"].split(":")[0].split("_")[:-1]).replace("_", " ")
        
        # Then we need to create the project and the invoice
        self.current_user = self._get_or_make_user_for_invoicing(user_last_name=last_name)
        self.current_project = self._get_or_make_project(ser["Project name"])
        self.current_invoice = self._get_or_make_invoice()
        
        # Then we need to create a consumable charge
        self.cur.execute(
            "INSERT INTO consumable_charges (name, unit_cost, quantity, date, invoice_id, project_id, PPMS_reference) VALUES (:name, :unit_cost, :quantity, :date, :invoice_id, :project_id, :PPMS_reference)",
            {"name": ser["Consumable name"], "unit_cost": ser["Unit price"], "quantity": ser["Quantity"], "date": ser["month"], "invoice_id": self.current_invoice.invoice_id, "project_id": self.current_project.project_id, "PPMS_reference": ser["Ref."]}
            )
        self.con.commit()
        
    def _consumable_charge_exits(self, project_id, invoice_id, name, date, unit_cost, quantity, ref):
        self.cur.execute(
            "SELECT charge_id from consumable_charges where project_id=:project_id AND invoice_id=:invoice_id AND name=:name AND date=:date AND unit_cost=:unit_cost AND quantity=:quantity AND PPMS_reference=:ref",
            {"project_id": project_id, "invoice_id": invoice_id, "name": name, "date": date, "unit_cost":unit_cost, "quantity": quantity, "ref":ref}
            )
        if len(self.cur.fetchall()) == 1:
            return True
        return False

    def _get_invoice_id(self, user_id):
        self.cur.execute(
            "SELECT invoice_id from invoices where user_id=:user_id AND first_month<=:first_month AND last_month>=:last_month", 
            {"first_month":int(self.first_month), "last_month":int(self.last_month), "user_id": user_id}
            )
        return self.cur.fetchall()[0][0]
    
    def _invoice_exists(self, user_id):
        self.cur.execute(
            "SELECT invoice_id from invoices where user_id=:user_id AND first_month<=:first_month AND last_month>=:last_month", 
            {"first_month":int(self.first_month), "last_month":int(self.last_month), "user_id": user_id}
            )
        if len(self.cur.fetchall()) == 1:
            return True
        else:
            return False

    def _get_project_id_user_id_from_project(self, project_name):
        project_title = ":".join(project_name.split(":")[1:]).strip()
        project_type = project_name.split(":")[0].split("_")[-1]
        self.cur.execute("SELECT project_id, user_id FROM projects WHERE project_title=:project_title AND projects.project_type=:project_type", {"project_title": project_title, "project_type": project_type})
        results = self.cur.fetchall()
        return results[0]

    def _project_exists(self, project_name):
        project_title = ":".join(project_name.split(":")[1:]).strip()
        project_type = project_name.split(":")[0].split("_")[-1]
        last_name = "_".join(project_name.split(":")[0].split("_")[:-1]).replace("_", " ")
        self.cur.execute("SELECT projects.project_id, projects.user_id FROM projects INNER JOIN users ON projects.user_id = users.user_id WHERE projects.project_title=:project_title AND projects.project_type=:project_type AND users.last_name=:last_name", {"project_title": project_title, "project_type": project_type, "last_name": last_name })
        results = self.cur.fetchall()
        if len(results) == 1:
            return True
        if len(results) > 1:
            raise RuntimeError("Multiple results returned when 1 was expected")
        else:
            return False

    def _do_ppms_input_staff_hours_csv_qc(self):
        # QC of PPMS_input_csvs
        self.ppms_input_csvs_paths = self.args.PPMS_input_staff_hours_csvs.split(",")
        for ppms_path in self.ppms_input_csvs_paths:
            if not os.path.exists(ppms_path):
                raise FileNotFoundError(f"{ppms_path} not found.")
        
        # If all found then put them into one big df where the columns are in the format YYYYMM
        # There are likely to be two files for any given period
        # one for ben and one for alyssa which separately log their respective hours
        # A dict with year as key and df as value
        year_to_df_dict = {}
        for ppms_path in self.ppms_input_csvs_paths:
            df = pd.read_csv(ppms_path)
            df.set_index("Project", inplace=True)
            year = list(df)[13].split(" ")[0]
            df.drop(["Type", f"{year} Total"], inplace=True, axis=1)
            # Check that we are left with only the 12 months
            if len(list(df)) != 12:
                raise RuntimeError("DataFrame formatting error")
            # Change the columns to format YYYMM
            new_cols = range(int(year)*100 + 1, (int(year)*100) + 13, 1)
            df.columns = new_cols
            df = df.fillna(0)

            # Remove SequAana projects from the PPMS
            df = df.loc[[_ for _ in df.index if "sequana" not in _.lower()], :]

            if year in year_to_df_dict:
                print(f"\n\nPPMS input {ppms_path} contains inputs for the same period as another PPMS input.")
                print("The two inputs will be merged")
                # Then a df for this year already exists and we need to merge the two dfs
                # making sure to summate any two projects with exactly the same name
                existing_df = year_to_df_dict[year]
                existing_indices = existing_df.index
                for ind, series in df.iterrows():
                    if ind in existing_indices:
                        # The exact same project exists in both dataframes
                        existing_df.loc[ind, :] = existing_df.loc[ind, :].add(series)
                    else:
                        existing_df.loc[ind] = series
                year_to_df_dict[year] = existing_df
            else:
                year_to_df_dict[year] = df
                
        
        self.hours_charged_df = pd.concat(year_to_df_dict.values(), axis=1).fillna(0)
        # Order the columns
        self.hours_charged_df = self.hours_charged_df.loc[:,sorted(list(self.hours_charged_df))]

        # Check that all of the months inbetween the first_month and last_month have data
        self.month_range = []
        for year in range(int(self.first_month[:4]), int(self.last_month[:4]) + 1, 1):
            if str(year) in self.first_month and str(year) in self.last_month:
                # Then we only want to include the months that are greater than or
                # equal to the first month and less than or equal to the last month
                months = range(int(self.first_month[4:]), int(self.last_month[4:]) + 1, 1)
            elif str(year) in self.first_month and str(year) not in self.last_month:
                # Then we want to start at the month of first_month and go to end of year
                months = range(int(self.first_month[4:]), 13, 1)
            elif str(year) in self.last_month and str(year) not in self.first_month:
                # Then we want the months from the beginning up until the month of last_month
                months = range(1, int(self.last_month[4:]) + 1, 1)
            elif str(year) not in self.first_month and str(year) not in self.last_month:
                months = range(1, 13, 1)
            
            self.month_range.extend([(year*100) + _ for _ in months])
        
        # sort the month_range
        self.month_range = sorted(self.month_range)

        missing_months = [_ for _ in self.month_range if _ not in list(self.hours_charged_df)]
        if missing_months:
            raise RuntimeError(f"The following months are missing in the PPMS input datasheets {missing_months}")

        # Check the project names to make sure that they are either exact or not similar
        projects = set()
        for proj_index in self.hours_charged_df.index:
            proj_name = proj_index.split(":")[-1].rstrip().strip()
            projects.add(proj_name)

        for proj_name in projects:
            close_matches = get_close_matches(proj_name, [_ for _ in projects if _ != proj_name], n=99, cutoff=0.8)
            if close_matches:
                if self._get_n_y_user_response(allow_skip=False, question_text=f"One of the project names in the PPMS output is very similar to one or more other project names:\n{proj_name} is similar to {close_matches}.\nDo you want to continue anyway?\nEnter n to exit the program and fix the PPMS input.\n [y/n]: ") != 'y':
                    sys.exit("Exiting at users request")

        # Trim down to the user specified months
        self.hours_charged_df = self.hours_charged_df.loc[:,self.month_range]

        # Remove any rows that sum to 0 i.e. have no billable hours
        self.hours_charged_df = self.hours_charged_df.loc[self.hours_charged_df.sum(axis=1) != 0,:]

        return self.month_range, self.hours_charged_df

    def _do_first_last_month_qc(self):
        # first and last month QC
        self.first_month = str(self.args.first_month)
        if len(self.first_month) != 6:
            raise FormatError("first_month should be in format YYYYMM")
        
        self.last_month = str(self.args.last_month)
        if len(self.last_month) != 6:
            raise FormatError("first_month should be in format YYYYMM")

        first_month_year = int(self.first_month[:4])
        if first_month_year < 2000:
            raise ValueError("first month year must be 2000 or later")
        
        last_month_year = int(self.last_month[:4])
        if last_month_year < first_month_year:
            raise ValueError("last month year must be equal or reater than the first month year")

        if int(self.last_month) < int(self.first_month):
            raise ValueError("last month must be equal or later than first month")

    def _parse_args(self):
        parser = argparse.ArgumentParser(prog='sequana')
        parser.set_defaults(func=parser.print_help)
        subparsers = parser.add_subparsers()
        
        # The subprograms of the invoicing program
        # There are currently 5 subprograms
        
        # create_invoices: this is used to make the invoices for users I.e. that 
        # have details of staff time and cosumable charges. These are then sent by
        # the users to accounting
        create_invoices_parser = subparsers.add_parser(
            'create_invoices',
            help='Create invoices to charge for staff time and consumables from PPMS output reports using a template document. \
            It can be used to make an invoice for one or more users across a single time period.'
            )
        create_invoices_parser.add_argument(
            '--first_month', action='store', required=True,
            help="The first month of the charging period to be charged. Fomat is YYYYMM."
            )
        create_invoices_parser.add_argument(
            '--last_month', action='store', required=True,
            help="The last moth of the charging period. Fomat is YYYYMM."
            )
        create_invoices_parser.add_argument(
            '--PPMS_input_staff_hours_csvs', action='store', required=True,
            help='The relative or full path to the report exported from PPMS that gives a per month, \
                per project breakdown of charged hours. Multiple files may be provided as a comma separated list. \
                E.g. ../2022_PPMS_billed_hours.csv,2021_PPMS_billed_hours.csv. Currently we are outputting separate csv \
                    files for Hume and Bell.'
        )
        create_invoices_parser.add_argument(
            '--PPMS_input_consumables_csv', action='store', required=False,
            help='The relative or full path to the report exported from PPMS that details the \
                consumables.'
        )
        create_invoices_parser.add_argument(
            '--users', action='store', required=False,
            help='Optional. If no user argument is provided invoices will be created for all users in the provided \
                PPMS_input_csvs for the charging period defined. \
                Users must be specified by last name using an underscore in place of any spaces. Multiple users should be provided as a comma separated string.'
        )
        create_invoices_parser.add_argument(
            '--template', action='store', required=False,
            help='Path to the invoice template. If not specified a *sequana_invoice_template.docx will be searched for in current directory and used.\
                If not found an error will be raised.'
        )
        create_invoices_parser.add_argument(
            '--output_dir', action='store', required=False, default='.',
            help='The directory in which the invoices will be written. Default is current directory.'
        )
        create_invoices_parser.add_argument(
            '--chargeable_account', action="store", required=False, default="1414 11171 08 2151040901",
            help="The account that charged users should transfer the money to. Default: 1414 11171 08 2151040901"
            )
        create_invoices_parser.add_argument(
            '--staff_hourly_rate_eur', action="store", required=False, default=20,
            help="The hourly rate in EUR for staff costs. Default: 20"
            )
        create_invoices_parser.add_argument(
            '--answer_yes', action="store_true", required=False,
            help="When passed, all interactive prompts will be skipped as though the answer 'y' was given."
            )
        create_invoices_parser.add_argument(
            '--db_backup_dir', action="store", required=False, default="db_backup",
            help="The directory in which backups of the database are made. Defaults to './db_backup'. A backup is automatically made after every successful run of create_invoices."
            )
        create_invoices_parser.set_defaults(func=self._init_create_invoices)

        # Create credit invoices
        # This is used to create a credit invoice which is different to a standard invoice in that it
        # is used to allow users to buy prepaid credit with sequana.
        create_credit_invoices_parser = subparsers.add_parser(
            'create_credit_invoices',
            help='Create invoices for users to buy prepaid credit for SequAna using a template document. \
            It can be used to make a credit invoice with details input using a csv file.'
            )
        create_credit_invoices_parser.add_argument('--input', help="The .csv file containing the credit invoice details.", required=True)
        create_credit_invoices_parser.add_argument(
            '--chargeable_account', action="store", required=False, default="1414 11171 08 2151040901",
            help="The account that charged users should transfer the money to. Default: 1414 11171 08 2151040901"
            )
        create_credit_invoices_parser.add_argument(
            '--template', action='store', required=True,
            help='Path to the credit invoice template.'
        )
        create_credit_invoices_parser.add_argument(
            '--output_dir', action='store', required=False, default='.',
            help='The directory in which the credit invoices will be written. Default is current directory.'
        )
        create_credit_invoices_parser.set_defaults(func=self._init_create_credit_invoices)

        # Set invoice as sent
        # This command is used to set an invoice (either a standard invoice or a credit invoice)
        # to the status sent.
        set_invoices_sent = subparsers.add_parser(
            'set_invoices_sent',
            help='Set the sent status of invoices in the database to True.'
            )
        set_invoices_sent.add_argument('--input', help="The .csv file containing the credit invoice details.", required=True)
        set_invoices_sent.set_defaults(func=self._set_invoice_to_sent)

        # Set invoice as paid
        # This command is used to set an invoice (either a standard invoice or a credit invoice)
        # to the status paid.
        set_invoices_paid = subparsers.add_parser(
            'set_invoices_paid',
            help='Set the paid status of invoices in the database to True.'
            )
        set_invoices_paid.add_argument('--input', help="The .csv file containing the credit invoice details.", required=True)
        set_invoices_paid.set_defaults(func=self._set_invoices_paid)

        # Create new user
        # This can be used to input a new user into the database
        make_new_user = subparsers.add_parser(
            'make_new_user',
            help='Make a new user'
            )
        make_new_user.set_defaults(func=self._make_new_user)

        self.args = parser.parse_args()
        self.args.func()

Invoicing()