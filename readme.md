# SequAna invoicing

This script is used for making invoices for SequAna from
the outputs of the PPMS time management system.

It consists of a python script called `invoicing.py` and a sqlite3 database called `invoicing.db`. There are also several python modules that the `invoicing.py` script makes use including:

- `projects.py`
- `user.py`
- `invoices.py`

# Running `invoicing.py`

## The sequana_admin conda environment
You can install the necessary packages to run invoicing.py using the sequana_admin conda environment for which there is a yml file in the main directory.

To make a new environment from this file and then activate it:

```
conda env create -f sequana_admin.yml
conda activate sequana_admin
```
## The command line interface
The invoicing.py commandline interface runs using subcommands. To see available commands execute the script either using python:

```
python3 invoicing.py
```

Or executing the script (it must have executable permissions for your user)

```
# from the same directory
./invoicing.py
```

The result will show the available commands with help text:

```
usage: sequana [-h]
               {create_invoices,create_credit_invoices,set_invoices_sent,set_invoices_paid,make_new_user}
               ...

positional arguments:
  {create_invoices,create_credit_invoices,set_invoices_sent,set_invoices_paid,make_new_user}
    create_invoices     Create invoices to charge for staff time and
                        consumables from PPMS output reports using a template
                        document. It can be used to make an invoice for one or
                        more users across a single time period.
    create_credit_invoices
                        Create invoices for users to buy prepaid credit for
                        SequAna using a template document. It can be used to
                        make a credit invoice with details input using a csv
                        file.
    set_invoices_sent   Set the sent status of invoices in the database to
                        True.
    set_invoices_paid   Set the paid status of invoices in the database to
                        True.
    make_new_user       Make a new user

optional arguments:
  -h, --help            show this help message and exit
```

For any given command you can see the available arguments using the help argument `-h` e.g.:

```
python3 invoicing.py create_invoices -h
usage: sequana create_invoices [-h] --first_month FIRST_MONTH --last_month
                               LAST_MONTH --PPMS_input_staff_hours_csvs
                               PPMS_INPUT_STAFF_HOURS_CSVS
                               [--PPMS_input_consumables_csv PPMS_INPUT_CONSUMABLES_CSV]
                               [--users USERS] [--template TEMPLATE]
                               [--output_dir OUTPUT_DIR]
                               [--chargeable_account CHARGEABLE_ACCOUNT]
                               [--staff_hourly_rate_eur STAFF_HOURLY_RATE_EUR]
                               [--answer_yes] [--db_backup_dir DB_BACKUP_DIR]

optional arguments:
  -h, --help            show this help message and exit
  --first_month FIRST_MONTH
                        The first month of the charging period to be charged.
                        Fomat is YYYYMM.
  --last_month LAST_MONTH
                        The last moth of the charging period. Fomat is YYYYMM.
  --PPMS_input_staff_hours_csvs PPMS_INPUT_STAFF_HOURS_CSVS
                        The relative or full path to the report exported from
                        PPMS that gives a per month, per project breakdown of
                        charged hours. Multiple files may be provided as a
                        comma separated list. E.g. ../2022_PPMS_billed_hours.c
                        sv,2021_PPMS_billed_hours.csv. Currently we are
                        outputting separate csv files for Hume and Bell.
  --PPMS_input_consumables_csv PPMS_INPUT_CONSUMABLES_CSV
                        The relative or full path to the report exported from
                        PPMS that details the consumables.
  --users USERS         Optional. If no user argument is provided invoices
                        will be created for all users in the provided
                        PPMS_input_csvs for the charging period defined. Users
                        must be specified by last name using an underscore in
                        place of any spaces. Multiple users should be provided
                        as a comma separated string.
  --template TEMPLATE   Path to the invoice template. If not specified a
                        *sequana_invoice_template.docx will be searched for in
                        current directory and used. If not found an error will
                        be raised.
  --output_dir OUTPUT_DIR
```

## Standard practice
The standard workflow is to create a set of invoices. This will create a set of word documents that should be examined for correctness before converting to .pdf. The invoice should then be sent to the user. The invoice sent status should then be set to True (see below). Once the payment has been made, the invoice paid status should then be set to True (see below)

## Creating standard invoices with invoicing.py

I use the term standard invoice to differentiate from a credit invoice (see below)

The command used for creating a standard SequAna invoice is `create_invoices`.

There are several required arguments to this command that are either required or that you will normally want to use:

- `--first_month`: SequAna invoices are generally generated on a monthly basis. So commonly the first month value will be the same as the last. This will result in an invoice for a single month. Otherwise the first month should be the first month that the invoice should cover. It is not possible to generate invoices at a resolution less than a single month.

- `--last_month`: The last month that should be included in the invoice. This will be the same as first_month if you are producing an invoice for a single month.

- `--PPMS_input_staff_hours_csvs`: This is comma delimited list of csv files output from the PPMS time management system. To output the files, go to 'Reports>Custom Report>Time logged per project per month/year per system' then select the member of staff you wish to output for and click 'run/refresh report'. Then save as a csv. If you wish to invoice for more than one member of staff (this is normal) then you can supply multiple comma separated file paths here.

- `--PPMS_input_consumables_csv`: To include invoicing for consumables (this is normal) you will also need to output a .csv report from PPMS ('Reports>Custom Report>Order Summary Report'). The path to this csv should be provided as the argument to this flag.

- `--template`: Full path to the template for the invoice. At the time of writing this documentation this was `/home/humebc/sequana_admin/invoices/invoice_templates/20221123_sequana_invoice_template.docx`. This template contains [jinja](https://jinja.palletsprojects.com/en/3.1.x/) fields that will be populated by the python code. If you modify the template you will need to modify the python code equivalently.

- `--output_dir`: Directory path where the invoices should be output to.

Example:
```
$ python3 invoicing.py create_invoices --first_month 202210 --last_month 202210 --PPMS_input_staff_hours_csvs /home/humebc/sequana_admin/invoices/invoices/202210/input_csvs/202210_hume.csv,/home/humebc/sequana_admin/invoices/invoices/202210/input_csvs/202210_bell.csv --PPMS_input_consumables_csv /home/humebc/sequana_admin/invoices/invoices/202210/input_csvs/202210_orders.csv --template /home/humebc/sequana_admin/invoices/invoice_templates/20221123_sequana_invoice_template.docx --output_dir ./invoices/202210/invoices --answer_yes
```

## Creating credit invoices with invoicing.py 

Credit invoices are invoices that are created to ask users to make a payment that will then be added as credit to their account.

They are created with the `create_credit_invoices` subcommand.

```
$ python3 invoicing.py create_credit_invoices
usage: sequana create_credit_invoices [-h] --input INPUT
                                      [--chargeable_account CHARGEABLE_ACCOUNT]
                                      --template TEMPLATE
                                      [--output_dir OUTPUT_DIR]
sequana create_credit_invoices: error: the following arguments are required: --input, --template
```

The two required inputs are:

- `--input`: Full path to the input csv. This is a csv that is currently called `invoice_input.csv`. It should have 2 columns with headers. The headers are `user_email` and `amount_payable`. The values for these columns must be an email of the user to whom you want to create the credit for and the amount in EURO that should be credited.

- `--template`: Full path to the credit invoice template. This is currently: `/home/humebc/sequana_admin/invoices/invoice_templates/20220811_SequAna_credit_invoice_template.docx`. This template contains [jinja](https://jinja.palletsprojects.com/en/3.1.x/) fields that will be populated by the python code. If you modify the template you will need to modify the python code equivalently.

## Setting invoices to sent
Once an invoice has been created it will be populated in the database with attribute 'sent' and 'paid' which will both be set to false. Once you have sent the invoice to the user, you should set the sent status to true in the database. You do this using the `set_invoices_sent` subcommand:

```
$ python3 invoicing.py set_invoices_sent -h
usage: sequana set_invoices_sent [-h] --input INPUT

optional arguments:
  -h, --help     show this help message and exit
  --input INPUT  The .csv file containing the credit invoice details.
```

This subcommand uses the same input file as the `create_credit_invoices` command. The program will look up the invoices according to user and payable amount and set the sent status to True.

## Setting invoices to paid
Once an invoice has been created it will be populated in the database with attribute 'sent' and 'paid' which will both be set to false. Once you have sent the invoice to the user and user has paid the invoice, you should set the paid status to true in the database. You do this using the `set_invoices_paid` subcommand:

```
$ python3 invoicing.py set_invoices_paid -h
usage: sequana set_invoices_paid [-h] --input INPUT

optional arguments:
  -h, --help     show this help message and exit
  --input INPUT  The .csv file containing the credit invoice details.
```

This subcommand uses the same input file as the `create_credit_invoices` command. The program will look up the invoices according to user and payable amount and set the paid status to True.

Upon setting the paid status of the invoice to True the credit will become available to the user and will be used to pay standard invoices that are created after the credit has become available (not before).

## Making new users
If a user does not exist in the database for a standard invoice that you are trying to create you will be prompted to enter in the information to create a new user on the command line.

Alternatively, you can run the sub command `make_new_user`.

There are no required inputs.


## Database backup
By default backups of the database are made every time standard invoices are made.

There are two types of back up made. The first is a copy of the sqlite database
in sqlite format.

The second is in .xlsx format where every table has been converted to an excel worksheet
and saved as an excel file. You will find these in `db_backup`.

The sqlite3 database is a simple file that is human readable. As such, if you mess something up, you can simply replace the current database file (`invoicing.db`) with a backup (making sure to change the name of the back up to be `invoicing.db`).

## Interacting with the database

The database can be accessed, queried and modified on the command line using the sqlite3 program by running:

```
sqlite3 invoicing.db
```

[This](https://www.sqlitetutorial.net/) is an excellent resource on how to work on the sqlite3 CLI.

## Database structure
To get an idea of the database structure either look at the database itself or look through `db_structure.txt`.