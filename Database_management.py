import json
import ast


def update_record(invoice_record):
    with open('database', 'a') as fp:
        fp.write("\n")
        json.dump(invoice_record, fp)


def read_record():
    itemlist = []
    with open('database', 'r') as fp:
        data = (fp.read()).split("\n")
    for i in data:
        itemlist.append(ast.literal_eval(i))
    return itemlist

def get_invoice_number():
    invoice_record = read_record()
    last_invoice_number_entry = invoice_record[len(invoice_record)-1]
    last_invoice_number = last_invoice_number_entry["Invoice_number"]
    return last_invoice_number

def check_record(invoice_number):
    invoice_record = read_record()
    for i in invoice_record:
        if i["Invoice_number"] == invoice_number:
            return "exist"
        else:
            return "do no exist"


def get_record_by_name(name):
    invoice_record = read_record()
    sample_list = []
    for i in invoice_record:
        if i["Customer_details"]["Name"] == name:
            sample_list.append(i)

    for i in sample_list:
        print(i)
