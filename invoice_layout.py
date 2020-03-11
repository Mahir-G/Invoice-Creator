from tkinter import *
from tkinter import ttk
import datetime
import tkinter.messagebox
import pdf_creator as pc
import pdf_creator_2 as pc2
import Database_management as dm

class invoice:
    def __init__(self, root):
        self.root = root
        self.root.title("Invoice Creator")
        self.root.geometry("1350x750+0+0")
        self.root.configure(background='powder blue')

        # ========================================================Variables======================================================

        Name = StringVar()
        Address = StringVar()
        GSTN = StringVar()
        State = StringVar()
        Code = StringVar()
        Phone = StringVar()

        SNo = 1
        ItemName = StringVar()
        HSNCode = StringVar()
        Qty = StringVar()
        Rate = StringVar()

        InvoiceNumber = StringVar()

        global invoice_number
        invoice_number = "Invoice #"+str(dm.get_invoice_number() + 1)

        # ===================================================DataBase============================================================

        Invoice = {'Invoice_number': "", 'Customer_details': "", 'Items': [], 'Amount': 0}
        Customer = {}


        # ===================================================Frame===============================================================

        MainFrame = Frame(self.root)
        MainFrame.grid()

        TitleFrame = Frame(MainFrame, bd=20, width=1365, padx=20, relief=RIDGE)
        TitleFrame.pack(side=TOP)

        self.lblTitle = Label(TitleFrame, font=('arial', 40, 'bold'), text="Invoice Creator", padx=2)
        self.lblTitle.grid()

        self.sublblTitle = Label(TitleFrame, font=('arial', 10, 'bold'), text=invoice_number, padx=2)
        self.sublblTitle.grid()

        FrameDetail = Frame(MainFrame, bd=20, width=1365, height=300, padx=20, relief=RIDGE)
        FrameDetail.pack(side=BOTTOM)

        ButtonFrame = Frame(MainFrame, bd=20, width=1365, height=50, padx=20, relief=RIDGE)
        ButtonFrame.pack(side=BOTTOM)

        DataFrame = Frame(MainFrame, bd=20, width=1365, height=245, padx=20, relief=RIDGE)
        DataFrame.pack(side=BOTTOM)

        DataFrameLeft = Frame(DataFrame, bd=10, width=450, height=245, padx=20, relief=RIDGE)
        DataFrameLeft.pack(side=LEFT)

        DataFrameLeftTop = LabelFrame(DataFrameLeft, bd=10, width=450, height=145, padx=20, relief=RIDGE,
                                      font=('arial', 12, 'bold'), text="Customer Details:")
        DataFrameLeftTop.pack(side=TOP)

        DataFrameRight = Frame(DataFrame, bd=10, width=450, height=245, padx=20, relief=RIDGE)
        DataFrameRight.pack(side=RIGHT)

        DataFrameRightTop = LabelFrame(DataFrameRight, bd=10, width=450, height=245, padx=20, relief=RIDGE,
                                       font=('arial', 12, 'bold'), text="Item Details:")
        DataFrameRightTop.pack(side=TOP)

        DataFrameRightBottom = LabelFrame(DataFrameRight, bd=10, width=450, height=245, padx=20, relief=RIDGE,
                                          font=('arial', 12, 'bold'), text="Actions:")
        DataFrameRightBottom.pack(side=BOTTOM)

        # ===================================================Function Declaration================================================

        def iExit():
            iExit = tkinter.messagebox.askyesno("Invoice Creator", "Confirm if you want to exit")
            if iExit > 0:
                root.destroy()
                return

        def iPrint():
            iSave()
            if len(Invoice['Items'])<=10:
                pc.layout(Invoice)
            elif len(Invoice['Items'])>10:
                pc2.layout(Invoice)

        def iSave():
            global invoice_number
            Number = dm.get_invoice_number() + 1
            Invoice['Invoice_number'] = Number
            invoice_number = "Invoice #" + str(Number+1)
            self.sublblTitle.configure(text=invoice_number)

            Date = datetime.date.today().strftime("%d-%B-%Y")
            Invoice['Date'] = Date

            Invoice['CGST'] = '9'
            Invoice['SGST'] = '9'
            Invoice['IGST'] = '0'

            Customer['Name'] = Name.get()
            Customer['Address'] = Address.get()
            Customer['GSTN'] = GSTN.get()
            Customer['State'] = State.get()
            Customer['Code'] = Code.get()
            Customer['Phone'] = Phone.get()

            Invoice['Customer_details'] = Customer

            dm.update_record(Invoice)

            Name.set("")
            Address.set("")
            GSTN.set("")
            State.set("")
            Code.set("")
            HSNCode.set("")
            Phone.set("")
            self.cboHSNCode.current(0)
            ItemName.set("")
            Rate.set("")
            Qty.set("")
            self.txtFrameDetail.delete('1.0', END)

            print(Invoice)

        def iAddItem():
            self.lblLabel = Label(FrameDetail, font=('arial', 10, 'bold'),
                                  text="\tSNo\tHSNCode\tItemName\tRate\tQty\tAmount\t")
            self.lblLabel.grid(row=0, column=0)

            Amount = str(int(Rate.get()) * int(Qty.get()))
            self.txtFrameDetail.insert(END, "\t\t\t\t\t\t" + str(
                SNo) + "\t" + HSNCode.get() + "\t\t" + ItemName.get() + "\t" + Rate.get() + "\t" + Qty.get() + "\t" + str(
                Amount) + "\n")

            item_description = {}
            item_description['ItemName'] = ItemName.get()
            item_description['HSNCode'] = HSNCode.get()
            item_description['Qty'] = Qty.get()
            item_description['Rate'] = Rate.get()

            Invoice['Items'].append(item_description.copy())

            Invoice['Amount'] += int(Amount)

            item_description.clear()
            HSNCode.set("")
            self.cboHSNCode.current(0)
            ItemName.set("")
            Rate.set("")
            Qty.set("")

            return

        def iDeleteItem():
            return

        # =============================================DataFrameLeftTop==========================================================

        self.lblName = Label(DataFrameLeftTop, font=('arial', 12, 'bold'), text="Name:", padx=2, pady=2)
        self.lblName.grid(row=0, column=0, sticky=W)

        self.txtName = Entry(DataFrameLeftTop, textvariable=Name, font=('arial', 12, 'bold'))
        self.txtName.grid(row=0, column=1)

        self.lblAddress = Label(DataFrameLeftTop, font=('arial', 12, 'bold'), text="Address:", padx=2, pady=2)
        self.lblAddress.grid(row=1, column=0, sticky=W)

        self.txtAddress = Entry(DataFrameLeftTop, textvariable=Address, font=('arial', 12, 'bold'))
        self.txtAddress.grid(row=1, column=1)

        self.lblGSTN = Label(DataFrameLeftTop, font=('arial', 12, 'bold'), text="GSTN:", padx=2, pady=2)
        self.lblGSTN.grid(row=2, column=0, sticky=W)

        self.txtGSTN = Entry(DataFrameLeftTop, textvariable=GSTN, font=('arial', 12, 'bold'))
        self.txtGSTN.grid(row=2, column=1)

        self.lblState = Label(DataFrameLeftTop, font=('arial', 12, 'bold'), text="State:", padx=2, pady=2)
        self.lblState.grid(row=3, column=0, sticky=W)

        self.txtState = Entry(DataFrameLeftTop, textvariable=State, font=('arial', 12, 'bold'))
        self.txtState.grid(row=3, column=1)

        self.lblCode = Label(DataFrameLeftTop, font=('arial', 12, 'bold'), text="Code:", padx=2, pady=2)
        self.lblCode.grid(row=4, column=0, sticky=W)

        self.txtCode = Entry(DataFrameLeftTop, textvariable=Code, font=('arial', 12, 'bold'))
        self.txtCode.grid(row=4, column=1)

        self.lblPhone = Label(DataFrameLeftTop, font=('arial', 12, 'bold'), text="Phone:", padx=2, pady=2)
        self.lblPhone.grid(row=5, column=0, sticky=W)

        self.txtPhone = Entry(DataFrameLeftTop, textvariable=Phone, font=('arial', 12, 'bold'))
        self.txtPhone.grid(row=5, column=1)

        # ===============================================DataFrameRightTop=======================================================

        self.lblHSNCode = Label(DataFrameRightTop, font=('arial', 12, 'bold'), text="HSN Code:", padx=2, pady=2)
        self.lblHSNCode.grid(row=0, column=0, sticky=W)

        self.cboHSNCode = ttk.Combobox(DataFrameRightTop, textvariable=HSNCode, state='readonly',
                                       font=('arial', 12, 'bold'), width=18)
        self.cboHSNCode['value'] = ('', '70072190', '70071100')
        self.cboHSNCode.current(0)
        self.cboHSNCode.grid(row=0, column=1)

        self.lblItemName = Label(DataFrameRightTop, font=('arial', 12, 'bold'), text="Item Name:", padx=2, pady=2)
        self.lblItemName.grid(row=1, column=0, sticky=W)

        self.txtItemName = Entry(DataFrameRightTop, textvariable=ItemName, font=('arial', 12, 'bold'))
        self.txtItemName.grid(row=1, column=1)

        self.lblQty = Label(DataFrameRightTop, font=('arial', 12, 'bold'), text="Qty:", padx=2, pady=2)
        self.lblQty.grid(row=2, column=0, sticky=W)

        self.txtQty = Entry(DataFrameRightTop, textvariable=Qty, font=('arial', 12, 'bold'))
        self.txtQty.grid(row=2, column=1)

        self.lblRate = Label(DataFrameRightTop, font=('arial', 12, 'bold'), text="Rate:", padx=2, pady=2)
        self.lblRate.grid(row=3, column=0, sticky=W)

        self.txtRate = Entry(DataFrameRightTop, textvariable=Rate, font=('arial', 12, 'bold'))
        self.txtRate.grid(row=3, column=1)

        # =================================================DataFrameRightBottom==================================================

        self.btnAddItem = Button(DataFrameRightBottom, text="Add Item", font=('arial', 12, 'bold'), width=12, bd=4,
                                 command=iAddItem)
        self.btnAddItem.grid(row=0, column=0)

        self.btnRemoveItem = Button(DataFrameRightBottom, text="Remove Item", font=('arial', 12, 'bold'), width=12,
                                    bd=4, command=iDeleteItem)
        self.btnRemoveItem.grid(row=0, column=1)

        # =================================================ButtonFrame===========================================================

        self.btnSave = Button(ButtonFrame, text="Save", font=('arial', 12, 'bold'), width=25, bd=4, command=iSave)
        self.btnSave.grid(row=1, column=0)

        self.btnPrint = Button(ButtonFrame, text="Print", font=('arial', 12, 'bold'), width=25, bd=4, command=iPrint)
        self.btnPrint.grid(row=1, column=1)

        self.btnExit = Button(ButtonFrame, text="Exit", font=('airal', 12, 'bold'), width=25, bd=4, command=iExit)
        self.btnExit.grid(row=1, column=2)

        # =================================================FrameDetail===========================================================

        self.lblLabel = Label(FrameDetail, font=('arial', 10, 'bold'), text="SNo\tHSNCode\tItemName\tRate\tQty\tAmount")
        self.lblLabel.grid(row=0, column=0)

        self.txtFrameDetail = Text(FrameDetail, font=('arial', 12, 'bold'), width=140, height=4, padx=2, pady=2)
        self.txtFrameDetail.grid(row=1, column=0)


if __name__ == '__main__':
    root = Tk()
    application = invoice(root)
    root.mainloop()
