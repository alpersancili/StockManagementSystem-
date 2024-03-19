# Import necessary modules for the application
from tkinter import *
from tkinter import ttk
from tkinter import messagebox
import tkinter
import random
# Import the pymysql module, which provides tools for interacting with MySQL databases
import pymysql
# Import the csv module, which provides functionality for reading and writing CSV files
import csv
from datetime import datetime
# It allows you to create and manipulate Excel workbooks
from openpyxl import Workbook

# Create the main application window
window = tkinter.Tk()
window.title("Stock Management System")
window.geometry("720x640")
# Create a Treeview widget for displaying data
my_tree = ttk.Treeview(window,show="headings",height=20)
style = ttk.Style()

# Define arrays and strings for generating random item IDs
placeholderArray=["","","","",""]
numeric="1234567890"
alpha="ABCDEFGHIJKLMNOPQRSTUVWXYZ"

# Define a function to establish a connection to the MySQL database
def connection():
    conn=pymysql.connect(
        host="localhost",
        user="root",
        password="",
        #database name
        db="reservation"
    )
    return conn

# Establish a connection to the database and create a cursor object
conn=connection()
cursor=conn.cursor()

# Initialize placeholderArray with Tkinter StringVar objects
for i in range(0,5):
    placeholderArray[i] = tkinter.StringVar()

# Define a function to retrieve data from the "stocks" table in the database
def read():
    # Check if the database connection is still alive
    cursor.connection.ping()
    # Construct an SQL query to select data from the "stocks" table
    sql=f"SELECT `item_id`, `name`, `price`, `quantity`, `category`, `date` FROM stocks ORDER BY `id` DESC"
    # Execute the SQL query and fetch all the data
    cursor.execute(sql)
    results=cursor.fetchall()
    conn.commit()
    conn.close()
    return results

# Define a function to update the Treeview widget with the latest data from the database
def refreshTable():
    # Iterate through all items in the Treeview widget and delete them
    for data in my_tree.get_children():
        my_tree.delete(data)
    # Iterate through the data retrieved from the database using the read() function
    for array in read():
        # Insert a new item into the Treeview widget with specified attributes
        # parent="" means it's a top-level item
        # index="end" means it's added at the end
        # iid=array sets the item's unique identifier
        # text="" is an empty string for the item's text
        # values=(array) sets the values for each column in the Treeview
        # tags="orow" assigns a tag to the item for styling purposes
        my_tree.insert(parent="",index="end",iid=array,text="",values=(array),tags="orow")
    # Configure the style of items with the "orow" tag, setting the background color to #EEEEEE
    my_tree.tag_configure("orow",background="#EEEEEE")
    # Pack the Treeview widget to ensure it's displayed
    my_tree.pack()

# Define a function to set placeholder values in the array
def setph(word,num):
    for ph in range(0,5):
        if ph == num:
            # Set the value of the corresponding StringVar in placeholderArray to the given 'word'

            placeholderArray[ph].set(word)

# Define a function to generate a random item ID
def generateRand():
    itemId=""
    for i in range(0,3):
        randno=random.randrange(0,(len(numeric)-1))
        itemId=itemId+str(numeric[randno])
    randno=random.randrange(0,(len(alpha)-1))
    itemId=itemId+"-"+str(alpha[randno])
    print("generated: "+itemId)
    setph(itemId,0)

# Define a function to save data to the database and refresh the table
def save():
    itemId=str(itemIdEntry.get())
    name=str(nameEntry.get())
    price=str(priceEntry.get())
    qnt=str(qntEntry.get())
    cat=str(categoryCombo.get())
    valid=True
    # Check if all entries are filled
    if not(itemId and itemId.strip()) or not(name and name.strip()) or not(price and price.strip()) or not(qnt and qnt.strip()) or not(cat and cat.strip()):
        messagebox.showwarning("", "Please fill up all entries")
        return
    # Validate item ID format
    if len(itemId) < 5:
        messagebox.showwarning("","Invalid Item Id")
        return
    if(not(itemId[3]=="-")):
        valid=False
    for i in range(0,3):
        if(not(itemId[i] in numeric)):
            valid=False
            break
    if(not(itemId[4] in alpha)):
        valid=False
    if not(valid):
        messagebox.showwarning("","Invalid Item Id")
        return
    try:
        # Check if the database connection is still alive
        cursor.connection.ping()
        # Check if the item ID already exists in the database
        sql = f"SELECT * FROM stocks WHERE `item_id`= '{itemId}' "
        # Execute the SQL query and fetch all the data
        cursor.execute(sql)
        checkItemNo=cursor.fetchall()
        if len(checkItemNo) > 0:
            messagebox.showwarning("","Item Id already used")
            return
        else:
            # Check if the database connection is still alive
            cursor.connection.ping()
            # Insert data into the database
            sql=f"INSERT INTO stocks (`item_id`, `name`, `price`, `quantity`, `category`) VALUES ('{itemId}','{name}','{price}','{qnt}','{cat}')"
            cursor.execute(sql)
        conn.commit()
        conn.close()
        # Clear placeholder values
        for num in range(0,5):
            setph("",(num))
    except:
        messagebox.showwarning("", "Error while saving")
        return
    # Refresh the table with updated data
    refreshTable()

# Define a function to update data in the database and refresh the table
def update():
    selectedItemId = ""
    try:
        # Get the selected item ID from the Treeview widget
        selectedItem = my_tree.selection()[0]
        selectedItemId = str(my_tree.item(selectedItem)["values"][0])
    except:
        messagebox.showwarning("", "Please select a data row")
    print(selectedItemId)
    itemId = str(itemIdEntry.get())
    name = str(nameEntry.get())
    price = str(priceEntry.get())
    qnt = str(qntEntry.get())
    cat = str(categoryCombo.get())
    # Check if all entries are filled
    if not(itemId and itemId.strip()) or not(name and name.strip()) or not(price and price.strip()) or not(qnt and qnt.strip()) or not(cat and cat.strip()):
        messagebox.showwarning("", "Please fill up all entries")
        return
    # Check if the item ID matches the selected item ID
    if(selectedItemId!=itemId):
        messagebox.showwarning("", "You can't change Item ID")
        return
    try:
        # Check if the database connection is still alive
        cursor.connection.ping()
        # Update data in the database
        sql = f"UPDATE stocks SET `name` = '{name}', `price` = '{price}', `quantity` = '{qnt}', `category` = '{cat}' WHERE `item_id` = '{itemId}' "
        cursor.execute(sql)
        conn.commit()
        conn.close()
        # Clear placeholder values
        for num in range(0,5):
            setph('',(num))
    except Exception as err:
        messagebox.showwarning("","Error occured ref: "+str(err))
        return
    # Refresh the table with updated data
    refreshTable()
# Define a function to delete data from the database and refresh the table
def delete():
    try:
        # Check if any item is selected in the Treeview widget
        if(my_tree.selection()[0]):
            decision = messagebox.askquestion("","Delete the selected data?")
            if(decision != "yes"):
                return
            else:
                # Retrieve the selected item and its Item ID
                selectedItem = my_tree.selection()[0]
                itemId = str(my_tree.item(selectedItem)["values"][0])
                try:
                    # Check if the database connection is still alive
                    cursor.connection.ping()
                    # Construct and execute an SQL query to delete data from the database
                    sql = f"DELETE FROM stocks WHERE `item_id` = '{itemId}' "
                    cursor.execute(sql)
                    conn.commit()
                    conn.close()
                    messagebox.showinfo("","Data has been succesfully deleted")
                except:
                    messagebox.showinfo("", "Sorry an error occured")
                refreshTable()
    except:
        messagebox.showinfo("", "Please select a data row")


def select():
    try:
        selectedItem = my_tree.selection()[0]
        itemId = str(my_tree.item(selectedItem)["values"][0])
        name = str(my_tree.item(selectedItem)["values"][1])
        price = str(my_tree.item(selectedItem)["values"][2])
        qnt = str(my_tree.item(selectedItem)["values"][3])
        cat = str(my_tree.item(selectedItem)["values"][4])
        setph(itemId,0)
        setph(name,1)
        setph(price,2)
        setph(qnt,3)
        setph(cat,4)
    except:
        messagebox.showwarning("", "Please select a data row")

def find():
    itemId = str(itemIdEntry.get())
    name = str(nameEntry.get())
    price = str(priceEntry.get())
    qnt = str(qntEntry.get())
    cat = str(categoryCombo.get())
    cursor.connection.ping()
    if(itemId and itemId.strip()):
        # Check which entry field has a value and construct an SQL query accordingly
        sql = f"SELECT `item_id`, `name`, `price`, `quantity`, `category`, `date` FROM stocks WHERE `item_id` LIKE '%{itemId}%' "
    elif (name and name.strip()):
        sql = f"SELECT `item_id`, `name`, `price`, `quantity`, `category`, `date` FROM stocks WHERE `name` LIKE '%{name}%' "
    elif (price and price.strip()):
        sql = f"SELECT `item_id`, `name`, `price`, `quantity`, `category`, `date` FROM stocks WHERE `price` LIKE '%{price}%' "
    elif (qnt and qnt.strip()):
        sql = f"SELECT `item_id`, `name`, `price`, `quantity`, `category`, `date` FROM stocks WHERE `qnt` LIKE '%{qnt}%' "
    elif (cat and cat.strip()):
        sql = f"SELECT `item_id`, `name`, `price`, `quantity`, `category`, `date` FROM stocks WHERE `cat` LIKE '%{cat}%' "
    else:
        messagebox.showwarning("","Please fill up one of the entries")
        return
    cursor.execute(sql)
    try:
        result = cursor.fetchall()
        for num in range(0, 5):
            setph(result[0][num], (num))
        conn.commit()
        conn.close()
    except:
        messagebox.showwarning("","No data found")

def clear():
    for num in range(0,5):
        setph("",(num))

def exportExcel():
    # Check if the database connection is still alive
    cursor.connection.ping()
    # Construct an SQL query to select data from the "stocks" table
    sql = f"SELECT `item_id`, `name`, `price`, `quantity`, `category`, `date` FROM stocks ORDER BY `id` DESC"
    # Execute the SQL query and fetch all the data
    cursor.execute(sql)
    dataraw = cursor.fetchall()

    # Create a new Excel workbook and select the active sheet
    wb = Workbook()
    ws = wb.active

    # Write the headers to the Excel file
    headers = ["item_id", "name", "price", "quantity", "category", "date"]
    ws.append(headers)

    # Write the data to the Excel file
    for record in dataraw:
        ws.append(record)

    # Save the Excel file with a timestamped filename
    date = datetime.now().strftime("%Y-%m-%d_%H-%M")
    excel_filename = f"stocks_{date}.xlsx"
    wb.save(excel_filename)

    print(f"Saved: {excel_filename}")

    cursor.connection.commit()
    cursor.connection.close()

    messagebox.showinfo("", "Excel file downloaded")


frame = tkinter.Frame(window,bg="#02577A")
frame.pack()

btnCoLor = "#196E78"

manageFrame = tkinter.LabelFrame(frame,text="Manage",borderwidth=5)
manageFrame.grid(row=0,column=0,sticky="w",padx=[10,200],pady=20,ipadx=[6])

saveBtn = Button(manageFrame,text="SAVE",width=10,borderwidth=3,bg=btnCoLor,fg="white",command=save)
updateBtn = Button(manageFrame,text="UPDATE",width=10,borderwidth=3,bg=btnCoLor,fg="white",command=update)
deleteBtn = Button(manageFrame,text="DELETE",width=10,borderwidth=3,bg=btnCoLor,fg="white",command=delete)
selectBtn = Button(manageFrame,text="SELECT",width=10,borderwidth=3,bg=btnCoLor,fg="white",command=select)
findBtn = Button(manageFrame,text="FİND",width=10,borderwidth=3,bg=btnCoLor,fg="white",command=find)
clearBtn = Button(manageFrame,text="CLEAR",width=10,borderwidth=3,bg=btnCoLor,fg="white",command=clear)
exportBtn = Button(manageFrame,text="EXPORT",width=10,borderwidth=3,bg=btnCoLor,fg="white",command=exportExcel)

saveBtn.grid(row=0,column=0,padx=5,pady=5)
updateBtn.grid(row=0,column=1,padx=5,pady=5)
deleteBtn.grid(row=0,column=2,padx=5,pady=5)
selectBtn.grid(row=0,column=3,padx=5,pady=5)
findBtn.grid(row=0,column=4,padx=5,pady=5)
clearBtn.grid(row=0,column=5,padx=5,pady=5)
exportBtn.grid(row=0,column=6,padx=5,pady=5)

entriesFrame = tkinter.LabelFrame(frame,text="Form",borderwidth=5)
entriesFrame.grid(row=1,column=0,sticky="w",padx=[10,200],pady=[0,20],ipadx=[6])

itemIdLabel = Label(entriesFrame,text="ITEM ID",anchor="e",width=10)
nameLabel = Label(entriesFrame,text="NAME",anchor="e",width=10)
priceLabel = Label(entriesFrame,text="PRICE",anchor="e",width=10)
qntLabel = Label(entriesFrame,text="QUANTITY",anchor="e",width=10)
categoryLabel = Label(entriesFrame,text="CATEGORY",anchor="e",width=10)

itemIdLabel.grid(row=0,column=0,pady=10)
nameLabel.grid(row=1,column=0,pady=10)
priceLabel.grid(row=2,column=0,pady=10)
qntLabel.grid(row=3,column=0,pady=10)
categoryLabel.grid(row=4,column=0,pady=10)

categoryArray = ["Networking Tools", "Computer Parts", "Repair Tools", "Gadgets"]

itemIdEntry = Entry(entriesFrame,width=50,textvariable=placeholderArray[0])
nameEntry = Entry(entriesFrame,width=50,textvariable=placeholderArray[1])
priceEntry = Entry(entriesFrame,width=50,textvariable=placeholderArray[2])
qntEntry = Entry(entriesFrame,width=50,textvariable=placeholderArray[3])
categoryCombo = ttk.Combobox(entriesFrame,width=48,textvariable=placeholderArray[4],values=categoryArray)

itemIdEntry.grid(row=0,column=2,padx=5,pady=5)
nameEntry.grid(row=1,column=2,padx=5,pady=5)
priceEntry.grid(row=2,column=2,padx=5,pady=5)
qntEntry.grid(row=3,column=2,padx=5,pady=5)
categoryCombo.grid(row=4,column=2,padx=5,pady=5)

generateIdBtn = Button(entriesFrame,text="GENERATE ID",borderwidth=3,bg=btnCoLor,fg="white",command=generateRand)
generateIdBtn.grid(row=0,column=3,padx=5,pady=5)

style.configure(window)
my_tree["columns"] = ("Item Id","Name","Price","Quantity","Category","Date")
my_tree.column("#0",width=0,stretch=NO)
my_tree.column("Item Id",anchor=W,width=70)
my_tree.column("Name",anchor=W,width=125)
my_tree.column("Price",anchor=W,width=125)
my_tree.column("Quantity",anchor=W,width=100)
my_tree.column("Category",anchor=W,width=150)
my_tree.column("Date",anchor=W,width=150)
my_tree.heading("Item Id",text="Item Id",anchor=W)
my_tree.heading("Name",text="Name",anchor=W)
my_tree.heading("Price",text="Price",anchor=W)
my_tree.heading("Quantity",text="Quantity",anchor=W)
my_tree.heading("Category",text="Category",anchor=W)
my_tree.heading("Date",text="Date",anchor=W)

my_tree.tag_configure("orow",background="#EEEEEE")
my_tree.pack()

refreshTable()

window.resizable(False,False)
window.mainloop()









