import os
import tkinter as tk
from tkinter import filedialog
import pandas as pd
import matplotlib.pyplot as plt
import openpyxl
import tkcalendar as tkc
from tkinter import ttk
def read_excel_data(file_path):
    try:
        # Citire date din Excel
        df = pd.read_excel(file_path, sheet_name=None)

        # Afisare nume sheets din fisierul Excel
        print("Sheets disponibile in fisier:")
        for sheet_name in df.keys():
            print(sheet_name)

        # Citire date din sheet-ul "vacation" (verificam mai intai daca exista acest sheet)
        if "vacation" in df:
            df_vacation = df["vacation"]
            print("\nDate din sheet-ul 'vacation':")
            print(df_vacation)
        else:
            print("Sheet-ul 'vacation' nu a fost gasit in fisierul Excel.")
            return None, None

        return df, df_vacation  # Returnam df_info si df_vacation

    except Exception as e:
        print(f"Erroare la incarcarea datelor din fisierul Excel: {e}")
        return None, None  # Returnam None pentru ambele date în cazul unei erori

# Functie pentru a incarca fisierul Excel
def load_excel_file():
    global file_path
    file_path= filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    return file_path
def calculate_values(df_vacation,start_date,end_date):
    global stored_value
    days_taken={}
    test = df_vacation[stored_value].unique()
    df_vacation=df_vacation[df_vacation["Document Status"].str.contains("Approved")== True]
    df_vacation["From"] = pd.to_datetime(df_vacation["From"]).dt.date
    df_vacation=df_vacation[df_vacation["From"]>start_date]
    df_vacation=df_vacation[df_vacation["From"]<end_date]
    for managers in test:
            filtered_row=df_vacation[df_vacation[stored_value]==managers]
            days_taken[managers]=filtered_row["Att./abs. days"].sum()
    list_managers=list(days_taken.keys())
    list_total_days=list(days_taken.values())
    return list_managers,list_total_days
    
def plot_vacation(list_managers,list_total_days):
    # Exemplu de creare a unui grafic simplu cu datele din sheet-ul "vacation"
    # cod pentru crearea graficului specific pentru orice nevoie din excelurile vietii.
    # Creare grafic bar chart
    formated_names=[]
    for nume in list_managers:
        nume= str(nume)
        formated_names.append(nume)
    plt.bar(formated_names, list_total_days)
    plt.xlabel("Tipul Concediului")
    plt.ylabel("Numarul de Concedii")
    plt.title(f"Numarul de Concedii in functie de {stored_value}")
    plt.xticks(rotation=45)
    plt.show()

def get_employee_id_list(file_path,stored_value):
    if stored_value in {"Department","Factory","Project Name"}:
        df_info = pd.read_excel(file_path, sheet_name="info")
    else:
        df_info = pd.read_excel(file_path, sheet_name="vacation")
    print(df_info)
    employee = df_info[stored_value]
    return employee.unique()


def get_factory_list(file_path):
    df_info = pd.read_excel(file_path, sheet_name="info")
    print(df_info)
    filtered_info = df_info["Factory"]
    return filtered_info.unique()

def show_checklist(root, file_path, filtered_info):
    filtered_info_var = []
    calculate_chart_button=[]
    if (stored_value != "Manager") and (stored_value != "Att./Absence type"):
        for display_text in filtered_info:
            filtered_info_var.append(tk.StringVar(value=display_text))
        for idx, display_text in enumerate(filtered_info):
            label = tk.Label(root,textvariable= filtered_info_var[idx])
            label.pack()
            calculate_chart_button.append(tk.Button(root, text="See days off info", command=lambda c=display_text,d=idx: days_off_calculator(c,d,file_path,root)))
            calculate_chart_button[idx].pack()
    generate_chart_button = tk.Button(root, text="Generate Chart", command=lambda : generate_chart(file_path,root))
    generate_chart_button.pack()
def days_off_calculator(display_text,idx,file_path,root):
    global stored_value
    df_info = pd.read_excel(file_path,sheet_name='info')
    df_vacation = pd.read_excel(file_path,sheet_name='vacation')
    collumns=["ID SAP","Factory","Department","Project Name","Entitlement 2023"]
    merged_df=df_vacation.merge(df_info[collumns],left_on="Employee ID", right_on="ID SAP",how="inner")
    selected_entity,total_days=calculate_values(merged_df,start_date,end_date)
    T = tk.Text(root, height = 5, width = 52)
    if stored_value == "Employee":
        filtered_row=merged_df[merged_df[stored_value]==display_text]
        entitlement_days= filtered_row["Entitlement 2023"]
        days_left=entitlement_days-total_days[idx]
        display_entitlement="".join(str(e) for e in entitlement_days.unique())
        T.insert(tk.END, f"{display_text} has : {total_days[idx]} days taken, {display_entitlement} entitled days \n")
    else:
        filtered_row=df_info[df_info[stored_value]==display_text]
        entitlement_days=filtered_row["Entitlement 2023"].sum()
        print(total_days)
        days_left=entitlement_days-total_days[idx]   
        T.insert(tk.END, f"{display_text} has : {total_days[idx]} days taken, {entitlement_days} entitled days \n")
    T.pack()        



def generate_chart(file_path,root):
        df_info = pd.read_excel(file_path,sheet_name='info')
        df_vacation = pd.read_excel(file_path,sheet_name='vacation')
        collumns=["ID SAP","Factory","Department","Project Name","Entitlement 2023"]
        merged_df=df_vacation.merge(df_info[collumns],left_on="Employee ID", right_on="ID SAP",how="inner")
        merged_df.to_excel("output.xlsx",index=False)
        global stored_value
        global selected_entity,total_days
        if file_path:
            selected_entity,total_days=calculate_values(merged_df,start_date,end_date)
            plot_vacation(selected_entity,total_days)
def print_answers(selected_option):
    global stored_value
    stored_value = "{}".format(selected_option.get())

def create_gui():
    root = tk.Tk()
    root.title("Aplicatie Concedii Angajati")
    label_start_data=""
    def dateentry_view_start():
        global top
        def print_sel():
            global start_date
            start_date=cal.get_date()
            
            print(cal.get_date())
            top.destroy()
        top = tk.Toplevel(root)

        tk.Label(top, text='Choose date')
        cal = tkc.DateEntry(top)
        cal.pack()
        tk.Button(top, text="ok", command=print_sel).pack()
    def dateentry_view_end():
        global top
        def print_sel():
            global end_date
            end_date=cal.get_date()
            print(cal.get_date())
            top.destroy()
        top = tk.Toplevel(root)

        tk.Label(top, text='Choose date')
        cal = tkc.DateEntry(top)
        cal.pack()
        tk.Button(top, text="ok", command=print_sel).pack()
    
    # Obtine dimensiunile ecranului
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    # Seteaza dimensiunile GUI (jumatate din dimediniunile ecranului)
    window_width = screen_width 
    window_height = screen_height
    x_offset = (screen_width - window_width) // 2
    y_offset = (screen_height - window_height) // 2
    root.geometry(f"{window_width}x{window_height}+{x_offset}+{y_offset}")
    main_frame= tk.Frame(root)
    main_frame.pack(fill=tk.BOTH,expand=1)
    canvas= tk.Canvas(main_frame)
    canvas.pack(side=tk.LEFT, fill=tk.BOTH,expand=1)
    scrollbar = ttk.Scrollbar(main_frame,orient=tk.VERTICAL,command=canvas.yview)
    scrollbar.pack(side=tk.RIGHT,fill=tk.Y)
    canvas.configure(yscrollcommand=scrollbar.set)
    canvas.bind('<Configure>',lambda e:canvas.configure(scrollregion=canvas.bbox("all")))
    second_frame=tk.Frame(canvas)
    canvas.create_window((0,0),window=second_frame,anchor="nw")
    tk.Button(main_frame, text='Start Date', command=dateentry_view_start).pack()
    tk.Button(main_frame, text='End Date', command=dateentry_view_end).pack()
    def on_load_button_click():
        load_excel_file()  # Nu transmitem niciun argument aici
        filtered_info = get_employee_id_list(file_path,stored_value)
        show_checklist(second_frame, file_path, filtered_info)
    options_list = ["Factory", "Department", "Project Name", "Manager","Employee","Att./Absence type"]
    selected_option = tk.StringVar(second_frame)
    selected_option.set("Pick a domain")
    question_menu = tk.OptionMenu(second_frame, selected_option,*options_list)
    question_menu.pack()
    submit_button = tk.Button(second_frame, text='Submit', command=lambda:print_answers(selected_option))
    submit_button.pack()
    load_button = tk.Button(second_frame, text="Incarca fisier Excel", command=on_load_button_click)
    load_button.pack()
 
    root.mainloop()

if __name__ == "__main__":
    create_gui()
