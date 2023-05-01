import tkinter
from tkinter import *
from tkinter import filedialog
from tkinter.ttk import Combobox
import pandas as pd

class Backend():
    def __init__(self,):
        self.week_var = None
        self.liste2 = None
        self.combo_file_type = None
        self.liste1 = None
        self.combo = None
    def browse_file(self):
        # Dosya seçme arayüzü oluşturma
        self.file_path = filedialog.askopenfilename(filetypes=[("Excel Dosyaları", "*.xlsx"), ("Bütün Dosyalar", "*.*")])
        # Excel dosyasını okuyun
        self.df2 = pd.read_excel(self.file_path, usecols=["Section"])

        self.sections = self.df2["Section"].tolist()
        self.combo["values"] = sorted(set(self.sections))

        self.combo.current(0)  # section 1 varsayılı değeri gösterir
        self.combo_file_type.current(0)  # txt varsayılı değeri gösterir

        self.on_Select()  # Section 1 bilgilerini yüklüyor


    def on_Select(self, event=None):
        self.selected = self.combo.get()
        self.df = pd.read_excel(self.file_path, usecols=["Id", "Name", "Section", "Department"], engine='openpyxl')

        # Verileri filtrele
        self.filtered_df = self.df[self.df["Section"] == self.selected][["Name", "Id", "Department"]]

        self.names = []
        self.dep_df = []
        for name, dep in zip(self.filtered_df["Name"].tolist(), self.filtered_df["Department"].tolist()):
            # isim ve soyisimleri ayırdığımız yer
            name_parts = name.split()
            # İnsanların 1 soyadı olduğu için soy ismini başa alırız ve geri kalan kısım insanların adı olur.
            if len(name_parts) > 1:
                # ad ve soyadı ters sırayla birleştir eğer
                last_name = name_parts[-1]
                first_names = " ".join(name_parts[:-1])
                self.names.append(last_name + ", " + first_names)
            else:
                self.names.append(name_parts[0])
            self.dep_df.append(dep)

        # self.dep_df dizisindeki değerleri self.filtered_df["Department"] eşleştiriyoruz
        self.filtered_df["Department"] = self.dep_df

        self.filtered_df["Name"] = self.names
        self.filtered_df = self.filtered_df.sort_values("Name")
        self.dep_df = self.filtered_df["Department"]  # Şu an verileri sortlayıp dep_df ye attık
        # for i in self.dep_df:
        #     print(i)
        # self.dep_df=self.filtered_df["Department"]
        self.name_id_column = self.filtered_df['Name'].astype(str) + ' , ' + self.filtered_df['Id'].astype(str)
        self.name_id_list = self.name_id_column.tolist()
        # self.name_id_list.sort() #Alfabetik olarak sıralanıyor
        # Listbox'ta göster
        self.liste1.delete(0, tkinter.END)
        self.liste2.delete(0, tkinter.END)
        for name_id, dep in zip(self.name_id_list, self.dep_df):
            self.liste1.insert(tkinter.END, name_id)


    def add_items(self):
        self.selected_items = self.liste1.curselection()
        for item in self.selected_items:
            self.liste2.insert(END, self.liste1.get(item))


    def remove_items(self):
        self.selected_items = self.liste2.curselection()
        self.selected_items = sorted(self.selected_items, reverse=True)
        for item in self.selected_items:
            self.liste2.delete(item)


    def submit(self):
        self.file_type_selected = self.combo_file_type.get()
        self.week = self.week_var.get()

        print(self.selected + " Week " + self.week + "." + self.file_type_selected)
        self.dosya_adi = self.selected + " Week " + self.week

        items = [(int(name.split(',')[2].strip()), f"{name.split(',')[1].strip()} {name.split(',')[0].strip()}") for name in
                 self.liste2.get(0, tkinter.END)]
        dep_dict = dict(
            zip([f"{name.split(',')[1].strip()} {name.split(',')[0].strip()}" for name in self.filtered_df["Name"]],
                self.dep_df))

        # for i in self.dep_df:
        #     print(i+"self.dep_df")
        # for i in dep_dict:
        #     print(i+dep_dict[i])
        new_items = []

        for item in items:
            new_item = list(item)
            department = dep_dict.get(item[1], "Bulunamadı")
            new_item.append(department)
            new_items.append(new_item)

        new_df = pd.DataFrame(new_items, columns=['Id', 'Name', 'Department'])

        self.output_file_path = filedialog.asksaveasfilename(defaultextension=self.file_type_selected,
                                                             initialfile=self.dosya_adi, filetypes=[
                (f"{self.file_type_selected.upper()} Dosyaları", f"*.{self.file_type_selected.lower()}")])

        if self.output_file_path:
            if self.file_type_selected == "xls":
                with pd.ExcelWriter(self.output_file_path, engine='openpyxl') as writer:
                    new_df.to_excel(writer, index=False)

                print("Listbox verileri Excel'e kaydedildi." + self.output_file_path)
            elif self.file_type_selected == "txt":
                # txt ye kaydetme
                new_df_string = new_df.to_string(index=False)
                with open(self.output_file_path, 'w') as f:
                    f.write(new_df_string.replace('\n', '\n\n'))
                # new_df.to_csv(self.output_file_path , sep="\t", index=False) bu yazım şekli de var ama farklı
                print("Listbox verileri metin dosyasına kaydedildi.", self.output_file_path)
            elif self.file_type_selected == "csv":
                raise BaseException("CSV Dosya tipi desteklenmiyor")
        # self.file_type_var.set("")

class AttandanceKeeperApp(tkinter.Frame, Backend):
    def __init__(self, parent):
        tkinter.Frame.__init__(self, parent)
        Backend.__init__(self)
        self.init_UI()

    def init_UI(self) -> None:
        self.pack()
        # Uygulamanın adı
        self.baslik = tkinter.Label(self, text="")

        self.baslik.grid(row=0, column=0)
        self.yazi = tkinter.Label(self.baslik, text="AttandanceKeeper v1.0", font=('Arial 23 bold'))
        self.yazi.grid(row=0, column=1)

        # bas_alt
        self.bas_alt = tkinter.Label(self, text="")
        self.bas_alt.grid(row=1, column=0, ipadx=140)
        self.bas_alt_yazi = tkinter.Label(self.bas_alt, text="Select student list Excel file: ", font=('Arial 15 bold'))
        self.bas_alt_yazi.grid(row=0, column=0)
        self.bas_alt_buton = tkinter.Button(self.bas_alt, text="Import List", command=self.browse_file, width=16,
                                            height=1,
                                            font=('Arial 12 bold'))
        self.bas_alt_buton.grid(row=0, column=1)

        # orta_üst
        self.orta_üst = tkinter.Label(self, text="")
        self.orta_üst.grid(row=2, column=0)

        self.select_student_label = tkinter.Label(self.orta_üst, text="Select a Student:", font=("Ariel 16 bold"))
        self.select_student_label.grid(row=0, column=0, )

        self.section_label = tkinter.Label(self.orta_üst, text="Section:", font=("Ariel 16 bold"), padx=70, )
        self.section_label.grid(row=0, column=1)

        self.attented_student_label = tkinter.Label(self.orta_üst, text="Attented Student:", font=("Ariel 16 bold"))
        self.attented_student_label.grid(row=0, column=2, padx=50)

        # orta
        self.orta = tkinter.Label(self, text="")
        self.orta.grid(row=3, column=0)

        self.liste1 = tkinter.Listbox(self.orta, width=45, selectmode=MULTIPLE)
        self.liste1.grid(row=0, column=0)

        self.scrollbar = Scrollbar(self.orta)
        self.scrollbar.grid(row=0, column=1, sticky='ns')
        self.liste1.config(yscrollcommand=self.scrollbar.set)
        self.scrollbar.config(command=self.liste1.yview)

        self.combo = Combobox(self.orta, state="readonly")
        self.combo.grid(row=0, column=2, sticky="n", ipadx=25)
        self.combo.bind("<<ComboboxSelected>>", self.on_Select)

        self.buton1 = tkinter.Button(self.orta, text="Add =>", command=self.add_items, width=18, height=1,
                                     font=('Arial 12 bold'))
        self.buton1.grid(row=0, column=2, sticky="n", pady=24)
        self.buton2 = tkinter.Button(self.orta, text="<= Remove", command=self.remove_items, width=18, height=1,
                                     font=('Arial 12 bold'))
        self.buton2.grid(row=0, column=2, sticky="n", pady=58)

        self.liste2 = tkinter.Listbox(self.orta, width=45, selectmode=MULTIPLE)
        self.liste2.grid(row=0, column=3)

        self.scrollbar2 = Scrollbar(self.orta)
        self.scrollbar2.grid(row=0, column=4, sticky='ns')
        self.liste2.config(yscrollcommand=self.scrollbar2.set)

        self.week_var = tkinter.StringVar()  # dosya tipini alcağımız stirng değer
        # alt
        self.alt = tkinter.Label(self, text="")
        self.alt.grid(row=4, column=0)
        self.file_type_label = tkinter.Label(self.alt, text="Please select file type:", font=("Ariel 12 bold"))
        self.file_type_label.grid(row=0, column=0, padx=60, )
        self.combo_file_type = Combobox(self.alt, values=["txt", "xls", "csv"], width=6,state="readonly")
        # self.combo_file_type.current(0)
        self.combo_file_type.grid(row=0, column=0, sticky="e")
        self.bosluk = tkinter.Label(self.alt, text="")
        self.bosluk.grid(row=0, column=1, padx=70)
        self.week_label = tkinter.Label(self.alt, text="Please enter week:", font=("Ariel 12 bold"))
        self.week_label.grid(row=0, column=3)
        self.file_entry = tkinter.Entry(self.alt, textvariable=self.week_var, font=('calibre', 10, 'normal'),
                                        width=14)
        self.file_entry.grid(row=0, column=4)
        self.send_buton = tkinter.Button(self.alt, text="Export as file", command=self.submit, font=("Ariel 9 bold"))
        self.send_buton.grid(row=0, column=5)



def main():
    root = tkinter.Tk()  # Pencere oluşturma
    root.title("Başlık")
    root.resizable(False, False)
    app = AttandanceKeeperApp(root)
    root.mainloop()

main()