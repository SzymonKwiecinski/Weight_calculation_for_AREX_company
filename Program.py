import pandas as pd
# import openpyxl as xl
# import numpy as np
import tkinter



file = 'Przelicznik.xlsx'
xl = pd.ExcelFile(file)
print(xl.sheet_names)
sheet_name = xl.sheet_names

def Prze_kg_na_100szt(value):
    new_value = format(value/10,'.3f')
    return new_value

def Prze_100szt_na_kg(value):
    new_value = format(10/value,'.3f')
    return new_value


top = tkinter.Tk()
top.title('Program do konwersji wagowo sztukowej - AREX')
#top.geometry("500x500")

label_norma = tkinter.Label(top,text='Ustaw\n\n1: NORME    \n2: ŚREDNICE \n3: DŁUGOŚĆ ')
label_norma.grid(row=0,column=0)

label_x = tkinter.Label(top,text=' X ')
# label_x.grid(row=0,column=5)

listbox_norma = tkinter.Listbox(top,selectmode = 'SINGLE',yscrollcommand=1
                                ,height = 5,width =30)
listbox_norma.grid(row=0,column=1)
for x,y in enumerate(sheet_name):
    listbox_norma.insert(x+1,y)


listbox_dlugosc = tkinter.Listbox(top,selectmode= 'SINGLE'
                                   ,yscrollcommand=1,height = 5)
listbox_dlugosc.grid(row=0,column=6)


listbox_srednica = tkinter.Listbox(top,selectmode= 'SINGLE'
                                   ,yscrollcommand=1,height = 5)
listbox_srednica.grid(row=0,column=3)


# scrollbar = tkinter.Scrollbar(top)
# scrollbar.grid(row=0,column=4)
# scrollbar.config( command = listbox_srednica.yview )
#
# scrollbar1 = tkinter.Scrollbar(top)
# scrollbar1.grid(row=0,column=7)
# scrollbar1.config( command = listbox_dlugosc.yview )

Sheet_DIN = xl.parse()

din = ''
srednica = ''
dlugosc = ''
napis = ''

waga_1000szt = float()



def select_din():
    global Sheet_DIN
    global sheet_name_now
    listbox_srednica.delete(0,'end')
    listbox_dlugosc.delete(0, 'end')
    index = listbox_norma.curselection()[0]
    my_sheet = listbox_norma.get(index)
    sheet_name_now = my_sheet
    if my_sheet == 'Nakretki_Podkladki':
        buton_select_srednica.configure(text='Wybierz norme')
        buton_select_dlugosc.configure(text='Wybierz Średnice')
    if my_sheet != 'Nakretki_Podkladki':
        buton_select_srednica.configure(text='Wybierz Średnice')
        buton_select_dlugosc.configure(text='Wybierz Długość')
    Sheet_DIN = xl.parse(my_sheet,index_col=0)
    # print(Sheet_DIN)
    for x in Sheet_DIN.columns.values:
        listbox_srednica.insert('end',x)
    # for x in Sheet_DIN.index.values:
    #     listbox_dlugosc.insert('end', x)
    #     print(x)
    # select_rozmiar()
    global din
    din = str(my_sheet)
    label_dane.configure(text=din,font=100)



def select_srednice():
    listbox_dlugosc.delete(0,'end')
    index = listbox_srednica.curselection()[0]
    my_sheet = listbox_srednica.get(index)
    for x in Sheet_DIN.index.values:
        if Sheet_DIN.loc[x,str(my_sheet)]!= 'None':
            listbox_dlugosc.insert('end', x)
    global srednica
    srednica = str(my_sheet)
    if sheet_name_now == 'Nakretki_Podkladki':
        label_dane.configure(text=(srednica))
    if sheet_name_now != 'Nakretki_Podkladki':
        label_dane.configure(text=(din + ' ' + srednica))

def select_dlugosc():
    index = listbox_dlugosc.curselection()[0]
    my_sheet = listbox_dlugosc.get(index)
    global dlugosc
    global waga_1000szt
    dlugosc = str(my_sheet)
    # label_dane.configure(text=(din+' '+srednica+'x'+str(dlugosc)))
    if sheet_name_now == 'Nakretki_Podkladki':
        label_dane.configure(text=(srednica + '  ' + str(dlugosc)))
        var = Sheet_DIN.loc[dlugosc, srednica]
        var = round(var,2)
        print(type(var))
        print(var)
        label_dane_z_tab_wyp.configure(font=15,text=' '+str(var)+' kg.'+'\n '+str(Prze_100szt_na_kg(var)+'\n '+str(Prze_kg_na_100szt(var))) )
        waga_1000szt = float(var)
    if sheet_name_now != 'Nakretki_Podkladki':
        label_dane.configure(text=(din + ' ' + srednica + 'x' + str(dlugosc)))
        var = Sheet_DIN.loc[float(dlugosc),srednica]
        label_dane_z_tab_wyp.configure(font=15,text=' '+str(var)+' kg.'+'\n '+str(Prze_100szt_na_kg(var)+'\n '+str(Prze_kg_na_100szt(var))) )
        waga_1000szt = float(var)

def oblicz_ile_to_kg():
    if entry_oblicz.get() != '':
        new_sztuki = float(entry_oblicz.get().replace(',','.'))
        kilogramy = (new_sztuki * waga_1000szt) / 1000.0
        label_wynik_obliczen.config(text=din + ' ' + srednica + 'x' + dlugosc +'\n'+ str(new_sztuki) + ' szt  =  '+ str(kilogramy)+' kg')

def oblicz_ile_to_sztuk():
    if entry_oblicz.get() != '':
        new_kg = float(entry_oblicz.get().replace(',','.'))
        sztuki = (new_kg / waga_1000szt) * 1000
        # label_wynik_obliczen.config(text=str(sztuki))
        label_wynik_obliczen.config(text=din + ' ' + srednica + 'x' + dlugosc +'\n'+ str(new_kg) + ' kg  =  '+ str(sztuki)+' szt')



buton_select_din = tkinter.Button(top, text='Wybierz Normę', command=select_din)
buton_select_din.grid(row=1,column=1)

buton_select_srednica = tkinter.Button(top, text='Wybierz Średnice', command=select_srednice)
buton_select_srednica.grid(row=1,column=3)

buton_select_dlugosc = tkinter.Button(top, text='Wybierz Dlugość', command=select_dlugosc)
buton_select_dlugosc.grid(row=1,column=6)

label_dane = tkinter.Label(top)
label_dane.grid(column=8,row=0,columnspan=2)

label_dane_z_tab = tkinter.Label(top)
label_dane_z_tab.grid(column=8,row=1)
label_dane_z_tab.configure(font=15,text='            1000 sztuk waży :\nPrzelicznik 100szt. na kg. :\nPrzelicznik kg. na 100szt. :')

label_dane_z_tab_wyp = tkinter.Label(top)
label_dane_z_tab_wyp.grid(column=9,row=1,padx=60)
# label_dane.configure(text='asdasd')
# ,columnspan=6
label_oblicz_sam = tkinter.Label(top)
label_oblicz_sam.grid(column=8,row=2,columnspan=2)
label_oblicz_sam.configure(font=15,text='\nWpisz sztuki lub kilogramy aby\n przeliczyć na inną jednostkę')

entry_oblicz = tkinter.Entry(top)
entry_oblicz.grid(row=3,column=8,columnspan=2)

buton_oblicz_kg = tkinter.Button(top,text='Oblicz ile\nto kilogramów',command=oblicz_ile_to_kg)
buton_oblicz_kg.grid(row=4,column=8)

buton_oblicz_szt = tkinter.Button(top,text="Oblicz ile\nto sztuk",command=oblicz_ile_to_sztuk)
buton_oblicz_szt.grid(row=4,column=9)

label_wynik_obliczen = tkinter.Label(top)
label_wynik_obliczen.grid(row=5, column=8,columnspan=2)
label_wynik_obliczen.config(font=10)

top.mainloop()
