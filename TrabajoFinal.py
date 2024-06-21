import xlsxwriter
import os
from tkinter import *
import tkinter as tk
from tkinter import ttk

raiz=Tk()
raiz.title("Velocidad de jugadores.")
raiz.config(height=300, width=300)
raiz.resizable(False,False)
frame=Frame(raiz,background="#53EA6A")
frame.grid()
a=[]
b=[]
suma=0
def Jugadores():
    pregunta_label=Label(frame, text="Nombre del jugador: ", font=("Arial",14),background="#53EA6A")
    pregunta_label.grid(column=0,row=0,padx=10,pady=10,sticky="W")
    jugador=tk.StringVar()
    deportista_entry=Entry(frame, textvariable=jugador)
    deportista_entry.grid(column=1,row=0,padx=10,pady=10,sticky="W")
    pregunta2_label=Label(frame, text="Tiempo del jugador: ", font=("Arial",14), background="#53EA6A")
    pregunta2_label.grid(column=0,row=1,padx=10,pady=10,sticky="W")
    tiempovar=tk.StringVar()
    deportista_entry=Entry(frame, textvariable=tiempovar)
    deportista_entry.grid(column=1,row=1,padx=10,pady=10,sticky="W")
    button=ttk.Button(frame,text="Agregar")
    button.grid(column=0,row=2,padx=10,pady=10,sticky="W")
    button2=ttk.Button(frame,text="Terminar")
    button2.grid(column=1,row=2,padx=10,pady=10,sticky="W")
    def Agregar_Jug():
        jug=jugador.get()
        tiempo=float(tiempovar.get())
        a.append(jug)
        b.append(tiempo)
        Jugadores()
    def terminar():
        raiz.destroy()
    button.config(command=Agregar_Jug)
    button2.config(command=terminar)
Jugadores()
raiz.mainloop()
for i in range(len(a)):
    suma=suma+b[i]           
prom=suma/len(a)
k=len(a)
for i in range(k-1):
    for j in range(i+1,k):
        if(b[i]>b[j]):
            aux=b[i]
            b[i]=b[j]
            b[j]=aux
            aux=a[i]
            a[i]=a[j]
            a[j]=aux           
workbook = xlsxwriter.Workbook('Jugadores.xlsx')
worksheet = workbook.add_worksheet()
#colores
cell_format = workbook.add_format({"font_color":"red"})
format1=workbook.add_format({"bg_color":"#00FFFF","border":1,})
format2=workbook.add_format({"bg_color":"#FF7F00","border":1})
format3=workbook.add_format({"font_color":"white"})

worksheet.write(0,0,"Jugadores",cell_format)
worksheet.write(0,1,"Tiempo (seg)",cell_format)
worksheet.write(0,2,"Condición: ",cell_format)
worksheet.write("F1","Promedio: ",cell_format)
worksheet.write("G1",prom)

for i in range (k):
    worksheet.write(i+1,0,a[i],format1)
    worksheet.write(i+1,1,b[i],format2)
    if(b[i]>prom):
        dif=b[i]-prom
        dif=dif
        worksheet.write(i+1,2,f"No superó por: {dif}")
    elif(b[i]<prom):
        dif=prom-b[i]
        worksheet.write(i+1,2,f"Superó por: {dif}")
    else:
        worksheet.write(i+1,2,"Es igual al promedio.")
    worksheet.write(i+1,6,prom,format3) #agrego el promedio en blanco
    
graf=workbook.add_chart({'type': 'column'})
graf.add_series({
    "name":"Tiempo",
    'categories': f'=Sheet1!$A$2:$A${k+1}',
    'values': f'=Sheet1!$B$2:$B${k+1}',
    "line" : {"color" : "green"},
    "fill" : {"color" : "green"},
})
graf2=workbook.add_chart({"type":"line"})
graf2.add_series({
    'name':"Promedio",
    'values':f'=Sheet1!$G$1:$G${k}',
    "line" : {"color" : "red"},
    'marker': {'type': 'diamond'} #tipo de punto
})

graf.combine(graf2)

graf.set_title({ 'name': 'tiempo vs promedio'})
graf.set_x_axis({'name': 'Jugadores'})
graf.set_y_axis({'name': 'Tiempo'})

worksheet.insert_chart('I1', graf)
workbook.close()
print(a,b)


os.startfile("jugadores.xlsx")