#!/usr/bin/python3.6

from anastruct import SystemElements
import uno
import subprocess as SP
import sys
import tkinter as tk
from tkinter import filedialog
import os
import numpy as np
import matplotlib.pyplot as plt

def rround(x): #define precision
    return round(x,3)



application_window = tk.Tk()
application_window.title("Calc")
screen_width = application_window.winfo_screenwidth()
screen_height = application_window.winfo_screenheight()
application_window.geometry("550x250+%d+%d" % (screen_width/2-275, screen_height/2-125))
application_window.lift()


def createWin():
    global application_window
    application_window = tk.Tk()
    application_window.title("Calc")
    screen_width = application_window.winfo_screenwidth()
    screen_height = application_window.winfo_screenheight()
    application_window.geometry("550x250+%d+%d" % (screen_width/2-275, screen_height/2-125))
    application_window.lift()

def addlabel(field):
    w=tk.Label(application_window, text=field)
    w.pack()

def valueInput(field):
    text = tk.simpledialog.askfloat("Valor: ", field,
                               parent=application_window,
                               )    
    return float(text)

ctx=None
def con(j=None,i=None,relative=False):
    global ctx
    # get the uno component context from the PyUNO runtime
    localContext = uno.getComponentContext()

    # create the UnoUrlResolver
    resolver = localContext.ServiceManager.createInstanceWithContext(
                    "com.sun.star.bridge.UnoUrlResolver", localContext )

    # connect to the running office
    ctx = resolver.resolve( "uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext" )
    smgr = ctx.ServiceManager

    # get the central desktop object
    desktop = smgr.createInstanceWithContext( "com.sun.star.frame.Desktop",ctx)

    # access the current writer document
    model = desktop.getCurrentComponent()
    # access the document's text property
    text = model.CurrentSelection

    sheets = model.getCurrentController()
    if i==None or j==None:
        i=text.CellAddress.Column
        j=text.CellAddress.Row
    if relative:
        i=text.CellAddress.Column+i        
        j=text.CellAddress.Row+j
       
    cell = sheets.getActiveSheet().getCellByPosition(i, j)
    return cell

def readSelected():
    cell=con()
    global ctx    
    load=cell.Value if type(cell.Value) is float else 0    
    ctx.ServiceManager
    return load

def readCell(i,j):
    cell=con(i,j)
    global ctx    
    load=cell.Value if type(cell.Value) is float else 0    
    ctx.ServiceManager
    return load

def writeSelectedCell(value,i=0,j=0):
    cell=con(i,j,relative=True)
    global ctx    
    cell.setFormula(str(value))
    ctx.ServiceManager
 

def loads(L):

    if len(L)<1:
        print("invalid argument!")
        return 0

    EI=readCell(0,0)
    EA=readCell(0,1)
    ss = SystemElements(EA=EA, EI=EI)  
    addlabel("EI: "+str(EI))     
    addlabel("EA: "+str(EA))  
    pl=0
    i=1
    start=[0,0]

    for l in L:
        end=[start[0]+l,0]
        ss.add_element(location=[start,end])
        start=end

    answer = tk.messagebox.askokcancel("Carga Concentrada","Existe carga Pontual?")
    if(answer):
        X=L.copy()
        X.insert(0,0)
        wasHingedinsert=False
        for l in X:
            pl=valueInput("\n carga pontual at id " + str(i)+": ")
            pl = float(pl) if (type(pl) is float or type(pl) is float) and pl>0 else -1
            if pl > 0:
                ss.point_load(Fy=-pl, node_id=i) 
                addlabel("Carga pontual no nó: " +str(i)+"  "+str(pl))
            else:
                if not wasHingedinsert:
                    ss.add_support_hinged(node_id=i)
                    wasHingedinsert=True
                else:
                    ss.add_support_roll(node_id=i)
            i+=1
    else:
        ss.add_support_hinged(node_id=i)
        i+=1
        for l in L:
            ss.add_support_roll(node_id=i)
            i+=1  

    load=readSelected()
    load =float(load) if type(load) is float else 0
    canswer = tk.messagebox.askokcancel("Carga distribuida","Aplicar celula celecionada ("+str(load)+")em toda estrutura?")    
    if canswer:
        i=1
        for l in L:
            ss.q_load(q=-load, element_id=i)
            i+=1  

        addlabel("Carga distribuída em toda estrutura: "+str(load))

    answer = tk.messagebox.askokcancel("Carga distribuida","Aperte cancel para mais cargas distribuídas")    

    if not answer:
        add=0 if not canswer else load
        i=1
        for l in L:
            load=valueInput("\n Concentrated load at id " + str(i)+": ")
            load = float(load) + add if type(load) is float and load>0 else 0
            if load!=0:
                ss.q_load(q=-load, element_id=i)
                addlabel("Elemento: "+str(i)+" carga: "+str(load))
            i+=1

    ss.solve()
    #ss.show_reaction_force()
    #ss.show_axial_force()
    ss.show_structure()       
    ss.show_shear_force()    
    ss.show_bending_moment()

    createWin()
    answer = tk.messagebox.askokcancel("Question","Wanna save?")
    path=False
    if answer:
        path=filedialog.asksaveasfilename(parent=application_window,
                                           initialdir="/media/matheus/OS/Users/mathe/Desktop/Concreto/pt2/Diagramas/",
                                           title="Please select a file name for saving:")
 
    if path:
        fig=ss.show_bending_moment(show=False)
        plt.title('Momento Fletor '+os.path.basename(path))
        plt.savefig(path+"_DEM.png")        
        fig=ss.show_structure(show=False)       
        plt.title('Estrutura '+os.path.basename(path))
        plt.savefig(path+"_Estrutura.png")
        fig=ss.show_shear_force(show=False)    
        plt.title('Esforço Cortante '+os.path.basename(path))
        plt.savefig(path+"_DEC.png")


#    fig=ss.show_displacement()
#    if path:
#        plt.title('Deslocamentos '+os.path.basename(path))
#        plt.savefig(path+".png")

    i=0
    O=ss.get_element_results(element_id=0)
    Mmin=[0]
    Mmax=[0]
    S=[]
    for l in O:
        if l["Mmin"]<0:
            Mmin.append(abs(l['Mmin']))
        if l["Mmax"]>0:
            Mmax.append(abs(l['Mmax']))

    #reactions
    v=ss.get_node_results_system(node_id=i+1)['Fy']
    writeSelectedCell(rround(v),10+i,0) 
    i+=1    
    for l in L:
        v=ss.get_node_results_system(node_id=i+1)['Fy']
        writeSelectedCell(rround(v),10+i,0)
        i+=1

 
    S=ss.get_element_result_range('shear') 
    Sh=[]
    for s in S:
        Sh.append(abs(s))

    writeSelectedCell(rround(max(Mmin)),0,2)
    writeSelectedCell(rround(max(Mmax)),0,3)
    writeSelectedCell(max(Sh),0,1)



if __name__ == "__main__":  
    L=eval(sys.argv[1])
    loads(L) 
#   loads([3.5,3.5]) 
#   application_window.destroy()