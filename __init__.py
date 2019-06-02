bl_info = {
    "name": "BlenderCalc",
    "author": "Matheus Fillipe",
    "version": (0, 1),
    "blender": (2, 7, 0),  
    "category": "Data Extraction",
}

import bpy
import bmesh
import sys
import os
import numpy as np

path2=os.path.dirname(os.path.realpath(sys.argv[0]))

sys.path.append(path2)
sys.path.append("/usr/lib/python3/dist-packages/")

import uno

location=bpy.types.CONSOLE_HT_header
locationTextEditor=bpy.types.VIEW3D_HT_header

def rround(x): #define precision
    return round(x,2)

def runSoffice():
    # run soffice as 'server'    
    from subprocess import Popen

    officepath = 'soffice' #respectivly the full path
    calc = '--calc'
    pipe = "--accept=socket,host=localhost,port=2002;urp;StarOffice.ServiceManager --norestore"
    Popen([officepath, calc, pipe]);

def getSelVerts():
    A=[]
    mesh=bmesh.from_edit_mesh(bpy.context.object.data)
    for v in mesh.verts:
        if v.select==1:
            P=[]
            P.append(v.co.x)
            P.append(v.co.y)
            A.append(P)
   
    return A

class Soffice(bpy.types.Operator):
    """Open Soffice"""      # Use this as a tooltip for menu items and buttons.
    bl_idname = "object.soffice"        # Unique identifier for buttons and menu items to reference.
    bl_label = "Soffice"         # Display name in the interface.
    bl_options = {'REGISTER', 'UNDO'}  # Enable undo for the operator.

    def execute(self, context):        # execute() is called when running the operator.
        runSoffice()
        return {'FINISHED'}            # Lets Blender know the operator finished successfully.


class VerticeTable(bpy.types.Operator):
    """Vertex Table"""      # Use this as a tooltip for menu items and buttons.
    bl_idname = "object.vertex"        # Unique identifier for buttons and menu items to reference.
    bl_label = "Vertex Table"         # Display name in the interface.
    bl_options = {'REGISTER', 'UNDO'}  # Enable undo for the operator.

    def execute(self, context):        # execute() is called when running the operator.

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

        V=getSelVerts()


        sheets = model.getSheets()
        i=text.CellAddress.Column
        j=text.CellAddress.Row


        if len(V)==1:
            cell = sheets.getByIndex(0).getCellByPosition(i, j)
            cell.setFormula( str(rround(V[0][0])))
            j+=1            
            cell = sheets.getByIndex(0).getCellByPosition(i, j)
            cell.setFormula(str(rround(V[0][1])))

        elif len(V)>1:
            for v in V:
                cell = sheets.getByIndex(0).getCellByPosition(i, j)
                cell.setFormula( str(rround(v[0])))
                i+=1            
                cell = sheets.getByIndex(0).getCellByPosition(i, j)
                cell.setFormula(str(rround(v[1])))
                j+=1     
                i-=1                        


        ctx.ServiceManager


        return {'FINISHED'}            # Lets Blender know the operator finished successfully.


class ObjectMoveX(bpy.types.Operator):
    """Edge size"""      # Use this as a tooltip for menu items and buttons.
    bl_idname = "object.edge"        # Unique identifier for buttons and menu items to reference.
    bl_label = "Edge size"         # Display name in the interface.
    bl_options = {'REGISTER', 'UNDO'}  # Enable undo for the operator.

    def execute(self, context):        # execute() is called when running the operator.

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

        V=getSelVerts()


        sheets = model.getSheets()
        i=text.CellAddress.Column
        j=text.CellAddress.Row


        if len(V)==1:
            cell = sheets.getByIndex(0).getCellByPosition(i, j)
            cell.setFormula( str(rround(V[0][0])))
            j+=1            
            cell = sheets.getByIndex(0).getCellByPosition(i, j)
            cell.setFormula(str(rround(V[0][1])))
          
        elif len(V)==2:
            d=str(rround(np.sqrt((V[0][0]-V[1][0])**2+(V[0][1]-V[1][1])**2)))
            cell = sheets.getByIndex(0).getCellByPosition(i, j)
            cell.setFormula( d)

        elif len(V)>2:
            for x,v in enumerate(V):
                if x == len(V)-1:
                    break
                d=str(rround(np.sqrt((V[x+0][0]-V[x+1][0])**2+(V[x+0][1]-V[x+1][1])**2)))
                cell = sheets.getByIndex(0).getCellByPosition(i, j)
                cell.setFormula( d)               
                j+=1

        ctx.ServiceManager


        return {'FINISHED'}            # Lets Blender know the operator finished successfully.


def add_object_button(self, context):  
        self.layout.operator(  
        ObjectMoveX.bl_idname,  
            text=ObjectMoveX.bl_label,  
            icon='PLAY') 
            
def add_vertex_button(self, context):  
        self.layout.operator(  
        VerticeTable.bl_idname,  
            text=VerticeTable.bl_label,  
            icon='PLAY') 


def add_soffice_button(self, context):  
        self.layout.operator(  
        Soffice.bl_idname,  
            text=Soffice.bl_label,  
            icon='PLAY') 



def register():
    bpy.utils.register_class(ObjectMoveX)
    locationTextEditor.append(add_object_button)
    bpy.utils.register_class(Soffice)
    locationTextEditor.append(add_soffice_button)
    bpy.utils.register_class(VerticeTable)
    locationTextEditor.append(add_vertex_button)



    
def unregister():
    bpy.utils.unregister_class(ObjectMoveX)
    locationTextEditor.remove(add_object_button)
    bpy.utils.unregister_class(Soffice)
    locationTextEditor.remove(add_soffice_button)
    bpy.utils.unregister_class(VerticeTable)
    locationTextEditor.remove(add_vertex_button)

if __name__ == "__main__":  
    register()  