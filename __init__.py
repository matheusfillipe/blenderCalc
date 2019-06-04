bl_info = {
    "name": "BlenderCalc",
    "author": "Matheus Fillipe",
    "version": (0, 1, 5),
    "blender": (2, 80, 0),  
    "category": "Data Extraction",
}

import bpy
import bmesh
import sys
import os
#TODO correct addon path, find alternativer to popen call on runBeam function, embed uno and anastruct libraries 
#TODO Add single menu instead of button mess

from itertools import permutations
import subprocess as SP
import numpy as np
#Replace with yours:
ADDON_PATH="/home/matheus/.config/blender/2.80/scripts/addons/blenderCalc/"

sys.path.append("/usr/lib/python3/dist-packages/")
import uno


location=bpy.types.CONSOLE_HT_header
locationTextEditor=bpy.types.VIEW3D_HT_header

def rround(x): #define precision
    return round(x,3)

def runSoffice():
    # run soffice as 'server'    
    from subprocess import Popen
    officepath = 'libreoffice' #respectivly the full path
    calc = '--calc'
    pipe = "--accept=socket,host=localhost,port=2002;urp;StarOffice.ServiceManager --norestore"
    Popen([officepath, calc, pipe]);

def runBeam(args):
    # run soffice as 'server'    
    from subprocess import Popen
    script_file = ADDON_PATH
    directory = script_file#os.path.dirname(script_file)
    officepath = directory+'beam.py' #respectivly the full path
    pipe = str(args)
    Popen([officepath, pipe]);


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

def PolygonArea(corners):
    n = len(corners) # of corners
    area = 0.0
    for i in range(n):
        j = (i + 1) % n
        area += corners[i][0] * corners[j][1]
        area -= corners[j][0] * corners[i][1]
    area = abs(area) / 2.0
    return area


class Soffice(bpy.types.Operator):
    """Open Soffice"""      # Use this as a tooltip for menu items and buttons.
    bl_idname = "object.soffice"        # Unique identifier for buttons and menu items to reference.
    bl_label = "Soffice"         # Display name in the interface.
    bl_options = {'REGISTER', 'UNDO'}  # Enable undo for the operator.

    def execute(self, context):        # execute() is called when running the operator.
        runSoffice()
        return {'FINISHED'}            # Lets Blender know the operator finished successfully.




class Beam(bpy.types.Operator):
    """Beam"""      # Use this as a tooltip for menu items and buttons.
    bl_idname = "object.beam"        # Unique identifier for buttons and menu items to reference.
    bl_label = "Beam"         # Display name in the interface.
    bl_options = {'REGISTER', 'UNDO'}  # Enable undo for the operator.

    def execute(self, context):        # execute() is called when running the operator.
        V=getSelVerts()
        if len(V)==2:            
                d=rround(np.sqrt((V[0][0]-V[1][0])**2+(V[0][1]-V[1][1])**2))
                runBeam([d])
               
        elif len(V)>2:
            L=[]
            perm=[list(p) for p in permutations(V)]
            P=[]
            for p in perm:
                S=[]
                for x,v in enumerate(p):
                    if x == len(p)-1:
                        break                        
                    d=np.sqrt((p[x+0][0]-p[x+1][0])**2+(p[x+0][1]-p[x+1][1])**2)
                    S.append(d)
                P.append(sum(S))
            
            V=perm[P.index(min(P))]
            V=sorted(V, key=sum)            
            
            for x,v in enumerate(V):
                if x == len(V)-1:
                    break
                d=rround(np.sqrt((V[x+0][0]-V[x+1][0])**2+(V[x+0][1]-V[x+1][1])**2))
                L.append(d)
            runBeam(L)

        return {'FINISHED'}            # Lets Blender know the operator finished successfully.



class Area(bpy.types.Operator):
    """Area"""      # Use this as a tooltip for menu items and buttons.
    bl_idname = "object.area"        # Unique identifier for buttons and menu items to reference.
    bl_label = "Area"         # Display name in the interface.
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

        sheets = model.getCurrentController()

        i=text.CellAddress.Column
        j=text.CellAddress.Row

        cell = sheets.getActiveSheet().getCellByPosition(i, j)
        
        per=[list(p) for p in permutations(getSelVerts())]
        areas=[]
        for p in per:
            areas.append(PolygonArea(p))
        cell.setFormula(str(rround(max(areas))))

        ctx.ServiceManager


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


        sheets = model.getCurrentController()
        i=text.CellAddress.Column
        j=text.CellAddress.Row


        if len(V)==1:
            cell = sheets.getActiveSheet().getCellByPosition(i, j)
            cell.setFormula( str(rround(V[0][0])))
            j+=1            
            cell = sheets.getActiveSheet().getCellByPosition(i, j)
            cell.setFormula(str(rround(V[0][1])))

        elif len(V)>1:
            for v in V:
                cell = sheets.getActiveSheet().getCellByPosition(i, j)
                cell.setFormula( str(rround(v[0])))
                i+=1            
                cell = sheets.getActiveSheet().getCellByPosition(i, j)
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


        sheets = model.getCurrentController()
        i=text.CellAddress.Column
        j=text.CellAddress.Row

        if len(V)==1:
            cell = sheets.getActiveSheet().getCellByPosition(i, j)
            cell.setFormula( str(rround(V[0][0])))
            j+=1            
            cell = sheets.getActiveSheet().getCellByPosition(i, j)
            cell.setFormula(str(rround(V[0][1])))
          
        elif len(V)==2:
            d=str(rround(np.sqrt((V[0][0]-V[1][0])**2+(V[0][1]-V[1][1])**2)))
            cell = sheets.getActiveSheet().getCellByPosition(i, j)
            cell.setFormula( d)

        elif len(V)>2:
            perm=[list(p) for p in permutations(V)]
            P=[]
            for p in perm:
                S=[]
                for x,v in enumerate(p):
                    if x == len(p)-1:
                        break                        
                    d=np.sqrt((p[x+0][0]-p[x+1][0])**2+(p[x+0][1]-p[x+1][1])**2)
                    S.append(d)
                P.append(sum(S))
            
            V=perm[P.index(min(P))]
            V=sorted(V, key=sum)            
            for x,v in enumerate(V):
                if x == len(V)-1:
                    break
                d=str(rround(np.sqrt((V[x+0][0]-V[x+1][0])**2+(V[x+0][1]-V[x+1][1])**2)))
                cell = sheets.getActiveSheet().getCellByPosition(i, j)
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

def add_area_button(self, context):  
        self.layout.operator(  
        Area.bl_idname,  
            text=Area.bl_label,  
            icon='PLAY') 

def add_beam_button(self, context):  
        self.layout.operator(  
        Beam.bl_idname,  
            text=Beam.bl_label,  
            icon='PLAY') 


def register():
    bpy.utils.register_class(ObjectMoveX)
    locationTextEditor.append(add_object_button)
    bpy.utils.register_class(Soffice)
    locationTextEditor.append(add_soffice_button)
    bpy.utils.register_class(VerticeTable)
    locationTextEditor.append(add_vertex_button)
    bpy.utils.register_class(Area)
    locationTextEditor.append(add_area_button)
    bpy.utils.register_class(Beam)
    locationTextEditor.append(add_beam_button)





    
def unregister():
    bpy.utils.unregister_class(ObjectMoveX)
    locationTextEditor.remove(add_object_button)
    bpy.utils.unregister_class(Soffice)
    locationTextEditor.remove(add_soffice_button)
    bpy.utils.unregister_class(VerticeTable)
    locationTextEditor.remove(add_vertex_button)
    bpy.utils.unregister_class(Area)
    locationTextEditor.remove(add_area_button)
    bpy.utils.unregister_class(Beam)
    locationTextEditor.remove(add_beam_button)



if __name__ == "__main__":  
    register()  