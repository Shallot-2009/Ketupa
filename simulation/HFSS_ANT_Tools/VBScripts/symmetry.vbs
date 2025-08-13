' ----------------------------------------------
' Script Recorded by Ansoft HFSS Version 11.1.3
' 2:47 PM  Jun 08, 2009
' ----------------------------------------------
Dim oAnsoftApp
Dim oDesktop
Dim oProject
Dim oDesign
Dim oEditor
Dim oModule
Set oAnsoftApp = CreateObject("AnsoftHfss.HfssScriptInterface")
Set oDesktop = oAnsoftApp.GetAppDesktop()
oDesktop.RestoreWindow

Set oProject = oDesktop.GetActiveProject()
Set oDesign = oProject.GetActiveDesign()

Set oEditor = oDesign.SetActiveEditor("3D Modeler")
objectnames = oEditor.GetMatchedObjectName("*")
totalobjects = oEditor.GetNumObjects



string_of_objects_split = objectnames(0)



for i=1 to totalobjects-1
  string_of_objects_split = string_of_objects_split & "," & objectnames(i)

next

split_plane = "ZX"  'YZ, ZX, XY
oEditor.Split Array("NAME:Selections", "Selections:=", string_of_objects_split, "NewPartsModelFlag:=",  _
  "Model"), Array("NAME:SplitToParameters", "CoordinateSystemID:=", -1, "SplitPlane:=",  _
  split_plane, "WhichSide:=", "PositiveOnly", "SplitCrossingObjectsOnly:=", false)

objectnames = oEditor.GetMatchedObjectName("*")
totalobjects = oEditor.GetNumObjects
  
  
'can't use any 2D objects with the section command so this will check for 2D and 3D objects
dim objects_section()
redim objects_section(totalobjects)
string_of_objects_section = ""
q=0
for i=0 to totalobjects-1
  all_props = oEditor.GetProperties("Geometry3DAttributeTab",objectnames(i))
  number_of_props = ubound(all_props)

  if number_of_props = 7 then  '7= 3D and 5 = 2D
    objects_section(q) = objectnames(i)
     string_of_objects_section = string_of_objects_section & "," & objectnames(i)
    q=q+1
  end if 
next
  
string_of_objects_section = Right(string_of_objects_section,len(string_of_objects_section)-1) 'removes leading comma from string


oEditor.Section Array("NAME:Selections", "Selections:=", string_of_objects_section , "NewPartsModelFlag:=", "Model"),_
 Array("NAME:SectionToParameters", "CoordinateSystemID:=",-1, "CreateNewObjects:=", false, "SectionPlane:=", split_plane) 


objectnames_for_sym = oEditor.GetMatchedObjectName("*_Section*")
size_of_objectnames_for_sym = Ubound(objectnames_for_sym)




Set oModule = oDesign.GetModule("BoundarySetup")
oModule.ChangeImpedanceMult 0.5
oModule.AssignSymmetry Array("NAME:Sym1", "Objects:=", objectnames_for_sym, "IsPerfectE:=", false)








  