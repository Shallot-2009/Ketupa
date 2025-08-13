' ----------------------------------------------
'  Ansoft HFSS Version 11.1.3
' 
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
oDesktop.NewProject
Set oProject = oDesktop.GetActiveProject
oProject.InsertDesign "HFSS", "Patch_Antenna_ADKv1", "DrivenModal", ""
Set oDesign = oProject.SetActiveDesign("Patch_Antenna_ADKv1")


dim circle_or_rect, solution_freq, patchX, patchY, subH, subX, subY, feedX, feedY, coax_inner_rad, coax_outer_rad, feed_length, units, israd

Locallang=getlocale()

Setlocale(1033)

' get arguments passed into script
on error resume next
dim args
Set args = AnsoftScript.arguments
if(IsEmpty(args)) then 
Set args = WSH.arguments
End if
on error goto 0
'At this point, args has the arguments no matter if you are running 
'under windows script host or Ansoft script hos


circle_or_rect = CInt( args(0))  '1 is rectangular, 0 is circular

dim light_speed
light_speed = 299792458


solution_freq = CDbl( args(1))

  dim freq_hz
  freq_hz=solution_freq*1e9

  dim WL 
  WL = light_speed/(freq_hz)



  


patchX = CDbl( args(2))
patchY = CDbl( args(3))
subH = args(4)
subX = CDbl( args(5))
subY = CDbl( args(6))
feedX = CDbl( args(7))
feedY = CDbl( args(8))


coax_inner_rad = CDbl( args(9))
coax_outer_rad = CDbl( args(10))
feed_length = CDbl( args(11))

units = args(12)

israd = args(13) ' 0 is for radiation surface, 1 is for PML

if patchX > subX then
  subX = patchX*1.1
end if
if patchY > subY then
  subY = patchY*1.1
end if


if patchX > subX then
  subX = patchX*1.1
end if
if patchY > subY then
  subY = patchY*1.1
end if





oDesign.ChangeProperty Array("NAME:AllTabs", Array("NAME:LocalVariableTab", Array("NAME:PropServers", "LocalVariables"), Array("NAME:NewProps", _
Array("NAME:--Patch Dimensions", "PropType:=", "SeparatorProp", "UserDef:=", true, "Value:=", ""), _
Array("NAME:patchX", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", patchX & units), _
Array("NAME:patchY", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", patchY & units), _
Array("NAME:--Substrate Dimensions", "PropType:=", "SeparatorProp", "UserDef:=", true, "Value:=", ""), _
Array("NAME:subH", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", subH ), _
Array("NAME:subX", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", subX & units), _
Array("NAME:subY", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", subY & units), _
Array("NAME:--Feed", "PropType:=", "SeparatorProp", "UserDef:=", true, "Value:=", ""), _
Array("NAME:feedX", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", feedX & units), _
Array("NAME:feedY", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", feedY & units), _
Array("NAME:coax_inner_rad", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", coax_inner_rad & units), _
Array("NAME:coax_outer_rad", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", coax_outer_rad & units), _
Array("NAME:feed_length", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", feed_length & units))))

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'create substrate material

Material_Name = args(14)
Material_Name = Material_Name & "_ADK"
Permittivity = args(15)
TandD = args(16)

Set oDefinitionManager = oProject.GetDefinitionManager()
oDefinitionManager.AddMaterial Array("NAME:" & Material_Name, "CoordinateSystemType:=",  _
  "Cartesian", Array("NAME:AttachedData"), Array("NAME:ModifierData"), "permittivity:=", Permittivity , "dielectric_loss_tangent:=",TandD)



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Draws Substrate and bottom metalization

Set oEditor = oDesign.SetActiveEditor("3D Modeler")
oEditor.CreateBox Array("NAME:BoxParameters", "CoordinateSystemID:=", -1, "XPosition:=",  _
  "-subX/2", "YPosition:=", "-subY/2" , "ZPosition:=", "0mm", "XSize:=", "subX", "YSize:=",  _
  "subY", "ZSize:=", "subH"), Array("NAME:Attributes", "Name:=", "sub", "Flags:=",  _
  "", "Color:=", "(132 132 193)", "Transparency:=", 0.8, "PartCoordinateSystem:=",  _
  "Global", "MaterialName:=", Material_Name, "SolveInside:=", true)
  
  
oEditor.CreateRectangle Array("NAME:RectangleParameters", "CoordinateSystemID:=",  _
  -1, "IsCovered:=", true, "XStart:=", "-subX/2", "YStart:=", "-subY/2", "ZStart:=",  _
  "0mm", "Width:=", "subX", "Height:=", "subY", "WhichAxis:=", "Z"), Array("NAME:Attributes", "Name:=",  _
  "Ground", "Flags:=", "", "Color:=", "(255 128 65)", "Transparency:=",  _
  0.8, "PartCoordinateSystem:=", "Global", "MaterialName:=",  _
  "pec", "SolveInside:=", false)

Set oModule = oDesign.GetModule("BoundarySetup")  

oModule.AssignPerfectE Array("NAME:Ground", "Objects:=", Array("Ground"), "InfGroundPlane:=", false)


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Draws patch

Set oEditor = oDesign.SetActiveEditor("3D Modeler")

if circle_or_rect = 1 then
  oEditor.CreateRectangle Array("NAME:RectangleParameters", "CoordinateSystemID:=",  _
  -1, "IsCovered:=", true, "XStart:=", "-patchX/2", "YStart:=", "-patchY/2", "ZStart:=",  _
  "subH", "Width:=", "patchX", "Height:=", "patchY", "WhichAxis:=", "Z"), Array("NAME:Attributes", "Name:=",  _
  "patch", "Flags:=", "", "Color:=", "(255 128 65)", "Transparency:=",  _
  0.3, "PartCoordinateSystem:=", "Global", "MaterialName:=",  _
  "pec", "SolveInside:=", false)
else
  
  oEditor.CreateEllipse Array("NAME:EllipseParameters", "CoordinateSystemID:=", -1, "IsCovered:=",  _
  true, "XCenter:=", "0mm", "YCenter:=", "0mm", "ZCenter:=", "subH", "MajRadius:=",  _
  "patchX/2", "Ratio:=", "patchY/patchX", "WhichAxis:=", "Z", "NumSegments:=", "0"), Array("NAME:Attributes", "Name:=",  _
  "patch", "Flags:=", "", "Color:=", "(255 128 65)", "Transparency:=",  _
  0.3, "PartCoordinateSystem:=", "Global", "MaterialName:=",  _
  "pec", "SolveInside:=", false)
  
 
end if 
  
Set oModule = oDesign.GetModule("BoundarySetup")  

oModule.AssignPerfectE Array("NAME:Patch", "Objects:=", Array("patch"), "InfGroundPlane:=", false)


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Draws feed

Set oEditor = oDesign.SetActiveEditor("3D Modeler")

'feed cutout
oEditor.CreateCircle Array("NAME:CircleParameters", "CoordinateSystemID:=", -1, "IsCovered:=",  _
  true, "XCenter:=", "feedX", "YCenter:=", "feedY", "ZCenter:=", "0mm", "Radius:=",  _
  "coax_outer_rad", "WhichAxis:=", "Z", "NumSegments:=", "0"), Array("NAME:Attributes", "Name:=",  _
  "feed_cutout", "Flags:=", "", "Color:=", "(255 128 65)", "Transparency:=",  _
  0.8, "PartCoordinateSystem:=", "Global", "MaterialName:=",  _
  "vacuum", "SolveInside:=", true)

oEditor.Subtract Array("NAME:Selections", "Blank Parts:=", "Ground", "Tool Parts:=",  _
  "feed_cutout"), Array("NAME:SubtractParameters", "CoordinateSystemID:=", -1, "KeepOriginals:=",  _
  false)

'feed pin
oEditor.CreateCylinder Array("NAME:CylinderParameters", "CoordinateSystemID:=", -1, "XCenter:=",  _
  "feedX", "YCenter:=", "feedY", "ZCenter:=", "0mm", "Radius:=", "coax_inner_rad", "Height:=",  _
  "subH", "WhichAxis:=", "Z", "NumSides:=", "0"), Array("NAME:Attributes", "Name:=",  _
  "feed_pin", "Flags:=", "", "Color:=", "(255 128 65)", "Transparency:=",  _
  0.8, "PartCoordinateSystem:=", "Global", "MaterialName:=",  _
  "pec", "SolveInside:=", false)
  
'feed coax
oEditor.CreateCylinder Array("NAME:CylinderParameters", "CoordinateSystemID:=", -1, "XCenter:=",  _
  "feedX", "YCenter:=", "feedY", "ZCenter:=", "0mm", "Radius:=", "coax_inner_rad", "Height:=",  _
  "-feed_length", "WhichAxis:=", "Z", "NumSides:=", "0"), Array("NAME:Attributes", "Name:=",  _
  "coax_pin", "Flags:=", "", "Color:=", "(132 132 193)", "Transparency:=",  _
  0.8, "PartCoordinateSystem:=", "Global", "MaterialName:=",  _
  "pec", "SolveInside:=", false) 
  
oEditor.CreateCylinder Array("NAME:CylinderParameters", "CoordinateSystemID:=", -1, "XCenter:=",  _
  "feedX", "YCenter:=", "feedY", "ZCenter:=", "0mm", "Radius:=", "coax_outer_rad", "Height:=",  _
  "-feed_length", "WhichAxis:=", "Z", "NumSides:=", "0"), Array("NAME:Attributes", "Name:=",  _
  "coax", "Flags:=", "", "Color:=", "(132 132 193)", "Transparency:=",  _
  0.8, "PartCoordinateSystem:=", "Global", "MaterialName:=",  _
  "Teflon (tm)", "SolveInside:=", true)

'Dim faceid, faceid2, half_feed_length, outface_posX,outface_posY
'half_feed_length = feed_length/2
'outface_posX = feedX+coax_inner_rad
'outface_posY = feedY



faceid = oEditor.GetFaceByPosition(Array("NAME:FaceParameters","BodyName:=", "coax", "XPosition:=", "feedX+coax_outer_rad", "YPosition:=","feedY", "ZPosition:=", "-feed_length/2"))



Set oModule = oDesign.GetModule("BoundarySetup")
oModule.AssignPerfectE Array("NAME:coax_outer", "Faces:=", Array(faceid), "InfGroundPlane:=",  _
  false)

'port
Set oEditor = oDesign.SetActiveEditor("3D Modeler")
oEditor.CreateCircle Array("NAME:CircleParameters", "CoordinateSystemID:=", -1, "IsCovered:=",  _
  true, "XCenter:=", "feedX", "YCenter:=", "feedY", "ZCenter:=", "-feed_length", "Radius:=",  _
  "coax_outer_rad", "WhichAxis:=", "Z", "NumSegments:=", "0"), Array("NAME:Attributes", "Name:=",  _
  "port1", "Flags:=", "", "Color:=", "(132 132 193)", "Transparency:=",  _
  0.8, "PartCoordinateSystem:=", "Global", "MaterialName:=",  _
  "vacuum", "SolveInside:=", true)
  
oEditor.CreateCylinder Array("NAME:CylinderParameters", "CoordinateSystemID:=", -1, "XCenter:=",  _
  "feedX", "YCenter:=", "feedY", "ZCenter:=", "-feed_length", "Radius:=", "coax_outer_rad", "Height:=",  _
  "-subH/10", "WhichAxis:=", "Z", "NumSides:=", "0"), Array("NAME:Attributes", "Name:=",  _
  "port_cap", "Flags:=", "", "Color:=", "(132 132 193)", "Transparency:=",  _
  0.8, "PartCoordinateSystem:=", "Global", "MaterialName:=",  _
  "pec", "SolveInside:=", false)   

oDesign.SetSolutionType "DrivenTerminal"

faceid = oEditor.GetFaceByPosition(Array("NAME:FaceParameters","BodyName:=", "port1", "XPosition:=", "feedX", "YPosition:=","feedY", "ZPosition:=", "-feed_length"))


Set oModule = oDesign.GetModule("BoundarySetup")
oModule.AutoIdentifyPorts Array("NAME:Faces", faceid), true, Array("NAME:ReferenceConductors",  _
  "port_cap")
  
  
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Solution Setup 
  



Set oModule = oDesign.GetModule("AnalysisSetup")

oModule.InsertSetup "HfssDriven", Array("NAME:Setup1", "Frequency:=", solution_freq&"GHZ", "PortsOnly:=",  _
  false, "MaxDeltaS:=", 0.02, "UseMatrixConv:=", false, "MaximumPasses:=", 15, "MinimumPasses:=",  _
  1, "MinimumConvergedPasses:=", 1, "PercentRefinement:=", 30, "BasisOrder:=", 1, "UseIterativeSolver:=",  _
  true, "DoLambdaRefine:=", true, "DoMaterialLambda:=", true, "SetLambdaTarget:=",  _
  false, "Target:=", 0.3333, "UseConvOutputVariable:=", false, "IsEnabled:=",  _
  true, "ExternalMesh:=", false, "UseMaxTetIncrease:=", false, "MaxTetIncrease:=",  _
  100000, "PortAccuracy:=", 2, "UseABCOnPort:=", false, "SetPortMinMaxTri:=",  _
  false)
  
dim start_freq
dim stop_freq
start_freq = .5*solution_freq
stop_freq = 1.5*solution_freq

 oModule.InsertFrequencySweep "Setup1", Array("NAME:Sweep1", "IsEnabled:=", true, "SetupType:=",  _
  "LinearCount", "StartValue:=", start_freq&"GHz", "StopValue:=", stop_freq&"GHz", "Count:=", 200, "Type:=",  _
  "Interpolating", "SaveFields:=", false, "InterpTolerance:=", 0.5, "InterpMaxSolns:=",  _
  50, "InterpMinSolns:=", 0, "InterpMinSubranges:=", 1, "ExtrapToDC:=", false, "InterpUseS:=",  _
  true, "InterpUseT:=", false, "InterpUsePortImped:=", false, "InterpUsePropConst:=",  _
  true, "UseFullBasis:=", true)   
  
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''  
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'get userlib directory where radition box script is sitting  
' and create radiation box
dim libdir
libdir = oDesktop.GetLibraryDirectory
dim fullpath
fullpath = libdir & "\userlib\AntennaDesignKit\VBScripts\rad_creation.vbs" 


dim extent_x_pos, extent_x_neg, extent_y_pos, extent_y_neg, extent_z_pos, extent_z_neg

extent_x_pos = "subX/2"
extent_x_neg = "-subX/2"

extent_y_pos = "subY/2"
extent_y_neg = "-subY/2"

extent_z_pos = "subH"
extent_z_neg = "-feed_length"

if israd = "ABC" then
 israd = "Rad"

end if


dim mycommand
' full command which invokes wscript to run desired VBScript and passes
mycommand = "wscript.exe " & """" & "./VBScripts/rad_creation.vbs" & """" &  " " & extent_x_pos & " " & extent_x_neg & " " & extent_y_pos & " " & extent_y_neg & " " & extent_z_pos & " " & extent_z_neg & " " & units & " " & israd& " " & 0
'msgbox(mycommand)
' run the desired VBScript
Set WSHShell = CreateObject("WScript.Shell")
WSHShell.Run mycommand, , True
Setlocale(locallang) 
  