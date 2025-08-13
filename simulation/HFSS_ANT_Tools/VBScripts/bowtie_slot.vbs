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
oProject.InsertDesign "HFSS", "Bow_Tie_Slot_Antenna_ADKv1", "DrivenModal", ""
Set oDesign = oProject.SetActiveDesign("Bow_Tie_Slot_Antenna_ADKv1")








dim  solution_freq, Inner_Width, Outer_Width, Arm_Length, Port_Gap_Width, feed_offset, subH, subX, subY, units, israd

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




dim light_speed
light_speed = 299792458

'dim solution_freq
solution_freq = CDbl( args(0))

  dim freq_hz
  freq_hz=solution_freq*1e9

  dim WL 
  WL = light_speed/(freq_hz)




Inner_width = CDbl( args(1))
Outer_Width = CDbl( args(2))
Arm_Length = CDbl( args(3))
Port_Gap_Width = CDbl( args(4))
feed_offset =  CDbl( args(5))

subH = ( args(6))
subX = CDbl( args(7))
subY = CDbl( args(8))


units = args(9)


israd = args(10) ' 0 is for radiation surface, 1 is for PML


if Arm_Length+Port_Gap_Width/2 > subY/2 then
  subY = 2*(Arm_Length+Port_Gap_Width/2)*1.1
end if
if Outer_Width > subX then
  subX = (Outer_Width)*1.1
end if

if feed_offset > Arm_Length then
  feed_offset = Arm_Length*.9
end if


oDesign.ChangeProperty Array("NAME:AllTabs", Array("NAME:LocalVariableTab", Array("NAME:PropServers", "LocalVariables"), Array("NAME:NewProps", _
Array("NAME:--Antenna Dimensions", "PropType:=", "SeparatorProp", "UserDef:=", true, "Value:=", ""), _
Array("NAME:Inner_width", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", Inner_width & units), _
Array("NAME:Outer_width", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", Outer_width & units), _
Array("NAME:Arm_Length", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", Arm_Length & units), _
Array("NAME:Port_Gap_Width", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", Port_Gap_Width & units), _
Array("NAME:Feed_Offset", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", feed_offset & units), _
Array("NAME:--Substrate Dimensions", "PropType:=", "SeparatorProp", "UserDef:=", true, "Value:=", ""), _
Array("NAME:subH", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", subH ), _
Array("NAME:subX", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", subX & units), _
Array("NAME:subY", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", subY & units))))


 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'create substrate material

Material_Name = args(11)
Material_Name = Material_Name & "_ADK"
Permittivity = args(12)
TandD = args(13)

Set oDefinitionManager = oProject.GetDefinitionManager()
oDefinitionManager.AddMaterial Array("NAME:" & Material_Name, "CoordinateSystemType:=",  _
  "Cartesian", Array("NAME:AttachedData"), Array("NAME:ModifierData"), "permittivity:=", Permittivity , "dielectric_loss_tangent:=",TandD)



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Draws Substrate

Set oEditor = oDesign.SetActiveEditor("3D Modeler")
oEditor.CreateBox Array("NAME:BoxParameters", "CoordinateSystemID:=", -1, "XPosition:=",  _
  "-subX/2", "YPosition:=", "-subY/2" , "ZPosition:=", "0mm", "XSize:=", "subX", "YSize:=",  _
  "subY", "ZSize:=", "-subH"), Array("NAME:Attributes", "Name:=", "sub", "Flags:=",  _
  "", "Color:=", "(132 132 193)", "Transparency:=", 0.8, "PartCoordinateSystem:=",  _
  "Global", "MaterialName:=", Material_Name, "SolveInside:=", true)
  
  


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Draws antenna

Set oEditor = oDesign.SetActiveEditor("3D Modeler")

oEditor.CreatePolyline Array("NAME:PolylineParameters", "CoordinateSystemID:=", -1, "IsPolylineCovered:=",  _
  true, "IsPolylineClosed:=", false, Array("NAME:PolylinePoints", Array("NAME:PLPoint", "X:=",  _
  "-Inner_Width/2", "Y:=", "Port_Gap_Width/2", "Z:=", "0mm"), Array("NAME:PLPoint", "X:=", "Inner_Width/2", "Y:=",  _
  "Port_Gap_Width/2", "Z:=", "0mm")), Array("NAME:PolylineSegments", Array("NAME:PLSegment", "SegmentType:=",  _
  "Line", "StartIndex:=", 0, "NoOfPoints:=", 2))), Array("NAME:Attributes", "Name:=",  _
  "Inner", "Flags:=", "", "Color:=", "(255 128 65)", "Transparency:=",  _
  0.3, "PartCoordinateSystem:=", "Global", "MaterialName:=",  _
  "vacuum", "SolveInside:=", true)
  
oEditor.CreatePolyline Array("NAME:PolylineParameters", "CoordinateSystemID:=", -1, "IsPolylineCovered:=",  _
  true, "IsPolylineClosed:=", false, Array("NAME:PolylinePoints", Array("NAME:PLPoint", "X:=",  _
  "-Outer_Width/2", "Y:=", "Port_Gap_Width/2+Arm_Length", "Z:=", "0mm"), Array("NAME:PLPoint", "X:=", "Outer_Width/2", "Y:=",  _
  "Port_Gap_Width/2+Arm_Length", "Z:=", "0mm")), Array("NAME:PolylineSegments", Array("NAME:PLSegment", "SegmentType:=",  _
  "Line", "StartIndex:=", 0, "NoOfPoints:=", 2))), Array("NAME:Attributes", "Name:=",  _
  "Arm", "Flags:=", "", "Color:=", "(255 128 65)", "Transparency:=",  _
  0.3, "PartCoordinateSystem:=", "Global", "MaterialName:=",  _
  "vacuum", "SolveInside:=", true)
  
  
oEditor.Connect Array("NAME:Selections", "Selections:=", "Arm,Inner")


oEditor.DuplicateAroundAxis Array("NAME:Selections", "Selections:=", "Arm", "NewPartsModelFlag:=",  _
  "Model"), Array("NAME:DuplicateAroundAxisParameters", "CoordinateSystemID:=", -1, "CreateNewObjects:=", _
  true, "WhichAxis:=", "Z", "AngleStr:=", "180deg", "NumClones:=", "2"), Array("NAME:Options", "DuplicateBoundaries:=", true)








''''''''
'port setup

oEditor.CreateRectangle Array("NAME:RectangleParameters", "CoordinateSystemID:=",  _
  -1, "IsCovered:=", true, "XStart:=", "-Inner_Width/2", "YStart:=", "-Port_Gap_Width/2", "ZStart:=",  _
  "0mm", "Width:=", "Inner_Width", "Height:=", "Port_Gap_Width", "WhichAxis:=", "Z"), Array("NAME:Attributes", "Name:=",  _
  "port1", "Flags:=", "", "Color:=", "(255 128 65)", "Transparency:=",  _
  0.8, "PartCoordinateSystem:=", "Global", "MaterialName:=",  _
  "vacuum", "SolveInside:=", true)

oEditor.CreateRectangle Array("NAME:RectangleParameters", "CoordinateSystemID:=",  _
  -1, "IsCovered:=", true, "XStart:=", "-SubX/2", "YStart:=", "-SubY/2", "ZStart:=",  _
  "0mm", "Width:=", "SubX", "Height:=", "SubY", "WhichAxis:=", "Z"), Array("NAME:Attributes", "Name:=",  _
  "slot", "Flags:=", "", "Color:=", "(255 128 65)", "Transparency:=",  _
  0.8, "PartCoordinateSystem:=", "Global", "MaterialName:=",  _
  "vacuum", "SolveInside:=", true)

oEditor.Subtract Array("NAME:Selections", "Blank Parts:=", "slot", "Tool Parts:=",  _
  "Arm,Arm_1"), Array("NAME:SubtractParameters", "KeepOriginals:=", false)
oEditor.Subtract Array("NAME:Selections", "Blank Parts:=", "slot", "Tool Parts:=",  _
  "port1"), Array("NAME:SubtractParameters", "KeepOriginals:=", true)

oEditor.CreateRectangle Array("NAME:RectangleParameters", "CoordinateSystemID:=",  _
  -1, "IsCovered:=", true, "XStart:=", "-Inner_Width/2", "YStart:=", "-Port_Gap_Width/2", "ZStart:=",  _
  "0mm", "Width:=", "-SubX/2+Inner_Width/2", "Height:=", "Port_Gap_Width", "WhichAxis:=", "Z"), Array("NAME:Attributes", "Name:=",  _
  "feed_offset_1", "Flags:=", "", "Color:=", "(255 128 65)", "Transparency:=",  _
  0.8, "PartCoordinateSystem:=", "Global", "MaterialName:=",  _
  "vacuum", "SolveInside:=", true)
  
oEditor.CreateRectangle Array("NAME:RectangleParameters", "CoordinateSystemID:=",  _
  -1, "IsCovered:=", true, "XStart:=", "Inner_Width/2", "YStart:=", "-Port_Gap_Width/2", "ZStart:=",  _
  "0mm", "Width:=", "SubX/2-Inner_Width/2", "Height:=", "Port_Gap_Width", "WhichAxis:=", "Z"), Array("NAME:Attributes", "Name:=",  _
  "feed_offset_2", "Flags:=", "", "Color:=", "(255 128 65)", "Transparency:=",  _
  0.8, "PartCoordinateSystem:=", "Global", "MaterialName:=",  _
  "vacuum", "SolveInside:=", true)  

dim end_vector_pointX, end_vector_pointY
dim start_vector_pointX, start_vector_pointY

end_vector_pointX = Inner_width/2
end_vector_pointX = end_vector_pointX & units
end_vector_pointY = 0
end_vector_pointY = end_vector_pointY & units


start_vector_pointX = -Inner_width/2
start_vector_pointX = start_vector_pointX & units
start_vector_pointY = 0
start_vector_pointY = start_vector_pointY & units


oDesign.SetSolutionType "DrivenModal"
Set oModule = oDesign.GetModule("BoundarySetup")
oModule.AssignLumpedPort Array("NAME:p1", "Objects:=", Array("port1"), Array("NAME:Modes", Array("NAME:Mode1", "ModeNum:=",  _
  1, "UseIntLine:=", true, Array("NAME:IntLine", "Start:=", Array(start_vector_pointX, start_vector_pointY,  _
  "0mm"), "End:=", Array(end_vector_pointX, end_vector_pointY, "0mm")), "CharImp:=", "Zpi", "RenormImp:=",  _
  "50ohm")), "FullResistance:=", "50ohm", "FullReactance:=", "0ohm")

Set oEditor = oDesign.SetActiveEditor("3D Modeler")
oEditor.Move Array("NAME:Selections", "Selections:=",  _
  "feed_offset_1,feed_offset_2,port1", "NewPartsModelFlag:=", "Model"), Array("NAME:TranslateParameters", "TranslateVectorX:=",  _
  "0cm", "TranslateVectorY:=", "Feed_Offset", "TranslateVectorZ:=", "0cm")

oEditor.Unite Array("NAME:Selections", "Selections:=",  _
  "slot,feed_offset_2,feed_offset_1"), Array("NAME:UniteParameters", "KeepOriginals:=",  _
  false)
  
    Set oModule = oDesign.GetModule("BoundarySetup")  
oModule.AssignPerfectE Array("NAME:Slot", "Objects:=", Array("slot"), "InfGroundPlane:=", false) 

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
start_freq = Round(.5*solution_freq,2)
stop_freq = Round(1.5*solution_freq,2)

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

extent_z_pos = "0"
extent_z_neg = "-subH"

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
  
  
