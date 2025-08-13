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
oProject.InsertDesign "HFSS", "Discone_Antenna_ADKv1", "DrivenModal", ""
Set oDesign = oProject.SetActiveDesign("Discone_Antenna_ADKv1")

dim  solution_freq, Inner_Radius, Outer_Radius, Cone_Height, Ground_Disk_Radius, Port_Gap, Port_Width, units, israd

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

Inner_Radius = CDbl( args(1))
Outer_Radius = CDbl( args(2))
Cone_Height = CDbl( args(3))
Ground_Disk_Radius =CDbl( args(4))
Port_Gap = CDbl( args(5))
Port_Width = CDbl( args(6))

if Port_Width > Inner_Radius*2 then
  Port_Width = Inner_Radius*2
end if

units = args(7)

israd = args(8) ' 0 is for radiation surface, 1 is for PML

oDesign.ChangeProperty Array("NAME:AllTabs", Array("NAME:LocalVariableTab", Array("NAME:PropServers", "LocalVariables"), Array("NAME:NewProps", _
Array("NAME:--Antenna Dimensions", "PropType:=", "SeparatorProp", "UserDef:=", true, "Value:=", ""), _
Array("NAME:Inner_Radius", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", Inner_Radius & units), _
Array("NAME:Outer_Radius", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", Outer_Radius & units), _
Array("NAME:Cone_Height", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", Cone_Height & units), _
Array("NAME:Ground_Disk_Radius", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", Ground_Disk_Radius & units), _
Array("NAME:Port_Gap", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", Port_Gap & units), _
Array("NAME:Port_Width", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", Port_Width & units))))




'''''''''''''''
' Draw Ground Disk
Set oEditor = oDesign.SetActiveEditor("3D Modeler")

oEditor.CreateCircle Array("NAME:CircleParameters", "CoordinateSystemID:=", -1, "IsCovered:=",  _
  true, "XCenter:=", "0mm", "YCenter:=", "0mm", "ZCenter:=", "0mm", "Radius:=",  _
  "Ground_Disk_Radius", "WhichAxis:=", "Z", "NumSegments:=", "0"), Array("NAME:Attributes", "Name:=",  _
  "Ground_Disk", "Flags:=", "", "Color:=", "(255 128 65)", "Transparency:=",  _
  0.8, "PartCoordinateSystem:=", "Global", "MaterialName:=",  _
  "vacuum", "SolveInside:=", true)  
Set oModule = oDesign.GetModule("BoundarySetup")  
oModule.AssignPerfectE Array("NAME:Ground", "Objects:=", Array("Ground_Disk"), "InfGroundPlane:=", false)  


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Draws antenna

Set oEditor = oDesign.SetActiveEditor("3D Modeler")

oEditor.CreatePolyline Array("NAME:PolylineParameters", "CoordinateSystemID:=", -1, "IsPolylineCovered:=",  _
  true, "IsPolylineClosed:=", false, Array("NAME:PolylinePoints", Array("NAME:PLPoint", "X:=",  _
  "Inner_Radius", "Y:=", "0mm", "Z:=", "Port_Gap"), Array("NAME:PLPoint", "X:=", "Outer_Radius", "Y:=",  _
  "0mm", "Z:=", "Port_Gap+Cone_Height")), Array("NAME:PolylineSegments", Array("NAME:PLSegment", "SegmentType:=",  _
  "Line", "StartIndex:=", 0, "NoOfPoints:=", 2))), Array("NAME:Attributes", "Name:=",  _
  "Cone", "Flags:=", "", "Color:=", "(255 128 65)", "Transparency:=",  _
  0.8, "PartCoordinateSystem:=", "Global", "MaterialName:=",  _
  "vacuum", "SolveInside:=", true)
  

oEditor.SweepAroundAxis Array("NAME:Selections", "Selections:=", "Cone", "NewPartsModelFlag:=",  _
  "Model"), Array("NAME:AxisSweepParameters", "CoordinateSystemID:=", -1, "DraftAngle:=",  _
  "0deg", "DraftType:=", "Round", "CheckFaceFaceIntersection:=", false, "SweepAxis:=",  _
  "Z", "SweepAngle:=", "360deg", "NumOfSegments:=", "0")
  
oEditor.CreateCircle Array("NAME:CircleParameters", "CoordinateSystemID:=", -1, "IsCovered:=",  _
  true, "XCenter:=", "0mm", "YCenter:=", "0mm", "ZCenter:=", "Port_Gap", "Radius:=",  _
  "Inner_Radius", "WhichAxis:=", "Z", "NumSegments:=", "0"), Array("NAME:Attributes", "Name:=",  _
  "Cone_End", "Flags:=", "", "Color:=", "(255 128 65)", "Transparency:=",  _
  0.8, "PartCoordinateSystem:=", "Global", "MaterialName:=",  _
  "vacuum", "SolveInside:=", true)  
  
oEditor.Unite Array("NAME:Selections", "Selections:=", "Cone,Cone_End"), Array("NAME:UniteParameters", "CoordinateSystemID:=",  _
  -1, "KeepOriginals:=", false)

Set oModule = oDesign.GetModule("BoundarySetup")  
oModule.AssignPerfectE Array("NAME:Cone", "Objects:=", Array("Cone"), "InfGroundPlane:=", false)





''''''''
'port setup

Set oEditor = oDesign.SetActiveEditor("3D Modeler")

oEditor.CreateRectangle Array("NAME:RectangleParameters", "CoordinateSystemID:=",  _
  -1, "IsCovered:=", true, "XStart:=", "0", "YStart:=", "-Port_Width/2", "ZStart:=",  _
  "0mm", "Width:=", "Port_Width", "Height:=", "Port_Gap", "WhichAxis:=", "X"), Array("NAME:Attributes", "Name:=",  _
  "port1", "Flags:=", "", "Color:=", "(255 128 65)", "Transparency:=",  _
  0.8, "PartCoordinateSystem:=", "Global", "MaterialName:=",  _
  "vacuum", "SolveInside:=", true)

dim end_vector_pointX, end_vector_pointY
dim start_vector_pointX, start_vector_pointY

end_vector_pointX = 0
end_vector_pointX = end_vector_pointX & units
end_vector_pointY = 0
end_vector_pointY = end_vector_pointY & units
end_vector_pointZ = Port_Gap
end_vector_pointZ = end_vector_pointZ & units

start_vector_pointX = 0
start_vector_pointX = start_vector_pointX & units
start_vector_pointY = 0
start_vector_pointY = start_vector_pointY & units
start_vector_pointZ = 0
start_vector_pointZ = start_vector_pointZ & units



oDesign.SetSolutionType "DrivenModal"
Set oModule = oDesign.GetModule("BoundarySetup")
oModule.AssignLumpedPort Array("NAME:p1", "Objects:=", Array("port1"), Array("NAME:Modes", Array("NAME:Mode1", "ModeNum:=",  _
  1, "UseIntLine:=", true, Array("NAME:IntLine", "Start:=", Array(start_vector_pointX, start_vector_pointY,  _
  start_vector_pointZ), "End:=", Array(end_vector_pointX, end_vector_pointY, end_vector_pointZ)), "CharImp:=", "Zpi", "RenormImp:=",  _
  "50ohm")), "FullResistance:=", "50ohm", "FullReactance:=", "0ohm")


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
  "LinearCount", "StartValue:=", start_freq&"GHz", "StopValue:=", stop_freq&"GHz", "Count:=", 101, "Type:=",  _
  "Fast", "SaveFields:=", true, "GenerateFieldsForAllFreqs:=", false, "ExtrapToDC:=",  _
  false)  
    
  
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''  
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'get userlib directory where radition box script is sitting  
' and create radiation box
dim libdir
libdir = oDesktop.GetLibraryDirectory
dim fullpath
fullpath = libdir & "\userlib\AntennaDesignKit\VBScripts\rad_creation.vbs" 


dim extent_x_pos, extent_x_neg, extent_y_pos, extent_y_neg, extent_z_pos, extent_z_neg



if CLng(Ground_Disk_Radius) > CLng(Outer_Radius) then
  extent_x_pos = "Ground_Disk_Radius"
  extent_x_neg = "-Ground_Disk_Radius"

  extent_y_pos = "Ground_Disk_Radius"
  extent_y_neg = "-Ground_Disk_Radius"

else
extent_x_pos = "Outer_Radius"
extent_x_neg = "-Outer_Radius"

extent_y_pos = "Outer_Radius"
extent_y_neg = "-Outer_Radius"
end if



extent_z_pos = "(Cone_Height+Port_Gap)"
extent_z_neg = "0"

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