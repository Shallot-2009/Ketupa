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
oProject.InsertDesign "HFSS", "CircularWG_Antenna_ADKv1", "DrivenModal", ""
Set oDesign = oProject.SetActiveDesign("CircularWG_Antenna_ADKv1")



dim  solution_freq, WG_Radius, WG_length, Wall_Thickness, units, israd
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



solution_freq = CDbl(args(0))

WG_Radius= CDbl( args(1))
WG_length = CDbl( args(2))
Wall_Thickness = CDbl( args(3))

units = args(4)

israd = args(5) ' 0 is for radiation surface, 1 is for PML


oDesign.ChangeProperty Array("NAME:AllTabs", Array("NAME:LocalVariableTab", Array("NAME:PropServers", "LocalVariables"), Array("NAME:NewProps", _
Array("NAME:--Waveguide Dimensions", "PropType:=", "SeparatorProp", "UserDef:=", true, "Value:=", ""), _
Array("NAME:WG_Radius", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", WG_Radius & units), _
Array("NAME:WG_length", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", WG_length & units), _
Array("NAME:Wall_Thickness", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", Wall_Thickness & units))))


'''''''''''''''''''''''
' draw waveguide
Set oEditor = oDesign.SetActiveEditor("3D Modeler")


''waveguide interior

oEditor.CreateCylinder Array("NAME:CylinderParameters", "CoordinateSystemID:=", -1, "XCenter:=",  _
  "0mm", "YCenter:=", "0mm", "ZCenter:=", "0mm", "Radius:=", "WG_Radius", "Height:=",  _
  "-WG_length", "WhichAxis:=", "Z", "NumSides:=", "0"), Array("NAME:Attributes", "Name:=",  _
  "WG_inner", "Flags:=", "", "Color:=", "(132 132 193)", "Transparency:=",  _
  0.8, "PartCoordinateSystem:=", "Global", "MaterialName:=",  _
  "vacuum", "SolveInside:=", true)


''''wall thickness  
oEditor.CreateCylinder Array("NAME:CylinderParameters", "CoordinateSystemID:=", -1, "XCenter:=",  _
  "0mm", "YCenter:=", "0mm", "ZCenter:=", "0mm", "Radius:=", "WG_Radius+Wall_Thickness", "Height:=",  _
  "-WG_length", "WhichAxis:=", "Z", "NumSides:=", "0"), Array("NAME:Attributes", "Name:=",  _
  "WG_outer", "Flags:=", "", "Color:=", "(132 132 193)", "Transparency:=",  _
  0.8, "PartCoordinateSystem:=", "Global", "MaterialName:=",  _
  "pec", "SolveInside:=", false)

   
oEditor.Subtract Array("NAME:Selections", "Blank Parts:=", "WG_outer", "Tool Parts:=", "WG_inner"), Array("NAME:SubtractParameters", "CoordinateSystemID:=", -1, "KeepOriginals:=", true)

'''''''''''''''''''''''''''''''''''
' create port

''portcap
oEditor.CreateCylinder Array("NAME:CylinderParameters", "CoordinateSystemID:=", -1, "XCenter:=",  _
  "0mm", "YCenter:=", "0mm", "ZCenter:=", "-WG_length", "Radius:=", "WG_Radius+Wall_Thickness", "Height:=",  _
  "-Wall_Thickness", "WhichAxis:=", "Z", "NumSides:=", "0"), Array("NAME:Attributes", "Name:=",  _
  "port_cap", "Flags:=", "", "Color:=", "(132 132 193)", "Transparency:=",  _
  0.8, "PartCoordinateSystem:=", "Global", "MaterialName:=",  _
  "pec", "SolveInside:=", false)   

oEditor.CreateCircle Array("NAME:CircleParameters", "CoordinateSystemID:=", -1, "IsCovered:=",  _
  true, "XCenter:=", "0mm", "YCenter:=", "0mm", "ZCenter:=", "-WG_length", "Radius:=",  _
  "WG_Radius", "WhichAxis:=", "Z", "NumSegments:=", "0"), Array("NAME:Attributes", "Name:=",  _
  "port1", "Flags:=", "", "Color:=", "(132 132 193)", "Transparency:=",  _
  0.8, "PartCoordinateSystem:=", "Global", "MaterialName:=",  _
  "vacuum", "SolveInside:=", true) 

dim end_vector_pointX, end_vector_pointY, end_vector_pointZ
dim start_vector_pointX, start_vector_pointY, start_vector_pointZ

end_vector_pointX = -WG_Radius
end_vector_pointX = end_vector_pointX & units

end_vector_pointY = 0
end_vector_pointY = end_vector_pointY & units

end_vector_pointZ = -WG_length
end_vector_pointZ = end_vector_pointZ & units

start_vector_pointX = WG_Radius
start_vector_pointX = start_vector_pointX & units

start_vector_pointY = 0
start_vector_pointY = start_vector_pointY & units

start_vector_pointZ = -WG_length
start_vector_pointZ = start_vector_pointZ & units

  Set oModule = oDesign.GetModule("BoundarySetup")
  

  oModule.AssignWavePort Array("NAME:p1", "Objects:=", Array("port1"), "NumModes:=",  _
  2, "PolarizeEField:=", true, "DoDeembed:=", false, Array("NAME:Modes", Array("NAME:Mode1", "ModeNum:=",  _
  1, "UseIntLine:=", true, Array("NAME:IntLine", "Start:=", Array(start_vector_pointX, start_vector_pointY,  _
  start_vector_pointZ), "End:=", Array(end_vector_pointX, end_vector_pointY, end_vector_pointZ)), "CharImp:=", "Zpv"), Array("NAME:Mode2", "ModeNum:=",  _
  2, "UseIntLine:=", true, Array("NAME:IntLine", "Start:=", Array(start_vector_pointY, start_vector_pointX, start_vector_pointZ), _
   "End:=", Array(end_vector_pointY, end_vector_pointX, end_vector_pointZ)), "CharImp:=", "Zpv")))
  

Set oModule = oDesign.GetModule("Solutions")
oModule.EditSources "NoIncidentWave", Array("NAME:Names", "p1"), Array("NAME:Modes", 2), Array("NAME:Magnitudes",  _
  "1", "1"), Array("NAME:Phases", "0deg", "90deg"), Array("NAME:Terminated"), Array("NAME:Impedances")


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
start_freq = Round(.8*solution_freq,2)
stop_freq = Round(1.2*solution_freq,2)

 
  
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

extent_x_pos = "WG_Radius"
extent_x_neg = "-WG_Radius"

extent_y_pos = "WG_Radius"
extent_y_neg = "-WG_Radius"

extent_z_pos = "0"
extent_z_neg = "-(WG_length+Wall_Thickness)"

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