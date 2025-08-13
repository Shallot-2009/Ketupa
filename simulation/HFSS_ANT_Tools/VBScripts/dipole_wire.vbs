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
oProject.InsertDesign "HFSS", "Dipole_Antenna_ADKv1", "DrivenModal", ""
Set oDesign = oProject.SetActiveDesign("Dipole_Antenna_ADKv1")

dim center_freq, wire_rad, port_gap, dipole_length, units, israd

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


center_freq = CDbl( args(0))

  dim freq_hz
  freq_hz=center_freq*1e9

  dim WL 
  WL = light_speed/(freq_hz)


wire_rad = CDbl( args(1))

port_gap = CDbl( args(2))

dipole_length = CDbl( args(3))

units = args(4)


israd = args(5) ' 0 is for radiation surface, 1 is for PML









oDesign.ChangeProperty Array("NAME:AllTabs", Array("NAME:LocalVariableTab", Array("NAME:PropServers", "LocalVariables"), Array("NAME:NewProps", _
Array("NAME:--Dipole Dimensions", "PropType:=", "SeparatorProp", "UserDef:=", true, "Value:=", ""), _
Array("NAME:wire_rad", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", wire_rad & units), _
Array("NAME:dipole_length", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", dipole_length & units), _
Array("NAME:port_gap", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", port_gap & units))))
 
  
  
Set oEditor = oDesign.SetActiveEditor("3D Modeler")

oEditor.CreateCylinder Array("NAME:CylinderParameters", "CoordinateSystemID:=", -1, "XCenter:=",  _
  "0mm", "YCenter:=", "0mm", "ZCenter:=", "port_gap/2", "Radius:=", "wire_rad", "Height:=",  _
  "dipole_length/2-port_gap/2", "WhichAxis:=", "Z", "NumSides:=", "0"), Array("NAME:Attributes", "Name:=",  _
  "arm1", "Flags:=", "", "Color:=", "(132 132 193)", "Transparency:=",  _
  0.8, "PartCoordinateSystem:=", "Global", "MaterialName:=",  _
  "pec", "SolveInside:=", false)
  

  
oEditor.DuplicateAroundAxis Array("NAME:Selections", "Selections:=", "arm1", "NewPartsModelFlag:=",  _
  "Model"), Array("NAME:DuplicateAroundAxisParameters", "CoordinateSystemID:=", -1, "CreateNewObjects:=",  _
  true, "WhichAxis:=", "X", "AngleStr:=", "180deg", "NumClones:=", "2"), Array("NAME:Options", "DuplicateBoundaries:=",  _
  true)
  
oEditor.CreateRectangle Array("NAME:RectangleParameters", "CoordinateSystemID:=",  _
  -1, "IsCovered:=", true, "XStart:=", "0mm", "YStart:=", "-wire_rad", "ZStart:=",  _
  "-port_gap/2", "Width:=", "wire_rad*2", "Height:=",  _
  "port_gap", "WhichAxis:=", "X"), Array("NAME:Attributes", "Name:=",  _
  "port1", "Flags:=", "", "Color:=", "(132 132 193)", "Transparency:=",  _
  0.8, "PartCoordinateSystem:=", "Global", "MaterialName:=",  _
  "vacuum", "SolveInside:=", true)
  

dim end_vector_point
dim start_vector_point
end_vector_point = port_gap/2
end_vector_point = end_vector_point & units


start_vector_point = -port_gap/2
start_vector_point = start_vector_point & units

oDesign.SetSolutionType "DrivenModal"
Set oModule = oDesign.GetModule("BoundarySetup")
oModule.AssignLumpedPort Array("NAME:p1", "Objects:=", Array("port1"), Array("NAME:Modes", Array("NAME:Mode1", "ModeNum:=",  _
  1, "UseIntLine:=", true, Array("NAME:IntLine", "Start:=", Array("0mm", "0mm",  _
  start_vector_point), "End:=", Array("0mm", "0mm", end_vector_point)), "CharImp:=", "Zpi", "RenormImp:=",  _
  "50ohm")), "FullResistance:=", "50ohm", "FullReactance:=", "0ohm")

  

  
Set oModule = oDesign.GetModule("AnalysisSetup")

oModule.InsertSetup "HfssDriven", Array("NAME:Setup1", "Frequency:=", center_freq&"GHZ", "PortsOnly:=",  _
  false, "MaxDeltaS:=", 0.02, "UseMatrixConv:=", false, "MaximumPasses:=", 15, "MinimumPasses:=",  _
  1, "MinimumConvergedPasses:=", 1, "PercentRefinement:=", 30, "BasisOrder:=", 1, "UseIterativeSolver:=",  _
  false, "DoLambdaRefine:=", true, "DoMaterialLambda:=", true, "SetLambdaTarget:=",  _
  false, "Target:=", 0.3333, "UseConvOutputVariable:=", false, "IsEnabled:=",  _
  true, "ExternalMesh:=", false, "UseMaxTetIncrease:=", false, "MaxTetIncrease:=",  _
  100000, "PortAccuracy:=", 2, "UseABCOnPort:=", false, "SetPortMinMaxTri:=",  _
  false)
  

dim start_freq
dim stop_freq
start_freq = .5*center_freq
stop_freq = 1.5*center_freq

 oModule.InsertFrequencySweep "Setup1", Array("NAME:Sweep1", "IsEnabled:=", true, "SetupType:=",  _
  "LinearCount", "StartValue:=", start_freq&"GHz", "StopValue:=", stop_freq&"GHz", "Count:=", 100, "Type:=",  _
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

extent_x_pos = "wire_rad"
extent_x_neg = "-wire_rad"

extent_y_pos = "wire_rad"
extent_y_neg = "-wire_rad"

extent_z_pos = "dipole_length/2"
extent_z_neg = "-(dipole_length/2)"

if israd = "ABC" then
 israd = "Rad"

end if


dim mycommand
' full command which invokes wscript to run desired VBScript and passes
mycommand = "wscript.exe " & """" & "./VBScripts/rad_creation.vbs" & """" &  " " & extent_x_pos & " " & extent_x_neg & " " & extent_y_pos & " " & extent_y_neg & " " & extent_z_pos & " " & extent_z_neg & " " & units & " " & israd& " " & 0& " " & 0
'test = "wscript.exe " & """" & "./VBScripts/rad_creation.vbs" & """" &  " " & extent_x_pos & " " & extent_x_neg & " " & extent_y_pos & " " & extent_y_neg & " " & extent_z_pos & " " & extent_z_neg & " " & units & " " & israd& " " & 0& " " & 0

'msgbox(mycommand)
' run the desired VBScript
Set WSHShell = CreateObject("WScript.Shell")
WSHShell.Run mycommand, , True

Setlocale(locallang)