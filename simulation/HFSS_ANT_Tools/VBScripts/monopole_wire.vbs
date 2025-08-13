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
oProject.InsertDesign "HFSS", "Monopole_Antenna_ADKv1", "DrivenModal", ""
Set oDesign = oProject.SetActiveDesign("Monopole_Antenna_ADKv1")

dim center_freq, wire_rad, port_gap, monopole_length, ground_width, units, israd

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

monopole_length = CDbl( args(3))

ground_width = CDbl( args(4))

units = args(5)


israd = args(6) ' 0 is for radiation surface, 1 is for PML









oDesign.ChangeProperty Array("NAME:AllTabs", Array("NAME:LocalVariableTab", Array("NAME:PropServers", "LocalVariables"), Array("NAME:NewProps", _
Array("NAME:--Monopole Dimensions", "PropType:=", "SeparatorProp", "UserDef:=", true, "Value:=", ""), _
Array("NAME:wire_rad", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", wire_rad & units), _
Array("NAME:monopole_length", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", monopole_length & units), _
Array("NAME:ground_width", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", ground_width & units), _
Array("NAME:port_gap", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", port_gap & units))))
 
  
  
Set oEditor = oDesign.SetActiveEditor("3D Modeler")

oEditor.CreateCylinder Array("NAME:CylinderParameters", "CoordinateSystemID:=", -1, "XCenter:=",  _
  "0mm", "YCenter:=", "0mm", "ZCenter:=", "port_gap", "Radius:=", "wire_rad", "Height:=",  _
  "monopole_length-port_gap", "WhichAxis:=", "Z", "NumSides:=", "0"), Array("NAME:Attributes", "Name:=",  _
  "arm1", "Flags:=", "", "Color:=", "(132 132 193)", "Transparency:=",  _
  0.8, "PartCoordinateSystem:=", "Global", "MaterialName:=",  _
  "pec", "SolveInside:=", false)
  

  

  
oEditor.CreateRectangle Array("NAME:RectangleParameters", "CoordinateSystemID:=",  _
  -1, "IsCovered:=", true, "XStart:=", "0mm", "YStart:=", "-wire_rad", "ZStart:=",  _
  "0", "Width:=", "wire_rad*2", "Height:=",  _
  "port_gap", "WhichAxis:=", "X"), Array("NAME:Attributes", "Name:=",  _
  "port1", "Flags:=", "", "Color:=", "(132 132 193)", "Transparency:=",  _
  0.8, "PartCoordinateSystem:=", "Global", "MaterialName:=",  _
  "vacuum", "SolveInside:=", true)
 
oEditor.CreateRectangle Array("NAME:RectangleParameters", "CoordinateSystemID:=",  _
  -1, "IsCovered:=", true, "XStart:=", "-ground_width/2", "YStart:=", "-ground_width/2", "ZStart:=",  _
  "0", "Width:=", "ground_width", "Height:=",  _
  "ground_width", "WhichAxis:=", "Z"), Array("NAME:Attributes", "Name:=",  _
  "Ground", "Flags:=", "", "Color:=", "(132 132 193)", "Transparency:=",  _
  0.8, "PartCoordinateSystem:=", "Global", "MaterialName:=",  _
  "vacuum", "SolveInside:=", true) 
  
  
Set oModule = oDesign.GetModule("BoundarySetup")
oModule.AssignPerfectE Array("NAME:PerfE2", "Objects:=", Array("Ground"), "InfGroundPlane:=",  _
  false)
  

dim end_vector_point
dim start_vector_point
end_vector_point = port_gap/2
end_vector_point = end_vector_point & units


start_vector_point = 0
start_vector_point = start_vector_point & units

oDesign.SetSolutionType "DrivenModal"
Set oModule = oDesign.GetModule("BoundarySetup")
oModule.AssignLumpedPort Array("NAME:p1", "Objects:=", Array("port1"), Array("NAME:Modes", Array("NAME:Mode1", "ModeNum:=",  _
  1, "UseIntLine:=", true, Array("NAME:IntLine", "Start:=", Array("0mm", "0mm",  _
  start_vector_point), "End:=", Array("0mm", "0mm", end_vector_point)), "CharImp:=", "Zpi", "RenormImp:=",  _
  "25ohm")), "FullResistance:=", "25ohm", "FullReactance:=", "0ohm")

  

  
Set oModule = oDesign.GetModule("AnalysisSetup")



oModule.InsertSetup "HfssDriven", Array("NAME:Setup1", "Frequency:=", center_freq&"GHZ", "PortsOnly:=",  _
  false, "MaxDeltaS:=", 0.02, "UseMatrixConv:=", false, "MaximumPasses:=", 15, "MinimumPasses:=",  _
  1, "MinimumConvergedPasses:=", 1, "PercentRefinement:=", 30, "IsEnabled:=",  _
  true, "BasisOrder:=", -1, "UseIterativeSolver:=", true, "IterativeResidual:=",  _
  0.0001, "DoLambdaRefine:=", true, "DoMaterialLambda:=", true, "SetLambdaTarget:=",  _
  false, "Target:=", 0.6667, "UseMaxTetIncrease:=", false, "MaxTetIncrease:=",  _
  1000000, "PortAccuracy:=", 2, "UseABCOnPort:=", false, "SetPortMinMaxTri:=",  _
  false, "EnableSolverDomains:=", false, "ThermalFeedback:=", false, "UsingNumSolveSteps:=",  _
  0, "ConstantDelta:=", "0s", "NumberSolveSteps:=", 1)
 
 
 
  

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

extent_x_pos = "ground_width/2"
extent_x_neg = "-ground_width/2"

extent_y_pos = "ground_width/2"
extent_y_neg = "-ground_width/2"

extent_z_pos = "monopole_length"
extent_z_neg = "0"

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