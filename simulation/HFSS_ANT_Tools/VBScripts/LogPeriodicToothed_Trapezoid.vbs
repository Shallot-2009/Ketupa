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
oProject.InsertDesign "HFSS", "LogPeriodicToothed_Antenna_ADKv1", "DrivenModal", ""
Set oDesign = oProject.SetActiveDesign("LogPeriodicToothed_Antenna_ADKv1")








dim  solution_freq, Length, Tau, Sigma, Delta_Angle, Beta_Angle, Port_Gap_Width, subH, subX, subY, units, israd

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




Length = CDbl(args(1))
Tau = CDbl(args(2))
Sigma = CDbl(args(3))
Delta_Angle = CDbl(args(4))
Beta_Angle = CDbl(args(5))
Port_Gap_Width = CDbl(args(6))

subH = (args(7))
subX = CDbl(args(8))
subY = CDbl(args(9))


units = args(10)

israd = args(11) ' 0 is for radiation surface, 1 is for PML


if Length > subY/2 then
  subY = Length*2
end if

if subX/2 < Length/tan(90*3.14159265358979323846/180-(Beta_Angle*3.14159265358979323846/180/2+Delta_Angle*3.14159265358979323846/180)) then
    subX = (Length/tan(90*3.14159265358979323846/180-(Beta_Angle*3.14159265358979323846/180/2+Delta_Angle*3.14159265358979323846/180)))*2.1
end if





''''''''''''''''''''''''''''''''


dim light_speed
light_speed = 299792458

select case units
 case "um"
   light_speed = 299800000000000
 case "mm"
   light_speed = 299800000000
 case "cm"
   light_speed = 29980000000
 case "m"
   light_speed = 299800000
  case "mil"
   light_speed = 11800000000000
  case "in"
   light_speed = 11800000000
  case "ft"
   light_speed = 983600000
end select



dim low_freq, high_freq

dim beta_rad, delta_rad, low_labda, high_lambda
beta_rad = Beta_Angle*3.14159/180
delta_rad = Delta_Angle*3.14159/180

  
  low_lambda = 2*Length*(tan(beta_rad/2)+tan(beta_rad/2+delta_rad))
  high_lambda = 2*Port_Gap_Width*(tan(beta_rad/2)+tan(beta_rad/2+delta_rad))


high_freq = Round(light_speed/(1.4^.5*high_lambda)*1e-9,1)


low_freq = Round(light_speed/(1.4^.5*low_lambda)*1e-9,1)

dim Bandwidth

Bandwidth = high_freq/low_freq

if Bandwidth > 5 then
solution_freq = Round((high_freq+low_freq)/2+low_freq,1)
end if

if Bandwidth >10 then
msgbox("Greater than 10:1 Bandwidth may not be practical. Please Check Dimensions.")
high_freq = Round(low_freq*10,1)
solution_freq = Round((high_freq+low_freq)/2+low_freq,1)
end if






oDesign.ChangeProperty Array("NAME:AllTabs", Array("NAME:LocalVariableTab", Array("NAME:PropServers", "LocalVariables"), Array("NAME:NewProps", _
Array("NAME:--Antenna Dimensions", "PropType:=", "SeparatorProp", "UserDef:=", true, "Value:=", ""), _
Array("NAME:Length", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", Length & units), _
Array("NAME:Tau", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", Tau), _
Array("NAME:Sigma", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", Sigma), _
Array("NAME:Delta_Angle", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", Delta_Angle & "deg"), _
Array("NAME:Beta_Angle", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", Beta_Angle & "deg"), _
Array("NAME:Port_Gap_Width", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", Port_Gap_Width & units), _
Array("NAME:--Substrate Dimensions", "PropType:=", "SeparatorProp", "UserDef:=", true, "Value:=", ""), _
Array("NAME:subH", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", subH ), _
Array("NAME:subX", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", subX & units), _
Array("NAME:subY", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", subY & units))))



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'create substrate material

Material_Name = args(12)
Material_Name = Material_Name & "_ADK"
Permittivity = args(13)
TandD = args(14)

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

oEditor.CreateUserDefinedPart Array("NAME:UserDefinedPrimitiveParameters", "DllName:=",  _
  "ADKv1/LogPeriodicToothed_Trapezoid_adkv1", "Version:=", "2.0", "NoOfParameters:=", 6, "Library:=", "userlib", Array("NAME:ParamVector", Array("NAME:Pair", "Name:=",  _
  "Length", "Value:=", "Length"), Array("NAME:Pair", "Name:=", "Tau", "Value:=",  _
  "Tau"), Array("NAME:Pair", "Name:=", "Sigma", "Value:=", "Sigma"), Array("NAME:Pair", "Name:=",  _
  "Delta_Angle", "Value:=", "Delta_Angle"), Array("NAME:Pair", "Name:=", "Beta_Angle", "Value:=",  _
  "Beta_Angle"), Array("NAME:Pair", "Name:=", "Port_Gap_Width", "Value:=", "Port_Gap_Width"))), Array("NAME:Attributes", "Name:=",  _
  "LogPeriodicToothedPlanarAntenna1", "Flags:=", "", "Color:=", "(255 128 65)", "Transparency:=",  _
  0.3, "PartCoordinateSystem:=", "Global", "MaterialName:=",  _
  "pec", "SolveInside:=", false)






'get edge ID associated with port, this edge name comes from within the UDP, this
' is just an easier way of creating the port because the inner edge location is not easily known from within the script
' this section will also create an object from that edge and connect orthogonal edges to create a port
  
dim port1_edgeID1, port1_edgeID2


port1_edgeID1 = oEditor.GetEdgeIDFromNameForFirstOperation("LogPeriodicToothedPlanarAntenna1", "port_edge1")
port1_edgeID2 = oEditor.GetEdgeIDFromNameForFirstOperation("LogPeriodicToothedPlanarAntenna1", "port_edge2")


oEditor.CreateObjectFromEdges Array("NAME:Selections", "Selections:=",  _
  "LogPeriodicToothedPlanarAntenna1", "NewPartsModelFlag:=", "Model"), Array("NAME:Parameters", Array("NAME:BodyFromEdgeToParameters", "CoordinateSystemID:=",  _
  -1, "Edges:=", Array(port1_edgeID1,port1_edgeID2)))

oEditor.ChangeProperty Array("NAME:AllTabs", Array("NAME:Geometry3DAttributeTab", Array("NAME:PropServers",  _
"LogPeriodicToothedPlanarAntenna1_ObjectFromEdge1"), Array("NAME:ChangedProps", Array("NAME:Name", "Value:=",  _
"port1_edge1"))))

oEditor.ChangeProperty Array("NAME:AllTabs", Array("NAME:Geometry3DAttributeTab", Array("NAME:PropServers",  _
"LogPeriodicToothedPlanarAntenna1_ObjectFromEdge2"), Array("NAME:ChangedProps", Array("NAME:Name", "Value:=",  _
"port1_edge2"))))


oEditor.Connect Array("NAME:Selections", "Selections:=","port1_edge1,port1_edge2")

oEditor.ChangeProperty Array("NAME:AllTabs", Array("NAME:Geometry3DAttributeTab", Array("NAME:PropServers",  _
"port1_edge1"), Array("NAME:ChangedProps", Array("NAME:Name", "Value:=",  _
"port1"))))


'this gets the face id for port 1 to be used in port creation
Dim faceid_for_port1
faceid_for_port1 = oEditor.GetFaceByPosition(Array("NAME:FaceParameters", "BodyName:=", "port1", _
"XPosition:=", "0mm", "YPosition:=", "0mm", "ZPosition:=", "0mm"))  
  
  
oEditor.SeparateBody Array("NAME:Selections", "Selections:=", "LogPeriodicToothedPlanarAntenna1", "NewPartsModelFlag:=",  _
  "Model")

  Set oModule = oDesign.GetModule("BoundarySetup")  
oModule.AssignPerfectE Array("NAME:arm1", "Objects:=", Array("LogPeriodicToothedPlanarAntenna1"), "InfGroundPlane:=", false)
oModule.AssignPerfectE Array("NAME:arm2", "Objects:=", Array("LogPeriodicToothedPlanarAntenna1_Separate1"), "InfGroundPlane:=", false)  
  

oDesign.SetSolutionType "DrivenTerminal"
Set oModule = oDesign.GetModule("BoundarySetup")
oModule.AutoIdentifyPorts Array("NAME:Faces", faceid_for_port1), false, Array("NAME:ReferenceConductors",  _
  "LogPeriodicToothedPlanarAntenna1")

oModule.EditTerminal "LogPeriodicToothedPlanarAntenna1_Separate1_T1", Array("NAME:LogPeriodicToothedPlanarAntenna1_Separate1_T1", "ParentBndID:=",  _
  "p1", "TerminalResistance:=", "188.5ohm")










''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Solution Setup 
  
Set oEditor = oDesign.SetActiveEditor("3D Modeler")
oEditor.SetModelUnits Array("NAME:Units Parameter", "Units:=", units, "Rescale:=",  _
  false)


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
start_freq = low_freq
stop_freq = high_freq

 oModule.InsertFrequencySweep "Setup1", Array("NAME:Sweep1", "IsEnabled:=", true, "SetupType:=",  _
  "LinearCount", "StartValue:=", low_freq&"GHz", "StopValue:=", high_freq&"GHz", "Count:=", 200, "Type:=",  _
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
mycommand = "wscript.exe " & """" & "./VBScripts/rad_creation.vbs" & """" &  " " & extent_x_pos & " " & extent_x_neg & " " & extent_y_pos & " " & extent_y_neg & " " & extent_z_pos & " " & extent_z_neg & " " & units & " " & israd& " " & low_freq
'msgbox(mycommand)
' run the desired VBScript
Set WSHShell = CreateObject("WScript.Shell")
WSHShell.Run mycommand, , True
Setlocale(locallang)
