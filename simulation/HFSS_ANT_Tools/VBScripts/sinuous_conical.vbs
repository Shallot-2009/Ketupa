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
oProject.InsertDesign "HFSS", "SinuousConical_Antenna_ADKv1", "DrivenModal", ""
Set oDesign = oProject.SetActiveDesign("SinuousConical_Antenna_ADKv1")








dim  solution_freq, NumberOfPoints, NumberOfCells, Alpha, GrowRate, OuterRadius, Delta, NumberOfArms, ConeHeight, port_extension, units, israd

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


solution_freq = CDbl( args(0))






NumberOfPoints = CDbl( args(1))
NumberOfCells = CDbl( args(2))
Alpha = CDbl( args(3))
GrowRate = CDbl( args(4))
OuterRadius = CDbl( args(5))
Delta = CDbl( args(6))
NumberOfArms = CDbl( args(7))
ConeHeight = CDbl( args(8))



port_extension = CDbl( args(9))

units = args(10)


israd = args(11) ' 0 is for radiation surface, 1 is for PML


'''''''''''''''''''''''''''''''''''''''''''''


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

dim low_freq, high_freq, low_lambda, high_lambda, InnerRadius_calc


dim scale_factor
scale_factor = 1.25


low_lambda = 4*OuterRadius*(Alpha*3.14159/180+Delta*3.14159/180)
low_freq = Round(scale_factor*light_speed/low_lambda*1e-9,2)  'in Ghz


InnerRadius_calc = GrowRate^(NumberOfCells-1)*OuterRadius
high_lambda = 4*2*InnerRadius_calc*(Alpha*3.14159/180+Delta*3.14159/180)
high_freq = Round(scale_factor*light_speed/high_lambda*1e-9,2)  'in GHz

dim Bandwidth

Bandwidth = high_freq/low_freq

if Bandwidth >14 then
msgbox("Greater than 14:1 Bandwidth may not be practical. Please Check Dimensions.")
high_freq = Round(low_freq*14,2)
solution_freq = (high_freq+low_freq)/2+low_freq
end if




''''''''''''''''''''''''''''''''''''''''''''





oDesign.ChangeProperty Array("NAME:AllTabs", Array("NAME:LocalVariableTab", Array("NAME:PropServers", "LocalVariables"), Array("NAME:NewProps", _
Array("NAME:--Spiral Dimensions", "PropType:=", "SeparatorProp", "UserDef:=", true, "Value:=", ""), _
Array("NAME:NumberOfPoints", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", NumberOfPoints), _
Array("NAME:NumberOfCells", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", NumberOfCells), _
Array("NAME:Alpha", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", Alpha & "deg"), _
Array("NAME:GrowRate", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", GrowRate), _
Array("NAME:OuterRadius", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", OuterRadius & units), _
Array("NAME:Delta", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", Delta & "deg"), _
Array("NAME:NumberOfArms", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", NumberOfArms), _
Array("NAME:ConeHeight", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", ConeHeight & units), _
Array("NAME:--Port", "PropType:=", "SeparatorProp", "UserDef:=", true, "Value:=", ""), _
Array("NAME:port_extension", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", port_extension & units))))






''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Draws spiral



Set oEditor = oDesign.SetActiveEditor("3D Modeler")
oEditor.CreateUserDefinedPart Array("NAME:UserDefinedPrimitiveParameters", "DllName:=",  _
  "ADKv1/sinuous_adkv1", "Version:=", "1.1.1", "NoOfParameters:=", 8, "Library:=", "userlib", Array("NAME:ParamVector", Array("NAME:Pair", "Name:=",  _
  "NumberOfPoints", "Value:=", "NumberOfPoints"), Array("NAME:Pair", "Name:=", "NumberOfCells", "Value:=",  _
  "NumberOfCells"), Array("NAME:Pair", "Name:=", "Alpha", "Value:=", "Alpha"), Array("NAME:Pair", "Name:=",  _
  "GrowRate", "Value:=", "GrowRate"), Array("NAME:Pair", "Name:=", "OuterRadius", "Value:=",  _
  "OuterRadius"), Array("NAME:Pair", "Name:=", "Delta", "Value:=", "Delta"), Array("NAME:Pair", "Name:=",  _
  "NumberOfArms", "Value:=", "NumberOfArms"), Array("NAME:Pair", "Name:=", "ConeHeight", "Value:=",  _
  "ConeHeight"))), Array("NAME:Attributes", "Name:=", "SinuousAntenna1", "Flags:=", "", "Color:=",  _
  "(255 128 65)", "Transparency:=", 0.3, "PartCoordinateSystem:=",  _
  "Global", "MaterialName:=", "vacuum", "SolveInside:=", true)
 
'get edge ID associated with port, this edge name comes from within the UDP, this
' is just an easier way of creating the port because the inner edge location is not easily known from within the script
' this section will also create an object from that edge and connect orthogonal edges to create a port
  
dim port1_edgeID1, port1_edgeID2, port2_edgeID1, port2_edgeID2


port1_edgeID1 = oEditor.GetEdgeIDFromNameForFirstOperation("SinuousAntenna1", "port1_edge1")
port1_edgeID2 = oEditor.GetEdgeIDFromNameForFirstOperation("SinuousAntenna1", "port1_edge2")


oEditor.CreateObjectFromEdges Array("NAME:Selections", "Selections:=",  _
  "SinuousAntenna1", "NewPartsModelFlag:=", "Model"), Array("NAME:Parameters", Array("NAME:BodyFromEdgeToParameters", "CoordinateSystemID:=",  _
  -1, "Edges:=", Array(port1_edgeID1,port1_edgeID2)))

oEditor.ChangeProperty Array("NAME:AllTabs", Array("NAME:Geometry3DAttributeTab", Array("NAME:PropServers",  _
"SinuousAntenna1_ObjectFromEdge1"), Array("NAME:ChangedProps", Array("NAME:Name", "Value:=",  _
"port1_edge1"))))

oEditor.ChangeProperty Array("NAME:AllTabs", Array("NAME:Geometry3DAttributeTab", Array("NAME:PropServers",  _
"SinuousAntenna1_ObjectFromEdge2"), Array("NAME:ChangedProps", Array("NAME:Name", "Value:=",  _
"port1_edge2"))))


oEditor.Connect Array("NAME:Selections", "Selections:=","port1_edge1,port1_edge2")

oEditor.ChangeProperty Array("NAME:AllTabs", Array("NAME:Geometry3DAttributeTab", Array("NAME:PropServers",  _
"port1_edge1"), Array("NAME:ChangedProps", Array("NAME:Name", "Value:=",  _
"port1"))))


'this gets the face id for port 1 to be used in port creation
Dim faceid_for_port1
faceid_for_port1 = oEditor.GetFaceByPosition(Array("NAME:FaceParameters", "BodyName:=", "port1", _
"XPosition:=", "0mm", "YPosition:=", "0mm", "ZPosition:=", ConeHeight & units))



if NumberOfArms = 4 then
port2_edgeID1 = oEditor.GetEdgeIDFromNameForFirstOperation("SinuousAntenna1", "port2_edge1")
port2_edgeID2 = oEditor.GetEdgeIDFromNameForFirstOperation("SinuousAntenna1", "port2_edge2")

oEditor.CreateObjectFromEdges Array("NAME:Selections", "Selections:=",  _
  "SinuousAntenna1", "NewPartsModelFlag:=", "Model"), Array("NAME:Parameters", Array("NAME:BodyFromEdgeToParameters", "CoordinateSystemID:=",  _
  -1, "Edges:=", Array(port2_edgeID1,port2_edgeID2)))
  
oEditor.ChangeProperty Array("NAME:AllTabs", Array("NAME:Geometry3DAttributeTab", Array("NAME:PropServers",  _
"SinuousAntenna1_ObjectFromEdge1"), Array("NAME:ChangedProps", Array("NAME:Name", "Value:=",  _
"port2_edge1"))))

oEditor.ChangeProperty Array("NAME:AllTabs", Array("NAME:Geometry3DAttributeTab", Array("NAME:PropServers",  _
"SinuousAntenna1_ObjectFromEdge2"), Array("NAME:ChangedProps", Array("NAME:Name", "Value:=",  _
"port2_edge2"))))  

oEditor.Connect Array("NAME:Selections", "Selections:=","port2_edge1,port2_edge2")
  
oEditor.ChangeProperty Array("NAME:AllTabs", Array("NAME:Geometry3DAttributeTab", Array("NAME:PropServers",  _
"port2_edge1"), Array("NAME:ChangedProps", Array("NAME:Name", "Value:=",  _
"port2"))))

oEditor.CreateObjectFromEdges Array("NAME:Selections", "Selections:=",  _
  "SinuousAntenna1", "NewPartsModelFlag:=", "Model"), Array("NAME:Parameters", Array("NAME:BodyFromEdgeToParameters", "CoordinateSystemID:=",  _
  -1, "Edges:=", Array(port2_edgeID1,port2_edgeID2)))
  
oEditor.ChangeProperty Array("NAME:AllTabs", Array("NAME:Geometry3DAttributeTab", Array("NAME:PropServers",  _
"SinuousAntenna1_ObjectFromEdge1"), Array("NAME:ChangedProps", Array("NAME:Name", "Value:=",  _
"port_ext1"))))

oEditor.ChangeProperty Array("NAME:AllTabs", Array("NAME:Geometry3DAttributeTab", Array("NAME:PropServers",  _
"SinuousAntenna1_ObjectFromEdge2"), Array("NAME:ChangedProps", Array("NAME:Name", "Value:=",  _
"port_ext2"))))  
    
oEditor.SweepAlongVector Array("NAME:Selections", "Selections:=",  _
  "port_ext1", "NewPartsModelFlag:=", "Model"), Array("NAME:VectorSweepParameters", "CoordinateSystemID:=",  _
  -1, "DraftAngle:=", "0deg", "DraftType:=", "Round", "CheckFaceFaceIntersection:=",  _
  false, "SweepVectorX:=", "0mm", "SweepVectorY:=", "0mm", "SweepVectorZ:=",  _
  "port_extension")  
  
oEditor.SweepAlongVector Array("NAME:Selections", "Selections:=",  _
  "port_ext2", "NewPartsModelFlag:=", "Model"), Array("NAME:VectorSweepParameters", "CoordinateSystemID:=",  _
  -1, "DraftAngle:=", "0deg", "DraftType:=", "Round", "CheckFaceFaceIntersection:=",  _
  false, "SweepVectorX:=", "0mm", "SweepVectorY:=", "0mm", "SweepVectorZ:=",  _
  "port_extension")
  
oEditor.Move Array("NAME:Selections", "Selections:=", "port2", "NewPartsModelFlag:=",  _
  "Model"), Array("NAME:TranslateParameters", "CoordinateSystemID:=", -1, "TranslateVectorX:=",  _
  "0mm", "TranslateVectorY:=", "0mm", "TranslateVectorZ:=", "port_extension")   
  
'this gets the face id for port 2 to be used in port creation
dim port2_height
port2_height = ConeHeight + port_extension 
Dim faceid_for_port2
faceid_for_port2 = oEditor.GetFaceByPosition(Array("NAME:FaceParameters", "BodyName:=", "port2", _
"XPosition:=", "0mm", "YPosition:=", "0mm", "ZPosition:=", port2_height & units))
  
end if



oEditor.SeparateBody Array("NAME:Selections", "Selections:=", "SinuousAntenna1", "NewPartsModelFlag:=",  _
  "Model")

  Set oModule = oDesign.GetModule("BoundarySetup")  
oModule.AssignPerfectE Array("NAME:arm1", "Objects:=", Array("SinuousAntenna1"), "InfGroundPlane:=", false)
oModule.AssignPerfectE Array("NAME:arm2", "Objects:=", Array("SinuousAntenna1_Separate1"), "InfGroundPlane:=", false)





oDesign.SetSolutionType "DrivenTerminal"
Set oModule = oDesign.GetModule("BoundarySetup")
oModule.AutoIdentifyPorts Array("NAME:Faces", faceid_for_port1), false, Array("NAME:ReferenceConductors",  _
  "SinuousAntenna1_Separate1")
  
oModule.EditTerminal "SinuousAntenna1_T1", Array("NAME:SinuousAntenna1_T1", "ParentBndID:=",  _
  "p1", "TerminalResistance:=", "266ohm")
  
  

if NumberOfArms = 4 then
  
  
  Set oModule = oDesign.GetModule("BoundarySetup")  
oModule.AssignPerfectE Array("NAME:port_ext1", "Objects:=", Array("port_ext1"), "InfGroundPlane:=", false)
oModule.AssignPerfectE Array("NAME:port_ext2", "Objects:=", Array("port_ext2"), "InfGroundPlane:=", false)   

oModule.AssignPerfectE Array("NAME:arm3", "Objects:=", Array("SinuousAntenna1_Separate2"), "InfGroundPlane:=", false)
oModule.AssignPerfectE Array("NAME:arm4", "Objects:=", Array("SinuousAntenna1_Separate3"), "InfGroundPlane:=", false)

oDesign.SetSolutionType "DrivenTerminal"
Set oModule = oDesign.GetModule("BoundarySetup")
oModule.AutoIdentifyPorts Array("NAME:Faces", faceid_for_port2), false, Array("NAME:ReferenceConductors",  _
  "port_ext2")

oModule.EditTerminal "port_ext1_T1", Array("NAME:port_ext1_T1", "ParentBndID:=",  _
  "p2", "TerminalResistance:=", "266ohm")

Set oModule = oDesign.GetModule("Solutions")
oModule.EditSources "NoIncidentWave", Array("NAME:Names", "1", "2"), Array("NAME:Terminals",  _
  1, 1), Array("NAME:Magnitudes", "1", "1"), Array("NAME:Phases", "0deg", "90deg"), Array("NAME:Terminated",  _
  false, false), Array("NAME:Impedances")  


end if



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


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

extent_x_pos = "OuterRadius"
extent_x_neg = "-OuterRadius"

extent_y_pos = "OuterRadius"
extent_y_neg = "-OuterRadius"

extent_z_pos = "ConeHeight"
extent_z_neg = "0"

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
  