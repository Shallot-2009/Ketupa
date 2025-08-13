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
oProject.InsertDesign "HFSS-IE", "Sinuous_Antenna_ADKv1", "DrivenModal", ""
Set oDesign = oProject.SetActiveDesign("Sinuous_Antenna_ADKv1")








dim  solution_freq, NumberOfPoints, NumberOfCells, Alpha, GrowRate, OuterRadius, Delta, NumberOfArms, subH, subX, subY, port_extension, units, israd

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

'subH = CDbl( args(8))
'subX = CDbl( args(9))
'subY = CDbl( args(10))

port_extension = CDbl( args(8))

units = args(9)


israd = args(10) ' 0 is for radiation surface, 1 is for PML


' this still constrains the substrate size even if it is larger than the
' outerradius. Does the greater than operation require a conversion to 
' number from string first?

'ans = IsNumeric(OuterRadius)
'msgbox(ans)

'if OuterRadius > (subX/2) then
'  subX = OuterRadius*2
'end if
'if OuterRadius > (subY/2) then
'  subY = OuterRadius*2
'end if

'calculate recommended min and max frequency for given dimensions

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


'''''''''''''''''''''''''''''''''''''




oDesign.ChangeProperty Array("NAME:AllTabs", Array("NAME:LocalVariableTab", Array("NAME:PropServers", "LocalVariables"), Array("NAME:NewProps", _
Array("NAME:--Spiral Dimensions", "PropType:=", "SeparatorProp", "UserDef:=", true, "Value:=", ""), _
Array("NAME:NumberOfPoints", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", NumberOfPoints), _
Array("NAME:NumberOfCells", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", NumberOfCells), _
Array("NAME:Alpha", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", Alpha & "deg"), _
Array("NAME:GrowRate", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", GrowRate), _
Array("NAME:OuterRadius", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", OuterRadius & units), _
Array("NAME:Delta", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", Delta & "deg"), _
Array("NAME:NumberOfArms", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", NumberOfArms), _
Array("NAME:--Port", "PropType:=", "SeparatorProp", "UserDef:=", true, "Value:=", ""), _
Array("NAME:port_extension", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", port_extension & units))))

'Array("NAME:--Substrate Dimensions", "PropType:=", "SeparatorProp", "UserDef:=", true, "Value:=", ""), _
'Array("NAME:subH", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", subH & units), _
'Array("NAME:subX", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", subX & units), _
'Array("NAME:subY", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", subY & units), _


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Draws Substrate and bottom metalization



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
  "0mm"))), Array("NAME:Attributes", "Name:=", "SinuousAntenna1", "Flags:=", "", "Color:=",  _
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
"XPosition:=", "0mm", "YPosition:=", "0mm", "ZPosition:=", "0mm"))



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
Dim faceid_for_port2
faceid_for_port2 = oEditor.GetFaceByPosition(Array("NAME:FaceParameters", "BodyName:=", "port2", _
"XPosition:=", "0mm", "YPosition:=", "0mm", "ZPosition:=", port_extension & units))
  
end if



oEditor.SeparateBody Array("NAME:Selections", "Selections:=", "SinuousAntenna1", "NewPartsModelFlag:=",  _
  "Model")

  Set oModule = oDesign.GetModule("BoundarySetup")  
oModule.AssignPerfectE Array("NAME:arm1", "Objects:=", Array("SinuousAntenna1"), "InfGroundPlane:=", false)
oModule.AssignPerfectE Array("NAME:arm2", "Objects:=", Array("SinuousAntenna1_Separate1"), "InfGroundPlane:=", false)





Set oModule = oDesign.GetModule("BoundarySetup")

'auto identify ports seems to add too many terminals, manually adding here
oModule.AssignLumpedPort Array("NAME:Port1", "Objects:=", Array("port1"), "RenormalizeAllTerminals:=",  _
  true, "TerminalIDList:=", Array())
oModule.AssignTerminal Array("NAME:port1_T1", "Edges:=", Array(port1_edgeID1), "ParentBndID:=",  _
  "Port1", "TerminalResistance:=", "266ohm")
   



  
  
  


if NumberOfArms = 4 then
  
  
  Set oModule = oDesign.GetModule("BoundarySetup")  
oModule.AssignPerfectE Array("NAME:port_ext1", "Objects:=", Array("port_ext1"), "InfGroundPlane:=", false)
oModule.AssignPerfectE Array("NAME:port_ext2", "Objects:=", Array("port_ext2"), "InfGroundPlane:=", false)   

oModule.AssignPerfectE Array("NAME:arm3", "Objects:=", Array("SinuousAntenna1_Separate2"), "InfGroundPlane:=", false)
oModule.AssignPerfectE Array("NAME:arm4", "Objects:=", Array("SinuousAntenna1_Separate3"), "InfGroundPlane:=", false)

'oDesign.SetSolutionType "DrivenTerminal"
Set oModule = oDesign.GetModule("BoundarySetup")


 
oModule.AssignLumpedPort Array("NAME:Port2", "Objects:=", Array("port2"), "RenormalizeAllTerminals:=",  _
  true, "TerminalIDList:=", Array())
oModule.AutoIdentifyTerminals Array("NAME:ReferenceConductors", "port_ext2"),  _
  "Port2", true
oModule.EditTerminal "port_ext1_T1", Array("NAME:port_ext1_T1", "ParentBndID:=",  _
  "Port2", "TerminalResistance:=", "266ohm")


Set oModule = oDesign.GetModule("Solutions")
oModule.EditSources Array("NAME:Sources", "port1_T1:=", Array(false, "1", "0deg"), "port_ext1_T1:=", Array( _
  false, "1", "0deg"))

end if



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Solution Setup 
  
Set oEditor = oDesign.SetActiveEditor("3D Modeler")
oEditor.SetModelUnits Array("NAME:Units Parameter", "Units:=", units, "Rescale:=",  _
  false)



  

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Solution Setup 
  
Set oModule = oDesign.GetModule("AnalysisSetup")

         
oModule.InsertSetup "HFIESetup", Array("NAME:Setup1", "MaximumPasses:=", 6, "MinimumPasses:=",  _
  1, "MinimumConvergedPasses:=", 1, "PercentRefinement:=", 30, "Enabled:=", true, "AdaptiveFreq:=",  _
  solution_freq & "GHz", "DoLambdaRefine:=", true, "UseDefaultLambdaTarget:=", true, "Target:=",  _
  0.25, "DoMaterialLambda:=", true, "MaxDeltaS:=", 0.02, "MaxDeltaE:=", 0.1, "UsingNumSolveSteps:=",  _
  0, "ConstantDelta:=", "0s", "NumberSolveSteps:=", 1)
  

           

     oModule.InsertSweep "Setup1", Array("NAME:Sweep1", "IsEnabled:=", true, "SetupType:=",  _
  "LinearCount", "StartValue:=", low_freq&"GHz", "StopValue:=", high_freq&"GHz", "Count:=",  _
  101, "Type:=", "Interpolating", "SaveFields:=", false, "InterpTolerance:=",  _
  0.5, "InterpMaxSolns:=", 250, "InterpMinSolns:=", 0, "InterpMinSubranges:=", 1, "ExtrapToDC:=",  _
  false)




  
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''Far field setup and Report Setup'  



Set oModule = oDesign.GetModule("RadField")
oModule.InsertFarFieldSphereSetup Array("NAME:infSphere", "UseCustomRadiationSurface:=",  _
false, "ThetaStart:=", "-180deg", "ThetaStop:=", "180deg", "ThetaStep:=", "5deg", "PhiStart:=",  _
"0deg", "PhiStop:=", "180deg", "PhiStep:=", "5deg", "UseLocalCS:=", false)


Set oModule = oDesign.GetModule("ReportSetup")
 
if NumberOfArms = 4 then   
oModule.CreateReport "Return Loss", "Solution Data", "Rectangular Plot",  _
"Setup1 : Sweep1", Array(), Array("Freq:=", Array("All")), Array("X Component:=", "Freq", "Y Component:=", Array("dB(S(1,1))","dB(S(2,2))")), Array()

oModule.CreateReport "Input Impedance", "Solution Data", "Smith Plot",_
"Setup1 : Sweep1", Array("Domain:=", "Sweep"), Array("Freq:=", Array("All")), _
Array("Polar Component:=", Array("S11","S22")),Array()

end if

if NumberOfArms = 2 then   
oModule.CreateReport "Return Loss", "Solution Data", "Rectangular Plot",  _
"Setup1 : Sweep1", Array(), Array("Freq:=", Array("All")), Array("X Component:=", "Freq", "Y Component:=", Array("dB(S(1,1))")), Array()

oModule.CreateReport "Input Impedance", "Solution Data", "Smith Plot",_
"Setup1 : Sweep1", Array("Domain:=", "Sweep"), Array("Freq:=", Array("All")), _
Array("Polar Component:=", Array("S11")),Array()

end if


oModule.CreateReport "ff_3D_GainRHCP", "Far Fields", "3D Polar Plot",  _
"Setup1 : LastAdaptive", Array("Context:=", "infSphere"), Array("Phi:=", Array( _
"All"), "Theta:=", Array("All")), Array("Phi Component:=",  _
"Phi", "Theta Component:=", "Theta", "Mag Component:=", Array("dB(GainRHCP)")), Array()

oModule.CreateReport "ff_3D_GainLHCP", "Far Fields", "3D Polar Plot",  _
"Setup1 : LastAdaptive", Array("Context:=", "infSphere"), Array("Phi:=", Array( _
"All"), "Theta:=", Array("All")), Array("Phi Component:=",  _
"Phi", "Theta Component:=", "Theta", "Mag Component:=", Array("dB(GainLHCP)")), Array()

oModule.CreateReport "ff_2D_GainRHCP", "Far Fields", "XY Plot",  _
"Setup1 : LastAdaptive", Array("Context:=", "infSphere"), Array("Theta:=", Array( _
"All"), "Phi:=", Array("0deg")), Array("X Component:=",  _
"Theta", "Y Component:=", Array("dB(GainRHCP)")), Array()

oModule.AddTraces "ff_2D_GainRHCP", "Setup1 : LastAdaptive", Array("Context:=",  _
"infSphere"), Array("Theta:=", Array("All"), "Phi:=", Array("90deg")_
), Array("X Component:=", "Theta", "Y Component:=", Array("dB(GainRHCP)")), Array()


oModule.CreateReport "ff_2D_GainLHCP", "Far Fields", "XY Plot",  _
"Setup1 : LastAdaptive", Array("Context:=", "infSphere"), Array("Theta:=", Array( _
"All"), "Phi:=", Array("0deg")), Array("X Component:=",  _
"Theta", "Y Component:=", Array("dB(GainLHCP)")), Array()

oModule.AddTraces "ff_2D_GainLHCP", "Setup1 : LastAdaptive", Array("Context:=",  _
"infSphere"), Array("Theta:=", Array("All"), "Phi:=", Array("90deg")_
), Array("X Component:=", "Theta", "Y Component:=", Array("dB(GainLHCP)")), Array()




Set oDesktop = oAnsoftApp.GetAppDesktop()
oDesktop.CloseAllWindows()
Set oModeler = oDesign.SetActiveEditor("3D Modeler")

oEditor.ShowWindow








Setlocale(locallang)


 
  