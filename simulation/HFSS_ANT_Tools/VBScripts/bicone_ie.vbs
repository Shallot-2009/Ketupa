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
oProject.InsertDesign "HFSS-IE", "Bicone_Antenna_ADKv1", "DrivenModal", ""
Set oDesign = oProject.SetActiveDesign("Bicone_Antenna_ADKv1")

dim  solution_freq, Inner_Radius, Outer_Radius, Cone_Height, Port_Gap, Port_Width, units, israd

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
Port_Gap = CDbl( args(4))
Port_Width = CDbl( args(5))
units = args(6)
israd = args(7) ' 0 is for radiation surface, 1 is for PML

if Port_Width > Inner_Radius*2 then
  Port_Width = Inner_Radius*2
end if

oDesign.ChangeProperty Array("NAME:AllTabs", Array("NAME:LocalVariableTab", Array("NAME:PropServers", "LocalVariables"), Array("NAME:NewProps", _
Array("NAME:--Antenna Dimensions", "PropType:=", "SeparatorProp", "UserDef:=", true, "Value:=", ""), _
Array("NAME:Inner_Radius", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", Inner_Radius & units), _
Array("NAME:Outer_Radius", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", Outer_Radius & units), _
Array("NAME:Cone_Height", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", Cone_Height & units), _
Array("NAME:Port_Gap", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", Port_Gap & units), _
Array("NAME:Port_Width", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", Port_Width & units))))

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Draws antenna

Set oEditor = oDesign.SetActiveEditor("3D Modeler")

oEditor.CreatePolyline Array("NAME:PolylineParameters", "CoordinateSystemID:=", -1, "IsPolylineCovered:=",  _
  true, "IsPolylineClosed:=", false, Array("NAME:PolylinePoints", Array("NAME:PLPoint", "X:=",  _
  "Inner_Radius", "Y:=", "0mm", "Z:=", "Port_Gap/2"), Array("NAME:PLPoint", "X:=", "Outer_Radius", "Y:=",  _
  "0mm", "Z:=", "Port_Gap/2+Cone_Height")), Array("NAME:PolylineSegments", Array("NAME:PLSegment", "SegmentType:=",  _
  "Line", "StartIndex:=", 0, "NoOfPoints:=", 2))), Array("NAME:Attributes", "Name:=",  _
  "Cone", "Flags:=", "", "Color:=", "(255 128 65)", "Transparency:=",  _
  0.8, "PartCoordinateSystem:=", "Global", "MaterialName:=",  _
  "vacuum", "SolveInside:=", true)
  

oEditor.SweepAroundAxis Array("NAME:Selections", "Selections:=", "Cone", "NewPartsModelFlag:=",  _
  "Model"), Array("NAME:AxisSweepParameters", "CoordinateSystemID:=", -1, "DraftAngle:=",  _
  "0deg", "DraftType:=", "Round", "CheckFaceFaceIntersection:=", false, "SweepAxis:=",  _
  "Z", "SweepAngle:=", "360deg", "NumOfSegments:=", "0")
  
oEditor.CreateCircle Array("NAME:CircleParameters", "CoordinateSystemID:=", -1, "IsCovered:=",  _
  true, "XCenter:=", "0mm", "YCenter:=", "0mm", "ZCenter:=", "Port_Gap/2", "Radius:=",  _
  "Inner_Radius", "WhichAxis:=", "Z", "NumSegments:=", "0"), Array("NAME:Attributes", "Name:=",  _
  "Cone_End", "Flags:=", "", "Color:=", "(255 128 65)", "Transparency:=",  _
  0.800000011920929, "PartCoordinateSystem:=", "Global", "MaterialName:=",  _
  "vacuum", "SolveInside:=", true)  
  
oEditor.Unite Array("NAME:Selections", "Selections:=", "Cone,Cone_End"), Array("NAME:UniteParameters", "CoordinateSystemID:=",  _
  -1, "KeepOriginals:=", false)

Set oModule = oDesign.GetModule("BoundarySetup")  
oModule.AssignPerfectE Array("NAME:Cone", "Objects:=", Array("Cone"), "InfGroundPlane:=", false)

Set oEditor = oDesign.SetActiveEditor("3D Modeler")

oEditor.DuplicateAroundAxis Array("NAME:Selections", "Selections:=", "Cone", "NewPartsModelFlag:=",  _
  "Model"), Array("NAME:DuplicateAroundAxisParameters", "CoordinateSystemID:=", -1, "CreateNewObjects:=", _
  true, "WhichAxis:=", "X", "AngleStr:=", "180deg", "NumClones:=", "2"), Array("NAME:Options", "DuplicateBoundaries:=", true)

''''''''
'port setup

oEditor.CreateRectangle Array("NAME:RectangleParameters", "CoordinateSystemID:=",  _
  -1, "IsCovered:=", true, "XStart:=", "0", "YStart:=", "-Port_Width/2", "ZStart:=",  _
  "-Port_Gap/2", "Width:=", "Port_Width", "Height:=", "Port_Gap", "WhichAxis:=", "X"), Array("NAME:Attributes", "Name:=",  _
  "port1", "Flags:=", "", "Color:=", "(255 128 65)", "Transparency:=",  _
  0.8, "PartCoordinateSystem:=", "Global", "MaterialName:=",  _
  "vacuum", "SolveInside:=", true)

dim end_vector_pointX, end_vector_pointY
dim start_vector_pointX, start_vector_pointY

end_vector_pointX = 0
end_vector_pointX = end_vector_pointX & units
end_vector_pointY = 0
end_vector_pointY = end_vector_pointY & units
end_vector_pointZ = Port_Gap/2
end_vector_pointZ = end_vector_pointZ & units

start_vector_pointX = 0
start_vector_pointX = start_vector_pointX & units
start_vector_pointY = 0
start_vector_pointY = start_vector_pointY & units
start_vector_pointZ = -Port_Gap/2
start_vector_pointZ = start_vector_pointZ & units


 edgeid =  oEditor.GetEdgeByPosition (Array("NAME:FaceParameters","BodyName:=", "port1", "XPosition:=", "0", "YPosition:=","0", "ZPosition:=", "Port_Gap/2"))

Set oModule = oDesign.GetModule("BoundarySetup")

  
  


Set oModule = oDesign.GetModule("BoundarySetup")
oModule.AssignLumpedPort Array("NAME:Port1", "Objects:=", Array("port1"), "RenormalizeAllTerminals:=",  _
  true, "TerminalIDList:=", Array())
oModule.AssignTerminal Array("NAME:port1_T1", "Edges:=", Array(edgeid), "ParentBndID:=",  _
  "Port1", "TerminalResistance:=", "50ohm")  

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
  


dim start_freq
dim stop_freq
start_freq = Round(.5*solution_freq,2)
stop_freq = Round(1.5*solution_freq,2)

           

     oModule.InsertSweep "Setup1", Array("NAME:Sweep1", "IsEnabled:=", true, "SetupType:=",  _
  "LinearCount", "StartValue:=", start_freq&"GHz", "StopValue:=", stop_freq&"GHz", "Count:=",  _
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
   
oModule.CreateReport "Return Loss", "Solution Data", "Rectangular Plot",  _
"Setup1 : Sweep1", Array(), Array("Freq:=", Array("All")), Array("X Component:=", "Freq", "Y Component:=", Array("dB(S(1,1))")), Array()

oModule.CreateReport "Input Impedance", "Solution Data", "Smith Plot",_
"Setup1 : Sweep1", Array("Domain:=", "Sweep"), Array("Freq:=", Array("All")), _
Array("Polar Component:=", Array("S11")),Array()




oModule.CreateReport "ff_3D_GainTotal", "Far Fields", "3D Polar Plot",  _
"Setup1 : LastAdaptive", Array("Context:=", "infSphere"), Array("Phi:=", Array( _
"All"), "Theta:=", Array("All")), Array("Phi Component:=",  _
"Phi", "Theta Component:=", "Theta", "Mag Component:=", Array("dB(GainTotal)")), Array()

oModule.CreateReport "ff_2D_GainTotal", "Far Fields", "XY Plot",  _
"Setup1 : LastAdaptive", Array("Context:=", "infSphere"), Array("Theta:=", Array( _
"All"), "Phi:=", Array("0deg")), Array("X Component:=",  _
"Theta", "Y Component:=", Array("dB(GainTotal)")), Array()

oModule.AddTraces "ff_2D_GainTotal", "Setup1 : LastAdaptive", Array("Context:=",  _
"infSphere"), Array("Theta:=", Array("All"), "Phi:=", Array("90deg")_
), Array("X Component:=", "Theta", "Y Component:=", Array("dB(GainTotal)")), Array()




Set oDesktop = oAnsoftApp.GetAppDesktop()
oDesktop.CloseAllWindows()
Set oModeler = oDesign.SetActiveEditor("3D Modeler")

oEditor.ShowWindow








Setlocale(locallang) 
  
