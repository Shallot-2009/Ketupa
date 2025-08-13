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
oProject.InsertDesign "HFSS-IE", "Bow_Tie_Antenna_ADKv1", "DrivenModal", ""
Set oDesign = oProject.SetActiveDesign("Bow_Tie_Antenna_ADKv1")








dim  solution_freq, Inner_Width, Outer_Width, Arm_Length, Outer_Radius, Port_Gap_Width, subH, subX, subY, units, israd

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
Outer_Radius = CDbl( args(4))
Port_Gap_Width = CDbl( args(5))

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


oDesign.ChangeProperty Array("NAME:AllTabs", Array("NAME:LocalVariableTab", Array("NAME:PropServers", "LocalVariables"), Array("NAME:NewProps", _
Array("NAME:--Antenna Dimensions", "PropType:=", "SeparatorProp", "UserDef:=", true, "Value:=", ""), _
Array("NAME:Inner_width", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", Inner_width & units), _
Array("NAME:Outer_width", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", Outer_width & units), _
Array("NAME:Arm_Length", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", Arm_Length & units), _
Array("NAME:Outer_Radius", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", Outer_Radius & units), _
Array("NAME:Port_Gap_Width", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", Port_Gap_Width & units), _
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


oEditor.CreateCircle Array("NAME:CircleParameters", "IsCovered:=", true, "XCenter:=",  _
  "0", "YCenter:=", "if(Outer_Radius>=Outer_Width/2,arm_length-outer_width/2/tan(asin(outer_width/2/outer_radius))+port_gap_width/2 ,arm_length)", "ZCenter:=", "0cm", "Radius:=", "Outer_Radius", "WhichAxis:=",  _
  "Z", "NumSegments:=", "0"), Array("NAME:Attributes", "Name:=", "Circle1", "Flags:=",  _
  "", "Color:=", "(132 132 193)", "Transparency:=", 0.8, "PartCoordinateSystem:=",  _
  "Global", "MaterialValue:=", "" & Chr(34) & "vacuum" & Chr(34) & "", "SolveInside:=",  _
  true)

oEditor.CreateRelativeCS Array("NAME:RelativeCSParameters", "OriginX:=", "0cm", "OriginY:=",  _
  "arm_length+port_gap_width/2", "OriginZ:=", "0cm", "XAxisXvec:=", "1cm", "XAxisYvec:=", "0cm", "XAxisZvec:=",  _
  "0cm", "YAxisXvec:=", "0cm", "YAxisYvec:=", "1cm", "YAxisZvec:=", "0cm"), Array("NAME:Attributes", "Name:=",  _
  "RelativeCS4")
oEditor.Split Array("NAME:Selections", "Selections:=", "Circle1", "NewPartsModelFlag:=",  _
  "Model"), Array("NAME:SplitToParameters", "SplitPlane:=", "ZX", "WhichSide:=",  _
  "PositiveOnly", "SplitCrossingObjectsOnly:=", false, "DeleteInvalidObjects:=",  _
  true)



oEditor.Unite Array("NAME:Selections", "Selections:=", "Arm,Circle1"), Array("NAME:UniteParameters", "KeepOriginals:=",  _
  false)

 oEditor.SetWCS Array("NAME:SetWCS Parameter", "Working Coordinate System:=", "Global")

oEditor.DuplicateAroundAxis Array("NAME:Selections", "Selections:=", "Arm", "NewPartsModelFlag:=",  _
  "Model"), Array("NAME:DuplicateAroundAxisParameters", "CoordinateSystemID:=", -1, "CreateNewObjects:=", _
  true, "WhichAxis:=", "Z", "AngleStr:=", "180deg", "NumClones:=", "2"), Array("NAME:Options", "DuplicateBoundaries:=", true)


  Set oModule = oDesign.GetModule("BoundarySetup")  
oModule.AssignPerfectE Array("NAME:Arm", "Objects:=", Array("Arm"), "InfGroundPlane:=", false) 
oModule.AssignPerfectE Array("NAME:Arm_1", "Objects:=", Array("Arm_1"), "InfGroundPlane:=", false)




''''''''
'port setup

oEditor.CreateRectangle Array("NAME:RectangleParameters", "CoordinateSystemID:=",  _
  -1, "IsCovered:=", true, "XStart:=", "-Inner_Width/2", "YStart:=", "-Port_Gap_Width/2", "ZStart:=",  _
  "0mm", "Width:=", "Inner_Width", "Height:=", "Port_Gap_Width", "WhichAxis:=", "Z"), Array("NAME:Attributes", "Name:=",  _
  "port1", "Flags:=", "", "Color:=", "(255 128 65)", "Transparency:=",  _
  0.8, "PartCoordinateSystem:=", "Global", "MaterialName:=",  _
  "vacuum", "SolveInside:=", true)

dim end_vector_pointX, end_vector_pointY
dim start_vector_pointX, start_vector_pointY

end_vector_pointX = 0
end_vector_pointX = end_vector_pointX & units
end_vector_pointY = -Port_Gap_Width/2
end_vector_pointY = end_vector_pointY & units


start_vector_pointX = 0
start_vector_pointX = start_vector_pointX & units
start_vector_pointY = Port_Gap_Width/2
start_vector_pointY = start_vector_pointY & units

 edgeid =  oEditor.GetEdgeByPosition (Array("NAME:FaceParameters","BodyName:=", "port1", "XPosition:=", "0", "YPosition:=","-Port_Gap_Width/2", "ZPosition:=", "0"))

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
  
  
