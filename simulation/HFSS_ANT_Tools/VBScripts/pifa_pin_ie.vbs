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
oProject.InsertDesign "HFSS-IE", "PIFA_Antenna_ADKv1", "DrivenModal", ""
Set oDesign = oProject.SetActiveDesign("PIFA_Antenna_ADKv1")



dim  solution_freq, patchX, patchY, subH, subX, subY, feedX, feedY, shortX, shortY, pin_rad, pin_outer, feed_length, units, israd

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


solution_freq = CDbl( args(0))

  dim freq_hz
  freq_hz=solution_freq*1e9

  dim WL 
  WL = light_speed/(freq_hz)

patchX = CDbl( args(1))
patchY = CDbl( args(2))
subH = ( args(3))
subX = CDbl( args(4))
subY = CDbl( args(5))
feedX = CDbl( args(6))
feedY = CDbl( args(7))
shortX = CDbl( args(8))
shortY = CDbl( args(9))
coax_inner_rad = CDbl( args(10))
coax_outer_rad = CDbl( args(11))
feed_length = CDbl( args(12))
units = args(13)

israd = args(14) ' 0 is for radiation surface, 1 is for PML

if patchX > subX then
  subX = patchX*1.1
end if
if patchY > subY then
  subY = patchY*1.1
end if



oDesign.ChangeProperty Array("NAME:AllTabs", Array("NAME:LocalVariableTab", Array("NAME:PropServers", "LocalVariables"), Array("NAME:NewProps", _
Array("NAME:--Patch Dimensions", "PropType:=", "SeparatorProp", "UserDef:=", true, "Value:=", ""), _
Array("NAME:patchX", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", patchX & units), _
Array("NAME:patchY", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", patchY & units), _
Array("NAME:--Shorting Pin", "PropType:=", "SeparatorProp", "UserDef:=", true, "Value:=", ""), _
Array("NAME:shortX", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", shortX & units), _
Array("NAME:shortY", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", shortY & units), _
Array("NAME:--Substrate Dimensions", "PropType:=", "SeparatorProp", "UserDef:=", true, "Value:=", ""), _
Array("NAME:subH", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", subH ), _
Array("NAME:subX", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", subX & units), _
Array("NAME:subY", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", subY & units), _
Array("NAME:--Feed", "PropType:=", "SeparatorProp", "UserDef:=", true, "Value:=", ""), _
Array("NAME:feedX", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", feedX & units), _
Array("NAME:feedY", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", feedY & units), _
Array("NAME:coax_inner_rad", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", coax_inner_rad & units), _
Array("NAME:coax_outer_rad", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", coax_outer_rad & units), _
Array("NAME:feed_length", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", feed_length & units))))

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'create substrate material

Material_Name = args(15)
Material_Name = Material_Name & "_ADK"
Permittivity = args(16)
TandD = args(17)

Set oDefinitionManager = oProject.GetDefinitionManager()
oDefinitionManager.AddMaterial Array("NAME:" & Material_Name, "CoordinateSystemType:=",  _
  "Cartesian", Array("NAME:AttachedData"), Array("NAME:ModifierData"), "permittivity:=", Permittivity , "dielectric_loss_tangent:=",TandD)


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Draws Substrate and bottom metalization

Set oEditor = oDesign.SetActiveEditor("3D Modeler")
oEditor.CreateBox Array("NAME:BoxParameters", "CoordinateSystemID:=", -1, "XPosition:=",  _
  "-subX/2", "YPosition:=", "-subY/2" , "ZPosition:=", "0mm", "XSize:=", "subX", "YSize:=",  _
  "subY", "ZSize:=", "subH"), Array("NAME:Attributes", "Name:=", "sub", "Flags:=",  _
  "", "Color:=", "(132 132 193)", "Transparency:=", 0.8, "PartCoordinateSystem:=",  _
  "Global", "MaterialName:=", Material_Name, "SolveInside:=", true)
  
  
oEditor.CreateRectangle Array("NAME:RectangleParameters", "CoordinateSystemID:=",  _
  -1, "IsCovered:=", true, "XStart:=", "-subX/2", "YStart:=", "-subY/2", "ZStart:=",  _
  "0mm", "Width:=", "subX", "Height:=", "subY", "WhichAxis:=", "Z"), Array("NAME:Attributes", "Name:=",  _
  "Ground", "Flags:=", "", "Color:=", "(255 128 65)", "Transparency:=",  _
  0.8, "PartCoordinateSystem:=", "Global", "MaterialName:=",  _
  "pec", "SolveInside:=", false)

Set oModule = oDesign.GetModule("BoundarySetup")  

oModule.AssignPerfectE Array("NAME:Ground", "Objects:=", Array("Ground"), "InfGroundPlane:=", false)


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Draws patch

Set oEditor = oDesign.SetActiveEditor("3D Modeler")


  oEditor.CreateRectangle Array("NAME:RectangleParameters", "CoordinateSystemID:=",  _
  -1, "IsCovered:=", true, "XStart:=", "-patchX/2", "YStart:=", "-patchY/2", "ZStart:=",  _
  "subH", "Width:=", "patchX", "Height:=", "patchY", "WhichAxis:=", "Z"), Array("NAME:Attributes", "Name:=",  _
  "patch", "Flags:=", "", "Color:=", "(255 128 65)", "Transparency:=",  _
  0.3, "PartCoordinateSystem:=", "Global", "MaterialName:=",  _
  "pec", "SolveInside:=", false)

  
Set oModule = oDesign.GetModule("BoundarySetup")  

oModule.AssignPerfectE Array("NAME:Patch", "Objects:=", Array("patch"), "InfGroundPlane:=", false)


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Shorting Pin

Set oEditor = oDesign.SetActiveEditor("3D Modeler")

oEditor.CreateCylinder Array("NAME:CylinderParameters", "CoordinateSystemID:=", -1, "XCenter:=",  _
  "shortX", "YCenter:=", "shortY", "ZCenter:=", "0mm", "Radius:=", "coax_inner_rad/3", "Height:=",  _
  "subH", "WhichAxis:=", "Z", "NumSides:=", "0"), Array("NAME:Attributes", "Name:=",  _
  "shorting_pin", "Flags:=", "", "Color:=", "(255 128 65)", "Transparency:=",  _
  0.3, "PartCoordinateSystem:=", "Global", "MaterialName:=",  _
  "pec", "SolveInside:=", false)


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Draws feed

Set oEditor = oDesign.SetActiveEditor("3D Modeler")

'feed cutout
oEditor.CreateCircle Array("NAME:CircleParameters", "CoordinateSystemID:=", -1, "IsCovered:=",  _
  true, "XCenter:=", "feedX", "YCenter:=", "feedY", "ZCenter:=", "0mm", "Radius:=",  _
  "coax_outer_rad", "WhichAxis:=", "Z", "NumSegments:=", "0"), Array("NAME:Attributes", "Name:=",  _
  "feed_cutout", "Flags:=", "", "Color:=", "(255 128 65)", "Transparency:=",  _
  0.8, "PartCoordinateSystem:=", "Global", "MaterialName:=",  _
  "vacuum", "SolveInside:=", true)

oEditor.Subtract Array("NAME:Selections", "Blank Parts:=", "Ground", "Tool Parts:=",  _
  "feed_cutout"), Array("NAME:SubtractParameters", "CoordinateSystemID:=", -1, "KeepOriginals:=",  _
  false)

'feed pin
oEditor.CreateCylinder Array("NAME:CylinderParameters", "CoordinateSystemID:=", -1, "XCenter:=",  _
  "feedX", "YCenter:=", "feedY", "ZCenter:=", "0mm", "Radius:=", "coax_inner_rad", "Height:=",  _
  "subH", "WhichAxis:=", "Z", "NumSides:=", "0"), Array("NAME:Attributes", "Name:=",  _
  "feed_pin", "Flags:=", "", "Color:=", "(255 128 65)", "Transparency:=",  _
  0.8, "PartCoordinateSystem:=", "Global", "MaterialName:=",  _
  "pec", "SolveInside:=", false)
  
'feed coax
oEditor.CreateCylinder Array("NAME:CylinderParameters", "CoordinateSystemID:=", -1, "XCenter:=",  _
  "feedX", "YCenter:=", "feedY", "ZCenter:=", "0mm", "Radius:=", "coax_inner_rad", "Height:=",  _
  "-feed_length", "WhichAxis:=", "Z", "NumSides:=", "0"), Array("NAME:Attributes", "Name:=",  _
  "coax_pin", "Flags:=", "", "Color:=", "(132 132 193)", "Transparency:=",  _
  0.8, "PartCoordinateSystem:=", "Global", "MaterialName:=",  _
  "pec", "SolveInside:=", false) 
  
oEditor.CreateCylinder Array("NAME:CylinderParameters", "CoordinateSystemID:=", -1, "XCenter:=",  _
  "feedX", "YCenter:=", "feedY", "ZCenter:=", "0mm", "Radius:=", "coax_outer_rad", "Height:=",  _
  "-feed_length", "WhichAxis:=", "Z", "NumSides:=", "0"), Array("NAME:Attributes", "Name:=",  _
  "coax", "Flags:=", "", "Color:=", "(132 132 193)", "Transparency:=",  _
  0.8, "PartCoordinateSystem:=", "Global", "MaterialName:=",  _
  "Teflon (tm)", "SolveInside:=", true)

'Dim faceid, faceid2, half_feed_length, outface_posX,outface_posY
'half_feed_length = feed_length/2
'outface_posX = feedX+coax_inner_rad
'outface_posY = feedY



faceid = oEditor.GetFaceByPosition(Array("NAME:FaceParameters","BodyName:=", "coax", "XPosition:=", "feedX+coax_outer_rad", "YPosition:=","feedY", "ZPosition:=", "-feed_length/2"))



Set oModule = oDesign.GetModule("BoundarySetup")
oModule.AssignPerfectE Array("NAME:coax_outer", "Faces:=", Array(faceid), "InfGroundPlane:=",  _
  false)

'port
Set oEditor = oDesign.SetActiveEditor("3D Modeler")
oEditor.CreateCircle Array("NAME:CircleParameters", "CoordinateSystemID:=", -1, "IsCovered:=",  _
  true, "XCenter:=", "feedX", "YCenter:=", "feedY", "ZCenter:=", "-feed_length", "Radius:=",  _
  "coax_outer_rad", "WhichAxis:=", "Z", "NumSegments:=", "0"), Array("NAME:Attributes", "Name:=",  _
  "port1", "Flags:=", "", "Color:=", "(132 132 193)", "Transparency:=",  _
  0.8, "PartCoordinateSystem:=", "Global", "MaterialName:=",  _
  "vacuum", "SolveInside:=", true)
  


faceid = oEditor.GetFaceByPosition(Array("NAME:FaceParameters","BodyName:=", "port1", "XPosition:=", "feedX", "YPosition:=","feedY", "ZPosition:=", "-feed_length"))

edgeid = oEditor.GetEdgeByPosition(Array("NAME:FaceParameters","BodyName:=", "coax_pin", "XPosition:=", "feedX+coax_inner_rad", "YPosition:=","feedY", "ZPosition:=", "-feed_length"))

Set oModule = oDesign.GetModule("BoundarySetup")

'auto identify ports seems to add too many terminals, manually adding here
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
start_freq = Round(.625*solution_freq,1)
stop_freq = Round(1.4*solution_freq,1)

           

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
  