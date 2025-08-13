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
oProject.InsertDesign "HFSS-IE", "Helix_Antenna_ADKv1", "DrivenTerminal", ""
Set oDesign = oProject.SetActiveDesign("Helix_Antenna_ADKv1")

dim solution_freq, groundY, groundX, helixD, helixS, N, wireD, direction, coaxouter, coaxinner, pinL,pinD, Zin, units, israd

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
groundY = CDbl( args(1))
groundX = CDbl( args(2))
helixD = CDbl( args(3))
helixS  = CDbl( args(4))
N = CDbl( args(5))
wireD = CDbl( args(6))
direction =  cstr(args(7))
coaxouter = CDbl( args(8))
coaxinner = CDbl( args(9))
pinL = CDbl( args(10))
pinD = CDbl( args(11))

Zin = CDbl( args(12))
units = args(13)
israd = args(14) ' 0 is for radiation surface, 1 is for PML



dim light_speed, freq_hz, WL
light_speed = 299792458
freq_hz =solution_freq*1e9
WL = light_speed/(freq_hz)


if  direction = "Left" then
      direction = 0
else
  direction = 1
end if


oDesign.ChangeProperty Array("NAME:AllTabs", Array("NAME:LocalVariableTab", Array("NAME:PropServers", "LocalVariables"), Array("NAME:NewProps", _
Array("NAME:--Helix Dimensions", "PropType:=", "SeparatorProp", "UserDef:=", true, "Value:=", ""), _
Array("NAME:Helix_Diameter", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", helixD & units), _
Array("NAME:Helix_Spacing", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", helixS & units), _
Array("NAME:Number_of_Turns", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", N), _
Array("NAME:Wire_Diameter", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", wireD & units), _   
Array("NAME:Direction", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", direction), _
Array("NAME:--Feed Dimensions", "PropType:=", "SeparatorProp", "UserDef:=", true, "Value:=", ""), _
Array("NAME:Coax_Inner_Diameter", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", coaxinner & units), _
Array("NAME:Coax_Outer_Diameter", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", coaxouter & units), _
Array("NAME:Coax_Length", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", helixD & units), _
Array("NAME:Pin_Height", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", pinL & units), _
Array("NAME:Pin_Diameter", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", pinD & units), _
Array("NAME:--Ground Plane Dimensions", "PropType:=", "SeparatorProp", "UserDef:=", true, "Value:=", ""), _
Array("NAME:Ground_Y", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", groundY & units), _
Array("NAME:Ground_X", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", groundX & units))))


  
  
Set oEditor = oDesign.SetActiveEditor("3D Modeler")


'draw helix'
oEditor.CreateUserDefinedPart Array("NAME:UserDefinedPrimitiveParameters", "DllName:=",  _
  "SegmentedHelix/PolygonHelix", "Version:=", "1.0", "NoOfParameters:=", 8, "Library:=",  _
  "syslib", Array("NAME:ParamVector", Array("NAME:Pair", "Name:=", "PolygonSegments", "Value:=",  _
  "8"), Array("NAME:Pair", "Name:=", "PolygonRadius", "Value:=", "Wire_Diameter/2"), Array("NAME:Pair", "Name:=",  _
  "StartHelixRadius", "Value:=", "Helix_Diameter/2"), Array("NAME:Pair", "Name:=", "RadiusChange", "Value:=",  _
  "0mm"), Array("NAME:Pair", "Name:=", "Pitch", "Value:=", "Helix_Spacing"), Array("NAME:Pair", "Name:=",  _
  "Turns", "Value:=", "Number_of_Turns"), Array("NAME:Pair", "Name:=", "SegmentsPerTurn", "Value:=",  _
  "16"), Array("NAME:Pair", "Name:=", "RightHanded", "Value:=", "Direction"))), Array("NAME:Attributes", "Name:=",  _
  "PolygonHelix1", "Flags:=", "", "Color:=", "(132 132 193)", "Transparency:=",  _
  0.8, "PartCoordinateSystem:=", "Global", "MaterialValue:=",  _
  "" & Chr(34) & "pec" & Chr(34) & "", "SolveInside:=", false)


 'draw ground plane
 oEditor.CreateRectangle Array("NAME:RectangleParameters", "CoordinateSystemID:=",  _
  -1, "IsCovered:=", true, "XStart:=", "-Ground_X/2", "YStart:=", "-Ground_Y/2", "ZStart:=",  _
  "-Pin_Height-Wire_Diameter/2", "Width:=", "Ground_X", "Height:=", "Ground_Y", "WhichAxis:=", "Z"), Array("NAME:Attributes", "Name:=",  _
  "Ground", "Flags:=", "", "Color:=", "(255 128 65)", "Transparency:=",  _
  0.8, "PartCoordinateSystem:=", "Global", "MaterialName:=",  _
  "pec", "SolveInside:=", false)


 oEditor.CreateCircle Array("NAME:CircleParameters", "CoordinateSystemID:=", -1, "IsCovered:=",  _
  true, "XCenter:=", "Helix_Diameter/2", "YCenter:=", "0", "ZCenter:=", "-Pin_Height-Wire_Diameter/2", "Radius:=",  _
  "Coax_Outer_Diameter/2", "WhichAxis:=", "Z", "NumSegments:=", "0"), Array("NAME:Attributes", "Name:=",  _
  "feed_cutout", "Flags:=", "", "Color:=", "(255 128 65)", "Transparency:=",  _
  0.8, "PartCoordinateSystem:=", "Global", "MaterialName:=",  _
  "vacuum", "SolveInside:=", true)


 oEditor.Subtract Array("NAME:Selections", "Blank Parts:=", "Ground", "Tool Parts:=",  _
  "feed_cutout"), Array("NAME:SubtractParameters", "CoordinateSystemID:=", -1, "KeepOriginals:=",  _
  false)



'feed pin
oEditor.CreateCylinder Array("NAME:CylinderParameters", "CoordinateSystemID:=", -1, "XCenter:=",  _
  "Helix_Diameter/2", "YCenter:=", "0", "ZCenter:=", "-Pin_Height-Wire_Diameter/2", "Radius:=", "Pin_Diameter/2", "Height:=",  _
  "Pin_Height+Wire_Diameter/2", "WhichAxis:=", "Z", "NumSides:=", "0"), Array("NAME:Attributes", "Name:=",  _
  "feed_pin", "Flags:=", "", "Color:=", "(255 128 65)", "Transparency:=",  _
  0.8, "PartCoordinateSystem:=", "Global", "MaterialName:=",  _
  "pec", "SolveInside:=", false)

oEditor.Unite Array("NAME:Selections", "Selections:=", "feed_pin,PolygonHelix1"), Array("NAME:UniteParameters", "KeepOriginals:=",  _
  false)

  
'feed coax
oEditor.CreateCylinder Array("NAME:CylinderParameters", "CoordinateSystemID:=", -1, "XCenter:=",  _
  "Helix_Diameter/2", "YCenter:=", "0", "ZCenter:=", "-Pin_Height-Wire_Diameter/2", "Radius:=", "Coax_Inner_Diameter/2", "Height:=",  _
  "-Coax_Length", "WhichAxis:=", "Z", "NumSides:=", "0"), Array("NAME:Attributes", "Name:=",  _
  "coax_pin", "Flags:=", "", "Color:=", "(132 132 193)", "Transparency:=",  _
  0.8, "PartCoordinateSystem:=", "Global", "MaterialName:=",  _
  "pec", "SolveInside:=", false) 
  
oEditor.CreateCylinder Array("NAME:CylinderParameters", "CoordinateSystemID:=", -1, "XCenter:=",  _
  "Helix_Diameter/2", "YCenter:=", "0", "ZCenter:=", "-Pin_Height-Wire_Diameter/2", "Radius:=", "Coax_Outer_Diameter/2", "Height:=",  _
  "-Coax_Length", "WhichAxis:=", "Z", "NumSides:=", "0"), Array("NAME:Attributes", "Name:=",  _
  "coax", "Flags:=", "", "Color:=", "(132 132 193)", "Transparency:=",  _
  0.8, "PartCoordinateSystem:=", "Global", "MaterialName:=",  _
  "Teflon (tm)", "SolveInside:=", true)


faceid = oEditor.GetFaceByPosition(Array("NAME:FaceParameters","BodyName:=", "coax", "XPosition:=", "Helix_Diameter/2+Coax_Outer_Diameter/2", "YPosition:=","0", "ZPosition:=", "-Pin_Height-Wire_Diameter/2-Coax_Length/2"))




 
Set oModule = oDesign.GetModule("BoundarySetup")  
oModule.AssignPerfectE Array("NAME:Ground", "Objects:=", Array("Ground"), "InfGroundPlane:=", false)
oModule.AssignPerfectE Array("NAME:coax_outer", "Faces:=", Array(faceid), "InfGroundPlane:=",  _
  false)



'port
Set oEditor = oDesign.SetActiveEditor("3D Modeler")
oEditor.CreateCircle Array("NAME:CircleParameters", "CoordinateSystemID:=", -1, "IsCovered:=",  _
  true, "XCenter:=", "Helix_Diameter/2", "YCenter:=", "0", "ZCenter:=", "-Pin_Height-Wire_Diameter/2-Coax_Length", "Radius:=",  _
  "Coax_Outer_Diameter/2", "WhichAxis:=", "Z", "NumSegments:=", "0"), Array("NAME:Attributes", "Name:=",  _
  "port1", "Flags:=", "", "Color:=", "(132 132 193)", "Transparency:=",  _
  0.8, "PartCoordinateSystem:=", "Global", "MaterialName:=",  _
  "vacuum", "SolveInside:=", true)
  




faceid = oEditor.GetFaceByPosition(Array("NAME:FaceParameters","BodyName:=", "port1", "XPosition:=", "Helix_Diameter/2", "YPosition:=","0", "ZPosition:=", "-Pin_Height-Wire_Diameter/2-Coax_Length"))









  edgeid = oEditor.GetEdgeByPosition(Array("NAME:FaceParameters","BodyName:=", "coax_pin", "XPosition:=", "Helix_Diameter/2+Coax_Inner_Diameter/2", "YPosition:=","0", "ZPosition:=", "-Pin_Height-Wire_Diameter/2-Coax_Length"))
   



Set oModule = oDesign.GetModule("BoundarySetup")

'auto identify ports seems to add too many terminals, manually adding here
oModule.AssignLumpedPort Array("NAME:Port1", "Objects:=", Array("port1"), "RenormalizeAllTerminals:=",  _
  true, "TerminalIDList:=", Array())
oModule.AssignTerminal Array("NAME:port1_T1", "Edges:=", Array(edgeid), "ParentBndID:=",  _
  "Port1", "TerminalResistance:=", Zin & "ohm")
   

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
start_freq = .75*solution_freq
stop_freq = 1.5*solution_freq

           

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
  