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
oProject.InsertDesign "HFSS-IE", "Monopole_Antenna_ADKv1", "DrivenModal", ""
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
  

faceid = oEditor.GetFaceByPosition(Array("NAME:FaceParameters","BodyName:=", "port1", "XPosition:=", "0", "YPosition:=","0", "ZPosition:=", "0"))

edgeid =  oEditor.GetEdgeByPosition (Array("NAME:FaceParameters","BodyName:=", "port1", "XPosition:=", "0", "YPosition:=","0", "ZPosition:=", "port_gap"))



Set oModule = oDesign.GetModule("BoundarySetup")

'auto identify ports seems to add too many terminals, manually adding here
oModule.AssignLumpedPort Array("NAME:Port1", "Objects:=", Array("port1"), "RenormalizeAllTerminals:=",  _
  true, "TerminalIDList:=", Array())
oModule.AssignTerminal Array("NAME:port1_T1", "Edges:=", Array(edgeid), "ParentBndID:=",  _
  "Port1", "TerminalResistance:=", "25ohm")
   
  
Set oModule = oDesign.GetModule("AnalysisSetup")

         
oModule.InsertSetup "HFIESetup", Array("NAME:Setup1", "MaximumPasses:=", 6, "MinimumPasses:=",  _
  1, "MinimumConvergedPasses:=", 1, "PercentRefinement:=", 30, "Enabled:=", true, "AdaptiveFreq:=",  _
  center_freq & "GHz", "DoLambdaRefine:=", true, "UseDefaultLambdaTarget:=", true, "Target:=",  _
  0.25, "DoMaterialLambda:=", true, "MaxDeltaS:=", 0.02, "MaxDeltaE:=", 0.1, "UsingNumSolveSteps:=",  _
  0, "ConstantDelta:=", "0s", "NumberSolveSteps:=", 1)
  



dim start_freq
dim stop_freq
start_freq = .5*center_freq
stop_freq = 1.5*center_freq

           

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