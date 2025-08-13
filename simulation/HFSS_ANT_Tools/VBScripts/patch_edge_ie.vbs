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
oProject.InsertDesign "HFSS-IE", "Patch_Antenna_ADKv1", "DrivenModal", ""
Set oDesign = oProject.SetActiveDesign("Patch_Antenna_ADKv1")


dim circle_or_rect, solution_freq, patchX, patchY, subH, subX, subY, EdgeFeedWidth, EdgeFeedLength, FeedWidth, FeedLength, units, israd

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


circle_or_rect = CInt( args(0))  '1 is rectangular, 0 is circular

dim light_speed
light_speed = 299792458


solution_freq = CDbl( args(1))

  dim freq_hz
  freq_hz=solution_freq*1e9

  dim WL 
  WL = light_speed/(freq_hz)



  


patchX = CDbl( args(2))
patchY = CDbl( args(3))
subH = args(4)
subX = CDbl( args(5))
subY = CDbl( args(6))
EdgeFeedWidth = CDbl( args(7))
EdgeFeedLength = CDbl( args(8))


FeedWidth = CDbl( args(9) )
FeedLength = CDbl( args(10))

units = args(11)

israd = args(12) ' 0 is for radiation surface, 1 is for PML

if patchX > subX then
  subX = patchX*1.1
end if
if patchY > subY then
  subY = patchY*1.1
end if


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
Array("NAME:--Substrate Dimensions", "PropType:=", "SeparatorProp", "UserDef:=", true, "Value:=", ""), _
Array("NAME:subH", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", subH ), _
Array("NAME:subX", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", subX & units), _
Array("NAME:subY", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", subY & units), _
Array("NAME:--Feed", "PropType:=", "SeparatorProp", "UserDef:=", true, "Value:=", ""), _
Array("NAME:EdgeFeedWidth", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", EdgeFeedWidth & units), _
Array("NAME:EdgeFeedLength", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", EdgeFeedLength & units), _
Array("NAME:FeedWidth", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", FeedWidth & units), _
Array("NAME:FeedLength", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", FeedLength & units))))

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'create substrate material

Material_Name = args(13)
Material_Name = Material_Name & "_ADK"
Permittivity = args(14)
TandD = args(15)

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

if circle_or_rect = 1 then
  oEditor.CreateRectangle Array("NAME:RectangleParameters", "CoordinateSystemID:=",  _
  -1, "IsCovered:=", true, "XStart:=", "-patchX/2", "YStart:=", "-patchY/2", "ZStart:=",  _
  "subH", "Width:=", "patchX", "Height:=", "patchY", "WhichAxis:=", "Z"), Array("NAME:Attributes", "Name:=",  _
  "patch", "Flags:=", "", "Color:=", "(255 128 65)", "Transparency:=",  _
  0.3, "PartCoordinateSystem:=", "Global", "MaterialName:=",  _
  "pec", "SolveInside:=", false)
else
  
  oEditor.CreateEllipse Array("NAME:EllipseParameters", "CoordinateSystemID:=", -1, "IsCovered:=",  _
  true, "XCenter:=", "0mm", "YCenter:=", "0mm", "ZCenter:=", "subH", "MajRadius:=",  _
  "patchX/2", "Ratio:=", "patchY/patchX", "WhichAxis:=", "Z", "NumSegments:=", "0"), Array("NAME:Attributes", "Name:=",  _
  "patch", "Flags:=", "", "Color:=", "(255 128 65)", "Transparency:=",  _
  0.3, "PartCoordinateSystem:=", "Global", "MaterialName:=",  _
  "pec", "SolveInside:=", false)
  
 
end if 

  oEditor.CreateRectangle Array("NAME:RectangleParameters", "CoordinateSystemID:=",  _
  -1, "IsCovered:=", true, "XStart:=", "-EdgeFeedWidth/2", "YStart:=", "0", "ZStart:=",  _
  "subH", "Width:=", "EdgeFeedWidth", "Height:=", "patchY/2+EdgeFeedLength", "WhichAxis:=", "Z"), Array("NAME:Attributes", "Name:=",  _
  "EdgeFeed", "Flags:=", "", "Color:=", "(255 128 65)", "Transparency:=",  _
  0.3, "PartCoordinateSystem:=", "Global", "MaterialName:=",  _
  "pec", "SolveInside:=", false) 

  oEditor.CreateRectangle Array("NAME:RectangleParameters", "CoordinateSystemID:=",  _
  -1, "IsCovered:=", true, "XStart:=", "-FeedWidth/2", "YStart:=", "patchY/2+EdgeFeedLength", "ZStart:=",  _
  "subH", "Width:=", "FeedWidth", "Height:=", "FeedLength", "WhichAxis:=", "Z"), Array("NAME:Attributes", "Name:=",  _
  "Feed", "Flags:=", "", "Color:=", "(255 128 65)", "Transparency:=",  _
  0.3, "PartCoordinateSystem:=", "Global", "MaterialName:=",  _
  "pec", "SolveInside:=", false)   

oEditor.Unite Array("NAME:Selections", "Selections:=", "patch,EdgeFeed,Feed"), Array("NAME:UniteParameters", "KeepOriginals:=",  _
  false)
  
  
oEditor.CreateRectangle Array("NAME:RectangleParameters", "IsCovered:=", true, "XStart:=",  _
  "-FeedWidth/2", "YStart:=", "patchY/2+EdgeFeedLength+FeedLength", "ZStart:=", "0", "Width:=", "subH", "Height:=",  _
  "FeedWidth", "WhichAxis:=", "Y"), Array("NAME:Attributes", "Name:=", "port1", "Flags:=",  _
  "", "Color:=", "(132 132 193)", "Transparency:=", 0.8, "PartCoordinateSystem:=",  _
  "Global", "MaterialValue:=", "" & Chr(34) & "vacuum" & Chr(34) & "", "SolveInside:=",  _
  true)  
  
  
Set oModule = oDesign.GetModule("BoundarySetup")  

oModule.AssignPerfectE Array("NAME:Patch", "Objects:=", Array("patch"), "InfGroundPlane:=", false)




faceid = oEditor.GetFaceByPosition(Array("NAME:FaceParameters","BodyName:=", "port1", "XPosition:=", "0", "YPosition:=","patchY/2+EdgeFeedLength+FeedLength", "ZPosition:=", "subH/2"))


Set oModule = oDesign.GetModule("BoundarySetup")



oModule.AutoIdentifyPorts Array("NAME:Faces", faceid), Array("NAME:ReferenceConductors",  _
  "Ground"), "LumpPort1", true
  
  
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
start_freq = .5*solution_freq
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
  