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
oProject.InsertDesign "HFSS", "Helix_Antenna_ADKv1", "DrivenTerminal", ""
Set oDesign = oProject.SetActiveDesign("Helix_Antenna_ADKv1")

dim solution_freq, groundY, groundX, helixD, helixS, N, wireD, direction, port_height, Zin, units, israd

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
 port_height =   CDbl( args(8))

Zin = CDbl( args(9))
units = args(10)
israd = args(11) ' 0 is for radiation surface, 1 is for PML



dim light_speed, freq_hz, WL
light_speed = 299792458
freq_hz =solution_freq*1e9
WL = light_speed/(freq_hz)


if  direction = "Left" then
      direction = 0
      direction2 = false
else
  direction = 1
  direction2  = true
end if


oDesign.ChangeProperty Array("NAME:AllTabs", Array("NAME:LocalVariableTab", Array("NAME:PropServers", "LocalVariables"), Array("NAME:NewProps", _
Array("NAME:--Helix Dimensions", "PropType:=", "SeparatorProp", "UserDef:=", true, "Value:=", ""), _
Array("NAME:Helix_Diameter", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", helixD & units), _
Array("NAME:Helix_Spacing", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", helixS & units), _
Array("NAME:Number_of_Turns", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", N), _
Array("NAME:Wire_Diameter", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", wireD & units), _   
Array("NAME:Direction", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", direction), _
Array("NAME:--Feed Dimensions", "PropType:=", "SeparatorProp", "UserDef:=", true, "Value:=", ""), _
Array("NAME:Port_Height", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=",  port_height & units), _
Array("NAME:--Ground Plane Dimensions", "PropType:=", "SeparatorProp", "UserDef:=", true, "Value:=", ""), _
Array("NAME:Ground_Y", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", groundY & units), _
Array("NAME:Ground_X", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", groundX & units))))



  
Set oEditor = oDesign.SetActiveEditor("3D Modeler")


'draw helix'
  oEditor.CreateCircle Array("NAME:CircleParameters", "IsCovered:=", true, "XCenter:=",  _
  "Helix_Diameter/2", "YCenter:=", "0cm", "ZCenter:=", "0cm", "Radius:=",  _
  "Wire_Diameter/2", "WhichAxis:=", "Y", "NumSegments:=", "0"), Array("NAME:Attributes", "Name:=",  _
  "Helix1", "Flags:=", "", "Color:=", "(132 132 193)", "Transparency:=",  _
  0.8, "PartCoordinateSystem:=", "Global", "MaterialValue:=",  _
  "" & Chr(34) & "pec" & Chr(34) & "", "SolveInside:=", false)
oEditor.CreateHelix Array("NAME:Selections", "Selections:=", "Helix1", "NewPartsModelFlag:=",  _
  "Model"), Array("NAME:HelixParameters", "XCenter:=", "0cm", "YCenter:=", "0cm", "ZCenter:=",  _
  "0cm", "XStartDir:=", "0cm", "YStartDir:=", "0cm", "ZStartDir:=", "1", "NumThread:=",  _
  "Number_of_Turns", "RightHand:=", direction2, "RadiusIncrement:=", "0cm", "Thread:=", "Helix_Spacing")







'oEditor.CreateUserDefinedPart Array("NAME:UserDefinedPrimitiveParameters", "DllName:=",  _
''  "SegmentedHelix/PolygonHelix", "Version:=", "1.0", "NoOfParameters:=", 8, "Library:=",  _
''  "syslib", Array("NAME:ParamVector", Array("NAME:Pair", "Name:=", "PolygonSegments", "Value:=",  _
''  "8"), Array("NAME:Pair", "Name:=", "PolygonRadius", "Value:=", "Wire_Diameter/2"), Array("NAME:Pair", "Name:=",  _
''  "StartHelixRadius", "Value:=", "Helix_Diameter/2"), Array("NAME:Pair", "Name:=", "RadiusChange", "Value:=",  _
''  "0mm"), Array("NAME:Pair", "Name:=", "Pitch", "Value:=", "Helix_Spacing"), Array("NAME:Pair", "Name:=",  _
''  "Turns", "Value:=", "Number_of_Turns"), Array("NAME:Pair", "Name:=", "SegmentsPerTurn", "Value:=",  _
''  "16"), Array("NAME:Pair", "Name:=", "RightHanded", "Value:=", "Direction"))), Array("NAME:Attributes", "Name:=",  _
''  "PolygonHelix1", "Flags:=", "", "Color:=", "(132 132 193)", "Transparency:=",  _
''  0.8, "PartCoordinateSystem:=", "Global", "MaterialValue:=",  _
''  "" & Chr(34) & "pec" & Chr(34) & "", "SolveInside:=", false)


 'draw ground plane
 oEditor.CreateRectangle Array("NAME:RectangleParameters", "CoordinateSystemID:=",  _
  -1, "IsCovered:=", true, "XStart:=", "-Ground_X/2", "YStart:=", "-Ground_Y/2", "ZStart:=",  _
  "-2*Port_Height-Wire_Diameter/2", "Width:=", "Ground_X", "Height:=", "Ground_Y", "WhichAxis:=", "Z"), Array("NAME:Attributes", "Name:=",  _
  "Ground", "Flags:=", "", "Color:=", "(255 128 65)", "Transparency:=",  _
  0.8, "PartCoordinateSystem:=", "Global", "MaterialName:=",  _
  "pec", "SolveInside:=", false)








'feed pin
oEditor.CreateCylinder Array("NAME:CylinderParameters", "CoordinateSystemID:=", -1, "XCenter:=",  _
  "Helix_Diameter/2", "YCenter:=", "0", "ZCenter:=", "-Port_Height-Wire_Diameter/2", "Radius:=", "Wire_Diameter/2", "Height:=",  _
  "Port_Height+Wire_Diameter/2", "WhichAxis:=", "Z", "NumSides:=", "0"), Array("NAME:Attributes", "Name:=",  _
  "feed_pin", "Flags:=", "", "Color:=", "(255 128 65)", "Transparency:=",  _
  0.8, "PartCoordinateSystem:=", "Global", "MaterialName:=",  _
  "pec", "SolveInside:=", false)

oEditor.Unite Array("NAME:Selections", "Selections:=", "feed_pin,Helix1"), Array("NAME:UniteParameters", "KeepOriginals:=",  _
  false)

  
'port

   oEditor.CreateRectangle Array("NAME:RectangleParameters", "CoordinateSystemID:=",  _
  -1, "IsCovered:=", true, "XStart:=", "Helix_Diameter/2-Wire_Diameter/2", "YStart:=", "0", "ZStart:=",  _
  "-Port_Height-Wire_Diameter/2", "Width:=", "-Port_Height", "Height:=", "Wire_Diameter", "WhichAxis:=", "Y"), Array("NAME:Attributes", "Name:=",  _
  "port1", "Flags:=", "", "Color:=", "(255 128 65)", "Transparency:=",  _
  0.8, "PartCoordinateSystem:=", "Global", "MaterialName:=",  _
  "pec", "SolveInside:=", false)



 
Set oModule = oDesign.GetModule("BoundarySetup")  
oModule.AssignPerfectE Array("NAME:Ground", "Objects:=", Array("Ground"), "InfGroundPlane:=", false)




'port



oDesign.SetSolutionType "DrivenTerminal"

faceid = oEditor.GetFaceByPosition(Array("NAME:FaceParameters","BodyName:=", "port1", "XPosition:=", "Helix_Diameter/2", "YPosition:=","0", "ZPosition:=", "-Port_Height-Wire_Diameter/2-Port_Height/2"))



Set oModule = oDesign.GetModule("BoundarySetup")


 oModule.AutoIdentifyPorts Array("NAME:Faces", faceid), false, Array("NAME:ReferenceConductors",  _
  "Ground"), "1", true


oDesign.ChangeProperty Array("NAME:AllTabs", Array("NAME:HfssTab", Array("NAME:PropServers",  _
  "BoundarySetup:port1_T1"), Array("NAME:ChangedProps", Array("NAME:Terminal Renormalizing Impedance", "Value:=",  _
  Zin & "ohm"))))









  oEditor.DuplicateAroundAxis Array("NAME:Selections", "Selections:=", "feed_pin,port1", "NewPartsModelFlag:=",  _
  "Model"), Array("NAME:DuplicateAroundAxisParameters", "CreateNewObjects:=", true, "WhichAxis:=",  _
  "Z", "AngleStr:=", "90deg", "NumClones:=", "4"), Array("NAME:Options", "DuplicateAssignments:=",  _
  true)





Set oModule = oDesign.GetModule("Solutions")
oModule.EditSources "NoIncidentWave", Array("NAME:Names", "1", "2", "3", "4"), Array("NAME:Terminals",  _
  1, 1, 1, 1), Array("NAME:Magnitudes", "1", "1", "1", "1"), Array("NAME:Phases", "0deg",  _
  "90deg", "180deg", "270deg"), Array("NAME:Terminated", false, false, false, false), Array("NAME:Impedances"),  _
  false



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Solution Setup 
  



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
start_freq = .5*solution_freq
stop_freq = 1.5*solution_freq

 oModule.InsertFrequencySweep "Setup1", Array("NAME:Sweep1", "IsEnabled:=", true, "SetupType:=",  _
  "LinearCount", "StartValue:=", start_freq&"GHz", "StopValue:=", stop_freq&"GHz", "Count:=", 200, "Type:=",  _
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

extent_x_pos = "Ground_X/2"
extent_x_neg = "-Ground_X/2"

extent_y_pos = "Ground_Y/2"
extent_y_neg = "-Ground_Y/2"

extent_z_pos = "Helix_Spacing*Number_of_Turns"
extent_z_neg = "-2*Port_Height"

if israd = "ABC" then
 israd = "Rad"

end if


dim mycommand
' full command which invokes wscript to run desired VBScript and passes
mycommand = "wscript.exe " & """" & "./VBScripts/rad_creation.vbs" & """" &  " " & extent_x_pos & " " & extent_x_neg & " " & extent_y_pos & " " & extent_y_neg & " " & extent_z_pos & " " & extent_z_neg & " " & units & " " & israd& " " & 0
'msgbox(mycommand)
' run the desired VBScript
Set WSHShell = CreateObject("WScript.Shell")
WSHShell.Run mycommand, , True



Setlocale(locallang)