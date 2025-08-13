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
oProject.InsertDesign "HFSS", "Slot_Antenna_ADKv1", "DrivenModal", ""
Set oDesign = oProject.SetActiveDesign("Slot_Antenna_ADKv1")








dim center_freq, Antenna_Width, feed_offset, slot_length, subH, subx, subY, units, israd

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




Antenna_Width = CDbl( args(1))
feed_offset = CDbl( args(2))
slot_length = CDbl( args(3))


subH =  args(4)
subX = CDbl( args(5))
subY = CDbl( args(6))


units = args(7)


israd = args(8) ' 0 is for radiation surface, 1 is for PML





if slot_length > subY then
  subY = 2*(slot_length)*1.1
end if



oDesign.ChangeProperty Array("NAME:AllTabs", Array("NAME:LocalVariableTab", Array("NAME:PropServers", "LocalVariables"), Array("NAME:NewProps", _
Array("NAME:--Antenna Dimensions", "PropType:=", "SeparatorProp", "UserDef:=", true, "Value:=", ""), _
Array("NAME:Antenna_Width", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", Antenna_Width & units), _
Array("NAME:feed_offset", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", feed_offset & units), _
Array("NAME:slot_length", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", slot_length & units), _
Array("NAME:--Substrate Dimensions", "PropType:=", "SeparatorProp", "UserDef:=", true, "Value:=", ""), _
Array("NAME:subH", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", subH), _
Array("NAME:subX", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", subX & units), _
Array("NAME:subY", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", subY & units))))

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'create substrate material

Material_Name = args(9)
Material_Name = Material_Name & "_ADK"
Permittivity = args(10)
TandD = args(11)

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
  
Set oModule = oDesign.GetModule("BoundarySetup")  



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Draws antenna

Set oEditor = oDesign.SetActiveEditor("3D Modeler")

oEditor.CreateRectangle Array("NAME:RectangleParameters", "CoordinateSystemID:=",  _
  -1, "IsCovered:=", true, "XStart:=", "-subX/2", "YStart:=", "-subY/2", "ZStart:=",  _
  "0mm", "Width:=", "subX", "Height:=", "subY", "WhichAxis:=", "Z"), Array("NAME:Attributes", "Name:=",  _
  "ground", "Flags:=", "", "Color:=", "(255 128 65)", "Transparency:=",  _
  0.3, "PartCoordinateSystem:=", "Global", "MaterialName:=",  _
  "pec", "SolveInside:=", false)
  
oEditor.CreateRectangle Array("NAME:RectangleParameters", "CoordinateSystemID:=",  _
  -1, "IsCovered:=", true, "XStart:=", "-Antenna_Width/2", "YStart:=", "-slot_length/2", "ZStart:=",  _
  "0mm", "Width:=", "Antenna_Width", "Height:=", "slot_length", "WhichAxis:=", "Z"), Array("NAME:Attributes", "Name:=",  _
  "slot", "Flags:=", "", "Color:=", "(255 128 65)", "Transparency:=",  _
  0.3, "PartCoordinateSystem:=", "Global", "MaterialName:=",  _
  "pec", "SolveInside:=", false)
  
  
 oEditor.Subtract Array("NAME:Selections", "Blank Parts:=", "ground", "Tool Parts:=",  _
  "slot"), Array("NAME:SubtractParameters", "KeepOriginals:=", false)




  Set oModule = oDesign.GetModule("BoundarySetup")  
oModule.AssignPerfectE Array("NAME:ground", "Objects:=", Array("ground"), "InfGroundPlane:=", false) 


''''''''
'port setup

oEditor.CreateRectangle Array("NAME:RectangleParameters", "CoordinateSystemID:=",  _
  -1, "IsCovered:=", true, "XStart:=", "-Antenna_Width/2", "YStart:=", "feed_offset-Antenna_Width/2", "ZStart:=",  _
  "0mm", "Width:=", "Antenna_Width", "Height:=", "Antenna_Width", "WhichAxis:=", "Z"), Array("NAME:Attributes", "Name:=",  _
  "port1", "Flags:=", "", "Color:=", "(255 128 65)", "Transparency:=",  _
  0.8, "PartCoordinateSystem:=", "Global", "MaterialName:=",  _
  "pec", "SolveInside:=", false)

dim end_vector_pointX, end_vector_pointY
dim start_vector_pointX, start_vector_pointY

end_vector_pointX = -Antenna_Width/2
end_vector_pointX = end_vector_pointX & units
end_vector_pointY = feed_offset
end_vector_pointY = end_vector_pointY & units


start_vector_pointX = Antenna_Width/2
start_vector_pointX = start_vector_pointX & units
start_vector_pointY = feed_offset
start_vector_pointY = start_vector_pointY & units


oDesign.SetSolutionType "DrivenModal"
Set oModule = oDesign.GetModule("BoundarySetup")
oModule.AssignLumpedPort Array("NAME:p1", "Objects:=", Array("port1"), Array("NAME:Modes", Array("NAME:Mode1", "ModeNum:=",  _
  1, "UseIntLine:=", true, Array("NAME:IntLine", "Start:=", Array(start_vector_pointX, start_vector_pointY,  _
  "0mm"), "End:=", Array(end_vector_pointX, end_vector_pointY, "0mm")), "CharImp:=", "Zpi", "RenormImp:=",  _
  "50ohm")), "FullResistance:=", "50ohm", "FullReactance:=", "0ohm")

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
start_freq = 0.5*solution_freq
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
mycommand = "wscript.exe " & """" & "./VBScripts/rad_creation.vbs" & """" &  " " & extent_x_pos & " " & extent_x_neg & " " & extent_y_pos & " " & extent_y_neg & " " & extent_z_pos & " " & extent_z_neg & " " & units & " " & israd& " " & 0
'msgbox(mycommand)
' run the desired VBScript
Set WSHShell = CreateObject("WScript.Shell")
WSHShell.Run mycommand, , True  
Setlocale(locallang)  