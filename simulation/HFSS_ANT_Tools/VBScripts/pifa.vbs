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
oProject.InsertDesign "HFSS", "PIFA_Antenna_ADKv1", "DrivenModal", ""
Set oDesign = oProject.SetActiveDesign("PIFA_Antenna_ADKv1")



dim  solution_freq, Length1, Length2, Antenna_Trace_Width, Antenna_Offset, Feed_Offset, Feed_Length, Feed_Width, subH, subX, subY, units, israd

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





Length1 = CDbl( args(1))
Length2 = CDbl( args(2))
Antenna_Trace_Width = CDbl( args(3))
Antenna_Offset =CDbl( args(4))
Feed_Offset = CDbl( args(5))
Feed_Length = CDbl( args(6))
Feed_Width = CDbl( args(7))

subH = ( args(8))
subX = CDbl( args(9))
subY = CDbl( args(10))




units = args(11)

israd = args(12) ' 0 is for radiation surface, 1 is for PML



if subY <= Length2+Antenna_Trace_Width+Feed_Length then
  subY = 2*(Length2+Antenna_Trace_Width+Feed_Length)
end if



oDesign.ChangeProperty Array("NAME:AllTabs", Array("NAME:LocalVariableTab", Array("NAME:PropServers", "LocalVariables"), Array("NAME:NewProps", _
Array("NAME:--Antenna Dimensions", "PropType:=", "SeparatorProp", "UserDef:=", true, "Value:=", ""), _
Array("NAME:Length1", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", Length1 & units), _
Array("NAME:Length2", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", Length2 & units), _
Array("NAME:Antenna_Trace_Width", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", Antenna_Trace_Width & units), _
Array("NAME:Antenna_Offset", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", Antenna_Offset & units), _
Array("NAME:--Feed", "PropType:=", "SeparatorProp", "UserDef:=", true, "Value:=", ""), _
Array("NAME:Feed_Offset", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", Feed_Offset & units), _
Array("NAME:Feed_Length", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", Feed_Length & units), _
Array("NAME:Feed_Width", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", Feed_Width & units), _
Array("NAME:--Substrate Dimensions", "PropType:=", "SeparatorProp", "UserDef:=", true, "Value:=", ""), _
Array("NAME:subH", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", subH ), _
Array("NAME:subX", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", subX & units), _
Array("NAME:subY", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", subY & units))))

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
  "-subX/2", "YPosition:=", "-subY+Length2+Antenna_Trace_Width" , "ZPosition:=", "0mm", "XSize:=", "subX", "YSize:=",  _
  "subY", "ZSize:=", "-subH"), Array("NAME:Attributes", "Name:=", "sub", "Flags:=",  _
  "", "Color:=", "(132 132 193)", "Transparency:=", 0.8, "PartCoordinateSystem:=",  _
  "Global", "MaterialName:=", Material_Name , "SolveInside:=", true)
  
  
oEditor.CreateRectangle Array("NAME:RectangleParameters", "CoordinateSystemID:=",  _
  -1, "IsCovered:=", true, "XStart:=", "-subX/2", "YStart:=", "-subY+Length2+Antenna_Trace_Width", "ZStart:=",  _
  "-subH", "Width:=", "subX", "Height:=", "subY-(Length2+Antenna_Trace_Width)", "WhichAxis:=", "Z"), Array("NAME:Attributes", "Name:=",  _
  "Ground", "Flags:=", "", "Color:=", "(255 128 65)", "Transparency:=",  _
  0.8, "PartCoordinateSystem:=", "Global", "MaterialName:=",  _
  "pec", "SolveInside:=", false)

Set oModule = oDesign.GetModule("BoundarySetup")  

oModule.AssignPerfectE Array("NAME:Ground", "Objects:=", Array("Ground"), "InfGroundPlane:=", false)


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Draws antenna

Set oEditor = oDesign.SetActiveEditor("3D Modeler")


  oEditor.CreateRectangle Array("NAME:RectangleParameters", "CoordinateSystemID:=",  _
  -1, "IsCovered:=", true, "XStart:=", "-Antenna_Trace_Width/2+Feed_Offset", "YStart:=", "0mm", "ZStart:=",  _
  "0mm", "Width:=", "Antenna_Trace_Width", "Height:=", "Length2", "WhichAxis:=", "Z"), Array("NAME:Attributes", "Name:=",  _
  "antenna_feed", "Flags:=", "", "Color:=", "(255 128 65)", "Transparency:=",  _
  0.3, "PartCoordinateSystem:=", "Global", "MaterialName:=",  _
  "pec", "SolveInside:=", false)
  
    oEditor.CreateRectangle Array("NAME:RectangleParameters", "CoordinateSystemID:=",  _
  -1, "IsCovered:=", true, "XStart:=", "-Antenna_Trace_Width/2-Antenna_Offset+Feed_Offset", "YStart:=", "0mm", "ZStart:=",  _
  "0mm", "Width:=", "-Antenna_Trace_Width", "Height:=", "Length2", "WhichAxis:=", "Z"), Array("NAME:Attributes", "Name:=",  _
  "antenna_short", "Flags:=", "", "Color:=", "(255 128 65)", "Transparency:=",  _
  0.3, "PartCoordinateSystem:=", "Global", "MaterialName:=",  _
  "pec", "SolveInside:=", false)
  
      oEditor.CreateRectangle Array("NAME:RectangleParameters", "CoordinateSystemID:=",  _
  -1, "IsCovered:=", true, "XStart:=", "-Antenna_Trace_Width/2-Antenna_Offset+Feed_Offset", "YStart:=", "0mm", "ZStart:=",  _
  "0mm", "Width:=", "-subH", "Height:=", "-Antenna_Trace_Width", "WhichAxis:=", "Y"), Array("NAME:Attributes", "Name:=",  _
  "antenna_short2", "Flags:=", "", "Color:=", "(255 128 65)", "Transparency:=",  _
  0.3, "PartCoordinateSystem:=", "Global", "MaterialName:=",  _
  "pec", "SolveInside:=", false)

    oEditor.CreateRectangle Array("NAME:RectangleParameters", "CoordinateSystemID:=",  _
  -1, "IsCovered:=", true, "XStart:=", "-Antenna_Trace_Width/2-Antenna_Offset-Antenna_Trace_Width+Feed_Offset", "YStart:=", "Length2", "ZStart:=",  _
  "0mm", "Width:=", "Length1", "Height:=", "-Antenna_Trace_Width", "WhichAxis:=", "Z"), Array("NAME:Attributes", "Name:=",  _
  "antenna", "Flags:=", "", "Color:=", "(255 128 65)", "Transparency:=",  _
  0.3, "PartCoordinateSystem:=", "Global", "MaterialName:=",  _
  "pec", "SolveInside:=", false)


  oEditor.Unite Array("NAME:Selections", "Selections:=", "antenna,antenna_short,antenna_feed,antenna_short2"), Array("NAME:UniteParameters", "CoordinateSystemID:=",  _
  -1, "KeepOriginals:=", false)

 Set oModule = oDesign.GetModule("BoundarySetup")  

oModule.AssignPerfectE Array("NAME:antenna", "Objects:=", Array("antenna"), "InfGroundPlane:=", false)


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Draws Microstrip Feed

if Feed_Length <> 0 then
Set oEditor = oDesign.SetActiveEditor("3D Modeler")


  oEditor.CreateRectangle Array("NAME:RectangleParameters", "CoordinateSystemID:=",  _
  -1, "IsCovered:=", true, "XStart:=", "-Feed_Width/2+Feed_Offset", "YStart:=", "0mm", "ZStart:=",  _
  "0mm", "Width:=", "Feed_Width", "Height:=", "-Feed_Length", "WhichAxis:=", "Z"), Array("NAME:Attributes", "Name:=",  _
  "microstrip_feed", "Flags:=", "", "Color:=", "(255 128 65)", "Transparency:=",  _
  0.3, "PartCoordinateSystem:=", "Global", "MaterialName:=",  _
  "pec", "SolveInside:=", false)

 Set oModule = oDesign.GetModule("BoundarySetup")  

oModule.AssignPerfectE Array("NAME:microstrip_feed", "Objects:=", Array("microstrip_feed"), "InfGroundPlane:=", false)

end if

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Draws Port

Set oEditor = oDesign.SetActiveEditor("3D Modeler")


  oEditor.CreateRectangle Array("NAME:RectangleParameters", "CoordinateSystemID:=",  _
  -1, "IsCovered:=", true, "XStart:=", "-Antenna_Trace_Width/2+Feed_Offset", "YStart:=", "-Feed_Length", "ZStart:=",  _
  "0mm", "Width:=", "-subH", "Height:=", "Feed_Width", "WhichAxis:=", "Y"), Array("NAME:Attributes", "Name:=",  _
  "port1", "Flags:=", "", "Color:=", "(132 132 193)", "Transparency:=",  _
  0.8, "PartCoordinateSystem:=", "Global", "MaterialName:=",  _
  "pec", "SolveInside:=", false)




faceid = oEditor.GetFaceByPosition(Array("NAME:FaceParameters","BodyName:=", "port1", "XPosition:=", "Feed_Offset", "YPosition:=","-Feed_Length", "ZPosition:=", "-subH/2"))

  
oDesign.SetSolutionType "DrivenTerminal"
Set oModule = oDesign.GetModule("BoundarySetup")
oModule.AutoIdentifyPorts Array("NAME:Faces", faceid), false, Array("NAME:ReferenceConductors",  _
  "Ground"), "p1", true  
  

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
start_freq = Round(.625*solution_freq,1)
stop_freq = Round(1.4*solution_freq,1)

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

extent_x_pos = "(subX/2)"
extent_x_neg = "(-subX/2)"

extent_y_pos = "(Length2+Antenna_Trace_Width)"
extent_y_neg = "(-(subY-(Length2+Antenna_Trace_Width)))"

extent_z_pos = "0"
extent_z_neg = "(-subH)"

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


  