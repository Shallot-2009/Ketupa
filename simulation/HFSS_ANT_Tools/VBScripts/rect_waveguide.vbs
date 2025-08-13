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
oProject.InsertDesign "HFSS", "WGRectangular_Antenna_ADKv1", "DrivenModal", ""
Set oDesign = oProject.SetActiveDesign("WGRectangular_Antenna_ADKv1")



dim  solution_freq, a, b, WG_length, Wall_Thickness, units, israd

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

a = CDbl( args(1))
b = CDbl( args(2))
WG_length = CDbl( args(3))
Wall_Thickness =CDbl( args(4))

units = args(5)

israd = args(6) ' 0 is for radiation surface, 1 is for PML


if b > a then
dim temp
temp = a
a=b
b=temp
end if

oDesign.ChangeProperty Array("NAME:AllTabs", Array("NAME:LocalVariableTab", Array("NAME:PropServers", "LocalVariables"), Array("NAME:NewProps", _
Array("NAME:--Waveguide Dimensions", "PropType:=", "SeparatorProp", "UserDef:=", true, "Value:=", ""), _
Array("NAME:a", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", a & units), _
Array("NAME:b", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", b & units), _
Array("NAME:WG_length", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", WG_length & units), _
Array("NAME:Wall_Thickness", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", Wall_Thickness & units))))


'''''''''''''''''''''''
' draw waveguide
Set oEditor = oDesign.SetActiveEditor("3D Modeler")


''waveguide interior
oEditor.CreateBox Array("NAME:BoxParameters", "CoordinateSystemID:=", -1, "XPosition:=",  _
  "-a/2", "YPosition:=", "-b/2", "ZPosition:=", "-WG_length", "XSize:=", "a", "YSize:=",  _
  "b", "ZSize:=", "WG_length"), Array("NAME:Attributes", "Name:=", "WG_inner", "Flags:=", "", "Color:=",  _
  "(132 132 193)", "Transparency:=", 0.8, "PartCoordinateSystem:=",  _
  "Global", "MaterialName:=", "vacuum", "SolveInside:=", true)

''''wall thickness  
oEditor.CreateBox Array("NAME:BoxParameters", "CoordinateSystemID:=", -1, "XPosition:=",  _
  "-a/2-Wall_Thickness", "YPosition:=", "-b/2-Wall_Thickness", "ZPosition:=", "-WG_length", "XSize:=", "a+2*Wall_Thickness", "YSize:=",  _
  "b+2*Wall_Thickness", "ZSize:=", "WG_length"), Array("NAME:Attributes", "Name:=", "WG_outer", "Flags:=", "", "Color:=",  _
  "(132 132 193)", "Transparency:=", 0.8, "PartCoordinateSystem:=",  _
  "Global", "MaterialName:=", "pec", "SolveInside:=", false)  




oEditor.Subtract Array("NAME:Selections", "Blank Parts:=", "WG_outer", "Tool Parts:=", "WG_inner"), Array("NAME:SubtractParameters", "CoordinateSystemID:=", -1, "KeepOriginals:=", true)

'''''''''''''''''''''''''''''''''''
' create port

''port cap 
oEditor.CreateBox Array("NAME:BoxParameters", "CoordinateSystemID:=", -1, "XPosition:=",  _
  "-a/2-Wall_Thickness", "YPosition:=", "-b/2-Wall_Thickness", "ZPosition:=", "-WG_length", "XSize:=", "a+2*Wall_Thickness", "YSize:=",  _
  "b+2*Wall_Thickness", "ZSize:=", "-Wall_Thickness"), Array("NAME:Attributes", "Name:=", "Port_Cap", "Flags:=", "", "Color:=",  _
  "(132 132 193)", "Transparency:=", 0.8, "PartCoordinateSystem:=",  _
  "Global", "MaterialName:=", "pec", "SolveInside:=", false)    

'port
oEditor.CreateRectangle Array("NAME:RectangleParameters", "CoordinateSystemID:=",  _
  -1, "IsCovered:=", true, "XStart:=", "-a/2", "YStart:=", "-b/2", "ZStart:=",  _
  "-WG_length", "Width:=", "a", "Height:=", "b", "WhichAxis:=", "Z"), Array("NAME:Attributes", "Name:=",  _
  "port1", "Flags:=", "", "Color:=", "(132 132 193)", "Transparency:=",  _
  0.8, "PartCoordinateSystem:=", "Global", "MaterialName:=",  _
  "vacuum", "SolveInside:=", true)  

dim end_vector_pointX, end_vector_pointY, end_vector_pointZ
dim start_vector_pointX, start_vector_pointY, start_vector_pointZ

end_vector_pointX = 0
end_vector_pointX = end_vector_pointX & units

end_vector_pointY = b/2
end_vector_pointY = end_vector_pointY & units

end_vector_pointZ = -WG_length
end_vector_pointZ = end_vector_pointZ & units

start_vector_pointX = 0
start_vector_pointX = start_vector_pointX & units

start_vector_pointY = -b/2
start_vector_pointY = start_vector_pointY & units

start_vector_pointZ = -WG_length
start_vector_pointZ = start_vector_pointZ & units

  Set oModule = oDesign.GetModule("BoundarySetup")
  
if a = b then
  oModule.AssignWavePort Array("NAME:p1", "Objects:=", Array("port1"), "NumModes:=",  _
  2, "PolarizeEField:=", true, "DoDeembed:=", false, Array("NAME:Modes", Array("NAME:Mode1", "ModeNum:=",  _
  1, "UseIntLine:=", true, Array("NAME:IntLine", "Start:=", Array(start_vector_pointX, start_vector_pointY,  _
  start_vector_pointZ), "End:=", Array(end_vector_pointX, end_vector_pointY, end_vector_pointZ)), "CharImp:=", "Zpi"), Array("NAME:Mode2", "ModeNum:=",  _
  2, "UseIntLine:=", true, Array("NAME:IntLine", "Start:=", Array(start_vector_pointY, start_vector_pointX, start_vector_pointZ), _
   "End:=", Array(end_vector_pointY, end_vector_pointX, end_vector_pointZ)), "CharImp:=", "Zpv")))
  

elseif a > b then



  oModule.AssignWavePort Array("NAME:p1", "Objects:=", Array("port1"), "NumModes:=",  _
  1, "PolarizeEField:=", false, "DoDeembed:=", false, Array("NAME:Modes", Array("NAME:Mode1", "ModeNum:=",  _
  1, "UseIntLine:=", true, Array("NAME:IntLine", "Start:=", Array(start_vector_pointX, start_vector_pointY,  _
  start_vector_pointZ), "End:=", Array(end_vector_pointX, end_vector_pointY, end_vector_pointZ)), "CharImp:=", "Zpv")))

end if



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
start_freq = Round(.8*solution_freq,2)
stop_freq = Round(1.2*solution_freq,2)

 
  
oModule.InsertFrequencySweep "Setup1", Array("NAME:Sweep1", "IsEnabled:=", true, "SetupType:=",  _
  "LinearCount", "StartValue:=", start_freq&"GHz", "StopValue:=", stop_freq&"GHz", "Count:=", 101, "Type:=",  _
  "Fast", "SaveFields:=", true, "GenerateFieldsForAllFreqs:=", false, "ExtrapToDC:=",  _
  false) 
  
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''  
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'get userlib directory where radition box script is sitting  
' and create radiation box
dim libdir
libdir = oDesktop.GetLibraryDirectory
dim fullpath
fullpath = libdir & "\userlib\AntennaDesignKit\VBScripts\rad_creation.vbs" 


dim extent_x_pos, extent_x_neg, extent_y_pos, extent_y_neg, extent_z_pos, extent_z_neg

extent_x_pos = "a/2"
extent_x_neg = "-a/2"

extent_y_pos = "b/2"
extent_y_neg = "-b/2"

extent_z_pos = "0"
extent_z_neg = "-WG_length"

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
  
  