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
oProject.InsertDesign "HFSS", "ArchimedeanConical_Antenna_ADKv1", "DrivenModal", ""
Set oDesign = oProject.SetActiveDesign("ArchimedeanConical_Antenna_ADKv1")








dim  solution_freq, NumberOfArms, InnerRadius, NumberOfTurns, OffsetAngle, ExpansionCoefficient, SpiralCoefficient, ConeHeight, port_extension, units, israd

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

NumberOfArms = CDbl( args(1) )' only 2 or 4 is going to be allowed even though UDP works with more
InnerRadius = CDbl( args(2))
NumberOfTurns = CDbl( args(3))
OffsetAngle = CDbl( args(4))
ExpansionCoefficient = CDbl( args(5))
SpiralCoefficient = CDbl( args(6))
ConeHeight = CDbl( args(7))


port_extension = CDbl( args(8))

units = args(9)


israd = args(10) ' 0 is for radiation surface, 1 is for PML




dim unit_conversion

select case units
  case "um"
   unit_conversion = 1e6
  case "mm"
   unit_conversion = 1e3
  case "cm"
   unit_conversion = 1e2
  case "meter"
   unit_conversion = 1
  case "ft"
   unit_conversion = 3.2808399
  case "in"
   unit_conversion = 39.3700787
  case "mil"
   unit_conversion = 39370.0787

end select



''''''''''''''''''''''''''''''''


dim light_speed
light_speed = 299792458

select case units
 case "um"
   light_speed = 299800000000000
 case "mm"
   light_speed = 299800000000
 case "cm"
   light_speed = 29980000000
 case "m"
   light_speed = 299800000
  case "mil"
   light_speed = 11800000000000
  case "in"
   light_speed = 11800000000
  case "ft"
   light_speed = 983600000
end select



dim low_freq, high_freq, approx_ro



approx_ro = InnerRadius+ExpansionCoefficient*(NumberOfTurns*2*3.14159265358979323846)^(1/SpiralCoefficient)*1e-3*unit_conversion  'the second part of this equation is assumed to be mm
OuterRadius_test = "InnerRadius+ExpansionCoefficient*(NumberOfTurns*2*3.14159265358979323846)^(1/SpiralCoefficient)*1e-3"

high_freq = Round(light_speed/(2*3.14159*InnerRadius)*1e-9,2)


low_freq = Round(light_speed/(2*3.14159*approx_ro)*1e-9,2)

dim Bandwidth

Bandwidth = high_freq/low_freq

if Bandwidth >10 then
msgbox("Greater than 10:1 Bandwidth may not be practical. Please Check Dimensions.")
high_freq = Round(low_freq*10,2)
solution_freq = (high_freq+low_freq)/2+low_freq
end if











oDesign.ChangeProperty Array("NAME:AllTabs", Array("NAME:LocalVariableTab", Array("NAME:PropServers", "LocalVariables"), Array("NAME:NewProps", _
Array("NAME:--Spiral Dimensions", "PropType:=", "SeparatorProp", "UserDef:=", true, "Value:=", ""), _
Array("NAME:NumberOfArms", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", NumberOfArms), _
Array("NAME:InnerRadius", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", InnerRadius & units), _
Array("NAME:NumberOfTurns", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", NumberOfTurns), _
Array("NAME:OffsetAngle", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", OffsetAngle & "deg"), _
Array("NAME:ExpansionCoefficient", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", ExpansionCoefficient), _
Array("NAME:SpiralCoefficient", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", SpiralCoefficient), _
Array("NAME:ConeHeight", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", ConeHeight & units), _
Array("NAME:--Port", "PropType:=", "SeparatorProp", "UserDef:=", true, "Value:=", ""), _
Array("NAME:port_extension", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", port_extension & units))))




oDesign.ChangeProperty Array("NAME:AllTabs", Array("NAME:LocalVariableTab", Array("NAME:PropServers",  _
  "LocalVariables"), Array("NAME:NewProps", Array("NAME:OuterRadius_calculated", "PropType:=", "VariableProp", "UserDef:=",  _
  true, "Value:=", OuterRadius_test))))
  
oDesign.ChangeProperty Array("NAME:AllTabs", Array("NAME:LocalVariableTab", Array("NAME:PropServers",  _
  "LocalVariables"), Array("NAME:ChangedProps", Array("NAME:OuterRadius_calculated", "ReadOnly:=",  _
  true), Array("NAME:OuterRadius_calculated", "Hidden:=", true))))

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Draws spiral



Set oEditor = oDesign.SetActiveEditor("3D Modeler")
oEditor.CreateUserDefinedPart Array("NAME:UserDefinedPrimitiveParameters", "DllName:=",  _
  "ADKv1/archimeadean_adkv1", "Version:=", "1.0", "NoOfParameters:=", 8, "Library:=",  _
  "userlib", Array("NAME:ParamVector", Array("NAME:Pair", "Name:=", "NumberOfPoints", "Value:=",  _
  "200"), Array("NAME:Pair", "Name:=", "NumberOfArms", "Value:=", "NumberOfArms"), Array("NAME:Pair", "Name:=",  _
  "InnerRadius", "Value:=", "InnerRadius"), Array("NAME:Pair", "Name:=", "NumberOfTurns", "Value:=",  _
  "NumberOfTurns"), Array("NAME:Pair", "Name:=", "Offset", "Value:=", "OffsetAngle"), Array("NAME:Pair", "Name:=",  _
  "ConeHeight", "Value:=", "ConeHeight"), Array("NAME:Pair", "Name:=", "ExpansionCoefficient", "Value:=",  _
  "ExpansionCoefficient"), Array("NAME:Pair", "Name:=", "SpiralCoefficient", "Value:=", "SpiralCoefficient"))), Array("NAME:Attributes", "Name:=",  _
  "SpiralAntenna1", "Flags:=", "", "Color:=", "(255 128 65)", "Transparency:=",  _
  0.3, "PartCoordinateSystem:=", "Global", "MaterialName:=",  _
  "vacuum", "SolveInside:=", true)
  


  
  oEditor.CreatePolyline Array("NAME:PolylineParameters", "CoordinateSystemID:=", -1, "IsPolylineCovered:=",  _
  true, "IsPolylineClosed:=", true, Array("NAME:PolylinePoints", Array("NAME:PLPoint", "X:=",  _
  "InnerRadius", "Y:=", "0mm", "Z:=", "ConeHeight"), Array("NAME:PLPoint", "X:=",  _
  "InnerRadius*cos(OffsetAngle)", "Y:=", "InnerRadius*sin(OffsetAngle)", "Z:=", "ConeHeight"), Array("NAME:PLPoint", "X:=",  _
  "-InnerRadius", "Y:=", "0mm", "Z:=", "ConeHeight"), Array("NAME:PLPoint", "X:=",  _
  "-InnerRadius*cos(OffsetAngle)", "Y:=", "-InnerRadius*sin(OffsetAngle)", "Z:=", "ConeHeight"), Array("NAME:PLPoint", "X:=",  _
  "InnerRadius", "Y:=", "0mm", "Z:=", "ConeHeight")), Array("NAME:PolylineSegments", Array("NAME:PLSegment", "SegmentType:=",  _
  "Line", "StartIndex:=", 0, "NoOfPoints:=", 2), Array("NAME:PLSegment", "SegmentType:=",  _
  "Line", "StartIndex:=", 1, "NoOfPoints:=", 2), Array("NAME:PLSegment", "SegmentType:=",  _
  "Line", "StartIndex:=", 2, "NoOfPoints:=", 2), Array("NAME:PLSegment", "SegmentType:=",  _
  "Line", "StartIndex:=", 3, "NoOfPoints:=", 2))), Array("NAME:Attributes", "Name:=",  _
  "port1", "Flags:=", "", "Color:=", "(255 128 65)", "Transparency:=",  _
  0.8, "PartCoordinateSystem:=", "Global", "MaterialName:=",  _
  "vacuum", "SolveInside:=", true)
  
  if NumberOfArms = 4 then
  
    oEditor.CreatePolyline Array("NAME:PolylineParameters", "CoordinateSystemID:=", -1, "IsPolylineCovered:=",  _
  true, "IsPolylineClosed:=", true, Array("NAME:PolylinePoints", Array("NAME:PLPoint", "X:=",  _
  "0mm", "Y:=", "-InnerRadius", "Z:=", "ConeHeight"), Array("NAME:PLPoint", "X:=",  _
  "0mm", "Y:=", "-InnerRadius", "Z:=", "ConeHeight+port_extension"), Array("NAME:PLPoint", "X:=",  _
  "InnerRadius*sin(OffsetAngle)", "Y:=", "-InnerRadius*cos(OffsetAngle)", "Z:=", "ConeHeight+port_extension"), Array("NAME:PLPoint", "X:=",  _
  "InnerRadius*sin(OffsetAngle)", "Y:=", "-InnerRadius*cos(OffsetAngle)", "Z:=", "ConeHeight"), Array("NAME:PLPoint", "X:=",  _
  "0mm", "Y:=", "-InnerRadius", "Z:=", "ConeHeight")), Array("NAME:PolylineSegments", Array("NAME:PLSegment", "SegmentType:=",  _
  "Line", "StartIndex:=", 0, "NoOfPoints:=", 2), Array("NAME:PLSegment", "SegmentType:=",  _
  "Line", "StartIndex:=", 1, "NoOfPoints:=", 2), Array("NAME:PLSegment", "SegmentType:=",  _
  "Line", "StartIndex:=", 2, "NoOfPoints:=", 2), Array("NAME:PLSegment", "SegmentType:=",  _
  "Line", "StartIndex:=", 3, "NoOfPoints:=", 2))), Array("NAME:Attributes", "Name:=",  _
  "port_ext1", "Flags:=", "", "Color:=", "(255 128 65)", "Transparency:=",  _
  0.8, "PartCoordinateSystem:=", "Global", "MaterialName:=",  _
  "vacuum", "SolveInside:=", true)
  
      oEditor.CreatePolyline Array("NAME:PolylineParameters", "CoordinateSystemID:=", -1, "IsPolylineCovered:=",  _
  true, "IsPolylineClosed:=", true, Array("NAME:PolylinePoints", Array("NAME:PLPoint", "X:=",  _
  "0mm", "Y:=", "InnerRadius", "Z:=", "ConeHeight"), Array("NAME:PLPoint", "X:=",  _
  "0mm", "Y:=", "InnerRadius", "Z:=", "ConeHeight+port_extension"), Array("NAME:PLPoint", "X:=",  _
  "-InnerRadius*sin(OffsetAngle)", "Y:=", "InnerRadius*cos(OffsetAngle)", "Z:=", "ConeHeight+port_extension"), Array("NAME:PLPoint", "X:=",  _
  "-InnerRadius*sin(OffsetAngle)", "Y:=", "InnerRadius*cos(OffsetAngle)", "Z:=", "ConeHeight"), Array("NAME:PLPoint", "X:=",  _
  "0mm", "Y:=", "InnerRadius", "Z:=", "ConeHeight")), Array("NAME:PolylineSegments", Array("NAME:PLSegment", "SegmentType:=",  _
  "Line", "StartIndex:=", 0, "NoOfPoints:=", 2), Array("NAME:PLSegment", "SegmentType:=",  _
  "Line", "StartIndex:=", 1, "NoOfPoints:=", 2), Array("NAME:PLSegment", "SegmentType:=",  _
  "Line", "StartIndex:=", 2, "NoOfPoints:=", 2), Array("NAME:PLSegment", "SegmentType:=",  _
  "Line", "StartIndex:=", 3, "NoOfPoints:=", 2))), Array("NAME:Attributes", "Name:=",  _
  "port_ext2", "Flags:=", "", "Color:=", "(255 128 65)", "Transparency:=",  _
  0.8, "PartCoordinateSystem:=", "Global", "MaterialName:=",  _
  "vacuum", "SolveInside:=", true)
  
    oEditor.CreatePolyline Array("NAME:PolylineParameters", "CoordinateSystemID:=", -1, "IsPolylineCovered:=",  _
  true, "IsPolylineClosed:=", true, Array("NAME:PolylinePoints", Array("NAME:PLPoint", "X:=",  _
  "0mm", "Y:=", "-InnerRadius", "Z:=", "ConeHeight+port_extension"), Array("NAME:PLPoint", "X:=",  _
  "InnerRadius*sin(OffsetAngle)", "Y:=", "-InnerRadius*cos(OffsetAngle)", "Z:=", "ConeHeight+port_extension"), Array("NAME:PLPoint", "X:=",  _
  "0mm", "Y:=", "InnerRadius", "Z:=", "ConeHeight+port_extension"), Array("NAME:PLPoint", "X:=",  _
  "-InnerRadius*sin(OffsetAngle)", "Y:=", "InnerRadius*cos(OffsetAngle)", "Z:=", "ConeHeight+port_extension"), Array("NAME:PLPoint", "X:=",  _
  "0mm", "Y:=", "-InnerRadius", "Z:=", "ConeHeight+port_extension")), Array("NAME:PolylineSegments", Array("NAME:PLSegment", "SegmentType:=",  _
  "Line", "StartIndex:=", 0, "NoOfPoints:=", 2), Array("NAME:PLSegment", "SegmentType:=",  _
  "Line", "StartIndex:=", 1, "NoOfPoints:=", 2), Array("NAME:PLSegment", "SegmentType:=",  _
  "Line", "StartIndex:=", 2, "NoOfPoints:=", 2), Array("NAME:PLSegment", "SegmentType:=",  _
  "Line", "StartIndex:=", 3, "NoOfPoints:=", 2))), Array("NAME:Attributes", "Name:=",  _
  "port2", "Flags:=", "", "Color:=", "(255 128 65)", "Transparency:=",  _
  0.8, "PartCoordinateSystem:=", "Global", "MaterialName:=",  _
  "vacuum", "SolveInside:=", true)
  
  Set oModule = oDesign.GetModule("BoundarySetup")  
oModule.AssignPerfectE Array("NAME:Port_Ext1", "Objects:=", Array("port_ext1"), "InfGroundPlane:=", false) 
oModule.AssignPerfectE Array("NAME:Port_Ext2", "Objects:=", Array("port_ext2"), "InfGroundPlane:=", false)
  
  end if
  
Set oModule = oDesign.GetModule("BoundarySetup")  
oModule.AssignPerfectE Array("NAME:SpiralAntenna1", "Objects:=", Array("SpiralAntenna1"), "InfGroundPlane:=", false) 

dim end_vector_pointX, end_vector_pointY
dim start_vector_pointX, start_vector_pointY

end_vector_pointX = -InnerRadius*cos(OffsetAngle*3.14159265358979323846/180)'InnerRadius*sin(OffsetAngle/2*3.14159/180)
end_vector_pointX = end_vector_pointX & units
end_vector_pointY = -InnerRadius*sin(OffsetAngle*3.14159265358979323846/180)'InnerRadius*cos(OffsetAngle/2*3.14159/180)
end_vector_pointY = end_vector_pointY & units


start_vector_pointX = InnerRadius'-InnerRadius*sin(OffsetAngle/2*3.14159/180)
start_vector_pointX = start_vector_pointX & units
start_vector_pointY = 0'-InnerRadius*cos(OffsetAngle/2*3.14159/180)
start_vector_pointY = start_vector_pointY & units


oDesign.SetSolutionType "DrivenModal"
Set oModule = oDesign.GetModule("BoundarySetup")
oModule.AssignLumpedPort Array("NAME:p1", "Objects:=", Array("port1"), Array("NAME:Modes", Array("NAME:Mode1", "ModeNum:=",  _
  1, "UseIntLine:=", true, Array("NAME:IntLine", "Start:=", Array(start_vector_pointX, start_vector_pointY,  _
  ConeHeight & units), "End:=", Array(end_vector_pointX, end_vector_pointY, ConeHeight & units)), "CharImp:=", "Zpi", "RenormImp:=",  _
  "188ohm")), "FullResistance:=", "50ohm", "FullReactance:=", "0ohm")
  
  
if NumberOfArms = 4 then

end_vector_pointY = InnerRadius*cos(OffsetAngle*3.14159265358979323846/180)'InnerRadius*sin(OffsetAngle/2*3.14159/180)
end_vector_pointY = end_vector_pointY & units
end_vector_pointX = -InnerRadius*sin(OffsetAngle*3.14159265358979323846/180)'InnerRadius*cos(OffsetAngle/2*3.14159/180)
end_vector_pointX = end_vector_pointX & units


start_vector_pointY = -InnerRadius'-InnerRadius*sin(OffsetAngle/2*3.14159/180)
start_vector_pointY = start_vector_pointY & units
start_vector_pointX = 0'-InnerRadius*cos(OffsetAngle/2*3.14159/180)
start_vector_pointX = start_vector_pointX & units

dim vector_pointZ
vector_pointZ = ConeHeight + port_extension
vector_pointZ =  vector_pointZ & units

oModule.AssignLumpedPort Array("NAME:p2", "Objects:=", Array("port2"), Array("NAME:Modes", Array("NAME:Mode1", "ModeNum:=",  _
  1, "UseIntLine:=", true, Array("NAME:IntLine", "Start:=", Array(start_vector_pointX, start_vector_pointY,  _
  vector_pointZ), "End:=", Array(end_vector_pointX, end_vector_pointY, vector_pointZ)), "CharImp:=", "Zpi", "RenormImp:=",  _
  "188ohm")), "FullResistance:=", "50ohm", "FullReactance:=", "0ohm")
  

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
start_freq = low_freq
stop_freq = high_freq

 oModule.InsertFrequencySweep "Setup1", Array("NAME:Sweep1", "IsEnabled:=", true, "SetupType:=",  _
  "LinearCount", "StartValue:=", low_freq&"GHz", "StopValue:=", high_freq&"GHz", "Count:=", 200, "Type:=",  _
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

extent_x_pos = "OuterRadius_calculated"
extent_x_neg = "-OuterRadius_calculated"

extent_y_pos = "OuterRadius_calculated"
extent_y_neg = "-OuterRadius_calculated"

extent_z_pos = "ConeHeight"
extent_z_neg = "0"

if israd = "ABC" then
 israd = "Rad"

end if


dim mycommand
' full command which invokes wscript to run desired VBScript and passes
mycommand = "wscript.exe " & """" & "./VBScripts/rad_creation.vbs" & """" &  " " & extent_x_pos & " " & extent_x_neg & " " & extent_y_pos & " " & extent_y_neg & " " & extent_z_pos & " " & extent_z_neg & " " & units & " " & israd& " " & low_freq
'msgbox(mycommand)
' run the desired VBScript
Set WSHShell = CreateObject("WScript.Shell")
WSHShell.Run mycommand, , True
  
  
Setlocale(locallang)  
  