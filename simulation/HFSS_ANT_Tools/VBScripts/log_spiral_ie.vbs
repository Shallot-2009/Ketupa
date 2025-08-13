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
oProject.InsertDesign "HFSS-IE", "LogSpiral_Antenna_ADKv1", "DrivenModal", ""
Set oDesign = oProject.SetActiveDesign("LogSpiral_Antenna_ADKv1")








dim  solution_freq, NumberOfArms, InnerRadius, NumberOfTurns, OffsetAngle, ExpansionCoefficient, port_extension, units, israd

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

NumberOfArms = CDbl( args(1)) ' only 2 or 4 is going to be allowed even though UDP works with more
InnerRadius = CDbl( args(2))
NumberOfTurns = CDbl( args(3))
OffsetAngle = CDbl( args(4))
ExpansionCoefficient = CDbl( args(5))



port_extension = CDbl( args(6))

units = args(7)


israd = args(8) ' 0 is for radiation surface, 1 is for PML


''''''''''''''''''''''''

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

dim scale_factor
scale_factor = 1.25

approx_ro = InnerRadius*exp(log(ExpansionCoefficient)*NumberOfTurns)
OuterRadius_test = "InnerRadius*exp(log(ExpansionCoefficient)*NumberOfTurns)"

high_freq = Round(scale_factor*light_speed/(2*3.14159*InnerRadius)*1e-9,2)


low_freq = Round(scale_factor*light_speed/(2*3.14159*approx_ro)*1e-9,2)

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
'Draws Substrate and bottom metalization


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Draws spiral



Set oEditor = oDesign.SetActiveEditor("3D Modeler")
oEditor.CreateUserDefinedPart Array("NAME:UserDefinedPrimitiveParameters", "DllName:=",  _
  "ADKv1/logspiral_adkv1", "Version:=", "1.0", "NoOfParameters:=", 8, "Library:=",  _
  "userlib", Array("NAME:ParamVector", Array("NAME:Pair", "Name:=", "NumberOfPoints", "Value:=",  _
  "200"), Array("NAME:Pair", "Name:=", "NumberOfArms", "Value:=", "NumberOfArms"), Array("NAME:Pair", "Name:=",  _
  "InnerRadius", "Value:=", "InnerRadius"), Array("NAME:Pair", "Name:=", "NumberOfTurns", "Value:=",  _
  "NumberOfTurns"), Array("NAME:Pair", "Name:=", "Offset", "Value:=", "OffsetAngle"), Array("NAME:Pair", "Name:=",  _
  "ConeHeight", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "ExpansionCoefficient", "Value:=",  _
  "ExpansionCoefficient"))), Array("NAME:Attributes", "Name:=",  _
  "LogSpiralAntenna1", "Flags:=", "", "Color:=", "(255 128 65)", "Transparency:=",  _
  0.3, "PartCoordinateSystem:=", "Global", "MaterialName:=",  _
  "vacuum", "SolveInside:=", true)
  
 oEditor.SeparateBody Array("NAME:Selections", "Selections:=", "LogSpiralAntenna1", "NewPartsModelFlag:=",  _
  "Model")

  
  oEditor.CreatePolyline Array("NAME:PolylineParameters", "CoordinateSystemID:=", -1, "IsPolylineCovered:=",  _
  true, "IsPolylineClosed:=", true, Array("NAME:PolylinePoints", Array("NAME:PLPoint", "X:=",  _
  "InnerRadius", "Y:=", "0mm", "Z:=", "0mm"), Array("NAME:PLPoint", "X:=",  _
  "InnerRadius*cos(OffsetAngle)", "Y:=", "InnerRadius*sin(OffsetAngle)", "Z:=", "0mm"), Array("NAME:PLPoint", "X:=",  _
  "-InnerRadius", "Y:=", "0mm", "Z:=", "0mm"), Array("NAME:PLPoint", "X:=",  _
  "-InnerRadius*cos(OffsetAngle)", "Y:=", "-InnerRadius*sin(OffsetAngle)", "Z:=", "0mm"), Array("NAME:PLPoint", "X:=",  _
  "InnerRadius", "Y:=", "0mm", "Z:=", "0mm")), Array("NAME:PolylineSegments", Array("NAME:PLSegment", "SegmentType:=",  _
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
  "0mm", "Y:=", "-InnerRadius", "Z:=", "0mm"), Array("NAME:PLPoint", "X:=",  _
  "0mm", "Y:=", "-InnerRadius", "Z:=", "port_extension"), Array("NAME:PLPoint", "X:=",  _
  "InnerRadius*sin(OffsetAngle)", "Y:=", "-InnerRadius*cos(OffsetAngle)", "Z:=", "port_extension"), Array("NAME:PLPoint", "X:=",  _
  "InnerRadius*sin(OffsetAngle)", "Y:=", "-InnerRadius*cos(OffsetAngle)", "Z:=", "0mm"), Array("NAME:PLPoint", "X:=",  _
  "0mm", "Y:=", "-InnerRadius", "Z:=", "0mm")), Array("NAME:PolylineSegments", Array("NAME:PLSegment", "SegmentType:=",  _
  "Line", "StartIndex:=", 0, "NoOfPoints:=", 2), Array("NAME:PLSegment", "SegmentType:=",  _
  "Line", "StartIndex:=", 1, "NoOfPoints:=", 2), Array("NAME:PLSegment", "SegmentType:=",  _
  "Line", "StartIndex:=", 2, "NoOfPoints:=", 2), Array("NAME:PLSegment", "SegmentType:=",  _
  "Line", "StartIndex:=", 3, "NoOfPoints:=", 2))), Array("NAME:Attributes", "Name:=",  _
  "port_ext1", "Flags:=", "", "Color:=", "(255 128 65)", "Transparency:=",  _
  0.8, "PartCoordinateSystem:=", "Global", "MaterialName:=",  _
  "vacuum", "SolveInside:=", true)
  
      oEditor.CreatePolyline Array("NAME:PolylineParameters", "CoordinateSystemID:=", -1, "IsPolylineCovered:=",  _
  true, "IsPolylineClosed:=", true, Array("NAME:PolylinePoints", Array("NAME:PLPoint", "X:=",  _
  "0mm", "Y:=", "InnerRadius", "Z:=", "0mm"), Array("NAME:PLPoint", "X:=",  _
  "0mm", "Y:=", "InnerRadius", "Z:=", "port_extension"), Array("NAME:PLPoint", "X:=",  _
  "-InnerRadius*sin(OffsetAngle)", "Y:=", "InnerRadius*cos(OffsetAngle)", "Z:=", "port_extension"), Array("NAME:PLPoint", "X:=",  _
  "-InnerRadius*sin(OffsetAngle)", "Y:=", "InnerRadius*cos(OffsetAngle)", "Z:=", "0mm"), Array("NAME:PLPoint", "X:=",  _
  "0mm", "Y:=", "InnerRadius", "Z:=", "0mm")), Array("NAME:PolylineSegments", Array("NAME:PLSegment", "SegmentType:=",  _
  "Line", "StartIndex:=", 0, "NoOfPoints:=", 2), Array("NAME:PLSegment", "SegmentType:=",  _
  "Line", "StartIndex:=", 1, "NoOfPoints:=", 2), Array("NAME:PLSegment", "SegmentType:=",  _
  "Line", "StartIndex:=", 2, "NoOfPoints:=", 2), Array("NAME:PLSegment", "SegmentType:=",  _
  "Line", "StartIndex:=", 3, "NoOfPoints:=", 2))), Array("NAME:Attributes", "Name:=",  _
  "port_ext2", "Flags:=", "", "Color:=", "(255 128 65)", "Transparency:=",  _
  0.8, "PartCoordinateSystem:=", "Global", "MaterialName:=",  _
  "vacuum", "SolveInside:=", true)
  
    oEditor.CreatePolyline Array("NAME:PolylineParameters", "CoordinateSystemID:=", -1, "IsPolylineCovered:=",  _
  true, "IsPolylineClosed:=", true, Array("NAME:PolylinePoints", Array("NAME:PLPoint", "X:=",  _
  "0mm", "Y:=", "-InnerRadius", "Z:=", "port_extension"), Array("NAME:PLPoint", "X:=",  _
  "InnerRadius*sin(OffsetAngle)", "Y:=", "-InnerRadius*cos(OffsetAngle)", "Z:=", "port_extension"), Array("NAME:PLPoint", "X:=",  _
  "0mm", "Y:=", "InnerRadius", "Z:=", "port_extension"), Array("NAME:PLPoint", "X:=",  _
  "-InnerRadius*sin(OffsetAngle)", "Y:=", "InnerRadius*cos(OffsetAngle)", "Z:=", "port_extension"), Array("NAME:PLPoint", "X:=",  _
  "0mm", "Y:=", "-InnerRadius", "Z:=", "port_extension")), Array("NAME:PolylineSegments", Array("NAME:PLSegment", "SegmentType:=",  _
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
Set oModule = oDesign.GetModule("BoundarySetup")  
oModule.AssignPerfectE Array("NAME:LogSpiralAntenna1", "Objects:=", Array("LogSpiralAntenna1", "LogSpiralAntenna1_Separate1"))


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


Set oModule = oDesign.GetModule("BoundarySetup")
  oModule.AssignLumpedPort Array("NAME:p1", "Objects:=", Array("port1"), "RenormalizeAllTerminals:=",  _
  true, "TerminalIDList:=", Array())
oModule.AutoIdentifyTerminals Array("NAME:ReferenceConductors", "LogSpiralAntenna1_Separate1"),  _
  "p1", true
oModule.EditTerminal "LogSpiralAntenna1_T1", Array("NAME:LogSpiralAntenna1_T1", "ParentBndID:=",  _
  "p1", "TerminalResistance:=", "188ohm")
  
if NumberOfArms = 4 then




   Set oModule = oDesign.GetModule("BoundarySetup")  
oModule.AssignPerfectE Array("NAME:LogSpiralAntenna_2", "Objects:=", Array("LogSpiralAntenna1_Separate2", "LogSpiralAntenna1_Separate3"))
  
  oModule.AssignLumpedPort Array("NAME:p2", "Objects:=", Array("port2"), "RenormalizeAllTerminals:=",  _
  true, "TerminalIDList:=", Array())
oModule.AutoIdentifyTerminals Array("NAME:ReferenceConductors", "port_ext2"),  _
  "p2", true
oModule.EditTerminal "port_ext1_T1", Array("NAME:port_ext1_T1", "ParentBndID:=",  _
  "p2", "TerminalResistance:=", "188ohm")
  
Set oModule = oDesign.GetModule("Solutions")
oModule.EditSources Array("NAME:Sources", "LogSpiralAntenna1_T1:=", Array(false, "1", "0deg"), "port_ext1_T1:=", Array( _
  false, "1", "0deg"))  

end if
  
  

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
start_freq = low_freq
stop_freq = high_freq

           

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
 
if NumberOfArms = 4 then   
oModule.CreateReport "Return Loss", "Solution Data", "Rectangular Plot",  _
"Setup1 : Sweep1", Array(), Array("Freq:=", Array("All")), Array("X Component:=", "Freq", "Y Component:=", Array("dB(S(1,1))","dB(S(2,2))")), Array()

oModule.CreateReport "Input Impedance", "Solution Data", "Smith Plot",_
"Setup1 : Sweep1", Array("Domain:=", "Sweep"), Array("Freq:=", Array("All")), _
Array("Polar Component:=", Array("S11","S22")),Array()

end if

if NumberOfArms = 2 then   
oModule.CreateReport "Return Loss", "Solution Data", "Rectangular Plot",  _
"Setup1 : Sweep1", Array(), Array("Freq:=", Array("All")), Array("X Component:=", "Freq", "Y Component:=", Array("dB(S(1,1))")), Array()

oModule.CreateReport "Input Impedance", "Solution Data", "Smith Plot",_
"Setup1 : Sweep1", Array("Domain:=", "Sweep"), Array("Freq:=", Array("All")), _
Array("Polar Component:=", Array("S11")),Array()

end if


oModule.CreateReport "ff_3D_GainRHCP", "Far Fields", "3D Polar Plot",  _
"Setup1 : LastAdaptive", Array("Context:=", "infSphere"), Array("Phi:=", Array( _
"All"), "Theta:=", Array("All")), Array("Phi Component:=",  _
"Phi", "Theta Component:=", "Theta", "Mag Component:=", Array("dB(GainRHCP)")), Array()

oModule.CreateReport "ff_3D_GainLHCP", "Far Fields", "3D Polar Plot",  _
"Setup1 : LastAdaptive", Array("Context:=", "infSphere"), Array("Phi:=", Array( _
"All"), "Theta:=", Array("All")), Array("Phi Component:=",  _
"Phi", "Theta Component:=", "Theta", "Mag Component:=", Array("dB(GainLHCP)")), Array()

oModule.CreateReport "ff_2D_GainRHCP", "Far Fields", "XY Plot",  _
"Setup1 : LastAdaptive", Array("Context:=", "infSphere"), Array("Theta:=", Array( _
"All"), "Phi:=", Array("0deg")), Array("X Component:=",  _
"Theta", "Y Component:=", Array("dB(GainRHCP)")), Array()

oModule.AddTraces "ff_2D_GainRHCP", "Setup1 : LastAdaptive", Array("Context:=",  _
"infSphere"), Array("Theta:=", Array("All"), "Phi:=", Array("90deg")_
), Array("X Component:=", "Theta", "Y Component:=", Array("dB(GainRHCP)")), Array()


oModule.CreateReport "ff_2D_GainLHCP", "Far Fields", "XY Plot",  _
"Setup1 : LastAdaptive", Array("Context:=", "infSphere"), Array("Theta:=", Array( _
"All"), "Phi:=", Array("0deg")), Array("X Component:=",  _
"Theta", "Y Component:=", Array("dB(GainLHCP)")), Array()

oModule.AddTraces "ff_2D_GainLHCP", "Setup1 : LastAdaptive", Array("Context:=",  _
"infSphere"), Array("Theta:=", Array("All"), "Phi:=", Array("90deg")_
), Array("X Component:=", "Theta", "Y Component:=", Array("dB(GainLHCP)")), Array()




Set oDesktop = oAnsoftApp.GetAppDesktop()
oDesktop.CloseAllWindows()
Set oModeler = oDesign.SetActiveEditor("3D Modeler")

oEditor.ShowWindow


  
Setlocale(locallang)  
  
  