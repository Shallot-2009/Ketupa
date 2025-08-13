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
oProject.InsertDesign "HFSS-IE", "LinearTaper_Antenna_ADKv1", "DrivenModal", ""
Set oDesign = oProject.SetActiveDesign("LinearTaper_Antenna_ADKv1")



dim cont_or_stepped, solution_freq, Wslot, Lfeed, Wtaper, Ltaper, Wbalun, Lbalun, Nsteps, Wstrip, Lstrip_offset, Lstrip, feed_offset, Wtotal, Ltotal, subH, units, israd

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



dim UDP_name
cont_or_stepped = Cint( args(0)) '0 for continuous curve, 1 for stepped

if cont_or_stepped = 1 then
 UDP_name = "ADKv1/vivaldi_stepped_adkv1"
else 
  UDP_name = "ADKv1/vivaldi_adkv1"
end if





solution_freq = CDbl( args(1))

  dim freq_hz
  freq_hz=solution_freq*1e9

  dim WL 
  WL = light_speed/(freq_hz)



Wslot = CDbl( args(2))
Lfeed = CDbl( args(3))
Wtaper = CDbl( args(4))
Ltaper = CDbl( args(5))
Wbalun = CDbl( args(6))
Lbalun = CDbl( args(7))
Nsteps = CDbl( args(8))



Wstrip = CDbl( args(9))
Lstrip_offset = CDbl( args(10))
Lstrip = CDbl( args(11))
feed_offset = CDbl( args(12))



Wtotal = CDbl( args(13))
Ltotal = CDbl( args(14))
subH = ( args(15))


units = args(16)

israd = args(17) ' 0 is for radiation surface, 1 is for PML

''''''''''''''''''''''''''''''''''''''''''
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
dim mid_freq
mid_freq=light_speed/(Wbalun*(2.2)^(1/2))/4*1e-9




low_freq = Round(light_speed/(Wtaper*2)*1e-9,1)

high_freq = Round(solution_freq,2)

if low_freq >= high_freq then
     low_freq = .9*high_freq
end if

if high_freq <= low_freq then
     high_freq = 1.1*low_freq
end if


dim Bandwidth

Bandwidth = high_freq/low_freq

if Bandwidth >6 then
msgbox("Greater than 6:1 Bandwidth may not be practical. Please Check Dimensions.")
high_freq = Round(low_freq*6,1)
solution_freq = (high_freq+low_freq)/2+low_freq
end if




oDesign.ChangeProperty Array("NAME:AllTabs", Array("NAME:LocalVariableTab", Array("NAME:PropServers", "LocalVariables"), Array("NAME:NewProps", _
Array("NAME:--Tapered Slot", "PropType:=", "SeparatorProp", "UserDef:=", true, "Value:=", ""), _
Array("NAME:Wslot", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", Wslot & units), _
Array("NAME:Lfeed", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", Lfeed & units), _
Array("NAME:Wtaper", "PropType:=", "VariableProp", "UserDef:=",true, "Value:=", Wtaper & units), _
Array("NAME:Ltaper", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", Ltaper & units), _
Array("NAME:Wbalun", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", Wbalun & units), _
Array("NAME:Lbalun", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", Lbalun & units), _
Array("NAME:Nsteps", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", Nsteps), _
Array("NAME:--Stripline Feed", "PropType:=", "SeparatorProp", "UserDef:=", true, "Value:=", ""), _
Array("NAME:Wstrip", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", Wstrip & units), _
Array("NAME:Lstrip_offset", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", Lstrip_offset & units), _
Array("NAME:Lstrip", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", Lstrip & units), _
Array("NAME:feed_offset", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", feed_offset & units), _
Array("NAME:--Substrate", "PropType:=", "SeparatorProp", "UserDef:=", true, "Value:=", ""), _
Array("NAME:Wtotal", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", Wtotal & units), _
Array("NAME:Ltotal", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", Ltotal & units), _
Array("NAME:subH", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", subH ))))

 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'create substrate material

Material_Name = args(18)
Material_Name = Material_Name & "_ADK"
Permittivity = args(19)
TandD = args(20)

Set oDefinitionManager = oProject.GetDefinitionManager()
oDefinitionManager.AddMaterial Array("NAME:" & Material_Name, "CoordinateSystemType:=",  _
  "Cartesian", Array("NAME:AttachedData"), Array("NAME:ModifierData"), "permittivity:=", Permittivity , "dielectric_loss_tangent:=",TandD)

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Draws Substrate and top and bottom metalization

Set oEditor = oDesign.SetActiveEditor("3D Modeler")
oEditor.CreateBox Array("NAME:BoxParameters", "CoordinateSystemID:=", -1, "XPosition:=",  _
  "0mm", "YPosition:=", "-Wtotal/2" , "ZPosition:=", "0mm", "XSize:=", "Ltotal", "YSize:=",  _
  "Wtotal", "ZSize:=", "subH"), Array("NAME:Attributes", "Name:=", "sub", "Flags:=",  _
  "", "Color:=", "(132 132 193)", "Transparency:=", 0.8, "PartCoordinateSystem:=",  _
  "Global", "MaterialName:=", Material_Name, "SolveInside:=", true)
  
  
  
oEditor.CreateRelativeCS Array("NAME:RelativeCSParameters", "CoordinateSystemID:=",  _
  -1, "OriginX:=", "0mm", "OriginY:=", "0mm", "OriginZ:=", "0mm", "XAxisXvec:=",  _
  "1mm", "XAxisYvec:=", "0mm", "XAxisZvec:=", "0mm", "YAxisXvec:=", "0mm", "YAxisYvec:=",  _
  "1mm", "YAxisZvec:=", "0mm"), Array("NAME:Attributes", "Name:=", "Bottom_CS")
  
oEditor.CreateRelativeCS Array("NAME:RelativeCSParameters", "CoordinateSystemID:=",  _
  -1, "OriginX:=", "0mm", "OriginY:=", "0mm", "OriginZ:=", "subH/2", "XAxisXvec:=",  _
  "1mm", "XAxisYvec:=", "0mm", "XAxisZvec:=", "0mm", "YAxisXvec:=", "0mm", "YAxisYvec:=",  _
  "1mm", "YAxisZvec:=", "0mm"), Array("NAME:Attributes", "Name:=", "Mid_CS")
  
oEditor.CreateRelativeCS Array("NAME:RelativeCSParameters", "CoordinateSystemID:=",  _
  -1, "OriginX:=", "0mm", "OriginY:=", "0mm", "OriginZ:=", "subH/2", "XAxisXvec:=",  _
  "1mm", "XAxisYvec:=", "0mm", "XAxisZvec:=", "0mm", "YAxisXvec:=", "0mm", "YAxisYvec:=",  _
  "1mm", "YAxisZvec:=", "0mm"), Array("NAME:Attributes", "Name:=", "Top_CS")  
  

  
  
oEditor.SetWCS Array("NAME:SetWCS Parameter", "Working Coordinate System:=",  _
  "Bottom_CS") 
     
oEditor.CreateUserDefinedPart Array("NAME:UserDefinedPrimitiveParameters", "DllName:=",  _
  UDP_name, "Version:=", "", "NoOfParameters:=", 9, "Library:=", "userlib", Array("NAME:ParamVector", Array("NAME:Pair", "Name:=",  _
  "Wslot", "Value:=", "Wslot"), Array("NAME:Pair", "Name:=", "Lfeed", "Value:=", "Lfeed"), Array("NAME:Pair", "Name:=",  _
  "Ltaper", "Value:=", "Ltaper"), Array("NAME:Pair", "Name:=", "N", "Value:=", "Nsteps", "ParamType:=",  _
  "IntParam"), Array("NAME:Pair", "Name:=", "Wtotal", "Value:=", "Wtotal"), Array("NAME:Pair", "Name:=",  _
  "Ltotal", "Value:=", "Ltotal"), Array("NAME:Pair", "Name:=", "Wbalun", "Value:=",  _
  "Wbalun"), Array("NAME:Pair", "Name:=", "Lbalun", "Value:=", "Lbalun"), Array("NAME:Pair", "Name:=",  _
  "Wtaper", "Value:=", "Wtaper"))), Array("NAME:Attributes", "Name:=", "Vivaldi_bot", "Flags:=",  _
  "", "Color:=", "(255 128 65)", "Transparency:=", 0.3, "PartCoordinateSystem:=",  _
  "Bottom_CS", "MaterialName:=", "pec", "SolveInside:=", false)
  
 oEditor.CreatePolyline Array("NAME:PolylineParameters", "IsPolylineCovered:=", true, "IsPolylineClosed:=",  _
  false, Array("NAME:PolylinePoints", Array("NAME:PLPoint", "X:=", "(Ltotal-Ltaper-Lfeed+feed_offset)", "Y:=",  _
  "-Wslot/2", "Z:=", "0"), Array("NAME:PLPoint", "X:=", "(Ltotal-Ltaper-Lfeed+feed_offset)", "Y:=",  _
  "Wslot/2", "Z:=", "0")), Array("NAME:PolylineSegments", Array("NAME:PLSegment", "SegmentType:=",  _
  "Line", "StartIndex:=", 0, "NoOfPoints:=", 2)), Array("NAME:PolylineXSection", "XSectionType:=",  _
  "None", "XSectionOrient:=", "Auto", "XSectionWidth:=", "0mm", "XSectionHeight:=",  _
  "0mm", "XSectionNumSegments:=", "0", "XSectionBendType:=", "Corner")), Array("NAME:Attributes", "Name:=",  _
  "Polyline1", "Flags:=", "", "Color:=", "(132 132 193)", "Transparency:=",  _
  0.8, "PartCoordinateSystem:=", "Global", "MaterialValue:=",  _
  "" & Chr(34) & "vacuum" & Chr(34) & "", "SolveInside:=", true)

oEditor.CreatePolyline Array("NAME:PolylineParameters", "IsPolylineCovered:=", true, "IsPolylineClosed:=",  _
  false, Array("NAME:PolylinePoints", Array("NAME:PLPoint", "X:=", "Ltotal", "Y:=",  _
  "-Wtaper/2", "Z:=", "0"), Array("NAME:PLPoint", "X:=", "Ltotal", "Y:=",  _
  "Wtaper/2", "Z:=", "0")), Array("NAME:PolylineSegments", Array("NAME:PLSegment", "SegmentType:=",  _
  "Line", "StartIndex:=", 0, "NoOfPoints:=", 2)), Array("NAME:PolylineXSection", "XSectionType:=",  _
  "None", "XSectionOrient:=", "Auto", "XSectionWidth:=", "0mm", "XSectionHeight:=",  _
  "0mm", "XSectionNumSegments:=", "0", "XSectionBendType:=", "Corner")), Array("NAME:Attributes", "Name:=",  _
  "Polyline2", "Flags:=", "", "Color:=", "(132 132 193)", "Transparency:=",  _
  0.8, "PartCoordinateSystem:=", "Global", "MaterialValue:=",  _
  "" & Chr(34) & "vacuum" & Chr(34) & "", "SolveInside:=", true)

 oEditor.Connect Array("NAME:Selections", "Selections:=", "Polyline1,Polyline2") 
  
oEditor.Subtract Array("NAME:Selections", "Blank Parts:=", "Vivaldi_bot", "Tool Parts:=",  _
  "Polyline1"), Array("NAME:SubtractParameters", "KeepOriginals:=", false) 
  
  
oEditor.SetWCS Array("NAME:SetWCS Parameter", "Working Coordinate System:=",  _
  "Top_CS")
  
    
oEditor.CreateUserDefinedPart Array("NAME:UserDefinedPrimitiveParameters", "DllName:=",  _
  UDP_name, "Version:=", "", "NoOfParameters:=", 9, "Library:=", "userlib", Array("NAME:ParamVector", Array("NAME:Pair", "Name:=",  _
  "Wslot", "Value:=", "Wslot"), Array("NAME:Pair", "Name:=", "Lfeed", "Value:=", "Lfeed"), Array("NAME:Pair", "Name:=",  _
  "Ltaper", "Value:=", "Ltaper"), Array("NAME:Pair", "Name:=", "N", "Value:=", "Nsteps", "ParamType:=",  _
  "IntParam"), Array("NAME:Pair", "Name:=", "Wtotal", "Value:=", "Wtotal"), Array("NAME:Pair", "Name:=",  _
  "Ltotal", "Value:=", "Ltotal"), Array("NAME:Pair", "Name:=", "Wbalun", "Value:=",  _
  "Wbalun"), Array("NAME:Pair", "Name:=", "Lbalun", "Value:=", "Lbalun"), Array("NAME:Pair", "Name:=",  _
  "Wtaper", "Value:=", "Wtaper"))), Array("NAME:Attributes", "Name:=", "Vivaldi_top", "Flags:=",  _
  "", "Color:=", "(255 128 65)", "Transparency:=", 0.3, "PartCoordinateSystem:=",  _
  "Top_CS", "MaterialName:=", "pec", "SolveInside:=", false) 

oEditor.CreatePolyline Array("NAME:PolylineParameters", "IsPolylineCovered:=", true, "IsPolylineClosed:=",  _
  false, Array("NAME:PolylinePoints", Array("NAME:PLPoint", "X:=", "(Ltotal-Ltaper-Lfeed+feed_offset)", "Y:=",  _
  "-Wslot/2", "Z:=", "0"), Array("NAME:PLPoint", "X:=", "(Ltotal-Ltaper-Lfeed+feed_offset)", "Y:=",  _
  "Wslot/2", "Z:=", "0")), Array("NAME:PolylineSegments", Array("NAME:PLSegment", "SegmentType:=",  _
  "Line", "StartIndex:=", 0, "NoOfPoints:=", 2)), Array("NAME:PolylineXSection", "XSectionType:=",  _
  "None", "XSectionOrient:=", "Auto", "XSectionWidth:=", "0mm", "XSectionHeight:=",  _
  "0mm", "XSectionNumSegments:=", "0", "XSectionBendType:=", "Corner")), Array("NAME:Attributes", "Name:=",  _
  "Polyline3", "Flags:=", "", "Color:=", "(132 132 193)", "Transparency:=",  _
  0.8, "PartCoordinateSystem:=", "Top_CS", "MaterialValue:=",  _
  "" & Chr(34) & "vacuum" & Chr(34) & "", "SolveInside:=", true)

oEditor.CreatePolyline Array("NAME:PolylineParameters", "IsPolylineCovered:=", true, "IsPolylineClosed:=",  _
  false, Array("NAME:PolylinePoints", Array("NAME:PLPoint", "X:=", "Ltotal", "Y:=",  _
  "-Wtaper/2", "Z:=", "0"), Array("NAME:PLPoint", "X:=", "Ltotal", "Y:=",  _
  "Wtaper/2", "Z:=", "0")), Array("NAME:PolylineSegments", Array("NAME:PLSegment", "SegmentType:=",  _
  "Line", "StartIndex:=", 0, "NoOfPoints:=", 2)), Array("NAME:PolylineXSection", "XSectionType:=",  _
  "None", "XSectionOrient:=", "Auto", "XSectionWidth:=", "0mm", "XSectionHeight:=",  _
  "0mm", "XSectionNumSegments:=", "0", "XSectionBendType:=", "Corner")), Array("NAME:Attributes", "Name:=",  _
  "Polyline4", "Flags:=", "", "Color:=", "(132 132 193)", "Transparency:=",  _
  0.8, "PartCoordinateSystem:=", "Top_CS", "MaterialValue:=",  _
  "" & Chr(34) & "vacuum" & Chr(34) & "", "SolveInside:=", true)

  oEditor.Connect Array("NAME:Selections", "Selections:=", "Polyline3,Polyline4") 
  
oEditor.Subtract Array("NAME:Selections", "Blank Parts:=", "Vivaldi_top", "Tool Parts:=",  _
  "Polyline3"), Array("NAME:SubtractParameters", "KeepOriginals:=", false)  


Set oModule = oDesign.GetModule("BoundarySetup")  

oModule.AssignPerfectE Array("NAME:PerfE1", "Objects:=", Array("Vivaldi_bot"), "InfGroundPlane:=", false)
oModule.AssignPerfectE Array("NAME:PerfE2", "Objects:=", Array("Vivaldi_top"), "InfGroundPlane:=", false)

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Draws Stripline Feed

Set oEditor = oDesign.SetActiveEditor("3D Modeler")

oEditor.SetWCS Array("NAME:SetWCS Parameter", "Working Coordinate System:=",  _
  "Mid_CS")

oEditor.CreateRectangle Array("NAME:RectangleParameters", "CoordinateSystemID:=",  _
  -1, "IsCovered:=", true, "XStart:=", "Ltotal-Ltaper-Lfeed+feed_offset", "YStart:=", "-Lstrip", "ZStart:=",  _
  "0mm", "Width:=", "-(Ltotal-Ltaper-Lfeed+feed_offset)", "Height:=", "Wstrip", "WhichAxis:=", "Z"), Array("NAME:Attributes", "Name:=",  _
  "Stripline", "Flags:=", "", "Color:=", "(255 128 65)", "Transparency:=",  _
  0.2, "PartCoordinateSystem:=", "Mid_CS", "MaterialName:=",  _
  "pec", "SolveInside:=", false)

oEditor.CreateRectangle Array("NAME:RectangleParameters", "CoordinateSystemID:=",  _
  -1, "IsCovered:=", true, "XStart:=", "Ltotal-Ltaper-Lfeed+feed_offset", "YStart:=", "0", "ZStart:=",  _
  "0mm", "Width:=", "Wstrip", "Height:=", "-Lstrip", "WhichAxis:=", "Z"), Array("NAME:Attributes", "Name:=",  _
  "Stripline2", "Flags:=", "", "Color:=", "(255 128 65)", "Transparency:=",  _
  0.2, "PartCoordinateSystem:=", "Mid_CS", "MaterialName:=",  _
  "pec", "SolveInside:=", false)


if Lstrip_offset > 0 Then
oEditor.CreateRectangle Array("NAME:RectangleParameters", "CoordinateSystemID:=",  _
  -1, "IsCovered:=", true, "XStart:=", "Ltotal-Ltaper-Lfeed+feed_offset", "YStart:=", "0", "ZStart:=",  _
  "0mm", "Width:=", "Wstrip", "Height:=", "Lstrip_offset", "WhichAxis:=", "Z"), Array("NAME:Attributes", "Name:=",  _
  "Stripline_offset", "Flags:=", "", "Color:=", "(255 128 65)", "Transparency:=",  _
  0.2, "PartCoordinateSystem:=", "Mid_CS", "MaterialName:=",  _
  "pec", "SolveInside:=", false)

oEditor.Unite Array("NAME:Selections", "Selections:=",  _
  "Stripline,Stripline2,Stripline_offset"), Array("NAME:UniteParameters", "CoordinateSystemID:=",  _
  -1, "KeepOriginals:=", false)

else

oEditor.Unite Array("NAME:Selections", "Selections:=",  _
  "Stripline,Stripline2"), Array("NAME:UniteParameters", "CoordinateSystemID:=",  _
  -1, "KeepOriginals:=", false)

end If

Set oModule = oDesign.GetModule("BoundarySetup")  

oModule.AssignPerfectE Array("NAME:Stripline", "Objects:=", Array("Stripline"), "InfGroundPlane:=", false)


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Draws and Assigns port

Set oEditor = oDesign.SetActiveEditor("3D Modeler")

oEditor.SetWCS Array("NAME:SetWCS Parameter", "Working Coordinate System:=",  _
  "Bottom_CS")
  



 

     oEditor.CreateRectangle Array("NAME:RectangleParameters", "CoordinateSystemID:=",  _
  -1, "IsCovered:=", true, "XStart:=", "0mm", "YStart:=", "-Lstrip", "ZStart:=",  _
  "0mm", "Width:=", "-Wstrip/4", "Height:=", "Wstrip", "WhichAxis:=", "Z"), Array("NAME:Attributes", "Name:=",  _
  "port_feed1", "Flags:=", "", "Color:=", "(255 128 65)", "Transparency:=",  _
  0.8, "PartCoordinateSystem:=", "Bottom_CS", "MaterialName:=",  _
  "vacuum", "SolveInside:=", true)
  
  
       oEditor.CreateRectangle Array("NAME:RectangleParameters", "CoordinateSystemID:=",  _
  -1, "IsCovered:=", true, "XStart:=", "0mm", "YStart:=", "-Lstrip", "ZStart:=",  _
  "SubH", "Width:=", "-Wstrip/4", "Height:=", "Wstrip", "WhichAxis:=", "Z"), Array("NAME:Attributes", "Name:=",  _
  "port_feed2", "Flags:=", "", "Color:=", "(255 128 65)", "Transparency:=",  _
  0.8, "PartCoordinateSystem:=", "Bottom_CS", "MaterialName:=",  _
  "vacuum", "SolveInside:=", true)

    oEditor.CreateRectangle Array("NAME:RectangleParameters", "CoordinateSystemID:=",  _
  -1, "IsCovered:=", true, "XStart:=", "-Wstrip/4", "YStart:=", "-Lstrip", "ZStart:=",  _
  "0mm", "Width:=", "Wstrip", "Height:=", "subH", "WhichAxis:=", "X"), Array("NAME:Attributes", "Name:=",  _
  "port_feed3", "Flags:=", "", "Color:=", "(255 128 65)", "Transparency:=",  _
  0.8, "PartCoordinateSystem:=", "Bottom_CS", "MaterialName:=",  _
  "vacuum", "SolveInside:=", true)
  
    oEditor.CreateRectangle Array("NAME:RectangleParameters", "CoordinateSystemID:=",  _
  -1, "IsCovered:=", true, "XStart:=", "0mm", "YStart:=", "-Lstrip", "ZStart:=",  _
  "SubH/2", "Width:=", "-Wstrip/4", "Height:=", "Wstrip", "WhichAxis:=", "Z"), Array("NAME:Attributes", "Name:=",  _
  "port1", "Flags:=", "", "Color:=", "(255 128 65)", "Transparency:=",  _
  0.8, "PartCoordinateSystem:=", "Bottom_CS", "MaterialName:=",  _
  "vacuum", "SolveInside:=", true)

Dim faceid
faceid = oEditor.GetFaceByPosition(Array("NAME:FaceParameters","BodyName:=", "port1", "XPosition:=", "-Wstrip/8", "YPosition:=","-Lstrip+Wstrip/2", "ZPosition:=", "subH/2"))


Set oModule = oDesign.GetModule("BoundarySetup")

 

oModule.AssignPerfectE Array("NAME:Port_Feed", "Objects:=", Array("port_feed1","port_feed2","port_feed3"), "InfGroundPlane:=", false)

oModule.AssignLumpedPort Array("NAME:Port1", "Objects:=", Array("port1"), "RenormalizeAllTerminals:=",  _
  true, "TerminalIDList:=", Array())
oModule.AutoIdentifyTerminals Array("NAME:ReferenceConductors", "port_feed3"), "Port1", true



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
stop_freq = 1*solution_freq

           

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




