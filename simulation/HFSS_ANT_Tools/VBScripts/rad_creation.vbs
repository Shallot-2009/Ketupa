' ----------------------------------------------------------------------
' -Create Parameterized Radiation bounding box in HFSS
' Based on the adaptive frequency.
' -Create a virtual object inside the radiation box and seed it to improve the far-field accuracy.
' -Set up far-field using internal virtual object as the integration surface for far-field calculation
' Written by D. Crawford, Ansoft Corp., 25-Aug-08 dcrawford@ansoft.com
' modified by A. Sligar on Jan 26 09 to be included in antenna design kit
' now takes in design variable names for xyz negative and positive extents. This will allow
' geometry size to be parameterized without intsecting with the radiation box for certain variations
' ------------------------------------------------------------------------------------------------
Dim oAnsoftApp
Dim oDesktop
Dim oProject
Dim oDesign
Dim oEditor
Dim oModule
Set oAnsoftApp = CreateObject("AnsoftHfss.HfssScriptInterface")
Set oDesktop = oAnsoftApp.GetAppDesktop()
oDesktop.RestoreWindow
Set oProject = oDesktop.GetActiveProject()
Set oDesign = oProject.GetActiveDesign()



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

extent_x_pos = args(0)

extent_x_neg = args(1)

extent_y_pos = args(2)
extent_y_neg = args(3)

extent_z_pos = args(4)
extent_z_neg = args(5)



units = args(6)

israd = args(7)

'msgbox(israd)
'if ubound(args) = 8 then
low_freq = args(8)
'else
'low_freq = 0 
'end if

Dim boundingBox
Dim lightSpeed
dim PML_Thickness

lightSpeed = 299792458 'm/s

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

lightSpeed = lightSpeed * unit_conversion







If oDesign Is Nothing Then
	Dim Designs : Set Designs = oProject.GetDesigns()
	If Designs.Count = 1 Then
		Set oDesign=Designs(0)	
	Else
		Err.Description = "Please select an HFSS design before running the script"
		MsgBox(Err.Description)
		Err.Raise vbObjectError + 1
	End If
End If


Set oEditor = oDesign.SetActiveEditor("3D Modeler")


'oEditor.SetModelUnits(ARRAY("NAME:Units Paramter","Units:=",units,"Rescale:=",false))



Set oModule = oDesign.GetModule("AnalysisSetup")
' setupList: a list of analysis setups
' numberOfSetups: self explanatory
' minAdaptFreq: will contain the lowest adapt frequency in all setups
' intFaces is a list of face ID's that will be used to create the integration surface for far-field calculations
' addExtend is the distance by which the outer radiation box will be extended beyond the 3D model.
' intScale is the scale factor used for the reduced size box on which the far-field integration is performed
'           for example, the box size will be x_extent_size/intScale.  "intScale" shows up as an integer in 
'           the HFSS model
Dim setupList, numberOfSetups, setup_index, count, adaptFreq, minAdaptFreq
Dim addExtend, intFaces(5), intFaces1(5), intFaces1VB(5) ,intScale, seedLen
intScale=8
count=0




setupList=oModule.GetSetups()
numberOfSetups = Ubound(setupList)
If numberOfSetups >= 0 Then  ' only proceed if setup exists
	Set oModule=oDesign.GetModule("Solutions")
	Do  ' Loop to determine the minimum adapt frequency
		If count=0 Then
			minAdaptFreq = oModule.GetAdaptiveFreq(setupList(count))
			setup_index=count
		Else
			adaptFreq = oModule.GetAdaptiveFreq(setupList(count))
			If adaptFreq < minAdaptFreq Then
				minAdaptFreq = adaptFreq
				setup_index=count
			End If
		End If
		count = count +1
	Loop while count < numberOfSetups
	
	if low_freq > 0 then
	minAdaptFreq = low_freq*1e9
	end if

	seedLen=CStr(Round(lightSpeed/minAdaptFreq/4,4)) 'Length for seeding radiatioun surfaces for far-field integration
	
	if israd="Rad" then
		addExtend=Round(lightSpeed/minAdaptFreq/3,4) 'Extend airbox by lambda/3
		addExtend_VB=Round(lightSpeed/minAdaptFreq/10,4) 'Extend virtual airbox by lambda/10
	else
	  PML_Thickness = Round(lightSpeed/minAdaptFreq/3,4)
		addExtend=Round(lightSpeed/minAdaptFreq/8,4) 'Extend airbox by lambda/8
		addExtend_VB=Round(lightSpeed/minAdaptFreq/10,4) 'Extend virtual airbox by lambda/10
	end if






	
if israd="PML" or israd ="Rad" then        'this will do the air box for PML or ABC, FEBI Will be done seperatly

oDesign.ChangeProperty Array("NAME:AllTabs", Array("NAME:LocalVariableTab", Array("NAME:PropServers", "LocalVariables"), Array("NAME:NewProps", _
Array("NAME:--Air Box", "PropType:=", "SeparatorProp", "UserDef:=", true, "Value:=", ""), _
Array("NAME:Airbox_dist", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", addExtend & units), _
Array("NAME:--Virtual Object Radiation Surface", "PropType:=", "SeparatorProp", "UserDef:=", true, "Value:=", ""), _
Array("NAME:VirtualObject_dist", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", addExtend_VB & units))))


if israd = "PML" then

oDesign.ChangeProperty Array("NAME:AllTabs", Array("NAME:LocalVariableTab", Array("NAME:PropServers",  _
  "LocalVariables"), Array("NAME:ChangedProps", Array("NAME:Airbox_dist", "ReadOnly:=",  _
  false))))

end if





	
	
' --------------------- Draw box with radiation boundary ----------------------------------------

dim xpos, xsize, ypos, ysize, zpos, zsize


xpos = extent_x_neg & " - Airbox_dist"
xsize = "abs(" & extent_x_neg & "-" & extent_x_pos & ") + 2*Airbox_dist"

ypos = extent_y_neg & " - Airbox_dist"
ysize = "abs(" & extent_y_neg & "-" & extent_y_pos & ") + 2*Airbox_dist"

zpos = extent_z_neg & " - Airbox_dist"
zsize = "abs(" & extent_z_neg & "-" & extent_z_pos & ") + 2*Airbox_dist"



Set oEditor = oDesign.SetActiveEditor("3D Modeler")


	oEditor.CreateBox Array("NAME:BoxParameters", "CoordinateSystemID:=", -1, "XPosition:=",  _
		xpos, "YPosition:=", ypos, "ZPosition:=",	 _
		zpos , "XSize:=", xsize, "YSize:=", ysize, "ZSize:=",  _
		zsize),	Array("NAME:Attributes", "Name:=", "AirBox", "Flags:=",	"",	"Color:=",	_
		"(128 128 255)", "Transparency:=", 0.9, "PartCoordinateSystem:=",	 _
		"Global", "MaterialName:=", "vacuum", "SolveInside:=", true)

	oEditor.ChangeProperty Array("NAME:AllTabs", Array("NAME:Geometry3DAttributeTab", Array("NAME:PropServers",  _
		"AirBox"), Array("NAME:ChangedProps", Array("NAME:Display Wireframe", "Value:=", true))))
		

'if israd = "PML" then    'this is no longer needed starting with v12 of HFSS and the pmls will move with faces'


''  oEditor.PurgeHistory Array("NAME:Selections", "Selections:=", "AirBox", "NewPartsModelFlag:=", "Model")


'end if


		
' --------------- Draw inner box to use for far-field integration -------------------------------		
dim xpos_vb, xsize_vb, ypos_vb, ysize_vb, zpos_vb, zsize_vb
xpos_vb = extent_x_neg & " - VirtualObject_dist"
xsize_vb = "abs(" & extent_x_neg & "-" & extent_x_pos & ") + 2*VirtualObject_dist"

ypos_vb = extent_y_neg & " - VirtualObject_dist"
ysize_vb = "abs(" & extent_y_neg & "-" & extent_y_pos & ") + 2*VirtualObject_dist"

zpos_vb = extent_z_neg & " - VirtualObject_dist"
zsize_vb = "abs(" & extent_z_neg & "-" & extent_z_pos & ") + 2*VirtualObject_dist"

		oEditor.CreateBox Array("NAME:BoxParameters", "CoordinateSystemID:=", -1, "XPosition:=",  _
		xpos_vb, "YPosition:=", ypos_vb, "ZPosition:=",	 _
		zpos_vb , "XSize:=", xsize_vb, "YSize:=", ysize_vb, "ZSize:=",  _
		zsize_vb),	Array("NAME:Attributes", "Name:=", "VirtualRadiation", "Flags:=",	"",	"Color:=",	_
		"(128 128 255)", "Transparency:=", 0.9, "PartCoordinateSystem:=",	 _
		"Global", "MaterialName:=", "vacuum", "SolveInside:=", true)

		
	oEditor.ChangeProperty Array("NAME:AllTabs", Array("NAME:Geometry3DAttributeTab", Array("NAME:PropServers",  _
		"VirtualRadiation"), Array("NAME:ChangedProps", Array("NAME:Display Wireframe", "Value:=", true))))
	
  
  



 
 dim xpos_face1, ypos_face1, zpos_face1, xpos_face2, ypos_face2, zpos_face2, xpos_face3, ypos_face3, zpos_face3
 dim xpos_face4, ypos_face4, zpos_face4, xpos_face5, ypos_face5, zpos_face5, xpos_face6, ypos_face6, zpos_face6
xpos_face1 = xpos
ypos_face1 = extent_y_neg & " - Airbox_dist" & "+ (abs(" & extent_y_neg & "-" & extent_y_pos & ") + 2*Airbox_dist)/2"
zpos_face1 = extent_z_neg & " - Airbox_dist" & "+ (abs(" & extent_z_neg & "-" & extent_z_pos & ") + 2*Airbox_dist)/2"

xpos_face2 = "(" & extent_x_pos & " + Airbox_dist)"
ypos_face2 = extent_y_neg & " - Airbox_dist" & "+ (abs(" & extent_y_neg & "-" & extent_y_pos & ") + 2*Airbox_dist)/2"
zpos_face2 = extent_z_neg & " - Airbox_dist" & "+ (abs(" & extent_z_neg & "-" & extent_z_pos & ") + 2*Airbox_dist)/2"

xpos_face3 = extent_x_neg & " - Airbox_dist" & "+ (abs(" & extent_x_neg & "-" & extent_x_pos & ") + 2*Airbox_dist)/2"
ypos_face3 = ypos
zpos_face3 = extent_z_neg & " - Airbox_dist" & "+ (abs(" & extent_z_neg & "-" & extent_z_pos & ") + 2*Airbox_dist)/2"

xpos_face4 = extent_x_neg & " - Airbox_dist" & "+ (abs(" & extent_x_neg & "-" & extent_x_pos & ") + 2*Airbox_dist)/2"
ypos_face4 = "(" & extent_y_pos & " + Airbox_dist)"
zpos_face4 = extent_z_neg & " - Airbox_dist" & "+ (abs(" & extent_z_neg & "-" & extent_z_pos & ") + 2*Airbox_dist)/2"

xpos_face5 = extent_x_neg & " - Airbox_dist" & "+ (abs(" & extent_x_neg & "-" & extent_x_pos & ") + 2*Airbox_dist)/2"
ypos_face5 = extent_y_neg & " - Airbox_dist" & "+ (abs(" & extent_y_neg & "-" & extent_y_pos & ") + 2*Airbox_dist)/2"
zpos_face5 = extent_z_neg & " - Airbox_dist"

xpos_face6 = extent_x_neg & " - Airbox_dist" & "+ (abs(" & extent_x_neg & "-" & extent_x_pos & ") + 2*Airbox_dist)/2"
ypos_face6 = extent_y_neg & " - Airbox_dist" & "+ (abs(" & extent_y_neg & "-" & extent_y_pos & ") + 2*Airbox_dist)/2"
zpos_face6 = "(" & extent_z_pos & " + Airbox_dist)"



 dim xpos_face1VB, ypos_face1VB, zpos_face1VB, xpos_face2VB, ypos_face2VB, zpos_face2VB, xpos_face3VB, ypos_face3VB, zpos_face3VB
 dim xpos_face4VB, ypos_face4VB, zpos_face4VB, xpos_face5VB, ypos_face5VB, zpos_face5VB, xpos_face6VB, ypos_face6VB, zpos_face6VB
xpos_face1VB = xpos_vb
ypos_face1VB = extent_y_neg & " - VirtualObject_dist" & "+ (abs(" & extent_y_neg & "-" & extent_y_pos & ") + 2*VirtualObject_dist)/2"
zpos_face1VB = extent_z_neg & " - VirtualObject_dist" & "+ (abs(" & extent_z_neg & "-" & extent_z_pos & ") + 2*VirtualObject_dist)/2"

xpos_face2VB = "(" & extent_x_pos & " + VirtualObject_dist)"

ypos_face2VB = extent_y_neg & " - VirtualObject_dist" & "+ (abs(" & extent_y_neg & "-" & extent_y_pos & ") + 2*VirtualObject_dist)/2"
zpos_face2VB = extent_z_neg & " - VirtualObject_dist" & "+ (abs(" & extent_z_neg & "-" & extent_z_pos & ") + 2*VirtualObject_dist)/2"

xpos_face3VB = extent_x_neg & " - VirtualObject_dist" & "+ (abs(" & extent_x_neg & "-" & extent_x_pos & ") + 2*VirtualObject_dist)/2"
ypos_face3VB = ypos_vb
zpos_face3VB = extent_z_neg & " - VirtualObject_dist" & "+ (abs(" & extent_z_neg & "-" & extent_z_pos & ") + 2*VirtualObject_dist)/2"

xpos_face4VB = extent_x_neg & " - VirtualObject_dist" & "+ (abs(" & extent_x_neg & "-" & extent_x_pos & ") + 2*VirtualObject_dist)/2"
ypos_face4VB = "(" & extent_y_pos & " + VirtualObject_dist)"
zpos_face4VB = extent_z_neg & " - VirtualObject_dist" & "+ (abs(" & extent_z_neg & "-" & extent_z_pos & ") + 2*VirtualObject_dist)/2"

xpos_face5VB = extent_x_neg & " - VirtualObject_dist" & "+ (abs(" & extent_x_neg & "-" & extent_x_pos & ") + 2*VirtualObject_dist)/2"
ypos_face5VB = extent_y_neg & " - VirtualObject_dist" & "+ (abs(" & extent_y_neg & "-" & extent_y_pos & ") + 2*VirtualObject_dist)/2"
zpos_face5VB = zpos_vb

xpos_face6VB = extent_x_neg & " - VirtualObject_dist" & "+ (abs(" & extent_x_neg & "-" & extent_x_pos & ") + 2*VirtualObject_dist)/2"
ypos_face6VB = extent_y_neg & " - VirtualObject_dist" & "+ (abs(" & extent_y_neg & "-" & extent_y_pos & ") + 2*VirtualObject_dist)/2"
zpos_face6VB = "(" & extent_z_pos & " + VirtualObject_dist)"


  	

	Set oModule=oDesign.GetModule("BoundarySetup")

	if israd="Rad" then
			oModule.AssignRadiation Array("NAME:Rad","Objects:=",Array("AirBox"))
			intFaces1(0) = oEditor.GetFaceByPosition(Array("NAME:FaceParameters","BodyName:=", "AirBox", "XPosition:=", xpos_face1, "YPosition:=",ypos_face1, "ZPosition:=", zpos_face1))
      intFaces1(1) = oEditor.GetFaceByPosition(Array("NAME:FaceParameters","BodyName:=", "AirBox", "XPosition:=", xpos_face2, "YPosition:=",ypos_face2, "ZPosition:=", zpos_face2))
      intFaces1(2) = oEditor.GetFaceByPosition(Array("NAME:FaceParameters","BodyName:=", "AirBox", "XPosition:=", xpos_face3, "YPosition:=",ypos_face3, "ZPosition:=", zpos_face3))
      intFaces1(3) = oEditor.GetFaceByPosition(Array("NAME:FaceParameters","BodyName:=", "AirBox", "XPosition:=", xpos_face4, "YPosition:=",ypos_face4, "ZPosition:=", zpos_face4))
      intFaces1(4) = oEditor.GetFaceByPosition(Array("NAME:FaceParameters","BodyName:=", "AirBox", "XPosition:=", xpos_face5, "YPosition:=",ypos_face5, "ZPosition:=", zpos_face5))
      intFaces1(5) = oEditor.GetFaceByPosition(Array("NAME:FaceParameters","BodyName:=", "AirBox", "XPosition:=", xpos_face6, "YPosition:=",ypos_face6, "ZPosition:=", zpos_face6))
      
	else
      intFaces1(0) = oEditor.GetFaceByPosition(Array("NAME:FaceParameters","BodyName:=", "AirBox", "XPosition:=", xpos_face1, "YPosition:=",ypos_face1, "ZPosition:=", zpos_face1))
			intFaces1(1) = oEditor.GetFaceByPosition(Array("NAME:FaceParameters","BodyName:=", "AirBox", "XPosition:=", xpos_face2, "YPosition:=",ypos_face2, "ZPosition:=", zpos_face2))
      intFaces1(2) = oEditor.GetFaceByPosition(Array("NAME:FaceParameters","BodyName:=", "AirBox", "XPosition:=", xpos_face3, "YPosition:=",ypos_face3, "ZPosition:=", zpos_face3))
      intFaces1(3) = oEditor.GetFaceByPosition(Array("NAME:FaceParameters","BodyName:=", "AirBox", "XPosition:=", xpos_face4, "YPosition:=",ypos_face4, "ZPosition:=", zpos_face4))
      intFaces1(4) = oEditor.GetFaceByPosition(Array("NAME:FaceParameters","BodyName:=", "AirBox", "XPosition:=", xpos_face5, "YPosition:=",ypos_face5, "ZPosition:=", zpos_face5))
      intFaces1(5) = oEditor.GetFaceByPosition(Array("NAME:FaceParameters","BodyName:=", "AirBox", "XPosition:=", xpos_face6, "YPosition:=",ypos_face6, "ZPosition:=", zpos_face6))
		oModule.CreatePML Array("UserDrawnGroup:=", false,_
		"PMLFaces:=",intFaces1, "CreateJoiningObjs:=", true,_
		"Thickness:=", PML_Thickness & units, "RadDist:=", addExtend & units,_ 
		"UseFreq:=", true, "MinFreq:=", minAdaptFreq & "Hz")
	end if
'---------Select the faces of the inner box -------------------   
     intFaces1VB(0) = oEditor.GetFaceByPosition(Array("NAME:FaceParameters","BodyName:=", "VirtualRadiation", "XPosition:=", xpos_face1VB, "YPosition:=",ypos_face1VB, "ZPosition:=", zpos_face1VB))
			intFaces1VB(1) = oEditor.GetFaceByPosition(Array("NAME:FaceParameters","BodyName:=", "VirtualRadiation", "XPosition:=", xpos_face2VB, "YPosition:=",ypos_face2VB, "ZPosition:=", zpos_face2VB))
      intFaces1VB(2) = oEditor.GetFaceByPosition(Array("NAME:FaceParameters","BodyName:=", "VirtualRadiation", "XPosition:=", xpos_face3VB, "YPosition:=",ypos_face3VB, "ZPosition:=", zpos_face3VB))
      intFaces1VB(3) = oEditor.GetFaceByPosition(Array("NAME:FaceParameters","BodyName:=", "VirtualRadiation", "XPosition:=", xpos_face4VB, "YPosition:=",ypos_face4VB, "ZPosition:=", zpos_face4VB))
      intFaces1VB(4) = oEditor.GetFaceByPosition(Array("NAME:FaceParameters","BodyName:=", "VirtualRadiation", "XPosition:=", xpos_face5VB, "YPosition:=",ypos_face5VB, "ZPosition:=", zpos_face5VB))
      intFaces1VB(5) = oEditor.GetFaceByPosition(Array("NAME:FaceParameters","BodyName:=", "VirtualRadiation", "XPosition:=", xpos_face6VB, "YPosition:=",ypos_face6VB, "ZPosition:=", zpos_face6VB))
     
'------ First apply seeding to the faces (lambda/6) seedLen----------
	'Set oModule = oDesign.GetModule("MeshSetup")
	'oModule.AssignLengthOp Array("NAME:seedVirtualBox","RefineInside:=", false, "Faces:=", intFaces1VB, _
	'"RestrictElem:=", false, "NumMaxElem:=", "1000", "RestrictLength:=",  _
	'true, "MaxLength:=", seedLen & units)
'------- Now create a face list for the far-field integration ----------------        
	oEditor.CreateEntityList Array("NAME:GeometryEntityListParameters", "EntityType:=",  _
	"Face", "EntityList:=", intFaces1VB), Array("NAME:Attributes", "Name:=",  _
	"radFaces")      
'------- Create far-field setup ------------------------------
	Set oModule = oDesign.GetModule("RadField")
	oModule.InsertFarFieldSphereSetup Array("NAME:infSphere", "UseCustomRadiationSurface:=",  _
	true, "CustomRadiationSurface:=", "radFaces", "ThetaStart:=", "-180deg", "ThetaStop:=",  _
	"180deg", "ThetaStep:=", "2deg", "PhiStart:=", "0deg", "PhiStop:=", "180deg", "PhiStep:=",  _
	"5deg", "UseLocalCS:=", false)
	
'--------- The following lines are intended to create an output variable (peak directitity)
'          and use this as an output convergence criterion.  I have not yet found an
'          easy way to "Edit" the setup.
'-----------------------------------------------------------
'------ Create output variable to use for convergence -----------------------------
'	Set oModule = oDesign.GetModule("OutputVariable")
'	oModule.CreateOutputVariable "peak_Dir", "mag(PeakDirectivity)",  _
'    setupList(setup_index) &" : LastAdaptive", "Far Fields", Array("Context:=", "infSphere")
'------ Add directivity as convergence criteria for antenna ------------------------------
'    Set oModule = oDesign.GetModule("AnalysisSetup")
'    oModule.EditSetup setupList(setup_index), ARRAY("NAME:" &setupList(setup_index),"UseConvOutputVariable:=", true, "ConvOutvar:=",  _
'    "peak_Dir", "ConvOutvarRelativeDelta:=", true, "ConvOutvarMaxDelta:=", 1, _
'    "ConvOutvarContext:=", "infSphere", "ConvOutvarIntrinsics:=", "Phi=" & Chr(39) & "0deg" & Chr(39) & _
'    " Theta=" & Chr(39) & "-180deg" & Chr(39) & "")

End If

end if ' this is to create rad boundary using ABC or PML. FEBI Will be done seperatly with Region Command



if israd="FEBI" then

FEBI_Offset = Round(lightSpeed/minAdaptFreq/10,4)

 oDesign.ChangeProperty Array("NAME:AllTabs", Array("NAME:LocalVariableTab", Array("NAME:PropServers", "LocalVariables"), Array("NAME:NewProps", _
Array("NAME:--Air Box", "PropType:=", "SeparatorProp", "UserDef:=", true, "Value:=", ""), _
Array("NAME:Airbox_dist", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", FEBI_Offset & units))))

FEBI_Offset = Round(lightSpeed/minAdaptFreq/10,4)

Set oEditor = oDesign.SetActiveEditor("3D Modeler")
oEditor.CreateRegion Array("NAME:RegionParameters", "+XPaddingType:=",  _
  "Absolute Offset", "+XPadding:=", "Airbox_dist", "-XPaddingType:=", "Absolute Offset", "-XPadding:=",  _
  "Airbox_dist", "+YPaddingType:=", "Absolute Offset", "+YPadding:=", "Airbox_dist", "-YPaddingType:=",  _
  "Absolute Offset", "-YPadding:=", "Airbox_dist", "+ZPaddingType:=", "Absolute Offset", "+ZPadding:=",  _
  "Airbox_dist", "-ZPaddingType:=", "Absolute Offset", "-ZPadding:=", "Airbox_dist"), Array("NAME:Attributes", "Name:=",  _
  "Region", "Flags:=", "Wireframe#", "Color:=", "(255 0 0)", "Transparency:=",  _
  0.800000011920929, "PartCoordinateSystem:=", "Global", "UDMId:=", "", "MaterialValue:=",  _
  "" & Chr(34) & "vacuum" & Chr(34) & "", "SolveInside:=", true)
  
Set oModule = oDesign.GetModule("BoundarySetup")
oModule.AssignRadiation Array("NAME:Rad1", "Objects:=", Array("Region"), "IsIncidentField:=",  _
  false, "IsEnforcedField:=", false, "IsFssReference:=", false, "IsForPML:=",  _
  false, "UseAdaptiveIE:=", true, "IncludeInPostproc:=", true)  
  

	Set oModule = oDesign.GetModule("RadField")
	oModule.InsertFarFieldSphereSetup Array("NAME:infSphere", "UseCustomRadiationSurface:=",  _
	false, "ThetaStart:=", "-180deg", "ThetaStop:=",  _
	"180deg", "ThetaStep:=", "2deg", "PhiStart:=", "0deg", "PhiStop:=", "180deg", "PhiStep:=",  _
	"5deg", "UseLocalCS:=", false)

End if






' this resizes the view to fit all on screen
if extent_z_pos <> "ConeHeight" then
  Set oEditor = oDesign.SetActiveEditor("3D Modeler")
  oEditor.SetModelUnits Array("NAME:Units Parameter", "Units:=", units, "Rescale:=", false)
end if


Set oDesktop = oAnsoftApp.GetAppDesktop()
oDesktop.CloseAllWindows()
Set oModeler = oDesign.SetActiveEditor("3D Modeler")

oEditor.ShowWindow


sol_type = oDesign.GetSolutionType



Set oModule = oDesign.GetModule("BoundarySetup")
numexcite = oModule.GetNumExcitations()

Set oModule = oDesign.GetModule("ReportSetup")

if sol_type = "driven modal" OR sol_type =  "DrivenModal" then

  if numexcite = 1 then
    oModule.CreateReport "Return Loss","Modal Solution Data", "XY Plot","Setup1 : Sweep1",_
    Array("Domain:=", "Sweep"), Array("Freq:=", Array("All")),Array("X Component:=", "Freq","Y Component:=",_
    Array("dB(S(1,1))")),Array()

    oModule.CreateReport "Input Impedance", "Modal Solution Data", "Smith Plot",_
    "Setup1 : Sweep1", Array("Domain:=", "Sweep"), Array("Freq:=", Array("All")), _
    Array("Polar Component:=", Array("S11")),Array()
    



    
    
  end if
  
  if numexcite = 2 then
  
    oModule.CreateReport "Return Loss","Modal Solution Data", "XY Plot","Setup1 : Sweep1",_
    Array("Domain:=", "Sweep"), Array("Freq:=", Array("All")),Array("X Component:=", "Freq","Y Component:=",_
    Array("dB(S(1,1))","dB(S(2,2))")),Array()

    oModule.CreateReport "Input Impedance", "Modal Solution Data", "Smith Plot",_
    "Setup1 : Sweep1", Array("Domain:=", "Sweep"), Array("Freq:=", Array("All")), _
    Array("Polar Component:=", Array("S(1,1)","S(2,2)")),Array()
    

    
  end if
  
    if numexcite = 4 then
  
    oModule.CreateReport "Return Loss","Modal Solution Data", "XY Plot","Setup1 : Sweep1",_
    Array("Domain:=", "Sweep"), Array("Freq:=", Array("All")),Array("X Component:=", "Freq","Y Component:=",_
    Array("dB(S(1,1))","dB(S(2,2))","dB(S(3,3))","dB(S(4,4))")),Array()

    oModule.CreateReport "Input Impedance", "Modal Solution Data", "Smith Plot",_
    "Setup1 : Sweep1", Array("Domain:=", "Sweep"), Array("Freq:=", Array("All")), _
    Array("Polar Component:=", Array("S(1,1)","S(2,2)","S(3,3)","S(4,4)")),Array()
    

    
  end if
  
end if


if sol_type = "driven terminal" OR sol_type =  "DrivenTerminal" then

  if numexcite = 1 then
    oModule.CreateReport "Return Loss","Terminal Solution Data", "XY Plot","Setup1 : Sweep1",_
    Array("Domain:=", "Sweep"), Array("Freq:=", Array("All")),Array("X Component:=", "Freq","Y Component:=",_
    Array("dB(St(1,1))")),Array()

    oModule.CreateReport "Input Impedance", "Terminal Solution Data", "Smith Plot",_
    "Setup1 : Sweep1", Array("Domain:=", "Sweep"), Array("Freq:=", Array("All")), _
    Array("Polar Component:=", Array("S11")),Array()
    

    
    
  end if
  
  if numexcite = 2 then
  
    oModule.CreateReport "Return Loss","Terminal Solution Data", "XY Plot","Setup1 : Sweep1",_
    Array("Domain:=", "Sweep"), Array("Freq:=", Array("All")),Array("X Component:=", "Freq","Y Component:=",_
    Array("dB(St(1,1))","dB(St(2,2))")),Array()

    oModule.CreateReport "Input Impedance", "Terminal Solution Data", "Smith Plot",_
    "Setup1 : Sweep1", Array("Domain:=", "Sweep"), Array("Freq:=", Array("All")), _
    Array("Polar Component:=", Array("St(1,1)","St(2,2)")),Array()
    
    

    
    
  end if
  
    if numexcite = 4 then
  
    oModule.CreateReport "Return Loss","Terminal Solution Data", "XY Plot","Setup1 : Sweep1",_
    Array("Domain:=", "Sweep"), Array("Freq:=", Array("All")),Array("X Component:=", "Freq","Y Component:=",_
    Array("dB(St(1,1))","dB(St(2,2))","dB(St(3,3))","dB(St(4,4))")),Array()

    oModule.CreateReport "Input Impedance", "Terminal Solution Data", "Smith Plot",_
    "Setup1 : Sweep1", Array("Domain:=", "Sweep"), Array("Freq:=", Array("All")), _
    Array("Polar Component:=", Array("St(1,1)","St(2,2)","St(3,3)","St(4,4)")),Array()
    
    

    
    
  end if
  
end if






design_name = oDesign.GetName

name_CP_antennas = "EllipticalHorn_Antenna_ADKv1" & "ConicalHorn_Antenna_ADKv1" & "Sinuous_Antenna_ADKv1" & "SinuousConical_Antenna_ADKv1"_
& "Archimedean_Antenna_ADKv1" & "ArchimedeanConical_Antenna_ADKv1" & "LogSpiral_Antenna_ADKv1" & "LogSpiralConical_Antenna_ADKv1"_
& "CircularWG_Antenna_ADKv1"  & "Helix_Antenna_ADKv1"


if InStr(name_CP_antennas, design_name) <> 0 then
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

else
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



end if








Set oDesktop = oAnsoftApp.GetAppDesktop()
oDesktop.CloseAllWindows()
Set oModeler = oDesign.SetActiveEditor("3D Modeler")

oEditor.ShowWindow

Setlocale(locallang)







