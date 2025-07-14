# --------------------------------------------------------------
#  [Ansys ver]Ansys Electronics Desktop Version 2024.2.0
#  [Date]  Jun 27, 2025
#  [File name]Planar_InvertedF_Antenna
#  [Notes]  Planar_InvertedF_Antenna for 2.4GHz
#  [Author's Email]  3405802009@qq.com
# --------------------------------------------------------------
import os
import win32com.client
oAnsoftApp = win32com.client.Dispatch('AnsoftHfss.HfssScriptInterface')
oDesktop = oAnsoftApp.GetAppDesktop()
oDesktop.RestoreWindow()
## oProject = oDesktop.GetActiveProject()
## path = os.getcwd()#获取当前py文件所在文件夹
## filename = 'IFA1.hfss'
## fullname = os.path.join(path, filename)
## oDesktop.OpenProject(fullname)
## oDesktop.OpenProject("C:\\00_Asenjo\\00_Project\\Ketupa\\simulation\\Ansys_HFSS\\HFSS_Projects\\Planar_InvertedF_Antenna.aedt")
## oProject = oDesktop.SetActiveProject("Planar_InvertedF_Antenna")
oProject = oDesktop.NewProject()
oProject.InsertDesign("HFSS", "InvertedF_Antenna_2.4G", "DrivenTerminal", "")
oDesign = oProject.SetActiveDesign("InvertedF_Antenna_2.4G")

oDesign.ChangeProperty(
	[
		"NAME:AllTabs",
		[
			"NAME:LocalVariableTab",
			[
				"NAME:PropServers", 
				"LocalVariables"
			],
			[
				"NAME:NewProps",
				[
					"NAME:--Antenna Dimensions",
					"PropType:="		, "SeparatorProp",
					"UserDef:="		, True,
					"Value:="		, ""
				],
				[
					"NAME:Length1",
					"PropType:="		, "VariableProp",
					"UserDef:="		, True,
					"Value:="		, "2.49cm"
				],
				[
					"NAME:Length2",
					"PropType:="		, "VariableProp",
					"UserDef:="		, True,
					"Value:="		, "0.8cm"
				],
				[
					"NAME:Antenna_Trace_Width",
					"PropType:="		, "VariableProp",
					"UserDef:="		, True,
					"Value:="		, "0.15cm"
				],
				[
					"NAME:Antenna_Offset",
					"PropType:="		, "VariableProp",
					"UserDef:="		, True,
					"Value:="		, "0.45cm"
				],
				[
					"NAME:--Feed",
					"PropType:="		, "SeparatorProp",
					"UserDef:="		, True,
					"Value:="		, ""
				],
				[
					"NAME:Feed_Offset",
					"PropType:="		, "VariableProp",
					"UserDef:="		, True,
					"Value:="		, "-0.5cm"
				],
				[
					"NAME:Feed_Length",
					"PropType:="		, "VariableProp",
					"UserDef:="		, True,
					"Value:="		, "0.01cm"
				],
				[
					"NAME:Feed_Width",
					"PropType:="		, "VariableProp",
					"UserDef:="		, True,
					"Value:="		, "0.15cm"
				],
				[
					"NAME:--Substrate Dimensions",
					"PropType:="		, "SeparatorProp",
					"UserDef:="		, True,
					"Value:="		, ""
				],
				[
					"NAME:subH",
					"PropType:="		, "VariableProp",
					"UserDef:="		, True,
					"Value:="		, "62mil"
				],
				[
					"NAME:subX",
					"PropType:="		, "VariableProp",
					"UserDef:="		, True,
					"Value:="		, "5cm"
				],
				[
					"NAME:subY",
					"PropType:="		, "VariableProp",
					"UserDef:="		, True,
					"Value:="		, "10cm"
				]
			]
		]
	])
oDefinitionManager = oProject.GetDefinitionManager()
oDefinitionManager.AddMaterial(
	[
		"NAME:my_mat_ADK",
		"CoordinateSystemType:=", "Cartesian",
		"BulkOrSurfaceType:="	, 1,
		[
			"NAME:PhysicsTypes",
			"set:="			, ["Electromagnetic"]
		],
		"permittivity:="	, "2.2",
		"dielectric_loss_tangent:=", "0.0009"
	])
oEditor = oDesign.SetActiveEditor("3D Modeler")
oEditor.CreateBox(
	[
		"NAME:BoxParameters",
		"XPosition:="		, "-subX/2",
		"YPosition:="		, "-subY+Length2+Antenna_Trace_Width",
		"ZPosition:="		, "0mm",
		"XSize:="		, "subX",
		"YSize:="		, "subY",
		"ZSize:="		, "-subH"
	], 
	[
		"NAME:Attributes",
		"Name:="		, "sub",
		"Flags:="		, "",
		"Color:="		, "(132 132 193)",
		"Transparency:="	, 0.8,
		"PartCoordinateSystem:=", "Global",
		"UDMId:="		, "",
		"MaterialValue:="	, "\"my_mat_ADK\"",
		"SurfaceMaterialValue:=", "\"\"",
		"SolveInside:="		, True,
		"ShellElement:="	, False,
		"ShellElementThickness:=", "0mm",
		"ReferenceTemperature:=", "20cel",
		"IsMaterialEditable:="	, True,
		"IsSurfaceMaterialEditable:=", True,
		"UseMaterialAppearance:=", False,
		"IsLightweight:="	, False
	])
oEditor.CreateRectangle(
	[
		"NAME:RectangleParameters",
		"IsCovered:="		, True,
		"XStart:="		, "-subX/2",
		"YStart:="		, "-subY+Length2+Antenna_Trace_Width",
		"ZStart:="		, "-subH",
		"Width:="		, "subX",
		"Height:="		, "subY-(Length2+Antenna_Trace_Width)",
		"WhichAxis:="		, "Z"
	], 
	[
		"NAME:Attributes",
		"Name:="		, "Ground",
		"Flags:="		, "",
		"Color:="		, "(255 128 65)",
		"Transparency:="	, 0.8,
		"PartCoordinateSystem:=", "Global",
		"UDMId:="		, "",
		"MaterialValue:="	, "\"pec\"",
		"SurfaceMaterialValue:=", "\"\"",
		"SolveInside:="		, False,
		"ShellElement:="	, False,
		"ShellElementThickness:=", "0mm",
		"ReferenceTemperature:=", "20cel",
		"IsMaterialEditable:="	, True,
		"IsSurfaceMaterialEditable:=", True,
		"UseMaterialAppearance:=", False,
		"IsLightweight:="	, False
	])
oModule = oDesign.GetModule("BoundarySetup")
oModule.AssignPerfectE(
	[
		"NAME:Ground",
		"Objects:="		, ["Ground"],
		"InfGroundPlane:="	, False
	])
oEditor.CreateRectangle(
	[
		"NAME:RectangleParameters",
		"IsCovered:="		, True,
		"XStart:="		, "-Antenna_Trace_Width/2+Feed_Offset",
		"YStart:="		, "0mm",
		"ZStart:="		, "0mm",
		"Width:="		, "Antenna_Trace_Width",
		"Height:="		, "Length2",
		"WhichAxis:="		, "Z"
	], 
	[
		"NAME:Attributes",
		"Name:="		, "antenna_feed",
		"Flags:="		, "",
		"Color:="		, "(255 128 65)",
		"Transparency:="	, 0.3,
		"PartCoordinateSystem:=", "Global",
		"UDMId:="		, "",
		"MaterialValue:="	, "\"pec\"",
		"SurfaceMaterialValue:=", "\"\"",
		"SolveInside:="		, False,
		"ShellElement:="	, False,
		"ShellElementThickness:=", "0mm",
		"ReferenceTemperature:=", "20cel",
		"IsMaterialEditable:="	, True,
		"IsSurfaceMaterialEditable:=", True,
		"UseMaterialAppearance:=", False,
		"IsLightweight:="	, False
	])
oEditor.CreateRectangle(
	[
		"NAME:RectangleParameters",
		"IsCovered:="		, True,
		"XStart:="		, "-Antenna_Trace_Width/2-Antenna_Offset+Feed_Offset",
		"YStart:="		, "0mm",
		"ZStart:="		, "0mm",
		"Width:="		, "-Antenna_Trace_Width",
		"Height:="		, "Length2",
		"WhichAxis:="		, "Z"
	], 
	[
		"NAME:Attributes",
		"Name:="		, "antenna_short",
		"Flags:="		, "",
		"Color:="		, "(255 128 65)",
		"Transparency:="	, 0.3,
		"PartCoordinateSystem:=", "Global",
		"UDMId:="		, "",
		"MaterialValue:="	, "\"pec\"",
		"SurfaceMaterialValue:=", "\"\"",
		"SolveInside:="		, False,
		"ShellElement:="	, False,
		"ShellElementThickness:=", "0mm",
		"ReferenceTemperature:=", "20cel",
		"IsMaterialEditable:="	, True,
		"IsSurfaceMaterialEditable:=", True,
		"UseMaterialAppearance:=", False,
		"IsLightweight:="	, False
	])
oEditor.CreateRectangle(
	[
		"NAME:RectangleParameters",
		"IsCovered:="		, True,
		"XStart:="		, "-Antenna_Trace_Width/2-Antenna_Offset+Feed_Offset",
		"YStart:="		, "0mm",
		"ZStart:="		, "0mm",
		"Width:="		, "-subH",
		"Height:="		, "-Antenna_Trace_Width",
		"WhichAxis:="		, "Y"
	], 
	[
		"NAME:Attributes",
		"Name:="		, "antenna_short2",
		"Flags:="		, "",
		"Color:="		, "(255 128 65)",
		"Transparency:="	, 0.3,
		"PartCoordinateSystem:=", "Global",
		"UDMId:="		, "",
		"MaterialValue:="	, "\"pec\"",
		"SurfaceMaterialValue:=", "\"\"",
		"SolveInside:="		, False,
		"ShellElement:="	, False,
		"ShellElementThickness:=", "0mm",
		"ReferenceTemperature:=", "20cel",
		"IsMaterialEditable:="	, True,
		"IsSurfaceMaterialEditable:=", True,
		"UseMaterialAppearance:=", False,
		"IsLightweight:="	, False
	])
oEditor.CreateRectangle(
	[
		"NAME:RectangleParameters",
		"IsCovered:="		, True,
		"XStart:="		, "-Antenna_Trace_Width/2-Antenna_Offset-Antenna_Trace_Width+Feed_Offset",
		"YStart:="		, "Length2",
		"ZStart:="		, "0mm",
		"Width:="		, "Length1",
		"Height:="		, "-Antenna_Trace_Width",
		"WhichAxis:="		, "Z"
	], 
	[
		"NAME:Attributes",
		"Name:="		, "antenna",
		"Flags:="		, "",
		"Color:="		, "(255 128 65)",
		"Transparency:="	, 0.3,
		"PartCoordinateSystem:=", "Global",
		"UDMId:="		, "",
		"MaterialValue:="	, "\"pec\"",
		"SurfaceMaterialValue:=", "\"\"",
		"SolveInside:="		, False,
		"ShellElement:="	, False,
		"ShellElementThickness:=", "0mm",
		"ReferenceTemperature:=", "20cel",
		"IsMaterialEditable:="	, True,
		"IsSurfaceMaterialEditable:=", True,
		"UseMaterialAppearance:=", False,
		"IsLightweight:="	, False
	])
oEditor.Unite(
	[
		"NAME:Selections",
		"Selections:="		, "antenna,antenna_short,antenna_feed,antenna_short2"
	], 
	[
		"NAME:UniteParameters",
		"KeepOriginals:="	, False,
		"TurnOnNBodyBoolean:="	, False
	])
oModule.AssignPerfectE(
	[
		"NAME:antenna",
		"Objects:="		, ["antenna"],
		"InfGroundPlane:="	, False
	])
oEditor.CreateRectangle(
	[
		"NAME:RectangleParameters",
		"IsCovered:="		, True,
		"XStart:="		, "-Feed_Width/2+Feed_Offset",
		"YStart:="		, "0mm",
		"ZStart:="		, "0mm",
		"Width:="		, "Feed_Width",
		"Height:="		, "-Feed_Length",
		"WhichAxis:="		, "Z"
	], 
	[
		"NAME:Attributes",
		"Name:="		, "microstrip_feed",
		"Flags:="		, "",
		"Color:="		, "(255 128 65)",
		"Transparency:="	, 0.3,
		"PartCoordinateSystem:=", "Global",
		"UDMId:="		, "",
		"MaterialValue:="	, "\"pec\"",
		"SurfaceMaterialValue:=", "\"\"",
		"SolveInside:="		, False,
		"ShellElement:="	, False,
		"ShellElementThickness:=", "0mm",
		"ReferenceTemperature:=", "20cel",
		"IsMaterialEditable:="	, True,
		"IsSurfaceMaterialEditable:=", True,
		"UseMaterialAppearance:=", False,
		"IsLightweight:="	, False
	])
oModule.AssignPerfectE(
	[
		"NAME:microstrip_feed",
		"Objects:="		, ["microstrip_feed"],
		"InfGroundPlane:="	, False
	])
oEditor.CreateRectangle(
	[
		"NAME:RectangleParameters",
		"IsCovered:="		, True,
		"XStart:="		, "-Antenna_Trace_Width/2+Feed_Offset",
		"YStart:="		, "-Feed_Length",
		"ZStart:="		, "0mm",
		"Width:="		, "-subH",
		"Height:="		, "Feed_Width",
		"WhichAxis:="		, "Y"
	], 
	[
		"NAME:Attributes",
		"Name:="		, "port1",
		"Flags:="		, "",
		"Color:="		, "(132 132 193)",
		"Transparency:="	, 0.8,
		"PartCoordinateSystem:=", "Global",
		"UDMId:="		, "",
		"MaterialValue:="	, "\"pec\"",
		"SurfaceMaterialValue:=", "\"\"",
		"SolveInside:="		, False,
		"ShellElement:="	, False,
		"ShellElementThickness:=", "0mm",
		"ReferenceTemperature:=", "20cel",
		"IsMaterialEditable:="	, True,
		"IsSurfaceMaterialEditable:=", True,
		"UseMaterialAppearance:=", False,
		"IsLightweight:="	, False
	])
oDesign.SetSolutionType("HFSS Hybrid Terminal Network", 
	[
		"NAME:Options",
		"EnableAutoOpen:="	, False
	])
oModule.AutoIdentifyPorts(
	[
		"NAME:Faces", 
		122
	], False, 
	[
		"NAME:ReferenceConductors", 
		"Ground"
	], "p1", True)
############################Analysis#######################################
oModule = oDesign.GetModule("AnalysisSetup")
oModule.InsertSetup("HfssDriven",
	[
		"NAME:Setup1",
		"SolveType:="		, "Single",
		"Frequency:="		, "2.4GHz",
		"MaxDeltaS:="		, 0.02,
		"UseMatrixConv:="	, False,
		"MaximumPasses:="	, 6,
		"MinimumPasses:="	, 1,
		"MinimumConvergedPasses:=", 1,
		"PercentRefinement:="	, 30,
		"IsEnabled:="		, True,
		[
			"NAME:MeshLink",
			"ImportMesh:="		, False
		],
		"BasisOrder:="		, 1,
		"DoLambdaRefine:="	, True,
		"DoMaterialLambda:="	, True,
		"SetLambdaTarget:="	, False,
		"Target:="		, 0.3333,
		"UseMaxTetIncrease:="	, False,
		"PortAccuracy:="	, 2,
		"UseABCOnPort:="	, False,
		"SetPortMinMaxTri:="	, False,
		"DrivenSolverType:="	, "Direct Solver",
		"EnhancedLowFreqAccuracy:=", False,
		"SaveRadFieldsOnly:="	, False,
		"SaveAnyFields:="	, True,
		"IESolverType:="	, "Auto",
		"LambdaTargetForIESolver:=", 0.15,
		"UseDefaultLambdaTgtForIESolver:=", True,
		"IE Solver Accuracy:="	, "Balanced",
		"InfiniteSphereSetup:="	, "",
		"MaxPass:="		, 10,
		"MinPass:="		, 1,
		"MinConvPass:="		, 1,
		"PerError:="		, 1,
		"PerRefine:="		, 30
	])
oModule.InsertFrequencySweep("Setup1",
	[
		"NAME:Sweep",
		"IsEnabled:="		, True,
		"RangeType:="		, "LinearCount",
		"RangeStart:="		, "0Hz",
		"RangeEnd:="		, "1Hz",
		"RangeCount:="		, 1,
		[
			"NAME:SweepRanges",
			[
				"NAME:Subrange",
				"RangeType:="		, "LogScale",
				"RangeStart:="		, "1Hz",
				"RangeEnd:="		, "100MHz",
				"RangeCount:="		, 401,
				"RangeSamples:="	, 20
			],
			[
				"NAME:Subrange",
				"RangeType:="		, "LinearStep",
				"RangeStart:="		, "100MHz",
				"RangeEnd:="		, "1GHz",
				"RangeStep:="		, "10MHz"
			],
			[
				"NAME:Subrange",
				"RangeType:="		, "LinearStep",
				"RangeStart:="		, "1GHz",
				"RangeEnd:="		, "5GHz",
				"RangeStep:="		, "10MHz"
			]
		],
		"Type:="		, "Interpolating",
		"SaveFields:="		, True,
		"SaveRadFields:="	, True,
		"InterpTolerance:="	, 0.5,
		"InterpMaxSolns:="	, 250,
		"InterpMinSolns:="	, 0,
		"InterpMinSubranges:="	, 1,
		"MinSolvedFreq:="	, "0.01GHz",
		"InterpUseS:="		, True,
		"InterpUsePortImped:="	, True,
		"InterpUsePropConst:="	, True,
		"UseDerivativeConvergence:=", False,
		"InterpDerivTolerance:=", 0.2,
		"UseFullBasis:="	, True,
		"EnforcePassivity:="	, True,
		"PassivityErrorTolerance:=", 0.0001,
		"EnforceCausality:="	, False
	])
##############################################################################################
oDesign.ChangeProperty(
	[
		"NAME:AllTabs",
		[
			"NAME:LocalVariableTab",
			[
				"NAME:PropServers", 
				"LocalVariables"
			],
			[
				"NAME:NewProps",
				[
					"NAME:--Air Box",
					"PropType:="		, "SeparatorProp",
					"UserDef:="		, True,
					"Value:="		, ""
				],
				[
					"NAME:Airbox_dist",
					"PropType:="		, "VariableProp",
					"UserDef:="		, True,
					"Value:="		, "4.1638cm"
				],
				[
					"NAME:--Virtual Object Radiation Surface",
					"PropType:="		, "SeparatorProp",
					"UserDef:="		, True,
					"Value:="		, ""
				],
				[
					"NAME:VirtualObject_dist",
					"PropType:="		, "VariableProp",
					"UserDef:="		, True,
					"Value:="		, "1.2491cm"
				]
			]
		]
	])
oEditor.CreateBox(
	[
		"NAME:BoxParameters",
		"XPosition:="		, "(-subX/2) - Airbox_dist",
		"YPosition:="		, "(-(subY-(Length2+Antenna_Trace_Width))) - Airbox_dist",
		"ZPosition:="		, "(-subH) - Airbox_dist",
		"XSize:="		, "abs((-subX/2)-(subX/2)) + 2*Airbox_dist",
		"YSize:="		, "abs((-(subY-(Length2+Antenna_Trace_Width)))-(Length2+Antenna_Trace_Width)) + 2*Airbox_dist",
		"ZSize:="		, "abs((-subH)-0) + 2*Airbox_dist"
	], 
	[
		"NAME:Attributes",
		"Name:="		, "AirBox",
		"Flags:="		, "",
		"Color:="		, "(128 128 255)",
		"Transparency:="	, 0.9,
		"PartCoordinateSystem:=", "Global",
		"UDMId:="		, "",
		"MaterialValue:="	, "\"air\"",
		"SurfaceMaterialValue:=", "\"\"",
		"SolveInside:="		, True,
		"ShellElement:="	, False,
		"ShellElementThickness:=", "0mm",
		"ReferenceTemperature:=", "20cel",
		"IsMaterialEditable:="	, True,
		"IsSurfaceMaterialEditable:=", True,
		"UseMaterialAppearance:=", False,
		"IsLightweight:="	, False
	])
oEditor = oDesign.SetActiveEditor("3D Modeler")
oEditor.ChangeProperty(
	[
		"NAME:AllTabs",
		[
			"NAME:Geometry3DAttributeTab",
			[
				"NAME:PropServers", 
				"AirBox"
			],
			[
				"NAME:ChangedProps",
				[
					"NAME:Display Wireframe",
					"Value:="		, True
				]
			]
		]
	])
oEditor = oDesign.SetActiveEditor("3D Modeler")
oEditor.CreateBox(
	[
		"NAME:BoxParameters",
		"XPosition:="		, "(-subX/2) - VirtualObject_dist",
		"YPosition:="		, "(-(subY-(Length2+Antenna_Trace_Width))) - VirtualObject_dist",
		"ZPosition:="		, "(-subH) - VirtualObject_dist",
		"XSize:="		, "abs((-subX/2)-(subX/2)) + 2*VirtualObject_dist",
		"YSize:="		, "abs((-(subY-(Length2+Antenna_Trace_Width)))-(Length2+Antenna_Trace_Width)) + 2*VirtualObject_dist",
		"ZSize:="		, "abs((-subH)-0) + 2*VirtualObject_dist"
	], 
	[
		"NAME:Attributes",
		"Name:="		, "VirtualRadiation",
		"Flags:="		, "",
		"Color:="		, "(128 128 255)",
		"Transparency:="	, 0.9,
		"PartCoordinateSystem:=", "Global",
		"UDMId:="		, "",
		"MaterialValue:="	, "\"air\"",
		"SurfaceMaterialValue:=", "\"\"",
		"SolveInside:="		, True,
		"ShellElement:="	, False,
		"ShellElementThickness:=", "0mm",
		"ReferenceTemperature:=", "20cel",
		"IsMaterialEditable:="	, True,
		"IsSurfaceMaterialEditable:=", True,
		"UseMaterialAppearance:=", False,
		"IsLightweight:="	, False
	])
oEditor = oDesign.SetActiveEditor("3D Modeler")
oEditor.ChangeProperty(
	[
		"NAME:AllTabs",
		[
			"NAME:Geometry3DAttributeTab",
			[
				"NAME:PropServers", 
				"VirtualRadiation"
			],
			[
				"NAME:ChangedProps",
				[
					"NAME:Display Wireframe",
					"Value:="		, True
				]
			]
		]
	])
oModule = oDesign.GetModule("BoundarySetup")
oModule.AssignRadiation(
	[
		"NAME:Rad",
		"Objects:="		, ["AirBox"]
	])
oEditor = oDesign.SetActiveEditor("3D Modeler")
oEditor.CreateEntityList(
	[
		"NAME:GeometryEntityListParameters",
		"EntityType:="		, "Face",
		"EntityList:="		, [153,154,155,156,157,158]
	], 
	[
		"NAME:Attributes",
		"Name:="		, "radFaces"
	])
oModule = oDesign.GetModule("RadField")
oModule.InsertInfiniteSphereSetup(
	[
		"NAME:infSphere",
		"UseCustomRadiationSurface:=", True,
		"CustomRadiationSurface:=", "radFaces",
		"CSDefinition:="	, "Theta-Phi",
		"Polarization:="	, "Linear",
		"ThetaStart:="		, "-180deg",
		"ThetaStop:="		, "180deg",
		"ThetaStep:="		, "2deg",
		"PhiStart:="		, "0deg",
		"PhiStop:="		, "180deg",
		"PhiStep:="		, "5deg",
		"UseLocalCS:="		, False
	])
oEditor.SetModelUnits(
	[
		"NAME:Units Parameter",
		"Units:="		, "cm",
		"Rescale:="		, False,
		"Max Model Extent:="	, 10000
	])
##################################Save and Analyze##############################
oProject.SaveAs("C:\\00_Asenjo\\00_Project\\Ketupa\\simulation\\Ansys_HFSS\\HFSS_Projects\\Planar_InvertedF_Antenna.aedt", True)
oDesign.AnalyzeAll()
##################################Create Reports################################
oModule = oDesign.GetModule("ReportSetup")
oModule.CreateReport("Terminal S Parameter Plot 1", "Terminal Solution Data", "Rectangular Plot", "Setup1 : Sweep",
	[
		"Domain:="		, "Sweep"
	],
	[
		"Freq:="		, ["All"],
		"Length1:="		, ["Nominal"],
		"Length2:="		, ["Nominal"],
		"Antenna_Trace_Width:="	, ["Nominal"],
		"Antenna_Offset:="	, ["Nominal"],
		"Feed_Offset:="		, ["Nominal"],
		"Feed_Length:="		, ["Nominal"],
		"Feed_Width:="		, ["Nominal"],
		"subH:="		, ["Nominal"],
		"subX:="		, ["Nominal"],
		"subY:="		, ["Nominal"],
		"Airbox_dist:="		, ["Nominal"],
		"VirtualObject_dist:="	, ["Nominal"]
	],
	[
		"X Component:="		, "Freq",
		"Y Component:="		, ["dB(St(microstrip_feed_T1,microstrip_feed_T1))"]
	])
oModule.CreateReport("Terminal S Parameter Plot 2", "Terminal Solution Data", "Rectangular Plot", "Setup1 : Sweep",
	[
		"Domain:="		, "Sweep"
	],
	[
		"Freq:="		, ["All"],
		"Length1:="		, ["Nominal"],
		"Length2:="		, ["Nominal"],
		"Antenna_Trace_Width:="	, ["Nominal"],
		"Antenna_Offset:="	, ["Nominal"],
		"Feed_Offset:="		, ["Nominal"],
		"Feed_Length:="		, ["Nominal"],
		"Feed_Width:="		, ["Nominal"],
		"subH:="		, ["Nominal"],
		"subX:="		, ["Nominal"],
		"subY:="		, ["Nominal"],
		"Airbox_dist:="		, ["Nominal"],
		"VirtualObject_dist:="	, ["Nominal"]
	],
	[
		"X Component:="		, "Freq",
		"Y Component:="		, ["dB(St(microstrip_feed_T1,microstrip_feed_T1))"]
	])
oModule.CreateReport("ff_3D_GainTotal", "Far Fields", "3D Polar Plot", "Setup1 : LastAdaptive",
	[
		"Context:="		, "infSphere"
	], 
	[
		"Phi:="			, ["All"],
		"Theta:="		, ["All"]
	], 
	[
		"Phi Component:="	, "Phi",
		"Theta Component:="	, "Theta",
		"Mag Component:="	, ["dB(GainTotal)"]
	])
oModule.CreateReport("ff_2D_GainTotal", "Far Fields", "Rectangular Plot", "Setup1 : LastAdaptive", 
	[
		"Context:="		, "infSphere"
	], 
	[
		"Theta:="		, ["All"],
		"Phi:="			, ["0deg"]
	], 
	[
		"X Component:="		, "Theta",
		"Y Component:="		, ["dB(GainTotal)"]
	])
oModule.AddTraces("ff_2D_GainTotal", "Setup1 : LastAdaptive", 
	[
		"Context:="		, "infSphere"
	], 
	[
		"Theta:="		, ["All"],
		"Phi:="			, ["90deg"]
	], 
	[
		"X Component:="		, "Theta",
		"Y Component:="		, ["dB(GainTotal)"]
	])
##################################Save Reports###########################################
oModule = oDesign.GetModule("ReportSetup")
oModule.ExportToFile("Terminal S Parameter Plot 2", "C:/00_Asenjo/00_Project/Ketupa/simulation/Ansys_HFSS/sim_results/Planar_InvertedF_Antenna.csv", False)
oModule.ExportImageToFile("Terminal S Parameter Plot 2", "C:/00_Asenjo/00_Project/Ketupa/simulation/Ansys_HFSS/sim_results/Planar_InvertedF_Antenna.bmp", 2030, 1102)
##################################Save s para############################################
oModule = oDesign.GetModule("Solutions")
oModule.ExportNetworkData("D=\'10um\' L=\'100um\' Lg=\'20um\' W=\'6um\'", ["Setup1:Sweep"], 3, "C:/00_Asenjo/00_Project/Ketupa/simulation/Ansys_HFSS/sim_results/Planar_InvertedF_Antenna.s1p",
						  ["All"], True, 50, "S", -1, 0, 15, True, True, False)
#######################################Save #############################################
oProject.SaveAs("C:\\00_Asenjo\\00_Project\\Ketupa\\simulation\\Ansys_HFSS\\HFSS_Projects\\Planar_InvertedF_Antenna.aedt", True)