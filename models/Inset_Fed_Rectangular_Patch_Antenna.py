# --------------------------------------------------------------
#  [Ansys ver]Ansys Electronics Desktop Version 2024.2.0
#  [Date]  Jun 27, 2025
#  [File name]Inset_Fed_Rectangular_Patch_Antenna
#  [Notes]  Inset_Fed_Rectangular_Patch_Antenna
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
## oDesktop.OpenProject("C:\\00_Asenjo\\00_Project\\Ketupa\\simulation\\Ansys_HFSS\\HFSS_Projects\\Inset_Fed_Rectangular_Patch_Antenna.aedt")
## oProject = oDesktop.SetActiveProject("Inset_Fed_Rectangular_Patch_Antenna")
oProject = oDesktop.NewProject()
oProject.InsertDesign("HFSS", "Rectangular_Patch_Antenna", "DrivenTerminal", "")
oDesign = oProject.SetActiveDesign("Rectangular_Patch_Antenna")

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
					"NAME:--Patch Dimensions",
					"PropType:="		, "SeparatorProp",
					"UserDef:="		, True,
					"Value:="		, ""
				],
				[
					"NAME:patchX",
					"PropType:="		, "VariableProp",
					"UserDef:="		, True,
					"Value:="		, "41mm"
				],
				[
					"NAME:patchY",
					"PropType:="		, "VariableProp",
					"UserDef:="		, True,
					"Value:="		, "32.5mm"
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
					"Value:="		, "1.5mm"
				],
				[
					"NAME:subX",
					"PropType:="		, "VariableProp",
					"UserDef:="		, True,
					"Value:="		, "70mm"
				],
				[
					"NAME:subY",
					"PropType:="		, "VariableProp",
					"UserDef:="		, True,
					"Value:="		, "70mm"
				],
				[
					"NAME:--Feed",
					"PropType:="		, "SeparatorProp",
					"UserDef:="		, True,
					"Value:="		, ""
				],
				[
					"NAME:InsetDistance",
					"PropType:="		, "VariableProp",
					"UserDef:="		, True,
					"Value:="		, "9.5mm"
				],
				[
					"NAME:InsetGap",
					"PropType:="		, "VariableProp",
					"UserDef:="		, True,
					"Value:="		, "4mm"
				],
				[
					"NAME:FeedWidth",
					"PropType:="		, "VariableProp",
					"UserDef:="		, True,
					"Value:="		, "3.2mm"
				],
				[
					"NAME:FeedLength",
					"PropType:="		, "VariableProp",
					"UserDef:="		, True,
					"Value:="		, "17.5mm"
				]
			]
		]
	])
oDefinitionManager = oProject.GetDefinitionManager()
oDefinitionManager.AddMaterial(
	[
		"NAME:Rogers_RO4232_ADK",
		"CoordinateSystemType:=", "Cartesian",
		"BulkOrSurfaceType:="	, 1,
		[
			"NAME:PhysicsTypes",
			"set:="			, ["Electromagnetic"]
		],
		"permittivity:="	, "3.2",
		"dielectric_loss_tangent:=", "0.0009"
	])
oEditor = oDesign.SetActiveEditor("3D Modeler")
oEditor.CreateBox(
	[
		"NAME:BoxParameters",
		"XPosition:="		, "-subX/2",
		"YPosition:="		, "-subY/2",
		"ZPosition:="		, "0mm",
		"XSize:="		, "subX",
		"YSize:="		, "subY",
		"ZSize:="		, "subH"
	], 
	[
		"NAME:Attributes",
		"Name:="		, "sub",
		"Flags:="		, "",
		"Color:="		, "(132 132 193)",
		"Transparency:="	, 0.8,
		"PartCoordinateSystem:=", "Global",
		"UDMId:="		, "",
		"MaterialValue:="	, "\"Rogers_RO4232_ADK\"",
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
		"YStart:="		, "-subY/2",
		"ZStart:="		, "0mm",
		"Width:="		, "subX",
		"Height:="		, "subY",
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
		"XStart:="		, "-patchX/2",
		"YStart:="		, "-patchY/2",
		"ZStart:="		, "subH",
		"Width:="		, "patchX",
		"Height:="		, "patchY",
		"WhichAxis:="		, "Z"
	], 
	[
		"NAME:Attributes",
		"Name:="		, "patch",
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
		"XStart:="		, "-FeedWidth/2-InsetGap",
		"YStart:="		, "patchY/2-InsetDistance",
		"ZStart:="		, "subH",
		"Width:="		, "FeedWidth+2*InsetGap",
		"Height:="		, "FeedLength",
		"WhichAxis:="		, "Z"
	], 
	[
		"NAME:Attributes",
		"Name:="		, "cutout",
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
oEditor.Subtract(
	[
		"NAME:Selections",
		"Blank Parts:="		, "patch",
		"Tool Parts:="		, "cutout"
	], 
	[
		"NAME:SubtractParameters",
		"KeepOriginals:="	, False,
		"TurnOnNBodyBoolean:="	, False
	])
oEditor.CreateRectangle(
	[
		"NAME:RectangleParameters",
		"IsCovered:="		, True,
		"XStart:="		, "-FeedWidth/2",
		"YStart:="		, "patchY/2-InsetDistance",
		"ZStart:="		, "subH",
		"Width:="		, "FeedWidth",
		"Height:="		, "FeedLength",
		"WhichAxis:="		, "Z"
	], 
	[
		"NAME:Attributes",
		"Name:="		, "Feed",
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
		"Selections:="		, "patch,Feed"
	], 
	[
		"NAME:UniteParameters",
		"KeepOriginals:="	, False,
		"TurnOnNBodyBoolean:="	, False
	])
oEditor.CreateRectangle(
	[
		"NAME:RectangleParameters",
		"IsCovered:="		, True,
		"XStart:="		, "-FeedWidth/2",
		"YStart:="		, "patchY/2-InsetDistance+FeedLength",
		"ZStart:="		, "0",
		"Width:="		, "subH",
		"Height:="		, "FeedWidth",
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
oModule.AssignPerfectE(
	[
		"NAME:Patch",
		"Objects:="		, ["patch"],
		"InfGroundPlane:="	, False
	])
oDesign.SetSolutionType("HFSS Hybrid Terminal Network", 
	[
		"NAME:Options",
		"EnableAutoOpen:="	, False
	])
oModule.AutoIdentifyPorts(
	[
		"NAME:Faces", 
		105
	], False, 
	[
		"NAME:ReferenceConductors", 
		"Ground"
	], "1", True)
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
					"Value:="		, "9.9931mm"
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
					"Value:="		, "2.9979mm"
				]
			]
		]
	])
oEditor.CreateBox(
	[
		"NAME:BoxParameters",
		"XPosition:="		, "-subX/2 - Airbox_dist",
		"YPosition:="		, "-subY/2 - Airbox_dist",
		"ZPosition:="		, "-subH - Airbox_dist",
		"XSize:="		, "abs(-subX/2-subX/2) + 2*Airbox_dist",
		"YSize:="		, "abs(-subY/2-subY/2) + 2*Airbox_dist",
		"ZSize:="		, "abs(-subH-subH) + 2*Airbox_dist"
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
		"XPosition:="		, "-subX/2 - VirtualObject_dist",
		"YPosition:="		, "-subY/2 - VirtualObject_dist",
		"ZPosition:="		, "-subH - VirtualObject_dist",
		"XSize:="		, "abs(-subX/2-subX/2) + 2*VirtualObject_dist",
		"YSize:="		, "abs(-subY/2-subY/2) + 2*VirtualObject_dist",
		"ZSize:="		, "abs(-subH-subH) + 2*VirtualObject_dist"
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
		"EntityList:="		, [136,137,138,139,140,141]
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
		"Units:="		, "mm",
		"Rescale:="		, False,
		"Max Model Extent:="	, 10000
	])
##################################Save and Analyze##############################
oProject.SaveAs("C:\\00_Asenjo\\00_Project\\Ketupa\\simulation\\Ansys_HFSS\\HFSS_Projects\\Inset_Fed_Rectangular_Patch_Antenna.aedt", True)
oDesign.AnalyzeAll()
##################################Create Reports################################
oModule = oDesign.GetModule("ReportSetup")
oModule.CreateReport("Terminal S Parameter Plot 1", "Terminal Solution Data", "Rectangular Plot", "Setup1 : Sweep",
	[
		"Domain:="		, "Sweep"
	],
	[
		"Freq:="		, ["All"],
		"patchX:="		, ["Nominal"],
		"patchY:="		, ["Nominal"],
		"subH:="		, ["Nominal"],
		"subX:="		, ["Nominal"],
		"subY:="		, ["Nominal"],
		"InsetDistance:="	, ["Nominal"],
		"InsetGap:="		, ["Nominal"],
		"FeedWidth:="		, ["Nominal"],
		"FeedLength:="		, ["Nominal"],
		"Airbox_dist:="		, ["Nominal"],
		"VirtualObject_dist:="	, ["Nominal"]
	],
	[
		"X Component:="		, "Freq",
		"Y Component:="		, ["dB(St(patch_T1,patch_T1))"]
	])
oModule.CreateReport("Terminal S Parameter Plot 2", "Terminal Solution Data", "Rectangular Plot", "Setup1 : Sweep",
	[
		"Domain:="		, "Sweep"
	],
	[
		"Freq:="		, ["All"],
		"patchX:="		, ["Nominal"],
		"patchY:="		, ["Nominal"],
		"subH:="		, ["Nominal"],
		"subX:="		, ["Nominal"],
		"subY:="		, ["Nominal"],
		"InsetDistance:="	, ["Nominal"],
		"InsetGap:="		, ["Nominal"],
		"FeedWidth:="		, ["Nominal"],
		"FeedLength:="		, ["Nominal"],
		"Airbox_dist:="		, ["Nominal"],
		"VirtualObject_dist:="	, ["Nominal"]
	],
	[
		"X Component:="		, "Freq",
		"Y Component:="		, ["dB(St(patch_T1,patch_T1))"]
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
oModule.ExportToFile("Terminal S Parameter Plot 2", "C:/00_Asenjo/00_Project/Ketupa/simulation/Ansys_HFSS/sim_results/Inset_Fed_Rectangular_Patch_Antenna.csv", False)
oModule.ExportImageToFile("Terminal S Parameter Plot 2", "C:/00_Asenjo/00_Project/Ketupa/simulation/Ansys_HFSS/sim_results/Inset_Fed_Rectangular_Patch_Antenna.bmp", 2030, 1102)
##################################Save s para############################################
oModule = oDesign.GetModule("Solutions")
oModule.ExportNetworkData("D=\'10um\' L=\'100um\' Lg=\'20um\' W=\'6um\'", ["Setup1:Sweep"], 3, "C:/00_Asenjo/00_Project/Ketupa/simulation/Ansys_HFSS/sim_results/Inset_Fed_Rectangular_Patch_Antenna.s1p",
						  ["All"], True, 50, "S", -1, 0, 15, True, True, False)
#######################################Save #############################################
oProject.SaveAs("C:\\00_Asenjo\\00_Project\\Ketupa\\simulation\\Ansys_HFSS\\HFSS_Projects\\Inset_Fed_Rectangular_Patch_Antenna.aedt", True)