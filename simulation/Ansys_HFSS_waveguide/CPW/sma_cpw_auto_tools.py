# --------------------------------------------------------------
#  [Ansys ver] Ansys Electronics Desktop Version 2024.2.0
#  [Date]  July 27, 2025
#  [File name] SMA_connectorized_CPW
#  [Notes]  SMA_connectorized_CPW
#  [Author's Email]  3405802009@qq.com
# --------------------------------------------------------------
import os
import win32com.client
oAnsoftApp = win32com.client.Dispatch('AnsoftHfss.HfssScriptInterface')
oDesktop = oAnsoftApp.GetAppDesktop()
oDesktop.RestoreWindow()
oProject = oDesktop.NewProject()
oProject.InsertDesign("HFSS", "sma_connectorized_cpw", "DrivenTerminal", "")
oDesign = oProject.SetActiveDesign("sma_connectorized_cpw")

oEditor = oDesign.SetActiveEditor("3D Modeler")
oEditor.Import(
	[
		"NAME:NativeBodyParameters",
		"HealOption:="		, 0,
		"Options:="		, "0",
		"FileType:="		, "UnRecognized",
		"MaxStitchTol:="	, -1,
		"ImportFreeSurfaces:="	, False,
		"GroupByAssembly:="	, False,
		"CreateGroup:="		, True,
		"STLFileUnit:="		, "Auto",
		"MergeFacesAngle:="	, 0.02,
		"HealSTL:="		, False,
		"ReduceSTL:="		, False,
		"ReduceMaxError:="	, 0,
		"ReducePercentage:="	, 100,
		"PointCoincidenceTol:="	, 1E-06,
		"CreateLightweightPart:=", False,
		"ImportMaterialNames:="	, True,
		"SeparateDisjointLumps:=", False,
		"SourceFile:="		, "C:\\00_Asenjo\\00_Project\\Ketupa\\simulation\\Ansys_HFSS_waveguide\\CPW\\dxf\\sma_cpw.sab"
	])

oEditor.AssignMaterial(
	[
		"NAME:Selections",
		"AllowRegionDependentPartSelectionForPMLCreation:=", True,
		"AllowRegionSelectionForPMLCreation:=", True,
		"Selections:="		, "sma_cpw_Unnamed_2"
	],
	[
		"NAME:Attributes",
		"MaterialValue:="	, "\"bronze\"",
		"SolveInside:="		, False,
		"ShellElement:="	, False,
		"ShellElementThickness:=", "nan ",
		"ReferenceTemperature:=", "nan ",
		"IsMaterialEditable:="	, True,
		"IsSurfaceMaterialEditable:=", True,
		"UseMaterialAppearance:=", False,
		"IsLightweight:="	, False
	])
oEditor = oDesign.SetActiveEditor("3D Modeler")
oEditor.AssignMaterial(
	[
		"NAME:Selections",
		"AllowRegionDependentPartSelectionForPMLCreation:=", True,
		"AllowRegionSelectionForPMLCreation:=", True,
		"Selections:="		, "sma_cpw_Unnamed_5"
	],
	[
		"NAME:Attributes",
		"MaterialValue:="	, "\"bronze\"",
		"SolveInside:="		, False,
		"ShellElement:="	, False,
		"ShellElementThickness:=", "nan ",
		"ReferenceTemperature:=", "nan ",
		"IsMaterialEditable:="	, True,
		"IsSurfaceMaterialEditable:=", True,
		"UseMaterialAppearance:=", False,
		"IsLightweight:="	, False
	])
oEditor.AssignMaterial(
	[
		"NAME:Selections",
		"AllowRegionDependentPartSelectionForPMLCreation:=", True,
		"AllowRegionSelectionForPMLCreation:=", True,
		"Selections:="		, "sma_cpw_Unnamed_3"
	],
	[
		"NAME:Attributes",
		"MaterialValue:="	, "\"Teflon (tm)\"",
		"SolveInside:="		, True,
		"ShellElement:="	, False,
		"ShellElementThickness:=", "nan ",
		"ReferenceTemperature:=", "nan ",
		"IsMaterialEditable:="	, True,
		"IsSurfaceMaterialEditable:=", True,
		"UseMaterialAppearance:=", False,
		"IsLightweight:="	, False
	])
oEditor.AssignMaterial(
	[
		"NAME:Selections",
		"AllowRegionDependentPartSelectionForPMLCreation:=", True,
		"AllowRegionSelectionForPMLCreation:=", True,
		"Selections:="		, "sma_cpw_Unnamed_4"
	],
	[
		"NAME:Attributes",
		"MaterialValue:="	, "\"copper\"",
		"SolveInside:="		, False,
		"ShellElement:="	, False,
		"ShellElementThickness:=", "nan ",
		"ReferenceTemperature:=", "nan ",
		"IsMaterialEditable:="	, True,
		"IsSurfaceMaterialEditable:=", True,
		"UseMaterialAppearance:=", False,
		"IsLightweight:="	, False
	])
oEditor.AssignMaterial(
	[
		"NAME:Selections",
		"AllowRegionDependentPartSelectionForPMLCreation:=", True,
		"AllowRegionSelectionForPMLCreation:=", True,
		"Selections:="		, "sma_cpw_Unnamed_9"
	],
	[
		"NAME:Attributes",
		"MaterialValue:="	, "\"copper\"",
		"SolveInside:="		, False,
		"ShellElement:="	, False,
		"ShellElementThickness:=", "nan ",
		"ReferenceTemperature:=", "nan ",
		"IsMaterialEditable:="	, True,
		"IsSurfaceMaterialEditable:=", True,
		"UseMaterialAppearance:=", False,
		"IsLightweight:="	, False
	])
oEditor.AssignMaterial(
	[
		"NAME:Selections",
		"AllowRegionDependentPartSelectionForPMLCreation:=", True,
		"AllowRegionSelectionForPMLCreation:=", True,
		"Selections:="		, "sma_cpw_Unnamed_8"
	],
	[
		"NAME:Attributes",
		"MaterialValue:="	, "\"tin\"",
		"SolveInside:="		, False,
		"ShellElement:="	, False,
		"ShellElementThickness:=", "nan ",
		"ReferenceTemperature:=", "nan ",
		"IsMaterialEditable:="	, True,
		"IsSurfaceMaterialEditable:=", True,
		"UseMaterialAppearance:=", False,
		"IsLightweight:="	, False
	])


oEditor.AssignMaterial(
	[
		"NAME:Selections",
		"AllowRegionDependentPartSelectionForPMLCreation:=", True,
		"AllowRegionSelectionForPMLCreation:=", True,
		"Selections:="		, "sma_cpw_Unnamed_7"
	],
	[
		"NAME:Attributes",
		"MaterialValue:="	, "\"Rogers TMM 3 (tm)\"",
		"SolveInside:="		, True,
		"ShellElement:="	, False,
		"ShellElementThickness:=", "nan ",
		"ReferenceTemperature:=", "nan ",
		"IsMaterialEditable:="	, True,
		"IsSurfaceMaterialEditable:=", True,
		"UseMaterialAppearance:=", False,
		"IsLightweight:="	, False
	])


oEditor.CreateBox(
	[
		"NAME:BoxParameters",
		"XPosition:="		, "0mm",
		"YPosition:="		, "16.4mm",
		"ZPosition:="		, "1.524mm",
		"XSize:="		, "60mm",
		"YSize:="		, "-2.8mm",
		"ZSize:="		, "0.050mm"
	],
	[
		"NAME:Attributes",
		"Name:="		, "micro_line",
		"Flags:="		, "",
		"Color:="		, "(143 175 143)",
		"Transparency:="	, 0,
		"PartCoordinateSystem:=", "Global",
		"UDMId:="		, "",
		"MaterialValue:="	, "\"copper\"",
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
oEditor.CreateBox(
	[
		"NAME:BoxParameters",
		"XPosition:="		, "0mm",
		"YPosition:="		, "17.9mm",
		"ZPosition:="		, "1.524mm",
		"XSize:="		, "60mm",
		"YSize:="		, "12.1mm",
		"ZSize:="		, "0.050mm"
	],
	[
		"NAME:Attributes",
		"Name:="		, "Gnd1",
		"Flags:="		, "",
		"Color:="		, "(143 175 143)",
		"Transparency:="	, 0,
		"PartCoordinateSystem:=", "Global",
		"UDMId:="		, "",
		"MaterialValue:="	, "\"copper\"",
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
oEditor.CreateBox(
	[
		"NAME:BoxParameters",
		"XPosition:="		, "0mm",
		"YPosition:="		, "12mm",
		"ZPosition:="		, "1.524mm",
		"XSize:="		, "60mm",
		"YSize:="		, "-12mm",
		"ZSize:="		, "0.050mm"
	],
	[
		"NAME:Attributes",
		"Name:="		, "Gnd2",
		"Flags:="		, "",
		"Color:="		, "(143 175 143)",
		"Transparency:="	, 0,
		"PartCoordinateSystem:=", "Global",
		"UDMId:="		, "",
		"MaterialValue:="	, "\"copper\"",
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
oEditor.Unite(
	[
		"NAME:Selections",
		"Selections:="		, "micro_line,sma_cpw_Unnamed_8"
	],
	[
		"NAME:UniteParameters",
		"KeepOriginals:="	, False,
		"TurnOnNBodyBoolean:="	, True
	])

oEditor.CreateCircle(
	[
		"NAME:CircleParameters",
		"IsCovered:="		, True,
		"XCenter:="		, "-8mm",
		"YCenter:="		, "15mm",
		"ZCenter:="		, "2.159mm",
		"Radius:="		, "2.6594136571808mm",
		"WhichAxis:="		, "X",
		"NumSegments:="		, "0"
	],
	[
		"NAME:Attributes",
		"Name:="		, "port1",
		"Flags:="		, "",
		"Color:="		, "(143 175 143)",
		"Transparency:="	, 0,
		"PartCoordinateSystem:=", "Global",
		"UDMId:="		, "",
		"MaterialValue:="	, "\"vacuum\"",
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
		"XStart:="		, "60mm",
		"YStart:="		, "0mm",
		"ZStart:="		, "0mm",
		"Width:="		, "30mm",
		"Height:="		, "5mm",
		"WhichAxis:="		, "X"
	],
	[
		"NAME:Attributes",
		"Name:="		, "port2",
		"Flags:="		, "",
		"Color:="		, "(143 175 143)",
		"Transparency:="	, 0,
		"PartCoordinateSystem:=", "Global",
		"UDMId:="		, "",
		"MaterialValue:="	, "\"vacuum\"",
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
oEditor.CreateRegion(
	[
		"NAME:RegionParameters",
		"+XPaddingType:="	, "Percentage Offset",
		"+XPadding:="		, "0",
		"-XPaddingType:="	, "Percentage Offset",
		"-XPadding:="		, "0",
		"+YPaddingType:="	, "Percentage Offset",
		"+YPadding:="		, "0",
		"-YPaddingType:="	, "Percentage Offset",
		"-YPadding:="		, "0",
		"+ZPaddingType:="	, "Percentage Offset",
		"+ZPadding:="		, "0",
		"-ZPaddingType:="	, "Percentage Offset",
		"-ZPadding:="		, "0",
		[
			"NAME:BoxForVirtualObjects",
			[
				"NAME:LowPoint",
				1,
				1,
				1
			],
			[
				"NAME:HighPoint",
				-1,
				-1,
				-1
			]
		]
	],
	[
		"NAME:Attributes",
		"Name:="		, "Air_Region",
		"Flags:="		, "Wireframe#",
		"Color:="		, "(143 175 143)",
		"Transparency:="	, 0,
		"PartCoordinateSystem:=", "Global",
		"UDMId:="		, "",
		"MaterialValue:="	, "\"air\"",
		"SurfaceMaterialValue:=", "\"\"",
		"SolveInside:="		, True,
		"ShellElement:="	, False,
		"ShellElementThickness:=", "nan ",
		"ReferenceTemperature:=", "nan ",
		"IsMaterialEditable:="	, True,
		"IsSurfaceMaterialEditable:=", True,
		"UseMaterialAppearance:=", False,
		"IsLightweight:="	, False
	])
oModule = oDesign.GetModule("BoundarySetup")
oModule.AutoIdentifyPorts(
	[
		"NAME:Faces",
		1242
	], False,
	[
		"NAME:ReferenceConductors",
		"sma_cpw_Unnamed_2"
	], "1", True)
oModule.AutoIdentifyPorts(
	[
		"NAME:Faces",
		1254
	], False,
	[
		"NAME:ReferenceConductors",
		"Gnd1",
		"Gnd2"
	], "2", True)
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
				"RangeEnd:="		, "8GHz",
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
##################################Save and Analyze##############################
oProject.SaveAs("C:\\00_Asenjo\\00_Project\\Ketupa\\simulation\\Ansys_HFSS_waveguide\\CPW\\HFSS_Projects\\sma_connectorized_cpw.aedt", True)
oDesign.AnalyzeAll()
##################################Create Reports################################
oModule = oDesign.GetModule("ReportSetup")
oModule.CreateReport("Terminal S Parameter Plot 1", "Terminal Solution Data", "Rectangular Plot", "Setup1 : Sweep",
	[
		"Domain:="		, "Sweep"
	],
	[
		"Freq:="		, ["All"]
	],
	[
		"X Component:="		, "Freq",
		"Y Component:="		, ["dB(St(sma_cpw_Unnamed_4_T1,sma_cpw_Unnamed_4_T1))","dB(St(micro_line_T1,sma_cpw_Unnamed_4_T1))"]
	])
##################################Save Reports###########################################
oModule = oDesign.GetModule("ReportSetup")
oModule.ExportToFile("Terminal S Parameter Plot 1", "C:\\00_Asenjo\\00_Project\\Ketupa\\simulation\\Ansys_HFSS_waveguide\\CPW\\sim_results\\sma_connectorized_cpw.csv", False)
oModule.ExportImageToFile("Terminal S Parameter Plot 1", "C:\\00_Asenjo\\00_Project\\Ketupa\\simulation\\Ansys_HFSS_waveguide\\CPW\\sim_results\\sma_connectorized_cpw.bmp", 2103, 1042)
##################################Save s para############################################
oModule = oDesign.GetModule("Solutions")
oModule.ExportNetworkData("D=\'10um\' L=\'100um\' Lg=\'20um\' W=\'6um\'", ["Setup1:Sweep"], 3, "C:\\00_Asenjo\\00_Project\\Ketupa\\simulation\\Ansys_HFSS_waveguide\\CPW\\sim_results\\sma_connectorized_cpw.s2p",
						  ["All"], True, 50, "S", -1, 0, 15, True, True, False)
#######################################Save #############################################
oProject.SaveAs("C:\\00_Asenjo\\00_Project\\Ketupa\\simulation\\Ansys_HFSS_waveguide\\CPW\\HFSS_Projects\\sma_connectorized_cpw.aedt", True)