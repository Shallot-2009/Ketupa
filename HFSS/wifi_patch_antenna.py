# --------------------------------------------------------------
#  [Ansys ver]Ansys Electronics Desktop Version 2024.2.0
#  [Date]  Jun 19, 2025
#  [File name]wifi_patch_antenna
#  [Notes]  wifi_patch_antenna for 2.4GHz
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
## oDesktop.OpenProject("C:\\00_Asenjo\\00_Project\\Ketupa\\HFSS\\HFSS_Projects\\Project1.aedt")
## oProject = oDesktop.SetActiveProject("Project1")
oProject = oDesktop.NewProject()
oProject.InsertDesign("HFSS", "wifi_2.4G", "DrivenTerminal", "")
oDesign = oProject.SetActiveDesign("wifi_2.4G")

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
					"NAME:X_0",
					"PropType:="		, "VariableProp",
					"UserDef:="		, True,
					"Value:="		, "15mm"
				],
                [
					"NAME:Y_0",
					"PropType:="		, "VariableProp",
					"UserDef:="		, True,
					"Value:="		, "15mm"
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
					"NAME:subH",
					"PropType:="		, "VariableProp",
					"UserDef:="		, True,
					"Value:="		, "1.5mm"
				],
                [
					"NAME:L_ant",
					"PropType:="		, "VariableProp",
					"UserDef:="		, True,
					"Value:="		, "41mm"
				],
                [
					"NAME:W_ant",
					"PropType:="		, "VariableProp",
					"UserDef:="		, True,
					"Value:="		, "32.5mm"
				],
                [
					"NAME:lamda",
					"PropType:="		, "VariableProp",
					"UserDef:="		, True,
					"Value:="		, "122mm"
				],
                [
					"NAME:T_HOZ_Plating",
					"PropType:="		, "VariableProp",
					"UserDef:="		, True,
					"Value:="		, "40um"
				],
                [
					"NAME:L_open",
					"PropType:="		, "VariableProp",
					"UserDef:="		, True,
					"Value:="		, "4mm"
				],
                [
					"NAME:W_open",
					"PropType:="		, "VariableProp",
					"UserDef:="		, True,
					"Value:="		, "9.5mm"
				],
                [
					"NAME:L_lamda",
					"PropType:="		, "VariableProp",
					"UserDef:="		, True,
					"Value:="		, "9.5mm"
				],
                [
					"NAME:W_lamda",
					"PropType:="		, "VariableProp",
					"UserDef:="		, True,
					"Value:="		, "3.2mm"
				],
                [
					"NAME:L_50ohm",
					"PropType:="		, "VariableProp",
					"UserDef:="		, True,
					"Value:="		, "8mm"
				],
                [
					"NAME:W_50ohm",
					"PropType:="		, "VariableProp",
					"UserDef:="		, True,
					"Value:="		, "3.2mm"
				]
			]
		]
	])
########################Create 3D Modeler ##############################
oEditor = oDesign.SetActiveEditor("3D Modeler")
oEditor.CreateBox(
	[
		"NAME:BoxParameters",
		"XPosition:="		, "-(5mm+lamda/2)",
		"YPosition:="		, "(60mm+lamda/2)",
		"ZPosition:="		, "(10mm+lamda)/2",
		"XSize:="		, "subX+lamda",
		"YSize:="		, "-(subY+lamda)",
		"ZSize:="		, "-(subH+lamda)"
	],
	[
		"NAME:Attributes",
		"Name:="		, "AirBox",
		"Flags:="		, "",
		"Color:="		, "(143 175 143)",
		"Transparency:="	, 0.6,
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
oEditor.CreateBox(
	[
		"NAME:BoxParameters",
		"XPosition:="		, "0",
		"YPosition:="		, "0",
		"ZPosition:="		, "0",
		"XSize:="		, "subX",
		"YSize:="		, "subY",
		"ZSize:="		, "subH"
	],
	[
		"NAME:Attributes",
		"Name:="		, "core",
		"Flags:="		, "",
		"Color:="		, "(143 175 143)",
		"Transparency:="	, 0,
		"PartCoordinateSystem:=", "Global",
		"UDMId:="		, "",
		"MaterialValue:="	, "\"Rogers RO4232 (tm)\"",
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
oEditor.CreateBox(
	[
		"NAME:BoxParameters",
		"XPosition:="		, "0",
		"YPosition:="		, "0",
		"ZPosition:="		, "0",
		"XSize:="		, "subX",
		"YSize:="		, "subY",
		"ZSize:="		, "-T_HOZ_Plating"
	],
	[
		"NAME:Attributes",
		"Name:="		, "Gnd",
		"Flags:="		, "",
		"Color:="		, "(253 187 66)",
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

oEditor = oDesign.SetActiveEditor("3D Modeler")
oEditor.CreateRectangle(
	[
		"NAME:RectangleParameters",
		"IsCovered:="		, True,
		"XStart:="		, "X_0",
		"YStart:="		, "Y_0",
		"ZStart:="		, "subH",
		"Width:="		, "L_ant",
		"Height:="		, "W_ant",
		"WhichAxis:="		, "Z"
	],
	[
		"NAME:Attributes",
		"Name:="		, "Ant",
		"Flags:="		, "",
		"Color:="		, "(253 187 66)",
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
oEditor = oDesign.SetActiveEditor("3D Modeler")
oEditor.CreateRectangle(
	[
		"NAME:RectangleParameters",
		"IsCovered:="		, True,
		"XStart:="		, "X_0+L_ant/2",
		"YStart:="		, "Y_0",
		"ZStart:="		, "subH",
		"Width:="		, "L_open+W_lamda/2",
		"Height:="		, "W_open",
		"WhichAxis:="		, "Z"
	],
	[
		"NAME:Attributes",
		"Name:="		, "Open_1",
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
oEditor.CreateRectangle(
	[
		"NAME:RectangleParameters",
		"IsCovered:="		, True,
		"XStart:="		, "X_0+L_ant/2",
		"YStart:="		, "Y_0",
		"ZStart:="		, "subH",
		"Width:="		, "-(L_open+W_lamda/2)",
		"Height:="		, "W_open",
		"WhichAxis:="		, "Z"
	],
	[
		"NAME:Attributes",
		"Name:="		, "Open_2",
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
oEditor.CreateRectangle(
	[
		"NAME:RectangleParameters",
		"IsCovered:="		, True,
		"XStart:="		, "X_0+L_ant/2",
		"YStart:="		, "Y_0+W_open",
		"ZStart:="		, "subH",
		"Width:="		, "-W_lamda/2",
		"Height:="		, "-L_lamda",
		"WhichAxis:="		, "Z"
	],
	[
		"NAME:Attributes",
		"Name:="		, "Rectangle_1",
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
oEditor.CreateRectangle(
	[
		"NAME:RectangleParameters",
		"IsCovered:="		, True,
		"XStart:="		, "X_0+L_ant/2",
		"YStart:="		, "Y_0+W_open",
		"ZStart:="		, "subH",
		"Width:="		, "W_lamda/2",
		"Height:="		, "-L_lamda",
		"WhichAxis:="		, "Z"
	],
	[
		"NAME:Attributes",
		"Name:="		, "Rectangle_2",
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
oEditor.CreateRectangle(
	[
		"NAME:RectangleParameters",
		"IsCovered:="		, True,
		"XStart:="		, "X_0+L_ant/2",
		"YStart:="		, "Y_0+W_open-L_lamda",
		"ZStart:="		, "subH",
		"Width:="		, "-W_50ohm/2",
		"Height:="		, "-L_50ohm",
		"WhichAxis:="		, "Z"
	],
	[
		"NAME:Attributes",
		"Name:="		, "Rectangle_3_50ohm_1",
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
oEditor.CreateRectangle(
	[
		"NAME:RectangleParameters",
		"IsCovered:="		, True,
		"XStart:="		, "X_0+L_ant/2",
		"YStart:="		, "Y_0+W_open-L_lamda",
		"ZStart:="		, "subH",
		"Width:="		, "W_50ohm/2",
		"Height:="		, "-L_50ohm",
		"WhichAxis:="		, "Z"
	],
	[
		"NAME:Attributes",
		"Name:="		, "Rectangle_4_50ohm_2",
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
oEditor.Subtract(
	[
		"NAME:Selections",
		"Blank Parts:="		, "Ant",
		"Tool Parts:="		, "Open_1,Open_2"
	],
	[
		"NAME:SubtractParameters",
		"KeepOriginals:="	, False,
		"TurnOnNBodyBoolean:="	, True
	])
oEditor.Unite(
	[
		"NAME:Selections",
		"Selections:="		, "Ant,Rectangle_1,Rectangle_2,Rectangle_3_50ohm_1,Rectangle_4_50ohm_2"
	],
	[
		"NAME:UniteParameters",
		"KeepOriginals:="	, False,
		"TurnOnNBodyBoolean:="	, True
	])
oEditor.ThickenSheet(
	[
		"NAME:Selections",
		"Selections:="		, "Ant",
		"NewPartsModelFlag:="	, "Model"
	],
	[
		"NAME:SheetThickenParameters",
		"Thickness:="		, "T_HOZ_plating",
		"BothSides:="		, False,
		[
			"NAME:ThickenAdditionalInfo",
			[
				"NAME:ShellThickenDirectionInfo",
				"SampleFaceID:="	, 100,
				"ComponentSense:="	, True,
				[
					"NAME:PointOnSampleFace",
					"X:="			, "23.5462377mm",
					"Y:="			, "28.8588775mm",
					"Z:="			, "1.5mm"
				],
				[
					"NAME:DirectionAtPoint",
					"X:="			, "-0mm",
					"Y:="			, "0mm",
					"Z:="			, "1mm"
				]
			]
		]
	])
##############################Add port###############################
oEditor = oDesign.SetActiveEditor("3D Modeler")
oEditor.CreateRectangle(
	[
		"NAME:RectangleParameters",
		"IsCovered:="		, True,
		"XStart:="		, "X_0+L_ant/2",
		"YStart:="		, "Y_0+W_open-L_lamda-L_50ohm",
		"ZStart:="		, "2*subH",
		"Width:="		, "-3*subH",
		"Height:="		, "-3*subH",
		"WhichAxis:="		, "Y"
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
		"XStart:="		, "X_0+L_ant/2",
		"YStart:="		, "Y_0+W_open-L_lamda-L_50ohm",
		"ZStart:="		, "2*subH",
		"Width:="		, "-3*subH",
		"Height:="		, "3*subH",
		"WhichAxis:="		, "Y"
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
oEditor.Unite(
	[
		"NAME:Selections",
		"Selections:="		, "port1,port2"
	],
	[
		"NAME:UniteParameters",
		"KeepOriginals:="	, False,
		"TurnOnNBodyBoolean:="	, True
	])
#########################Add Radiation##############3######################
oModule = oDesign.GetModule("BoundarySetup")
oModule.AssignRadiation(
	[
		"NAME:Rad1",
		"Objects:="		, ["AirBOX"]
	])
oModule.AutoIdentifyPorts(
	[
		"NAME:Faces",
		272
	], False,
	[
		"NAME:ReferenceConductors",
		"Gnd"
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
##################################Save and Analyze##############################
oProject.SaveAs("C:\\00_Asenjo\\00_Project\\Ketupa\\HFSS\\HFSS_Projects\\Project1.aedt", True)
oDesign.AnalyzeAll()
##################################Create Reports################################
oModule = oDesign.GetModule("ReportSetup")
oModule.CreateReport("Terminal S Parameter Plot 1", "Terminal Solution Data", "Rectangular Plot", "Setup1 : Sweep",
	[
		"Domain:="		, "Sweep"
	],
	[
		"Freq:="		, ["All"],
		"X_0:="			, ["Nominal"],
		"Y_0:="			, ["Nominal"],
		"subX:="		, ["Nominal"],
		"subY:="		, ["Nominal"],
		"subH:="		, ["Nominal"],
		"L_ant:="		, ["Nominal"],
		"W_ant:="		, ["Nominal"],
		"lamda:="		, ["Nominal"],
		"T_HOZ_Plating:="	, ["Nominal"],
		"L_open:="		, ["Nominal"],
		"W_open:="		, ["Nominal"],
		"L_lamda:="		, ["Nominal"],
		"W_lamda:="		, ["Nominal"],
		"L_50ohm:="		, ["Nominal"],
		"W_50ohm:="		, ["Nominal"]
	],
	[
		"X Component:="		, "Freq",
		"Y Component:="		, ["dB(St(Ant_T1,Ant_T1))"]
	])
oModule.CreateReport("Terminal S Parameter Plot 2", "Terminal Solution Data", "3D Rectangular Plot", "Setup1 : Sweep", [],
	[
		"Freq:="		, ["All"],
		"X_0:="			, ["All"],
		"Y_0:="			, ["Nominal"],
		"subX:="		, ["Nominal"],
		"subY:="		, ["Nominal"],
		"subH:="		, ["Nominal"],
		"L_ant:="		, ["Nominal"],
		"W_ant:="		, ["Nominal"],
		"lamda:="		, ["Nominal"],
		"T_HOZ_Plating:="	, ["Nominal"],
		"L_open:="		, ["Nominal"],
		"W_open:="		, ["Nominal"],
		"L_lamda:="		, ["Nominal"],
		"W_lamda:="		, ["Nominal"],
		"L_50ohm:="		, ["Nominal"],
		"W_50ohm:="		, ["Nominal"]
	],
	[
		"X Component:="		, "Freq",
		"Y Component:="		, "X_0",
		"Z Component:="		, ["dB(St(Ant_T1,Ant_T1))"]
	])
oModule = oDesign.GetModule("FieldsReporter")
oModule.CreateFieldPlot(
	[
		"NAME:Mag_E1",
		"SolutionName:="	, "Setup1 : LastAdaptive",
		"UserSpecifyName:="	, 0,
		"UserSpecifyFolder:="	, 0,
		"QuantityName:="	, "Mag_E",
		"PlotFolder:="		, "E Field",
		"StreamlinePlot:="	, False,
		"AdjacentSidePlot:="	, False,
		"FullModelPlot:="	, False,
		"IntrinsicVar:="	, "Freq=\'2.3999999999999999GHz\' Phase=\'0deg\'",
		"PlotGeomInfo:="	, [1,"Volume","ObjList",1,"Ant"],
		"FilterBoxes:="		, [0],
		[
			"NAME:PlotOnVolumeSettings",
			"PlotIsoSurface:="	, True,
			"PointSize:="		, 1,
			"Refinement:="		, 0,
			"CloudSpacing:="	, 0.5,
			"CloudMinSpacing:="	, -1,
			"CloudMaxSpacing:="	, -1,
			"ShadingType:="		, 0,
			"IsoMapTransparency:="	, True,
			"IsoTransparency:="	, 0.899999976158142,
			"IsoTransScaleThreshold:=", 0.200000002980232,
			[
				"NAME:Arrow3DSpacingSettings",
				"ArrowUniform:="	, True,
				"ArrowSpacing:="	, 0,
				"MinArrowSpacing:="	, 0,
				"MaxArrowSpacing:="	, 0
			]
		],
		"EnableGaussianSmoothing:=", False,
		"SurfaceOnly:="		, False
	], "Field")
##################################Save Reports###########################################
oModule = oDesign.GetModule("ReportSetup")
oModule.ExportToFile("Terminal S Parameter Plot 1", "C:/00_Asenjo/00_Project/Ketupa/HFSS/Sim_results/Terminal S Parameter Plot 1.csv", False)
oModule.ExportImageToFile("Terminal S Parameter Plot 1", "C:/00_Asenjo/00_Project/Ketupa/HFSS/Sim_results/Terminal S Parameter Plot 1.bmp", 2030, 1102)
##################################Save s para############################################
oModule = oDesign.GetModule("Solutions")
oModule.ExportNetworkData("D=\'10um\' L=\'100um\' Lg=\'20um\' W=\'6um\'", ["Setup1:Sweep"], 3, "C:/00_Asenjo/00_Project/Ketupa/HFSS/Sim_results/tutorial_Design1.s1p",
						  ["All"], True, 50, "S", -1, 0, 15, True, True, False)
#######################################Save #############################################
oProject.SaveAs("C:\\00_Asenjo\\00_Project\\Ketupa\\HFSS\\HFSS_Projects\\Project1.aedt", True)
