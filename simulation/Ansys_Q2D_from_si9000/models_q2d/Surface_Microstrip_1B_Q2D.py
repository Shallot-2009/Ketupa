# ----------------------------------------------------------------------------------------------------
#  [Ansys ver]Ansys Electronics Desktop Version 2024.2.0
#  [Date]  Aug 8, 2025
#  [File name] Surface_Microstrip_1B_Q2D
#  [Notes]  The impedance is a single-ended microstrip line of 50 ohm.
#  [Author's Email]  3405802009@qq.com
# -----------------------------------------------------------------------------------------------------
import os
import sys
import ansys.aedt.core
import win32com.client

import pandas as pd
import camelot.io as  camelot

from datetime import datetime
from pathlib import Path
##################################Read input_files from si 9000 ###########################################
print(f"-------------------------------------------------------------------------------------------------")
tables=camelot.read_pdf('../input_files/Surface_Microstrip_1B.pdf' , pages='1', flavor='stream')
tables.export('../input_files/Surface_Microstrip_1B.csv',f='csv')

excel_file_path="../input_files/Surface_Microstrip_1B-page-1-table-1.csv"
sheet_name="Surface_Microstrip_1B"
# Read excel.
df = pd.read_csv(excel_file_path, sep=',' , header=0, encoding='gb18030')
print(df)
#value=df.iloc[:,:3]
#print(value)
print(f"-------------------------------------------------------------------------------------------------")
##############################################Mkdir#######################################################
current_path = os.path.abspath(__file__)
parent_dir_1 = os.path.dirname(current_path)
parent_dir = os.path.dirname(parent_dir_1)
os.environ['PATH_DIR'] = parent_dir
print(f"Environment Variable Path：{parent_dir}")

path_dir = os.getenv('PATH_DIR')
if not path_dir:
    raise ValueError("ValueError: Environment Variable PATH_DIR not set ")

target_path_1 = Path(path_dir) / 'HFSS_Projects'
target_path_1.mkdir(parents=True, exist_ok=True)
print(f"The project directory has been created：{target_path_1}")

target_path_2 = Path(path_dir) / 'sim_results'
target_path_2.mkdir(parents=True, exist_ok=True)
print(f"The sim results directory has been created：{target_path_2}")

target_path_3 = Path(path_dir) / 'logs'
target_path_3.mkdir(parents=True, exist_ok=True)
print(f"The logs directory has been created：{target_path_3}")
print(f"-------------------------------------------------------------------------------------------------")
############################################################################################################
Z0=df.iloc[5,2]       #Impedance
Microstrip_Impedance= f"{float(Z0) :.3f}" + " ohm"
print(f"Polar Si9000 Impedance：{Microstrip_Impedance}")
##############################################Add Time######################################################
now = datetime.now()     # Get the current date and time
time_string = now.strftime("%Y%m%d_%H%M%S")  # Format the date and time as a string
clean_name = "Surface_Microstrip_1B"
project_file_name =  f"{clean_name}_{Microstrip_Impedance}_{time_string}.aedt"     # The project_file_name of the file to save
project_file =  f"{clean_name}_{Microstrip_Impedance}_{time_string}"               # The  project_file to save
print(f"The Q2D project file name：{project_file_name}")
print(f"-------------------------------------------------------------------------------------------------")
###########################################  Q2D  ############################################################
oAnsoftApp = win32com.client.Dispatch('AnsoftHfss.HfssScriptInterface')
oDesktop = oAnsoftApp.GetAppDesktop()
oDesktop.RestoreWindow()
oProject = oDesktop.NewProject()
oProject.InsertDesign("2D Extractor",  "Surface_Microstrip_1B_(1)",  "","")
oDesign = oProject.SetActiveDesign("Surface_Microstrip_1B_(1)")
oProject.SaveAs(target_path_1/project_file_name, True)
#oProject.SaveAs("$PATH_DIR/Q2D_Projects/Surface_Microstrip_1B.aedt", True)
print(f"The Q2D project has been saved as：{target_path_1/project_file_name}")
print(f"-------------------------------------------------------------------------------------------------")
###############################################################################################################
# Define constants.
AEDT_VERSION = "2025.1"
NUM_CORES = 4
# Create a temporary directory where downloaded data or dumped data can be stored.
# If you’d like to retrieve the project data for subsequent use, the temporary folder name is given by .temp_folder.name.

q2d=ansys.aedt.core.Q2d(target_path_1/project_file_name)
# Define variables.

core_h = "H1"
e_factor = "$Er1"
Lower_Trace_Width = "W1"
Upper_Trace_Width = "W2"
cond_h = "T1"

value1=df.iloc[0,2]
value2=df.iloc[1,2]
value3=df.iloc[2,2]
value4=df.iloc[3,2]
value5=df.iloc[4,2]

for var_name, var_value in {
    "H1": f"{value1:.2f}" + "um",
    "$Er1":f"{value2:.2f}",
    "W1": f"{value3:.2f}" + "um",
    "W2": f"{value4:.2f}" + "um",
    "T1": f"{value5:.2f}" + "um",
	"subX": "5" "mm",
    "subY": "5" "mm",
    "subH": f"{float(value1) / 1000:.4f}" + "mm",
}.items():
    q2d[var_name] = var_value

# 创建基元
# 创建基元并定义层高度

layer_1_lh = 0
layer_1_uh = cond_h
layer_2_lh = layer_1_uh + "+" + core_h
layer_2_uh = layer_2_lh + "+" + cond_h

q2d.save_project(target_path_1/project_file_name, overwrite=True)
q2d.release_desktop(close_projects=True)
############################################################################################################
oAnsoftApp = win32com.client.Dispatch('AnsoftHfss.HfssScriptInterface')
oDesktop = oAnsoftApp.GetAppDesktop()
oDesktop.RestoreWindow()
oDesktop.OpenProject(target_path_1/project_file_name)
oProject = oDesktop.SetActiveProject(project_file)
oDesign = oProject.SetActiveDesign("Surface_Microstrip_1B_(1)")
########################Create 2D Modeler ##################################
oEditor = oDesign.SetActiveEditor("3D Modeler")
oEditor.CreatePolyline(
	[
		"NAME:PolylineParameters",
		"IsPolylineCovered:="	, True,
		"IsPolylineClosed:="	, False,
		[
			"NAME:PolylinePoints",
			[
				"NAME:PLPoint",
				"X:="			, "0mm",
				"Y:="			, "0mm",
				"Z:="			, "0mm"
			],
			[
				"NAME:PLPoint",
				"X:="			, "0mm",
				"Y:="			, "-T1",
				"Z:="			, "0mm"
			]
		],
		[
			"NAME:PolylineSegments",
			[
				"NAME:PLSegment",
				"SegmentType:="		, "Line",
				"StartIndex:="		, 0,
				"NoOfPoints:="		, 2
			]
		],
		[
			"NAME:PolylineXSection",
			"XSectionType:="	, "None",
			"XSectionOrient:="	, "Auto",
			"XSectionWidth:="	, "0mm",
			"XSectionTopWidth:="	, "0mm",
			"XSectionHeight:="	, "0mm",
			"XSectionNumSegments:="	, "0",
			"XSectionBendType:="	, "Corner"
		]
	],
	[
		"NAME:Attributes",
		"Name:="		, "Gnd",
		"Flags:="		, "",
		"Color:="		, "(255 128 64)",
		"Transparency:="	, 0,
		"PartCoordinateSystem:=", "Global",
		"UDMId:="		, "",
		"MaterialValue:="	, "\"copper\"",
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
oEditor = oDesign.SetActiveEditor("3D Modeler")
oEditor.ChangeProperty(
	[
		"NAME:AllTabs",
		[
			"NAME:Geometry3DCmdTab",
			[
				"NAME:PropServers",
				"Gnd:CreatePolyline:1"
			],
			[
				"NAME:ChangedProps",
				[
					"NAME:Type",
					"Value:="		, "Line"
				]
			]
		]
	])
oEditor.ChangeProperty(
	[
		"NAME:AllTabs",
		[
			"NAME:Geometry3DCmdTab",
			[
				"NAME:PropServers",
				"Gnd:CreatePolyline:1"
			],
			[
				"NAME:ChangedProps",
				[
					"NAME:Width/Diameter",
					"Value:="		, "subX"
				]
			]
		]
	])

oEditor.CreatePolyline(
	[
		"NAME:PolylineParameters",
		"IsPolylineCovered:="	, True,
		"IsPolylineClosed:="	, False,
		[
			"NAME:PolylinePoints",
			[
				"NAME:PLPoint",
				"X:="			, "0mm",
				"Y:="			, "0mm",
				"Z:="			, "0mm"
			],
			[
				"NAME:PLPoint",
				"X:="			, "0mm",
				"Y:="			, "H1",
				"Z:="			, "0mm"
			]
		],
		[
			"NAME:PolylineSegments",
			[
				"NAME:PLSegment",
				"SegmentType:="		, "Line",
				"StartIndex:="		, 0,
				"NoOfPoints:="		, 2
			]
		],
		[
			"NAME:PolylineXSection",
			"XSectionType:="	, "None",
			"XSectionOrient:="	, "Auto",
			"XSectionWidth:="	, "0mm",
			"XSectionTopWidth:="	, "0mm",
			"XSectionHeight:="	, "0mm",
			"XSectionNumSegments:="	, "0",
			"XSectionBendType:="	, "Corner"
		]
	],
	[
		"NAME:Attributes",
		"Name:="		, "Dielectric",
		"Flags:="		, "",
		"Color:="		, "(143 175 143)",
		"Transparency:="	, 0,
		"PartCoordinateSystem:=", "Global",
		"UDMId:="		, "",
		"MaterialValue:="	, "\"vacuum\"",
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
oEditor = oDesign.SetActiveEditor("3D Modeler")
oEditor.ChangeProperty(
	[
		"NAME:AllTabs",
		[
			"NAME:Geometry3DCmdTab",
			[
				"NAME:PropServers",
				"Dielectric:CreatePolyline:1"
			],
			[
				"NAME:ChangedProps",
				[
					"NAME:Type",
					"Value:="		, "Line"
				]
			]
		]
	])
oEditor.ChangeProperty(
	[
		"NAME:AllTabs",
		[
			"NAME:Geometry3DCmdTab",
			[
				"NAME:PropServers",
				"Dielectric:CreatePolyline:1"
			],
			[
				"NAME:ChangedProps",
				[
					"NAME:Width/Diameter",
					"Value:="		, "subX"
				]
			]
		]
	])
oDefinitionManager = oProject.GetDefinitionManager()
oDefinitionManager.AddMaterial(
	[
		"NAME:Substrate Dielectric 1",
		"CoordinateSystemType:=", "Cartesian",
		"BulkOrSurfaceType:="	, 1,
		[
			"NAME:PhysicsTypes",
			"set:="			, ["Electromagnetic"]
		],
		"permittivity:="	, "$Er1",
		"permeability:="	, "1.00001",
		"dielectric_loss_tangent:=", "0.085"
	])
oDesign = oProject.SetActiveDesign("Surface_Microstrip_1B_(1)")
oEditor = oDesign.SetActiveEditor("3D Modeler")
oEditor.AssignMaterial(
	[
		"NAME:Selections",
		"AllowRegionDependentPartSelectionForPMLCreation:=", True,
		"AllowRegionSelectionForPMLCreation:=", True,
		"Selections:="		, "Dielectric"
	],
	[
		"NAME:Attributes",
		"MaterialValue:="	, "\"Substrate Dielectric 1\"",
		"SolveInside:="		, False,
		"ShellElement:="	, False,
		"ShellElementThickness:=", "nan ",
		"ReferenceTemperature:=", "nan ",
		"IsMaterialEditable:="	, True,
		"IsSurfaceMaterialEditable:=", True,
		"UseMaterialAppearance:=", False,
		"IsLightweight:="	, False
	])


oEditor.CreatePolyline(
	[
		"NAME:PolylineParameters",
		"IsPolylineCovered:="	, True,
		"IsPolylineClosed:="	, False,
		[
			"NAME:PolylinePoints",
			[
				"NAME:PLPoint",
				"X:="			, "0mm",
				"Y:="			, "-T1-(subY/100)",    ###############For importing subY to HFSS ###############
				"Z:="			, "0mm"
			],
			[
				"NAME:PLPoint",
				"X:="			, "0mm",
				"Y:="			, "5*(H1+T1)",
				"Z:="			, "0mm"
			]
		],
		[
			"NAME:PolylineSegments",
			[
				"NAME:PLSegment",
				"SegmentType:="		, "Line",
				"StartIndex:="		, 0,
				"NoOfPoints:="		, 2
			]
		],
		[
			"NAME:PolylineXSection",
			"XSectionType:="	, "None",
			"XSectionOrient:="	, "Auto",
			"XSectionWidth:="	, "0mm",
			"XSectionTopWidth:="	, "0mm",
			"XSectionHeight:="	, "0mm",
			"XSectionNumSegments:="	, "0",
			"XSectionBendType:="	, "Corner"
		]
	],
	[
		"NAME:Attributes",
		"Name:="		, "AirBox",
		"Flags:="		, "",
		"Color:="		, "(143 175 143)",
		"Transparency:="	, 1,
		"PartCoordinateSystem:=", "Global",
		"UDMId:="		, "",
		"MaterialValue:="	, "\"air\"",
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
oEditor = oDesign.SetActiveEditor("3D Modeler")
oEditor.ChangeProperty(
	[
		"NAME:AllTabs",
		[
			"NAME:Geometry3DCmdTab",
			[
				"NAME:PropServers",
				"AirBox:CreatePolyline:1"
			],
			[
				"NAME:ChangedProps",
				[
					"NAME:Type",
					"Value:="		, "Line"
				]
			]
		]
	])
oEditor.ChangeProperty(
	[
		"NAME:AllTabs",
		[
			"NAME:Geometry3DCmdTab",
			[
				"NAME:PropServers",
				"AirBox:CreatePolyline:1"
			],
			[
				"NAME:ChangedProps",
				[
					"NAME:Width/Diameter",
					"Value:="		, "subX"
				]
			]
		]
	])

oEditor.CreatePolyline(
	[
		"NAME:PolylineParameters",
		"IsPolylineCovered:="	, True,
		"IsPolylineClosed:="	, True,
		[
			"NAME:PolylinePoints",
			[
				"NAME:PLPoint",
				"X:="			, "-(W1)/2",
				"Y:="			, "H1",
				"Z:="			, "0mm"
			],
			[
				"NAME:PLPoint",
				"X:="			, "(W1)/2",
				"Y:="			, "H1",
				"Z:="			, "0mm"
			],
			[
				"NAME:PLPoint",
				"X:="			, "(W2)/2",
				"Y:="			, "H1+T1",
				"Z:="			, "0mm"
			],
			[
				"NAME:PLPoint",
				"X:="			, "-(W2)/2",
				"Y:="			, "H1+T1",
				"Z:="			, "0mm"
			],
			[
				"NAME:PLPoint",
				"X:="			, "-(W1)/2",
				"Y:="			, "H1",
				"Z:="			, "0mm"
			]
		],
		[
			"NAME:PolylineSegments",
			[
				"NAME:PLSegment",
				"SegmentType:="		, "Line",
				"StartIndex:="		, 0,
				"NoOfPoints:="		, 2
			],
			[
				"NAME:PLSegment",
				"SegmentType:="		, "Line",
				"StartIndex:="		, 1,
				"NoOfPoints:="		, 2
			],
			[
				"NAME:PLSegment",
				"SegmentType:="		, "Line",
				"StartIndex:="		, 2,
				"NoOfPoints:="		, 2
			],
			[
				"NAME:PLSegment",
				"SegmentType:="		, "Line",
				"StartIndex:="		, 3,
				"NoOfPoints:="		, 2
			]
		],
		[
			"NAME:PolylineXSection",
			"XSectionType:="	, "None",
			"XSectionOrient:="	, "Auto",
			"XSectionWidth:="	, "0mm",
			"XSectionTopWidth:="	, "0mm",
			"XSectionHeight:="	, "0mm",
			"XSectionNumSegments:="	, "0",
			"XSectionBendType:="	, "Corner"
		]
	],
	[
		"NAME:Attributes",
		"Name:="		, "Microstrip_1B",
		"Flags:="		, "",
		"Color:="		, "(255 128 64)",
		"Transparency:="	, 0,
		"PartCoordinateSystem:=", "Global",
		"UDMId:="		, "",
		"MaterialValue:="	, "\"copper\"",
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

oProject.Save()

oEditor = oDesign.SetActiveEditor("3D Modeler")
oEditor.Copy(
	[
		"NAME:Selections",
		"Selections:="		, "Gnd,Dielectric,AirBox,Microstrip_1B"
	])
###########################################  HFSS  ###########################################################
oProject.InsertDesign("HFSS", "Surface_Microstrip_1B_(2)", "DrivenTerminal", "")
oDesign = oProject.SetActiveDesign("Surface_Microstrip_1B_(2)")
oProject.SaveAs(target_path_1/project_file_name, True)
print(f"The HFSS project has been saved as：{target_path_1/project_file_name}")
print(f"-------------------------------------------------------------------------------------------------")
###############################################################################################################
oDesign = oProject.SetActiveDesign("Surface_Microstrip_1B_(2)")
oEditor = oDesign.SetActiveEditor("3D Modeler")
oEditor.Paste()

oEditor.Rotate(
	[
		"NAME:Selections",
		"Selections:="		, "Gnd,Dielectric,AirBox,Microstrip_1B",
		"NewPartsModelFlag:="	, "Model"
	],
	[
		"NAME:RotateParameters",
		"RotateAxis:="		, "X",
		"RotateAngle:="		, "90deg"
	])
oEditor.ThickenSheet(
	[
		"NAME:Selections",
		"Selections:="		, "Gnd,Dielectric,AirBox",
		"NewPartsModelFlag:="	, "Model"
	],
	[
		"NAME:SheetThickenParameters",
		"Thickness:="		, "subY",
		"BothSides:="		, False,
		[
			"NAME:ThickenAdditionalInfo",
			[
				"NAME:ShellThickenDirectionInfo",
				"SampleFaceID:="	, 7,
				"ComponentSense:="	, True,
				[
					"NAME:PointOnSampleFace",
					"X:="			, "-0.367865mm",
					"Y:="			, "-1.04444492124002e-18mm",
					"Z:="			, "-0.01705708mm"
				],
				[
					"NAME:DirectionAtPoint",
					"X:="			, "0mm",
					"Y:="			, "1mm",
					"Z:="			, "-6.12323399573677e-17mm"
				]
			]
		]
	])
oEditor.ThickenSheet(
	[
		"NAME:Selections",
		"Selections:="		, "Microstrip_1B",
		"NewPartsModelFlag:="	, "Model"
	],
	[
		"NAME:SheetThickenParameters",
		"Thickness:="		, "-subY",
		"BothSides:="		, False,
		[
			"NAME:ThickenAdditionalInfo",
			[
				"NAME:ShellThickenDirectionInfo",
				"SampleFaceID:="	, 7,
				"ComponentSense:="	, True,
				[
					"NAME:PointOnSampleFace",
					"X:="			, "-0.367865mm",
					"Y:="			, "-1.04444492124002e-18mm",
					"Z:="			, "-0.01705708mm"
				],
				[
					"NAME:DirectionAtPoint",
					"X:="			, "0mm",
					"Y:="			, "1mm",
					"Z:="			, "-6.12323399573677e-17mm"
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
		"XStart:="		, "-(subX)/2",
		"YStart:="		, "0",
		"ZStart:="		, "-T1-(subY/100)",
		"Width:="		, "5*(H1+T1)+(subY/100)",
		"Height:="		, "subX",
		"WhichAxis:="		, "Y"
	],
	[
		"NAME:Attributes",
		"Name:="		, "port1",
		"Flags:="		, "",
		"Color:="		, "(143 175 143)",
		"Transparency:="	, 0.8,
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
		"XStart:="		, "-(subX)/2",
		"YStart:="		, "subY",
		"ZStart:="		, "-T1-(subY/100)",
		"Width:="		, "5*(H1+T1)+(subY/100)",
		"Height:="		, "subX",
		"WhichAxis:="		, "Y"
	],
	[
		"NAME:Attributes",
		"Name:="		, "port2",
		"Flags:="		, "",
		"Color:="		, "(143 175 143)",
		"Transparency:="	, 0.8,
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
#########################Add Radiation#####################################
oModule = oDesign.GetModule("BoundarySetup")
oModule.AssignRadiation(
	[
		"NAME:Rad1",
		"Objects:="		, ["AirBox"]
	])
oModule.AutoIdentifyPorts(
	[
		"NAME:Faces",
		173
	], True,
	[
		"NAME:ReferenceConductors",
		"Gnd"
	], "1", True)
oModule.AutoIdentifyPorts(
	[
		"NAME:Faces",
		185
	], True,
	[
		"NAME:ReferenceConductors",
		"Gnd"
	], "2", True)
###########################MeshSetup#########################################
oModule = oDesign.GetModule("MeshSetup")
oModule.AssignLengthOp(
	[
		"NAME:Length1",
		"RefineInside:="	, False,
		"Enabled:="		, True,
		"Objects:="		, ["Microstrip_1B"],
		"RestrictElem:="	, False,
		"NumMaxElem:="		, "1000",
		"RestrictLength:="	, True,
		"MaxLength:="		, "0.1mm"
	])
############################Analysis#######################################
oModule = oDesign.GetModule("AnalysisSetup")
oModule.InsertSetup("HfssDriven",
	[
		"NAME:Setup1",
		"SolveType:="		, "Single",
		"Frequency:="		, "4GHz",
		"MaxDeltaS:="		, 0.01,
		"UseMatrixConv:="	, False,
		"MaximumPasses:="	, 30,
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
				"RangeEnd:="		, "16GHz",
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
oProject.SaveAs(target_path_1/project_file_name, True)
oDesign.AnalyzeAll()
##################################Create Reports################################
oModule = oDesign.GetModule("ReportSetup")
oModule.CreateReport("Terminal S Parameter Plot 1", "Terminal Solution Data", "Rectangular Plot", "Setup1 : Sweep",
	[
		"Domain:="		, "Sweep"
	],
	[
		"Freq:="		, ["All"],
		"H1:="			, ["Nominal"],
		"W1:="			, ["Nominal"],
		"W2:="			, ["Nominal"],
		"T1:="			, ["Nominal"],
		"subX:="		, ["Nominal"],
		"subY:="		, ["Nominal"],
		"subH:="		, ["Nominal"],
		"$Er1:="		, ["Nominal"]
	],
	[
		"X Component:="		, "Freq",
		"Y Component:="		, ["dB(St(Microstrip_1B_T1,Microstrip_1B_T1))","dB(St(Microstrip_1B_T2,Microstrip_1B_T1))"]
	])
oModule.CreateReport("Terminal TDR Impedance Plot 2", "Terminal Solution Data", "Rectangular Plot", "Setup1 : Sweep",
	[
		"Domain:="		, "Time",
		"HoldTime:="		, 1,
		"RiseTime:="		, 2E-10,
		"StepTime:="		, 4E-11,
		"Step:="		, True,
		"WindowWidth:="		, 1,
		"WindowType:="		, 4,
		"KaiserParameter:="	, 1,
		"MaximumTime:="		, 4E-09
	],
	[
		"Time:="		, ["All"],
		"H1:="			, ["Nominal"],
		"W1:="			, ["Nominal"],
		"W2:="			, ["Nominal"],
		"T1:="			, ["Nominal"],
		"subX:="		, ["Nominal"],
		"subY:="		, ["Nominal"],
		"subH:="		, ["Nominal"],
		"$Er1:="		, ["Nominal"]
	],
	[
		"X Component:="		, "Time",
		"Y Component:="		, ["TDRZt(Microstrip_1B_T1)"]
	])
##################################Save Reports###########################################
sim_results_file =  f"{clean_name}_{Microstrip_Impedance}_{time_string}"
oModule = oDesign.GetModule("ReportSetup")
oModule.ExportToFile("Terminal S Parameter Plot 1", f"{str(target_path_2)}/{sim_results_file}.csv", False)
oModule.ExportImageToFile("Terminal S Parameter Plot 1", f"{str(target_path_2)}/{sim_results_file}.bmp", 2030, 1102)

oModule.ExportImageToFile("Terminal TDR Impedance Plot 2", f"{str(target_path_2)}/{sim_results_file}_TDR.bmp", 2030, 1102)
##################################Save s para############################################
oModule = oDesign.GetModule("Solutions")
oModule.ExportNetworkData("D=\'10um\' L=\'100um\' Lg=\'20um\' W=\'6um\'", ["Setup1:Sweep"], 3, f"{str(target_path_2)}/{sim_results_file}.s4p",
						  ["All"], True, 50, "S", -1, 0, 15, True, True, False)
#######################################Save #############################################
oProject.SaveAs(target_path_1/project_file_name, True)
print(f"Successful")
##########################################################################################
log_file =  f"{clean_name}_{Microstrip_Impedance}_{time_string}"
sys.stdout=open(f"{str(target_path_3)}/{log_file }.log",'w')
sys.stdout.close()