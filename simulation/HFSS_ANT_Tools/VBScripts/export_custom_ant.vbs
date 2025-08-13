

Dim oHfssApp
Dim oDesktop
Dim oProject
Dim oDesign
Dim oEditor
Dim oModule
Set oHfssApp  = CreateObject("AnsoftHfss.HfssScriptInterface")
Set oDesktop = oHfssApp.GetAppDesktop()
oDesktop.RestoreWindow
Set oProject = oDesktop.GetActiveProject

'
' Setup Arrays and vars used to store variables and model names etc
'

Dim Local_Variables_Array
Dim Project_Variables_Array
Dim Model_List
Dim OutText
Dim Proj
Dim ProjPath
Dim FilePath
Dim LocalTotal
Dim Version





Locallang=getlocale()

Setlocale(1033)









' get arguments passed into script
'on error resume next
dim args
'Set args = AnsoftScript.arguments
'if(IsEmpty(args)) then 
Set args = WSH.arguments
'End if
on error goto 0
'At this point, args has the arguments no matter if you are running 
'under windows script host or Ansoft script hos

projectname_argument = args(0)

      current_time_temp = Replace(time,":"," ")
      current_time_temp = Replace(current_time_temp," ","")
      

Dim sCurPath
sCurPath = CreateObject("Scripting.FileSystemObject").GetAbsolutePathName(".")
Custom_Library_Path = sCurPath & "\custom_library\"

    Const OverwriteExisting = True
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    objFSO.CopyFile projectname_argument , Custom_Library_Path, OverwriteExisting









argument_projectname_pos = InStrRev(projectname_argument,"\")
argument_projectname_len = Len(projectname_argument)
projectname = Right(projectname_argument,argument_projectname_len-argument_projectname_pos)


argument_projectname_pos = InStrRev(projectname,".")
argument_projectname_len = Len(projectname)
projectname = Left(projectname,argument_projectname_pos-1)


Set objFSO = CreateObject("Scripting.FileSystemObject")
objFSO.MoveFile Custom_Library_Path & projectname & ".hfss" , Custom_Library_Path & projectname & current_time_temp & ".hfss"



oDesktop.OpenProject Custom_Library_Path & projectname & current_time_temp & ".hfss"

temp_original_project_name = projectname & current_time_temp




Proj = projectname & current_time_temp
Projpath=Custom_Library_Path & projectname & current_time_temp & ".hfss"

Version=oDesktop.GetVersion()
major_version_pos = InStr(Version,".")
major_version = Left(Version,major_version_pos-1)


Set oProject = oDesktop.SetActiveProject (projectname & current_time_temp)

string_list_of_designs = ""
dim designs 
set designs = oProject.GetDesigns()
for j = 0 to designs.Count-1
designs(j).GetName()
string_list_of_designs = string_list_of_designs & j+1 & ": " & designs(j).GetName() & Chr(13) & Chr(13)
next 

number_of_designs = j
  
if cint(number_of_designs) > cint(1) then
  design_number = InputBox("Please Select the Design to be included as a Custom Antenna" & Chr(13) & Chr(13) & string_list_of_designs, "Choose Design", "1")
  

  
  Do
    i=1
    
    if IsNumeric(design_number) = False then
      design_number = InputBox("Input must be numeric" & Chr(13) & Chr(13) & string_list_of_designs, "Choose Design", "1")
      i=0
    elseif cint(design_number) > cint(number_of_designs)  then
      design_number = InputBox("Design Number Must be less <=" & number_of_designs  & Chr(13) & Chr(13) & string_list_of_designs, "Choose Design", "1")
      i=0   
    elseif cint(design_number) <= cint(0)  then
      design_number = InputBox("Design Number Must Greater Than Zero" & Chr(13) & Chr(13) & string_list_of_designs, "Choose Design", "1")
      i=0  
    end if 
 
  Loop Until i=1
  
  design_number = cint(design_number)

else
  design_number = 1
end if

Set oDesign = oProject.SetActiveDesign (designs(design_number-1).GetName())

design_name = oDesign.GetName()



'Custom_Library_Path = "C:\Documents and Settings\Arien\Desktop\Antenna_Design_Kit\ADK_v1.1_working\custom_library\"
FilePath = Custom_Library_Path & projectname & ".ant"
'msgbox(FilePath)










' Create the file for output
'

Set paramFSO = CreateObject("Scripting.FileSystemObject")
i=0
Do
  If paramFSO.FileExists(FilePath) Then
  
      current_time = Replace(time,":",".")
      current_time = Replace(current_time," ","")
      projectname  =  InputBox("File name already exists, please rename", "Antenna Design Kit Custom Antenna File Rename", projectname  &  "_" & current_time)
      name_param = projectname  & ".ant"
      FilePath = Custom_Library_Path & name_param
      i=0
  
  else
	   Set paramFSO = CreateObject("Scripting.FileSystemObject")
	   Set ofile = paramFSO.CreateTextFile (FilePath)
     i=1
  end if
Loop Until i=1






'
' Check for existence and number of Project Variables
'

	Project_Variables_Array = oProject.GetVariables()
	if IsArray(Project_Variables_Array) = 0 then

		nPROJECT = -1
	else
		nPROJECT=UBound(Project_Variables_Array)
	 
	end if

'
' If project variables exist, read them and write to the output file
'
	ofile.writeline projectname & "," & design_name


	if nPROJECT<>-1 then

		for var = 0 to nPROJECT

			Project_Variable_Value= oProject.GetVariableValue(Project_Variables_Array(Var))
   		ofile.writeline Project_Variables_Array(var)& "," & Project_Variable_Value
			next
  end if




'
' Check for existence of Local Variables
'
		Local_Variables_Array = oDesign.GetVariables()

		if IsArray(Local_Variables_Array) = 0 then

			nLOCAL = -1
		else
			nLocal=UBound(Local_Variables_Array)
	
		end if
'
' If local variables exist, read them and write to the output file
'

		if nLocal <>-1 then

			for var = 0 to nLocal
		
				Local_Variable_Value= oDesign.GetVariableValue(Local_Variables_Array(Var))
   			ofile.writeline Local_Variables_Array(var) & "," & Local_Variable_Value
					Local_Total= Local_Total+1
			next 
						
					
		end if


'
' Close the outputfile.
'
	ofile.close


'


oProject.CopyDesign design_name

Set oProject = oDesktop.NewProject

oProject.SaveAs Custom_Library_Path & "temp_"& projectname  & ".hfss", true

arrayEntities = oProject.Paste
oProject.SaveAs Custom_Library_Path & "temp_"& projectname  & ".hfss", true

oDesktop.CloseProject "temp_"& projectname 


Set oFSO = CreateObject("Scripting.FileSystemObject")
sDestinationFile = Custom_Library_Path & projectname  & ".adk"

original_project_name = projectname 
i=0
Do
  If oFSO.FileExists(sDestinationFile) Then
  

      projectname  =  projectname  &  "_" & current_time
      sDestinationFile = projectname 
      
      i=0
  
  else
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    objFSO.MoveFile Custom_Library_Path & "temp_"& original_project_name & ".hfss" , Custom_Library_Path & projectname  & ".adk"

    Set objFSO = CreateObject("Scripting.FileSystemObject")
    objFSO.DeleteFolder(Custom_Library_Path & "temp_"& original_project_name & ".hfssresults")
    i=1
  end if
Loop Until i=1


if major_version >= 12 then
    Set oProject = oDesktop.GetActiveProject()
    Set oDesign = oProject.GetActiveDesign()
    Set oEditor = oDesign.SetActiveEditor("3D Modeler")
    oEditor.ExportModelImageToFile  _
      Custom_Library_Path & original_project_name & ".jpg", 250, 375, Array("NAME:SaveImageParams", "ShowAxis:=",  _
      "Default", "ShowGrid:=", "Default")
end if


'if discription file with temp name excists, rename to correct name. 
' temp name is generated in ADK.exe when user is prompted for discription

    Set objFSO = CreateObject("Scripting.FileSystemObject")
    objFSO.MoveFile Custom_Library_Path & "temp_ant_d.txt" , Custom_Library_Path & projectname  & ".txt"


oDesktop.CloseProject Proj


    Set objFSO = CreateObject("Scripting.FileSystemObject")
    objFSO.DeleteFolder(Custom_Library_Path & temp_original_project_name & ".hfssresults")
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    objFSO.DeleteFile(Custom_Library_Path & temp_original_project_name  & ".hfss")
    
 if major_version >= 12 then
 success_msg = msgbox("Project successfully imported into Antenna Design Kit." & (Chr(13)) & _
 "Files imported into custom_library folder in ADK installation directory:" & chr(13) & chr(13) & _
 "Project File - " & chr(13) &  projectname  & ".adk" & chr(13) & chr(13) & _
 "Parameter File - " & chr(13) & projectname  & ".ant" & chr(13) &_
 "All variables in design will be included in the parameter file. You can edit this file to change which variables are displayed in ADK." & chr(13) & chr(13) & _
 "Image File - " & chr(13) &  projectname  & ".jpg" & chr(13) & chr(13) & _
 "Antenna Discription File - " & chr(13) &  projectname  & ".txt" ,64,"Import Successful") 
  end if
  
   if major_version < 12 then
   
  success_msg = msgbox("Project successfully imported into Antenna Design Kit." & (Chr(13)) & _
 "Files imported into custom_library folder in ADK installation directory:" & chr(13) & chr(13) & _
 "Project File - " & chr(13) & projectname  & ".adk" & chr(13) & chr(13) & _ 
 "Parameter File - " & chr(13) & projectname  & ".ant" & chr(13) & chr(13) &_
  "All variables in design will be included in the parameter file. You can edit this file " & chr(13) & "to change which variables are displayed in ADK." & chr(13) & chr(13) & _
 "Antenna Discription File - " & chr(13) &  projectname  & ".txt" & chr(13) & chr(13) & _
 "Image File - " & chr(13) & _
 "Image import is not supported prior to HFSS v12. " & chr(13) & "To add an image to Antenna Design Kit, create a .jpg file and " & chr(13) & "place in the custom_library folder located in your ADK" & chr(13) &_
 "installation directory named " & projectname  & ".jpg",64,"Import Successful")   
   
  end if
  
  
    
   dim objShell
    Set objShell = WScript.CreateObject( "WScript.Shell" )
    objShell.Run("HFSS_ADK.exe")
    Set objShell = Nothing

Setlocale(locallang)
