Dim oAnsoftApp
Dim oDesktop
Dim oProject
Dim oDesign
Dim oEditor
Dim oModule
Set oAnsoftApp = CreateObject("AnsoftHfss.HfssScriptInterface")
Set oDesktop = oAnsoftApp.GetAppDesktop()

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


personal_lib_dir = oDesktop.GetProjectDirectory & "\PersonalLib\"

projectname = args(0)
designname = args(1)
number_of_parameters = CDbl(args(2))

''''''''''''''''''''''''''''''''''''''''''''
' copy .adk file to personal lib and rename extension to .hfss
Dim sCurPath
sCurPath = CreateObject("Scripting.FileSystemObject").GetAbsolutePathName(".")


project_pathandname = sCurPath & "\custom_library\" & projectname & ".adk"



Dim oFSO
sSourceFile = project_pathandname
sDestinationFile = personal_lib_dir & projectname & ".hfss"

list_of_open_projects = oDesktop.GetProjectList()


'first section checks to see if the file is open already. if it is the user can 
' update the project with current variable or than can create a new project with a new name

is_open = 0
for n=0 to (UBound(list_of_open_projects))

 if list_of_open_projects(n) = projectname then
   is_open = 1
   response = msgbox("Project Is Already Open -" & (Chr(13)) & "Yes - Update Parameters" & (Chr(13)) & "No - Rename File", 4, "File Open") 'returns 6 for yes, 7 for no 'returns 6 for yes, 7 for no
      if response = 7 then
          current_time = Replace(time,":",".")
          current_time = Replace(current_time," ","")
          projectname =  InputBox("Enter a new file name", "Antenna Design Kit File Rename", projectname &  "_" & current_time )
      
          Set oFSO = CreateObject("Scripting.FileSystemObject")
          sDestinationFile = personal_lib_dir & projectname & ".hfss"
    i=0
    Do  
      If oFSO.FileExists(sDestinationFile) Then
        current_time = Replace(time,":",".")
        current_time = Replace(current_time," ","")
        projectname =  InputBox("File name already exists, please rename", "Antenna Design Kit File Rename", projectname &  "_" & current_time)
        sDestinationFile = personal_lib_dir & projectname & ".hfss"
        'oFSO.CopyFile sSourceFile, sDestinationFile
        'oDesktop.OpenProject sDestinationFile
        i=0
      else
        oFSO.CopyFile sSourceFile, sDestinationFile
        oDesktop.OpenProject sDestinationFile
        i=1
      end if
    Loop Until i=1
      
 end if
 
 
 end if
 
next

'this second section will only run if the previos section determined that no file was currently open with the
'same project name. if a file excists with teh same project name the user has the option to open the file and update it or
'rename and save as a new file name

if is_open = 0 then
	Set oFSO = CreateObject("Scripting.FileSystemObject")
  

  If oFSO.FileExists(sDestinationFile) Then	

        
        response = msgbox("File Exists -" & (Chr(13)) & "Yes - Update Parameters" & (Chr(13)) & "No - Rename File", 4, "File Exists") 'returns 6 for yes, 7 for no
      
      
        if response = 7 then
          current_time = Replace(time,":",".")
          current_time = Replace(current_time," ","")
          projectname =  InputBox("Enter a new file name", "Antenna Design Kit File Rename", projectname &  "_" & current_time)
          sDestinationFile = personal_lib_dir & projectname & ".hfss"
          i=0
          Do
            If oFSO.FileExists(sDestinationFile) Then	
                current_time = Replace(time,":",".")
                current_time = Replace(current_time," ","")
                projectname =  InputBox("Enter a new file name", "Antenna Design Kit File Rename", projectname &  "_" & current_time)
                sDestinationFile = personal_lib_dir & projectname & ".hfss"
                i=0
            else
              oFSO.CopyFile sSourceFile, sDestinationFile
              oDesktop.OpenProject sDestinationFile
              i=1
            end if 
          Loop Until i=1
        else 
          If oFSO.FileExists(sDestinationFile & ".lock") Then  'check to see if a lock file excists and gives user option to delete
              lock_file_response = msgbox("A lock file exists -"  & (Chr(13)) & "Yes - Delete Lock File and Open" & (Chr(13)) & "No - To Cancel", 4, "Lock File") 'returns 6 for yes, 7 for no
                  if lock_file_response = 6 then
                    oFSO.DeleteFile sDestinationFile & ".lock" 
                  else
                    WScript.Quit
                  end if
        end if
    
   
    
    oDesktop.OpenProject sDestinationFile
    i=1
    
    end if
    

Else

	oFSO.CopyFile sSourceFile, sDestinationFile
  oDesktop.OpenProject sDestinationFile
'Wend 
End If
	' Clean Up

	Set oFSO = Nothing
end if



Set oProject = oDesktop.SetActiveProject(projectname)
Set oDesign = oProject.SetActiveDesign(designname)




dim number_of_global
dim number_of_local
number_of_global = 0
number_of_local = 0


'counts total number of global and local params
For count = 1 to 2*number_of_parameters
is_global_var=InStr(CStr(args(2+count)),"$")
  if is_global_var = 1 then
       number_of_global = number_of_global+1
       count = count + 1 'increment 1 extra because variables are passed in in format Param_Name Param_Value
  else
    number_of_local = number_of_local+1
    count = count + 1 'increment 1 extra because variables are passed in in format Param_Name Param_Value
end if
Next

dim ChangedProps_local_array
dim ChangedProps_global_array
Dim props_array_local
Dim props_array_global
 
redim ChangedProps_local_array(CInt(number_of_local))  'array to construct ChangedProps command
redim ChangedProps_global_array(CInt(number_of_global))  'array to construct ChangedProps command
dim local_var_count
dim global_var_count
global_var_count = 0
local_var_count = 0

'Change variables to meet value
For count = 1 to 2*number_of_parameters

is_global_var=InStr(CStr(args(2+count)),"$")


  if is_global_var = 1 then
    props_array_global = Array("NAME:" & args(2+count), "Value:=",args(2+(count+1)))
    count = count + 1 'increment 1 extra because variables are passed in in format Param_Name Param_Value
    ChangedProps_global_array(global_var_count+1) = props_array_global 
    global_var_count = global_var_count+1
       
  else
    props_array_local = Array("NAME:" & args(2+count), "Value:=",args(2+(count+1)))
    count = count + 1 'increment 1 extra because variables are passed in in format Param_Name Param_Value
    ChangedProps_local_array(local_var_count+1) = props_array_local 
    local_var_count = local_var_count+1
    
  end if
  

Next

ChangedProps_local_array(0)="NAME:ChangedProps"
ChangedProps_global_array(0)="NAME:ChangedProps"

if number_of_local >= 1 then
changedproperty_command_local = Array("NAME:AllTabs", Array("NAME:LocalVariableTab", Array("NAME:PropServers",  _
  "LocalVariables"), ChangedProps_local_array))
oDesign.ChangeProperty changedproperty_command_local
end if
  
if number_of_global >= 1 then
changedproperty_command_global = Array("NAME:AllTabs", Array("NAME:ProjectVariableTab", Array("NAME:PropServers",  _
  "ProjectVariables"), ChangedProps_global_array))  
oDesign.ChangeProperty changedproperty_command_global

end if

   

    
    
 Setlocale(locallang) 







