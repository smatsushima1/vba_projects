Attribute VB_Name = "Module1"
Public error_code As String

Sub Macro6()

Dim folder, error_code As String

folder = "C:\Users\smats\Documents\OFFICE\FOLDER_GENERATOR_DEV"
error_message = Error(1)

If Dir(folder, vbDirectory) <> "" Then
  MsgBox "folder exists"
Else
  MsgBox error_message
End If

Macro2

End Sub
Sub main_dev()

Macro1
MsgBox error_code

Macro2

Macro3
MsgBox error_code

End Sub
Sub Macro1()

error_code = 0

MsgBox "macro1"

End Sub
Sub Macro2()

MsgBox "macro2"

End Sub
Sub Macro3()

If error_code = 0 Then
  Exit Sub
End If

MsgBox "macro3"

End Sub
Sub macro4()

Dim doc As Document

Set doc = Documents("macro_dev.docm")

MsgBox doc.Name

End Sub
Sub macro5()

MsgBox ThisDocument.Name

End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'- Public variables must be saved outside of functions to be used throughout
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public error_code, _
       error_message, _
       file_name, _
       description, _
       initials, _
       pr, _
       ige, _
       supp_service, _
       psc, _
       naics, _
       ja, _
       delivery_date, _
       requirement_type, _
       it, _
       directory_name, _
       sap_folder, _
       large_folder1, _
       large_folder2, _
       large_folder3, _
       large_folder4, _
       large_folder5, _
       large_folder6 _
  As String
Sub generateRequirement()

Application.ScreenUpdating = False

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'STEP 1 - RESET VARIABLES
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

resetVariables

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'STEP 2 - SAVE INPUT AS PUBLIC VARIABLES
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

saveInput

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'STEP 3 - CREATE INITIAL FOLDER IN SPECIFIED LOCATION
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

createFolder

If error_code = 1 Then
  MsgBox error_message, Title:="Error"
  resetVariables
  Application.ScreenUpdating = True
  Exit Sub
End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'STEP 4 - CREATE SUBFOLDERS WITHIN MAIN FOLDER
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

createSubfolders

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'STEP 5 - POPULATE FORMS WITH DATA AND SAVE IN SUBFOLDERS
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

populateForms

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'FINISHED
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

MsgBox "E-file requirement successfully generated!", Title:="Success!"

resetVariables

Application.ScreenUpdating = True

End Sub
Sub resetVariables()

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'STEP 1 - RESET VARIABLES
'- reset all variables before running all other steps
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

error_code = vbNullString
error_message = vbNullString
file_name = vbNullString
description = vbNullString
initials = vbNullString
pr = vbNullString
ige = vbNullString
supp_service = vbNullString
psc = vbNullString
naics = vbNullString
ja = vbNullString
delivery_date = vbNullString
requirement_type = vbNullString
it = vbNullString
directory_name = vbNullString
sap_folder = vbNullString
large_folder1 = vbNullString
large_folder2 = vbNullString
large_folder3 = vbNullString
large_folder4 = vbNullString
large_folder5 = vbNullString
large_folder6 = vbNullString

End Sub
Sub saveInput()

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'STEP 2 - SAVE INPUT AS PUBLIC VARIABLES
'- first reset all public variables, then reassign them
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

file_name = "macro_dev"

With Documents(file_name)
  description = .SelectContentControlsByTitle("DESCRIPTION").Item(1).Range.Text
  initials = .SelectContentControlsByTitle("INITIALS").Item(1).Range.Text
  pr = .SelectContentControlsByTitle("PR").Item(1).Range.Text
  ige = .SelectContentControlsByTitle("IGE").Item(1).Range.Text
  supp_service = .SelectContentControlsByTitle("SUPPLY/SERVICE").Item(1).Range.Text
  psc = .SelectContentControlsByTitle("PSC").Item(1).Range.Text
  naics = .SelectContentControlsByTitle("NAICS").Item(1).Range.Text
  ja = .SelectContentControlsByTitle("J&A").Item(1).Range.Text
  delivery_date = .SelectContentControlsByTitle("DELIVERY DATE").Item(1).Range.Text
  requirement_type = .SelectContentControlsByTitle("REQUIREMENT TYPE").Item(1).Range.Text
  it = .SelectContentControlsByTitle("IT").Item(1).Range.Text
End With

End Sub
Sub createFolder()

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'STEP 3 - CREATE INITIAL FOLDER IN SPECIFIED LOCATION
'- if folder is selected (Show = -1), then save path as folder_path; else, exit
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim message, folder_path As String

message = "Select the folder where your requirement will be saved"

With Application.FileDialog(msoFileDialogFolderPicker)
  MsgBox message & ".", Title:="Select Folder"
  .Title = message
  If .Show = -1 Then
    folder_path = .SelectedItems(1)
  Else
    error_code = 1
    error_message = "No folder selected."
    Exit Sub
  End If
End With

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'- Concatenate pr, initials, and description with folder_path
'- if folder already exists, exit so as not to overwrite files
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

directory_name = folder_path & "\" & pr & ", " & initials & ", " & description

If Dir(directory_name, vbDirectory) = "" Then
  MkDir directory_name
Else
  error_code = 1
  error_message = "Folder already exists."
  Exit Sub
End If

End Sub
Sub createSubfolders()

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'STEP 4 - CREATE SUBFOLDERS WITHIN MAIN FOLDER
'- if SAP, then just create "WORKING" folder
'- if Large, then create subfolders
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim working As String

working = "\WORKING"

If requirement_type = "SAP" Then
  sap_folder = directory_name & working
  MkDir sap_folder
Else
  large_folder1 = directory_name & "\1 PLANNING"
  large_folder2 = directory_name & "\2 SOLICITATION"
  large_folder3 = directory_name & "\3 EVALUATION"
  large_folder4 = directory_name & "\4 AWARD"
  large_folder5 = directory_name & "\5 POST AWARD"
  large_folder6 = directory_name & "\6 CONTRACT AND MODS"
  MkDir large_folder1
  MkDir large_folder2
  MkDir large_folder3
  MkDir large_folder4
  MkDir large_folder5
  MkDir large_folder6
  large_folder1 = large_folder1 & working
  large_folder2 = large_folder2 & working
  large_folder3 = large_folder3 & working
  large_folder4 = large_folder4 & working
  large_folder5 = large_folder5 & working
  large_folder6 = large_folder6 & working
  MkDir large_folder1
  MkDir large_folder2
  MkDir large_folder3
  MkDir large_folder4
  MkDir large_folder5
  MkDir large_folder6
End If

End Sub
Sub populateForms()

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'STEP 5 - POPULATE FORMS WITH DATA AND SAVE IN SUBFOLDERS
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim forms_path As String

forms_path = Documents(file_name).Path & "\FORMS\"

If requirement_type = "SAP" Then
  Documents.Open forms_path & "Blank Form.docx"
  With Documents("Blank Form.docx")
    .SelectContentControlsByTitle("PR").Item(1).Range.Text = pr
    .SelectContentControlsByTitle("IGE").Item(1).Range.Text = ige
    .SaveAs2 sap_folder & "\Blank Form.docx"
    .Close
  End With
End If

End Sub
Private Sub GENERATE_Click()

generateRequirement

End Sub

