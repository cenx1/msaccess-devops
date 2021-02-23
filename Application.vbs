'---------------------------------------------------------------------------------------
' Application 	: Script to launch Access Database (Installed via ClickOnce)
' Author    	: Adam Waller
' Date      	: 8/25/2017
'---------------------------------------------------------------------------------------

'=================================================
'	** SET THESE PARAMETERS **
'=================================================

Const strAppName = "Contacts CRM"
Const strAppFile = "Contacts CRM.accdb"

'=================================================


' Set application as trusted.
VerifyTrustedLocation strAppName

' Use Windows Shell to launch the application.
CreateObject("WScript.Shell").Run  """msaccess.exe"" """ & ScriptPath & strAppFile & """ /cmd ""ojHYrvAwMudK8pezm7AR"""


'---------------------------------------------------------------------------------------
' Function 	: ScriptPath
' Author    : Adam Waller
' Date      : 2/8/2017
' Purpose   : Get the path to the folder where this script is running.
'---------------------------------------------------------------------------------------
Function ScriptPath()
	
	Dim oFSO
	Dim oFile
	
	Set oFSO = CreateObject("Scripting.FileSystemObject")
	set oFile = oFSO.GetFile(Wscript.ScriptFullName)
	ScriptPath = oFSO.GetParentFolderName(oFile) & "\"
	Set oFSO = Nothing
	
End Function


'---------------------------------------------------------------------------------------
' Function 	: VerifyTrustedLocation
' Author    : Adam Waller
' Date      : 1/24/2017
' Purpose   : Run this proceedure on startup to make sure the database is located
'           : in a trusted location. (Adding an entry if needed.)
'---------------------------------------------------------------------------------------
'
Function VerifyTrustedLocation(strAppName)

    Dim oShell
    Dim oFSO
    Dim oFile
    Dim strVersion
    Dim strPath
    Dim strAppPath
    Dim blnCreate
    Dim strVal
    
    Set oShell = CreateObject("WScript.Shell")
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    
    ' Get Access version
    On Error Resume Next
    strVal = oShell.RegRead("HKEY_CLASSES_ROOT\Access.Application\CurVer\")
    If Err Then
        Err.Clear
    Else
        ' Parse the version number
        strVal = Right(strVal, 2)
        If IsNumeric(strVal) Then strVersion = strVal & ".0"
    End If
    On Error GoTo 0
    
    ' Make sure we actually found a version number
    If Len(strVersion) <> 4 Then
        MsgBox "Unable to determine Microsoft Access Version.", vbCritical
        Exit Function
    End If
    
    ' Get application name
    'strAppName = Application.VBE.ActiveVBProject.Name
    
    ' Get registry path for trusted locations
    strPath = "HKEY_CURRENT_USER\Software\Microsoft\Office\" & _
        strVersion & "\Access\Security\Trusted Locations\" & strAppName & "\"
    
    ' Attempt to read the key
    On Error Resume Next
    strVal = oShell.RegRead(strPath & "Path")
    If Err Then
        Err.Clear
        blnCreate = True
    End If
    On Error GoTo 0
    
    ' Get script location, to find application path
    strAppPath = WScript.ScriptFullName
    'strAppPath = CodeProject.Path & "\" & CodeProject.Name
    Set oFile = oFSO.GetFile(strAppPath)
    strAppPath = oFSO.GetParentFolderName(oFile)
    
    If blnCreate = True Then
        ' Create values
        oShell.RegWrite strPath & "Path", strAppPath
        oShell.RegWrite strPath & "Date", Now()
        oShell.RegWrite strPath & "Description", strAppName
        oShell.RegWrite strPath & "AllowSubfolders", 1, "REG_DWORD"
    Else
        ' Verify path location
        strVal = oShell.RegRead(strPath & "Path")
        If strVal <> strAppPath Then
            ' Update value
            oShell.RegWrite strPath & "Path", strAppPath
            oShell.RegWrite strPath & "Date", Now()
        End If
    End If
    
    ' Release references
    Set oShell = Nothing
    Set oFSO = Nothing
    
End Function