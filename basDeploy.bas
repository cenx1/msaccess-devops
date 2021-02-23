'=======================================================================================
'                   This module is updated automatically from the Code Templates
'   **IMPORTANT**   database. If you need to make custom changes, please change
'                   the following line to turn off the automatic updating for this
'                   item. (True for automatic updates, False to disable updates)
'=======================================================================================
'@AutoUpdate = True


'---------------------------------------------------------------------------------------
' Module    : basDeploy
' Author    : Adam Waller
' Date      : 8/24/2019
' Purpose   : Deploy an update to an Access Database application.
'           : Version number is stored in a custom property in the local database.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Private Module
Option Explicit

'---------------------------------------------------------------------------------------
'   DEFAULT USER CONFIGURED OPTIONS
'   (Uses saved registry values instead, if they exist. See SaveOptions sub below)
'---------------------------------------------------------------------------------------

' Specify the path to the deployment folder. (UNC path not supported)
Private Const DEPLOY_FOLDER As String = "T:\Apps\Deploy\"

' Pipe delimited list of root project folders that should not be deployed.
Private Const IGNORE_FOLDERS As String = "dev|cache"
'---------------------------------------------------------------------------------------


' Used for debug output display
Private Const cstrSpacer As String = "---------------------------------------------------------------------------------------"

' Constants so we don't have to use the VBE reference in projects
Private Const vbext_ct_StdModule As Integer = 1
Private Const vbext_ct_ClassModule As Integer = 2

' Collection of versions as read from `Latest Versions.csv`
Private mVersions As Collection
Private mUpdates As Collection

' Enum to improve code readability.
' These match the columns in the
' Latest Versions.csv file.
Private Enum eVersion
    evName
    evVersion
    evDate
    evFile
    evType
    evNotes
End Enum


'---------------------------------------------------------------------------------------
' Procedure : Deploy
' Author    : Adam Waller
' Date      : 1/5/2017
' Purpose   : Deploys the program for end users to install and run.
'           : Returns true if the deployement process was completed.
'---------------------------------------------------------------------------------------
'
Public Function Deploy(Optional blnIgnorePendingUpdates As Boolean = False) As Boolean
    
    Dim strPath As String
    Dim strTools As String
    Dim strName As String
    Dim strCmd As String
    Dim strIcon As String
    Dim objComponent As Object
    Dim strFile As String
    Dim intCnt As Integer
    
    ' Make sure we don't accidentally deploy a nested library!
    If CodeProject.FullName <> CurrentProject.FullName Then
        Debug.Print " ** WARNING ** " & CodeProject.Name & " is not the top-level project!"
        Debug.Print " Switching to " & CurrentProject.Name & "..."
        Set VBE.ActiveVBProject = GetVBProjectForCurrentDB
        ' Fire off deployment from primary database.
        Run "[" & GetVBProjectForCurrentDB.Name & "].Deploy"
        Exit Function
    End If
    
    ' Show debug output
    Debug.Print vbCrLf & cstrSpacer
    Debug.Print "Deployment Started - " & Now()
    
    ' Check for any updates to dependent libraries
    If CheckForUpdates Then
        If Not blnIgnorePendingUpdates Then
            Debug.Print cstrSpacer
            Debug.Print " *** UPDATES AVAILABLE *** "
            Debug.Print "Please install before deployment or set flag "
            Debug.Print "to continue deployment anyway. I.e. `Deploy True`" & vbCrLf & cstrSpacer
            Exit Function
        End If
    End If
    
    ' Check for reference issues with dependent modules
    If HasDuplicateProjects Then
        Select Case Eval("MsgBox('Would you like to run ''LocalizeReferences'' first?@Some VBA projects appear duplicated which usually indicates non-local references.@Select ''No'' to continue anyway or ''Cancel'' to cancel the deployment.@" & _
                "(Library databases that are only used as a part of other applications are typically not deployed as ClickOnce installers.)@',35)")
            Case vbYes
                Call LocalizeReferences
                Exit Function
            Case vbNo
                ' Continue anyway.
            Case Else
                Exit Function
        End Select
    End If
    
    ' Increment build number
    IncrementBuildVersion
    
    ' List project and new build number
    Debug.Print " ~ " & VBE.ActiveVBProject.Name & " ~ Version " & AppVersion
    Debug.Print cstrSpacer
    
    ' Update project description
    VBE.ActiveVBProject.Description = "Version " & AppVersion & " deployed on " & Date
    
    ' Get deployment folder (Create if needed)
    ' Note: This is the version-specific folder for this release.
    strPath = GetDeploymentVersionFolder
    
    ' Check flag for ClickOnce deployment.
    If IsClickOnce Then
    
        ' Copy project files
        Debug.Print "Copying Files";
        Debug.Print vbCrLf & CopyFiles(CodeProject.Path & "\", strPath, True) & " files copied."
        
        ' Get tools folder
        strTools = GetDeploymentFolder & "_Tools\"
        
        ' Copy manifest templates to project
        strName = VBE.ActiveVBProject.Name
        
        ' Build shell command
        strCmd = "cmd /s /c " & strTools & "Deploy.bat """ & strName & """ " & AppVersion
        
        ' Add application icon if one exists in the application folder.
        strIcon = Dir(CodeProject.Path & "\*.ico")
        If strIcon <> "" Then strCmd = strCmd & " """ & strIcon & """"
        
        ' Compile and build clickonce installation
        Shell strCmd, vbNormalFocus
        
        ' Print final status message.
        Debug.Print "Files Copied. Please review command window for any errors." & vbCrLf & cstrSpacer
    
    Else
        ' Code templates are handled just a little differently
        If CodeProject.Name = "Code Templates.accdb" Then
            
            ' Build path for exported template components.
            strPath = GetDeploymentFolder & "Code Templates\"
            
            ' Loop through all component objects, exporting each one to a file.
            For Each objComponent In GetVBProjectForCurrentDB.VBComponents
                With objComponent
                    If .Name <> "basInternal" Then
                        strFile = strPath & .Name & ".bas"
                        ' Remove any existing file before exporting
                        If Dir(strFile) <> vbNullString Then Kill strFile
                        .Export strFile
                        intCnt = intCnt + 1
                    End If
                End With
            Next objComponent
            Debug.Print intCnt & " templates deployed to " & strPath
        Else
            ' Probably a code library without a click-once installer.
            ' Deploy just the library versioned folder. (Do not include dependent
            ' libraries or we could cause some real issues with versions overwriting
            ' each other's dependencies as different libraries are updated.)
            Debug.Print "Copying " & CodeProject.Name
            CreateObject("Scripting.FileSystemObject").CopyFile CodeProject.FullName, strPath
            Debug.Print "Library deployed."
        End If
    End If
    
    ' Update list of latest versions.
    LoadVersionList
    UpdateVersionInList
    SaveVersionList
    Deploy = True
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : UpdateDependencies
' Author    : Adam Waller
' Date      : 4/3/2020
' Purpose   : Updates all dependencies in current project with latest versions.
'           : Displays errors if files are not found.
'---------------------------------------------------------------------------------------
'
Public Sub UpdateDependencies()

    Dim blnUpdatedLibraries As Boolean
    
    ' Make sure we check for updates before updating dependencies.  :-)
    If mUpdates Is Nothing Then CheckForUpdates
    
    ' See if we actually have some updates to process.
    If mUpdates.Count = 0 Then
        ' Nothing to update
        Debug.Print "No updates found."
    Else
        blnUpdatedLibraries = (UpdateLibraries > 0)
        UpdateVBAComponents
        Debug.Print cstrSpacer
        Debug.Print "Update complete. ";
        
        ' Prompt user to restart database if needed.
        If blnUpdatedLibraries Then
            Debug.Print "DATABASE RESTART REQUIRED"
            Debug.Print "Please close and reopen this database to apply changes."
        Else
            Debug.Print
            ' Make sure we are working in the current project
            Set VBE.ActiveVBProject = GetVBProjectForCurrentDB
            Debug.Print "Compiling and saving modules..."
            ' Compile and save all code modules
            Application.RunCommand acCmdCompileAndSaveAllModules
            Debug.Print "Done."
        End If
    End If
    
    ' Reset the updates collection
    Set mUpdates = Nothing
       
End Sub


'---------------------------------------------------------------------------------------
' Procedure : UpdateVBAComponents
' Author    : Adam Waller
' Date      : 4/3/2020
' Purpose   : Update the VBA objects like modules and classes
'---------------------------------------------------------------------------------------
'
Private Function UpdateVBAComponents() As Integer

    Dim varUpdate As Variant
    Dim strFile As String
    Dim intUpdated As Integer
    Dim cmp As Object ' VBComponent
    
    ' Loop through updates and process any components
    For Each varUpdate In mUpdates
        ' Check for component type update
        If varUpdate(evType) = "Component" Then
            ' Make sure it comes from Code Templates
            If GetFileNameFromPath(CStr(varUpdate(evFile))) = "Code Templates.accdb" Then
                ' Coming from our code templates. Get path to latest file.
                strFile = GetDeploymentFolder & "Code Templates\" & varUpdate(evName) & ".bas"
                ' Make sure file exists
                If Dir(strFile) = "" Then
                    Debug.Print "ERROR: Could not find " & strFile
                Else
                    If varUpdate(evName) = "basDeploy" Then
                        UpdateDeployModule
                    Else
                        Set cmp = GetVBProjectForCurrentDB.VBComponents(varUpdate(evName))
                        If cmp.Type = vbext_ct_ClassModule Or cmp.Type = vbext_ct_StdModule Then
                            ' Remove existing module and replace with file
                            With GetVBProjectForCurrentDB.VBComponents
                                .Remove .Item(varUpdate(evName))
                                .Import strFile
                            End With
                        Else
                            ' Other components like forms. Replace code module from file.
                            ' (Could extend this later to replace entire object, but start with this.)
                            With cmp.CodeModule
                                .DeleteLines 1, .CountOfLines
                                .AddFromFile strFile
                            End With
                        End If
                        Debug.Print "Updated " & varUpdate(evName)
                        intUpdated = intUpdated + 1
                    End If
                End If
            End If
        End If
    Next varUpdate
    
    ' Return number of components updated
    UpdateVBAComponents = intUpdated

End Function


'---------------------------------------------------------------------------------------
' Procedure : UpdateLibraries
' Author    : Adam Waller
' Date      : 4/3/2020
' Purpose   : Update library databases. (Only auto-update the linked library itself,
'           : not any other dependencies, lest we create version issues.)
'---------------------------------------------------------------------------------------
'
Private Function UpdateLibraries() As Integer

    Dim varUpdate As Variant
    Dim strExisting As String
    Dim strFile As String
    Dim intUpdated As Integer
    
    ' Loop through updates and process any components
    For Each varUpdate In mUpdates
        ' Check for (library) file type update
        If varUpdate(evType) = "File" Then
            ' Build full path to file
            strFile = GetDeploymentFolder & varUpdate(evName) & "\" & varUpdate(evVersion) & "\" & varUpdate(evFile)
            ' Make sure file exists
            If Dir(strFile) = "" Then
                Debug.Print "ERROR: Could not find " & strFile
            Else
                ' Check for referenced file
                strExisting = CodeProject.Path & "\" & varUpdate(evFile)
                If Dir(strExisting) = vbNullString Then
                    Debug.Print "ERROR: Could not find existing library: " & strExisting
                Else
                    ' Replace existing file
                    CreateObject("Scripting.FileSystemObject").CopyFile strFile, strExisting, True
                    Debug.Print "Updated " & varUpdate(evName) & " to version " & varUpdate(evVersion)
                    intUpdated = intUpdated + 1
                End If
            End If
        End If
    Next varUpdate

    ' Return number of libraries updated.
    UpdateLibraries = intUpdated
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetIgnoredFolders
' Author    : Adam Waller
' Date      : 8/24/2019
' Purpose   : Returns the pipe delimited list of ignored folders
'---------------------------------------------------------------------------------------
'
Private Function GetIgnoredFolders() As String

    ' Use value saved in registry, or fall back to default constant
    GetIgnoredFolders = GetSetting("DevOps", "SE API", "Ignore Folders", IGNORE_FOLDERS)

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetDeploymentFolder
' Author    : Adam Waller
' Date      : 8/24/2019
' Purpose   : Returns path to base deployment folder.
'---------------------------------------------------------------------------------------
'
Private Function GetDeploymentFolder() As String
    
    Dim strPath As String
    Dim strTest As String
    
    ' Use path saved in registry, or fall back to default constant
    strPath = GetSetting("DevOps", "SE API", "Deploy Folder", DEPLOY_FOLDER)
    
    ' Make sure the folder exists before we continue.
    ' (It might not in an environment where the registry override is being used,
    '  and the application is being deployed from a new development computer or
    '  user profile that doesn't have the override configured.)
    
TestPath:
    strTest = Left(strPath, Len(strPath) - 1)
    If Len(Nz(Dir(strTest, vbDirectory))) < 3 Then
        If MsgBox("Deployment path '" & strPath & "' not found." & vbCrLf & _
            "Would you like to enter a custom path to use instead?" & vbCrLf & _
            "The custom path will be saved in this user profile for future deployments.", vbQuestion + vbYesNo) = vbYes Then
            strPath = InputBox("Enter path to deployment folder:", , DEPLOY_FOLDER)
            ' Save the new selection
            SaveOptions strPath, GetIgnoredFolders
            ' Test the newly entered path before using it.
            GoTo TestPath
        Else
            ' Revert to default if they didn't want to create a custom one.
            strPath = DEPLOY_FOLDER
        End If
    End If
    
    ' Return path
    GetDeploymentFolder = strPath
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : SaveOptions
' Author    : Adam Waller
' Date      : 8/24/2019
' Purpose   : Save a set of user options to the registry, to override the default
'           : constants for this computer/user profile.
'---------------------------------------------------------------------------------------
'
Private Sub SaveOptions(strDeployFolder As String, strIgnoreFolders As String)

    SaveSetting "DevOps", "SE API", "Deploy Folder", strDeployFolder
    SaveSetting "DevOps", "SE API", "Ignore Folders", strIgnoreFolders
    Debug.Print "Settings saved."
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetDeploymentVersionFolder
' Author    : Adam Waller
' Date      : 8/24/2019
' Purpose   : Returns the full path to the deployment folder used for this release,
'           : including both the application name and version number.
'           : (I.e. T:\Apps\Deploy\SE API\1.0.0.12\
'---------------------------------------------------------------------------------------
'
Private Function GetDeploymentVersionFolder() As String
    
    Dim strPath As String
    Dim strProject As String
    Dim strVersion As String
    
    strPath = GetDeploymentFolder
    strProject = VBE.ActiveVBProject.Name
    strVersion = AppVersion
    
    ' Build out full path for deployment
    strPath = strPath & strProject
    If Dir(strPath, vbDirectory) = "" Then
        ' Create project folder
        MkDir strPath
    End If
    
    strPath = strPath & "\" & strVersion
    If Dir(strPath, vbDirectory) = "" Then
        ' Create version folder
        MkDir strPath
    End If
    
    ' Return full path
    GetDeploymentVersionFolder = strPath & "\"
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : AppVersion
' Author    : Adam Waller
' Date      : 1/5/2017
' Purpose   : Get the version from the database property.
'---------------------------------------------------------------------------------------
'
Public Property Get AppVersion() As String
    Dim strVersion As String
    strVersion = GetDBProperty("AppVersion")
    If strVersion = "" Then strVersion = "1.0.0.0"
    AppVersion = strVersion
End Property


'---------------------------------------------------------------------------------------
' Procedure : AppVersion
' Author    : Adam Waller
' Date      : 1/5/2017
' Purpose   : Set version property in current database.
'---------------------------------------------------------------------------------------
'
Public Property Let AppVersion(strVersion As String)
    SetDBProperty "AppVersion", strVersion
End Property


'---------------------------------------------------------------------------------------
' Procedure : GetDBProperty
' Author    : Adam Waller
' Date      : 9/1/2017
' Purpose   : Get a database property
'---------------------------------------------------------------------------------------
'
Public Function GetDBProperty(strName As String) As Variant

    Dim prp As Object   ' Access.AccessObjectProperty
    
    For Each prp In PropertyParent.Properties
        If prp.Name = strName Then
            GetDBProperty = prp.Value
            Exit For
        End If
    Next prp
    
    Set prp = Nothing
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : SetDBProperty
' Author    : Adam Waller
' Date      : 9/1/2017
' Purpose   : Set a database property
'---------------------------------------------------------------------------------------
'
Public Sub SetDBProperty(strName As String, varValue, Optional prpType = DB_TEXT)

    Dim prp As Object   ' Access.AccessObjectProperty
    Dim prpAccdb As Property
    Dim blnFound As Boolean
    Dim dbs As Database
    
    For Each prp In PropertyParent.Properties
        If prp.Name = strName Then
            blnFound = True
            ' Skip set on matching value
            If prp.Value = varValue Then Exit Sub
            Exit For
        End If
    Next prp
    
    On Error Resume Next
    If blnFound Then
        PropertyParent.Properties(strName).Value = varValue
    Else
        If CurrentProject.ProjectType = acADP Then
            PropertyParent.Properties.Add strName, varValue
        Else
            ' Normal accdb database property
            Set dbs = CurrentDb
            Set prpAccdb = dbs.CreateProperty(strName, DB_TEXT, varValue)
            dbs.Properties.Append prpAccdb
            Set dbs = Nothing
        End If
    End If
    If Err Then Err.Clear
    On Error GoTo 0

End Sub


'---------------------------------------------------------------------------------------
' Procedure : PropertyParent
' Author    : Adam Waller
' Date      : 1/30/2017
' Purpose   : Get the correct parent type for database properties (including custom)
'---------------------------------------------------------------------------------------
'
Private Function PropertyParent() As Object
    ' Get correct parent project type
    If CurrentProject.ProjectType = acADP Then
        Set PropertyParent = CurrentProject
    Else
        Set PropertyParent = CurrentDb
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : IncrementBuildVersion
' Author    : Adam Waller
' Date      : 1/6/2017
' Purpose   : Increments the build version (1.0.0.x)
'---------------------------------------------------------------------------------------
'
Public Sub IncrementBuildVersion()
    Dim varParts As Variant
    Dim intVer As Integer
    varParts = Split(AppVersion, ".")
    If UBound(varParts) < 3 Then Exit Sub
    intVer = varParts(UBound(varParts))
    varParts(UBound(varParts)) = intVer + 1
    AppVersion = Join(varParts, ".")
End Sub


'---------------------------------------------------------------------------------------
' Procedure : CopyFiles
' Author    : Adam Waller
' Date      : 1/5/2017
' Purpose   : Recursive function to copy files from one folder to another.
'           : (Set to ignore certain files)
'---------------------------------------------------------------------------------------
'
Private Function CopyFiles(strSource As String, strDest As String, blnOverwrite As Boolean) As Double
    
    Dim strFile As String
    Dim dblCnt As Double
    Dim strBase As String
    Dim objFSO As Object
    Dim oFolder As Object
    Dim oFile As Object
    Dim blnExists As Boolean
    Dim varSkip As Variant
    Dim varItem As Variant
    
    ' Requires FSO to copy open database files. (VBA.FileCopy gives a permission denied error.)
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set oFolder = objFSO.GetFolder(strSource)
    
    ' Ignore certain types of base folders
    strBase = CurrentProject.Path
    varSkip = Split(GetIgnoredFolders, "|")
    For Each varItem In varSkip
        If strSource = strBase & "\" & CStr(varItem) & "\" Then
            ' Ignore this folder
            Exit Function
        End If
    Next varItem
    
    ' Copy files then folders
    For Each oFile In objFSO.GetFolder(strSource).Files
        strFile = oFile.Name
        Select Case True
            ' Files to skip
            Case strFile Like ".git*"
            Case strFile Like "*.laccdb"
            Case Else
                blnExists = Dir(strDest & strFile) <> ""
                If blnExists And Not blnOverwrite Then
                    ' Skip this file
                Else
                    If blnExists Then Kill strDest & strFile
                    oFile.Copy strDest & strFile
                    ' Show progress point as each file is copied
                    dblCnt = dblCnt + 1
                    Debug.Print ".";
                End If
        End Select
    Next oFile
    
    ' Copy folders
    For Each oFolder In objFSO.GetFolder(strSource).SubFolders
        strFile = oFolder.Name
        Select Case True
            ' Files to skip
            Case strFile = CodeProject.Name & ".src"    ' This project
            Case strFile Like "*.src"                   ' Other source files
            Case strFile Like ".git*"
            Case Else
                ' Check if folder already exists in destination
                If Dir(strDest & strFile, vbDirectory) = "" Then
                    MkDir strDest & strFile
                    ' Show progress after creating folder but before copying files
                    Debug.Print ".";
                End If
                ' Recursively copy files from this folder
                dblCnt = dblCnt + CopyFiles(strSource & strFile & "\", strDest & strFile & "\", blnOverwrite)
        End Select
    Next oFolder
    
    ' Release reference to objects.
    Set objFSO = Nothing
    Set oFile = Nothing
    Set oFolder = Nothing
    
    ' Return count of files copied.
    CopyFiles = dblCnt
    
End Function



'---------------------------------------------------------------------------------------
' Procedure : CheckForUpdates
' Author    : Adam Waller
' Date      : 1/27/2017
' Purpose   : Check for updates to library databases or template modules
'---------------------------------------------------------------------------------------
'
Public Function CheckForUpdates() As Boolean

    Const vbext_rk_Project As Integer = 1
    
    Dim ref As Access.Reference
    Dim varLatest As Variant
    Dim strCurrent As String
    Dim strLatest As String
    Dim objComponent As Object
    Dim strName As String
    Dim intCnt As Integer
    Dim intLines As Integer
    Dim blnUpdatesAvailable As Boolean
    Dim intAutoUpdateCount As Integer
    
    ' We shouldn't be running this on deployed applications.
    If InStr(1, CurrentProject.Path, "\AppData\") > 1 Then Exit Function
    
    ' Reload version file before checking for updates.
    LoadVersionList
    Set mUpdates = New Collection
    
    ' Check references for updates.
    For Each ref In Application.References
        If ref.Kind = vbext_rk_Project Then
            strCurrent = GetCurrentRefVersion(ref)
            varLatest = GetLatestVersionDetails(ref.Name)
            If IsArray(varLatest) Then
                If UBound(varLatest) > 2 Then
                    strLatest = varLatest(1)
                    If strLatest <> "" Then
                        ' Compare current with latest.
                        If strCurrent <> strLatest Then
                            Debug.Print "UPDATE AVAILABLE: " & ref.Name & " (" & _
                                GetFileNameFromPath(VBE.VBProjects(ref.Name).FileName) & _
                                ") can be updated from " & strCurrent & " to " & strLatest
                            blnUpdatesAvailable = True
                            intAutoUpdateCount = intAutoUpdateCount + 1
                            mUpdates.Add varLatest
                        End If
                    End If
                End If
            End If
        End If
    Next ref
    
    ' Check code modules for updates
    For Each objComponent In GetVBProjectForCurrentDB.VBComponents
        strName = objComponent.Name
        ' Look for matching item in list
        For intCnt = 2 To mVersions.Count
            If UBound(mVersions(intCnt)) = 4 Then
                If (mVersions(intCnt)(evName) = strName) _
                    And (mVersions(intCnt)(evType) = "Component") Then
                    ' Check for different "version"
                    intLines = GetCodeLineCount(objComponent.CodeModule)
                    If mVersions(intCnt)(1) <> intLines _
                        And mVersions(intCnt)(evFile) <> CurrentProject.Name Then
                        If AllowAutoUpdate(objComponent.CodeModule) Then
                            Debug.Print "MODULE UPDATE AVAILABLE: " & strName & _
                                " can be updated from """ & mVersions(intCnt)(evFile) & """ (" & _
                                mVersions(intCnt)(evVersion) - intLines & " lines on " & _
                                mVersions(intCnt)(evDate) & ".)"
                            blnUpdatesAvailable = True
                            intAutoUpdateCount = intAutoUpdateCount + 1
                            mUpdates.Add mVersions(intCnt)
                        Else
                            Debug.Print "Manual* update available: " & strName & _
                                " can be updated from """ & mVersions(intCnt)(evFile) & """ (" & _
                                mVersions(intCnt)(evVersion) - intLines & " lines on " & _
                                mVersions(intCnt)(evDate) & ".) *This module is currently flagged to disable automatic updates."
                        End If
                    End If
                End If
            End If
        Next intCnt
    Next objComponent
    
    ' Offer to run auto-update on the available components.
    If intAutoUpdateCount > 0 Then
        Debug.Print "=========================================================================="
        Debug.Print " " & intAutoUpdateCount;
        If intAutoUpdateCount = 1 Then
            Debug.Print " update is ";
        Else
            Debug.Print " updates are ";
        End If
        Debug.Print "available for automatic installation. If you would like" & vbCrLf _
            & " to apply these updates now, please run the following command:"
        Debug.Print "=========================================================================="
        Debug.Print "UpdateDependencies"
        Debug.Print
    End If
    
    Set ref = Nothing
    Set objComponent = Nothing
    
    CheckForUpdates = blnUpdatesAvailable
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : LoadVersionList
' Author    : Adam Waller
' Date      : 1/27/2017
' Purpose   : Loads a list of the current versions.
'---------------------------------------------------------------------------------------
'
Private Function LoadVersionList() As Boolean
    
    Dim strFile As String
    Dim intFile As Integer
    Dim strLine As String
    
    strFile = GetDeploymentFolder & "Latest Versions.csv"
    intFile = FreeFile
    
    ' Initialize collection
    Set mVersions = New Collection
    
    ' Start with header if file does not exist.
    If Dir(strFile) = "" Then
        ' Create a new list.
        mVersions.Add Array("Name", "Version", "Date", "File", "Type", "Notes")
    Else
        ' Read entries in the file
        Open strFile For Input As #intFile
            Do While Not EOF(intFile)
                Line Input #intFile, strLine
                mVersions.Add Split(strLine, ",")
            Loop
        Close intFile
    End If
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : UpdateVersionInList
' Author    : Adam Waller
' Date      : 1/27/2017
' Purpose   : Update the version info in the list of current versions.
'---------------------------------------------------------------------------------------
'
Private Function UpdateVersionInList()
    
    Dim intCnt As Integer
    Dim strName As String
    Dim strNotes As String
    
    If mVersions Is Nothing Then
        MsgBox "Must load version list first.", vbExclamation
        Exit Function
    End If
    
    ' Structure of list entry:
    'varItem = Array(Name, Version, Date, File, [Type], [Notes])
    
    ' Get current project name
    strName = GetVBProjectForCurrentDB.Name
    
    ' Look for matching item in list
    For intCnt = 2 To mVersions.Count
        If UBound(mVersions(intCnt)) >= 3 Then
            If mVersions(intCnt)(0) = strName Then
                mVersions.Remove intCnt
                Exit For
            End If
        End If
    Next intCnt
    
    ' Add any release notes. (Only used in some projects)
    strNotes = GetDBProperty("VersionReleaseNotes")
    
    ' Add to list
    mVersions.Add Array(strName, AppVersion, Now, CodeProject.Name, "File", strNotes), , , 1
    
    ' Save any code templates
    If CurrentProject.Name = "Code Templates.accdb" Then SaveCodeTemplateVersions
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : SaveCodeTemplateVersions
' Author    : Adam Waller
' Date      : 2/10/2017
' Purpose   : Saves code template modules, using line count as "version".
'---------------------------------------------------------------------------------------
'
Private Sub SaveCodeTemplateVersions()
    
    Dim objComponent As Object
    Dim intCnt As Integer
    Dim blnSkip As Boolean
    Dim intLines As Integer
    Dim strName As String
    
    For Each objComponent In GetVBProjectForCurrentDB.VBComponents
        strName = objComponent.Name
        Select Case strName
        
            ' Skip anything listed here
            Case "basInternal"
                
            ' Any other components
            Case Else
            
                ' Look for matching item in list
                blnSkip = False ' Reset flag
                intLines = GetCodeLineCount(objComponent.CodeModule)
                For intCnt = 2 To mVersions.Count
                    If UBound(mVersions(intCnt)) = 4 Then
                        If (mVersions(intCnt)(0) = strName) _
                            And (mVersions(intCnt)(4) = "Component") Then
                            ' Check for different "version"
                            If mVersions(intCnt)(1) <> intLines Then
                                mVersions.Remove intCnt
                            Else
                                blnSkip = True
                            End If
                            Exit For
                        End If
                    End If
                Next intCnt
                
                ' Add to list
                If Not blnSkip Then mVersions.Add Array(objComponent.Name, intLines, Now, CodeProject.Name, "Component"), , , 1
        End Select
    Next objComponent
    
    Set objComponent = Nothing
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetLatestVersionDetails
' Author    : Adam Waller
' Date      : 1/27/2017
' Purpose   : Return an array of the latest version details.
'---------------------------------------------------------------------------------------
'
Private Function GetLatestVersionDetails(strName As String) As Variant

    Dim varItem As Variant
    
    If mVersions Is Nothing Then LoadVersionList
    
    For Each varItem In mVersions
        If UBound(varItem) > 2 Then
            If varItem(0) = strName Then
                GetLatestVersionDetails = varItem
                Exit Function
            End If
        End If
    Next varItem

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetCurrentRefVersion
' Author    : Adam Waller
' Date      : 1/31/2017
' Purpose   : Return the version of the currently installed reference.
'---------------------------------------------------------------------------------------
'
Private Function GetCurrentRefVersion(ref As Access.Reference) As String

    Dim wrk As Workspace
    Dim dbs As Database
    
    Set wrk = DBEngine(0)
    Set dbs = wrk.OpenDatabase(ref.FullPath, , True)
    
    ' Attempt to read custom property
    On Error Resume Next
    GetCurrentRefVersion = dbs.Properties("AppVersion")
    If Err Then Err.Clear
    On Error GoTo 0
    
    dbs.Close
    DoEvents
    Set dbs = Nothing
    Set wrk = Nothing
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : SaveVersionList
' Author    : Adam Waller
' Date      : 1/27/2017
' Purpose   : Write the version list to a file.
'---------------------------------------------------------------------------------------
'
Private Function SaveVersionList()
    
    Dim strFile As String
    Dim intFile As Integer
    Dim varLine As Variant
    
    If mVersions Is Nothing Then
        MsgBox "Please load version list before saving", vbExclamation
        Exit Function
    End If
    
    strFile = GetDeploymentFolder & "Latest Versions.csv"
    intFile = FreeFile
    
    ' Read entries in the file
    Open strFile For Output As #intFile
        For Each varLine In mVersions
            ' Write in CSV format
            Print #intFile, Join(varLine, ",")
        Next varLine
    Close intFile

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetFileNameFromPath
' Author    : http://stackoverflow.com/questions/1743328/how-to-extract-file-name-from-path
' Date      : 1/31/2017
' Purpose   : Return file name from path.
'---------------------------------------------------------------------------------------
'
Private Function GetFileNameFromPath(strFullPath As String) As String
    GetFileNameFromPath = Right(strFullPath, Len(strFullPath) - InStrRev(strFullPath, "\"))
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetVBProjectForCurrentDB
' Author    : Adam Waller
' Date      : 7/25/2017
' Purpose   : Get the actual VBE project for the current top-level database.
'           : (This is harder than you would think!)
'---------------------------------------------------------------------------------------
'
Public Function GetVBProjectForCurrentDB() As Object   ' As VBProject

    Dim objProj As Object
    Dim strPath As String
    
    strPath = CurrentProject.FullName
    If VBE.ActiveVBProject.FileName = strPath Then
        ' Use currently active project
        Set GetVBProjectForCurrentDB = VBE.ActiveVBProject
    Else
        ' Search for project with matching filename.
        For Each objProj In VBE.VBProjects
            If objProj.FileName = strPath Then
                Set GetVBProjectForCurrentDB = objProj
                Exit For
            End If
        Next objProj
    End If
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetCodeLines
' Author    : Adam Waller
' Date      : 2/14/2017
' Purpose   : A more robust way of counting the lines of code in a module.
'           : (Simply using LineCount can give varying results, due to white
'           :  spacing differences at the end of a code module.)
'---------------------------------------------------------------------------------------
'
Private Function GetCodeLineCount(objCodeModule As Object) As Long
    
    Dim lngLine As Long
    Dim lngLen As Long
    Dim strLine As String
    
    lngLen = objCodeModule.CountOfLines
    
    ' Start from the last line, and work backwards till we find
    ' the last line of code or comment in the module.
    For lngLine = lngLen To 1 Step -1
        ' Remove line break characters
        strLine = Replace(objCodeModule.Lines(lngLine, lngLine), vbCrLf, "")
        If Trim(strLine) <> "" Then
            ' Found code or comment
            GetCodeLineCount = lngLine
            Exit For
        End If
    Next lngLine
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : AllowAutoUpdate
' Author    : Adam Waller
' Date      : 4/3/2020
' Purpose   : Return true if the module explicitly allows auto-updating, or if the
'           : flag is not present in the module. (Allows legacy modules to be
'           : automatically migrated to auto-updating modules.)
'---------------------------------------------------------------------------------------
'
Private Function AllowAutoUpdate(objCodeModule) As Boolean

    Const cstrFlag As String = "'@AutoUpdate"
    
    Dim blnManualUpdate As Boolean
    Dim varLines As Variant
    Dim intLineCount As Integer
    Dim intLine As Integer
    
    ' We should find this flag within the declaration lines at the top of the module.
    With objCodeModule
        intLineCount = .CountOfDeclarationLines
        varLines = Split(.Lines(1, intLineCount), vbCrLf)
    End With
    
    ' Loop through lines
    For intLine = 1 To intLineCount - 1
        If UCase(Left(varLines(intLine), Len(cstrFlag))) = UCase(cstrFlag) Then
            ' Found the flag line. Let's see if there is a false value on the same line.
            blnManualUpdate = (InStr(1, Replace(varLines(intLine), " ", ""), _
                "=False", vbTextCompare) > Len(cstrFlag))
            ' No need to keep looking, now that we found the flag.
            Exit For
        End If
    Next intLine
    
    ' Return true to allow update unless we found a flag with the value of False
    AllowAutoUpdate = Not blnManualUpdate
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : IsClickOnce
' Author    : Adam Waller
' Date      : 2/14/2017
' Purpose   : Returns true if this application should be deployed as a ClickOnce
'           : application. (Stored as custom property rather than module constant
'           : to make updates easier.)
'---------------------------------------------------------------------------------------
'
Private Function IsClickOnce() As Boolean

    Const cstrName As String = "ClickOnce Deployment"
    Dim prp As Object   ' Access.AccessObjectProperty
    Dim strValue As String
    Dim prpAccdb As Object
    Dim dbs As Database
    
    For Each prp In PropertyParent.Properties
        If prp.Name = cstrName Then
            strValue = prp.Value
            Exit For
        End If
    Next prp
    
    Select Case strValue
        
        Case "True", "False"
            ' Use defined value
        
        Case Else
        
            ' Ask user to define preference
            If Eval("MsgBox('Use ClickOnce Deployment for this application?@Select ''Yes'' to create an application " & _
                "that will be installed on the user''s computer, or click ''No'' to simply update the version number.@" & _
                "(Library databases that are only used as a part of other applications are typically not deployed as ClickOnce installers.)@',36)") = vbYes Then
                strValue = "True"
            Else
                strValue = "False"
            End If
            
            ' Save to this database
            If CodeProject.ProjectType = acADP Then
                PropertyParent.Properties.Add cstrName, strValue
            Else
                ' Normal accdb database property
                Set dbs = CurrentDb
                Set prpAccdb = dbs.CreateProperty(cstrName, DB_TEXT, strValue)
                dbs.Properties.Append prpAccdb
                Set dbs = Nothing
            End If
            
    End Select
    
    Set prp = Nothing
    Set prpAccdb = Nothing
    
    ' Return the existing or newly defined value.
    IsClickOnce = CBool(strValue)
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : LocalizeReferences
' Author    : Adam Waller
' Date      : 2/22/2017
' Purpose   : Make sure references are local
'---------------------------------------------------------------------------------------
'
Public Sub LocalizeReferences()

    Dim oApp As Access.Application
    'Set oApp = New Access.Application
    Set oApp = CreateObject("Access.Application")
    
    With oApp
        .UserControl = True ' Turn visible and stay open.
        .OpenCurrentDatabase GetDeploymentFolder & "_Tools\Localize References.accdb"
        .Eval "LocalizeReferencesForRemoteDB(""" & CurrentProject.FullName & """)"
    End With
    
    Set oApp = Nothing
    Application.Quit acQuitSaveAll
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : HasDuplicateProjects
' Author    : Adam Waller
' Date      : 2/22/2017
' Purpose   : Returns true if duplicate projects exist with the same name.
'           : (Typically caused by non-localized references.)
'---------------------------------------------------------------------------------------
'
Private Function HasDuplicateProjects() As Boolean
    
    Dim colProjects As Collection
    Dim objProj As Object
    Dim strName As String
    Dim varProj As Variant
    
    Set colProjects = New Collection
    
    For Each objProj In VBE.VBProjects
        strName = objProj.Name
        
        ' See if we have already seen this project name.
        For Each varProj In colProjects
            If strName = varProj Then
                HasDuplicateProjects = True
                Exit For
            End If
        Next varProj
        If HasDuplicateProjects Then Exit For
        
        ' Add to list of project names
        colProjects.Add strName
    Next objProj
    
    Set objProj = Nothing
    Set colProjects = Nothing
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : PrintDebugMsg
' Author    : Adam Waller
' Date      : 2/22/2017
' Purpose   : Print a debug message to the immediate window.
'           : Used in COM automation processes such as LocalizeReferences.
'---------------------------------------------------------------------------------------
'
Public Function PrintDebugMsg(strMsg) As String
    Debug.Print strMsg
End Function


'---------------------------------------------------------------------------------------
' Procedure : UpdateDeployModule
' Author    : Adam Waller
' Date      : 4/3/2020
' Purpose   : Uses a couple little tricks to effectively replace the running module.
'---------------------------------------------------------------------------------------
'
Private Sub UpdateDeployModule()

    Const vbext_ct_StdModule = 1
    Dim proj As Object  ' VBProject
    Set proj = GetVBProjectForCurrentDB
    
    ' Rename this module as a backup
    proj.VBComponents("basDeploy").Name = "basDeployBak"
    
    ' Import the basDeploy module
    proj.VBComponents.Import(GetDeploymentFolder & "\Code Templates\basDeploy.bas").Name = "basDeploy"
    
    ' Now, call a function in the new module to remove this (now) backup module.
    basDeploy.RemoveBackupDeploymentModule
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : RemoveBackupDeploymentModule
' Author    : Adam Waller
' Date      : 4/3/2020
' Purpose   : Removes the backup deployment module created when updating basDeploy.
'---------------------------------------------------------------------------------------
'
Public Sub RemoveBackupDeploymentModule()

    Dim proj As Object  '  VBProject
    Set proj = GetVBProjectForCurrentDB
    
    ' Remove the backup module
    With proj.VBComponents
        .Remove .Item("basDeployBak")
    End With
    
End Sub


