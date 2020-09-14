Option Strict Off
Option Explicit On
Imports System.IO

Module modLoadDefaults
    Public m_DBPath As String
    Public m_DataPath As String
    Public m_PPTPath As String
    Public m_PPTView As String
    Public m_Office_Dir As String
    Public m_Office_App_Dir As String
    Public m_AppDirPath = My.Computer.FileSystem.CurrentDirectory.ToString
    Public m_DefPrgFiles As String = My.Computer.FileSystem.SpecialDirectories.ProgramFiles.ToString & "\"
    'Public m_AppDirPath As String = My.Application.Info.DirectoryPath
    Public Function m_LoadDefaults() As Boolean
        'Declare directory path as var
        ' Declare local variable sDefDBPath
        Dim sDefDBpath As String

        ' set the default location of database file
        sDefDBpath = m_AppDirPath & "\DB\songs.mdb"
        Debug.Print("sDefDBPath :" & sDefDBpath)
        ' Declare local variable for sDefDataPath
        Dim sDefDataPath As String
        ' set default location for the data files
        sDefDataPath = m_AppDirPath & "\data\"
        ' Declare local variable for Power Point Viewer Path
        Dim sDefPPTViewPath As String
        ' set Power Point Viewer Path to Empty String
        sDefPPTViewPath = ""
        ' Declare Power Point Path Variable
        Dim sDefPPTPath As String
        ' set power point path to empty string
        sDefPPTPath = ""
        ' declare MyPath Variable
        Dim MyPath As String
        ' Declare 64bitprgfiles variable

        ' set bit program files directory
        Debug.Print("m_DefPrgFiles " & m_DefPrgFiles)
        'CreateObject for directory exists
        On Error GoTo RaiseError
        If Not My.Computer.FileSystem.FileExists(sDefDBpath) And Not My.Computer.FileSystem.DirectoryExists(sDefDataPath) Then
            Return False
            Exit Function
        End If
        If My.Computer.FileSystem.DirectoryExists(m_DefPrgFiles & "microsoft office\") Then
            m_Office_Dir = m_DefPrgFiles & "Microsoft Office\"
            Debug.Print(m_Office_Dir)
            'LocateOfficeDir()
            If My.Computer.FileSystem.DirectoryExists(m_Office_Dir) Then
                MyPath = m_Office_Dir
                If My.Computer.FileSystem.FileExists(MyPath & "PowerPNT.exe") Then
                    sDefPPTPath = MyPath & "PowerPNT.exe"
                End If
                If My.Computer.FileSystem.FileExists(MyPath & "PPTVIEW.EXE") <> "" Then
                    sDefPPTViewPath = MyPath & "PPTVIEW.EXE"
                End If
            End If
            m_DBPath = sDefDBpath
            m_DataPath = sDefDataPath
            m_PPTPath = sDefPPTPath
            m_PPTView = sDefPPTViewPath
        End If
        Return True
RaiseError:
        MsgBox(Err.Description, MsgBoxStyle.Exclamation)
        Return False
        Exit Function

    End Function

    Public Sub LocateOfficeDir()
        Dim OfficeDirFound As Boolean
        Dim MyPath As String
        MyPath = m_Office_Dir & "Office\"
        ' Set value of office dir to false
        OfficeDirFound = False
        Dim counter As Short
        counter = 9
        Do Until OfficeDirFound = True
            If My.Computer.FileSystem.DirectoryExists(MyPath) Then
                OfficeDirFound = True
                m_Office_App_Dir = MyPath
                Debug.Print("Directory Found: " & MyPath)
            Else

                OfficeDirFound = False

                If counter = 17 Then Exit Do

                counter = counter + 1

                MyPath = m_Office_Dir & "office" & CStr(LTrim(CStr(counter))) & "\"
                Debug.Print(MyPath)
            End If
        Loop

    End Sub

    Public Sub SaveRegistry()

        
        SaveSetting("Church Songs", "Default", "AppDirPath", m_AppDirPath)
        SaveSetting("Church Songs", "Default", "DataPath", m_DataPath)
        SaveSetting("Church Songs", "Default", "DBPath", m_DBPath)
        SaveSetting("Church Songs", "Default", "PPTPath", m_PPTPath)
        SaveSetting("Church Songs", "Default", "PPTView", m_PPTView)
        SaveSetting("Church Songs", "Default", "OfficeDir", m_Office_Dir)
        SaveSetting("Church Songs", "Default", "OfficeAppDir", m_Office_App_Dir)
        SaveSetting("Church Songs", "Default", "DefPrgFiles", m_DefPrgFiles)
    End Sub
    Public Sub GetRegistry()
        m_AppDirPath = GetSetting("Church Songs", "Default", "AppDirPath","")
        m_DataPath = GetSetting("Church Songs", "Default", "DataPath", "")
        m_DBPath = GetSetting("Church Songs", "Default", "DBPath", "")
        m_PPTPath = GetSetting("Church Songs", "Default", "PPTPath", "")
        m_PPTView = GetSetting("Church Songs", "Default", "PPTView", "")
        m_Office_Dir = GetSetting("Church Songs", "Default", "OfficeDir", "")
        m_Office_App_Dir = GetSetting("Church Songs", "Default", "OfficeAppDir", "")
        m_DefPrgFiles = GetSetting("Church Songs", "Default", "DefPrgFiles", "")
    End Sub
End Module
m_AppDirPath