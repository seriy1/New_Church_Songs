Option Explicit On
Option Strict Off
Module modGetMyData
    Public Property DAODBEngine_definst As Object
    'Public Property datTodayTmp As Object
    Public Property dbfsongs As Object

    Public Sub GetData()
        'UPGRADE_WARNING: Arrays in structure recSelect may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
        Dim dbTemp As dao.Database
        Dim recSelect As dao.Recordset
        Dim strSQL As String
        Dim myDBPath As String
        If m_DBPath = "" Then Exit Sub
        On Error GoTo RaiseError

        myDBPath = m_DBPath

        ' Get the database name (Path) and open database.
        'UPGRADE_ISSUE: Data property datTodayTmp.DatabaseName was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
        'datTodayTmp.Name = myDBPath
        'datTodayTmp.Refresh()

        dbTemp = DAODBEngine_definst.Workspaces(0).OpenDatabase(myDBPath)
        ' Open a snapshot database.
        strSQL = "Select [SongNum], [SongName]from [todaytemp] " & "ORDER BY [SongName]"

        ' Create the RecordSet
        recSelect = dbTemp.OpenRecordset(strSQL, dao.RecordsetTypeEnum.dbOpenSnapshot)

        ' Check and see if database has any records and then continue loading data
        If recSelect.RecordCount > 0 Then
            recSelect.MoveFirst()
            Do While Not recSelect.EOF
                ' lstToday.Items.Add(New VB6.ListBoxItem(recSelect.Fields("SongName").Value, Val(recSelect.Fields("SongNum").Value)))
                recSelect.MoveNext()
            Loop
        End If
        'UPGRADE_ISSUE: Data property datTodayTmp.DatabaseName was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
        'datTodayTmp.Name = ""
        'UPGRADE_ISSUE: Data property datTodayTmp.RecordSource was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
        'datTodayTmp.RecordSource = ""
RaiseError:
        MsgBox(Err.Description, MsgBoxStyle.Exclamation)
        Exit Sub
    End Sub

    Public Sub MoveBack(ByRef data As String)
        Dim tdftable As dao.TableDef
        Dim blnTableFound1 As Boolean
        Dim blnTableFound2 As Boolean
        Dim dbpath As String
        Dim strSQL As String

        On Error GoTo RaiseError
        dbpath = m_DBPath

        blnTableFound1 = False
        blnTableFound2 = False
        ' check and see if temp table
        ' exist
        dbfsongs = DAODBEngine_definst.Workspaces(0).OpenDatabase(dbpath)
        ' for each statement determines if tables exists
        For Each tdftable In dbfsongs.TableDefs
            If tdftable.Name = "temp" Then
                blnTableFound1 = True
                Exit For
            End If
        Next tdftable
        ' check and see if todaytemp table exist.
        For Each tdftable In dbfsongs.TableDefs
            If tdftable.Name = "todaytemp" Then
                blnTableFound2 = True
                Exit For
            End If
        Next tdftable
        If blnTableFound1 = True And blnTableFound2 = True Then
            ' build sql command
            ' This string for INSERTING the data to temp db table from
            ' todaytemp table

            strSQL = "INSERT INTO [temp] ([SongNum], [SongName], [SongComplete], [Content]) " & "SELECT todaytemp.SongNum, todaytemp.SongName, todaytemp.SongComplete, todaytemp.Content " & "FROM todaytemp " & "WHERE [SongName] = """ & data & """"
            ' execute command
            dbfsongs.Execute((strSQL))

            ' build sql command to remove data
            ' from todaytemp table.
            strSQL = "DELETE FROM [todaytemp] " & "WHERE [todaytemp].[SongName] = " & """" & data & """"
            ' execute the command
            dbfsongs.Execute((strSQL))
        End If
        Exit Sub
RaiseError:
        MsgBox(Err.Description, MsgBoxStyle.Exclamation)
        Exit Sub
    End Sub
End Module
