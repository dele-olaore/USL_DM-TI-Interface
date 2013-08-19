VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DMInterface2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim S_OK As Integer
Dim ErrorMessage As String
Dim ErrorCode As Long
Dim RegCode As Integer

Dim UserId As String, Password As String, Library As String, TrialCount As Integer, WebTopUrl As String
Dim Profile As String, Query As String, PassordEncrypted As String
Dim DST As String, Group As String

Public Type MyHash
    key As String
    value As String
End Type

Dim transidValues(0 To 1000) As MyHash 'max of 1000 session at a time
Dim transid_ As String

Dim FSO As New FileSystemObject
Dim TS As TextStream
Dim logFile As String

Enum AttachStatus
    NotComplete = 0
    Complete = 1
    Error1 = 2
End Enum

Private Sub LoadRegValues()
    On Error GoTo RegFail
    ' load the registry values
    
    Dim sKey As String '***Key under which to create the value
    Dim sValue As String '***Value name to set
    Dim vSetting As Variant

    sKey = "Software\TIPlus_DMInterface\DMParameters"
    
    sValue = "UserId"
    '***Read the String value
    vSetting = QueryValue(HKEY_CURRENT_USER, sKey, sValue)
    UserId = "" & vSetting
    
    sValue = "Password"
    '***Read the String value
    vSetting = QueryValue(HKEY_CURRENT_USER, sKey, sValue)
    Password = "" & vSetting
    
    sValue = "Library"
    '***Read the String value
    vSetting = QueryValue(HKEY_CURRENT_USER, sKey, sValue)
    Library = "" & vSetting
    
    sValue = "TrialCount"
    '***Read the String value
    vSetting = QueryValue(HKEY_CURRENT_USER, sKey, sValue)
    TrialCount = CInt("" & vSetting)
    
    sValue = "WebTopUrl"
    '***Read the String value
    vSetting = QueryValue(HKEY_CURRENT_USER, sKey, sValue)
    WebTopUrl = "" & vSetting
    
    sKey = "Software\TIPlus_DMInterface\DocForms"
    
    sValue = "Profile"
    '***Read the String value
    vSetting = QueryValue(HKEY_CURRENT_USER, sKey, sValue)
    Profile = "" & vSetting
    
    sValue = "Query"
    '***Read the String value
    vSetting = QueryValue(HKEY_CURRENT_USER, sKey, sValue)
    Query = "" & vSetting
    
    sKey = "Software\TIPlus_DMInterface\OtherParams"
    
    sValue = "PassordEncrypted"
    '***Read the String value
    vSetting = QueryValue(HKEY_CURRENT_USER, sKey, sValue)
    PassordEncrypted = "" & vSetting
    
    RegCode = 1
    log "Registry values read successfully!"
    
    Exit Sub
    
RegFail:
    RegCode = 0
    log "Error reading registry values. Source: " & Err.Source & ", Number: " & Err.Number & ", Message: " & Err.description
End Sub

Public Sub Class_Initialize()
    RegCode = 0
    logFile = "c:\dminterface-log.txt"
    
    S_OK = 1
    
    DST = ""
    Group = ""
End Sub

Public Sub Class_Terminate()
    On Error Resume Next
    
    'Set TS = Nothing
    'Set FSO = Nothing
End Sub

Public Sub log(ByVal message As String)
    On Error GoTo closeStream
    
    Dim TempS As String
    Dim Final As String
    
    Set TS = FSO.OpenTextFile(logFile, ForReading, True)
    Do Until TS.AtEndOfStream
        TempS = TS.ReadLine
        Final = Final & TempS & vbCrLf
    Loop
    TS.Close
    
    Final = Final & "log at " & Now() & " - " & message & vbCrLf
    
    Set TS = FSO.OpenTextFile(logFile, ForWriting, True)
    TS.Write Final
    TS.Close
    
    Exit Sub
    
closeStream:
    On Error Resume Next
    TS.Close
    Set TS = Nothing
End Sub

Private Sub AddToKeyList(ByVal MisysSessionId As String, ByVal trnsid_ As String)
    Dim curKey As MyHash
    curKey.key = MisysSessionId
    curKey.value = trnsid_
    
    Dim exists As Boolean
    Dim pos As Long
    Dim lngPosition As Long 'for iterating through the array
    
    exists = False
    For lngPosition = LBound(transidValues) To UBound(transidValues)
        Dim existKey As MyHash
        existKey = transidValues(lngPosition)
        If Not (IsNull(existKey)) Then
            If existKey.key = curKey.key Then
                exists = True
                pos = lngPosition
                Exit For
            End If
        Else
            Exit For
        End If
    Next lngPosition
    
    If Not exists Then
        'ReDim Preserve transidValues(0 To UBound(transidValues) + 1) As MyHash
        'Set transidValues(UBound(transidValues)) = curKey
        transidValues(lngPosition) = curKey
    Else
        transidValues(pos) = curKey
    End If
End Sub

Private Function GetKeyValue(ByVal MisysSessionId As String) As String
    Dim value As String
    Dim lngPosition As Long 'for iterating through the array
    For lngPosition = LBound(transidValues) To UBound(transidValues)
        Dim existKey As MyHash
        existKey = transidValues(lngPosition)
        If Not (IsNull(existKey)) Then
            If existKey.key = MisysSessionId Then
                value = existKey.value
                Exit For
            End If
        End If
    Next lngPosition
    
    GetKeyValue = value
    
End Function

Private Sub RemoveKey(ByVal MisysSessionId As String)
    Dim lngPosition As Long 'for iterating through the array
    For lngPosition = LBound(transidValues) To UBound(transidValues)
        Dim existKey As MyHash
        existKey = transidValues(lngPosition)
        If Not (IsNull(existKey)) Then
            If existKey.key = MisysSessionId Then
                transidValues(lngPosition) = Null
                Exit For
            End If
        End If
    Next lngPosition
End Sub

Public Function View(Optional ByVal docId As String, Optional ByVal windowTitle As String, Optional ByVal MisysSessionId As String, Optional ByRef WebPageUrl As String, Optional ByRef success As Boolean) As Integer
    On Error GoTo LogError
    S_OK = 1
    
    success = 0
    log "TI calling View Function"
    
    If Len(docId) = 0 Then
        ErrorCode = 800
        ErrorMessage = "Document Id must be Specifed"
        log "Error: 800, Description: Document Id must be Specifed"
    Else
        If RegCode = 0 Then
            LoadRegValues
        End If
        
        If RegCode = 1 Then
            Dim webstr As String
            'webstr = "http://10.1.11.12/cyberdocs"
            webstr = RTrim(LTrim(WebTopUrl))
            Dim docidstr As String
            docidstr = RTrim(LTrim(docId))
            Dim userstr As String
            'userstr = "dmadmin"
            userstr = RTrim(LTrim(UserId))
            Dim passstr As String
            'passstr = "10loopDS$"
            passstr = RTrim(LTrim(Password))
            Dim libstr As String
            'libstr = "DMSLIV"
            libstr = RTrim(LTrim(Library))
            
            Dim url As String
            url = webstr & "/quickstart.asp?show=VIEW:" & docidstr & ":native&library=" & libstr & "&noframes=yes&eclip=yes&papi=bypass&user=" & userstr & ":" & passstr
            
            WebPageUrl = RTrim(LTrim(url))
            
            success = 1
            log "View returned WebPageUrl: " & WebPageUrl
        Else
            ErrorCode = 100
            ErrorMessage = "Registry Parameters not Fetched Properly"
            log "Error: 100, Description: Registry Parameters not Fetched Properly"
        End If
    End If
    
    log "View Function returned"
    View = S_OK
    
    Exit Function
    
LogError:
    log "Error: Source: " & Err.Source & ", Number: " & Err.Number & ", Description: " & Err.description
    View = S_OK
End Function

Public Function IsAvailable(Optional ByRef success As Boolean) As Integer
    On Error GoTo LogError
    S_OK = 1
    success = 0
    log "TI calling IsAvailable Function"

    success = ConnectToDM()

    IsAvailable = S_OK
    
    Exit Function
    
LogError:
    log "Error: Source: " & Err.Source & ", Number: " & Err.Number & ", Description: " & Err.description
    IsAvailable = S_OK
End Function

Public Function AttachTo(Optional ByVal name As String, Optional ByVal description As String, Optional ByVal windowTitle As String, Optional ByVal MisysSessionId As String, Optional ByRef WebPageUrl As String, Optional ByRef success As Boolean) As Integer
    On Error GoTo LogError
    S_OK = 1
    success = 0
    log "TI calling AttachTo Function"
    
    If RegCode = 0 Then
        LoadRegValues
    End If
    
    If RegCode = 1 Then
        
        Dim webstr As String
        'webstr = "http://10.1.11.12/cyberdocs"
        webstr = RTrim(LTrim(WebTopUrl))
        Dim userstr As String
        'userstr = "dmadmin"
        userstr = RTrim(LTrim(UserId))
        Dim passstr As String
        'passstr = "10loopDS$"
        passstr = RTrim(LTrim(Password))
        Dim libstr As String
        'libstr = "DMSLIV"
        libstr = RTrim(LTrim(Library))
        Dim qbestr As String
        qbestr = RTrim(LTrim(Query)) '"TI_PROF"
        
        Dim tistr As String
        tistr = MisysSessionId
        
        transid_ = Str(Now())
        
        'AddToKeyList MisysSessionId, transid_
        
        Dim namestr As String
        namestr = ""
        Dim descstr As String
        descstr = ""
    
        Dim url As String
    
        url = webstr & "/quickstart.asp?show=SEARCHRESULTSACT:TI_ATTACHED=N:_FORMNAME=" & qbestr & "&library=" & libstr & "&eclip=yes&papi=bypass&user=" & userstr & ":" & passstr & "&tiname=" & namestr & "&tidescription=" & descstr & "&TIPlusSessionID=" & tistr & "&TRANSID=" & tistr 'transid_
    
        WebPageUrl = RTrim(LTrim(url))
        success = 1
        log "AttachTo returned WebPageUrl: " & url
    Else
        ErrorCode = 100
        ErrorMessage = "Registry Parameters not Fetched Properly"
        log "Error: 100, Description: Registry Parameters not Fetched Properly"
    End If
    
    log "AttachTo Function returned"
    AttachTo = S_OK
    
    Exit Function
    
LogError:
    log "Error: Source: " & Err.Source & ", Number: " & Err.Number & ", Description: " & Err.description
    AttachTo = S_OK
End Function

Public Function EnhancedAttachTo(Optional ByVal name As String, Optional ByVal description As String, Optional ByVal windowTitle As String, Optional ByVal MisysSessionId As String, Optional ByVal XmlSearchCriteria As String, Optional ByRef WebPageUrl As String, Optional ByRef success As Boolean) As Integer
    On Error GoTo LogError
    S_OK = 1
    success = 0
    log "TI calling EnhancedAttachTo Function"

    If RegCode = 0 Then
        LoadRegValues
    End If
        
    If RegCode = 1 Then
    
        'If Len(name) = 0 Or Len(description) = 0 Then
        '    ErrorCode = 800
        '    ErrorMessage = "Name and Description are required."
        '    log "Error: Number: " & ErrorCode & ", Description: " & ErrorMessage
        'Else
            Dim webstr As String
            'webstr = "http://10.1.11.12/cyberdocs"
            webstr = RTrim(LTrim(WebTopUrl))
            Dim userstr As String
            'userstr = "dmadmin"
            userstr = RTrim(LTrim(UserId))
            Dim passstr As String
            'passstr = "10loopDS$"
            passstr = RTrim(LTrim(Password))
            Dim libstr As String
            'libstr = "DMSLIV"
            libstr = RTrim(LTrim(Library))
            Dim qbestr As String
            qbestr = RTrim(LTrim(Query)) '"TI_PROF"
            
            Dim tistr As String
            tistr = MisysSessionId
            
            transid_ = Str(Now())
            
            'AddToKeyList MisysSessionId, transid_
            
            Dim namestr As String
            namestr = ""
            Dim descstr As String
            descstr = ""
            
            Dim url As String
            
            url = webstr & "/quickstart.asp?show=SEARCHRESULTSACT:TI_ATTACHED=N:_FORMNAME=" & qbestr & "&library=" & libstr & "&eclip=yes&papi=bypass&user=" & userstr & ":" & passstr & "&tiname=" & namestr & "&tidescription=" & descstr & "&TIPlusSessionID=" & tistr & "&TRANSID=" & tistr 'transid_
            
            WebPageUrl = RTrim(LTrim(url))
            success = 1
            log "EnhancedAttachTo returned WebPageUrl: " & url
        'End If
    Else
        ErrorCode = 100
        ErrorMessage = "Registry Parameters not Fetched Properly"
        log "Error: 100, Description: Registry Parameters not Fetched Properly"
    End If

    log "EnhancedAttachTo Function returned"
    EnhancedAttachTo = S_OK
    
    Exit Function
    
LogError:
    log "Error: Source: " & Err.Source & ", Number: " & Err.Number & ", Description: " & Err.description
    EnhancedAttachTo = S_OK
End Function

Public Function AttachComplete(Optional ByVal MisysSessionId As String, Optional ByRef docId As String, Optional ByRef status As AttachStatus, Optional ByRef success As Boolean) As Integer
    On Error GoTo LogError
    S_OK = 1
    status = NotComplete
    success = 0
    log "TI calling AttachComplete Function"
    
    If Len(DST) = 0 Then
        ConnectToDM
    End If
    
    If Len(DST) > 0 Then
        Dim rec As New PCDSearch
        Dim libstr As String
        libstr = RTrim(LTrim(Library))
        Dim qstr As String
        qstr = RTrim(LTrim(Query))
        
        'Dim trnsid As String
        'trnsid = GetKeyValue(MisysSessionId)
        
        Dim id As String
        id = "" & MisysSessionId & MisysSessionId 'trnsid 'transid_
        
        rec.SetDST DST
        rec.AddSearchLib libstr
        rec.SetSearchObject qstr
        rec.AddReturnProperty "DOCNUMBER"
        rec.AddSearchCriteria "TI_REF1", id
        
        rec.Execute
        If (rec.ErrNumber <> 0) Then
            ErrorCode = rec.ErrNumber
            ErrorMessage = rec.ErrDescription
            log "Error searching for document in Attach Complete from DM: Number: " & rec.ErrNumber & ", Description: " & rec.ErrDescription
        Else
            Dim Count As Integer
            Count = rec.GetRowsFound
            If Count > 0 Then
                If Count = 1 Then
                    If rec.NextRow Then
                        Dim DOCNUMBER As String
                        DOCNUMBER = "" & rec.GetPropertyValue("DOCNUMBER")
                        
                        docId = DOCNUMBER
                        status = Complete
                        success = 1
                        
                        RemoveKey MisysSessionId
                        
                        log "Attach complete for DOC: " & DOCNUMBER
                    End If
                Else
                    status = Error1
                    ErrorCode = 990
                    ErrorMessage = "No or Multiple Row Returned  for same DocumentID"
                    log "Error searching for document in Attach Complete from DM: Number: " & ErrorCode & ", Description: " & ErrorMessage
                End If
            End If
        End If
        
        rec.ReleaseResults
    Else
        status = Error1
        success = 0
        log "Error: Could not connect to DM from AttachComplete."
    End If
    
    log "AttachComplete Function returned"
    AttachComplete = S_OK
    
    Exit Function
    
LogError:
    log "Error: Source: " & Err.Source & ", Number: " & Err.Number & ", Description: " & Err.description
    AttachComplete = S_OK
End Function

Public Function AttachTimeout(Optional ByVal MisysSessionId As String, Optional ByRef success As Boolean) As Integer
    On Error GoTo LogError
    S_OK = 1
    success = 1
    log "TI calling AttachTimeout Function"

    log "AttachTimeout Function returned"
    AttachTimeout = S_OK
    
    Exit Function
    
LogError:
    log "Error: Source: " & Err.Source & ", Number: " & Err.Number & ", Description: " & Err.description
    AttachTimeout = S_OK
End Function

Public Function Delete(Optional ByVal docId As String, Optional ByRef success As Boolean) As Integer
    On Error GoTo LogError
    S_OK = 1
    success = 0
    log "TI calling Delete Function"

    If Len(DST) = 0 Then
        ConnectToDM
    End If
    
    If Len(DST) > 0 Then
        Dim docobj As New PCDDocObject
        Dim libstr As String
        libstr = RTrim(LTrim(Library))
        Dim pstr As String
        pstr = RTrim(LTrim(Profile))
        
        docobj.SetDST DST
        docobj.SetObjectType pstr
        docobj.SetProperty "%TARGET_LIBRARY", libstr
        docobj.SetProperty "%OBJECT_IDENTIFIER", docId
        
        docobj.Delete
        
        ErrorCode = docobj.ErrNumber
        ErrorMessage = docobj.ErrDescription
        
        If ErrorCode = 0 Then
            success = 1
            log "Delete successful"
        Else
            ErrorCode = 997
            ErrorMessage = "Failed to Delete Record"
            log "Error deleting document from Delete in DM: Number: " & docobj.ErrNumber & ", Description: " & docobj.ErrDescription
        End If
    Else
        success = 0
        log "Error: Could not connect to DM from Delete."
    End If

    log "Delete Function returned"
    Delete = S_OK
    
    Exit Function
    
LogError:
    log "Error: Source: " & Err.Source & ", Number: " & Err.Number & ", Description: " & Err.description
    Delete = S_OK
End Function

Public Function Detach(Optional ByVal docId As String, Optional ByRef success As Boolean) As Integer
    On Error GoTo LogError
    S_OK = 1
    success = 0
    log "TI calling Detach Function"

    If Len(DST) = 0 Then
        ConnectToDM
    End If
    
    If Len(DST) > 0 Then
        Dim docobj As New PCDDocObject
        Dim libstr As String
        libstr = RTrim(LTrim(Library))
        Dim pstr As String
        pstr = RTrim(LTrim(Profile))
        
        docobj.SetDST DST
        docobj.SetObjectType pstr
        docobj.SetProperty "%TARGET_LIBRARY", libstr
        docobj.SetProperty "%OBJECT_IDENTIFIER", docId
        docobj.SetProperty "TI_ATTACHED", "N"
        
        docobj.Update
        
        ErrorCode = docobj.ErrNumber
        ErrorMessage = docobj.ErrDescription
        
        If ErrorCode = 0 Then
            success = 1
            log "Detach successful"
        Else
            ErrorCode = 997
            ErrorMessage = "Failed to Detach Record"
            log "Error detaching document from Detach in DM: Number: " & docobj.ErrNumber & ", Description: " & docobj.ErrDescription
        End If
    Else
        success = 0
        log "Error: Could not connect to DM from Detach."
    End If

    log "Detach Function returned"
    Detach = S_OK
    
    Exit Function
    
LogError:
    log "Error: Source: " & Err.Source & ", Number: " & Err.Number & ", Description: " & Err.description
    Detach = S_OK
End Function

Public Function SetName(Optional ByVal docId As String, Optional ByVal name As String, Optional ByRef success As Boolean) As Integer
    On Error GoTo LogError
    S_OK = 1
    success = 0
    log "TI calling SetName Function"

    If Len(DST) = 0 Then
        ConnectToDM
    End If
    
    If Len(docId) > 0 Then
        If Len(DST) > 0 Then
            Dim docobj As New PCDDocObject
            Dim libstr As String
            libstr = RTrim(LTrim(Library))
            Dim pstr As String
            pstr = RTrim(LTrim(Profile))
            
            docobj.SetDST DST
            docobj.SetObjectType pstr
            docobj.SetProperty "%TARGET_LIBRARY", libstr
            docobj.SetProperty "%OBJECT_IDENTIFIER", docId
            docobj.SetProperty "TI_NAME", name
            
            docobj.Update
            
            ErrorCode = docobj.ErrNumber
            ErrorMessage = docobj.ErrDescription
            
            If ErrorCode = 0 Then
                success = 1
                log "SetName updated doc: " & docId & ", name: " & name
            Else
                ErrorCode = 997
                ErrorMessage = "Failed to Update name"
                log "Error updating document from SetName in DM: Number: " & docobj.ErrNumber & ", Description: " & docobj.ErrDescription
            End If
        Else
            success = 0
            log "Error: Could not connect to DM from SetName."
        End If
    Else
        ErrorCode = 801
        ErrorMessage = "Document Number and Name must be Specifed"
        log "Error: Number: " & ErrorCode & ", Description: " & ErrorMessage
    End If

    log "SetName Function returned"
    SetName = S_OK
    
    Exit Function
    
LogError:
    log "Error: Source: " & Err.Source & ", Number: " & Err.Number & ", Description: " & Err.description
    SetName = S_OK
End Function

Public Function SetDescription(Optional ByVal docId As String, Optional ByVal description As String, Optional ByRef success As Boolean) As Integer
    On Error GoTo LogError
    S_OK = 1
    success = 0
    log "TI calling SetDescription Function"

    If Len(DST) = 0 Then
        ConnectToDM
    End If
    
    If Len(docId) > 0 Then
        If Len(DST) > 0 Then
            Dim docobj As New PCDDocObject
            Dim libstr As String
            libstr = RTrim(LTrim(Library))
            Dim pstr As String
            pstr = RTrim(LTrim(Profile))
            
            docobj.SetDST DST
            docobj.SetObjectType pstr
            docobj.SetProperty "%TARGET_LIBRARY", libstr
            docobj.SetProperty "%OBJECT_IDENTIFIER", docId
            docobj.SetProperty "TI_DESCRIPTION", description
            
            docobj.Update
            
            ErrorCode = docobj.ErrNumber
            ErrorMessage = docobj.ErrDescription
            
            If ErrorCode = 0 Then
                success = 1
                log "SetDescription updated doc: " & docId & ", description: " & description
            Else
                ErrorCode = 997
                ErrorMessage = "Failed to Update description"
                log "Error updating document from SetDescription in DM: Number: " & docobj.ErrNumber & ", Description: " & docobj.ErrDescription
            End If
        Else
            success = 0
            log "Error: Could not connect to DM from SetDescription."
        End If
    Else
        ErrorCode = 801
        ErrorMessage = "Document Number and Description must be Specifed"
        log "Error: Number: " & ErrorCode & ", Description: " & ErrorMessage
    End If

    log "SetDescription Function returned"
    SetDescription = S_OK
    
    Exit Function
    
LogError:
    log "Error: Source: " & Err.Source & ", Number: " & Err.Number & ", Description: " & Err.description
    SetDescription = S_OK
End Function

Public Function GetName(Optional ByVal docId As String, Optional ByRef name As String, Optional ByRef success As Boolean) As Integer
    On Error GoTo errLabel
    S_OK = 1
    success = 0
    log "TI calling GetName Function"

    If Len(DST) = 0 Then
        ConnectToDM
    End If
    
    If Len(docId) > 0 Then
        If Len(DST) > 0 Then
            Dim rec As New PCDSearch
            Dim libstr As String
            libstr = RTrim(LTrim(Library))
            Dim qstr As String
            qstr = RTrim(LTrim(Query))
            
            rec.SetDST DST
            rec.AddSearchLib libstr
            rec.SetSearchObject qstr
            rec.AddReturnProperty "TI_NAME"
            rec.AddSearchCriteria "DOCNUM", docId
            
            rec.Execute
            
            ErrorCode = rec.GetErrNumber
            ErrorMessage = rec.GetErrDescription

            If ErrorCode = 0 Then
                Dim Count As Integer
                Count = rec.GetRowsFound
                If Count = 1 Then
                    If rec.NextRow Then
                        Dim value As String
                        value = "" & rec.GetPropertyValue("TI_NAME")
                        name = value
                        success = 1
                        log "GetName success: DOCNUMBER: " & docId & ", Name: " & name
                    End If
                Else
                    ErrorCode = 990
                    ErrorMessage = "No or Multiple Row Returned for same DocumentID"
                    log "Error getting document name: Number: " & ErrorCode & ", Description: " & ErrorMessage
                End If
            Else
                ErrorCode = 997
                ErrorMessage = "Failed to get name"
                log "Error getting document name from GetName in DM: Number: " & ErrorCode & ", Description: " & ErrorMessage
            End If
            
            rec.ReleaseResults
        Else
            success = 0
            log "Error: Could not connect to DM from GetName."
        End If
    Else
        ErrorCode = 801
        ErrorMessage = "Document Number must be Specifed"
        log "Error: Number: " & ErrorCode & ", Description: " & ErrorMessage
    End If

    log "GetName Function returned"
    GetName = S_OK
    
    Exit Function
errLabel:
    log "Error Source: " & Err.Source & ", Number: " & Err.Number & ", Description: " & Err.description
    GetName = S_OK
End Function

Public Function GetDescription(Optional ByVal docId As String, Optional ByRef description As String, Optional ByRef success As Boolean) As Integer
    On Error GoTo errLabel
    S_OK = 1
    success = 0
    log "TI calling GetDescription Function"

    If Len(DST) = 0 Then
        ConnectToDM
    End If
    
    If Len(docId) > 0 Then
        If Len(DST) > 0 Then
            Dim rec As New PCDSearch
            Dim libstr As String
            libstr = RTrim(LTrim(Library))
            Dim qstr As String
            qstr = RTrim(LTrim(Query))
            
            rec.SetDST DST
            rec.AddSearchLib libstr
            rec.SetSearchObject qstr
            rec.AddReturnProperty "TI_DESCRIPTION"
            rec.AddSearchCriteria "DOCNUM", docId
            
            rec.Execute
            
            ErrorCode = rec.GetErrNumber
            ErrorMessage = rec.GetErrDescription

            If ErrorCode = 0 Then
                Dim Count As Integer
                Count = rec.GetRowsFound
                If Count = 1 Then
                    If rec.NextRow Then
                        Dim value As String
                        value = "" & rec.GetPropertyValue("TI_DESCRIPTION")
                        description = value
                        success = 1
                        log "GetDescription success: DOCNUMBER: " & docId & ", Description: " & description
                    End If
                Else
                    ErrorCode = 990
                    ErrorMessage = "No or Multiple Row Returned for same DocumentID"
                    log "Error getting document description: Number: " & ErrorCode & ", Description: " & ErrorMessage
                End If
            Else
                ErrorCode = 997
                ErrorMessage = "Failed to get description"
                log "Error getting document description from GetDescription in DM: Number: " & ErrorCode & ", Description: " & ErrorMessage
            End If
            
            rec.ReleaseResults
        Else
            success = 0
            log "Error: Could not connect to DM from GetDescription."
        End If
    Else
        ErrorCode = 801
        ErrorMessage = "Document Number must be Specifed"
        log "Error: Number: " & ErrorCode & ", Description: " & ErrorMessage
    End If

    log "GetDescription Function returned"
    GetDescription = S_OK
    
    Exit Function
    
errLabel:
    log "Error Source: " & Err.Source & ", Number: " & Err.Number & ", Description: " & Err.description
    GetDescription = S_OK
End Function

Public Function GetErrorCode(Optional ByRef pVal As Long) As Integer
    S_OK = 1

    pVal = ErrorCode

    GetErrorCode = S_OK
End Function

Public Function GetErrorMessage(Optional ByRef pVal As String) As Integer
    S_OK = 1

    pVal = ErrorMessage

    GetErrorMessage = S_OK
End Function

Private Function ConnectToDM() As Boolean
    On Error GoTo LogError
    Dim success As Boolean
    success = 0
    log "DMInterface connecting to DM"
    
    If RegCode = 0 Then
        LoadRegValues
    End If
    
    If RegCode = 1 Then
        Dim OK As Boolean
        Dim login As New PCDLogin
        
        Dim userstr As String
        userstr = RTrim(LTrim(UserId))
        Dim passstr As String
        passstr = RTrim(LTrim(Password))
        Dim libstr As String
        libstr = RTrim(LTrim(Library))
        
        login.AddLogin 0, libstr, userstr, passstr
        
        OK = (login.ErrNumber = 0)
        If (OK) Then
            login.Execute
            OK = (login.ErrNumber = 0)
        Else
            ErrorCode = login.ErrNumber
            ErrorMessage = login.ErrDescription
            log "Error connecting to DM: Number: " & login.ErrNumber & ", Description: " & login.ErrDescription
        End If
        
        If OK Then
            Group = login.GetPrimaryGroup()
            DST = login.GetDST
            
            success = 1
            log "Connected to DM: DST: " & DST
        Else
            ErrorCode = login.ErrNumber
            ErrorMessage = login.ErrDescription
            log "Error connecting to DM: Number: " & login.ErrNumber & ", Description: " & login.ErrDescription
        End If
    Else
        ErrorCode = 100
        ErrorMessage = "Registry Parameters not Fetched Properly"
        log "Error: 100, Description: Registry Parameters not Fetched Properly"
    End If
    
    ConnectToDM = success
    Exit Function
    
LogError:
    log "Error: Source: " & Err.Source & ", Number: " & Err.Number & ", Description: " & Err.description
    ConnectToDM = success
End Function

