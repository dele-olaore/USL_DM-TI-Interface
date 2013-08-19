VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DMInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Dim S_OK As Integer
Dim ErrorMessage As String
Dim ErrorCode As Long
'Dim RegCode As Integer

'Dim UserId As String, Password As String, Library As String, TrialCount As Integer, WebTopUrl As String
'Dim Profile As String, Query As String, PassordEncrypted As String
'Dim DST As String, Group As String

Dim transidHash As New Hashtable
Dim transid_ As String

Dim logFile As String

Enum AttachStatus
    NotComplete = 0
    Complete = 1
    Error1 = 2
End Enum

Private Function LoadRegValues(ByRef UserId As String, ByRef Password As String, ByRef Library As String, ByRef TrialCount As Integer, ByRef WebTopUrl As String, ByRef Profile As String, ByRef Query As String, ByRef PassordEncrypted As String) As Integer
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
    'UserId = "dmadmin"
    
    sValue = "Password"
    '***Read the String value
    vSetting = QueryValue(HKEY_CURRENT_USER, sKey, sValue)
    Password = "" & vSetting
    'Password = "00loopDS$"
    
    sValue = "Library"
    '***Read the String value
    vSetting = QueryValue(HKEY_CURRENT_USER, sKey, sValue)
    Library = "" & vSetting
    'Library = "DMSLIV"
    
    sValue = "TrialCount"
    '***Read the String value
    vSetting = QueryValue(HKEY_CURRENT_USER, sKey, sValue)
    TrialCount = CInt("" & vSetting)
    'TrialCount = 3
    
    sValue = "WebTopUrl"
    '***Read the String value
    vSetting = QueryValue(HKEY_CURRENT_USER, sKey, sValue)
    WebTopUrl = "" & vSetting
    'WebTopUrl = "http://10.1.11.14/cyberdocs" 'http://10.1.11.12/cyberdocs
    
    sKey = "Software\TIPlus_DMInterface\DocForms"
    
    sValue = "Profile"
    '***Read the String value
    vSetting = QueryValue(HKEY_CURRENT_USER, sKey, sValue)
    Profile = "" & vSetting
    'Profile = "TI_PROF"
    
    sValue = "Query"
    '***Read the String value
    vSetting = QueryValue(HKEY_CURRENT_USER, sKey, sValue)
    Query = "" & vSetting
    'Query = "TI_QBE"
    
    sKey = "Software\TIPlus_DMInterface\OtherParams"
    
    sValue = "PassordEncrypted"
    '***Read the String value
    vSetting = QueryValue(HKEY_CURRENT_USER, sKey, sValue)
    PassordEncrypted = "" & vSetting
    'PassordEncrypted = "N"
    
    LoadRegValues = 1
    log "Registry values read successfully!"
    
    Exit Function
    
RegFail:
    LoadRegValues = 0
    log "Error reading registry values. Source: " & Err.source & ", Number: " & Err.Number & ", Message: " & Err.description
End Function

Public Sub Class_Initialize()
    'RegCode = 0
    logFile = "c:\dminterface-log.txt"
    
    'S_OK = 1
    
    'DST = ""
    'Group = ""
End Sub

Public Sub Class_Terminate()
    On Error Resume Next
    
    'Set TS = Nothing
    'Set FSO = Nothing
End Sub

Public Sub log(ByVal message As String)
    On Error Resume Next
    
    Dim FSO As New FileSystemObject
    Dim TS As TextStream
    
    Dim TempS As String
    Dim Final As String
    
    'Set TS = FSO.OpenTextFile(logFile, ForReading, True)
    'Do Until TS.AtEndOfStream
    '    TempS = TS.ReadLine
    '    Final = Final & TempS & vbCrLf
    'Loop
    'TS.Close
    
    'Final = Final & "log at " & Now() & " - " & message & vbCrLf
    
    Set TS = FSO.OpenTextFile(logFile, ForAppending, True)
    TS.Write "log at " & Now() & " - " & message & vbCrLf 'Final
    TS.Close
    
    'Exit Sub
    
'closeStream:
    'log "log: Error: Source: " & Err.source & ", Number: " & Err.Number & ", Description: " & Err.description
    'On Error Resume Next
    'TS.Close
    'Set TS = Nothing
End Sub

Public Function View(Optional ByVal docId As String, Optional ByVal windowTitle As String, Optional ByVal MisysSessionId As String, Optional ByRef WebPageUrl As String, Optional ByRef success As Boolean) As Integer
    On Error GoTo LogError
    
    log Str(Now()) & ": TI calling View Function for Session: " & MisysSessionId & " Parameters(docId: " & docId & ", windowTitle: " & windowTitle & "MisysSessionId: " & MisysSessionId & ")"
    
    Dim UserId As String, Password As String, Library As String, TrialCount As Integer, WebTopUrl As String
    Dim Profile As String, Query As String, PasswordEncrypted As String
    Dim DST As String, Group As String
    Dim RegCode As Integer
    
    Dim S_OK As Integer
    
    S_OK = 1
    
    success = 0
    
    If Len(docId) = 0 Then
        ErrorCode = 800
        ErrorMessage = "Document Id must be Specifed"
        log "Error: 800, Description: Document Id must be Specifed"
    Else
        'If RegCode = 0 Then
            RegCode = LoadRegValues(UserId, Password, Library, TrialCount, WebTopUrl, Profile, Query, PasswordEncrypted)
        'End If
        
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
            log Str(Now()) & ": View Function returned for session: " & MisysSessionId & ", WebPageUrl: " & WebPageUrl
        Else
            ErrorCode = 100
            ErrorMessage = "Registry Parameters not Fetched Properly"
            log "Error: 100, Description: Registry Parameters not Fetched Properly"
        End If
    End If
    
    log Str(Now()) & ": View Function returned for session: " & MisysSessionId
    View = S_OK
    
    Exit Function
    
LogError:
    log Str(Now()) & ": Error in View Function for Session: " & MisysSessionId & ", Source: " & Err.source & ", Number: " & Err.Number & ", Description: " & Err.description
    View = S_OK
End Function

Public Function IsAvailable(Optional ByRef success As Boolean) As Integer
    On Error GoTo LogError
    
    log Str(Now()) & ": TI calling IsAvailable Function"
    
    Dim UserId As String, Password As String, Library As String, TrialCount As Integer, WebTopUrl As String
    Dim Profile As String, Query As String, PasswordEncrypted As String
    Dim DST As String, Group As String
    Dim RegCode As Integer
    
    Dim S_OK As Integer
    
    S_OK = 1
    success = 0

    RegCode = LoadRegValues(UserId, Password, Library, TrialCount, WebTopUrl, Profile, Query, PasswordEncrypted)

    If RegCode = 1 Then
        success = ConnectToDM(UserId, Password, Library, Group, DST)
    End If

    IsAvailable = S_OK
    log Str(Now()) & ": IsAvailable Function returned"
    Exit Function
    
LogError:
    log Str(Now()) & ": Error in IsAvailable Function, Source: " & Err.source & ", Number: " & Err.Number & ", Description: " & Err.description
    IsAvailable = S_OK
End Function

Public Function AttachTo(Optional ByVal name As String, Optional ByVal description As String, Optional ByVal windowTitle As String, Optional ByVal MisysSessionId As String, Optional ByRef WebPageUrl As String, Optional ByRef success As Boolean) As Integer
    On Error GoTo LogError
    
    log Str(Now()) & ": TI calling AttachTo Function for Session: " & MisysSessionId & " Parameters(name:" & name & ",description:" & description & ",windowtitle:" & windowTitle & ",MisysSessionId:" & MisysSessionId & ")"
    
    Dim UserId As String, Password As String, Library As String, TrialCount As Integer, WebTopUrl As String
    Dim Profile As String, Query As String, PasswordEncrypted As String
    Dim DST As String, Group As String
    Dim RegCode As Integer
    
    Dim S_OK As Integer
    
    S_OK = 1
    success = 0
    
    RegCode = LoadRegValues(UserId, Password, Library, TrialCount, WebTopUrl, Profile, Query, PasswordEncrypted)
    
    'If RegCode = 0 Then
    '    LoadRegValues
    'End If
    
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
        If transidHash.Exists(MisysSessionId) Then
            transidHash.Remove MisysSessionId
        End If
        transidHash.Add MisysSessionId, transid_
        
        Dim namestr As String
        namestr = ""
        Dim descstr As String
        descstr = ""
    
        Dim url As String
    
        url = webstr & "/quickstart.asp?show=SEARCHRESULTSACT:TI_ATTACHED=N:_FORMNAME=" & qbestr & "&library=" & libstr & "&eclip=yes&papi=bypass&user=" & userstr & ":" & passstr & "&tiname=" & namestr & "&tidescription=" & descstr & "&TIPlusSessionID=" & tistr & "&TRANSID=" & transid_
    
        WebPageUrl = RTrim(LTrim(url))
        success = 1
        log Str(Now()) & ": AttachTo Function returned for Session: " & MisysSessionId & ", Trans ID: " & transid_ & ", in hash: " & transidHash.item(MisysSessionId) & ", WebPageUrl: " & url
    Else
        ErrorCode = 100
        ErrorMessage = "Registry Parameters not Fetched Properly"
        log Str(Now()) & ": Error in AttachTo Function for Session: " & MisysSessionId & " Number: 100, Description: Registry Parameters not Fetched Properly"
    End If
    
    log Str(Now()) & ": AttachTo Function returned for Session: " & MisysSessionId
    AttachTo = S_OK
    
    Exit Function
    
LogError:
    log Str(Now()) & ": Error in AttachTo for Session: " & MisysSessionId & ", Source: " & Err.source & ", Number: " & Err.Number & ", Description: " & Err.description
    AttachTo = S_OK
End Function

Public Function EnhancedAttachTo(Optional ByVal name As String, Optional ByVal description As String, Optional ByVal windowTitle As String, Optional ByVal MisysSessionId As String, Optional ByVal XmlSearchCriteria As String, Optional ByRef WebPageUrl As String, Optional ByRef success As Boolean) As Integer
    On Error GoTo LogError
    
    log Str(Now()) & ": TI calling EnhancedAttachTo Function for Session: " & MisysSessionId & " Parameters(name:" & name & ",description:" & description & ",windowtitle:" & windowTitle & ",MisysSessionId:" & MisysSessionId & ", XmlSearchCriteria:" & XmlSearchCriteria & ")"
    
    Dim UserId As String, Password As String, Library As String, TrialCount As Integer, WebTopUrl As String
    Dim Profile As String, Query As String, PasswordEncrypted As String
    Dim DST As String, Group As String
    Dim RegCode As Integer
    
    Dim S_OK As Integer
    
    S_OK = 1
    success = 0
    
    RegCode = LoadRegValues(UserId, Password, Library, TrialCount, WebTopUrl, Profile, Query, PasswordEncrypted)

    'If RegCode = 0 Then
    '    LoadRegValues
    'End If
        
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
            If transidHash.Exists(MisysSessionId) Then
                transidHash.Remove MisysSessionId
            End If
            transidHash.Add MisysSessionId, transid_
            
            Dim namestr As String
            namestr = ""
            Dim descstr As String
            descstr = ""
            
            Dim url As String
            
            url = webstr & "/quickstart.asp?show=SEARCHRESULTSACT:TI_ATTACHED=N:_FORMNAME=" & qbestr & "&library=" & libstr & "&eclip=yes&papi=bypass&user=" & userstr & ":" & passstr & "&tiname=" & namestr & "&tidescription=" & descstr & "&TIPlusSessionID=" & tistr & "&TRANSID=" & transid_
            
            WebPageUrl = RTrim(LTrim(url))
            success = 1
            log Str(Now()) & ": EnhancedAttachTo Function for Session: " & MisysSessionId & ", Trans ID: " & transid_ & ", in hash: " & transidHash.item(MisysSessionId) & ", WebPageUrl: " & url
        'End If
    Else
        ErrorCode = 100
        ErrorMessage = "Registry Parameters not Fetched Properly"
        log Str(Now()) & ": Error in EnhancedAttachTo Function for Session: " & MisysSessionId & " Number: 100, Description: Registry Parameters not Fetched Properly"
    End If

    log Str(Now()) & ": EnhancedAttachTo Function returned for Session: " & MisysSessionId
    EnhancedAttachTo = S_OK
    
    Exit Function
    
LogError:
    log Str(Now()) & ": Error in EnhancedAttachTo for Session: " & MisysSessionId & ", Source: " & Err.source & ", Number: " & Err.Number & ", Description: " & Err.description
    EnhancedAttachTo = S_OK
End Function

Public Function AttachComplete(Optional ByVal MisysSessionId As String, Optional ByRef docId As String, Optional ByRef status As AttachStatus, Optional ByRef success As Boolean) As Integer
    On Error GoTo LogError
    
    log Str(Now()) & ": TI calling AttachComplete Function for Session: " & MisysSessionId & " Parameters(MisysSessionId:" & MisysSessionId & ")"
    
    Dim UserId As String, Password As String, Library As String, TrialCount As Integer, WebTopUrl As String
    Dim Profile As String, Query As String, PasswordEncrypted As String
    Dim DST As String, Group As String
    Dim RegCode As Integer
    
    Dim S_OK As Integer
    
    S_OK = 1
    status = NotComplete
    success = 0
    
    RegCode = LoadRegValues(UserId, Password, Library, TrialCount, WebTopUrl, Profile, Query, PasswordEncrypted)
    
    'If Len(DST) = 0 Then
        ConnectToDM UserId, Password, Library, Group, DST
    'End If
    
    If Len(DST) > 0 Then
        Dim rec As New PCDSearch
        Dim libstr As String
        libstr = RTrim(LTrim(Library))
        Dim qstr As String
        qstr = RTrim(LTrim(Query))
        
        Dim transid_ As String
        transid_ = transidHash.item(MisysSessionId)
        
        log Str(Now()) & ": AttachComplete: Session: " & MisysSessionId & ", Trans ID: " & transid_
        
        Dim id As String
        id = "" & MisysSessionId & transid_
        
        rec.SetDST DST
        rec.AddSearchLib libstr
        rec.SetSearchObject qstr
        rec.AddReturnProperty "DOCNUMBER"
        rec.AddSearchCriteria "TI_REF1", id
        
        rec.Execute
        If (rec.ErrNumber <> 0) Then
            ErrorCode = rec.ErrNumber
            ErrorMessage = rec.ErrDescription
            log Str(Now()) & ": Error searching for document in AttachComplete from DM: Session: " & MisysSessionId & ", Number: " & rec.ErrNumber & ", Description: " & rec.ErrDescription
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
                        
                        transidHash.Remove MisysSessionId
                        
                        log Str(Now()) & ": AttachComplete for Session: " & MisysSessionId & ", DOC: " & DOCNUMBER
                    End If
                Else
                    status = Error1
                    ErrorCode = 990
                    ErrorMessage = "No or Multiple Row Returned  for same DocumentID"
                    log Str(Now()) & ": Error searching for document in AttachComplete from DM: Session: " & MisysSessionId & ", Number: " & ErrorCode & ", Description: " & ErrorMessage
                End If
            End If
        End If
        
        rec.ReleaseResults
    Else
        status = NotComplete ' Error1
        success = 0
        log Str(Now()) & ": AttachComplete for Session: " & MisysSessionId & ", Error: Could not connect to DM from AttachComplete."
    End If
    
    log Str(Now()) & ": AttachComplete Function returned for Session: " & MisysSessionId
    AttachComplete = S_OK
    
    Exit Function
    
LogError:
    log Str(Now()) & ": Error in AttachComplete for Session: " & MisysSessionId & ", Source: " & Err.source & ", Number: " & Err.Number & ", Description: " & Err.description
    AttachComplete = S_OK
End Function

Public Function AttachTimeout(Optional ByVal MisysSessionId As String, Optional ByRef success As Boolean) As Integer
    On Error GoTo LogError
    
    log Str(Now()) & ": TI calling AttachTimeout Function for Session: " & MisysSessionId & " Parameters(MisysSessionId:" & MisysSessionId & ")"
    
    Dim S_OK As Integer
    
    S_OK = 1
    success = 1
    'log Str(Now()) & ": TI calling AttachTimeout Function"

    If transidHash.Exists(MisysSessionId) Then
        transidHash.Remove MisysSessionId
    End If

    log Str(Now()) & ": AttachTimeout Function returned for Session: " & MisysSessionId
    AttachTimeout = S_OK
    
    Exit Function
    
LogError:
    log Str(Now()) & ": Error in AttachTimeout for Session: " & MisysSessionId & ", Source: " & Err.source & ", Number: " & Err.Number & ", Description: " & Err.description
    AttachTimeout = S_OK
End Function

Public Function Delete(Optional ByVal docId As String, Optional ByRef success As Boolean) As Integer
    On Error GoTo LogError
    
    log Str(Now()) & ": TI calling Delete Function Parameters(docId:" & docId & ")"
    
    Dim S_OK As Integer
    
    S_OK = 1
    success = 0
    
    Dim UserId As String, Password As String, Library As String, TrialCount As Integer, WebTopUrl As String
    Dim Profile As String, Query As String, PasswordEncrypted As String
    Dim DST As String, Group As String
    Dim RegCode As Integer

    RegCode = LoadRegValues(UserId, Password, Library, TrialCount, WebTopUrl, Profile, Query, PasswordEncrypted)

    'If Len(DST) = 0 Then
        ConnectToDM UserId, Password, Library, Group, DST
    'End If
    
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
            log Str(Now()) & ": Delete Function successful"
        Else
            ErrorCode = 997
            ErrorMessage = "Failed to Delete Record"
            log Str(Now()) & ": Error deleting document from Delete in DM: DocID: " & docId & ", Number: " & docobj.ErrNumber & ", Description: " & docobj.ErrDescription
        End If
    Else
        success = 0
        log Str(Now()) & ": Error in Delete Function, Message: Could not connect to DM from Delete."
    End If

    log Str(Now()) & ": Delete Function returned for docId: " & docId
    Delete = S_OK
    
    Exit Function
    
LogError:
    log Str(Now()) & ": Error in Delete Function, docId: " & docId & ", Source: " & Err.source & ", Number: " & Err.Number & ", Description: " & Err.description
    Delete = S_OK
End Function

Public Function Detach(Optional ByVal docId As String, Optional ByRef success As Boolean) As Integer
    On Error GoTo LogError
    
    log Str(Now()) & ": TI calling Detach Function Parameters(docId:" & docId & ")"
    
    Dim S_OK As Integer
    
    S_OK = 1
    success = 0
    
    Dim UserId As String, Password As String, Library As String, TrialCount As Integer, WebTopUrl As String
    Dim Profile As String, Query As String, PasswordEncrypted As String
    Dim DST As String, Group As String
    Dim RegCode As Integer

    RegCode = LoadRegValues(UserId, Password, Library, TrialCount, WebTopUrl, Profile, Query, PasswordEncrypted)
    
    'If Len(DST) = 0 Then
        ConnectToDM UserId, Password, Library, Group, DST
    'End If
    
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
            log Str(Now()) & ": Detach Function successful"
        Else
            ErrorCode = 997
            ErrorMessage = "Failed to Detach Record"
            log Str(Now()) & ": Error detaching document from Detach in DM: DocID: " & docId & ", Number: " & docobj.ErrNumber & ", Description: " & docobj.ErrDescription
        End If
    Else
        success = 0
        log Str(Now()) & ": Error: Could not connect to DM from Detach, DocID: " & docId
    End If

    log Str(Now()) & ": Detach Function returned for DocId: " & docId
    Detach = S_OK
    
    Exit Function
    
LogError:
    log Str(Now()) & ": Error in Detach Function, docId: " & docId & ", Source: " & Err.source & ", Number: " & Err.Number & ", Description: " & Err.description
    Detach = S_OK
End Function

Public Function SetName(Optional ByVal docId As String, Optional ByVal name As String, Optional ByRef success As Boolean) As Integer
    On Error GoTo LogError
    
    log Str(Now()) & ": TI calling SetName Function Parameters(docId:" & docId & ", name: " & name & ")"
    
    Dim S_OK As Integer
    
    S_OK = 1
    success = 0
    
    Dim UserId As String, Password As String, Library As String, TrialCount As Integer, WebTopUrl As String
    Dim Profile As String, Query As String, PasswordEncrypted As String
    Dim DST As String, Group As String
    Dim RegCode As Integer

    RegCode = LoadRegValues(UserId, Password, Library, TrialCount, WebTopUrl, Profile, Query, PasswordEncrypted)

    'If Len(DST) = 0 Then
        ConnectToDM UserId, Password, Library, Group, DST
    'End If
    
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
                log Str(Now()) & ": SetName updated doc: " & docId & ", name: " & name
            Else
                ErrorCode = 997
                ErrorMessage = "Failed to Update name"
                log Str(Now()) & ": Error updating document from SetName in DM, DocId: " & docId & ", Number: " & docobj.ErrNumber & ", Description: " & docobj.ErrDescription
            End If
        Else
            success = 0
            log Str(Now()) & ": Error in SetName for DocId: " & docId & ", Could not connect to DM from SetName."
        End If
    Else
        ErrorCode = 801
        ErrorMessage = "Document Number and Name must be Specifed"
        log Str(Now()) & ": Error in SetName for DocId: " & docId & ", Number: " & ErrorCode & ", Description: " & ErrorMessage
    End If

    log Str(Now()) & ": SetName Function returned for DocId: " & docId
    SetName = S_OK
    
    Exit Function
    
LogError:
    log Str(Now()) & ": Error in SetName Function, docId: " & docId & ", Source: " & Err.source & ", Number: " & Err.Number & ", Description: " & Err.description
    SetName = S_OK
End Function

Public Function SetDescription(Optional ByVal docId As String, Optional ByVal description As String, Optional ByRef success As Boolean) As Integer
    On Error GoTo LogError
    
    log Str(Now()) & ": TI calling SetDescription Function Parameters(docId:" & docId & ", description: " & description & ")"
    
    Dim S_OK As Integer
    
    S_OK = 1
    success = 0
    
    Dim UserId As String, Password As String, Library As String, TrialCount As Integer, WebTopUrl As String
    Dim Profile As String, Query As String, PasswordEncrypted As String
    Dim DST As String, Group As String
    Dim RegCode As Integer

    RegCode = LoadRegValues(UserId, Password, Library, TrialCount, WebTopUrl, Profile, Query, PasswordEncrypted)

    'If Len(DST) = 0 Then
        ConnectToDM UserId, Password, Library, Group, DST
    'End If
    
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
                log Str(Now()) & ": SetDescription updated doc: " & docId & ", description: " & description
            Else
                ErrorCode = 997
                ErrorMessage = "Failed to Update description"
                log Str(Now()) & ": Error updating document from SetDescription in DM DocId: " & docId & ", Number: " & docobj.ErrNumber & ", Description: " & docobj.ErrDescription
            End If
        Else
            success = 0
            log Str(Now()) & ": Error in SetDescription for docId: " & docId & ", Could not connect to DM from SetDescription."
        End If
    Else
        ErrorCode = 801
        ErrorMessage = "Document Number and Description must be Specifed"
        log Str(Now()) & ": Error in SetDescription for docId: " & docId & ", Number: " & ErrorCode & ", Description: " & ErrorMessage
    End If

    log Str(Now()) & ": SetDescription Function returned for docId: " & docId
    SetDescription = S_OK
    
    Exit Function
    
LogError:
    log Str(Now()) & ": Error in SetDescription Function, docId: " & docId & ", Source: " & Err.source & ", Number: " & Err.Number & ", Description: " & Err.description
    SetDescription = S_OK
End Function

Public Function GetName(Optional ByVal docId As String, Optional ByRef name As String, Optional ByRef success As Boolean) As Integer
    On Error GoTo errLabel
    
    log Str(Now()) & ": TI calling GetName Function Parameters(docId:" & docId & ")"
    
    Dim S_OK As Integer
    
    S_OK = 1
    success = 0
    
    Dim UserId As String, Password As String, Library As String, TrialCount As Integer, WebTopUrl As String
    Dim Profile As String, Query As String, PasswordEncrypted As String
    Dim DST As String, Group As String
    Dim RegCode As Integer

    RegCode = LoadRegValues(UserId, Password, Library, TrialCount, WebTopUrl, Profile, Query, PasswordEncrypted)

    'If Len(DST) = 0 Then
        ConnectToDM UserId, Password, Library, Group, DST
    'End If
    
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
                        log Str(Now()) & ": GetName success: DOCNUMBER: " & docId & ", Name: " & name
                    End If
                Else
                    ErrorCode = 990
                    ErrorMessage = "No or Multiple Row Returned for same DocumentID"
                    log Str(Now()) & ": Error getting document name: DocId: " & docId & ", Number: " & ErrorCode & ", Description: " & ErrorMessage
                End If
            Else
                ErrorCode = 997
                ErrorMessage = "Failed to get name"
                log Str(Now()) & ": Error getting document name from GetName in DM: Number: " & ErrorCode & ", Description: " & ErrorMessage
            End If
            
            rec.ReleaseResults
        Else
            success = 0
            log Str(Now()) & ": Error: Could not connect to DM from GetName."
        End If
    Else
        ErrorCode = 801
        ErrorMessage = "Document Number must be Specifed"
        log Str(Now()) & ": Error: Number: " & ErrorCode & ", Description: " & ErrorMessage
    End If

    log Str(Now()) & ": GetName Function returned for DocId: " & docId
    GetName = S_OK
    
    Exit Function
errLabel:
    log Str(Now()) & ": Error in GetName Function, docId: " & docId & ", Source: " & Err.source & ", Number: " & Err.Number & ", Description: " & Err.description
    GetName = S_OK
End Function

Public Function GetDescription(Optional ByVal docId As String, Optional ByRef description As String, Optional ByRef success As Boolean) As Integer
    On Error GoTo errLabel
    
    log Str(Now()) & ": TI calling GetDescription Function Parameters(docId:" & docId & ")"
    
    Dim S_OK As Integer
    
    S_OK = 1
    success = 0
    
    Dim UserId As String, Password As String, Library As String, TrialCount As Integer, WebTopUrl As String
    Dim Profile As String, Query As String, PasswordEncrypted As String
    Dim DST As String, Group As String
    Dim RegCode As Integer

    RegCode = LoadRegValues(UserId, Password, Library, TrialCount, WebTopUrl, Profile, Query, PasswordEncrypted)

    'If Len(DST) = 0 Then
        ConnectToDM UserId, Password, Library, Group, DST
    'End If
    
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
                        log Str(Now()) & ": GetDescription success: DOCNUMBER: " & docId & ", Description: " & description
                    End If
                Else
                    ErrorCode = 990
                    ErrorMessage = "No or Multiple Row Returned for same DocumentID"
                    log Str(Now()) & ": Error getting document description: Number: " & ErrorCode & ", Description: " & ErrorMessage
                End If
            Else
                ErrorCode = 997
                ErrorMessage = "Failed to get description"
                log Str(Now()) & ": Error getting document description from GetDescription in DM: Number: " & ErrorCode & ", Description: " & ErrorMessage
            End If
            
            rec.ReleaseResults
        Else
            success = 0
            log Str(Now()) & ": Error: Could not connect to DM from GetDescription."
        End If
    Else
        ErrorCode = 801
        ErrorMessage = "Document Number must be Specifed"
        log Str(Now()) & ": Error: Number: " & ErrorCode & ", Description: " & ErrorMessage
    End If

    log Str(Now()) & ": GetDescription Function returned"
    GetDescription = S_OK
    
    Exit Function
    
errLabel:
    log Str(Now()) & ": Error in GetDescription Function, docId: " & docId & ", Source: " & Err.source & ", Number: " & Err.Number & ", Description: " & Err.description
    GetDescription = S_OK
End Function

Public Function GetErrorCode(Optional ByRef pVal As Long) As Integer
    Dim S_OK As Integer

    S_OK = 1

    pVal = ErrorCode

    GetErrorCode = S_OK
End Function

Public Function GetErrorMessage(Optional ByRef pVal As String) As Integer
    Dim S_OK As Integer

    S_OK = 1

    pVal = ErrorMessage

    GetErrorMessage = S_OK
End Function

Private Function ConnectToDM(ByVal UserId As String, ByVal Password As String, ByVal Library As String, ByRef Group As String, ByRef DST As String) As Boolean
    On Error GoTo LogError
    Dim success As Boolean
    success = 0
    log "DMInterface connecting to DM"
    
    'If RegCode = 0 Then
    '    LoadRegValues
    'End If
    
    'If RegCode = 1 Then
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
            log "Connected to DM: DST: " ' & DST
        Else
            ErrorCode = login.ErrNumber
            ErrorMessage = login.ErrDescription
            log "Error connecting to DM: Number: " & login.ErrNumber & ", Description: " & login.ErrDescription
        End If
    'Else
    '    ErrorCode = 100
    '    ErrorMessage = "Registry Parameters not Fetched Properly"
    '    log "Error: 100, Description: Registry Parameters not Fetched Properly"
    'End If
    
    ConnectToDM = success
    Exit Function
    
LogError:
    log "Error: Source: " & Err.source & ", Number: " & Err.Number & ", Description: " & Err.description
    ConnectToDM = success
End Function
