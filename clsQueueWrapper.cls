VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsQueueWrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3BCC961B0380"
'Encapsulate the common functionality of the MSMQQueue, MSMQQueueInfo,
'MSMQMessage, and MSMQEvent objects.
Option Explicit

'Private variables for the properties
Private mlngEventTimeout As Long
Private mblnEnableNotification As Boolean
Private mintAccessMode As MSMQ.MQACCESS
Private mintShareMode As MSMQ.MQSHARE
Private mstrQueueLabel As String
Private mobjQueue As MSMQ.MSMQQueue
Private mobjQueueInfo As MSMQ.MSMQQueueInfo
Private mobjQueueMessage As MSMQ.MSMQMessage
Private WithEvents mobjQueueEvent As MSMQ.MSMQEvent
Attribute mobjQueueEvent.VB_VarHelpID = -1

'Fires to indicate the arrival of a message in the underlying
'queue when AccessMode = MQ_RECEIVE_ACCESS or MQ_PEEK_ACCESS.
'
Public Event Arrived()
Private Sub Class_Initialize()
    On Error Resume Next
    Set mobjQueue = New MSMQ.MSMQQueue
    Set mobjQueueInfo = New MSMQ.MSMQQueueInfo
    Set mobjQueueMessage = New MSMQ.MSMQMessage
    Set mobjQueueEvent = New MSMQ.MSMQEvent
    mlngEventTimeout = 500
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    Set mobjQueue = Nothing
    Set mobjQueueInfo = Nothing
    Set mobjQueueMessage = Nothing
    Set mobjQueueEvent = Nothing
End Sub
'Fires when a message appears in the queue.
Private Sub mobjQueueEvent_Arrived(ByVal Queue As Object, ByVal Cursor As Long)
    On Error Resume Next
    'Make sure we'll have a queue to receive from and
    'that it was opened for Peek or Receive mode
    If mobjQueue.IsOpen = False Then
        Exit Sub
    ElseIf mintAccessMode = MQ_SEND_ACCESS Then
        Exit Sub
    End If
    
    'Notify the client
    RaiseEvent Arrived
           
End Sub
Private Sub mobjQueueEvent_ArrivedError(ByVal Queue As Object, ByVal ErrorCode As Long, ByVal Cursor As Long)
    'only care for non-send queues
    If mintAccessMode <> MQ_SEND_ACCESS Then
        If ErrorCode = MQ_ERROR_IO_TIMEOUT Then
            'if it was just a timeout error, re-enable
            'notification if necessary
            If mblnEnableNotification = True Then
                mobjQueue.EnableNotification mobjQueueEvent, , mlngEventTimeout
            End If
        End If
    End If
End Sub

'Purpose:   Returns a recordset formatted for the SendMessage method.
'Comment:   If the Originating ID is omitted, the MSS Gateway service
'           will insert its ID prior to transmission.
Public Function CreateMessage(ByVal pstrQueueLabel As String, _
    ByVal pstrAux As String, ByVal pstrDestinationMnem As String, _
    ByVal pstrBody As String, Optional pstrOrigID) As ADODB.Recordset
    
    On Error Resume Next
    Dim strOrigID As String
    
    'Default to 5 space string
    strOrigID = Space(5)
    If Not IsMissing(pstrOrigID) Then
        strOrigID = pstrOrigID
    End If
    
    If pstrBody = "" Then
        pstrBody = " "
    End If
    
    
    'IMPORTANT NOTE:
    'The wrong number of delimiters can bring down the
    'NCIC connection (trust me), so strip them out to
    'insure no delimiters are in the fields.
    strOrigID = Replace(strOrigID, ".", " ")
    pstrAux = Replace(pstrAux, ".", " ")
    pstrDestinationMnem = Replace(pstrDestinationMnem, ".", " ")
    
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    rs.Fields.Append "MsgDate", adDate, 8, adFldIsNullable
    rs.Fields.Append "QueueLabel", adVarChar, 8, adFldIsNullable
    rs.Fields.Append "OrigID", adChar, 5, adFldIsNullable
    rs.Fields.Append "Aux", adChar, 4, adFldIsNullable
    rs.Fields.Append "Mnem", adVarChar, 255, adFldIsNullable
    rs.Fields.Append "Delimiter", adChar, 1
    rs.Fields.Append "Body", adLongVarWChar, Len(pstrBody), adFldIsNullable
    rs.Open
    
    rs.AddNew
    rs("MsgDate") = Now
    rs("QueueLabel") = pstrQueueLabel
    rs("OrigID") = Left(Format(Trim(strOrigID), "!@@@@@"), 5)
    rs("Aux") = Left(Format(Trim(pstrAux), "!@@@@"), 4)
    rs("Mnem") = Left(Format(Trim(pstrDestinationMnem), "!@@@@@"), 5)
    rs("Delimiter") = "."
    rs("Body") = pstrBody
    rs.Update
    
    Set CreateMessage = rs.Clone
    Set rs = Nothing
    
End Function

'Close the reference to the queue.
Public Sub QueueClose()
    On Error Resume Next
    mobjQueue.Close
End Sub

'Locate and open a queue with pintAccessMode and pintShareMode functionality
'based upon the queue's name.
'
'If the Optional pstrMachinePath value is not supplied, attempt
'to locate the queue using a directory service. If pstrMachinePath
'is supplied, use it to open the queue.
'
'The Optional ReceiveTimeout value is used for the event notification
'on non-send queues.
'
'The Optional pblnEnableNotification is used to determine whether the
'queue event object should raise an event and notify the hosting
'process when events occur.
Public Sub QueueOpen(ByVal pstrQueueName As String, _
    ByVal pintAccessMode As MSMQ.MQACCESS, _
    ByVal pintShareMode As MSMQ.MQSHARE, _
    Optional pstrMachinePath, _
    Optional plngReceiveTimeout, _
    Optional pblnEnableNotification)
    
    On Error GoTo Err_QueueOpen
    Dim objQueueQuery As MSMQ.MSMQQuery
    Dim objQueueInfos As MSMQ.MSMQQueueInfos
    Dim intCnt As Integer
    Dim blnEnableNotification As Boolean
    
    
    'Can only open a queue with a name
    If pstrQueueName = "" Then
        Exit Sub
    End If
    
    'capture a new value for the event timeout
    If Not IsMissing(plngReceiveTimeout) Then
        If IsNumeric(plngReceiveTimeout) Then
            If plngReceiveTimeout > 0 Or plngReceiveTimeout = -1 Then
                mlngEventTimeout = plngReceiveTimeout
            End If
        End If
    End If
    
    'assume true
    mblnEnableNotification = True
    If Not IsMissing(pblnEnableNotification) Then
        mblnEnableNotification = pblnEnableNotification
    End If
    
    If Not IsMissing(pstrMachinePath) Then
        'Use the explicit machine name path
        
        'If passed an empty string, use local machine
        If pstrMachinePath = "" Then
            pstrMachinePath = ".\PRIVATE$\"
        End If
        If Right(pstrMachinePath, 1) <> "\" Then
            pstrMachinePath = pstrMachinePath & "\"
        End If
        mobjQueueInfo.FormatName = "DIRECT=OS:" & _
            pstrMachinePath & pstrQueueName
    Else
        'Use a directory service to locate the Queue by label.
        Set objQueueQuery = New MSMQ.MSMQQuery
        Set objQueueInfos = objQueueQuery.LookupQueue(Label:=pstrQueueName)
        '
        'Verfy there is only one queue with this name
        objQueueInfos.Reset
        intCnt = 0
        While Not (mobjQueueInfo Is Nothing)
            intCnt = intCnt + 1
        Wend
        If intCnt <> 1 Then
            'Shouldn't happen, but multiple queues found
            'with this name
            Exit Sub
        End If
        '
        'Set the queue info object
        objQueueInfos.Reset
        Set mobjQueueInfo = objQueueInfos.Next
    End If
    
    'Close the old queue if necessary
    If CBool(mobjQueue.IsOpen) = True Then
        mobjQueue.Close
    End If
    
    'Update the internal variables and open the queue.
    mstrQueueLabel = pstrQueueName
    mintAccessMode = pintAccessMode
    mintShareMode = pintShareMode
    
    Set mobjQueue = mobjQueueInfo.Open(pintAccessMode, pintShareMode)
    
    'Enable our notification for the correct access type
    If mintAccessMode <> MQ_SEND_ACCESS And mblnEnableNotification = True Then
        mobjQueue.EnableNotification mobjQueueEvent, , mlngEventTimeout
    End If
    
    'Clean up resources
    On Error Resume Next
    Set objQueueQuery = Nothing
    Set objQueueInfos = Nothing
    
    Exit Sub
Err_QueueOpen:
    Err.Raise Err.Number, App.EXEName & ".QueueOpen: ", _
        Err.Description
End Sub

'Returns a recordset containing the received message from the remote MSS
'System. Valid only when AccessMode = MQ_RECEIVE_ACCESS or MQ_PEEK_ACCESS.
'
'When AccessMode = MQ_RECEIVE_ACCESS and pblnPeekOnly = True, the message is
'not removed from the underlying queue. If AccessMode = MQ_PEEK_ACCESS, the
'message is not removed from the underlying queue and pblnPeekOnly is ignored.
'
'plngReceiveTimeout is an optional parameter used to determine the number of
'milliseconds to wait when reading the queue. Used for performance tuning, the
'default value of 500ms should suffice due to event notification. A value of
'-1 disables the waiting period and the application will wait for an infinite
'amount of time if the queue is empty.
Public Function ReceiveMessage(ByVal pblnPeekOnly As Boolean, _
    Optional plngReceiveTimeout) As ADODB.Recordset
    On Error GoTo Err_ReceiveMessage
    On Error Resume Next
    
    Dim lngReceiveTimeout As Long
    Dim rs As ADODB.Recordset
    
    Set rs = New ADODB.Recordset
    rs.Fields.Append "MsgDate", adDate, 8, adFldIsNullable
    rs.Fields.Append "QueueLabel", adVarChar, 8, adFldIsNullable
    rs.Fields.Append "OrigID", adChar, 5, adFldIsNullable
    rs.Fields.Append "Aux", adChar, 4, adFldIsNullable
    rs.Fields.Append "Mnem", adVarChar, 255
    rs.Fields.Append "Delimiter", adChar, 1
    rs.Fields.Append "Body", adLongVarWChar
    rs.Open
    
    'initialize an empty recordset
    Set ReceiveMessage = rs
    
    lngReceiveTimeout = mlngEventTimeout
    
    If mobjQueue.IsOpen = False Then
        Err.Raise MQ_ERROR_QUEUE_NOT_AVAILABLE, _
            App.EXEName & ".ReceiveMessage: ", Err.Description
        Exit Function
    ElseIf mobjQueue.Access = MQ_SEND_ACCESS Then
        Err.Raise MQ_ERROR_ACCESS_DENIED, _
            App.EXEName & ".ReceiveMessage: ", Err.Description
        Exit Function
    End If
    
    If Not IsMissing(plngReceiveTimeout) Then
        If IsNumeric(plngReceiveTimeout) Then
            If plngReceiveTimeout > 0 Or plngReceiveTimeout = -1 Then
                lngReceiveTimeout = plngReceiveTimeout
            End If
        End If
    End If
    
    If mintAccessMode = MQ_PEEK_ACCESS Or pblnPeekOnly = True Then
        'peek mode
        Set mobjQueueMessage = mobjQueue.Peek(, , lngReceiveTimeout)
    Else
        'receive mode
        Set mobjQueueMessage = mobjQueue.Receive(, , , lngReceiveTimeout)
    End If
    'should be a recordset
    Set ReceiveMessage = mobjQueueMessage.Body
    
    'Re-enable the event notification
    If mblnEnableNotification = True Then
        mobjQueue.EnableNotification mobjQueueEvent, , mlngEventTimeout
    End If
    Exit Function

Err_ReceiveMessage:
    'Re-enable the event notification
    If mblnEnableNotification = True Then
        mobjQueue.EnableNotification mobjQueueEvent, , mlngEventTimeout
    End If
    Err.Raise Err.Number, App.EXEName & ".ReceiveMessage: ", _
        Err.Description
End Function

'Valid only when AccessMode = MQ_SEND_ACCESS. Pass a recordset and optional
'Priority value for transmission to the underlying queue object. Alternatively,
'invoking CreateMessage will return a recordset formatted for
'SendMessage.
'
'The recordset should consist of the following fields:
'
'   QueueLabel - string identifier that uniquely determines the message type.
'   OrigID -    Originating Terminal ID. All applications should pass a received
'               OrigID value when generating a response.
'   MsgDate -   Date of the message.
'   Aux -       Auxilliary field. All applications should pass a received Aux value
'               when generating a response.
'   Mnem -      Destination mnemonic. Application specific code for the remote MSS
'               system.
'   Delimiter   "."
'   Body -      Complete message body to be sent to the remote MSS System.
Public Sub SendMessage(ByVal prsMessage As ADODB.Recordset, _
    Optional pbytPriority)
    On Error GoTo Err_SendMessage
    
    If mobjQueue.IsOpen = False Then
        Exit Sub
    ElseIf mobjQueue.Access <> MQ_SEND_ACCESS Then
        Exit Sub
    End If
    
    mobjQueueMessage.Label = prsMessage("QueueLabel")
    
    'Even though we are passing an object,
    'do not use the Set operator per MSDN instructions
    mobjQueueMessage.Body = prsMessage
    
    mobjQueueMessage.Delivery = MQMSG_DELIVERY_RECOVERABLE
    mobjQueueMessage.Journal = MQMSG_DEADLETTER
    'timeout after 30 seconds if it can't be written to the queue
    mobjQueueMessage.MaxTimeToReachQueue = 30
    'allow the message to reside in the queue until read
    mobjQueueMessage.MaxTimeToReceive = -1
    'Set the message priority (future use)
    mobjQueueMessage.Priority = 3
    If Not IsMissing(pbytPriority) Then
        If pbytPriority < 8 Then
            mobjQueueMessage.Priority = pbytPriority
        End If
    End If
    mobjQueueMessage.Send mobjQueue
    Exit Sub
Err_SendMessage:
    Err.Raise Err.Number, App.EXEName & ".SendMessage: ", _
        Err.Description
End Sub
'Returns the access mode of the queue. Read-only.
Public Property Get AccessMode() As MSMQ.MQACCESS
    On Error Resume Next
    AccessMode = mintAccessMode
End Property
'Returns whether a queue is open. Read-only.
Public Property Get IsOpen() As Boolean
    On Error Resume Next
    If IsObject(mobjQueue) Then
        IsOpen = CBool(mobjQueue.IsOpen)
    End If
End Property

'Returns the Label of the Queue. Read-only.
Public Property Get QueueLabel() As String
    On Error Resume Next
    QueueLabel = mstrQueueLabel
End Property

'Returns a reference to the underlying MSMQ Queue object.
Public Property Get Queue() As MSMQ.MSMQQueue
    On Error Resume Next
    Set Queue = mobjQueue
End Property
'Sets a reference for the underlying MSMQ Queue object.
Public Property Set Queue(ByVal pobjQueue As MSMQ.MSMQQueue)
    On Error Resume Next
    Set mobjQueue = pobjQueue
End Property
'Returns a reference to the underlying MSMQQueueInfo object.
Public Property Get QueueInfo() As MSMQ.MSMQQueueInfo
    On Error Resume Next
    Set QueueInfo = mobjQueueInfo
End Property
'Sets a reference for the underlying MSMQQueueInfo object.
'This affects the underlying queue, so close and re-open
'the queue using the new QueueInfo.Label
Public Property Set QueueInfo(ByVal pobjQueueInfo As MSMQ.MSMQQueueInfo)
    On Error Resume Next
    Set mobjQueueInfo = pobjQueueInfo
    mobjQueue.Close
    QueueOpen mobjQueueInfo.Label, mintAccessMode, mintShareMode
End Property
'Returns a reference to the underlying MSMQMessage object.
Public Property Get QueueMessage() As MSMQ.MSMQMessage
    On Error Resume Next
    Set QueueMessage = mobjQueueMessage
End Property
'Sets a reference for the underlying MSMQMessage object.
Public Property Set QueueMessage(ByVal pobjQueueMessage As MSMQ.MSMQMessage)
    On Error Resume Next
    Set mobjQueueMessage = pobjQueueMessage
End Property

'Returns the share mode of the queue. Read-only when the queue is open.
Public Property Get ShareMode() As MSMQ.MQSHARE
    On Error Resume Next
    ShareMode = mintShareMode
End Property


