'
' // Callback ISampleGrabberCB object implementaion
' // by The trick 2018
'
' // Because of a callbacks are received from different threads, we can't implement such
' // object in a class module. A CSampleGrabberBuffer object uses circular buffering.
' // It assigns an offset with the playback time and we can get the playback offset using
' // such information. Playback offset always is before writing time and we should hold
' // some samples in the buffer. Reset method is used to reset the time asiigned to offset.
'
' // --> time
' // ---------+ +----------------------------------------------+ +---------
' //          | |       playback             writing           | |
' // \        | |          |                    |              | |        \
' // /        | |          |                    |              | |        /
' // \        | |         \ /                  \ /             | |        \
' //    13 14 | | 00 01 02 03 04 05 06 07 08 09 10 11 12 13 14 | | 01 02
' // ---------+ +----------------------------------------------+ +---------
'
' // Playback cursor is controlled by IMediaPosition::CurrentPosition in the main thread
' // periodically from the timer event. Writing cursor is controlled by SampleCB method is
' // called by Sample Grabber in arbitrary thread. When writing cursor reaches the end of
' // the buffer it's moved to begin. The same behavior is in playback cursor.
' // If a callback object is reset the writing cursor is assigned to the sample time.
' // A callback object uses that time to calculate offsets to the next writing position and
' // the playback position.
'

Option Explicit

Public Const IIDSTR_ISampleGrabberCB        As String = "{0579154A-2B53-4994-B0D0-E773148EFF85}"
Public Const IIDSTR_ISampleGrabberBuffer    As String = "{cc7df97a-9f40-4ac4-8e6a-8d0e8f4eeaca}"
Public Const WMMEDIATYPE_Audio              As String = "{73647561-0000-0010-8000-00AA00389B71}"
Public Const FORMAT_WaveFormatEx            As String = "{05589F81-C356-11CE-BF01-00AA0055595A}"

' // User sample grabber virtual methods table
Private Type ISampleGrabberBuffer_vtbl
    pfnQueryInterface       As Long
    pfnAddRef               As Long
    pfnRelease              As Long
    pfnSampleCB             As Long
    pfnBufferCB             As Long
    pfnReset                As Long
    pfnGetOffsetByTime      As Long
End Type

' // User sample grabber object
Private Type CSampleGrabberBuffer
    pVtbl                   As Long         ' // Pointer to vTable
    lRefCounter             As Long         ' // Reference counter
    pData                   As Long         ' // Pointer to buffer
    lDataSize               As Long         ' // Size of buffer
    lOffset                 As Long         ' // Buffer offset write position
    dCurrentTime            As Double       ' // Current time of the offset
    bIsReset                As Boolean      ' // If true the time isn't defined
    tFormat                 As WAVEFORMATEX ' // Buffer format
End Type

Private mtVtable                    As ISampleGrabberBuffer_vtbl
Private mtIID_IUnknown              As UUID
Private mtIID_ISampleGrabberCB      As UUID
Private mtIID_ISampleGrabberBuffer  As UUID

' // Create user sample grabber object
Public Function CreateSampleGrabberCB( _
                ByVal pData As Long, _
                ByVal lSize As Long, _
                ByRef tFormat As WAVEFORMATEX) As ISampleGrabberBuffer
    Dim tLocObj As CSampleGrabberBuffer
    Dim pObject As Long
    
    If mtVtable.pfnQueryInterface = 0 Then
        
        ' // Setup vTable
        mtVtable.pfnQueryInterface = FARPROC(AddressOf QueryInterface)
        mtVtable.pfnAddRef = FARPROC(AddressOf AddRef)
        mtVtable.pfnRelease = FARPROC(AddressOf Release)
        mtVtable.pfnSampleCB = FARPROC(AddressOf SampleCB)
        mtVtable.pfnBufferCB = FARPROC(AddressOf BufferCB)
        mtVtable.pfnReset = FARPROC(AddressOf Reset)
        mtVtable.pfnGetOffsetByTime = FARPROC(AddressOf GetOffsetByTime)
        
        ' // Setup IIDs
        GUIDFromString StrPtr(IIDSTR_IUnknown), mtIID_IUnknown
        GUIDFromString StrPtr(IIDSTR_ISampleGrabberCB), mtIID_ISampleGrabberCB
        GUIDFromString StrPtr(IIDSTR_ISampleGrabberBuffer), mtIID_ISampleGrabberBuffer
        
    End If
    
    ' // Setup object
    With tLocObj
    
    .lRefCounter = 1
    .pVtbl = VarPtr(mtVtable)
    .lDataSize = lSize
    .pData = pData
    .tFormat = tFormat
    .bIsReset = True
    
    End With
    
    ' // Alloc memory for object
    pObject = CoTaskMemAlloc(Len(tLocObj))
    
    ' // Copy object
    MoveMemory ByVal pObject, tLocObj, Len(tLocObj)
    
    ' // Cast
    GetMem4 pObject, CreateSampleGrabberCB
    
End Function

Public Function MakeTrue( _
                ByRef bValue As Boolean) As Boolean
    bValue = True
    MakeTrue = True
End Function

' // Query interface callback
Private Function QueryInterface( _
                 ByRef cObject As CSampleGrabberBuffer, _
                 ByRef tIId As UUID, _
                 ByRef ppOut As Long) As Long
                 
    If IsEqualGUID(tIId, mtIID_IUnknown) Or _
        IsEqualGUID(tIId, mtIID_ISampleGrabberCB) Or _
        IsEqualGUID(tIId, mtIID_ISampleGrabberBuffer) Then
        ppOut = VarPtr(cObject)
        AddRef cObject
    Else
        QueryInterface = E_NOINTERFACE
    End If
                 
End Function

' // AddRef callback
Private Function AddRef( _
                 ByRef cObject As CSampleGrabberBuffer) As Long
    AddRef = InterlockedIncrement(cObject.lRefCounter)
End Function

' // Release callback
Private Function Release( _
                 ByRef cObject As CSampleGrabberBuffer) As Long
    
    Release = InterlockedDecrement(cObject.lRefCounter)
    
    If Release <= 0 Then
        CoTaskMemFree VarPtr(cObject)
    End If
    
End Function

' // SampleCB callback (not implemented)
Private Function SampleCB( _
                 ByRef cObject As CSampleGrabberBuffer, _
                 ByVal dSampleTime As Double, _
                 ByVal cSample As IUnknown) As Long
    SampleCB = E_NOTIMPL
End Function

' // Called by Sample Grabber filter
Private Function BufferCB( _
                 ByRef cObject As CSampleGrabberBuffer, _
                 ByVal dSampleTime As Double, _
                 ByVal pBuffer As Long, _
                 ByVal lBufferLen As Long) As Long
    Dim lWritten    As Long
    Dim dCurTime    As Double
    Dim lOffset     As Long
    
    ' // Get offset according to previous sample
    If Not cObject.bIsReset Then
        lOffset = GetOffsetByTime(cObject, dSampleTime)
    End If
    
    ' // Ensure circular bounds
    If lOffset + lBufferLen > cObject.lDataSize Then
        lWritten = cObject.lDataSize - lOffset
    Else
        lWritten = lBufferLen
    End If

    MoveMemory ByVal cObject.pData + lOffset, ByVal pBuffer, lWritten
    
    If lWritten <> lBufferLen Then
        MoveMemory ByVal cObject.pData, ByVal pBuffer + lWritten, lBufferLen - lWritten
    End If
    
    ' // Update time (if need)
    dCurTime = dSampleTime + lBufferLen / cObject.tFormat.nAvgBytesPerSec
    
    If dCurTime > cObject.dCurrentTime Or cObject.bIsReset Then
    
        cObject.dCurrentTime = dCurTime
        cObject.lOffset = lOffset + lBufferLen
        cObject.bIsReset = False
    
    End If
    
End Function

' // Reset time (when user select new playback position by slider)
Private Function Reset( _
                 ByRef cObject As CSampleGrabberBuffer) As Long
    cObject.bIsReset = True
End Function

' // Get the offset by time (using previous writing time)
Private Function GetOffsetByTime( _
                 ByRef cObject As CSampleGrabberBuffer, _
                 ByVal dTime As Double) As Long
    Dim lOffset As Long
    
    With cObject
    
    lOffset = ((dTime - .dCurrentTime) * .tFormat.nAvgBytesPerSec) And (Not (.tFormat.nBlockAlign - 1))
    
    lOffset = (.lOffset + lOffset) Mod .lDataSize
    
    If lOffset < 0 Then lOffset = .lDataSize + lOffset
    
    End With
    
    GetOffsetByTime = lOffset
    
End Function

Private Function FARPROC( _
                 ByVal pfn As Long) As Long
    FARPROC = pfn
End Function