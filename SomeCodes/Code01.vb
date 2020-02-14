' // DirectShow audio visualizer using Sample Grabber filter
' // by The trick 2018
'
' // This example inserts the sample grabber filter between a playback device and wave-PCM data.
' // During playing it grabs the audio samples (L + R) / 2 and shows them on the picturebox.
' // It uses the different grabbing techniques in the compiled form and IDE.
' // Example has a callback object that receives the samples in rendering thread in the compiled form.
' // It uses ISampleGrabber::GetCurrentBuffer method to grab the samples in IDE (poor quality).
'
' // You can install the additional DS filters (for example LAVFilters https://github.com/nevcairiel/lavfilters)
' // to support additional formats.
'
' // If a media has the black window it probably means the graph uses DirectVobSub filter (subtitles) you can
' // rebuild the graph to connect a video output from decoder to video input of DirectVobSub filter.
' // This behavior isn't implemented in this example.

Option Explicit

' // Slider states
Private Enum eSliderAction
    SA_UPDATE   ' // During playing time to update slider position according to playback time
    SA_MOVE     ' // User moves the slider to set playback position
End Enum

' // Since it uses ISampleGrabber::GetCurrentBuffer method in IDE we can't know the actual samples time.
' // An audio-renderer device has the buffers with some latency (on my PC about 2.5 sec.), we can change this
' // constant to compensate one. Notice, we'll have the delay between playing time and updation.
Private Const IDE_DELAY_TIME        As Double = 2.5

Dim mcGraphMgr      As FilgraphManager      ' // Graph manager
Dim mcPosition      As IMediaPosition       ' // Playback position control
Dim mcCallback      As ISampleGrabberBuffer ' // User callback grabber
Dim miBuffer()      As Integer              ' // Buffer with 16-bit, stereo, 44100Hz format
Dim mtFormat        As WAVEFORMATEX         ' // Format of the needed samples
Dim mcGrabberNative As ISampleGrabber       ' // Grabber native interface
Dim mcGrabber       As IFilterInfo          ' // Grabber quartz interface
Dim meSliderAction  As eSliderAction        ' // Current slider action
Dim mbHasAudio      As Boolean              ' // A media has audio stream
Dim mbHasVideo      As Boolean              ' // A media has video stream
Dim mcEvents        As IMediaEventEx        ' // Event interface (to track the EOF)
Dim mlPalette(255)  As Long                 ' // Palette of waveform samples amplitudes

Private Sub Form_Load()
    Dim lIndex  As Long
    
    ' // Set up palette
    For lIndex = 0 To 255
        mlPalette(lIndex) = ColorHLSToRGB(80 - lIndex / 255 * 80, 120, 240)
    Next
    
    ' // Set up the audio format
    With mtFormat
    
    .nChannels = 2
    .nSamplesPerSec = 44100
    .wBitsPerSample = 16
    .wFormatTag = 1
    .nBlockAlign = (.wBitsPerSample \ 8) * .nChannels
    .nAvgBytesPerSec = .nSamplesPerSec * .nBlockAlign
    
    ' // Create 4 sec buffer. Callback object will write to it and we'll read data from timer event
    ReDim miBuffer((.nAvgBytesPerSec \ 2) * 4 - 1)
    
    End With
    
End Sub

' // Open the media
Private Sub OpenFile( _
            ByRef sFileName As String)
    Dim cInAudioRenderer    As IPinInfo         ' // Input pin of audio-renderer
    Dim cOutAudioFilter     As IPinInfo         ' // Output pin of PCM provider
    Dim tMediaType          As AM_MEDIA_TYPE    ' // Media type for grabber
    Dim cWindow             As IVideoWindow     ' // Window control
    Dim bIsInIDE            As Boolean          ' // IDE flag
    
    On Error GoTo error_handler
    
    ' // Disable controls
    tmrUpdation.Enabled = False
    sldPosition.Enabled = False
    
    mbHasAudio = False
    mbHasVideo = False
    
    Debug.Assert MakeTrue(bIsInIDE)
    
    Set mcGraphMgr = New FilgraphManager
    Set mcPosition = mcGraphMgr

    ' // Add filters
    mcGraphMgr.RenderFile sFileName

    ' // Search for input pin of the audio renderer filter
    Set cInAudioRenderer = FindAudioInPin()
    
    If Not cInAudioRenderer Is Nothing Then
        
        mbHasAudio = True
        
        ' // Get PCM provider pin
        Set cOutAudioFilter = cInAudioRenderer.ConnectedTo
        
        ' // Add sample grabber
        Set mcGrabber = GetSampleGrabber()
        
        ' // Set media type
        Set mcGrabberNative = mcGrabber.Filter
        
        GUIDFromString StrPtr(WMMEDIATYPE_Audio), tMediaType.majortype
        GUIDFromString StrPtr(FORMAT_WaveFormatEx), tMediaType.formattype
        
        tMediaType.bFixedSizeSamples = 1
        tMediaType.lSampleSize = mtFormat.nBlockAlign
        tMediaType.cbFormat = Len(mtFormat)
        tMediaType.pbFormat = VarPtr(mtFormat)
        
        mcGrabberNative.SetMediaType tMediaType
        mcGrabberNative.SetBufferSamples 1
        
        ' // Disconnect audio renderer input
        cInAudioRenderer.Disconnect
        
        ' // Insert grabber between previous filter and renderer
        cOutAudioFilter.ConnectDirect GetPinByDirection(mcGrabber, 0)
        cInAudioRenderer.ConnectDirect GetPinByDirection(mcGrabber, 1)
        
        ' // Create callback
        Set mcCallback = CreateSampleGrabberCB(VarPtr(miBuffer(0)), (UBound(miBuffer) + 1) * 2, mtFormat)
        
        If Not bIsInIDE Then
            ' // In compiled form use BufferCB method
            mcGrabberNative.SetCallback mcCallback, 1
        End If
        
    End If
    
    ' // Check if a Video renderer exists in the graph
    Set cWindow = mcGraphMgr
    
    If IsVideoRendererExist(cWindow) Then
        
        mbHasVideo = True
        
        cWindow.Owner = picFrame.hwnd
        cWindow.MessageDrain = picFrame.hwnd
        cWindow.WindowStyle = &H56010000
        cWindow.Left = 0
        cWindow.Top = 0
        cWindow.Width = picFrame.ScaleWidth
        cWindow.Height = picFrame.ScaleHeight
        
    End If
    
    Set mcEvents = mcGraphMgr

    mcGraphMgr.Run
    
    If mcPosition.Duration >= 1 Then
        sldPosition.Max = mcPosition.Duration
        sldPosition.Enabled = True
    Else
        sldPosition.Max = 1
        sldPosition.Enabled = False
    End If
    
    sldPosition.Value = 0

    tmrUpdation.Enabled = True
    
    Exit Sub
    
error_handler:
    
    MsgBox "An error occurs 0x" & Hex$(Err.Number)
    
End Sub

' // Search for input pin of audio renderer filter
Private Function FindAudioInPin() As IPinInfo
    Dim cFilter     As IFilterInfo
    Dim cPin        As IPinInfo
    Dim cPin2       As IPinInfo
    Dim cMediaInfo  As Object
    
    On Error Resume Next
    
    ' // Find the first audio pin
    For Each cFilter In mcGraphMgr.FilterCollection
        
        For Each cPin In cFilter.Pins
            
            Set cMediaInfo = cPin.ConnectionMediaType
            
            If cMediaInfo.Type = WMMEDIATYPE_Audio Then
                
                ' // Search for last input pin in chain
                Do
                    
                    Set cPin2 = GetPinByDirection(cPin.FilterInfo, 1)
                    
                    If cPin2 Is Nothing Then
                        
                        If cPin.Direction = 1 Then
                            Set cPin = GetPinByDirection(cPin.FilterInfo, 0)
                        End If
                        
                        Set FindAudioInPin = cPin
                        
                        Exit Function
                        
                    End If
                    
                    Set cPin = cPin2.ConnectedTo
                    
                Loop
                
            End If
            
        Next
        
    Next

End Function

' // Get pin by its direction
Private Function GetPinByDirection( _
                 ByRef cFilterInfo As IFilterInfo, _
                 ByRef lDirection As Long) As IPinInfo
    Dim cPin    As IPinInfo

    For Each cPin In cFilterInfo.Pins
        If cPin.Direction = lDirection Then
            Set GetPinByDirection = cPin
            Exit Function
        End If
    Next
  
End Function

' // Add the sample grabber and get the object reference to it
Private Function GetSampleGrabber() As IUnknown
    Dim cRegFilter  As IRegFilterInfo

    For Each cRegFilter In mcGraphMgr.RegFilterCollection

        If cRegFilter.Name = "SampleGrabber" Then
            
            cRegFilter.Filter GetSampleGrabber
            Exit Function
            
        End If

    Next

End Function

Private Sub Form_Unload(Cancel As Integer)
    
    If Not mcGraphMgr Is Nothing Then
        mcGraphMgr.Stop
    End If
    
    ' // Stop callback
    If Not mcGrabberNative Is Nothing Then
        mcGrabberNative.SetCallback Nothing, 1
    End If
    
    Set mcGrabberNative = Nothing
    Set mcCallback = Nothing
    Set mcGraphMgr = Nothing
    
End Sub

Private Sub mnuOpen_Click()
    Dim sFileName   As String
    
    sFileName = GetOpenFile(Me.hwnd, "Open media", "All files" & vbNullChar & "*.*" & vbNullChar)
    
    If Len(sFileName) = 0 Then Exit Sub
    
    OpenFile sFileName
    
End Sub

Private Sub sldPosition_Change()
    
    If meSliderAction = SA_UPDATE Then Exit Sub
    
    mcPosition.CurrentPosition = sldPosition
    
    ' // Reset offset time
    If mbHasAudio Then
        mcCallback.Reset
    End If
    
    meSliderAction = SA_UPDATE
    
End Sub

' // Start user action
Private Sub sldPosition_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    meSliderAction = SA_MOVE
End Sub

' // Check if the media has a video renderer
Private Function IsVideoRendererExist( _
                 ByVal cVideoWindow As IVideoWindow) As Boolean
    Dim lUnused As Long
    
    On Error GoTo error_handler
    
    ' // Access to any method
    ' // From MSDN:
    ' // For the Filter Graph Manager's implementation, if the graph does not contain a video renderer
    ' // filter, all methods return E_NOINTERFACE.
    
    lUnused = cVideoWindow.BorderColor
    
    IsVideoRendererExist = True
    
error_handler:
    
End Function

' // Update audio waveform
Private Sub tmrUpdation_Timer()
    Dim lSize       As Long
    Dim lIndex      As Long
    Dim lSampleIdx  As Long
    Dim fSample     As Single
    Dim lPixels     As Long
    Dim mlReadPos   As Long
    Dim bIsInIDE    As Boolean
    Dim dTime       As Double
      
    ' // Check if the media has been completed
    If IsComplete() Then
        
        mcGraphMgr.Stop
        
        tmrUpdation.Enabled = False
        sldPosition.Enabled = False
    
    End If
    
    Debug.Assert MakeTrue(bIsInIDE)
    
    ' // Get current playback position
    dTime = mcPosition.CurrentPosition
    
    If mbHasAudio Then
        
        If bIsInIDE Then
            ' // Fill buffer using GetCurrentBuffer method with IDE offset
            FillCurrentBufferIDE dTime
        End If
        
        lSize = (UBound(miBuffer) + 1)
        
        ' // Get the offset from begin of the buffer according to time
        mlReadPos = mcCallback.GetOffsetByTime(dTime) \ 2
        
        picWaveform.Cls
        
        ' // Draw the grid
        For lIndex = 1 To 5
            
            picWaveform.Line (0, picWaveform.ScaleHeight / 2 + lIndex * (picWaveform.ScaleHeight / 2 / 5))- _
                            Step(picWaveform.ScaleWidth, 0), &H101010
            picWaveform.Line (0, picWaveform.ScaleHeight / 2 - lIndex * (picWaveform.ScaleHeight / 2 / 5))- _
                            Step(picWaveform.ScaleWidth, 0), &H101010
        Next
        
        For lIndex = 0 To picWaveform.ScaleWidth Step 10
            picWaveform.Line (lIndex, 0)-Step(0, picWaveform.ScaleHeight), &H151515
        Next
        
        For lIndex = 0 To picWaveform.ScaleWidth - 1
                
            ' // Calcualte sample index
            lSampleIdx = (lIndex * 2 + mlReadPos) Mod lSize
            
            ' // Mix the left and the right sample and cast it to 0..1
            fSample = (CLng(miBuffer(lSampleIdx)) + miBuffer(lSampleIdx + 1)) / 65536
            
            ' // Convert to pixels offset
            lPixels = fSample * picWaveform.ScaleHeight / 2 + picWaveform.ScaleHeight / 2
            
            miBuffer(lSampleIdx) = 0
            miBuffer(lSampleIdx + 1) = 0
           
'            ' // Line waveform
'            If lIndex Then
'                picWaveform.Line -(lIndex, lPixels), mlPalette(Abs(fSample) * 255)
'            Else
'                picWaveform.PSet (lIndex, lPixels)
'            End If
            
            picWaveform.Line (lIndex, picWaveform.ScaleHeight / 2)-(lIndex, lPixels), mlPalette(Abs(fSample) * 255)
            
        Next
        
    End If
    
    If meSliderAction = SA_UPDATE Then
        sldPosition.Value = dTime
    End If
    
End Sub

' // Check if media is complete
Private Function IsComplete() As Boolean
    Dim lEvent  As Long
    Dim lParam1 As Long
    Dim lParam2 As Long
    
    On Error GoTo error_handler
    
    ' // Search for EC_COMPLETE event
    mcEvents.GetEvent lEvent, lParam1, lParam2, 1
    
    mcEvents.FreeEventParams lEvent, lParam1, lParam2
    
    IsComplete = lEvent = EC_COMPLETE
    
error_handler:
    
End Function

' // Fill buffer using GetCurrentBuffer method
Private Sub FillCurrentBufferIDE( _
            ByVal dTime As Double)
    Dim lSize       As Long
    Dim lWritten    As Long
    Dim bData()     As Byte

    On Error Resume Next
    
    Do
        
        Err.Clear
        
        mcGrabberNative.GetCurrentBuffer lSize, ByVal 0&
        
        ReDim bData(lSize - 1)
        
        mcGrabberNative.GetCurrentBuffer lSize, bData(0)
        
        ' // If lSize too small (if between ReDim it increases buffer) continue obtaining
    Loop While Err.Number = 7
    
    ' // Write to buffer using IDE offset
    mcCallback.BufferCB dTime - IDE_DELAY_TIME, bData(0), lSize
    
End Sub