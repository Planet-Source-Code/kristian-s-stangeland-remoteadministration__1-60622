VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGdiPlus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' http://www.syix.com/wpsjr1/index.html
' Class for GDI+ Access
' Requires my gdi+.tlb

' Note: if you are unfamiliar with tlb (type library)
' They are files containing Declares, Enums and Constants (and can also contain interfaces and other data structures)
' Type Libraries are compiled into the exe, and do NOT need to be available externally.

' Feel free to use my tlb, with the stipulation that you will
' name your first born male child after me. :P

Event Error(ByVal lGdiError As Long, ByVal sErrorDesc As String)

Dim m_lToken As Long         ' Startup/Shutdown token
Dim tGuids() As CLSID       ' Array of GUIDs for codecs
Dim colCodecs As Collection   ' Search into codec index on file extension
Dim m_lNumCodecs As Long

Private Sub Class_Initialize()

    Dim gsi As GdiplusStartupInput
    Dim lError As Long
    
    gsi.GdiplusVersion = 1
    lError = GdiplusStartup(m_lToken, gsi, ByVal 0)
    
    If Not lError Then
        EnumEncoders
    Else
        RaiseEvent Error(lError, GdiErrorString(lError))
    End If

End Sub

Private Sub Class_Terminate()

    If m_lToken Then
        GdiplusShutdown m_lToken
    End If

End Sub

Private Function EnumEncoders() As Long

    Dim lNumEncoders As Long
    Dim lEncoderSize As Long
    Dim lError As Long
    Dim b() As Byte
    Dim i As Long
    Dim codecs() As ImageCodecInfo
    Dim aTemp As Variant
    Dim Tell As Long
    
    lError = GdipGetImageEncodersSize(lNumEncoders, lEncoderSize)
    
    If Not lError Then
    
        ReDim codecs(lNumEncoders - 1)
        ReDim b(lEncoderSize - 1)
        
        lError = GdipGetImageEncoders(lNumEncoders, lEncoderSize, b(0))
        
        If Not lError Then
        
            RtlMoveMemory codecs(0), b(0), lNumEncoders * LenB(codecs(0))
            
            ReDim tGuids(lNumEncoders - 1)
            m_lNumCodecs = lNumEncoders
            
            Set colCodecs = Nothing
            Set colCodecs = New Collection
            
            Do While lNumEncoders
            
                lNumEncoders = lNumEncoders - 1
                tGuids(lNumEncoders) = codecs(lNumEncoders).CLSID
                aTemp = Split(StringFromPointerW(codecs(lNumEncoders).pwszFilenameExtension), ";")
                
                For Tell = UBound(aTemp) To LBound(aTemp) Step -1
                    colCodecs.Add lNumEncoders, aTemp(Tell)
                Next
                
            Loop
        Else
            RaiseEvent Error(lError, GdiErrorString(lError))
        End If
    
    Else
        RaiseEvent Error(lError, GdiErrorString(lError))
    End If

End Function

' do not compare this to a boolean
'  returns -1 for not found, 0-positive GUID index for found
Private Function ExtensionExists(ByRef sKey As String) As Long

    On Error GoTo errorhandler
    ExtensionExists = True ' invalid index
    
    If Not colCodecs Is Nothing Then
        ExtensionExists = colCodecs.Item(sKey)
    End If
    
    Exit Function
errorhandler:
    ' exit silently

End Function

Private Function StringToGuid(ByRef sGuid As String) As CLSID

    CLSIDFromString sGuid, StringToGuid

End Function

' saves the contents of a picturebox to a file
' supports GIF/JPG/TIF/PNG and various others
Public Function PictureToFile(ByVal picHandle As Long, ByRef sFilename As String, Optional lQuality As Long = 85) As Long

    Dim sExtension As String
    Dim bitmap     As Long
    Dim lError     As Long
    Dim params     As EncoderParameters
    Dim lQual      As Long
    Dim lIndex     As Long
    Dim tguid      As CLSID
    
    sExtension = GetExtension(sFilename)
    
    lIndex = ExtensionExists("*." & sExtension)
    
    If lIndex > -1 Then
    
        lError = GdipCreateBitmapFromHBITMAP(picHandle, 0, bitmap)
        
        If Not lError Then
        
            If (Asc(sExtension) And Not 32) = vbKeyJ Then ' lazy JPEG/JPG/JPE/JFIF checking :P
                lQual = lQuality
                params.Count = 1
                params.Parameter.CLSID = StringToGuid(EncoderQuality)
                params.Parameter.NumberOfValues = 1
                params.Parameter.Type = EncoderParameterValueTypeLong
                params.Parameter.Value = VarPtr(lQual)
                lError = GdipSaveImageToFile(bitmap, sFilename, tGuids(lIndex), params)
            Else
                lError = GdipSaveImageToFile(bitmap, sFilename, tGuids(lIndex), ByVal 0)
            End If
            
            If Not lError Then
                PictureToFile = True ' tada!
            Else
                RaiseEvent Error(lError, GdiErrorString(lError))
            End If
        Else
            RaiseEvent Error(lError, GdiErrorString(lError))
        End If
    End If

End Function

Public Function GetExtension(sFile As String) As String

    On Error Resume Next
    GetExtension = Right(sFile, Len(sFile) - InStrRev(sFile, "."))

End Function

Private Function StringFromPointerW(ByVal lPointer As Long) As String

    Dim lLength As Long
    
    If lPointer Then
        lLength = lstrlenW(lPointer)
        StringFromPointerW = Space$(lLength)
        RtlMoveMemory ByVal StrPtr(StringFromPointerW), ByVal lPointer, lLength * 2
    End If

End Function

Public Function MakeScreenshot(sFileExtension As String, Optional Width As Long, Optional Height As Long, Optional srcWidth As Long, Optional srcHeight As Long) As String

    Dim hdc As Long, hBitmap As Long, destDC As Long, oldBitmap As Long, sFile As String
    
    If Width <= 0 Then
        Width = Screen.Width / Screen.TwipsPerPixelX
    End If
    
    If srcWidth <= 0 Then
        srcWidth = Width
    End If
    
    If Height <= 0 Then
        Height = Screen.Height / Screen.TwipsPerPixelY
    End If
    
    If srcHeight <= 0 Then
        srcHeight = Height
    End If
    
    destDC = GetDC(0)
    
    hdc = CreateCompatibleDC(destDC)
    hBitmap = CreateCompatibleBitmap(destDC, Width, Height)
    oldBitmap = SelectObject(hdc, hBitmap)
    
    StretchBlt hdc, 0, 0, Width, Height, destDC, 0, 0, srcWidth, srcHeight, vbSrcCopy
    
    sFile = Communication.Main.TempFile & "." & sFileExtension
    PictureToFile hBitmap, sFile
    
    ' Return file name
    MakeScreenshot = sFile
    
    ' Clean up
    SelectObject hdc, oldBitmap
    DeleteObject hBitmap
    DeleteDC hdc

End Function

Public Function GdiErrorString(ByVal lError As Status) As String

    Dim s As String
    
    Select Case lError
    Case GenericError:              s = "Generic Error"
    Case InvalidParameter:          s = "Invalid Parameter"
    Case OutOfMemory:               s = "Out Of Memory"
    Case ObjectBusy:                s = "Object Busy"
    Case InsufficientBuffer:        s = "Insufficient Buffer"
    Case NotImplemented:            s = "Not Implemented"
    Case Win32Error:                s = "Win32 Error"
    Case WrongState:                s = "Wrong State"
    Case Aborted:                   s = "Aborted"
    Case FileNotFound:              s = "File Not Found"
    Case ValueOverflow:             s = "Value Overflow"
    Case AccessDenied:              s = "Access Denied"
    Case UnknownImageFormat:        s = "Unknown Image Format"
    Case FontFamilyNotFound:        s = "FontFamily Not Found"
    Case FontStyleNotFound:         s = "FontStyle Not Found"
    Case NotTrueTypeFont:           s = "Not TrueType Font"
    Case UnsupportedGdiplusVersion: s = "Unsupported Gdiplus Version"
    Case GdiplusNotInitialized:     s = "Gdiplus Not Initialized"
    Case PropertyNotFound:          s = "Property Not Found"
    Case PropertyNotSupported:      s = "Property Not Supported"
    Case Else:                      s = "Unknown GDI+ Error"
    End Select
    
    GdiErrorString = s

End Function
