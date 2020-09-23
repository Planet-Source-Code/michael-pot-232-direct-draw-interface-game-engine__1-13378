VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DDrawInterface"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'**********************************************************
'       Title: Direct Draw Interface (DDI)
'Project File: DDrawInterface
'      Author: Michael Pote (Michaelpote@worldonline.co.za)
'
' Description: Graphics Engine for quick access to direct draw,
'              by loading DDI files (Made with DDIFileMaker).
'              Also supports wave file loading and playing
'              via Direct Sound.
'
'    Requires: DX7VB.dll (class wrapper for the DirectX API.)
'**********************************************************


Private DX As New DirectX7
Private DD As DirectDraw7
Private DS As DirectSound
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public MainForm As Form

Private Type DDIImageBank
Top As Long
Width As Long
Height As Long
AnimCount As Long
End Type

Private ScreenRect As RECT

Private Type DDIFileType
Surf As DirectDrawSurface7
DDSD As DDSURFACEDESC2
Transparent As Boolean
ImageCount As Long
DDIImages() As DDIImageBank
End Type

Public DDICount As Long
Public SndCount As Long
Private SndFile() As DirectSoundBuffer
Private DDIFile() As DDIFileType
Dim Pddsd As DDSURFACEDESC2, ScrDdsd As DDSURFACEDESC2
Dim Primary As DirectDrawSurface7
Public Backbuffer As DirectDrawSurface7
Dim Ending As Boolean, bRestore As Boolean

Private Sub WriteINI(Path As String, Section As String, Nam As String, Vaule As String)
Dim V As String
V = Vaule
WritePrivateProfileString Section, Nam, V, Path
DoEvents
End Sub
Private Function ReadINI(Path As String, Section As String, Nam As String) As String
Static r As String * 200
r = ""
GetPrivateProfileString Section, Nam, "Error Reading INI", r, 200, Path
ReadINI = Trim(r)
If Asc(Right(ReadINI, 1)) = 0 Then
ReadINI = Mid(ReadINI, 1, Len(ReadINI) - 1)
End If
End Function

'Loads a wave file into an array of directsound buffers.
'Returns index of the newly created sound file in the array
'*************************************************
Public Function LoadSoundFile(WaveFile As String) As Integer
SndCount = SndCount + 1
ReDim Preserve SndFile(1 To SndCount) As DirectSoundBuffer
LoadWave WaveFile, SndFile(SndCount)
LoadSoundFile = SndCount
End Function

Public Sub PlaySound(SndIndex As Long, Optional Volume As Long = 0, Optional Frequency As Long = 44100, Optional Pan As Long = 0, Optional Looping As Boolean)
With SndFile(SndIndex)
.SetVolume Volume
.SetPan Pan
.SetFrequency Frequency
If Looping Then .Play DSBPLAY_LOOPING Else .Play DSBPLAY_DEFAULT
End With
End Sub

Public Sub StopSoundLooping(SndIndex As Long)
SndFile(SndIndex).Stop
End Sub

'Draws a DDI file from the array onto the screen.
'************************************************
Public Sub DrawDDIFile(DDIIndex As Long, X As Long, Y As Long, ImageBankIndex As Long, Optional AnimFrame As Long = 0)
Dim TotHeight As Long
With DDIFile(DDIIndex).DDIImages(ImageBankIndex)
If AnimFrame > .AnimCount Then AnimFrame = .AnimCount
BltFast X, Y, AnimFrame * .Width, CSng(.Top), CSng(.Width), CSng(.Height), DDIFile(DDIIndex).Surf, DDIFile(DDIIndex).Transparent
End With
End Sub

'Private routine to blit things to the backbuffer easily
'*******************************************************
Private Sub BltFast(DX As Long, DY As Long, SrcX As Single, SrcY As Single, SrcWid As Single, SrcHgt As Single, ByRef Surf As DirectDrawSurface7, ColourKey As Boolean)
On Local Error GoTo Errot
Dim SrcRect As RECT, Retval, TestRect As RECT

With SrcRect

If DY < 0 Then
.Top = SrcY + (-DY)
DY = 1
Else
.Top = SrcY
End If

If DX < 0 Then
.Left = SrcX + (-DX)
DX = 1
Else
.Left = SrcX
End If

If DY + SrcHgt > 600 Then
.Bottom = SrcY + SrcHgt - ((DY + SrcHgt) - 600)
Else
.Bottom = SrcY + SrcHgt
End If

If DX + SrcWid > 800 Then
.Right = SrcX + SrcWid - ((DX + SrcWid) - 800)
Else
.Right = SrcX + SrcWid
End If

End With
If ColourKey Then
Retval = Backbuffer.BltFast(DX, DY, Surf, SrcRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
Else
Retval = Backbuffer.BltFast(DX, DY, Surf, SrcRect, DDBLTFAST_WAIT)
End If
Exit Sub
Errot:
EndIt
End Sub

'This routine loads a DDI file and it's DDG counterpart in
'to a array of DDIfiles and directdraw surfaces respectivly.
'It Returns the index of the newly loaded file.
'******************************************************
Public Function LoadDDIFile(Filename As String) As Integer
DDICount = DDICount + 1
Dim T As String, I As Long, ColourKey As Boolean, TransColour As Long, Width As Long, Height As Long, BitmapName As String

ReDim Preserve DDIFile(1 To DDICount) As DDIFileType
With DDIFile(DDICount)
T = ReadINI(Filename, "General", "Transparent")
If T = "1" Then .Transparent = True Else .Transparent = False
TransColour = Val(ReadINI(Filename, "General", "TransColour"))
.ImageCount = Val(ReadINI(Filename, "General", "Count"))
Height = Val(ReadINI(Filename, "General", "Height"))
Width = Val(ReadINI(Filename, "General", "Width"))

ReDim .DDIImages(1 To .ImageCount) As DDIImageBank
For I = 1 To .ImageCount
.DDIImages(I).AnimCount = Val(ReadINI(Filename, "DDI" & I, "AnimCount"))
.DDIImages(I).Height = Val(ReadINI(Filename, "DDI" & I, "Height"))
.DDIImages(I).Width = Val(ReadINI(Filename, "DDI" & I, "Width"))
.DDIImages(I).Top = Val(ReadINI(Filename, "DDI" & I, "Top"))
Next

BitmapName = Mid(Filename, 1, Len(Filename) - 3) & "DDG"

CreateSurface DDICount, BitmapName, Width, Height, .Transparent, TransColour
LoadDDIFile = DDICount
End With
End Function


'Private Function to create a direct draw surface from a
'file easily.
'*******************************************************
Private Function CreateSurface(DDIIndex As Long, Filename As String, Wid As Long, Hgt As Long, ColourKey As Boolean, TransColour As Long) As Long
On Local Error GoTo SIE
With DDIFile(DDIIndex)
Set .Surf = Nothing                                   'Clear the surface
.DDSD.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
.DDSD.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
.DDSD.lWidth = Wid
.DDSD.lHeight = Hgt
Set .Surf = DD.CreateSurfaceFromFile(Filename, .DDSD) 'Load the bitmap
If ColourKey = True Then
Dim key As DDCOLORKEY
'key.low = TransColour
'key.high = TransColour
key.low = 0
key.high = 0
.Surf.SetColorKey DDCKEY_SRCBLT, key
End If
End With
Exit Function
SIE:
EndIt
End Function

'Flips the backbuffer to expose it to the screen
'*****************************************
Public Sub MainLoopFlip()
Primary.Flip Nothing, DDFLIP_WAIT

bRestore = False
Do Until ExModeActive
DoEvents
bRestore = True
Loop
DoEvents
If bRestore Then
bRestore = False
DD.RestoreAllSurfaces
'InitSurfaces
End If

End Sub
'Clears the screen in a colour.
'direct draw colours seem to be differently placed
'from normal RGB colours though.
'************************************************
Public Sub ClearScreen(Colour As Long)
Backbuffer.BltColorFill ScreenRect, Colour
End Sub

'Sets-up direct draw to handle drawing.
'Bpp - bits per pixel (8-bit (256 colours), 16-bit, 24-bit, 32-bit)
'**********************************************
Public Function Initialize(ResolutionX As Long, ResolutionY As Long, Bpp As Long) As Boolean
On Local Error GoTo ErrorOut
Set DD = DX.DirectDrawCreate("")
Set DS = DX.DirectSoundCreate("")

MainForm.Show
DS.SetCooperativeLevel MainForm.hWnd, DSSCL_NORMAL
DD.SetCooperativeLevel MainForm.hWnd, DDSCL_FULLSCREEN Or DDSCL_ALLOWMODEX Or DDSCL_EXCLUSIVE
DD.SetDisplayMode ResolutionX, ResolutionY, Bpp, 0, DDSDM_DEFAULT

ScreenRect.Top = 0
ScreenRect.Left = 0
ScreenRect.Right = ResolutionX
ScreenRect.Bottom = ResolutionY

     
     Pddsd.lFlags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT
     Pddsd.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE Or DDSCAPS_FLIP Or DDSCAPS_COMPLEX
     Pddsd.lBackBufferCount = 1
     DoEvents
     Set Primary = DD.CreateSurface(Pddsd)
     Dim Caps As DDSCAPS2
     Caps.lCaps = DDSCAPS_BACKBUFFER
     Set Backbuffer = Primary.GetAttachedSurface(Caps)
     Backbuffer.GetSurfaceDesc ScrDdsd
     Backbuffer.SetFontTransparency True
     Backbuffer.SetForeColor vbWhite


Initialize = True
Exit Function
ErrorOut:
EndIt
End Function

Public Sub EndIt()
Ending = True
DD.RestoreDisplayMode
DD.SetCooperativeLevel MainForm.hWnd, DDSCL_NORMAL
If Err.Description <> "" Then MsgBox Err.Description, vbCritical
DoEvents
Unload MainForm

End Sub

'Private routine to test if direct draw is stable.
'**************************************************
Private Function ExModeActive() As Boolean
     Dim TestCoopRes As Long 'holds the return value of the test.

     TestCoopRes = DD.TestCooperativeLevel 'Tells DDraw to do the test

     If (TestCoopRes = DD_OK) Then
         ExModeActive = True 'everything is fine
     Else
         ExModeActive = False 'this computer doesn't support this mode
     End If
End Function

'Loads a wave file into a direct sound buffer
'****************************************
Private Sub LoadWave(File As String, ByRef Buffer As DirectSoundBuffer)

    Dim bufferDesc As DSBUFFERDESC
    Dim waveFormat As WAVEFORMATEX
    Set Buffer = Nothing
    
    bufferDesc.lFlags = DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME Or DSBCAPS_STATIC
    waveFormat.nFormatTag = WAVE_FORMAT_PCM
    waveFormat.nChannels = 2    '
    waveFormat.lSamplesPerSec = 44100
    waveFormat.nBitsPerSample = 16
    waveFormat.nBlockAlign = waveFormat.nBitsPerSample / 8 * waveFormat.nChannels
    waveFormat.lAvgBytesPerSec = waveFormat.lSamplesPerSec * waveFormat.nBlockAlign
    Set Buffer = DS.CreateSoundBufferFromFile(File, bufferDesc, waveFormat)
    'checks for any errors
    If Err.Number <> 0 Then
        MsgBox "Error with Sound!: Unable to find " + File, vbCritical
    EndIt
    End If

End Sub


