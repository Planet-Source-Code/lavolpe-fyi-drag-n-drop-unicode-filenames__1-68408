VERSION 5.00
Begin VB.Form frmDNDunicode 
   Caption         =   "Form1"
   ClientHeight    =   4200
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5925
   LinkTopic       =   "Form1"
   ScaleHeight     =   280
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   395
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picNoUnicodeSupport 
      Height          =   2640
      Left            =   2970
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   172
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   179
      TabIndex        =   1
      Top             =   510
      Width           =   2745
   End
   Begin VB.PictureBox picUnicodeSupported 
      Height          =   2640
      Left            =   105
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   172
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   179
      TabIndex        =   0
      Top             =   495
      Width           =   2745
   End
   Begin VB.Image Image1 
      Height          =   765
      Left            =   120
      Picture         =   "frmDNDunicode.frx":0000
      Top             =   3405
      Width           =   2325
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "And the above does not"
      Height          =   315
      Index           =   1
      Left            =   3450
      TabIndex        =   4
      Top             =   3180
      Width           =   2265
   End
   Begin VB.Label Label2 
      Caption         =   "Supports Unicode Characters. Example:"
      Height          =   315
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   3165
      Width           =   3030
   End
   Begin VB.Label Label1 
      Caption         =   "Drag and drop bitmap file names on the following boxes to test.  Or Copy bitmap file and Paste below (Ctrl+V)"
      Height          =   360
      Left            =   120
      TabIndex        =   2
      Top             =   90
      Width           =   5700
   End
End
Attribute VB_Name = "frmDNDunicode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Important information
' NOTES REGARDING THE TLB FILE..... Needed to ensure unciode filename support

' ^^ NO OBSOLETE.

' Using a new method of calling an interface's function without defining the
' interface in a TLB. Notes and references in the bas module's GetDroppedFiles

' SO WHAT DOES THIS PROJECT DO?  IT ALLOWS OLE DROP AND PASTE EVENTS OF
' FILE NAMES THAT CONTAIN UNICODE CHARACTERS. THAT'S IT; NOTHING MORE.
' I have also included some nice-to-have routines in the bas module
' that you may want to use with other projects.

' To test the limitations of VB there are two picture boxes on the form.
' One supports unicode characters and the other does not. Both support
' dropping of files and pasting of files.

' How do you get a unicode file name to test this with?  The easiest way I can think
' of, assuming you have NT-based system and/or XP, are the following steps:

' a. Copy any bitmap file to your C: drive
' b. Open Internet Explorer and navigate to a website like the following
'       http://topic.csdn.net/t/20060201/23/4538188.html#
' c. Then copy some of the non-ANSI characters (i.e., highlight & Ctrl+C)
' d. Right click on the bitmap you copied and choose Rename
' e. Paste the copied unicode characters as the bitmap name. Preserve the .bmp extension


' Good.  Run the project and drag the file on both picture boxes.

Private Declare Function DragQueryFile Lib "shell32.dll" Alias "DragQueryFileA" (ByVal hDrop As Long, ByVal UINT As Long, ByVal lpStr As String, ByVal ch As Long) As Long
Private Declare Function DragQueryFileW Lib "shell32.dll" (ByVal hDrop As Long, ByVal UINT As Long, ByVal lpStr As Long, ByVal ch As Long) As Long
Private Declare Function OpenClipboard Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function GetClipboardData Lib "user32.dll" (ByVal wFormat As Long) As Long
Private Declare Function CloseClipboard Lib "user32.dll" () As Long

' ///////////////////////////////////////////////////////////////////////////////////

'   U N I C O D E   S U P P O R T

' ///////////////////////////////////////////////////////////////////////////////////


Private Sub picUnicodeSupported_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyV Then
        If (Shift And vbCtrlMask) = vbCtrlMask Then
            
            Dim sFiles() As String
            If GetPastedFiles(sFiles) > 0 Then
            
                On Error Resume Next
                Set picUnicodeSupported.Picture = LoadPictureW(sFiles(1))
                If Err Then
                    MsgBox Err.Description, vbExclamation + vbOKOnly
                    Err.Clear
                End If
                
            End If
        End If
    End If
    
End Sub

Private Sub picUnicodeSupported_OLEDragOver(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    ' only accept files for this demo
    If data.GetFormat(vbCFFiles) = False Then Effect = vbDropEffectNone
End Sub

Private Sub picUnicodeSupported_OLEDragDrop(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Dim sFiles() As String
    Dim tPicture As StdPicture
    
    If GetDroppedFiles(data, sFiles) > 0 Then
    
        On Error Resume Next
        Set picUnicodeSupported = LoadPictureW(sFiles(1))
        If Err Then
            MsgBox Err.Description, vbExclamation + vbOKOnly
            Err.Clear
        End If
        
    End If
End Sub


' ///////////////////////////////////////////////////////////////////////////////////

'   N O N - U N I C O D E   S U P P O R T

' ///////////////////////////////////////////////////////////////////////////////////

Private Sub picNoUnicodeSupport_OLEDragOver(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    ' only accept files for this demo
    If data.GetFormat(vbCFFiles) = False Then Effect = vbDropEffectNone
End Sub

Private Sub picNoUnicodeSupport_OLEDragDrop(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error Resume Next
    If data.GetFormat(vbCFFiles) = True Then
        Set picNoUnicodeSupport.Picture = LoadPicture(data.Files(1))
        If Err Then
            MsgBox Err.Description, vbExclamation + vbOKOnly
            Err.Clear
        End If
    End If
End Sub

Private Sub picNoUnicodeSupport_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyV Then
        If (Shift And vbCtrlMask) = vbCtrlMask Then
            
            Dim hDrop As Long, lLen As Long, sFile As String
            
            If OpenClipboard(0) Then
                hDrop = GetClipboardData(vbCFFiles)
                If Not hDrop = 0 Then
                    
                    On Error Resume Next
                    lLen = DragQueryFile(hDrop, 0, vbNullString, 0) ' get length
                    sFile = String$(lLen, 0)                        ' build buffer
                    DragQueryFile hDrop, 0, sFile, lLen + 1         ' transfer to buffer
                    
                    Set picNoUnicodeSupport.Picture = LoadPicture(sFile)
                    If Err Then
                        MsgBox Err.Description, vbExclamation + vbOKOnly
                        Err.Clear
                    End If
                    
                    ' Note that if the filename has unicode characters then we could have
                    ' used DragQueryFileW version to get the Unicode vs ANSI file name
                    ' but that still won't make the filename compatible 'cause
                    ' LoadPicture cannot use unicode characters
                    
                End If
                CloseClipboard
            End If
        End If
    End If
    
End Sub

