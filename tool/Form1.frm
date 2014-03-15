VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Demo membuat tool dongle"
   ClientHeight    =   5985
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   7950
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstDrive 
      Height          =   3960
      Left            =   120
      TabIndex        =   3
      Top             =   1455
      Width           =   7695
   End
   Begin VB.CommandButton cmdCreateDongleKey 
      Caption         =   "Create Dongle Key"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   5535
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   " [ Info ] "
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7695
      Begin VB.Label Label1 
         Height          =   885
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   7455
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'***************************************************************************
' MMMM  MMMMM  OMMM   MMMO    OMMM    OMMM    OMMMMO     OMMMMO    OMMMMO  '
'  MM    MM   MM MM    MMMO  OMMM    MM MM    MM   MO   OM    MO  OM    MO '
'  MM  MM    MM  MM    MM  OO  MM   MM  MM    MM   MO   OM    MO       OMO '
'  MMMM     MMMMMMMM   MM  MM  MM  MMMMMMMM   MMMMMO     OMMMMO      OMO   '
'  MM  MM        MM    MM      MM       MM    MM   MO   OM    MO   OMO     '
'  MM    MM      MM    MM      MM       MM    MM    MO  OM    MO  OM   MM  '
' MMMM  MMMM    MMMM  MMMM    MMMM     MMMM  MMMM  MMMM  OMMMMO   MMMMMMM  '
'                                                                          '
' K4m4r82's Laboratory                                                     '
' http://coding4ever.wordpress.com                                         '
'***************************************************************************

Private Declare Function GetLogicalDrives Lib "kernel32" () As Long
Private Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long

Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Const SYNCHRONIZE       As Long = &H100000
Private Const INFINITE          As Long = &HFFFF

Private Const SECURITY_CODE     As String = "-eB03DVVsA5RFyvKh" 'ini bisa diganti

Private Sub writeDongleFile(ByVal fileName As String, ByVal key As String)
    Dim fso As Scripting.FileSystemObject
    Dim ts  As Scripting.TextStream

    Set fso = New Scripting.FileSystemObject
    Set ts = fso.OpenTextFile(fileName, ForWriting, True)
    ts.Write key & vbCrLf
    ts.Close
    Set ts = Nothing
    Set fso = Nothing
End Sub

Private Function fileExists(ByVal namaFile As String) As Boolean
    Dim fso As Scripting.FileSystemObject
    
    On Error GoTo errHandle
    
    If Not (Len(namaFile) > 0) Then fileExists = False: Exit Function
    
    Set fso = New Scripting.FileSystemObject
    fileExists = fso.fileExists(namaFile)
    Set fso = Nothing
    
    Exit Function
errHandle:
    fileExists = False
End Function

Private Function generateKeyByMD5(ByVal serialNumber As String) As String
    Dim objMD5  As clsMD5
    
    Set objMD5 = New clsMD5
    generateKeyByMD5 = objMD5.CalculateMD5(serialNumber)
    Set objMD5 = Nothing
End Function

Private Sub loadDrive(ByVal lst As ListBox)
    Dim lDs             As Long
    Dim cnt             As Long
    Dim serial          As Long
    
    Dim strLabel        As String
    Dim fSName          As String
    Dim formatHex       As String
    Dim driveName       As String
    Dim serialNumber    As String
    Dim generateKey     As String
    Dim dongleFile      As String
    Dim cmd             As String
    
    Dim shellX          As Long
    Dim lPid            As Long
    Dim lHnd            As Long
    Dim lRet            As Long
    
    'get the available drives
    lDs = GetLogicalDrives
    lst.Clear
                
    For cnt = 0 To 25
        If (lDs And 2 ^ cnt) <> 0 Then
            driveName = Chr$(65 + cnt) & ":\"
                            
            'Drive Type :
            '***************
            '2 = Removable/flash disk
            '3 = Drive Fixed
            '4 = Remote
            '5 = Cd-Rom
            '6 = Ram disk
            
            If GetDriveType(driveName) = 2 Then 'hanya flash disk yang kita proses
                dongleFile = driveName & "donglekey"
                
                If fileExists(dongleFile) Then 'sudah ada file dongle
                    'tampilkan file donglekey sebelumnya
                    'kalo tidak akan terjadi error waktu menjalankan perintah kill
                    cmd = "attrib -s -h " & dongleFile
                    
                    shellX = Shell(cmd, vbHide)
                    lPid = shellX
                    If lPid <> 0 Then
                        lHnd = OpenProcess(SYNCHRONIZE, 0, lPid)
                        If lHnd <> 0 Then
                            lRet = WaitForSingleObject(lHnd, INFINITE)
                            CloseHandle (lHnd)
                        End If
                    End If
                    
                    'hapus file dongle sebelumnya
                    'kalo tidak akan terjadi error waktu menulis ulang file dongle
                    'kenapa error,karena file dongle disembunyikan dg attribut +s -> dianggap file system
                    Kill dongleFile
                End If
                
                strLabel = String$(255, Chr$(0))
                GetVolumeInformation driveName, strLabel, 255, serial, 0, 0, fSName, 255
                strLabel = Left$(strLabel, InStr(1, strLabel, Chr$(0)) - 1)
                
                GetVolumeInformation driveName, vbNullString, 255, serial, 0, 0, vbNullString, 255
                
                formatHex = Format(Hex(serial), "00000000")
                serialNumber = Left(formatHex, 4) & "-" & Right(formatHex, 4) 'serial number - plain text
                                                
                generateKey = generateKeyByMD5(serialNumber & SECURITY_CODE) 'serial number + security code yang sudah dienkripsi
                
                Call writeDongleFile(dongleFile, generateKey) 'tulis file dongle ke flash disk
                DoEvents
                Call Shell("attrib +s +h " & dongleFile) 'sembunyikan file dongle
                
                lst.AddItem strLabel & "(" & Chr$(65 + cnt) & ":" & ") -> Serial Number : " & serialNumber & " -> Generate Key : " & generateKey
            End If
        End If
    Next cnt
    
    If Not (lst.ListCount > 0) Then lst.AddItem ">> Belom ada flash disk yang di coloxin <<"
End Sub

Private Sub cmdCreateDongleKey_Click()
    Call loadDrive(lstDrive)
End Sub


Private Sub Form_Load()
    Label1.Caption = "Sebelum mengklik tombol 'Create Dongle Key', pasang terlebih dahulu flash disk dongle, dan pastikan semua flash disk sudah tampil di Windows Explorer." & vbCrLf & vbCrLf & _
                     "File donglekey otomatis tersimpan di flash disk dan disembunyikan."
End Sub


