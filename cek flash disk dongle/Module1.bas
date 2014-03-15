Attribute VB_Name = "Module1"
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

Private Const SECURITY_CODE As String = "-eB03DVVsA5RFyvKh" 'ini bisa diganti

Private Function generateKeyByMD5(ByVal serialNumber As String) As String
    Dim objMD5  As clsMD5
    
    Set objMD5 = New clsMD5
    generateKeyByMD5 = objMD5.CalculateMD5(serialNumber)
    Set objMD5 = Nothing
End Function

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

Private Function dongleKeyFile(ByVal fileName As String) As String
    Dim fso As Scripting.FileSystemObject
    Dim ts  As Scripting.TextStream
    Dim tmp As String
    
    On Error GoTo errHandle
    
    If fileExists(fileName) Then
        Set fso = New Scripting.FileSystemObject
        Set ts = fso.OpenTextFile(fileName, ForReading, False)
        Do While Not ts.AtEndOfStream
            tmp = ts.ReadLine
            If Len(tmp) > 0 Then Exit Do
        Loop
        ts.Close
        Set ts = Nothing
        Set fso = Nothing
    End If
    
    dongleKeyFile = tmp
    
    Exit Function
errHandle:
    dongleKeyFile = ""
End Function

Private Function isValidDongle() As Boolean
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
    
    lDs = GetLogicalDrives
    
    For cnt = 0 To 25
        If (lDs And 2 ^ cnt) <> 0 Then
            driveName = Chr$(65 + cnt) & ":\"
            
            If GetDriveType(driveName) = 2 Then 'hanya flash disk yang kita proses
                dongleFile = driveName & "donglekey"
                
                strLabel = String$(255, Chr$(0))
                GetVolumeInformation driveName, strLabel, 255, serial, 0, 0, fSName, 255
                strLabel = Left$(strLabel, InStr(1, strLabel, Chr$(0)) - 1)
                
                GetVolumeInformation driveName, vbNullString, 255, serial, 0, 0, vbNullString, 255
                
                formatHex = Format(Hex(serial), "00000000")
                serialNumber = Left(formatHex, 4) & "-" & Right(formatHex, 4) 'serial number - plain text
                
                'serial number + security code yang sudah dienkripsi
                'security code -> harus sama dg yang di tool dongle
                generateKey = generateKeyByMD5(serialNumber & SECURITY_CODE)
                
                If generateKey = dongleKeyFile(dongleFile) Then
                    isValidDongle = True: Exit For
                End If
            End If
        End If
    Next cnt
End Function

Public Sub Main()
    If isValidDongle Then
        'TODO : tampilkan Form Utama disini
    Else
        MsgBox "Donglenya enggak valid atau belum dipasang.", vbExclamation, "Peringatan"
    End If
End Sub
