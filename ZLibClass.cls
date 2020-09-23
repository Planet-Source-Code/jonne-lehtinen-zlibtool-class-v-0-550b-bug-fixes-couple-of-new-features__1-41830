VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ZLibClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************
'*     Zlibtool Class Example by      *
'*          Jonne Lehtinen            *
'*          Version 0.550b            *
'* Purpose:                           *
'*    To demonstrate packing of many  *
'*    files in one file with zlib     *
'*                                    *
'* Requires:                          *
'*    Zlibtool component              *
'*                                    *
'* Contact:                           *
'*    Email:                          *
'*       Eclipser@vorlonempire.com    *
'*                                    *
'* Author note:                       *
'*    Ok, so the algorithm might not  *
'*    be the best possible, but       *
'*    atleast it works, even though   *
'*    there might still be bugs       *
'*    and there's couple of functions *
'*    missing. For example, the class *
'*    can't extract single file from  *
'*    compressed file.                *
'*    I won 't take any               *
'*    responsibility on what may      *
'*    happen when using this class.   *
'*    You may report me the bugs and  *
'*    I'll try to fix them if I find  *
'*    the time to do this.            *
'*    Report by E-Mail, query in IRC  *
'*    (Eclipser @ IRCnet, Efnet,      *
'*    FreeNode or Aniverse) or via    *
'*    ICQ (24147962).                 *
'*    Feel free to modify the class   *
'*    as you'd like it as long as     *
'*    you don't remove this text.     *
'*    This class is free of charge    *
'*    and every VB coder may use it   *
'*    freely in their projects even   *
'*    if they want money from the     *
'*    program they're making          *
'*    (crediting me would be nice).   *
'*    Also, if someone's wants to     *
'*    translate this code in other    *
'*    programming languages, it must  *
'*    be open source.                 *
'*    There might be some other issues*
'*    too, I just don't know them ^_^.*
'*    Respect the Copyright.          *
'**************************************

Option Explicit

Private Type FileData
    path As String
    filename As String
    bytes As Long
End Type

Private Const TempPathBuf As Long = 250
Private Const Header As String = "T4SSuS0f7zl1b"    ' This you may change to whatever you like, it's used to check if the file is compressed with this lib

Private pFiles As Long
Private File() As FileData
Private sTempPath As String
Private bAbort As Boolean
Private ReadBuf As Long ' The higher this is the less there's HD time for other programs to run, so don't set it too high

Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" _
(ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

'*********Class_Terminate()***********
'*         Terminate Event           *
'*                                   *
'*    Raised when set to nothing     *
'*                                   *
'* Uninitializes everything possible *
'*************************************

Private Sub Class_Terminate()
    pFiles = -1             ' Set pFiles to -1 to represent that there's no files
    Erase File              ' Remove all filedata from memory
End Sub

'*********Class_Initialize()***********
'*       Initialization Event         *
'*                                    *
'*  Raised when new instance is made  *
'*                                    *
'*   Sets all variable ready to use   *
'**************************************

Private Sub Class_Initialize()
    Dim i As Long
    sTempPath = Space(TempPathBuf)          ' Search for Windows Temporary files path
    i = GetTempPath(TempPathBuf, sTempPath) '   |
    If i > TempPathBuf Then                 '   |
        sTempPath = Space(i)                '   |
        i = GetTempPath(i, sTempPath)       '   |
    End If                                  '   |
    sTempPath = Left(TempPath, i)           '   V  Done
    pFiles = -1                             ' no files added yet so pFiles = -1
    ReadBuf = 1024
End Sub

Public Sub Abort(ByVal zlib As ZlibTool)    ' isn't this obvious?
    zlib.Abort
    bAbort = True
End Sub

Public Property Let TempPath(NewTempPath As String)
    If Dir(NewTempPath, vbDirectory) Then
        If Mid(NewTempPath, Len(NewTempPath)) <> "\" Then NewTempPath = NewTempPath & "\"
        sTempPath = NewTempPath
    End If
End Property

Public Property Get TempPath() As String
    TempPath = sTempPath
End Property

Public Property Get Files() As Long
    ' Usage example: lngVar = ZLibClass.Files
    Files = pFiles
End Property

Public Property Get ReadBuffer() As Long
    ReadBuffer = ReadBuf
End Property

Public Property Let ReadBuffer(Buffer As Long)
    If Buffer <> 0 And Buffer > -1 Then
        ReadBuf = Buffer
    End If
End Property

Public Property Get bytes(index As Long) As Long
    ' Usage example: lngVar = ZLibClass.bytes(index)
    If Not (index > pFiles) And (index > -1) Then
        bytes = File(index).bytes        ' return bytes
    Else
        bytes = -1                       ' an error occured
    End If
End Property

Public Property Let filename(index As Long, fn As String)
    ' Usage example: ZLibClass.filename(index) = "filename"
    If Not (index > pFiles) And (index > -1) Then
        If InStr(fn, ":") = 2 Then fn = Right(fn, Len(fn) - 2)  ' filename might be something like "C:\blah.tmp" but we don't want the "C:" part there, so let's remove it if it's there
        If InStr(fn, "\") <> 1 Then fn = "\" & fn               ' also, the filename might miss the "\" from the beginning, if that's so then we should add it there
        File(index).filename = fn
    End If
End Property

Public Property Get filename(index As Long) As String
    ' Usage example: strVar = ZLibClass.filename(index)
    If Not (index > pFiles) And (index > -1) Then
        filename = File(index).filename     ' return filename
    Else
        filename = ""                       ' an error occured
    End If
End Property

Private Sub FileAdd(ByVal filename As String)
    pFiles = pFiles + 1                                         '   Add new file
    ReDim Preserve File(pFiles)                                 '   reserve space for it in memory
    File(pFiles).path = filename
    File(pFiles).filename = Right(filename, Len(filename) - 2)  '   saves filename of the file
    File(pFiles).bytes = FileLen(filename)                      '   How many bytes in file
End Sub

'**************AddFile()**************
'*    Function Adds one file to      *
'*             filelist (File())     *
'*                                   *
'* Arguments: filename as String     *
'*             // File's name to add *
'* Optional attrib As VbFileAttribute*
'*  If wildcard is specified you may *
'*  specify attributes for filesearch*
'*                                   *
'* Return value: Boolean             *
'* if the file existed or not        *
'*************************************

Public Function AddFile(ByVal filename As String, Optional attrib As VbFileAttribute = vbNormal) As Boolean
    If Dir(filename, attrib) <> "" Then                                         ' If the file exists {
        If InStr(filename, "*") Or InStr(filename, "?") Then
            Dim temp As String, path As String, i As Long
            path = filename
            i = Len(path)
            Do While InStr(Mid(path, i), "\") <> 1
                i = i - 1
            Loop
            path = Mid(path, 1, i)
            temp = Dir(filename, attrib)
            Do While temp <> ""
                If (GetAttr(path & temp) And vbDirectory) <> vbDirectory Then
                    FileAdd path & temp
                End If
                temp = Dir
            Loop
        ElseIf attrib <> vbDirectory Then
            FileAdd filename
        End If
        AddFile = True                                                          '   the file was found so return true
    Else                                                                        ' } else {
        AddFile = False                                                         '   return false
    End If                                                                      ' }
End Function

'************RemoveFile()*************
'* Removes file(index) and organizes *
'* the files list again              *
'*                                   *
'* Arguments: index as Long          *
'*            // file index to remove*
'*                                   *
'* Return value: boolean             *
'* Wether the file was removed or not*
'*************************************

Public Function RemoveFile(index As Long) As Boolean
    If Not (index > pFiles) And index > -1 Then             ' Check if out of boundaries
        If pFiles = 0 Then                                  ' if there's only one file
            Erase File                                      ' release all memory
        Else                                                '
            Dim i As Long                                   '
            File(index).bytes = 0                           ' Remove the file
            File(index).filename = ""                       '
            pFiles = pFiles - 1                             ' decrement pFiles by one (pFiles--)
            For i = index To pFiles                         ' if there's a "hole" (empty file, no filename, bytedata etc) somewhere in file()
                File(i) = File(i + 1)                       ' replace it by next files in file()
            Next i                                          '
            ReDim Preserve File(pFiles)                     ' release memory
        End If                                              '
        RemoveFile = True                                   ' File was removed
    Else                                                    '
        RemoveFile = False                                  ' File was not removed
    End If                                                  '
End Function

'************Compress()***************
'*  Compress all files in file()     *
'*                                   *
'* Arguments:                        *
'* zlib as ZlibTool, output as string*
'* //zlibtool object, output filename*
'*************************************

Public Sub Compress(ByVal zlib As ZlibTool, ByVal output As String)
    Dim i As Long, a As Long, FileNum As Integer, FileNum2 As Integer, byt As String
    If pFiles > -1 Then
        FileNum = FreeFile                                      '
        Open TempPath & "zlibtemp" For Binary As #FileNum       ' Add all filedata in one file
            Put #FileNum, , Header                              ' Add a small header which to use as detection that the file is zlib and made with this class
            Put #FileNum, , pFiles                              ' how many files are in this file
            For i = 0 To pFiles                                  ' loop through all filedata
                If Not bAbort Then
                    Put #FileNum, , Len(File(i).filename)           ' Lenght of the filename, used to make a buffer to retrieve the filename
                    Put #FileNum, , File(i).filename                ' Filename
                    Put #FileNum, , File(i).bytes                   ' amount of bytes in file
                    FileNum2 = FreeFile
                    Open File(i).path For Binary As #FileNum2                   ' open file and put all data to the temp file
                        For a = 1 To Fix(LOF(FileNum2) / ReadBuf)
                            DoEvents
                            If Not bAbort Then
                                byt = Input(ReadBuf, FileNum2)                  ' Read
                                Put #FileNum, , byt                             ' Write
                            Else
                                Exit For                                        ' if aborted
                            End If
                        Next a
                        If Not bAbort Then
                            byt = Input((LOF(FileNum2) Mod ReadBuf), FileNum2)  ' read the rest of the file
                            Put #FileNum, , byt                                 ' Write
                        End If
                    Close #FileNum2
                Else
                    Exit For
                End If
            Next i                                              '
        Close #1                                                '
        If Not bAbort Then
            zlib.InputFile = TempPath & "zlibtemp"                  ' Compress
            zlib.OutputFile = output                                ' the temporary
            zlib.Compress                                           ' file
            Kill TempPath & "zlibtemp"                              ' destroy temporary file
        Else
            If Dir(TempPath & "zlibtemp") <> "" Then Kill TempPath & "zlibtemp"
            bAbort = False
        End If
    Else
        MsgBox "No Files to Compress", vbCritical Or vbOKOnly
    End If
End Sub

'*************CheckDir()***************
'* Checks the existance of the output *
'* path and creates it if neccessary  *
'* Note: This function is private so  *
'* the class user can't access this   *
'* from code.                         *
'*                                    *
'* Arguments:                         *
'*  directory As String               *
'*     The path originally saved to   *
'*     zlib file                      *
'*                                    *
'*  outputpath As String              *
'*    User defined path to extract to *
'*                                    *
'*  UsePaths As Boolean               *
'*    The code will ignore directory  *
'*    if this is false                *
'*                                    *
'* Return Value: String (filepath)    *
'**************************************

Private Function CheckDir(directory As String, outputpath As String, UsePaths As Boolean) As String
    Dim temp As String, i As Long
    If InStr(directory, ":") = 2 Then directory = Right(directory, Len(directory) - 2) ' if, for some odd reason, the Driveletter is still there, remove it
    Do While Mid(outputpath, Len(outputpath)) = "\"
        outputpath = Left(outputpath, Len(outputpath) - 1)
    Loop ' the outputpath shouldn't have "\" as the last character, remove it if there is one
    If InStr(outputpath, ":") <> 2 Then         ' The outputpath should also have drive
        If InStr(outputpath, "\") <> 1 Then     ' if it isn't, assume "C:"
            outputpath = "C:\" & outputpath     ' and add it there
        Else
            outputpath = "C:" & outputpath
        End If
    End If
    If Not UsePaths Then        ' if not using directory paths which were saved to the archive
        i = Len(directory)                  ' Then we shouldn't use then
        Do While Mid(directory, i) <> "\"
            i = i - 1
        Loop
        directory = Mid(directory, i)       ' Directories stripped away
    End If
    If Dir(outputpath & directory) = "" Then        ' if the file doesn't exist
        i = 4
        Do                                      ' make all directories which are missing but required
            DoEvents
            temp = Mid(outputpath & directory, 1, i)
            i = i + 1
            If Mid(temp, Len(temp)) = "\" And Dir(temp, vbDirectory) = "" Then If Dir(temp) = "" Then MkDir temp
            If temp = outputpath & directory Then Exit Do
        Loop
    Else    ' if the file exists
        Dim response As Long
        response = MsgBox("File " & outputpath & directory & " already exists, overwrite?", vbQuestion Or vbYesNo) ' Ask the user if he/she wants to overwrite the file
        If response = vbNo Then ' prevent overwrite if the users doesn't want to overwrite the file
            outputpath = ""
            directory = ""
        ElseIf response = vbYes Then ' destroy the file to overwrite, the client's always right
            Kill outputpath & directory
        End If
    End If
    CheckDir = outputpath & directory
End Function

'************Decompress()**************
'*  Decompresses all files from file  *
'*  compressed with Compress() sub    *
'*                                    *
'* Arguments:                         *
'*  ByVal zlib As ZlibTool            *
'*    Which zlibtool to use, if there *
'*    are many in the project         *
'*                                    *
'*  ByVal CompressedFile As String    *
'*    Path and filename of the file   *
'*    to decompress                   *
'*                                    *
'*  ByVal outputpath As String        *
'*    The path where the files are    *
'*    supposed to be extracted        *
'*                                    *
'*  UsePaths As Boolean               *
'*    Check CheckDir() comments       *
'**************************************

Public Sub Decompress(ByVal zlib As ZlibTool, ByVal CompressedFile As String, ByVal outputpath As String, UsePaths As Boolean)
    If Dir(CompressedFile) <> "" Then
        Dim h As String, g As Long, e As Long, i As Long, a As Long, _
        strFileName As String, byt As String, FileNum1 As Integer, FileNum2 As Integer
        zlib.InputFile = CompressedFile                                     ' Set Compressed File as zlibtool input
        zlib.OutputFile = TempPath & "zlibtemp"                             ' decompress it to windows temporary files directory
        zlib.Decompress                                                     '
        If Not bAbort Then
            FileNum1 = FreeFile                                             '
            Open TempPath & "zlibtemp" For Binary As #FileNum1              ' open the temp file
                h = Space(Len(Header))                                      ' create headerbuffer
                Get #FileNum1, , h                                          ' get header
                If h = Header Then                                          ' check if header if correct
                    Get #FileNum1, , e                                      ' check how many files are compressed in the file
                    For g = 0 To e                                          ' loop through all files
                        If Not bAbort Then
                            DoEvents                                        '
                            Get #FileNum1, , i                              ' get lenght of filename to create filenamebuffer
                            strFileName = Space(i)                          ' create filenamebuffer
                            Get #FileNum1, , strFileName                    ' get filename
                            Get #FileNum1, , i                              ' get bytes of file
                            h = outputpath                                  '
                            strFileName = CheckDir(strFileName, h, UsePaths) ' check dir (see description of CheckDir function)
                            If strFileName <> "" Then                       ' if the file didn't exist or the user chose to overwrite
                                FileNum2 = FreeFile                         '
                                Open strFileName For Binary As #FileNum2    ' open for write
                                    For a = 1 To Fix(i / ReadBuf)
                                        DoEvents
                                        If Not bAbort Then
                                            byt = Input(ReadBuf, FileNum1)  ' get bytes of the file from compressed file
                                            Put #FileNum2, , byt            ' and put them to the destination file
                                        Else
                                            Exit For
                                        End If
                                    Next a
                                    If Not bAbort Then
                                        byt = Input(i Mod ReadBuf, FileNum1) ' Read the rest of the file
                                        Put #FileNum2, , byt                 ' and write it
                                    End If
                                Close #FileNum2                             '
                                If Not bAbort Then strFileName = ""
                            Else                                            '
                                Seek #FileNum1, Loc(FileNum1) + i + 1       ' file existed and user didn't want to overwrite it; hop over the filedata
                            End If                                          '
                        Else
                            Exit For
                        End If
                    Next g                                                  '
                Else                                                        ' There was an error with the header
                    MsgBox "Not a " & Header & "-zLib file or corrupted", vbCritical Or vbOKOnly
                End If                                                      '
            Close #FileNum1
            If bAbort Then
                If Dir(strFileName) <> "" Then Kill strFileName             '
                bAbort = False
            End If
        End If                                                   '
        Kill zlib.OutputFile
    Else                                                                    '
        MsgBox "File not found", vbCritical Or vbOKOnly                     ' CompressedFile was not found
    End If                                                                  '
End Sub
