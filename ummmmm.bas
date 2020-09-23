Attribute VB_Name = "ummmmm"
'--------START GLOBAL STRINGS FOR THIS PROJECT-----
Global strSvrURL As String
Global URL As String
Global RESUMEFILE As Boolean
Global FilePathName As String
Global Filename As String
Global FileLength As Single
Global Sec%, Min%, Hr%
Global Unix As Boolean
Public Function GETDATAHEAD(DATA As Variant, ToRetrieve As String)
On Error Resume Next
If DATA = "" Then Exit Function
Dim EndBYTES%, A$, LENGTHEND%, PART%, Part2%, RetrieveLength%
If InStr(DATA, ToRetrieve) > 0 Then
LENGTHEND = Len(DATA)
PART = InStr(DATA, ToRetrieve)
RetrieveLength = Len(ToRetrieve)
A = Right(DATA, LENGTHEND - PART - RetrieveLength)
LENGTHEND = Len(A)
If InStr(A, vbCrLf) > 0 Then
Part2 = InStr(A, vbCrLf)
A = Left(A, Part2 - 1)
End If
GETDATAHEAD = A
End If
End Function
Public Function OutFileName(File$) As String
Dim P%
    P = InStr(File$, ".") 'Check for the period in the file
    If P = 0 Then
        OutFileName = File & "ext" & ".rsm" 'If no period then add a period and extension to it
        Exit Function
    End If
If LCase(Right(File$, 3) = "rsm") Then 'Check to see if its extension is the resuming file extension used by downloader
Dim Length%, A$, B$
    P = InStr(File$, ".")
    A = Left(File$, P - 1) 'Trimming off the filename without added extension
    B = Right(A, 3) 'Getting extension of original filename
    Length = Len(A$)
    A = Left(A, Length - 3) 'get rid of the original extension
    OutFileName = A & "." & B 'add original extension back on with period
Else 'if its not a resumable file then make it one!
Dim Dot%, One$, Ext$, SLength%
    Dot = InStr(File$, ".") 'get position of period
    
    One = Left(File$, Dot - 1) 'Get the filename by itself
    Ext = Right(File$, 3) 'Get the extension by itself
    OutFileName = One & Ext & ".rsm" 'Put the rsm file extension onto the file!
End If
End Function
