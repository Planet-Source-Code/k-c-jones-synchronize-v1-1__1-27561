Attribute VB_Name = "startup"
Option Explicit
Public Sub Main()
    Dim cArgs
    Dim i As String
    Dim a, b
    Dim oDir
    Dim oFso As New FileSystemObject

    Dim oFiles
    Dim File
    
    'Set oFso = CreateObject("fso.Object")
    i = "e:\test1 e:\test2"
    cArgs = Splitter(i, " ")
    
    'cArgs = Splitter(Command, " ")
    For Each a In cArgs
        b = b + 1
        
    Next
    'Check to make sure only two arguments where passed
    If b = 2 Then
        'check to see if source param is a folder
        If oFso.FolderExists(cArgs(1)) Then
            'check to see if second param is a folder for dest
            If oFso.FolderExists(cArgs(2)) Then
                'now det the source folder
                Set oDir = oFso.GetFolder(cArgs(1))
                'get all the files
                Set oFiles = oDir.Files
                'copy the files with overwrite
                For Each File In oFiles
                    'Check if file exists
                    If Not oFso.FileExists(cArgs(2) & "\" & File.Name) Then
                        'if not copy it
                        File.Copy cArgs(2) & "\", True
                    Else
                        'check time stamp if diffrent then replace it
                        If Not oFso.GetFile(cArgs(2) & "\" & File.Name).DateLastModified = File.DateLastModified Then
                            File.Copy cArgs(2) & "\", True
                        End If
                    End If
                Next
                'Now reverse the process
                Set oDir = Nothing
                Set oDir = oFso.GetFolder(cArgs(2))
                'get all the files
                Set oFiles = oDir.Files
                'delete the files that do not exist anymore
                For Each File In oFiles
                    
                    If Not oFso.FileExists(cArgs(1) & "\" & File.Name) Then
                        File.Delete
                    End If
                Next
                
            Else
                MsgBox ("I'm sorry but the Destination must be a folder")
            End If
        Else
            MsgBox ("I'm sorry but the Source must be a folder")
        End If
                
    Else
        MsgBox "I'm sorry you must specify synchronize.exe [SOURCE] [DESTINATION]"
        End
    End If
    
    
End Sub
Public Function Splitter(SplitString As String, SplitLetter As String) As Variant
    ReDim SplitArray(1 To 1) As Variant
    Dim TempLetter As String
    Dim TempSplit As String
    Dim i As Integer
    Dim x As Integer
    Dim StartPos As Integer
    
    SplitString = SplitString & SplitLetter


    For i = 1 To Len(SplitString)
        TempLetter = Mid(SplitString, i, Len(SplitLetter))


        If TempLetter = SplitLetter Then
            TempSplit = Mid(SplitString, (StartPos + 1), (i - StartPos) - 1)


            If TempSplit <> "" Then
                x = x + 1
                ReDim Preserve SplitArray(1 To x) As Variant
                SplitArray(x) = TempSplit
            End If
            StartPos = i
        End If
    Next i
    Splitter = SplitArray
End Function

