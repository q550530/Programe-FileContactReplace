Imports System.IO
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb.OleDbConnection
Imports System.Data.OleDb
Imports System.Xml
Imports Word = Microsoft.Office.Interop.Word



Public Class Form1

    

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim documentFormat As Object = 8
        Dim DirInfo As IO.DirectoryInfo
        Dim WithoutExe, outputFilename As String
        DirInfo = New IO.DirectoryInfo(TextBox3.Text)

        If sFileType.Text = "xml" Then

            For Each file In DirInfo.GetFiles("*." & sFileType.Text, IO.SearchOption.TopDirectoryOnly)
                WithoutExe = Path.GetFileNameWithoutExtension(file.Name)

                Dim fOut As StreamWriter = New StreamWriter(TextBox3.Text & "\output\" & file.Name)
                Using sr As StreamReader = New StreamReader(TextBox3.Text & "\" & file.Name)
                    Dim line As String
                    Do
                        line = sr.ReadLine()
                        If line Is Nothing Then
                            Exit Do
                        End If
                        fOut.WriteLine(line.Replace(TextBox1.Text, TextBox2.Text))

                    Loop Until line Is Nothing
                    sr.Close()
                    fOut.Close()
                End Using


            Next file
            MessageBox.Show("Complete")

        Else
            For Each file In DirInfo.GetFiles("*." & sFileType.Text, IO.SearchOption.TopDirectoryOnly) 'Find only for docx file



                WithoutExe = Path.GetFileNameWithoutExtension(file.Name)

                Dim objApp As New Word.Application
                objApp.Visible = True

                'Open an existing document.  
                Dim objDoc As Word.Document = objApp.Documents.Open(TextBox3.Text & "\" & file.Name)
                objDoc = objApp.ActiveDocument

                'Find and replace some text  
                objDoc.Content.Find.Execute(FindText:=TextBox1.Text, ReplaceWith:=TextBox2.Text, Replace:=Word.WdReplace.wdReplaceAll)
                While objDoc.Content.Find.Execute(FindText:="  ", Wrap:=Word.WdFindWrap.wdFindContinue)
                    objDoc.Content.Find.Execute(FindText:="  ", ReplaceWith:=" ", Replace:=Word.WdReplace.wdReplaceAll, Wrap:=Word.WdFindWrap.wdFindContinue)
                End While

                'outputFilename = System.IO.Path.ChangeExtension(file.Name, "htm")
                ''Close
                ' objDoc.SaveAs(TextBox3.Text & "\" & WithoutExe, documentFormat)
                objDoc.Save()
                objDoc.Close()
                objApp.Quit()
                objDoc = Nothing
                objApp = Nothing

            Next file
            MessageBox.Show("Complete")
        End If




    End Sub

   
End Class
