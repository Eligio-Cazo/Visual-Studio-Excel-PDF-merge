
Imports iTextSharp.text
Imports iTextSharp.text.pdf
Imports System.IO


Public Class Hoja1

    Private Sub Hoja1_Startup() Handles Me.Startup


    End Sub

    Private Sub Hoja1_Shutdown() Handles Me.Shutdown

    End Sub


    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Dim m, j
        Dim A$(100)
        m = 3
        A$(1) = Cells(1, 2).value
        A$(2) = Cells(2, 2).value
        A$(3) = Cells(3, 2).value

        'MsgBox(A$(1))
        'A$(1) = "c:\acadprg\515.pdf"
        'A$(2) = "c:\acadprg\515.pdf"
        'A$(3) = "c:\acadprg\excel.pdf"

        For j = 1 To m - 1
            If A$(j) = A$(m) Then
                MsgBox("Out Filename error", vbExclamation, "I/o Error")
                Exit Sub
            End If

        Next

        Dim Lista As New List(Of String)
        Dim pd$(51)

        For i = 1 To m
            pd$(i) = A$(i)
        Next i

        For i = 1 To m - 1
            Lista.Add(pd$(i))
        Next i

        Dim sFileJoin As String = A$(m)

        Dim Doc As New Document()

        Try
            Dim fs As New FileStream(sFileJoin, FileMode.Create, FileAccess.Write, FileShare.None)
            Dim copy As New PdfCopy(Doc, fs)
            Doc.Open()
            Dim Rd As PdfReader
            Dim n As Integer

            For Each file In Lista

                Rd = New PdfReader(file)
                n = Rd.NumberOfPages
                Dim page As Integer = 0

                Do While page < n
                    page += 1
                    copy.AddPage(copy.GetImportedPage(Rd, page))
                Loop

                copy.FreeReader(Rd)
                Rd.Close()


            Next

            MsgBox("Succes", vbExclamation, "Merge pdfs")

        Catch ex As Exception

            MsgBox(ex.Message, vbExclamation, "Error in merge pdfs")

        Finally

            Doc.Close()

        End Try



    End Sub
End Class
