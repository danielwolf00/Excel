Public Class Form1
    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim ExcelApp = New Microsoft.Office.Interop.Excel.Application
        Dim libro = ExcelApp.Workbooks.Add

        libro.Sheets(1).cells(1, 1) = "hola mundo"
        libro.SaveAs(Filename:="C:\Users\danie\OneDrive\Escritorio\test1.xlsx")

        Label1.Text = "El registro fue un exito"

        ExcelApp.Quit()
        libro = Nothing
        ExcelApp = Nothing

    End Sub
End Class
