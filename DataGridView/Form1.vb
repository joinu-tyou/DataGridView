Option Explicit On

Imports System.Reflection.MethodBase
Imports System.Drawing.ColorTranslator
Imports System.Text.RegularExpressions
Imports System.Windows.Forms

Public Class Form1


    'グリッドタイトルの色
    Public Const G_GRID_TITLE_COLOR As String = "#ffdddddd"

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load

        InitGridNoDataSource(fg1)

    End Sub


    Private Sub InitGridNoDataSource(ByRef fg As System.Windows.Forms.DataGridView)

        On Error GoTo ERR_HANDLE

        fg.SuspendLayout()

        '初期化
        fg.AllowUserToOrderColumns = False
        fg.AllowUserToDeleteRows = False
        fg.AllowUserToAddRows = False
        fg.AllowUserToResizeColumns = False
        fg.AllowUserToResizeRows = False
        fg.RowHeadersVisible = False
        fg.ColumnHeadersVisible = False
        fg.RowTemplate.Height = 285 / 10



        fg.RowCount = 4
        fg.ColumnCount = 3


        '第一列の設定
        fg.Columns(0).Width = 50
        fg.Columns(0).ReadOnly = True
        fg.Columns(0).DefaultCellStyle.BackColor = FromHtml(G_GRID_TITLE_COLOR)

        '第二列の設定
        fg.Columns(1).Width = 80

        '第三列の設定
        fg.Columns(2).Width = 80

        '第一行の設定
        fg.Rows(0).ReadOnly = True
        fg.Rows(0).DefaultCellStyle.BackColor = FromHtml(G_GRID_TITLE_COLOR)

        '第一行第一列の設定
        fg.Rows(0).Cells(0).ReadOnly = False
        fg.Rows(0).Cells(0).Style.BackColor = Color.White




        fg.Rows(0).Cells(0).Value = "名前"


        fg.Rows(0).Cells(1).Value = "点数"
        fg.Rows(0).Cells(2).Value = "偏差値"

        fg.Rows(1).Cells(0).Value = "英語"
        fg.Rows(1).Cells(1).Value = ""
        fg.Rows(1).Cells(2).Value = ""

        fg.Rows(2).Cells(0).Value = "物理"
        fg.Rows(2).Cells(1).Value = ""
        fg.Rows(2).Cells(2).Value = ""

        fg.Rows(3).Cells(0).Value = "化学"
        fg.Rows(3).Cells(1).Value = ""
        fg.Rows(3).Cells(2).Value = ""

        'For i As Integer = 0 To 3
        '    fg.Rows(i).Height = 285 / 15 '.set_RowHeight(i, 285)
        'Next

        fg.Rows(0).Cells(0).Style.BackColor = Color.White

        fg.ResumeLayout()
        Exit Sub
ERR_HANDLE:
        fg.ResumeLayout()
        MsgBox(GetCurrentMethod.Name & vbCrLf & Err.Description)


    End Sub

End Class
