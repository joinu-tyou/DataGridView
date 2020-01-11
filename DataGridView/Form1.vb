Option Explicit On

Imports System.Reflection.MethodBase
Imports System.Drawing.ColorTranslator
Imports System.Text.RegularExpressions
Imports System.Windows.Forms

Public Class Form1


    'グリッドタイトルの色
    Public Const G_GRID_TITLE_COLOR As String = "#ffdddddd"

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load

        'MsgBox(Application.StartupPath & vbCrLf & System.Environment.CurrentDirectory)


        InitGridNoDataSource(fg1)
        InitGridNoDataSourceSort(fg2)
        InitGridDataSource(fg3)

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

        '第二行の設定
        fg.Rows(1).Cells(0).Value = "英語"
        fg.Rows(1).Cells(1).Value = ""
        fg.Rows(1).Cells(2).Value = ""

        fg.Rows(2).Cells(0).Value = "物理"
        fg.Rows(2).Cells(1).Value = ""
        fg.Rows(2).Cells(2).Value = ""

        fg.Rows(3).Cells(0).Value = "化学"
        fg.Rows(3).Cells(1).Value = ""
        fg.Rows(3).Cells(2).Value = ""

        fg.ResumeLayout()
        Exit Sub
ERR_HANDLE:
        fg.ResumeLayout()
        MsgBox(GetCurrentMethod.Name & vbCrLf & Err.Description)
    End Sub

    Private Sub InitGridNoDataSourceSort(ByRef fg As System.Windows.Forms.DataGridView)

        On Error GoTo ERR_HANDLE
        fg.SuspendLayout()
        '初期化
        fg.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        fg.ColumnHeadersHeight = 285 / 10
        fg.ColumnHeadersDefaultCellStyle.BackColor = FromHtml(G_GRID_TITLE_COLOR)


        fg.AllowUserToOrderColumns = True
        fg.AllowUserToDeleteRows = True
        fg.AllowUserToAddRows = False
        fg.AllowUserToResizeColumns = False
        fg.AllowUserToResizeRows = False
        fg.RowHeadersVisible = False
        fg.ColumnHeadersVisible = True
        'fg.ColumnHeadersHeight = 20
        fg.RowTemplate.Height = 285 / 10



        'AddHandler fg.ColumnHeaderMouseClick, AddressOf DataGridViewCellMouseEventHandler


        fg.ResumeLayout()


        fg.RowCount = 2
        fg.ColumnCount = 2



        fg.Columns(0).Name = "日付"
        fg.Columns(0).SortMode = DataGridViewColumnSortMode.Automatic
        fg.Columns(0).HeaderCell.Style.BackColor = FromHtml(G_GRID_TITLE_COLOR)
        fg.Columns(0).DefaultCellStyle.BackColor = FromHtml(G_GRID_TITLE_COLOR)


        fg.Columns(1).Name = "イベント"
        'fg.Columns(1).SortMode = DataGridViewColumnSortMode.Automatic

        fg.Rows(0).Cells(0).Value = 20200101L
        fg.Rows(0).Cells(1).Value = 1L
        fg.Rows(1).Cells(0).Value = 20200102L
        fg.Rows(1).Cells(1).Value = 2L
        Exit Sub
ERR_HANDLE:
        fg.ResumeLayout()
        MsgBox(GetCurrentMethod.Name & vbCrLf & Err.Description)
    End Sub


    Private Sub InitGridDataSource(ByRef fg As System.Windows.Forms.DataGridView)

        On Error GoTo ERR_HANDLE

        Dim csvFilePath As String

        csvFilePath = System.IO.Path.Combine(Application.StartupPath, "TestData.Csv")

        If System.IO.File.Exists(csvFilePath) Then


            '接続文字列
            Dim conString As String =
                "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" _
                + Application.StartupPath + ";Extended Properties=""text;HDR=Yes;FMT=Delimited"""
            Dim con As New System.Data.OleDb.OleDbConnection(conString)

            Dim commText As String = "SELECT * FROM [" + "TestData.Csv" + "]"
            Dim da As New System.Data.OleDb.OleDbDataAdapter(commText, con)

            'DataTableに格納する
            Dim dt As New DataTable
            da.Fill(dt)
            fg.SuspendLayout()
            fg.DataSource = dt
            fg.ResumeLayout()
        End If





        Exit Sub
ERR_HANDLE:
        fg.ResumeLayout()
        MsgBox(GetCurrentMethod.Name & vbCrLf & Err.Description)
    End Sub


End Class
