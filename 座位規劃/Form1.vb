Imports Excel = Microsoft.Office.Interop.Excel
Public Class Form1
    Dim seat(7, 7) As String '顯示座號的二為陣列
    Dim skipNumArr As String() = {""} '跳過的座號
    Dim thisCnum As Integer = 1
    Dim Max As Integer = 40 '最大值
    Dim cancelNum As Integer = 0 '取消排序座位
    Dim setStart As Boolean = False '尚未設定初始值
    Dim directionBool As Boolean = True '上下左右(T) 左右上下(F)
    Dim hsort As Boolean = True '向左(T) 向右(F)
    Dim vsort As Boolean = True '向上(T) 向下(F)
    Dim randomRun As Boolean = False '亂數排序(預設未啟用)
    Dim ii, jj As Integer

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
    End Sub

    Dim changcolor As Boolean = True
    'Private Sub Timer1_Tick(sender As Object, e As EventArgs)
    '    If changcolor Then
    '        Title.ForeColor = Color.LightCyan
    '    Else
    '        Title.ForeColor = Color.MediumSlateBlue
    '    End If
    '    changcolor = Not (changcolor)
    'End Sub

    Private Sub count_run_Click(sender As Object, e As EventArgs) Handles count_run.Click
        If lastSeat.Text < 0 Then MsgBox("人數超出座位數!! 無法排序", 16, "錯誤") : Return
        If Not (setStart) Then MsgBox("請在座位表中任意一格雙擊左鍵 設定初始座位", 64, "警告") : Return
        If maxNum.ForeColor = Color.Black Then
            MsgBox(" '最後一號座號' 修改數值後未按下確認紐或Enter鍵，數值並未更新 !", 64, "警告")
            maxNum.ForeColor = Color.Red '更換顏色提醒使用者
            Return
        End If
        If Not (directionBool) Then '水平先排列起始位置對稱對調
            Dim Register As Integer = ii
            ii = jj
            jj = Register
        End If
        maxNum.Enabled = False '禁用textbox
        skipNum.Enabled = False
        Dim randomSeat As Integer() = {} '亂數陣列(尚未去除跳過座號)
        If randomRun Then randomSeat = creatlist()
        Do While thisCnum <= Max
            Do While ii >= 0 AndAlso ii < 7 AndAlso thisCnum <= Max
                Dim num As Integer = thisCnum '當格應填入的座號
                If randomRun Then num = randomSeat(thisCnum)
#Region "跳過指定座號"
                Do While Array.IndexOf(skipNumArr, num & "") <> -1
                    thisCnum += 1
                    If thisCnum > Max Then
                        showSeat(seat)
                        Return
                    End If
                    If randomRun Then
                        num = randomSeat(thisCnum)
                    Else
                        num = thisCnum
                    End If
                Loop
#End Region
#Region "跳過 X"
                If Me.Controls("L" & ii & jj).Text <> "X" Then '跳過 X
                    If directionBool Then
                        seat(ii, jj) = num
                    Else
                        seat(jj, ii) = num
                    End If
                    thisCnum += 1
                End If
#End Region
                If (directionBool And vsort) Or (Not (directionBool) And hsort) Then '向上遞減
                    ii -= 1
                Else '向下
                    ii += 1
                End If
            Loop
            ii = ii Mod 7 '上數循環
            If ii < 0 Then ii += 7 '下數循環

            If (directionBool And hsort) Or (Not (directionBool) And vsort) Then '向左
                jj -= 1
                If jj < 0 Then jj = 6
            Else '向右
                jj = (jj + 1) Mod 7
            End If
        Loop
        showSeat(seat)
        count_run.Enabled = False
    End Sub
#Region "輸出"
    Function showSeat(ByVal array(,) As String)
        For i = 0 To 6
            For j = 0 To 6
                Me.Controls("L" & i & j).Text = seat(i, j)
            Next
        Next
        Return 0
    End Function
#End Region

#Region "產生亂數數列"
    Function creatlist()
        Dim n As Integer = Max
        Dim A(n) As Integer
        For i = 1 To n
            A(i) = i
        Next
        Randomize()
        For i = n To 1 Step -1
            Dim num As Integer = Fix(i * Rnd()) + 1 'Fix()無條件捨去小數
            Dim temp As Integer = A(i)
            A(i) = A(num) '交換位置
            A(num) = temp
        Next
        Return A
    End Function
#End Region

#Region "設定初始值"
    Private Sub set_start_value(sender As Object, e As EventArgs) Handles L00.DoubleClick, L66.DoubleClick, L56.DoubleClick, L46.DoubleClick, L36.DoubleClick, L26.DoubleClick, L16.DoubleClick, L06.DoubleClick, L65.DoubleClick, L55.DoubleClick, L45.DoubleClick, L35.DoubleClick, L25.DoubleClick, L15.DoubleClick, L05.DoubleClick, L64.DoubleClick, L54.DoubleClick, L44.DoubleClick, L34.DoubleClick, L24.DoubleClick, L14.DoubleClick, L04.DoubleClick, L63.DoubleClick, L53.DoubleClick, L43.DoubleClick, L33.DoubleClick, L23.DoubleClick, L13.DoubleClick, L03.DoubleClick, L62.DoubleClick, L52.DoubleClick, L42.DoubleClick, L32.DoubleClick, L22.DoubleClick, L12.DoubleClick, L02.DoubleClick, L61.DoubleClick, L51.DoubleClick, L41.DoubleClick, L31.DoubleClick, L21.DoubleClick, L11.DoubleClick, L01.DoubleClick, L60.DoubleClick, L50.DoubleClick, L40.DoubleClick, L30.DoubleClick, L20.DoubleClick, L10.DoubleClick
        If setStart Then
            If sender.text = "✪" Then
                sender.text = ""
                setStart = False
            End If
            Return
        End If
        sender.text = "✪"
        ii = Int(sender.name.Substring(1, 1))
        jj = Int(sender.name.Substring(2))
        setStart = True
    End Sub
#End Region

#Region "設置空座位"
    Private Sub cancel_this(sender As Object, e As MouseEventArgs) Handles L66.MouseClick, L56.MouseClick, L46.MouseClick, L36.MouseClick, L26.MouseClick, L16.MouseClick, L06.MouseClick, L65.MouseClick, L55.MouseClick, L45.MouseClick, L35.MouseClick, L25.MouseClick, L15.MouseClick, L05.MouseClick, L64.MouseClick, L54.MouseClick, L44.MouseClick, L34.MouseClick, L24.MouseClick, L14.MouseClick, L04.MouseClick, L63.MouseClick, L53.MouseClick, L43.MouseClick, L33.MouseClick, L23.MouseClick, L13.MouseClick, L03.MouseClick, L62.MouseClick, L52.MouseClick, L42.MouseClick, L32.MouseClick, L22.MouseClick, L12.MouseClick, L02.MouseClick, L61.MouseClick, L51.MouseClick, L41.MouseClick, L31.MouseClick, L21.MouseClick, L11.MouseClick, L01.MouseClick, L60.MouseClick, L50.MouseClick, L40.MouseClick, L30.MouseClick, L20.MouseClick, L10.MouseClick, L00.MouseClick
        If e.Button = Windows.Forms.MouseButtons.Right Then
            Me.Controls.Remove(clickTitle) '移除體醒
            If sender.text = "X" Then
                sender.text = ""
                cancelNum -= 1
                lastSeat.Text += 1
            ElseIf sender.text = "" Then
                sender.text = "X"
                cancelNum += 1
                lastSeat.Text -= 1
            Else
            End If
        End If
    End Sub
#End Region

#Region "輕量事件"
    '排列方向
    Private Sub directionChanged(sender As Object, e As EventArgs) Handles RBtopL.CheckedChanged, RBleftT.CheckedChanged
        directionBool = RBtopL.Checked = True '僅限於兩種狀況 
    End Sub
    '垂直排列
    Private Sub vsort_change(sender As Object, e As EventArgs) Handles RBtop.CheckedChanged, RBbottom.CheckedChanged
        vsort = RBtop.Checked '僅限於兩種狀況 
    End Sub
    '水平排列
    Private Sub hsort_change(sender As Object, e As EventArgs) Handles RBleft.CheckedChanged, RBright.CheckedChanged
        hsort = RBleft.Checked '僅限於兩種狀況 
    End Sub
    '儲存最後一個座號
    Private Sub maxNum_KeyPress(sender As Object, e As KeyPressEventArgs) Handles maxNum.KeyPress
        If e.KeyChar = ChrW(Keys.Return) Then '偵測 Enter鍵
            btn_lastNum_Click(sender, e)
        End If
    End Sub
    Dim lastOne As Integer = 40
    Private Sub btn_lastNum_Click(sender As Object, e As EventArgs) Handles btn_lastNum.Click
        If maxNum.Text < 0 OrElse maxNum.Text > 50 Then MsgBox("人數不符合邏輯，請輸入正確數值", 16, "錯誤") : Return
        Max = Int(maxNum.Text)
        lastSeat.Text += (lastOne - Max)
        lastOne = Max
        maxNum.ForeColor = Color.DodgerBlue
    End Sub
    Private Sub maxNum_TextChanged(sender As Object, e As EventArgs) Handles maxNum.TextChanged
        If sender.text <> "" AndAlso sender.text <> lastOne Then maxNum.ForeColor = Color.Black
    End Sub
    '切割跳過座號存入陣列
    Dim skiplength As Integer = 0
    Private Sub skipNum_TextChanged(sender As Object, e As EventArgs) Handles skipNum.TextChanged
        skipNumArr = sender.text.split(",")
        If skipNumArr(0) = "" Then '歸零
            skiplength = 0
            lastSeat.Text = 49 - Max - cancelNum
            Return
        End If
        If skipNumArr.Length > skiplength Then '累加
            skiplength = skipNumArr.Length
            lastSeat.Text += 1
        ElseIf skipNumArr.Length < skiplength Then '遞減
            skiplength = skipNumArr.Length
            lastSeat.Text -= 1
        End If
    End Sub
    '亂數排序
    Private Sub CBrandom_CheckedChanged(sender As Object, e As EventArgs) Handles CBrandom.CheckedChanged
        randomRun = sender.checked
    End Sub
    '剩餘座位更動事件
    Dim lastNum As Integer = 9
    Private Sub lastSeat_TextChanged(sender As Object, e As EventArgs) Handles lastSeat.TextChanged
        Dim num As Integer = sender.text
        If num < 0 Then
            sender.forecolor = Color.Red
            '僅提出一次警告
            If num < lastNum AndAlso lastNum >= 0 Then MsgBox("人數將超出座位數!! 可能導致無法正常排序~", 64, "警告")
        Else
                sender.forecolor = Color.Black
        End If
        lastNum = num
    End Sub
#End Region
#Region "重置"
    Private Sub reset_Click(sender As Object, e As EventArgs) Handles reset.Click
        For i = 0 To 6
            For j = 0 To 6
                seat(i, j) = ""
            Next
        Next
        ii = 0
        jj = 0
        showSeat(seat)
        setStart = False
        thisCnum = 1
        cancelNum = 0
        skipNumArr = {""}
        skipNum.Text = ""
        lastSeat.Text = 49 - Max
        maxNum.Enabled = True '啟用textbox
        skipNum.Enabled = True
        count_run.Enabled = True '啟用Buttton
    End Sub
#End Region
#Region "輸出excel"
    Dim app As New Excel.Application
    Dim book As Excel.Workbook
    Dim sheet As Excel.Worksheet
    Dim range As Excel.Range
    Dim have As Boolean = False '檔案室否已存在(預設不存在)

    Private Sub outputToexcel_Click(sender As Object, e As EventArgs) Handles outputToexcel.Click
        '將路徑拉至桌面
        Dim displayPath As String = System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
        If System.IO.File.Exists(displayPath & "\考試座位表\考試座位表.xlsx") Then
            book = app.Workbooks.Open(displayPath & "\考試座位表\考試座位表.xlsx") '開啟一張已存在的 Excel 檔案
            sheet = book.Worksheets.Add()
            have = True
        Else
            have = False
            If Not (IO.Directory.Exists(displayPath & "\考試座位表")) Then
                '如不存在，建立資料夾
                IO.Directory.CreateDirectory(displayPath & "\考試座位表")
                MsgBox("先幫你在桌面建一個資料夾了 (資料夾名稱: 考試座位表)", 64, "建立資料夾成功")
            End If
            '不存在，建立空 Excel
            book = app.Workbooks.Add() '建立一個空 Excel
            sheet = book.Sheets(1)
        End If
        Dim sheetName As String = ""
        sheetName = InputBox("請輸入 ~新增工作表的名稱~", "提示", "(使用預設值) => 請直接按確定即可")
        If sheetName <> "(使用預設值) => 請直接按確定即可" Then sheet.Name = sheetName '變更工作表名稱

        For i = 0 To 6
            For j = 0 To 6
                range = sheet.Cells(i + 2, j + 5) '(0+1)+(向下位移1, 向右位移4)
                '寫入 Cell(1,1)
                range.Value = seat(i, j)
            Next
        Next
        If have Then
            book.Save() '存檔
        Else '另存檔案到 c:\...\桌面\考試座位表\考試座位表.xlsx"
            book.SaveAs(displayPath & "\考試座位表\考試座位表.xlsx")
        End If
        book.Close()
        MsgBox("已經幫你把資料輸出至excel放在資料夾了 >_< (檔名: 考試座位表.xlsx)", 64, "輸出完成")
        app.Quit() '結束操作
    End Sub
#End Region
End Class
