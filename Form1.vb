Imports System
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.Drawing
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Text.Json
Imports System.Threading.Tasks
Imports System.Windows.Forms

Public Class Form1
    Inherits Form

    ' --- Constants ---
    Private Const BookingURL As String = "https://c40.radioboss.fm/u/98"
    Private Const RowHeight As Integer = 52
    Private Const InitialRows As Integer = 10

    ' --- Win32 (top-most + dragging) ---
    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function SetWindowPos(hWnd As IntPtr, hWndInsertAfter As IntPtr,
                                         X As Integer, Y As Integer, cx As Integer, cy As Integer, uFlags As UInteger) As Boolean
    End Function

    <DllImport("user32.dll")>
    Private Shared Function ReleaseCapture() As Boolean
    End Function

    <DllImport("user32.dll")>
    Private Shared Function SendMessage(hWnd As IntPtr, msg As Integer, wParam As Integer, lParam As Integer) As IntPtr
    End Function

    Private Const WM_NCLBUTTONDOWN As Integer = &HA1
    Private Const HTCAPTION As Integer = 2

    Private Shared ReadOnly HWND_TOPMOST As New IntPtr(-1)
    Private Shared ReadOnly HWND_NOTOPMOST As New IntPtr(-2)
    Private Const SWP_NOSIZE As UInteger = &H1
    Private Const SWP_NOMOVE As UInteger = &H2
    Private Const SWP_SHOWWINDOW As UInteger = &H40

    ' --- Colors ---
    Private ReadOnly NeonGreen As Color = Color.FromArgb(0, 255, 0)
    Private ReadOnly BgColor As Color = Color.Black

    ' --- UI (single container approach) ---
    Private outerPanel As Panel
    Private contentPanel As Panel
    Private contentLayout As TableLayoutPanel
    Private headerRow As TableLayoutPanel
    Private rowsTable As TableLayoutPanel
    Private buttonsBar As FlowLayoutPanel

    Private btnBookingPage As Button
    Private btnClose As Button
    Private btnSave As Button
    Private btnLoad As Button
    Private btnReset As Button
    Private btnAlwaysOnTop As Button

    ' --- Row management ---
    Private rowControls As New List(Of RowControl)()

    ' --- Persistence ---
    Private ReadOnly repo As New RowRepository()

    ' --- State ---
    Private isForcedTopMost As Boolean = False

    Public Sub New()
        Text = "DJ Booking — Staff App"
        FormBorderStyle = FormBorderStyle.None
        StartPosition = FormStartPosition.CenterScreen
        Width = 1100
        Height = 900
        BackColor = Color.Black
        DoubleBuffered = True

        InitializeComponent()
        LoadRowsAsync()
        UpdateTopMostState()
        UpdateAlwaysOnTopButtonStyle()
    End Sub

    Private Async Sub LoadRowsAsync()
        Dim rows = Await repo.LoadAsync()
        If rows Is Nothing OrElse rows.Count = 0 Then
            For i As Integer = 1 To InitialRows
                AddRow(String.Empty, String.Empty, String.Empty)
            Next
        Else
            For Each r In rows
                AddRow(If(r.Time, String.Empty), If(r.DjName, String.Empty), If(r.DjLink, String.Empty))
            Next
        End If
    End Sub

    Private Sub InitializeComponent()
        ' Outer border (draggable)
        outerPanel = New Panel() With {
            .Dock = DockStyle.Fill,
            .BackColor = BgColor,
            .Padding = New Padding(10)
        }
        AddHandler outerPanel.Paint, AddressOf OuterPanel_Paint
        AddHandler outerPanel.MouseDown, AddressOf Header_MouseDown
        Controls.Add(outerPanel)

        ' Single content panel (everything lives here)
        contentPanel = New Panel() With {
            .Dock = DockStyle.Fill,
            .BackColor = Color.Black,
            .Padding = New Padding(8, 6, 8, 8)
        }
        AddHandler contentPanel.Paint, AddressOf ContentPanel_Paint
        AddHandler contentPanel.MouseDown, AddressOf Header_MouseDown
        outerPanel.Controls.Add(contentPanel)

        ' Content layout with 3 rows: header labels, rows table, bottom buttons
        contentLayout = New TableLayoutPanel() With {
            .Dock = DockStyle.Fill,
            .BackColor = Color.Black,
            .ColumnCount = 1,
            .RowCount = 3,
            .AutoSize = False,
            .AutoSizeMode = AutoSizeMode.GrowAndShrink,
            .Padding = New Padding(0),
            .Margin = New Padding(0)
        }
        contentLayout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))
        contentLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))   ' headerRow
        contentLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))   ' rowsTable
        contentLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))   ' buttonsBar
        contentPanel.Controls.Add(contentLayout)

        ' Header labels aligned to columns
        headerRow = New TableLayoutPanel() With {
            .Dock = DockStyle.Top,
            .BackColor = Color.Black,
            .ColumnCount = 4,
            .AutoSize = True,
            .AutoSizeMode = AutoSizeMode.GrowAndShrink,
            .Margin = New Padding(0, 0, 0, 6),
            .Padding = New Padding(6, 2, 6, 2)
        }
        headerRow.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 96))   ' Time
        headerRow.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 220))  ' DJ Name
        headerRow.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))   ' DJ Link
        headerRow.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 160))  ' Actions

        headerRow.Controls.Add(MakeHeader("Time"), 0, 0)
        headerRow.Controls.Add(MakeHeader("DJ Name"), 1, 0)
        headerRow.Controls.Add(MakeHeader("DJ Link"), 2, 0)
        headerRow.Controls.Add(MakeHeader("Actions"), 3, 0)
        AddHandler headerRow.MouseDown, AddressOf Header_MouseDown
        contentLayout.Controls.Add(headerRow, 0, 0)

        ' Rows table (adds rows dynamically)
        rowsTable = New TableLayoutPanel() With {
            .Dock = DockStyle.Top,
            .BackColor = Color.Black,
            .ColumnCount = 4,
            .CellBorderStyle = TableLayoutPanelCellBorderStyle.None,
            .AutoSize = True,
            .AutoSizeMode = AutoSizeMode.GrowAndShrink,
            .Margin = New Padding(0, 0, 0, 8),
            .Padding = New Padding(0)
        }
        rowsTable.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 96))
        rowsTable.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 220))
        rowsTable.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))
        rowsTable.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 160))
        AddHandler rowsTable.MouseDown, AddressOf Header_MouseDown
        contentLayout.Controls.Add(rowsTable, 0, 1)

        ' Buttons bar (flow, centered)
        buttonsBar = New FlowLayoutPanel() With {
            .Dock = DockStyle.Top,
            .BackColor = Color.Black,
            .AutoSize = True,
            .AutoSizeMode = AutoSizeMode.GrowAndShrink,
            .WrapContents = False,
            .Margin = New Padding(0),
            .Padding = New Padding(0, 4, 0, 0),
            .FlowDirection = FlowDirection.LeftToRight
        }
        AddHandler contentPanel.Resize, AddressOf CenterButtons
        contentLayout.Controls.Add(buttonsBar, 0, 2)

        ' Create buttons
        btnBookingPage = CreateNeonButton("Booking Page")
        btnSave = CreateNeonButton("Save")
        btnLoad = CreateNeonButton("Load")
        btnReset = CreateNeonButton("Reset")
        btnAlwaysOnTop = CreateNeonButton("Always On Top")
        btnClose = CreateNeonButton("Close")

        ' Add to bar
        buttonsBar.Controls.AddRange(New Control() {btnBookingPage, btnSave, btnLoad, btnReset, btnAlwaysOnTop, btnClose})

        ' Wire events
        AddHandler btnBookingPage.Click, AddressOf BtnBookingPage_Click
        AddHandler btnSave.Click, AddressOf BtnSave_Click
        AddHandler btnLoad.Click, AddressOf BtnLoad_Click
        AddHandler btnReset.Click, AddressOf BtnReset_Click
        AddHandler btnAlwaysOnTop.Click, AddressOf BtnAlwaysOnTop_Click
        AddHandler btnClose.Click, AddressOf BtnClose_Click

        AddHandler Me.Paint, AddressOf Form_Paint
    End Sub

    Private Sub CenterButtons(sender As Object, e As EventArgs)
        Dim totalWidth As Integer = 0
        Dim spacing As Integer = 12
        For i = 0 To buttonsBar.Controls.Count - 1
            Dim c = buttonsBar.Controls(i)
            c.Width = 150
            c.Height = 40
            c.Margin = New Padding(If(i = 0, 0, spacing \ 2), 4, If(i = buttonsBar.Controls.Count - 1, 0, spacing \ 2), 8)
            totalWidth += c.Width
            If i > 0 Then totalWidth += spacing
        Next
        Dim leftPad = Math.Max(12, (contentPanel.ClientSize.Width - totalWidth) \ 2)
        buttonsBar.Padding = New Padding(leftPad, 4, 0, 0)
    End Sub

    Private Function MakeHeader(text As String) As Control
        Dim lbl As New Label() With {
            .Text = text,
            .AutoSize = True,
            .ForeColor = NeonGreen,
            .BackColor = Color.Black,
            .Font = New Font("Segoe UI", 10, FontStyle.Bold),
            .Margin = New Padding(2)
        }
        Return lbl
    End Function

    Private Sub Header_MouseDown(sender As Object, e As MouseEventArgs)
        If e.Button = MouseButtons.Left Then
            ReleaseCapture()
            SendMessage(Me.Handle, WM_NCLBUTTONDOWN, HTCAPTION, 0)
        End If
    End Sub

    Private Sub DisposeRowControl(rc As RowControl)
        Try
            rc.NameTextBox?.Dispose()
            rc.LinkTextBox?.Dispose()
            rc.TimeCombo?.Dispose()
            rc.CopyButton?.Dispose()
            rc.DeleteButton?.Dispose()
            rc.TimePanel?.Dispose()
            rc.NamePanel?.Dispose()
            rc.LinkPanel?.Dispose()
            rc.ButtonsPanel?.Dispose()
        Catch ex As Exception
            Debug.WriteLine("DisposeRowControl failed: " & ex.ToString())
        End Try
    End Sub

    Private Sub OuterPanel_Paint(sender As Object, e As PaintEventArgs)
        Using p As New Pen(NeonGreen, 4)
            Dim r = outerPanel.ClientRectangle
            r.Inflate(-2, -2)
            e.Graphics.DrawRectangle(p, r)
        End Using
    End Sub

    Private Sub ContentPanel_Paint(sender As Object, e As PaintEventArgs)
        Using p As New Pen(NeonGreen, 2)
            Dim r = contentPanel.ClientRectangle
            r.Inflate(-6, -6)
            e.Graphics.DrawRectangle(p, r)
        End Using
    End Sub

    Private Sub Form_Paint(sender As Object, e As PaintEventArgs)
        Using p As New Pen(NeonGreen, 1)
            Dim r = Me.ClientRectangle
            r.Inflate(-4, -4)
            e.Graphics.DrawRectangle(p, r)
        End Using
    End Sub

    Private Sub ReindexRows()
        For i As Integer = 0 To rowControls.Count - 1
            Dim rc = rowControls(i)
            rowsTable.SetRow(rc.TimePanel, i)
            rowsTable.SetRow(rc.NamePanel, i)
            rowsTable.SetRow(rc.LinkPanel, i)
            rowsTable.SetRow(rc.ButtonsPanel, i)
        Next
    End Sub

    Private Sub DeleteRow(rc As RowControl)
        rowsTable.Controls.Remove(rc.TimePanel)
        rowsTable.Controls.Remove(rc.NamePanel)
        rowsTable.Controls.Remove(rc.LinkPanel)
        rowsTable.Controls.Remove(rc.ButtonsPanel)
        DisposeRowControl(rc)
        rowControls.Remove(rc)
        ReindexRows()
    End Sub

    Private Function CreateNeonButton(text As String) As Button
        Dim b As New Button() With {
            .Text = text,
            .Width = 150,
            .Height = 40,
            .BackColor = Color.Black,
            .ForeColor = NeonGreen,
            .FlatStyle = FlatStyle.Flat,
            .Font = New Font("Segoe UI", 10, FontStyle.Regular),
            .Margin = New Padding(0)
        }
        b.FlatAppearance.BorderColor = NeonGreen
        b.FlatAppearance.BorderSize = 3
        b.FlatAppearance.MouseOverBackColor = Color.FromArgb(24, 24, 24)
        Return b
    End Function

    Private Sub AddRow(time As String, djName As String, djLink As String)
        Dim index = rowControls.Count
        rowsTable.RowCount += 1
        rowsTable.RowStyles.Add(New RowStyle(SizeType.Absolute, RowHeight))

        ' Panels per cell (for neon borders)
        Dim timePanel As New Panel() With {.Dock = DockStyle.Fill, .BackColor = Color.Black, .Padding = New Padding(2), .Margin = New Padding(2)}
        Dim namePanel As New Panel() With {.Dock = DockStyle.Fill, .BackColor = Color.Black, .Padding = New Padding(4), .Margin = New Padding(2)}
        Dim linkPanel As New Panel() With {.Dock = DockStyle.Fill, .BackColor = Color.Black, .Padding = New Padding(4), .Margin = New Padding(2)}
        Dim buttonsPanel As New Panel() With {.Dock = DockStyle.Fill, .BackColor = Color.Black, .Padding = New Padding(4), .Margin = New Padding(2)}

        AddHandler timePanel.Paint, Sub(s, e)
                                        Using pen As New Pen(NeonGreen, 2)
                                            Dim r = timePanel.ClientRectangle : r.Inflate(-1, -1)
                                            e.Graphics.DrawRectangle(pen, r)
                                        End Using
                                    End Sub
        AddHandler namePanel.Paint, Sub(s, e)
                                        Using pen As New Pen(NeonGreen, 2)
                                            Dim r = namePanel.ClientRectangle : r.Inflate(-1, -1)
                                            e.Graphics.DrawRectangle(pen, r)
                                        End Using
                                    End Sub
        AddHandler linkPanel.Paint, Sub(s, e)
                                        Using pen As New Pen(NeonGreen, 2)
                                            Dim r = linkPanel.ClientRectangle : r.Inflate(-1, -1)
                                            e.Graphics.DrawRectangle(pen, r)
                                        End Using
                                    End Sub
        AddHandler buttonsPanel.Paint, Sub(s, e)
                                           Using pen As New Pen(NeonGreen, 2)
                                               Dim r = buttonsPanel.ClientRectangle : r.Inflate(-1, -1)
                                               e.Graphics.DrawRectangle(pen, r)
                                           End Using
                                       End Sub

        ' Time combo
        Dim cb As New ComboBox() With {
            .DropDownStyle = ComboBoxStyle.DropDownList,
            .BackColor = Color.Black,
            .ForeColor = NeonGreen,
            .FlatStyle = FlatStyle.Flat,
            .Dock = DockStyle.Left,
            .Width = 88
        }
        cb.Items.AddRange(GetDefaultTimesOnHour())
        If Not String.IsNullOrEmpty(time) AndAlso cb.Items.Contains(time) Then
            cb.SelectedItem = time
        End If
        timePanel.Controls.Add(cb)

        ' Name textbox
        Dim txtName As New TextBox() With {
            .Dock = DockStyle.Fill,
            .Multiline = False,
            .Font = New Font("Segoe UI", 10),
            .ForeColor = NeonGreen,
            .BackColor = Color.Black,
            .BorderStyle = BorderStyle.None,
            .Text = djName
        }
        namePanel.Controls.Add(txtName)

        ' Link textbox
        Dim txtLink As New TextBox() With {
            .Dock = DockStyle.Fill,
            .Multiline = False,
            .Font = New Font("Segoe UI", 10),
            .ForeColor = NeonGreen,
            .BackColor = Color.Black,
            .BorderStyle = BorderStyle.None,
            .Text = djLink
        }
        linkPanel.Controls.Add(txtLink)

        ' Buttons (Copy + Delete)
        Dim btnCopy As New Button() With {
            .Text = "Copy",
            .BackColor = Color.Black,
            .ForeColor = NeonGreen,
            .FlatStyle = FlatStyle.Flat,
            .Height = 24,
            .Dock = DockStyle.Top
        }
        btnCopy.FlatAppearance.BorderColor = NeonGreen
        btnCopy.FlatAppearance.BorderSize = 2
        AddHandler btnCopy.Click, Sub(s, e)
                                      Try
                                          Clipboard.SetText(txtLink.Text)
                                          btnCopy.Text = "Copied"
                                          Dim t = New Timer() With {.Interval = 900}
                                          AddHandler t.Tick, Sub(ss, ee)
                                                                 btnCopy.Text = "Copy"
                                                                 t.Stop()
                                                                 t.Dispose()
                                                             End Sub
                                          t.Start()
                                      Catch ex As Exception
                                          MessageBox.Show("Copy failed: " & ex.Message)
                                      End Try
                                  End Sub

        Dim btnDelete As New Button() With {
            .Text = "Delete",
            .BackColor = Color.Black,
            .ForeColor = Color.Red,
            .FlatStyle = FlatStyle.Flat,
            .Height = 24,
            .Dock = DockStyle.Top
        }
        btnDelete.FlatAppearance.BorderColor = NeonGreen
        btnDelete.FlatAppearance.BorderSize = 2

        Dim rcObj As New RowControl With {
            .TimePanel = timePanel,
            .NamePanel = namePanel,
            .LinkPanel = linkPanel,
            .ButtonsPanel = buttonsPanel,
            .TimeCombo = cb,
            .NameTextBox = txtName,
            .LinkTextBox = txtLink,
            .CopyButton = btnCopy,
            .DeleteButton = btnDelete
        }
        AddHandler btnDelete.Click, Sub(s, e) DeleteRow(rcObj)

        buttonsPanel.Controls.Add(btnDelete)
        buttonsPanel.Controls.Add(btnCopy)

        rowsTable.Controls.Add(timePanel, 0, index)
        rowsTable.Controls.Add(namePanel, 1, index)
        rowsTable.Controls.Add(linkPanel, 2, index)
        rowsTable.Controls.Add(buttonsPanel, 3, index)

        rowControls.Add(rcObj)
    End Sub

    Private Shared Function GetDefaultTimesOnHour() As String()
        Dim times As New List(Of String)
        For h As Integer = 0 To 23
            times.Add($"{h:00}:00")
        Next
        Return times.ToArray()
    End Function

    Private Async Sub BtnSave_Click(sender As Object, e As EventArgs)
        Dim data As New List(Of RowModel)
        For Each rc In rowControls
            data.Add(New RowModel With {
                .Time = If(rc.TimeCombo.SelectedItem?.ToString(), String.Empty),
                .DjName = rc.NameTextBox.Text,
                .DjLink = rc.LinkTextBox.Text
            })
        Next
        Await repo.SaveAsync(data)
        MessageBox.Show("Saved.", "Save", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Async Sub BtnLoad_Click(sender As Object, e As EventArgs)
        Dim data = Await repo.LoadAsync()
        If data Is Nothing OrElse data.Count = 0 Then
            MessageBox.Show("No saved data found.", "Load", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        While rowControls.Count > 0
            DeleteRow(rowControls(0))
        End While

        For Each r In data
            AddRow(If(r.Time, String.Empty), If(r.DjName, String.Empty), If(r.DjLink, String.Empty))
        Next
    End Sub

    Private Sub BtnReset_Click(sender As Object, e As EventArgs)
        If MessageBox.Show("Reset to 10 empty slots?", "Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            While rowControls.Count > 0
                DeleteRow(rowControls(0))
            End While
            For i As Integer = 1 To InitialRows
                AddRow(String.Empty, String.Empty, String.Empty)
            Next
        End If
    End Sub

    Private Sub BtnAlwaysOnTop_Click(sender As Object, e As EventArgs)
        isForcedTopMost = Not isForcedTopMost
        UpdateTopMostState()
        UpdateAlwaysOnTopButtonStyle()
    End Sub

    Private Sub UpdateAlwaysOnTopButtonStyle()
        If isForcedTopMost Then
            btnAlwaysOnTop.BackColor = NeonGreen
            btnAlwaysOnTop.ForeColor = Color.Black
            btnAlwaysOnTop.Text = "Staying On Top"
        Else
            btnAlwaysOnTop.BackColor = Color.Black
            btnAlwaysOnTop.ForeColor = NeonGreen
            btnAlwaysOnTop.Text = "Always On Top"
        End If
    End Sub

    Private Sub BtnClose_Click(sender As Object, e As EventArgs)
        Try
            Application.Exit()
        Catch
        End Try
        Try
            Dim t As New Timer() With {.Interval = 200}
            AddHandler t.Tick, Sub()
                                   t.Stop()
                                   t.Dispose()
                                   Try
                                       Process.GetCurrentProcess().Kill()
                                   Catch
                                   End Try
                               End Sub
            t.Start()
        Catch
        End Try
    End Sub

    Private Sub BtnBookingPage_Click(sender As Object, e As EventArgs)
        Try
            Process.Start(New ProcessStartInfo(BookingURL) With {.UseShellExecute = True})
        Catch ex As Exception
            Debug.WriteLine("BtnBookingPage_Click failed: " & ex.ToString())
        End Try
    End Sub

    Private Sub UpdateTopMostState()
        Dim shouldBeTopMost = isForcedTopMost
        Me.TopMost = shouldBeTopMost
        Try
            Dim h = Me.Handle
            SetWindowPos(h, If(shouldBeTopMost, HWND_TOPMOST, HWND_NOTOPMOST), 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW)
        Catch ex As Exception
            Debug.WriteLine("UpdateTopMostState failed: " & ex.ToString())
        End Try
    End Sub

    ' --- Helper classes ---
    Private Class RowControl
        Public Property TimePanel As Panel
        Public Property NamePanel As Panel
        Public Property LinkPanel As Panel
        Public Property ButtonsPanel As Panel
        Public Property TimeCombo As ComboBox
        Public Property NameTextBox As TextBox
        Public Property LinkTextBox As TextBox
        Public Property CopyButton As Button
        Public Property DeleteButton As Button
    End Class

    Private Class RowModel
        Public Property Time As String
        Public Property DjName As String
        Public Property DjLink As String
    End Class

    Private Class RowRepository
        Private ReadOnly folder As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "DJSchedulerApp")
        Private ReadOnly filePath As String = Path.Combine(folder, "rows.json")
        Private ReadOnly options As JsonSerializerOptions = New JsonSerializerOptions With {.WriteIndented = True}

        Public Sub New()
            Try
                Directory.CreateDirectory(folder)
            Catch ex As Exception
                Debug.WriteLine("RowRepository.New failed creating folder: " & ex.ToString())
            End Try
        End Sub

        Public Async Function SaveAsync(rows As List(Of RowModel)) As Task
            Try
                Dim json = JsonSerializer.Serialize(rows, options)
                Await File.WriteAllTextAsync(filePath, json, Encoding.UTF8)
            Catch ex As Exception
                Debug.WriteLine("Failed saving rows: " & ex.ToString())
            End Try
        End Function

        Public Async Function LoadAsync() As Task(Of List(Of RowModel))
            Try
                If Not File.Exists(filePath) Then Return New List(Of RowModel)()
                Dim json = Await File.ReadAllTextAsync(filePath, Encoding.UTF8)
                Dim list = JsonSerializer.Deserialize(Of List(Of RowModel))(json, options)
                Return If(list, New List(Of RowModel)())
            Catch ex As Exception
                Debug.WriteLine("Failed loading rows: " & ex.ToString())
                Return New List(Of RowModel)()
            End Try
        End Function
    End Class
End Class