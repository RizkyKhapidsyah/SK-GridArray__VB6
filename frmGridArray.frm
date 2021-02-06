VERSION 5.00
Begin VB.Form frmGridArray 
   Caption         =   "Display Data"
   ClientHeight    =   8415
   ClientLeft      =   1755
   ClientTop       =   525
   ClientWidth     =   14280
   LinkTopic       =   "Form1"
   ScaleHeight     =   8415
   ScaleWidth      =   14280
   Begin VB.Frame Frame1 
      Height          =   8295
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   14055
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   7920
         Width           =   13815
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   7455
         Left            =   13680
         TabIndex        =   3
         Top             =   360
         Width           =   255
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         Height          =   7455
         Left            =   120
         ScaleHeight     =   7395
         ScaleWidth      =   13395
         TabIndex        =   1
         Top             =   360
         Width           =   13455
         Begin VB.PictureBox Picture2 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FFFFFF&
            Height          =   7095
            Left            =   120
            ScaleHeight     =   7035
            ScaleWidth      =   13035
            TabIndex        =   2
            Top             =   120
            Width           =   13095
            Begin VB.TextBox Text1 
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   375
               Left            =   120
               TabIndex        =   5
               Top             =   120
               Width           =   1815
            End
         End
         Begin VB.Shape Shape1 
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00C0C0C0&
            FillStyle       =   0  'Solid
            Height          =   7095
            Left            =   240
            Top             =   240
            Width           =   13095
         End
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileSub 
         Caption         =   "P&rint Doc"
         Enabled         =   0   'False
         Index           =   0
      End
      Begin VB.Menu mnuFileSub 
         Caption         =   "&Preview"
         Index           =   1
      End
      Begin VB.Menu mnuFileSub 
         Caption         =   "E&xit"
         Index           =   2
      End
   End
   Begin VB.Menu mnuAllign 
      Caption         =   "Allign"
      Begin VB.Menu mnuAllignSub 
         Caption         =   "Left Justify"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu mnuAllignSub 
         Caption         =   "Center Text"
         Index           =   1
      End
      Begin VB.Menu mnuAllignSub 
         Caption         =   "Right Justify"
         Index           =   2
      End
   End
End
Attribute VB_Name = "frmGridArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public prtObj As Object 'Print Object target
Public DocDate As String 'Store Todays Date
Public Title As String 'Stores report title
Public Page As Integer 'Stores Page Number
Public TopMargin As Single 'Stores TopMargin
Public LeftMargin As Single 'Stores Left Margin
Public RightMargin As Single 'Stores Right Margin
Public BottomMargin As Single 'Stores Bottom Margin
Public HeaderHeight As Single 'Height of the Title
Public FooterHeight As Single 'Height of the Footer
Public BodyTop As Single 'Store Top of Body
Public BodyWidth As Single 'Width of Document Body
Public BodyHeight As Single 'Height of Document Body
Dim Lines() As Single 'Stores Each Line Coordinates
Dim ColWidth() As Single 'Store Width of Each Column
Dim ColLeft() As Single 'Stores the left coordinates of each Column
Dim ColNames() As String 'Stores Column Names
Dim DataArray()
Dim Row As Integer, Col As Integer, RowBottom As Single
Dim RowHeight As Single, ColHeight As Single
Public ColNameHeight As Single, DocWidth As Single, DocHeight As Single
Public I As Integer, L As Integer, R As Integer, C As Integer
Private Sub LoadData()
    ReDim DataArray(UBound(Lines, 1) - 2, UBound(ColLeft, 1) - 1)
    'Print the body data R is Row C is Column when printing from a
    'database R would be record C would be field
    For R = 0 To UBound(Lines, 1) - 2
        For C = 0 To UBound(ColLeft, 1) - 1
            DataArray(R, C) = "Row " & Str(R) & " Col " & C
        Next
    Next
End Sub
Private Sub CellLocate(X1 As Single, Y1 As Single)
    For I = 0 To UBound(ColLeft, 1) - 1
        If X1 > ColLeft(I) And X1 < ColLeft(I + 1) Then Col = I
    Next
    If X1 > ColLeft(UBound(ColLeft, 1) - 1) And X1 < RightMargin Then
        Col = I - 1
    End If
    Row = Int((Y1 - BodyTop) / RowHeight)
End Sub
Sub PageSetup()
    L = 1
    Page = 1
    DocDate = Date 'Current Date
    'Set up document margins
    TopMargin = 720 '1/2 inch in TWIPS
    LeftMargin = 720 '1/2 inch
    RightMargin = prtObj.Width - LeftMargin '# of TWIPS from Left
    BottomMargin = prtObj.Height - TopMargin 'Height - 1/2 in
    'Height of header and footer
    HeaderHeight = 1440 'Inch title
    FooterHeight = 1440 'Inch footer
    'Calculating data body width and height
    BodyWidth = RightMargin - LeftMargin
    BodyHeight = prtObj.ScaleHeight - (TopMargin + HeaderHeight + FooterHeight + (prtObj.ScaleHeight - BottomMargin))
    'Column Info
    ReDim ColLeft(5) 'Number of Columns
    ReDim ColNames(UBound(ColLeft, 1)) 'Array to hold column names
    ReDim ColWidth(UBound(ColLeft, 1)) 'Array to hold Column Widths
    ColLeft(0) = LeftMargin 'Position of first column in TWIPS
    prtObj.FontSize = 14 'Set Printer Object to column fonts
    ColNameHeight = prtObj.TextHeight("A")
    'Create body top cooridinates
    BodyTop = TopMargin + HeaderHeight + ColNameHeight
    'Create Column Info
    For I = 0 To UBound(ColLeft, 1) - 1
        'Column Names
        ColNames(I) = "Column " & Str(I)
        'Column Widths
        ColWidth(I) = BodyWidth / UBound(ColLeft, 1)
        'Left Position of columns 1 to the # of Columns
        If I > 0 Then ColLeft(I) = ColLeft(I - 1) + ColWidth(I)
    Next
    prtObj.FontSize = 12 'Set the Data FONT object
    'Calculate Number of lines
    ReDim Lines(Int(BodyHeight / prtObj.TextHeight("A")))
    'Position the First Line
    Lines(0) = BodyTop + prtObj.TextHeight("A")
    'Store the top position of each line to the Lines Array
    For I = 1 To UBound(Lines, 1) - 2
        Lines(I) = BodyTop + (prtObj.TextHeight(Lines(I)) * (I + 1))
    Next
End Sub

Sub SetScrollBars(DisplayObj As Object, Container As Object)
    With VScroll1
        .Left = Container.Left + Container.Width
        .Top = Container.Top
        .Max = (DisplayObj.Height - Container.ScaleHeight) + 200 '32,767
        .Min = -200
        .Value = .Min
        .Height = Container.Height
        .SmallChange = Container.Height / 10
        .LargeChange = Container.Height
    End With
    If DisplayObj.ScaleHeight > Container.ScaleHeight Then
        VScroll1.Visible = True
    Else
        VScroll1.Visible = False
    End If
    With HScroll1
        .Left = Container.Left
        .Top = Container.Top + Container.Height
        .Min = -200
        .Width = Container.Width
        .Value = .Min
        .Max = (DisplayObj.Width - Container.ScaleWidth) + 200
        .SmallChange = Container.ScaleWidth / 10
        .LargeChange = Container.Width
    End With
    If DisplayObj.ScaleWidth > Container.ScaleWidth Then
        HScroll1.Visible = True
    Else
        HScroll1.Visible = False
    End If
End Sub

Private Sub PrintReport()
    Title = "Report Title" 'Report Title
    prtObj.FontSize = 48
    prtObj.FontBold = True
    'Position at top margin
    prtObj.CurrentY = TopMargin
    'Center and print the title
    prtObj.CurrentX = (prtObj.ScaleWidth / 2) - (prtObj.TextWidth(Title) / 2)
    prtObj.Print Title
    'Set the column names FONT
    prtObj.FontSize = 14
    'Print Line before the Column Names
    prtObj.Line (LeftMargin, BodyTop - (ColNameHeight + 50))-(RightMargin, BodyTop - (ColNameHeight + 50))
    'Position at body top - the height of the column text
    prtObj.CurrentY = BodyTop - ColNameHeight
    DoEvents
    'Position and print the column names
    For I = 0 To UBound(ColLeft, 1) - 1
        prtObj.CurrentX = ColLeft(I)
        prtObj.Print ColNames(I);
    Next
    'Advance to next line
    prtObj.Print
    'Print line under column names
    prtObj.Line (LeftMargin, BodyTop - 50)-(RightMargin, BodyTop - 50)
    'Set the print object font to the data font
    prtObj.FontSize = 12
    'set Position to the top of the body
    prtObj.CurrentY = BodyTop
    prtObj.FontBold = False
    RowHeight = prtObj.TextHeight("A")
    'Print the body data R is Row C is Column when printing from a
    'database R would be record C would be field
    For R = 0 To UBound(DataArray, 1)
        For C = 0 To UBound(DataArray, 2)
            If mnuAllignSub(0).Checked Then prtObj.CurrentX = ColLeft(C) 'Allign Left
            If mnuAllignSub(1).Checked Then prtObj.CurrentX = ColLeft(C) + ((ColWidth(C) - prtObj.TextWidth(DataArray(R, C))) / 2) 'Allign center
            If mnuAllignSub(2).Checked Then prtObj.CurrentX = ColLeft(C) + (ColWidth(C) - prtObj.TextWidth(DataArray(R, C))) 'Allign right
            prtObj.Print DataArray(R, C);
        Next
        prtObj.CurrentY = Lines(R)
    Next
    RowBottom = prtObj.CurrentY
    'Print Line above the footer
    prtObj.Line (LeftMargin, prtObj.CurrentY)-(RightMargin, prtObj.CurrentY)
End Sub

Private Sub Form_Load()
    DocWidth = Printer.ScaleWidth
    DocHeight = Printer.ScaleHeight
    'Special Effect using a filled shape for the document shadow
    Shape1.Width = DocWidth
    Shape1.Height = DocHeight
End Sub

Private Sub HScroll1_Change()
    'Set picture box left to horizontal scroll bar value
    Picture2.Left = -HScroll1.Value
    'Set shadow shape left to horizontal scroll bar value + offset
    Shape1.Left = -HScroll1.Value + 100
End Sub

Private Sub HScroll1_Scroll()
    'Horizontal scroll when dragging scroll bar
    HScroll1_Change
End Sub

Private Sub mnuAllignSub_Click(Index As Integer)
    For I = 0 To mnuAllignSub.Count - 1
        mnuAllignSub(I).Checked = False
    Next
    mnuAllignSub(Index).Checked = True
    prtObj.Cls
    'Print report data on the print object
    PrintReport
End Sub

Private Sub mnuFileSub_Click(Index As Integer)
    Select Case mnuFileSub(Index).Caption
        Case "P&rint Doc"
            Text1.Visible = False
            'Set print object to printer
            Set prtObj = Printer
            If UBound(DataArray, 2) = 0 Then LoadData
            'If the report page setup
            PageSetup
            'Print Report Data
            PrintReport
            'Force last page from printer
            prtObj.EndDoc
            Set prtObj = Picture2
        Case "&Preview"
            mnuFileSub(0).Enabled = True
            'Set the print object to the picture box
            Set prtObj = Picture2
            'Resize it to the printer for a WYSIWYG preview
            prtObj.Width = DocWidth
            prtObj.Height = DocHeight
            'Show the previewer picture boxes and scroll bars that are in
            'the frame1 container
            Frame1.Visible = True
            'Clear the picture box this is not a printer method
            prtObj.Cls
            'If the report page setup
            PageSetup
            LoadData
            'Set the Scroll Bar Values
            SetScrollBars prtObj, Picture1
            'Print report data on the print object
            PrintReport
        Case "E&xit"
            End
    End Select
End Sub


Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Text1.Visible = False
    If Text1 <> "" Then
        prtObj.Line (Text1.Left, Text1.Top)-(Text1.Left + Text1.Width, Text1.Top + Text1.Height), prtObj.BackColor, BF
        'Print report data on the print object
        PrintReport
    End If
    Text1 = ""
    Text1.Font = prtObj.Font
    Text1.FontSize = prtObj.FontSize
    CellLocate X, Y
    Text1.Left = ColLeft(Col)
    Text1.Width = ColWidth(Col)
    Text1.Height = RowHeight
    Text1.Top = (Row * RowHeight) + BodyTop
    Text1 = DataArray(Row, Col)
    Text1.Visible = True
    Text1.SetFocus
    Text1.SelLength = Len(Trim(Text1))
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not Text1.Visible Then
        'Calculate cursor location cell
        If Y > BodyTop And Y < RowBottom And X > LeftMargin And X < RightMargin Then
            CellLocate X, Y
            'Display location in caption
            Caption = "Row = " & Row & "   Col = " & Col & "   Data in Cell - " & DataArray(Row, Col)
        End If
    End If
End Sub

Private Sub Text1_Change()
    If Text1.Visible Then DataArray(Row, Col) = Text1
End Sub

Private Sub Text1_LostFocus()
    If Text1.Visible Then
        Text1.Visible = False
        If Text1 <> "" Then
            prtObj.Line (Text1.Left, Text1.Top)-(Text1.Left + Text1.Width, Text1.Top + Text1.Height), prtObj.BackColor, BF
            'Print report data on the print object
            PrintReport
        End If
    End If
End Sub

Private Sub VScroll1_Change()
    'Position the picture box to the vertical scroll value
    Picture2.Top = -VScroll1.Value
    Shape1.Top = -VScroll1.Value + 100
End Sub

Private Sub VScroll1_Scroll()
    VScroll1_Change
End Sub
