VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FormPrintPreview 
   Caption         =   "Предварительный просмотр"
   ClientHeight    =   6264
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   7140
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   7.8
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6264
   ScaleWidth      =   7140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox PictureBuffer 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   1332
      Left            =   5760
      ScaleHeight     =   1308
      ScaleWidth      =   1308
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2520
      Width           =   1332
   End
   Begin VB.PictureBox PicturePreview 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5292
      Left            =   1200
      ScaleHeight     =   5268
      ScaleWidth      =   4308
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   600
      Width           =   4332
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7140
      _ExtentX        =   12594
      _ExtentY        =   529
      ButtonWidth     =   1926
      ButtonHeight    =   487
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Обновить"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Печать"
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Shape ShapeShadow 
      BackColor       =   &H80000010&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   5292
      Left            =   1320
      Top             =   720
      Width           =   4332
   End
End
Attribute VB_Name = "FormPrintPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    '<EhHeader>
    On Error GoTo Toolbar_ButtonClick_Err
    '</EhHeader>

    If (Button.index = 1) Then
    
         Dim dRatio As Double
         
         dRatio = ScalePicPreviewToPrinterInches(PicturePreview)
         
         PrintRoutine2 PicturePreview, dRatio
                  
    End If

    If (Button.index = 2) Then
             
         Printer.ScaleMode = vbInches
         
         PrintRoutine2 Printer
         
         Printer.EndDoc
                  
    End If
    
    '<EhFooter>
    Exit Sub

Toolbar_ButtonClick_Err:
    Logger.Info "[cop.FormPrintPreview.Toolbar_ButtonClick]: " & GetErrorMessageById( _
            Err.Number, Err.Description)

    Resume Next

    '</EhFooter>
End Sub

Private Function ScalePicPreviewToPrinterInches _
         (picPreview As PictureBox) As Double
    '<EhHeader>
    On Error GoTo ScalePicPreviewToPrinterInches_Err
    '</EhHeader>

         Dim Ratio As Double ' Ratio between Printer and Picture
         Dim LRGap As Double, TBGap As Double
         Dim HeightRatio As Double, WidthRatio As Double
         Dim PgWidth As Double, PgHeight As Double
         Dim smtemp As Long

         ' Get the physical page size in Inches:
         PgWidth = Printer.Width / 1440
         PgHeight = Printer.Height / 1440

         ' Find the size of the non-printable area on the printer to
         ' use to offset coordinates. These formulas assume the
         ' printable area is centered on the page:
         smtemp = Printer.ScaleMode
         Printer.ScaleMode = vbInches
         LRGap = (PgWidth - Printer.ScaleWidth) / 2
         TBGap = (PgHeight - Printer.ScaleHeight) / 2
         Printer.ScaleMode = smtemp

         ' Scale PictureBox to Printer's printable area in Inches:
         picPreview.ScaleMode = vbInches

         ' Compare the height and with ratios to determine the
         ' Ratio to use and how to size the picture box:
         HeightRatio = picPreview.ScaleHeight / PgHeight
         WidthRatio = picPreview.ScaleWidth / PgWidth

         If HeightRatio < WidthRatio Then
         
            Ratio = HeightRatio
            smtemp = picPreview.Container.ScaleMode
            picPreview.Container.ScaleMode = vbInches
            picPreview.Width = PgWidth * Ratio
            picPreview.Container.ScaleMode = smtemp
            
         Else
         
            Ratio = WidthRatio
            smtemp = picPreview.Container.ScaleMode
            picPreview.Container.ScaleMode = vbInches
            picPreview.Height = PgHeight * Ratio
            picPreview.Container.ScaleMode = smtemp
            
         End If

         ' Set default properties of picture box to match printer
         ' There are many that you could add here:
         picPreview.Scale (0, 0)-(PgWidth, PgHeight)
         picPreview.Font.Name = Printer.Font.Name
         picPreview.FontSize = Printer.FontSize * Ratio
         picPreview.ForeColor = Printer.ForeColor
         picPreview.Cls

         ScalePicPreviewToPrinterInches = Ratio
         
    '<EhFooter>
    Exit Function

ScalePicPreviewToPrinterInches_Err:
    Logger.Info "[cop.FormPrintPreview.ScalePicPreviewToPrinterInches]: " & GetErrorMessageById( _
            Err.Number, Err.Description)

    Resume Next

    '</EhFooter>
End Function

Private Sub PrintRoutine(objPrint As Object, _
                               Optional Ratio As Double = 1)
    '<EhHeader>
    On Error GoTo PrintRoutine_Err
    '</EhHeader>
         
         
         
         ' All dimensions in inches:

         ' Print some graphics to the control object
         objPrint.Line (1, 1)-(1 + 6.5, 1 + 9), , B
         objPrint.Line (1.1, 2)-(1.1, 2)
         
         objPrint.Line (2.1, 1.2)-(2.1 + 5.2, 1.2 + 0.7), _
                        RGB(200, 200, 200), BF

         ' Print a title
         With objPrint
         
            .Font.Name = "Arial"
            .CurrentX = 2.3
            .CurrentY = 1.3
            .FontSize = 35 * Ratio
            objPrint.Print "Visual Basic Printing"
            
         End With

         ' Print some circles
         Dim X As Single
         
         For X = 3 To 5.5 Step 0.2
         
            objPrint.Circle (X, 3.5), 0.75
            
         Next

         ' Print some text
         With objPrint
         
            .Font.Name = "Courier New"
            .FontSize = 30 * Ratio
            .CurrentX = 1.5
            .CurrentY = 5
            objPrint.Print "It is possible to do"

            .FontSize = 24 * Ratio
            .CurrentX = 1.5
            .CurrentY = 6.5
            objPrint.Print "It is possible to do print"

            .FontSize = 18 * Ratio
            .CurrentX = 1.5
            .CurrentY = 8
            objPrint.Print "It is possible to do print preview"
            
         End With
         
    '<EhFooter>
    Exit Sub

PrintRoutine_Err:
    Logger.Info "[cop.FormPrintPreview.PrintRoutine]: " & GetErrorMessageById( _
            Err.Number, Err.Description)

    Resume Next

    '</EhFooter>
End Sub

Private Sub PrintRoutine2(objPrint As Object, _
                               Optional Ratio As Double = 1)
    '<EhHeader>
    On Error GoTo PrintRoutine2_Err
    '</EhHeader>
         
    Const intLINE_START_POS As Integer = 6
    Const intLINES_PER_PAGE As Integer = 60

    Dim intLineCtr As Integer
    Dim intPageCtr As Integer
    Dim intX As Integer
    Dim strCustFileName As String
    Dim strBackSlash As String
    Dim intCustFileNbr As Integer
    
    Dim strFirstName As String
    Dim strLastName As String
    Dim strAddr As String
    Dim strCity As String
    Dim strState As String
    Dim strZip As String
 
    Dim hSourceDC As Long
 
    PictureBuffer.Cls
       
    hSourceDC = GetDC(FormMain.StepGrid(0).hWnd)

    BitBlt PictureBuffer.hdc, _
        0, 0, _
        FormMain.StepGrid(0).Width, FormMain.StepGrid(0).Height, _
        hSourceDC, _
        0, 0, _
        vbSrcCopy

    ReleaseDC FormMain.StepGrid(0).hWnd, hSourceDC
        
    PictureBuffer.PSet (100, 100), RGB(192, 192, 192)
    PictureBuffer.Circle (100, 100), 50, RGB(192, 192, 192)
    
    objPrint.PaintPicture PictureBuffer.Image, 0!, 0!, PictureBuffer.Width, _
        PictureBuffer.Height, _
        , , , , vbSrcCopy _

    Exit Sub
 
    With objPrint

        ' Set the printer font to Courier, if available (otherwise, we would be
        ' relying on the default font for the Windows printer, which may or
        ' may not be set to an appropriate font) ...

        .FontName = "Courier New"
        '.FontSize = 10

        ' initialize report variables ...
        intPageCtr = 0
        intLineCtr = 99 ' initialize line counter to an arbitrarily high number
 
        ' to force the first page break
        ' prepare file name & number
        strBackSlash = IIf(Right$(App.Path, 1) = "\", "", "\")
        strCustFileName = App.Path & strBackSlash & "customer.txt"
 
        ' Print 4 blank lines, which provides a for top margin. These four lines do NOT
        ' count toward the limit of 60 lines.

        objPrint.Print
        objPrint.Print
        objPrint.Print
        objPrint.Print

        ' Print the main headings
        
        objPrint.Print Tab(intLINE_START_POS); _
        "Print Date: "; _
        Format$(Date, "mm/dd/yy"); _
        Tab(intLINE_START_POS + 31); _
        "THE VBPROGRAMMER.COM"; _
        Tab(intLINE_START_POS + 73); _
        "Page:"; _
        Format$(intPageCtr, "@@@")
        
        objPrint.Print Tab(intLINE_START_POS); _
        "Print Time: "; _
        Format$(Time, "hh:nn:ss"); _
        Tab(intLINE_START_POS + 33); _
        "CUSTOMER LIST"

        objPrint.Print

        ' Print the column headings
        
        objPrint.Print Tab(intLINE_START_POS); _
        "CUSTOMER NAME"; _
        Tab(21 + intLINE_START_POS); _
        "ADDRESS"; _
        Tab(48 + intLINE_START_POS); _
        "CITY"; _
        Tab(72 + intLINE_START_POS); _
        "ST"; _
        Tab(76 + intLINE_START_POS); _
        "ZIP"

        objPrint.Print Tab(intLINE_START_POS); _
        "-------------"; _
        Tab(21 + intLINE_START_POS); _
        "-------"; _
        Tab(48 + intLINE_START_POS); _
        "----"; _
        Tab(72 + intLINE_START_POS); _
        "--"; _
        Tab(76 + intLINE_START_POS); _
        "---"
        

            
    End With
         
    '<EhFooter>
    Exit Sub

PrintRoutine2_Err:
    Logger.Info "[cop.FormPrintPreview.PrintRoutine2]: " & GetErrorMessageById( _
            Err.Number, Err.Description)

    Resume Next

    '</EhFooter>
End Sub


