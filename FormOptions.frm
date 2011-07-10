VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FormOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Настройки"
   ClientHeight    =   4908
   ClientLeft      =   2568
   ClientTop       =   1500
   ClientWidth     =   6132
   Icon            =   "FormOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4908
   ScaleWidth      =   6132
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab SSTab 
      Height          =   4212
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5892
      _ExtentX        =   10393
      _ExtentY        =   7430
      _Version        =   393216
      Style           =   1
      TabHeight       =   420
      TabCaption(0)   =   "Основные"
      TabPicture(0)   =   "FormOptions.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "picOptions(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Лог "
      TabPicture(1)   =   "FormOptions.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "picOptions(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Другое"
      TabPicture(2)   =   "FormOptions.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "picOptions(2)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VB.PictureBox picOptions 
         BorderStyle     =   0  'None
         Height          =   3708
         Index           =   2
         Left            =   -74880
         ScaleHeight     =   3708
         ScaleWidth      =   5688
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   360
         Width           =   5688
         Begin VB.Frame FrameOptionsFilesHistory 
            Caption         =   "История файлов"
            Height          =   1188
            Left            =   0
            TabIndex        =   27
            Top             =   0
            Width           =   5676
            Begin VB.TextBox TextFilesHistoryCount 
               Alignment       =   1  'Right Justify
               Height          =   288
               Left            =   1680
               MaxLength       =   2
               TabIndex        =   30
               Text            =   "4"
               Top             =   312
               Width           =   408
            End
            Begin VB.CheckBox CheckFilesHistoryLimitPaths 
               Caption         =   "Ограничивать длину пути в меню"
               Height          =   252
               Left            =   120
               TabIndex        =   29
               Top             =   720
               Width           =   4932
            End
            Begin MSComCtl2.UpDown UpDownFilesHistoryCount 
               Height          =   288
               Left            =   2088
               TabIndex        =   28
               Top             =   312
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   508
               _Version        =   393216
               Value           =   4
               BuddyControl    =   "TextFilesHistoryCount"
               BuddyDispid     =   196627
               OrigLeft        =   2028
               OrigTop         =   300
               OrigRight       =   2268
               OrigBottom      =   612
               Min             =   1
               SyncBuddy       =   -1  'True
               BuddyProperty   =   65547
               Enabled         =   -1  'True
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Помнить не более"
               Height          =   192
               Left            =   120
               TabIndex        =   32
               Top             =   360
               Width           =   1428
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "файлов"
               Height          =   192
               Left            =   2424
               TabIndex        =   31
               Top             =   360
               Width           =   612
            End
         End
      End
      Begin VB.PictureBox picOptions 
         BorderStyle     =   0  'None
         Height          =   3732
         Index           =   1
         Left            =   -74880
         ScaleHeight     =   3732
         ScaleWidth      =   5688
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   360
         Width           =   5688
         Begin VB.Frame Frame4 
            Caption         =   "Параметры"
            Height          =   3708
            Left            =   0
            TabIndex        =   21
            Top             =   0
            Width           =   5676
            Begin VB.CheckBox CheckRewriteLogFile 
               Caption         =   "Перезаписывать файл лога при запуске"
               Height          =   252
               Left            =   120
               TabIndex        =   24
               Top             =   360
               Width           =   3612
            End
            Begin VB.TextBox TextLogFilePath 
               Height          =   288
               Left            =   120
               TabIndex        =   23
               Top             =   1008
               Width           =   4212
            End
            Begin VB.CommandButton cmdBrowse 
               Caption         =   "Обзор..."
               Height          =   372
               Left            =   4440
               TabIndex        =   22
               Top             =   972
               Width           =   1092
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "Путь к файлу:"
               Height          =   192
               Left            =   120
               TabIndex        =   25
               Top             =   720
               Width           =   1092
            End
         End
      End
      Begin VB.PictureBox picOptions 
         BorderStyle     =   0  'None
         Height          =   3708
         Index           =   0
         Left            =   120
         ScaleHeight     =   3708
         ScaleWidth      =   5688
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   360
         Width           =   5688
         Begin VB.Frame Frame1 
            Caption         =   "Импорт/экспорт"
            Height          =   732
            Left            =   0
            TabIndex        =   15
            Top             =   2950
            Width           =   5676
            Begin VB.ComboBox Combo1 
               Height          =   288
               ItemData        =   "FormOptions.frx":0060
               Left            =   1380
               List            =   "FormOptions.frx":0067
               Style           =   2  'Dropdown List
               TabIndex        =   17
               Top             =   240
               Width           =   852
            End
            Begin VB.ComboBox Combo2 
               Height          =   288
               ItemData        =   "FormOptions.frx":0071
               Left            =   4440
               List            =   "FormOptions.frx":0078
               Style           =   2  'Dropdown List
               TabIndex        =   16
               Top             =   240
               Width           =   852
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "Импорт из:"
               Height          =   192
               Left            =   240
               TabIndex        =   19
               Top             =   288
               Width           =   864
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "Экспорт в:"
               Height          =   192
               Left            =   3336
               TabIndex        =   18
               Top             =   288
               Width           =   828
            End
         End
         Begin VB.Frame fraSample1 
            Caption         =   "Вид"
            Height          =   2868
            Left            =   0
            TabIndex        =   4
            Top             =   0
            Width           =   5676
            Begin VB.TextBox Text1 
               Height          =   288
               Left            =   2040
               TabIndex        =   10
               Top             =   264
               Width           =   612
            End
            Begin VB.TextBox Text2 
               Height          =   288
               Left            =   2040
               TabIndex        =   9
               Top             =   684
               Width           =   612
            End
            Begin VB.CommandButton cmdFont 
               Caption         =   "Шрифт"
               Height          =   372
               Left            =   2880
               TabIndex        =   8
               Top             =   720
               Width           =   972
            End
            Begin VB.Frame Frame2 
               Height          =   492
               Left            =   2880
               TabIndex        =   6
               Top             =   120
               Width           =   2652
               Begin VB.Label LabelFont 
                  AutoSize        =   -1  'True
                  Caption         =   "FontName"
                  Height          =   192
                  Left            =   120
                  TabIndex        =   7
                  Top             =   192
                  Width           =   756
               End
            End
            Begin VB.CommandButton cmdApply 
               Caption         =   "Применить"
               Height          =   372
               Left            =   4440
               TabIndex        =   5
               Top             =   720
               Width           =   1092
            End
            Begin MSFlexGridLib.MSFlexGrid StepsView 
               Height          =   1452
               Left            =   120
               TabIndex        =   11
               TabStop         =   0   'False
               Top             =   1320
               Width           =   5412
               _ExtentX        =   9546
               _ExtentY        =   2561
               _Version        =   393216
               Rows            =   16
               Cols            =   81
               AllowBigSelection=   0   'False
               AllowUserResizing=   1
               BorderStyle     =   0
               Appearance      =   0
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Ширина столбца шага:"
               Height          =   192
               Left            =   120
               TabIndex        =   14
               Top             =   360
               Width           =   1740
            End
            Begin VB.Label Label2 
               Caption         =   "Высота строк таблиц:"
               Height          =   252
               Left            =   120
               TabIndex        =   13
               Top             =   720
               Width           =   1812
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Образец:"
               Height          =   192
               Left            =   120
               TabIndex        =   12
               Top             =   1080
               Width           =   732
            End
         End
      End
   End
   Begin MSComDlg.CommonDialog SaveFileDialog 
      Left            =   600
      Top             =   4440
      _ExtentX        =   677
      _ExtentY        =   677
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog FontDialog 
      Left            =   120
      Top             =   4440
      _ExtentX        =   677
      _ExtentY        =   677
      _Version        =   393216
      DialogTitle     =   "Шрифт"
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Отмена"
      Height          =   375
      Left            =   4920
      TabIndex        =   1
      Top             =   4455
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   3696
      TabIndex        =   0
      Top             =   4455
      Width           =   1095
   End
End
Attribute VB_Name = "FormOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**
'@author <a href="mailto:unihomelab@ya.ru">Мезенцев Вячеслав</a>
'@revision Дата ревизии: 16.06.2011 г., время: 4:19:56
'@rem <h1><b>FormOptions</b></h1>
'<pre>
'--------------------------------------------------------------------------------
' Проект   :       Конфигуратор управляющих программ
' Модуль   :       FormOptions
' Описание :       Диалоговая форма редактирования настроек программы
' Автор    :       Мезенцев Вячеслав
' Изменён  :       16.06.2011 г., время: 4:19:56
'--------------------------------------------------------------------------------
'</pre>
Option Explicit

'**
'@rem <h2>cmdApply_Click</h2>
'Обработчик кнопки "Применить". При её нажатии на образце таблицы будут показаны
'результаты сделанных изменений в настройках интерфейса.
Private Sub cmdApply_Click()
    '<EhHeader>
    On Error GoTo cmdApply_Click_Err
    '</EhHeader>
    
    RefreshStepsView
    
    '<EhFooter>
    Exit Sub

cmdApply_Click_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & " [INFO] [cop.FormOptions.cmdApply_Click]: " _
       & GetErrorMessageById(Err.Number, Err.Description), VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
    Resume Next
    '</EhFooter>
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

'**
'@rem <h2>cmdFont_Click</h2>
'Обработчик кнопки "Шрифт". Вызывается стандартное окно выбора параметров
'шрифта.
Private Sub cmdFont_Click()
    '<EhHeader>
    On Error GoTo cmdFont_Click_Err
    '</EhHeader>

    FontDialog.FontBold = Settings.StepsViewFontBold
    FontDialog.FontItalic = Settings.StepsViewFontItalic
    FontDialog.FontName = Settings.StepsViewFontName
    FontDialog.FontSize = Settings.StepsViewFontSize
    FontDialog.Flags = cdlCFBoth
    
    FontDialog.ShowFont
    
    LabelFont.FontBold = FontDialog.FontBold
    LabelFont.FontItalic = FontDialog.FontItalic
    LabelFont.FontName = FontDialog.FontName
    LabelFont.FontSize = FontDialog.FontSize
    LabelFont.Caption = LabelFont.FontName & ", " & CInt(LabelFont.FontSize)
    
    '<EhFooter>
    Exit Sub

cmdFont_Click_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & " [INFO] [cop.FormOptions.cmdFont_Click]: " _
       & GetErrorMessageById(Err.Number, Err.Description), VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
    Resume Next
    '</EhFooter>
End Sub

Private Sub cmdOK_Click()
    '<EhHeader>
    On Error GoTo cmdOK_Click_Err
    '</EhHeader>

    Settings.StepColWidth = CInt(Text1.Text)
    Settings.RowHeight = CInt(Text2.Text)
    
    Settings.StepsViewFontBold = LabelFont.FontBold
    Settings.StepsViewFontItalic = LabelFont.FontItalic
    Settings.StepsViewFontName = LabelFont.FontName
    Settings.StepsViewFontSize = LabelFont.FontSize
    
    Settings.RewriteLogFile = CheckRewriteLogFile.Value > 0
    Settings.LogFilePath = TextLogFilePath.Text
    
    Settings.FilesHistoryLimitPaths = CheckFilesHistoryLimitPaths.Value > 0
    MRUFileList.MaxFileCount = CInt(TextFilesHistoryCount.Text)

    Settings.SaveSettings
    
    FormMain.RefreshComponents False
    Unload Me
    
    '<EhFooter>
    Exit Sub

cmdOK_Click_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & " [INFO] [cop.FormOptions.cmdOK_Click]: " _
       & GetErrorMessageById(Err.Number, Err.Description), VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
    Resume Next
    '</EhFooter>
End Sub

Private Sub cmdBrowse_Click()
    '<EhHeader>
    On Error GoTo cmdBrowse_Click_Err
    '</EhHeader>

    SaveFileDialog.FileName = Settings.LogFilePath
    SaveFileDialog.DialogTitle = "Обзор..."
    SaveFileDialog.DefaultExt = ".log"
    SaveFileDialog.Filter = "Файл лога (*.log)|*.log|Все файлы (*.*)|(*.*)"
    SaveFileDialog.FilterIndex = 1
    SaveFileDialog.MaxFileSize = 32767
    SaveFileDialog.InitDir = CurrentDir
    SaveFileDialog.CancelError = True
    
    SaveFileDialog.ShowSave

    TextLogFilePath.Text = SaveFileDialog.FileName
    
    '<EhFooter>
    Exit Sub

cmdBrowse_Click_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & " [INFO] [cop.FormOptions.cmdBrowse_Click]: " _
        & GetErrorMessageById(Err.Number, Err.Description), VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
    Resume Next
    '</EhFooter>
End Sub

Private Sub RefreshStepsView()
    '<EhHeader>
    On Error GoTo RefreshStepsView_Err
    '</EhHeader>

    Dim s As String
    Dim col As Integer, row As Integer
    
    StepsView.Visible = False
    
    StepsView.Font.Bold = LabelFont.FontBold
    StepsView.Font.Italic = LabelFont.FontItalic
    StepsView.Font.Name = LabelFont.FontName
    StepsView.Font.Size = LabelFont.FontSize
    
    StepsView.Cols = 10 + 1
    
    s = "<   |"

    For col = 1 To StepsView.Cols - 1
        
        If col < StepsView.Cols - 1 Then

            If col < 10 Then
                s = s & "0" & col & "|"
            Else
                s = s & col & "|"
            End If
        Else
            s = s & col
        End If
        StepsView.col = col
        StepsView.row = 0
        StepsView.CellAlignment = flexAlignCenterCenter
    Next
    
    StepsView.FormatString = s
       
    s = ";|" _
       & "Клапан горячей воды" & "|" _
       & "Клапан холодной воды 1" & "|" _
       & "Клапан холодной воды 2" & "|" _
       & "Клапан МС 1" & "|" _
       & "Клапан МС 2" & "|" _
       & "Клапан МС 3" & "|" _
       & "Клапан МС 4" & "|" _
       & "Клапан МС 5" & "|" _
       & "Клапан МС 6" & "|" _
       & "Клапан МС 7" & "|" _
       & "Клапан МС 8" & "|" _
       & "Клапан МС 9" & "|" _
       & "Замок люка 1" & "|" _
       & "Замок люка 2" & "|" _
       & "Слив 1" & "|" _
       & "Слив 2" & "|" _
       & "Нагрев" & "|" _
       & "Мотор"
    
    StepsView.FormatString = s
    
    ' "Тушим" все ячейки таблицы

    For row = 1 To StepsView.rows - 1
        StepsView.RowHeight(row) = CInt(Text2.Text)
        
        For col = 1 To 10
            StepsView.ColWidth(col) = CInt(Text1.Text)
            StepsView.col = col
            StepsView.row = row
            StepsView.CellBackColor = &H8000000F
        Next
    Next
    
    StepsView.col = 1
    StepsView.row = 1
    
    StepsView.Visible = True
    
    '<EhFooter>
    Exit Sub

RefreshStepsView_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & " [INFO] [cop.FormOptions.RefreshStepsView]: " _
       & GetErrorMessageById(Err.Number, Err.Description), VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
    Resume Next
    '</EhFooter>
End Sub

Private Sub Form_Load()
    '<EhHeader>
    On Error GoTo Form_Load_Err
    '</EhHeader>

    Text1.Text = "" & Settings.StepColWidth
    Text2.Text = "" & Settings.RowHeight
    
    Combo1.ListIndex = 0
    Combo2.ListIndex = 0
    
    LabelFont.FontBold = Settings.StepsViewFontBold
    LabelFont.FontItalic = Settings.StepsViewFontItalic
    LabelFont.FontName = Settings.StepsViewFontName
    LabelFont.FontSize = Settings.StepsViewFontSize
    LabelFont.Caption = LabelFont.FontName & ", " & CInt(LabelFont.FontSize)
    
    TextFilesHistoryCount.Text = MRUFileList.MaxFileCount
    
    Select Case Settings.FilesHistoryLimitPaths
        Case False: CheckFilesHistoryLimitPaths.Value = 0
        Case True: CheckFilesHistoryLimitPaths.Value = 1
    End Select
    
    Select Case Settings.RewriteLogFile
        Case False: CheckRewriteLogFile.Value = 0
        Case True: CheckRewriteLogFile.Value = 1
    End Select

    TextLogFilePath.Text = Settings.LogFilePath
    
    RefreshStepsView
    
    'center the form
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    
    '<EhFooter>
    Exit Sub

Form_Load_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & " [INFO] [cop.FormOptions.Form_Load]: " _
        & GetErrorMessageById(Err.Number, Err.Description), VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
    Resume Next
    '</EhFooter>
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    '<EhHeader>
    On Error GoTo Text1_KeyPress_Err
    '</EhHeader>

    If KeyAscii = VBRUN.KeyCodeConstants.vbKeyReturn Then KeyAscii = 0
    
    '<EhFooter>
    Exit Sub

Text1_KeyPress_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & " [INFO] [cop.FormOptions.Text1_KeyPress]: " _
       & GetErrorMessageById(Err.Number, Err.Description), VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
    Resume Next
    '</EhFooter>
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    '<EhHeader>
    On Error GoTo Text2_KeyPress_Err
    '</EhHeader>

    If KeyAscii = VBRUN.KeyCodeConstants.vbKeyReturn Then KeyAscii = 0
    
    '<EhFooter>
    Exit Sub

Text2_KeyPress_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & " [INFO] [cop.FormOptions.Text2_KeyPress]: " _
       & GetErrorMessageById(Err.Number, Err.Description), VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
    Resume Next
    '</EhFooter>
End Sub

'**
'@rem <h2>TextFilesHistoryCount_Change</h2>
Private Sub TextFilesHistoryCount_Change()
    '<EhHeader>
    On Error GoTo TextFilesHistoryCount_Change_Err
    '</EhHeader>
    
    If CInt(TextFilesHistoryCount.Text) > 10 Then
        TextFilesHistoryCount.Text = "10"
    ElseIf CInt(TextFilesHistoryCount.Text) < 1 Then
        TextFilesHistoryCount.Text = "1"
    End If
    '<EhFooter>
    Exit Sub

TextFilesHistoryCount_Change_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & " [INFO] [cop.FormOptions.TextFilesHistoryCount_Change]: " _
        & GetErrorMessageById(Err.Number, Err.Description), VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
    Resume Next
    '</EhFooter>
End Sub
