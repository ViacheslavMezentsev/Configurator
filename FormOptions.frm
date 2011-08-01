VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
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
      TabCaption(0)   =   "Вид "
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
      TabCaption(2)   =   "Другое "
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
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   360
         Width           =   5688
         Begin VB.Frame Frame1 
            Caption         =   "Импорт/экспорт"
            Height          =   732
            Left            =   0
            TabIndex        =   32
            Top             =   2400
            Width           =   5676
            Begin VB.ComboBox ComboExportFormat 
               Height          =   288
               ItemData        =   "FormOptions.frx":0060
               Left            =   4440
               List            =   "FormOptions.frx":0067
               Style           =   2  'Dropdown List
               TabIndex        =   34
               Top             =   240
               Width           =   852
            End
            Begin VB.ComboBox ComboImportFormat 
               Height          =   288
               ItemData        =   "FormOptions.frx":0071
               Left            =   1380
               List            =   "FormOptions.frx":0078
               Style           =   2  'Dropdown List
               TabIndex        =   33
               Top             =   240
               Width           =   852
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Экспорт в:"
               Height          =   192
               Left            =   3336
               TabIndex        =   36
               Top             =   288
               Width           =   828
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Импорт из:"
               Height          =   192
               Left            =   240
               TabIndex        =   35
               Top             =   288
               Width           =   864
            End
         End
         Begin VB.Frame FrameOptionsFilesHistory 
            Caption         =   "История файлов"
            Height          =   1188
            Left            =   0
            TabIndex        =   26
            Top             =   0
            Width           =   5676
            Begin VB.CheckBox CheckFilesHistoryLimitPaths 
               Caption         =   "Ограничивать длину пути в меню"
               Height          =   252
               Left            =   120
               TabIndex        =   28
               Top             =   720
               Width           =   4932
            End
            Begin VB.TextBox TextFilesHistoryCount 
               Alignment       =   1  'Right Justify
               Height          =   288
               Left            =   1680
               MaxLength       =   2
               TabIndex        =   27
               Text            =   "4"
               Top             =   312
               Width           =   408
            End
            Begin MSComCtl2.UpDown UpDownFilesHistoryCount 
               Height          =   288
               Left            =   2088
               TabIndex        =   29
               Top             =   312
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   508
               _Version        =   393216
               Value           =   4
               BuddyControl    =   "TextFilesHistoryCount"
               BuddyDispid     =   196617
               OrigLeft        =   2028
               OrigTop         =   300
               OrigRight       =   2268
               OrigBottom      =   612
               Min             =   1
               SyncBuddy       =   -1  'True
               BuddyProperty   =   65547
               Enabled         =   -1  'True
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "файлов"
               Height          =   192
               Left            =   2424
               TabIndex        =   31
               Top             =   360
               Width           =   612
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Помнить не более"
               Height          =   192
               Left            =   120
               TabIndex        =   30
               Top             =   360
               Width           =   1428
            End
         End
         Begin VB.Frame FrameUpdateOptions 
            Caption         =   "Обновление"
            Height          =   972
            Left            =   0
            TabIndex        =   22
            Top             =   1320
            Width           =   5676
            Begin VB.CheckBox CheckEnableAutoUpdate 
               Caption         =   "Включить автоматическое обновление"
               Height          =   252
               Left            =   120
               TabIndex        =   24
               Top             =   240
               Width           =   5412
            End
            Begin VB.ComboBox ComboAutoUpdatePeriod 
               Enabled         =   0   'False
               Height          =   288
               ItemData        =   "FormOptions.frx":0082
               Left            =   360
               List            =   "FormOptions.frx":008F
               Style           =   2  'Dropdown List
               TabIndex        =   23
               Top             =   552
               Width           =   1692
            End
            Begin VB.Label LabelAutoUpdatePeriod 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "период автообновления"
               Enabled         =   0   'False
               Height          =   192
               Left            =   2160
               TabIndex        =   25
               Top             =   600
               Width           =   1932
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
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   360
         Width           =   5688
         Begin VB.Frame Frame4 
            Caption         =   "Параметры"
            Height          =   3708
            Left            =   0
            TabIndex        =   15
            Top             =   0
            Width           =   5676
            Begin VB.CheckBox CheckRewriteLogFile 
               Caption         =   "Перезаписывать файл лога при запуске"
               Height          =   252
               Left            =   120
               TabIndex        =   18
               Top             =   360
               Width           =   3612
            End
            Begin VB.TextBox TextLogFilePath 
               Height          =   288
               Left            =   120
               TabIndex        =   17
               Top             =   1008
               Width           =   4212
            End
            Begin VB.CommandButton cmdBrowse 
               Caption         =   "Обзор..."
               Height          =   372
               Left            =   4440
               TabIndex        =   16
               Top             =   972
               Width           =   1092
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "Путь к файлу:"
               Height          =   192
               Left            =   120
               TabIndex        =   19
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
         Begin VB.Frame fraSample1 
            Caption         =   "Вид"
            Height          =   3708
            Left            =   0
            TabIndex        =   21
            Top             =   0
            Width           =   5676
            Begin VB.PictureBox PictureHSelRight 
               AutoRedraw      =   -1  'True
               BackColor       =   &H8000000D&
               BorderStyle     =   0  'None
               Height          =   252
               Left            =   5400
               ScaleHeight     =   252
               ScaleWidth      =   24
               TabIndex        =   46
               TabStop         =   0   'False
               Top             =   2520
               Visible         =   0   'False
               Width           =   24
            End
            Begin VB.PictureBox PictureHSelBottom 
               AutoRedraw      =   -1  'True
               BackColor       =   &H8000000D&
               BorderStyle     =   0  'None
               Height          =   20
               Left            =   120
               ScaleHeight     =   24
               ScaleWidth      =   5316
               TabIndex        =   45
               TabStop         =   0   'False
               Top             =   2760
               Visible         =   0   'False
               Width           =   5316
            End
            Begin VB.PictureBox PictureHSelLeft 
               AutoRedraw      =   -1  'True
               BackColor       =   &H8000000D&
               BorderStyle     =   0  'None
               Height          =   252
               Left            =   120
               ScaleHeight     =   252
               ScaleWidth      =   24
               TabIndex        =   44
               TabStop         =   0   'False
               Top             =   2520
               Visible         =   0   'False
               Width           =   24
            End
            Begin VB.PictureBox PictureHSelTop 
               AutoRedraw      =   -1  'True
               BackColor       =   &H8000000D&
               BorderStyle     =   0  'None
               Height          =   20
               Left            =   120
               ScaleHeight     =   24
               ScaleWidth      =   5316
               TabIndex        =   43
               TabStop         =   0   'False
               Top             =   2520
               Visible         =   0   'False
               Width           =   5316
            End
            Begin VB.PictureBox PictureVSelBottom 
               AutoRedraw      =   -1  'True
               BackColor       =   &H8000000D&
               BorderStyle     =   0  'None
               Height          =   20
               Left            =   2760
               ScaleHeight     =   24
               ScaleWidth      =   396
               TabIndex        =   42
               TabStop         =   0   'False
               Top             =   3600
               Visible         =   0   'False
               Width           =   400
            End
            Begin VB.PictureBox PictureVSelTop 
               AutoRedraw      =   -1  'True
               BackColor       =   &H8000000D&
               BorderStyle     =   0  'None
               Height          =   20
               Left            =   2760
               ScaleHeight     =   24
               ScaleWidth      =   396
               TabIndex        =   41
               TabStop         =   0   'False
               Top             =   1800
               Visible         =   0   'False
               Width           =   400
            End
            Begin VB.PictureBox PictureVSelRight 
               AutoRedraw      =   -1  'True
               BackColor       =   &H8000000D&
               BorderStyle     =   0  'None
               Height          =   1812
               Left            =   3120
               ScaleHeight     =   1812
               ScaleWidth      =   24
               TabIndex        =   40
               TabStop         =   0   'False
               Top             =   1800
               Visible         =   0   'False
               Width           =   20
            End
            Begin VB.PictureBox PictureVSelLeft 
               AutoRedraw      =   -1  'True
               BackColor       =   &H8000000D&
               BorderStyle     =   0  'None
               Height          =   1812
               Left            =   2760
               ScaleHeight     =   1812
               ScaleWidth      =   24
               TabIndex        =   39
               TabStop         =   0   'False
               Top             =   1800
               Visible         =   0   'False
               Width           =   20
            End
            Begin VB.CheckBox CheckHSelector 
               Caption         =   "Горизонтальный селектор"
               Height          =   252
               Left            =   120
               TabIndex        =   38
               Top             =   1440
               Width           =   3132
            End
            Begin VB.CheckBox CheckVSelector 
               Caption         =   "Вертикальный селектор"
               Height          =   252
               Left            =   120
               TabIndex        =   37
               Top             =   1200
               Width           =   3252
            End
            Begin VB.TextBox Text1 
               Height          =   288
               Left            =   2040
               TabIndex        =   9
               Top             =   264
               Width           =   612
            End
            Begin VB.TextBox Text2 
               Height          =   288
               Left            =   2040
               TabIndex        =   8
               Top             =   684
               Width           =   612
            End
            Begin VB.CommandButton cmdFont 
               Caption         =   "Шрифт"
               Height          =   372
               Left            =   2880
               TabIndex        =   7
               Top             =   720
               Width           =   972
            End
            Begin VB.Frame Frame2 
               Height          =   492
               Left            =   2880
               TabIndex        =   5
               Top             =   120
               Width           =   2652
               Begin VB.Label LabelFont 
                  AutoSize        =   -1  'True
                  Caption         =   "FontName"
                  Height          =   192
                  Left            =   120
                  TabIndex        =   6
                  Top             =   192
                  Width           =   756
               End
            End
            Begin VB.CommandButton cmdApply 
               Caption         =   "Применить"
               Height          =   372
               Left            =   4440
               TabIndex        =   4
               Top             =   720
               Width           =   1092
            End
            Begin MSFlexGridLib.MSFlexGrid StepsView 
               Height          =   1452
               Left            =   120
               TabIndex        =   10
               TabStop         =   0   'False
               Top             =   2160
               Width           =   5412
               _ExtentX        =   9546
               _ExtentY        =   2561
               _Version        =   393216
               Rows            =   16
               Cols            =   81
               AllowBigSelection=   0   'False
               ScrollBars      =   0
               AllowUserResizing=   1
               BorderStyle     =   0
               Appearance      =   0
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Ширина столбца шага:"
               Height          =   192
               Left            =   120
               TabIndex        =   13
               Top             =   360
               Width           =   1740
            End
            Begin VB.Label Label2 
               Caption         =   "Высота строк таблиц:"
               Height          =   252
               Left            =   120
               TabIndex        =   12
               Top             =   720
               Width           =   1812
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Образец:"
               Height          =   192
               Left            =   120
               TabIndex        =   11
               Top             =   1920
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

Private Sub CheckEnableAutoUpdate_Click()

    ComboAutoUpdatePeriod.Enabled = CheckEnableAutoUpdate.Value > 0
    LabelAutoUpdatePeriod.Enabled = CheckEnableAutoUpdate.Value > 0

End Sub

Private Sub ShowHorizontalSelector()
    '<EhHeader>
    On Error GoTo ShowHorizontalSelector_Err
    '</EhHeader>

    ' Отображаем вертикальный селектор
    If CheckHSelector.Value > 0 Then
    
        PictureHSelLeft.Top = StepsView.Top + StepsView.RowPos(StepsView.RowSel) - Settings.StepsSelectorWidth / 2
        PictureHSelLeft.Left = StepsView.Left
        PictureHSelLeft.Width = Settings.StepsSelectorWidth
        PictureHSelLeft.Height = StepsView.RowHeight(StepsView.RowSel)
        
        PictureHSelRight.Top = PictureHSelLeft.Top
        PictureHSelRight.Left = PictureHSelLeft.Left + StepsView.ColWidth(0) + StepsView.ColWidth(1) * (StepsView.Cols - 1)
        PictureHSelRight.Height = PictureHSelLeft.Height
        PictureHSelRight.Width = Settings.StepsSelectorWidth
        
        PictureHSelTop.Left = PictureHSelLeft.Left
        PictureHSelTop.Top = PictureHSelLeft.Top
        PictureHSelTop.Height = Settings.StepsSelectorWidth
        PictureHSelTop.Width = PictureHSelRight.Left
        
        PictureHSelBottom.Left = PictureHSelLeft.Left
        PictureHSelBottom.Top = PictureHSelLeft.Top + PictureHSelLeft.Height
        PictureHSelBottom.Height = Settings.StepsSelectorWidth
        PictureHSelBottom.Width = PictureHSelTop.Width
        
        If StepsView.RowIsVisible(StepsView.RowSel) Then
        
            PictureHSelLeft.Visible = True
            PictureHSelRight.Visible = True
            PictureHSelTop.Visible = True
            PictureHSelBottom.Visible = True
        
        Else
        
            PictureHSelLeft.Visible = False
            PictureHSelRight.Visible = False
            PictureHSelTop.Visible = False
            PictureHSelBottom.Visible = False
        
        End If

    Else
        
        PictureHSelLeft.Visible = False
        PictureHSelRight.Visible = False
        PictureHSelTop.Visible = False
        PictureHSelBottom.Visible = False
            
    End If

    '<EhFooter>
    Exit Sub

ShowHorizontalSelector_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & _
            " [INFO] [cop.FormOptions.ShowHorizontalSelector]: " & GetErrorMessageById( _
            Err.Number, Err.Description), _
            VBRUN.LogEventTypeConstants.vbLogEventTypeInformation

    Resume Next

    '</EhFooter>
End Sub

Private Sub ShowVerticalSelector()
    '<EhHeader>
    On Error GoTo ShowVerticalSelector_Err
    '</EhHeader>

    ' Отображаем вертикальный селектор
    If CheckVSelector.Value > 0 Then
            
        PictureVSelLeft.Top = StepsView.Top
        PictureVSelLeft.Left = StepsView.Left + StepsView.ColPos(StepsView.ColSel) - Settings.StepsSelectorWidth / 2
        PictureVSelLeft.Width = Settings.StepsSelectorWidth
        PictureVSelLeft.Height = StepsView.RowHeight(StepsView.RowSel) * StepsView.rows
        
        PictureVSelRight.Top = PictureVSelLeft.Top
        PictureVSelRight.Left = PictureVSelLeft.Left + StepsView.ColWidth(StepsView.ColSel)
        PictureVSelRight.Height = PictureVSelLeft.Height
        PictureVSelRight.Width = Settings.StepsSelectorWidth
        
        PictureVSelTop.Left = PictureVSelLeft.Left
        PictureVSelTop.Top = StepsView.Top
        PictureVSelTop.Height = Settings.StepsSelectorWidth
        PictureVSelTop.Width = StepsView.ColWidth(StepsView.ColSel)
        
        PictureVSelBottom.Left = PictureVSelLeft.Left
        PictureVSelBottom.Top = PictureVSelLeft.Height - Settings.StepsSelectorWidth
        PictureVSelBottom.Height = Settings.StepsSelectorWidth
        PictureVSelBottom.Width = PictureVSelTop.Width
        
        If StepsView.ColIsVisible(StepsView.ColSel) Then
        
            PictureVSelLeft.Visible = True
            PictureVSelRight.Visible = True
            PictureVSelTop.Visible = True
            PictureVSelBottom.Visible = True
        
        Else
        
            PictureVSelLeft.Visible = False
            PictureVSelRight.Visible = False
            PictureVSelTop.Visible = False
            PictureVSelBottom.Visible = False
        
        End If
        
    Else
    
        PictureVSelLeft.Visible = False
        PictureVSelRight.Visible = False
        PictureVSelTop.Visible = False
        PictureVSelBottom.Visible = False
        
    End If

    '<EhFooter>
    Exit Sub

ShowVerticalSelector_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & _
            " [INFO] [cop.FormOptions.ShowVerticalSelector]: " & GetErrorMessageById( _
            Err.Number, Err.Description), _
            VBRUN.LogEventTypeConstants.vbLogEventTypeInformation

    Resume Next

    '</EhFooter>
End Sub

Private Sub CheckHSelector_Click()
    '<EhHeader>
    On Error GoTo CheckHSelector_Click_Err
    '</EhHeader>

    ' Отображаем горизонтальный селектор
    ShowHorizontalSelector
    
    '<EhFooter>
    Exit Sub

CheckHSelector_Click_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & _
            " [INFO] [cop.FormOptions.CheckHSelector_Click]: " & GetErrorMessageById( _
            Err.Number, Err.Description), _
            VBRUN.LogEventTypeConstants.vbLogEventTypeInformation

    Resume Next

    '</EhFooter>
End Sub

Private Sub CheckVSelector_Click()
    '<EhHeader>
    On Error GoTo CheckVSelector_Click_Err
    '</EhHeader>

    ' Отображаем вертикальный селектор
    ShowVerticalSelector
    
    '<EhFooter>
    Exit Sub

CheckVSelector_Click_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & _
            " [INFO] [cop.FormOptions.CheckVSelector_Click]: " & GetErrorMessageById( _
            Err.Number, Err.Description), _
            VBRUN.LogEventTypeConstants.vbLogEventTypeInformation

    Resume Next

    '</EhFooter>
End Sub

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

    Settings.StepsColWidth = CInt(Text1.Text)
    Settings.StepsRowHeight = CInt(Text2.Text)

    Settings.StepsVSelectorEnabled = CheckVSelector.Value > 0
    Settings.StepsHSelectorEnabled = CheckHSelector.Value > 0

    Settings.StepsViewFontBold = LabelFont.FontBold
    Settings.StepsViewFontItalic = LabelFont.FontItalic
    Settings.StepsViewFontName = LabelFont.FontName
    Settings.StepsViewFontSize = LabelFont.FontSize
    
    ' [Лог]
    Settings.RewriteLogFile = CheckRewriteLogFile.Value > 0
    Settings.LogFilePath = TextLogFilePath.Text
    
    ' [История файлов]
    Settings.FilesHistoryLimitPaths = CheckFilesHistoryLimitPaths.Value > 0
    MRUFileList.MaxFileCount = CInt(TextFilesHistoryCount.Text)

    ' [Обновление]
    Settings.AutoUpdateEnabled = CheckEnableAutoUpdate.Value > 0
    Settings.AutoUpdatePeriod = ComboAutoUpdatePeriod.ListIndex + 1

    ' Сохраняем изменения настроек в файле конфигурации
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
            StepsView.CellBackColor = &HC8D0D4
            
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

    Text1.Text = "" & Settings.StepsColWidth
    Text2.Text = "" & Settings.StepsRowHeight

    ComboImportFormat.ListIndex = 0
    ComboExportFormat.ListIndex = 0

    Select Case Settings.StepsVSelectorEnabled
        Case False: CheckVSelector.Value = 0
        Case True: CheckVSelector.Value = 1
    End Select
    
    Select Case Settings.StepsHSelectorEnabled
        Case False: CheckHSelector.Value = 0
        Case True: CheckHSelector.Value = 1
    End Select
    
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

    ' [Обновление]
    Select Case Settings.AutoUpdateEnabled

        Case False: CheckEnableAutoUpdate.Value = 0

        Case True: CheckEnableAutoUpdate.Value = 1

    End Select

    ComboAutoUpdatePeriod.ListIndex = Settings.AutoUpdatePeriod - 1
    ComboAutoUpdatePeriod.Enabled = CheckEnableAutoUpdate.Value > 0
    LabelAutoUpdatePeriod.Enabled = CheckEnableAutoUpdate.Value > 0

    RefreshStepsView

    StepsView.row = 1
    StepsView.col = 1

    ShowHorizontalSelector
    ShowVerticalSelector
    
    ' Перемещаем форму в центр экрана
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
