VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Univest Digital projects v3"
   ClientHeight    =   6630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10635
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   10635
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView lstProjects 
      Height          =   4215
      Left            =   120
      TabIndex        =   17
      Top             =   480
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   7435
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name (english A-Z and nums 0-9):"
         Object.Width           =   5080
      EndProperty
   End
   Begin VB.CommandButton cmdEditProject 
      Height          =   495
      Left            =   2280
      Picture         =   "frmMain.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Edit project name"
      Top             =   4920
      Width           =   495
   End
   Begin VB.CommandButton cmdRemoveProject 
      Height          =   495
      Left            =   1320
      Picture         =   "frmMain.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Remove project from list"
      Top             =   4920
      Width           =   495
   End
   Begin VB.CommandButton cmdAddProj 
      Height          =   495
      Left            =   240
      Picture         =   "frmMain.frx":0E0E
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Add new project"
      Top             =   4920
      Width           =   495
   End
   Begin VB.Frame Frame1 
      Caption         =   "Parameters for selected project:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5535
      Left            =   3240
      TabIndex        =   5
      Top             =   120
      Width           =   7095
      Begin VB.CheckBox chkFotoMe 
         Caption         =   "Project is FotoMe-compatible (small size - 74x105 mm - and different XML-file)"
         ForeColor       =   &H00808000&
         Height          =   375
         Left            =   240
         TabIndex        =   22
         Top             =   3120
         Width           =   3495
      End
      Begin VB.ComboBox cmbFont 
         Height          =   315
         ItemData        =   "frmMain.frx":1250
         Left            =   2520
         List            =   "frmMain.frx":1281
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   735
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   18
         Text            =   "frmMain.frx":12C1
         Top             =   4680
         Width           =   6615
      End
      Begin VB.CommandButton btnChangeBackImage 
         Caption         =   "Select &background JPG image"
         Height          =   615
         Left            =   240
         TabIndex        =   15
         Top             =   3840
         Width           =   3495
      End
      Begin VB.TextBox txtInput 
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "C:\"
         Top             =   765
         Width           =   3495
      End
      Begin VB.TextBox txtOutput 
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   8
         Text            =   "C:\"
         Top             =   1725
         Width           =   3495
      End
      Begin VB.CommandButton btnSelectInput 
         Caption         =   "Select &input folder"
         Height          =   375
         Left            =   1680
         TabIndex        =   7
         Top             =   360
         Width           =   2055
      End
      Begin VB.CommandButton btnSelectOutput 
         Caption         =   "Select &output folder"
         Height          =   375
         Left            =   1680
         TabIndex        =   6
         Top             =   1320
         Width           =   2055
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(default is 15)"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   21
         Top             =   2520
         Width           =   945
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Font size for ORDER_ID:"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   20
         Top             =   2280
         Width           =   1785
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   4095
         Left            =   3960
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2895
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Input folder path:"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   480
         Width           =   1200
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Output folder path:"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   1440
         Width           =   1320
      End
   End
   Begin VB.CommandButton btnStart 
      Caption         =   "Start &processing"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3570
      TabIndex        =   3
      Top             =   5880
      Width           =   1815
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save settings"
      Default         =   -1  'True
      Height          =   495
      Left            =   1890
      TabIndex        =   2
      Top             =   5880
      Width           =   1455
   End
   Begin VB.CommandButton cmdDiscard 
      Caption         =   "&Exit programm"
      Height          =   495
      Left            =   5610
      TabIndex        =   1
      Top             =   5880
      Width           =   1455
   End
   Begin VB.CommandButton cmdHide 
      Caption         =   "&Hide this window"
      Height          =   495
      Left            =   7290
      TabIndex        =   0
      Top             =   5880
      Width           =   1455
   End
   Begin VB.PictureBox pictWorking 
      Enabled         =   0   'False
      Height          =   495
      Left            =   960
      Picture         =   "frmMain.frx":134F
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   4
      Top             =   5760
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Timer tmrWatch 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   240
      Top             =   5760
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Projects list:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   12
      Top             =   120
      Width           =   1065
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Const WM_NCACTIVATE = &H86
Const SC_CLOSE = &HF060
Const MF_BYCOMMAND = &H0&
Const TXT_INPUT As String = "input.txt"
Const TXT_OUTPUT As String = "output.txt"
Const TXT_PROJECTS As String = "projects.txt"
Const TXT_FONTSIZE As String = "fonts.txt"

Private WithEvents FormSys As FrmSysTray
Attribute FormSys.VB_VarHelpID = -1

'Private gGflAx As GflAx.GflAx
Private gGflAx As Object

Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long

Private Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, _
            ByVal wFlags As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, _
            ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
            
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
            
Private Const LOCALE_SDECIMAL = &HE         '  decimal separator

Private Declare Function SetLocaleInfo Lib "kernel32" Alias "SetLocaleInfoA" _
  (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String) As Long

Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" _
  (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long

Private Declare Function GetUserDefaultLCID% Lib "kernel32" ()
Private Declare Function GetSystemDefaultLCID& Lib "kernel32" ()
Private Declare Function GetACP& Lib "kernel32" ()
            
Private Type aIMG
    in_1        As Variant
    in_2        As Variant
    in_3        As Variant
    in_4        As Variant
    in_5        As Variant
    in_6        As Variant
    in_7        As Variant
    in_8        As Variant
'Need some explain about next strange parameter
'Adobe programs is writing JPEG CMYK images in strange color encoding named YCCK :(
'If a program for viewing JPEG doesnt understand that sucks format then we see INVERTED image :(
'Fuck
    in_9_ycck   As Boolean
End Type


Private sInputFolder() As String
Private sOutputFolder() As String
Private sProjects() As String
Private bProjectIsFotoMeCompat() As Boolean
Private FSO As Object, FO As Object, FI As Object
Private xmlDoc As Object, currNode As Object
Private sINP() As String, sOUTP() As String
Private iFN As Integer, iFN1 As Integer, iFN2 As Integer, i As Integer
Private bInProcessing As Boolean
Private MAX_PROJECT_NUM As Integer

Private sPDFencoding() As String
Private sOutJPG As Variant

Private Const sCyrAPDF = "ÀÁÂÃÄÅ¨ÆÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖ×ØÙÜÚÛÝÞß²¯ª" & _
                         "àáâãäå¸æçèéêëìíîïðñòóôõö÷øùüúûýþÿ³¿º" & _
                         "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ" & _
                         "« 1234567890-+=_`~!@#$%^&*()[]{};:'»\|/?,.<>¹"
                         
Private Const sCyrBPDF = "\037•\036•\035•\034•\033•\032•\031•\030•\027•\026•\025•\024•\023•\022•\021•\020•\017•" & _
                         "\016•\r•\f•\013•\n•\t•\b•\007•\006•\005•\004•\003•\002•\001••" & _
                         "\201•\215•\217•\220•\235•\240•\255•\200•\202•\203•\204•\205•\206•\207•\210•\211•" & _
                         "\212•\213•\214•\216•\221•\222•\223•\224•\225•\226•\227•\230•\231•\232•\233•\234•\236•" & _
                         "\237•\241•\242•\243•\244•\245•\246•a•b•c•d•e•f•g•h•i•j•k•l•m•n•o•p•q•r•s•t•u•v•w•x•y•z•" & _
                         "A•B•C•D•E•F•G•H•I•J•K•L•M•N•O•P•Q•R•S•T•U•V•W•X•Y•Z•\253• •1•2•3•4•5•6•7•8•9•0•" & _
                         "-•+•=•_•`•~•!•@•#•$•%•^•&•*•\(•\)•[•]•{•}•;•:•\247•\273•\\•|•/•?•,•.•<•>•#"
                         
Private Const sFileds = "job_phb_format|job_copy|job_phb_pages|job_phb_binding|job_phb_cover_color|job_phb_cut|" & _
                        "name|tel|type_of_delivery|address_of_delivery_region|address_of_delivery_city|" & _
                        "address_of_delivery_index|address_of_delivery_street|address_of_delivery_house|" & _
                        "address_of_delivery_flat|address_of_delivery_domophone|type_of_payment|spm_price|" & _
                        "spm_t_price|spm_value|spm_t_value|order_id|date_of_receipt"

'Private Const sFiledsFotoMe = "orderID|orderDate|paymentType|item_albumName|customerData_name|customerData_telephone|" & _
                              "customerData_cellPhone|deliveryType|deliveryCustomerData_city|deliveryCustomerData_postalCode|" & _
                              "deliveryCustomerData_address|deliveryCustomerData_name|deliveryCustomerData_telephone|" & _
                              "deliveryCustomerData_cellPhone|store_name|store_telephone|store_city|subtotal|" & _
                              "shipCost|discount|orderData_total"
                              
Private Const sFiledsFotoMe = "orderID|orderDate|paymentType|item_albumName|customerData_name|customerData_telephone|" & _
                              "customerData_cellPhone|deliveryType|deliveryCustomerData_city|deliveryCustomerData_postalCode|" & _
                              "deliveryCustomerData_address|deliveryCustomerData_name|deliveryCustomerData_telephone|" & _
                              "deliveryCustomerData_cellPhone|store_name|store_telephone|store_city|" & _
                              "item_unitPrice|item_quantity|item_discountQuantity|item_subtotal|item_id_1|item_id_2"


Private Sub btnChangeBackImage_Click()
    Dim sPath As String
    If (MAX_PROJECT_NUM = 0) Or (Me.lstProjects.SelectedItem.Index = 0) Then
        MsgBox "No selected projects!", vbCritical, "Univest Digital"
        Exit Sub
    End If
    sPath = ShowOpenFileDialog("JPG images (*.jpg)|*.jpg", "jpg", , _
            OFN_EXPLORER Or OFN_FILEMUSTEXIST Or OFN_LONGNAMES, Me.hwnd)
    If Len(sPath) > 0 Then
        If sPath = App.Path & "\" & sProjects(Me.lstProjects.SelectedItem.Index) & ".jpg" Or _
            sPath = App.Path & "\" & sProjects(Me.lstProjects.SelectedItem.Index) & "_preview.jpg" _
            Then
            MsgBox "Selected file can't be a background image!!!" & vbCrLf & _
            "Please select another image!", vbCritical, "Univest Digital"
            Exit Sub
        End If
    
        sOutJPG = ParseJPG(sPath)
        If TypeName(sOutJPG) = "Boolean" Then
            MsgBox "Error loading JPG image!", vbCritical, "Univest Digital"
            Exit Sub
        End If
        
        FSO.DeleteFile App.Path & "\" & sProjects(Me.lstProjects.SelectedItem.Index) & ".jpg", True
        FSO.DeleteFile App.Path & "\" & sProjects(Me.lstProjects.SelectedItem.Index) & "_preview.jpg", True
        
        FSO.CopyFile sPath, App.Path & "\" & sProjects(Me.lstProjects.SelectedItem.Index) & ".jpg", True
        
        sPath = App.Path & "\" & sProjects(Me.lstProjects.SelectedItem.Index) & ".jpg"
                        
        gGflAx.LoadBitmap sPath
        gGflAx.SaveFormat = 1 'AX_JPEG
        gGflAx.SaveBitmap App.Path & "\" & sProjects(Me.lstProjects.SelectedItem.Index) & "_preview.jpg"
        Me.Image1.Picture = LoadPicture(App.Path & "\" & sProjects(Me.lstProjects.SelectedItem.Index) & "_preview.jpg")
    End If
    DoEvents
End Sub

Private Sub btnSelectInput_Click()
    Dim sPath As String
    sPath = BrowseForFolder(Me.hwnd)

    If Len(sPath) > 0 Then
        sInputFolder(Me.lstProjects.SelectedItem.Index) = sPath
        Me.txtInput.Text = sPath
        WriteHotFolders
    End If
End Sub

Private Sub btnSelectOutput_Click()
    Dim sPath As String
    sPath = BrowseForFolder(Me.hwnd)

    If Len(sPath) > 0 Then
        sOutputFolder(Me.lstProjects.SelectedItem.Index) = sPath
        Me.txtOutput.Text = sPath
        WriteHotFolders
    End If
End Sub

Private Sub btnStart_Click()
    If Me.btnStart.Caption = "Start &processing" Then
        If FodlersIsWrong Then
            Me.Show
            MsgBox "Please select properly hot folders!", vbExclamation + vbOKOnly, "Univest Digital warning"
            Exit Sub
        End If
        
        Call WriteProjects
        Call WriteHotFolders
    
        Me.btnStart.Caption = "Stop &processing"
        bInProcessing = True
        Me.Hide
        FormSys.TrayIcon = "DEFAULT"
        FormSys.Tooltip = "Watching folders"
        
        Me.btnSelectInput.Enabled = False
        Me.btnSelectOutput.Enabled = False
        Me.lstProjects.Enabled = False
        Me.btnChangeBackImage.Enabled = False
        Me.cmdAddProj.Enabled = False
        Me.cmdEditProject.Enabled = False
        Me.cmdRemoveProject.Enabled = False
        Me.cmdSave.Enabled = False
        Me.cmbFont.Enabled = False
        Me.chkFotoMe.Enabled = False
        Me.tmrWatch.Enabled = True
        DoEvents
    Else
        Me.tmrWatch.Enabled = False

        Me.btnStart.Caption = "Start &processing"
        
        Me.cmbFont.Enabled = Not bProjectIsFotoMeCompat(Me.lstProjects.SelectedItem.Index)
        Me.btnSelectInput.Enabled = True
        Me.btnSelectOutput.Enabled = True
        Me.lstProjects.Enabled = True
        Me.btnChangeBackImage.Enabled = True
        Me.cmdAddProj.Enabled = True
        Me.cmdEditProject.Enabled = True
        Me.cmdRemoveProject.Enabled = True
        Me.cmdSave.Enabled = True
        Me.chkFotoMe.Enabled = True

        bInProcessing = False
        FormSys.TrayIcon = Me
        FormSys.Tooltip = "Univest Digital projects"
        DoEvents
    End If
End Sub


Private Sub chkFotoMe_Click()
    Dim lRess As Long, bChange As Boolean
    bChange = False
    If Me.Visible And Me.chkFotoMe.Enabled And (bProjectIsFotoMeCompat(Me.lstProjects.SelectedItem.Index) <> -(Me.chkFotoMe.Value)) Then
        bChange = True
        lRess = MsgBox("If you want to leave current background image with new project format, press <YES>," & vbCrLf & _
            "if you want to reset it to default, press <NO>. To discard changes press <CANCEL>.", vbQuestion + vbYesNoCancel + vbDefaultButton2)
        If lRess = vbCancel Then
            Me.chkFotoMe.Value = 1 - Me.chkFotoMe.Value
            Exit Sub
        End If
    End If
    bProjectIsFotoMeCompat(Me.lstProjects.SelectedItem.Index) = -(Me.chkFotoMe.Value)
    Me.cmbFont.Enabled = Not bProjectIsFotoMeCompat(Me.lstProjects.SelectedItem.Index)
    If Me.cmbFont.Enabled Then
        Me.Text1.Text = "Background image MUST be saved in a JPG format from PHOTOSHOP" & vbCrLf & _
                        "(Gray, RGB or CMYK). Image will be STRETCHED to full page size" & vbCrLf & _
                        "(152x214 mm)"
    Else
        Me.Text1.Text = "Background image MUST be saved in a JPG format from PHOTOSHOP" & vbCrLf & _
                        "(Gray, RGB or CMYK). Image will be STRETCHED to full page size" & vbCrLf & _
                        "(74x105 mm)"
    End If
    If bChange Then
        If lRess = vbNo Then
            If bProjectIsFotoMeCompat(Me.lstProjects.SelectedItem.Index) Then
                FSO.CopyFile App.Path & "\blank_me.jpg", App.Path & "\" & sProjects(Me.lstProjects.SelectedItem.Index) & ".jpg", True
            Else
                FSO.CopyFile App.Path & "\blank.jpg", App.Path & "\" & sProjects(Me.lstProjects.SelectedItem.Index) & ".jpg", True
            End If
            gGflAx.LoadBitmap App.Path & "\" & sProjects(Me.lstProjects.SelectedItem.Index) & ".jpg"
            gGflAx.SaveFormat = 1 'AX_JPEG
            gGflAx.SaveBitmap App.Path & "\" & sProjects(Me.lstProjects.SelectedItem.Index) & "_preview.jpg"
            Me.Image1.Picture = LoadPicture(App.Path & "\" & sProjects(Me.lstProjects.SelectedItem.Index) & "_preview.jpg")
        End If
    End If
    WriteProjects
End Sub

Private Sub cmbFont_Click()
    On Error Resume Next
    iFN = FreeFile()
    Open App.Path & "\" & TXT_FONTSIZE For Output As #iFN
        Print #iFN, Me.cmbFont.Text
    Close #iFN
    Err.Clear
End Sub

Private Sub cmdAddProj_Click()
    Dim sStr As String, jj As Integer
    If Me.lstProjects.ListItems.Count = 99 Then
        MsgBox "You have 99 active projects. This is not enough???" & _
                "Adding more is impossible. Try to reorganize your business ;)", _
                vbQuestion, "Univest Digital projects"
        Exit Sub
    End If
    sStr = InputBox("Please enter name for a new project:" & vbCrLf & _
            "(only english letters (A-Z) and numbers (0 to 9) allowed!)", _
            "Create new project", vbNullString)
    sStr = CleanStringFromWrongSymbols(sStr)
    If Len(sStr) = 0 Then MsgBox "Invalid name!", vbCritical, "Create new project": Exit Sub
    sStr = UCase$(sStr)
    For jj = 1 To MAX_PROJECT_NUM
        If sStr = sProjects(jj) Then MsgBox "Project with this name already exists!", vbCritical, "Create new project": Exit Sub
    Next jj
    MAX_PROJECT_NUM = MAX_PROJECT_NUM + 1
    ReDim Preserve sProjects(1 To MAX_PROJECT_NUM)
    ReDim Preserve bProjectIsFotoMeCompat(1 To MAX_PROJECT_NUM)
    ReDim Preserve sInputFolder(1 To MAX_PROJECT_NUM)
    ReDim Preserve sOutputFolder(1 To MAX_PROJECT_NUM)
    ReDim Preserve sOUTP(1 To MAX_PROJECT_NUM)
    ReDim Preserve sINP(1 To MAX_PROJECT_NUM)
    sProjects(MAX_PROJECT_NUM) = sStr
    bProjectIsFotoMeCompat(MAX_PROJECT_NUM) = False
    sInputFolder(MAX_PROJECT_NUM) = "C:\"
    sOutputFolder(MAX_PROJECT_NUM) = "C:\"
    sINP(i) = App.Path & "\" & CStr(MAX_PROJECT_NUM) & TXT_INPUT
    sOUTP(i) = App.Path & "\" & CStr(MAX_PROJECT_NUM) & TXT_OUTPUT
    Me.lstProjects.ListItems.Add MAX_PROJECT_NUM, sStr, sStr
    FSO.CopyFile App.Path & "\blank.jpg", App.Path & "\" & sProjects(MAX_PROJECT_NUM) & ".jpg", True
        gGflAx.LoadBitmap App.Path & "\" & sProjects(MAX_PROJECT_NUM) & ".jpg"
        gGflAx.SaveFormat = 1 'AX_JPEG
        gGflAx.SaveBitmap App.Path & "\" & sProjects(MAX_PROJECT_NUM) & "_preview.jpg"
        Me.Image1.Picture = LoadPicture(App.Path & "\" & sProjects(MAX_PROJECT_NUM) & "_preview.jpg")
    Me.lstProjects.ListItems(MAX_PROJECT_NUM).Selected = True
    Me.txtInput.Text = sInputFolder(MAX_PROJECT_NUM)
    Me.txtOutput.Text = sOutputFolder(MAX_PROJECT_NUM)
    Me.chkFotoMe.Value = 0
    Me.cmbFont.Enabled = True
    WriteProjects
    WriteHotFolders
End Sub

Private Sub cmdDiscard_Click()
    Call FormSys.mExit_Click
End Sub

Private Sub cmdEditProject_Click()
    Me.lstProjects.SetFocus
    Me.lstProjects.StartLabelEdit
End Sub

Private Sub cmdHide_Click()
    Me.Hide
End Sub

Private Sub cmdRemoveProject_Click()
    If Me.lstProjects.ListItems.Count = 1 Then
        MsgBox "You have only one project! It couldn't be deleted!", vbCritical, "Univest Digital"
        Exit Sub
    End If
    If MsgBox("Do you really want to delete selected project???", _
            vbQuestion + vbYesNo + vbDefaultButton2, "Univest Digital") = vbYes Then
        If MsgBox("Last question: ARE YOU SHURE?", vbCritical + vbYesNo + vbDefaultButton2, _
                "Confirm deleting project") = vbYes Then
            Call RemoveProject(Me.lstProjects.SelectedItem.Index)
        End If
    End If
End Sub

Private Sub cmdSave_Click()
    If FodlersIsWrong Then
        MsgBox "Please select properly hot folders!", vbExclamation + vbOKOnly, "Univest Digital projects warning"
        Exit Sub
    End If
    Call WriteProjects
    Call WriteHotFolders
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Me.Hide
End Sub

Private Sub Form_Load()
    Dim arrData() As Byte
    Dim oTest As Object
    
    If App.PrevInstance Then End
    
    'If (GetUserDefaultLCID <> 1049) Or (GetSystemDefaultLCID <> 1049) Or (GetACP <> 1251) Then
    If (GetSystemDefaultLCID <> 1049) Or (GetACP <> 1251) Then
        MsgBox "This program is designed for usage with FULL CYRILLIC system support." & vbCrLf & _
        "Your system locale or user locale or default system code page is NOT cyrillic." & vbCrLf & _
        "Go to <Control Panel>-<Regional Settings> and set <Regional options>-<Standards and formats> to RUSSIAN" & vbCrLf & _
        "and <Advanced>-<Language for non-Unicode programs> also to RUSSIAN. Then REBOOT and try again.", _
                vbCritical, "Univest Digital projects"
        End
    End If
    
    Me.Caption = "Univest Digital projects v" & CStr(App.Major) & "." & CStr(App.Minor) & "." & CStr(App.Revision) ' & " beta"
    
    DisableCloseButton Me.hwnd
    
    On Error Resume Next
    Set gGflAx = CreateObject("GflAx.GflAx")
    If Err.Number <> 0 Then
        Err.Clear
        If Not FSO.FileExists(App.Path & "\GflAx.dll") Then
            arrData = LoadResData(103, "CUSTOM")
            Open App.Path & "\GflAx.dll" For Binary Access Write As #1
                Put #1, , arrData
            Close #1
        End If
        Err.Clear
        Call Shell(Environ$("SYSTEMROOT") & "\system32\regsvr32.exe /s " & Chr$(34) & App.Path & "\GflAx.dll" & Chr$(34), vbHide)
        If Err.Number <> 0 Then
            MsgBox "Error accessing windows system folder!" & vbCrLf & _
            "Please ensure you have administrator access!", _
                vbCritical, "Univest Digital projects"
            End
        End If
        
        Err.Clear
        Set gGflAx = CreateObject("GflAx.GflAx")
        If Err.Number <> 0 Then
            MsgBox "Error creating and/or accessing GflAx object!" & _
            "You have to reinstall programm and use it under administrator access!", _
                vbCritical, "Univest Digital projects"
            End
        End If
    End If
    gGflAx.EnableLZW = True
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    If Err.Number <> 0 Then
        MsgBox "Error accessing windows scripting runtime! Looks like you have Windows 98 :) , not XP." & vbCrLf & _
                "You have to google 'Microsoft XML Core Services' (MSXML) and install it.", _
                vbCritical, "Univest Digital projects"
        End
    End If
        
    Set oTest = CreateObject("MSXML.DOMDocument")
    If Err.Number <> 0 Then
        MsgBox "Error accessing Microsoft XML runtime! Looks like you have Windows 98 :) , not XP." & vbCrLf & _
                "You have to google 'Microsoft XML' and install it.", _
                vbCritical, "Univest Digital projects"
        End
    End If
    
    Err.Clear
    
    Set xmlDoc = CreateObject("MSXML.DOMDocument")
    
    If Not FSO.FileExists(App.Path & "\blank.jpg") Then
        arrData = LoadResData(101, "CUSTOM")
        Open App.Path & "\blank.jpg" For Binary Access Write As #1
            Put #1, , arrData
        Close #1
    End If
    
    If Not FSO.FileExists(App.Path & "\blank_me.jpg") Then
        arrData = LoadResData(104, "CUSTOM")
        Open App.Path & "\blank_me.jpg" For Binary Access Write As #1
            Put #1, , arrData
        Close #1
    End If
    
    If Not FSO.FileExists(App.Path & "\shablon.pdf") Then
        arrData = LoadResData(102, "CUSTOM")
        Open App.Path & "\shablon.pdf" For Binary Access Write As #1
            Put #1, , arrData
        Close #1
    End If
    
    If Not FSO.FileExists(App.Path & "\shablon_me.pdf") Then
        arrData = LoadResData(105, "CUSTOM")
        Open App.Path & "\shablon_me.pdf" For Binary Access Write As #1
            Put #1, , arrData
        Close #1
    End If
    
    If Err.Number <> 0 Then
        MsgBox "Error while writing data to program folder! Looks like you have readonly device :(" & vbCrLf & _
                "You have to install program to normal harddisk ant try again!", _
                vbCritical, "Univest Digital projects"
        End
    End If
    Err.Clear
    On Error GoTo 0
    
    Set FormSys = New FrmSysTray
    Load FormSys
    Set FormSys.FSys = Me
    FormSys.TrayIcon = Me
    FormSys.Tooltip = "Univest Digital projects"
    
    If Not ReadProjects Then
        MsgBox "Error accessing disk! Device is read-only?", _
               vbCritical + vbOKOnly, "Univest Digital projects I/O error"
        End
    End If
    
    Call Init
    
    If Not FodlersIsWrong Then
        Call btnStart_Click
    Else
        Me.Show
    End If
    
End Sub

Private Sub DisableCloseButton(ByVal FormHWND As Long)
    Dim hMenu As Long, Success As Long
    hMenu = GetSystemMenu(FormHWND, 0)
    Success = DeleteMenu(hMenu, SC_CLOSE, MF_BYCOMMAND)
    SendMessage FormHWND, WM_NCACTIVATE, 0&, 0&
    SendMessage FormHWND, WM_NCACTIVATE, 1&, 0
End Sub

Private Function GetZakazXMLData(ByVal sFile As String, iProjectNumber As Integer) As Variant
    'Dim xmlDoc As Object, currNode As Object
    Dim bPDFResult As Variant, ii As Long, meCount As Long, bRes As Boolean, sPDFs() As String, vPDFs As Variant, sTmp As String
    
    If Not FSO.FileExists(sFile) Then GetZakazXMLData = False: Exit Function
    
    On Error Resume Next
    'Set xmlDoc = Nothing
    Set currNode = Nothing
    'Set xmlDoc = CreateObject("MSXML.DOMDocument")
    
    If xmlDoc Is Nothing Then Set xmlDoc = CreateObject("MSXML.DOMDocument")
    
    bRes = True
    If bProjectIsFotoMeCompat(iProjectNumber) Then
        vPDFs = SimplifyFotoMeXML(sFile)
        If TypeName(vPDFs) = "Boolean" Then
            GetZakazXMLData = False
            FSO.DeleteFile sFile, True
            Exit Function
        End If
        meCount = UBound(vPDFs, 1)
        ReDim sPDFs(1 To meCount)
        For ii = 1 To meCount
            sTmp = Replace(sFile, ".xml", "_" & CStr(vPDFs(ii, 1)) & ".xml")
            If xmlDoc Is Nothing Then Set xmlDoc = CreateObject("MSXML.DOMDocument")
            xmlDoc.async = False
            xmlDoc.Load sTmp
            If (xmlDoc.parseError.errorCode <> 0) Then
                'Error parsing XML - maybe not XML? :( We shall to kill this bad file :)
                bRes = False
                FSO.DeleteFile sTmp, True
                'Exit Function
            Else
                Set currNode = xmlDoc.selectSingleNode("//Order")
                If currNode Is Nothing Then
                    Set currNode = xmlDoc.selectSingleNode("//order")
                    If currNode Is Nothing Then
                        Err.Clear
                        bRes = False
                        'Exit Function
                    End If
                End If
                sOutJPG = ParseJPG(App.Path & "\" & sProjects(iProjectNumber) & ".jpg")
                sPDFs(ii) = ParsePDF(App.Path & "\shablon_me.pdf", sOutJPG, True)
                FSO.DeleteFile sTmp, True
            End If
        Next ii
    Else
        xmlDoc.async = False
        xmlDoc.Load sFile
        If (xmlDoc.parseError.errorCode <> 0) Then
            'Error parsing XML - maybe not XML? :( We shall to kill this bad file :)
            GetZakazXMLData = False
            FSO.DeleteFile sFile, True
            Exit Function
        Else
            Set currNode = xmlDoc.selectSingleNode("//Order")
            If currNode Is Nothing Then
                Set currNode = xmlDoc.selectSingleNode("//order")
                If currNode Is Nothing Then
                    Err.Clear
                    GetZakazXMLData = False
                    Exit Function
                End If
            End If
            sOutJPG = ParseJPG(App.Path & "\" & sProjects(iProjectNumber) & ".jpg")
            bPDFResult = ParsePDF(App.Path & "\shablon.pdf", sOutJPG, False)
        End If
    End If
    
    Set currNode = Nothing
    'Set xmlDoc = Nothing
    If (Err.Number <> 0) Or (TypeName(bPDFResult) = "Boolean" Or bRes = False) Then
        GetZakazXMLData = False
    Else
        If bProjectIsFotoMeCompat(iProjectNumber) Then
            GetZakazXMLData = sPDFs
        Else
            GetZakazXMLData = bPDFResult
        End If
    End If
    Err.Clear
    On Error GoTo 0
End Function

Private Function SimplifyFotoMeXML(ByVal sFile As String) As Variant
    Dim FT1 As Object, FT2 As Object, FI1 As Object, i As Long
    Dim S1 As String, vArr() As String, lArr() As Long, S2 As String

    On Error Resume Next
    If Not FSO.FileExists(sFile) Then SimplifyFotoMeXML = False: Exit Function
    
    'MyMoveFileOrFolder Me.hwnd, sFile, sFile & "_origin.xml"
    
    Set FI1 = FSO.GetFile(sFile) ' & "_origin.xml")
    
    'Set FT2 = FSO.CreateTextFile(sFile, True, False)
    
    Set FT1 = FI1.OpenAsTextStream(1) 'ForReading)
    S1 = FT1.ReadAll
    
    vArr = Split(S1, "<item id=" & Chr$(34), , vbTextCompare)
    ReDim lArr(1 To UBound(vArr), 1 To 2)
    
    For i = 1 To UBound(lArr)
        lArr(i, 1) = Val(vArr(i))
        S1 = Mid$(vArr(i), InStr(1, UCase$(vArr(i)), "ITEMID=", vbTextCompare) + 8)
        lArr(i, 2) = Val(S1)
    
        Set FT1 = FI1.OpenAsTextStream(1) 'ForReading)
        
        Set FT2 = FSO.CreateTextFile(Replace(sFile, ".xml", "_" & CStr(lArr(i, 1)) & ".xml"), True, False)
    
    'Set FT2 = FI2.OpenAsTextStream(2) 'ForWriting)
        Do
            S1 = FT1.ReadLine
            If Trim$(S1) = "<order>" Then S1 = "<Order>"
            If Trim$(S1) = "</order>" Then S1 = "</Order>"
            If Trim$(S1) = "<customerData>" Then
                Do While Trim$(S1) <> "</customerData>"
                    S1 = FT1.ReadLine
                    If Trim$(S1) Like "<name>*</name>" Then
                        S1 = Replace(S1, "<name>", "<customerData_name>", 1, 1, vbTextCompare)
                        S1 = Replace(S1, "</name>", "</customerData_name>", 1, 1, vbTextCompare)
                        FT2.WriteLine S1
                    End If
                    If Trim$(S1) Like "<telephone>*</telephone>" Then
                        S1 = Replace(S1, "<telephone>", "<customerData_telephone>", 1, 1, vbTextCompare)
                        S1 = Replace(S1, "</telephone>", "</customerData_telephone>", 1, 1, vbTextCompare)
                        FT2.WriteLine S1
                    End If
                    If Trim$(S1) Like "<cellPhone>*</cellPhone>" Then
                        S1 = Replace(S1, "<cellPhone>", "<customerData_cellPhone>", 1, 1, vbTextCompare)
                        S1 = Replace(S1, "</cellPhone>", "</customerData_cellPhone>", 1, 1, vbTextCompare)
                        FT2.WriteLine S1
                    End If
                Loop
                S1 = vbNullString
            End If
            If Trim$(S1) = "<deliveryCustomerData>" Then
                Do While Trim$(S1) <> "</deliveryCustomerData>"
                    S1 = FT1.ReadLine
                    If Trim$(S1) Like "<name>*</name>" Then
                        S1 = Replace(S1, "<name>", "<deliveryCustomerData_name>", 1, 1, vbTextCompare)
                        S1 = Replace(S1, "</name>", "</deliveryCustomerData_name>", 1, 1, vbTextCompare)
                        FT2.WriteLine S1
                    End If
                    If Trim$(S1) Like "<address>*</address>" Then
                        S1 = Replace(S1, "<address>", "<deliveryCustomerData_address>", 1, 1, vbTextCompare)
                        S1 = Replace(S1, "</address>", "</deliveryCustomerData_address>", 1, 1, vbTextCompare)
                        FT2.WriteLine S1
                    End If
                    If Trim$(S1) Like "<postalCode>*</postalCode>" Then
                        S1 = Replace(S1, "<postalCode>", "<deliveryCustomerData_postalCode>", 1, 1, vbTextCompare)
                        S1 = Replace(S1, "</postalCode>", "</deliveryCustomerData_postalCode>", 1, 1, vbTextCompare)
                        FT2.WriteLine S1
                    End If
                    If Trim$(S1) Like "<city>*</city>" Then
                        S1 = Replace(S1, "<city>", "<deliveryCustomerData_city>", 1, 1, vbTextCompare)
                        S1 = Replace(S1, "</city>", "</deliveryCustomerData_city>", 1, 1, vbTextCompare)
                        FT2.WriteLine S1
                    End If
                    If Trim$(S1) Like "<telephone>*</telephone>" Then
                        S1 = Replace(S1, "<telephone>", "<deliveryCustomerData_telephone>", 1, 1, vbTextCompare)
                        S1 = Replace(S1, "</telephone>", "</deliveryCustomerData_telephone>", 1, 1, vbTextCompare)
                        FT2.WriteLine S1
                    End If
                    If Trim$(S1) Like "<cellPhone>*</cellPhone>" Then
                        S1 = Replace(S1, "<cellPhone>", "<deliveryCustomerData_cellPhone>", 1, 1, vbTextCompare)
                        S1 = Replace(S1, "</cellPhone>", "</deliveryCustomerData_cellPhone>", 1, 1, vbTextCompare)
                        FT2.WriteLine S1
                    End If
                Loop
                S1 = vbNullString
            End If
            If Trim$(S1) = "<store>" Then
                Do While Trim$(S1) <> "</store>"
                    S1 = FT1.ReadLine
                    If Trim$(S1) Like "<name>*</name>" Then
                        S1 = Replace(S1, "<name>", "<store_name>", 1, 1, vbTextCompare)
                        S1 = Replace(S1, "</name>", "</store_name>", 1, 1, vbTextCompare)
                        FT2.WriteLine S1
                    End If
                    If Trim$(S1) Like "<city>*</city>" Then
                        S1 = Replace(S1, "<city>", "<store_city>", 1, 1, vbTextCompare)
                        S1 = Replace(S1, "</city>", "</store_city>", 1, 1, vbTextCompare)
                        FT2.WriteLine S1
                    End If
                    If Trim$(S1) Like "<telephone>*</telephone>" Then
                        S1 = Replace(S1, "<telephone>", "<store_telephone>", 1, 1, vbTextCompare)
                        S1 = Replace(S1, "</telephone>", "</store_telephone>", 1, 1, vbTextCompare)
                        FT2.WriteLine S1
                    End If
                Loop
                S1 = vbNullString
            End If
            If Trim$(S1) Like "<item id=*" Then
                S1 = Mid$(S1, InStr(1, UCase$(S1), "<ITEM ID=", vbTextCompare) + 10)
                If Val(S1) <> i Then
                    Do While Trim$(S1) <> "</item>"
                        S1 = FT1.ReadLine
                    Loop
                    S1 = vbNullString
                Else
                    Do While Trim$(S1) <> "</item>"
                        S1 = FT1.ReadLine
                        If Trim$(S1) Like "<albumName>*</albumName>" Then
                            S1 = Replace(S1, "<albumName>", "<item_albumName>", 1, 1, vbTextCompare)
                            S1 = Replace(S1, "</albumName>", "</item_albumName>", 1, 1, vbTextCompare)
                            FT2.WriteLine S1
                        End If
                        If Trim$(S1) Like "<Filename>*</Filename>" Then
                            S1 = Replace(S1, "<Filename>", "<item_Filename>", 1, 1, vbTextCompare)
                            S1 = Replace(S1, "</Filename>", "</item_Filename>", 1, 1, vbTextCompare)
                            FT2.WriteLine S1
                        End If
                        If Trim$(S1) Like "<unitPrice>*</unitPrice>" Then
                            S1 = Replace(S1, "<unitPrice>", "<item_unitPrice>", 1, 1, vbTextCompare)
                            S1 = Replace(S1, "</unitPrice>", "</item_unitPrice>", 1, 1, vbTextCompare)
                            FT2.WriteLine S1
                        End If
                        If Trim$(S1) Like "<quantity>*</quantity>" Then
                            S1 = Replace(S1, "<quantity>", "<item_quantity>", 1, 1, vbTextCompare)
                            S1 = Replace(S1, "</quantity>", "</item_quantity>", 1, 1, vbTextCompare)
                            FT2.WriteLine S1
                        End If
                        If Trim$(S1) Like "<discountQuantity>*</discountQuantity>" Then
                            S1 = Replace(S1, "<discountQuantity>", "<item_discountQuantity>", 1, 1, vbTextCompare)
                            S1 = Replace(S1, "</discountQuantity>", "</item_discountQuantity>", 1, 1, vbTextCompare)
                            FT2.WriteLine S1
                        End If
                        If Trim$(S1) Like "<subtotal>*</subtotal>" Then
                            S1 = Replace(S1, "<subtotal>", "<item_subtotal>", 1, 1, vbTextCompare)
                            S1 = Replace(S1, "</subtotal>", "</item_subtotal>", 1, 1, vbTextCompare)
                            FT2.WriteLine S1
                        End If
                    Loop
                    S1 = vbNullString
                    FT2.WriteLine "<item_id_1>" & CStr(i) & "</item_id_1>"
                    FT2.WriteLine "<item_id_2>" & CStr(lArr(i, 2)) & "</item_id_2>"
                End If
            End If
                
            'If Trim$(S1) Like "*deliveryCustomerData>*" Then S1 = vbNullString
            If Trim$(S1) Like "*orderData>*" Then S1 = vbNullString
            'If Trim$(S1) Like "*<item id=*" Then S1 = vbNullString
            'If Trim$(S1) Like "*</item>*" Then S1 = vbNullString
            If Trim$(S1) Like "<total>*</total>" Then
                S1 = Replace(S1, "<total>", "<orderData_total>", 1, 1, vbTextCompare)
                S1 = Replace(S1, "</total>", "</orderData_total>", 1, 1, vbTextCompare)
            End If
            If Trim$(S1) Like "<subtotal>*</subtotal>" Then
                S1 = Replace(S1, "<subtotal>", "<orderData_subtotal>", 1, 1, vbTextCompare)
                S1 = Replace(S1, "</subtotal>", "</orderData_subtotal>", 1, 1, vbTextCompare)
            End If
            If Trim$(S1) Like "<discount>*</discount>" Then
                S1 = Replace(S1, "<discount>", "<orderData_discount>", 1, 1, vbTextCompare)
                S1 = Replace(S1, "</discount>", "</orderData_discount>", 1, 1, vbTextCompare)
            End If
           
           
            If Len(S1) > 0 Then FT2.WriteLine S1
        Loop Until FT1.AtEndOfStream
        FT2.Close
    
    Next i
    
    FT1.Close
    'FT2.Close
    
    'FI1.Delete
    Set FI1 = Nothing
    Set FT1 = Nothing
    Set FT2 = Nothing
    If Err.Number = 0 Then SimplifyFotoMeXML = lArr Else SimplifyFotoMeXML = False
    Err.Clear
    On Error GoTo 0
End Function

Private Sub Init()
  
    sPDFencoding = Split(sCyrBPDF, "•")

    For i = 1 To MAX_PROJECT_NUM
        sINP(i) = App.Path & "\" & CStr(i) & TXT_INPUT
        sOUTP(i) = App.Path & "\" & CStr(i) & TXT_OUTPUT
    Next i

    If ReadHotFolders = False Then
        MsgBox "Error accessing disk! Device is read-only?", _
               vbCritical + vbOKOnly, "Univest Digital projects I/O error"
        End
     End If
     
    For i = 1 To MAX_PROJECT_NUM
        If Not FSO.FolderExists(sInputFolder(i)) Then
            sInputFolder(i) = "C:\"
        End If
    
        If Not FSO.FolderExists(sOutputFolder(i)) Then
            sOutputFolder(i) = "C:\"
        End If
        
        If Not FSO.FileExists(App.Path & "\" & sProjects(i) & ".jpg") Then
            If bProjectIsFotoMeCompat(i) Then
                FSO.CopyFile App.Path & "\blank_me.jpg", App.Path & "\" & sProjects(i) & ".jpg", True
            Else
                FSO.CopyFile App.Path & "\blank.jpg", App.Path & "\" & sProjects(i) & ".jpg", True
            End If
        End If
        
        Me.lstProjects.ListItems.Add i, sProjects(i), sProjects(i)
    Next i
    
    Me.lstProjects.ListItems(1).Selected = True
    Me.chkFotoMe.Value = -(bProjectIsFotoMeCompat(1))
    Me.cmbFont.Enabled = Not bProjectIsFotoMeCompat(1)
    gGflAx.LoadBitmap App.Path & "\" & sProjects(1) & ".jpg"
    gGflAx.SaveFormat = 1 'AX_JPEG
    gGflAx.SaveBitmap App.Path & "\" & sProjects(1) & "_preview.jpg"
    Call lstProjects_ItemClick(Me.lstProjects.ListItems(1))
    'Me.lstProjects.SetFocus
    
    bInProcessing = False
    
    
   
End Sub

Public Function ReadHotFolders() As Boolean
Dim sFontSize As String

    On Error Resume Next
    
    iFN = FreeFile()
    
    If FSO.FileExists(App.Path & "\" & TXT_FONTSIZE) Then
        Open App.Path & "\" & TXT_FONTSIZE For Input As #iFN
            Line Input #iFN, sFontSize
        Close #iFN
        Err.Clear
        If Val(sFontSize) > 9 And Val(sFontSize) < 25 Then
            Me.cmbFont.Text = Format$(Val(sFontSize), "00")
        Else
            Me.cmbFont.Text = "15"
        End If
    Else
        Me.cmbFont.Text = "15"
    End If
    Err.Clear


 For i = 1 To MAX_PROJECT_NUM
    sInputFolder(i) = vbNullString
    sOutputFolder(i) = vbNullString

    
    iFN = FreeFile()
    Open sINP(i) For Input As iFN
    Close iFN
        If Err.Number <> 0 Then ' no file!!!
            'Close iFN
            Err.Clear
            Open sINP(i) For Output As iFN
                If Err.Number <> 0 Then 'error creating file!!! Device redonly?
                    Err.Clear
                    On Error GoTo 0
                    ReadHotFolders = False
                    Exit Function
                End If
                Print #iFN, "C:\"
                sInputFolder(i) = "C:\"
            Close iFN
        End If
    Open sINP(i) For Input As iFN
        Line Input #iFN, sInputFolder(i)
    Close iFN
    Err.Clear
    
'    iFN = FreeFile()
    
    Open sOUTP(i) For Input As iFN
    Close iFN
        If Err.Number <> 0 Then ' no file!!!
            Close iFN
            Err.Clear
            Open sOUTP(i) For Output As iFN
                If Err.Number <> 0 Then 'error creating file!!! Device redonly?
                    Err.Clear
                    On Error GoTo 0
                    ReadHotFolders = False
                    Exit Function
                End If
                Print #iFN, "C:\"
                sOutputFolder(i) = "C:\"
            Close iFN
        End If
    Open sOUTP(i) For Input As iFN
        Line Input #iFN, sOutputFolder(i)
    Close iFN
    Err.Clear
 Next i
 

    On Error GoTo 0


 ReadHotFolders = True
End Function

Public Function WriteHotFolders() As Boolean
    Dim sTmpFile As String
    On Error Resume Next
    For i = 1 To 99 ' some cleaning :)
        sTmpFile = App.Path & "\" & CStr(i) & TXT_INPUT
        If FSO.FileExists(sTmpFile) Then FSO.DeleteFile sTmpFile
        sTmpFile = App.Path & "\" & CStr(i) & TXT_OUTPUT
        If FSO.FileExists(sTmpFile) Then FSO.DeleteFile sTmpFile
    Next i
    Err.Clear
    
    
    For i = 1 To MAX_PROJECT_NUM

     iFN = FreeFile()
     Open sINP(i) For Output As iFN

     If Err.Number <> 0 Then 'error creating file!!! Device redonly?

         Err.Clear
            On Error GoTo 0
         WriteHotFolders = False
            Exit Function

        End If

     Print #iFN, sInputFolder(i)
     Close iFN
    
     iFN = FreeFile()
     Open sOUTP(i) For Output As iFN

     If Err.Number <> 0 Then 'error creating file!!! Device redonly?

         Err.Clear
            On Error GoTo 0
         WriteHotFolders = False
            Exit Function

        End If

     Print #iFN, sOutputFolder(i)
     Close iFN

    Next i
    
  
        On Error GoTo 0
     WriteHotFolders = True

End Function


Private Sub lstProjects_AfterLabelEdit(Cancel As Integer, NewString As String)
    Dim sTemp As String
    sTemp = CleanStringFromWrongSymbols(UCase$(NewString))
    If (Len(sTemp) > 0) And (sProjects(Me.lstProjects.SelectedItem.Index) <> sTemp) Then
        MyMoveFileOrFolder Me.hwnd, _
                App.Path & "\" & sProjects(Me.lstProjects.SelectedItem.Index) & ".jpg", _
                App.Path & "\" & sTemp & ".jpg"
        MyMoveFileOrFolder Me.hwnd, _
                App.Path & "\" & sProjects(Me.lstProjects.SelectedItem.Index) & "_preview.jpg", _
                App.Path & "\" & sTemp & "_preview.jpg"
        sProjects(Me.lstProjects.SelectedItem.Index) = sTemp
        Me.lstProjects.SelectedItem.Key = sTemp
'        gGflAx.LoadBitmap App.Path & "\" & sProjects(MAX_PROJECT_NUM) & ".jpg"
'        gGflAx.SaveFormat = 1 'AX_JPEG
'        gGflAx.SaveBitmap App.Path & "\" & sProjects(MAX_PROJECT_NUM) & "_preview.jpg"
        Me.Image1.Picture = LoadPicture(App.Path & "\" & sProjects(MAX_PROJECT_NUM) & "_preview.jpg")
        Me.chkFotoMe.Value = -(bProjectIsFotoMeCompat(Me.lstProjects.SelectedItem.Index))
        Me.cmbFont.Enabled = Not bProjectIsFotoMeCompat(bProjectIsFotoMeCompat(Me.lstProjects.SelectedItem.Index))
        WriteProjects
    Else
        Cancel = -1
    End If
End Sub

Private Sub lstProjects_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Me.txtInput.Text = sInputFolder(Item.Index)
    Me.txtOutput.Text = sOutputFolder(Item.Index)
'    gGflAx.LoadBitmap App.Path & "\" & sProjects(Item.Index) & ".jpg"
'    gGflAx.SaveFormat = 1 'AX_JPEG
'    gGflAx.SaveBitmap App.Path & "\" & sProjects(Item.Index) & "_preview.jpg"
    Me.Image1.Picture = LoadPicture(App.Path & "\" & sProjects(Item.Index) & "_preview.jpg")
    Me.chkFotoMe.Value = -(bProjectIsFotoMeCompat(Item.Index))
    Me.cmbFont.Enabled = Not bProjectIsFotoMeCompat(Item.Index)
End Sub

Private Sub tmrWatch_Timer()
    Dim sWorkPDF As Variant, nNum As Long, varArr As Variant
    On Error Resume Next
    For i = 1 To MAX_PROJECT_NUM
        If Not FSO.FolderExists(sInputFolder(i)) Then
            MsgBox "Input folder " & sInputFolder(i) & " does not exists!!!", vbCritical + vbOKOnly, "Univest Digital projects"
            Me.Show
            Err.Clear
            On Error GoTo 0
            Call btnStart_Click
        End If
        If Not FSO.FolderExists(sOutputFolder(i)) Then
            MsgBox "Output folder " & sOutputFolder(i) & " does not exists!!!", vbCritical + vbOKOnly, "Univest Digital projects"
            Me.Show
            Err.Clear
            On Error GoTo 0
            Call btnStart_Click
        End If
        Set FO = FSO.GetFolder(sInputFolder(i))
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo 0
            Exit Sub
        End If
        For Each FI In FO.Files
            If FI.Name Like "*.xml" Then
                If bInProcessing = False Then Me.tmrWatch.Enabled = False: Exit Sub
                sWorkPDF = GetZakazXMLData(FI.Path, i)
                If TypeName(sWorkPDF) <> "Boolean" Then
                    If TypeName(sWorkPDF) = "String" Then
                        FSO.CopyFile sWorkPDF, sOutputFolder(i) & "\" & FSO.GetFile(sWorkPDF).Name, True
                        FSO.DeleteFile sWorkPDF, True
                    End If
                    If TypeName(sWorkPDF) = "String()" Then
                        For nNum = LBound(sWorkPDF) To UBound(sWorkPDF)
                            FSO.CopyFile sWorkPDF(nNum), sOutputFolder(i) & "\" & FSO.GetFile(sWorkPDF(nNum)).Name, True
                            FSO.DeleteFile sWorkPDF(nNum), True
                        Next nNum
                    End If
                End If
                FI.Name = FI.Name & ".tmp"
            End If
            If Err.Number <> 0 Then Err.Clear: Exit Sub
            Call Sleep(250)
            DoEvents
        Next
        FSO.DeleteFile sInputFolder(i) & "\*.tmp", True
        Err.Clear
    Next i
    Err.Clear
    On Error GoTo 0
End Sub

Private Function EncodeTextForPDF(ByRef sText As String, Optional ByVal lStartPos As Long = 1, _
            Optional ByVal lEndPos As Long = 0) As String
    Dim i As Long, sRes As String, sSym As String * 1, lPos As Long
    Dim lFrom As Long, lTo As Long, lTemp As Long, sRepl As String
    
    If Len(sText) = 0 Then EncodeTextForPDF = vbNullString: Exit Function
    
    lFrom = lStartPos: lTo = lEndPos
    
    If lTo > 0 And lTo < lFrom Then
        lTemp = lTo: lTo = lFrom: lFrom = lTemp
    End If
    If lTo = 0 Then lTo = Len(sText)
    If lTo > Len(sText) Then lTo = Len(sText)
    If lFrom < 1 Then lFrom = 1
    sRes = vbNullString
    For i = lFrom To lTo
        sSym = Mid$(sText, i, 1)
        lPos = InStr(1, sCyrAPDF, sSym, vbBinaryCompare)
        If lPos > 0 Then
            sRepl = sPDFencoding(lPos - 1)
            sRes = sRes & sRepl
        ElseIf sSym = Chr$(34) Then
            sRes = sRes & sSym
        Else
            sRes = sRes & "?"
        End If
    Next i
    EncodeTextForPDF = Left$(sText, lFrom - 1) & sRes & Right$(sText, Len(sText) - lTo)
End Function

Public Sub SetDecimalSeparator(ByVal sDecSeparator As String)
  Dim iLocale As Integer, sTmpStr As String, lRes As Long
  On Error Resume Next
  If Len(sDecSeparator) = 0 Then Exit Sub
  sTmpStr = sDecSeparator
  If Len(sTmpStr) > 4 Then sTmpStr = Left$(sTmpStr, 4)
  sTmpStr = sTmpStr & Chr$(0)
  iLocale = GetUserDefaultLCID()
  lRes = SetLocaleInfo(iLocale, LOCALE_SDECIMAL, sTmpStr)
  Err.Clear
  On Error GoTo 0
End Sub

Public Function GetDecimalSeparator() As String
  Dim iLocale As Integer, sTmpStr As String, lRes As Long, aLen As Long
  On Error Resume Next
  sTmpStr = String$(255, " ") & Chr$(0)
  aLen = 1
  iLocale = GetUserDefaultLCID()
  lRes = GetLocaleInfo(iLocale, LOCALE_SDECIMAL, sTmpStr, aLen)
  GetDecimalSeparator = Left$(sTmpStr, aLen)
  Err.Clear
  On Error GoTo 0
End Function

Private Function FodlersIsWrong() As Boolean
    Dim bbRes As Boolean
    bbRes = False
    For i = 1 To MAX_PROJECT_NUM
        bbRes = bbRes Or (sInputFolder(i) = "C:\")
        bbRes = bbRes Or (sOutputFolder(i) = "C:\")
        bbRes = bbRes Or (sInputFolder(i) = sOutputFolder(i))
        bbRes = bbRes Or (FSO.FolderExists(sInputFolder(i)) = False)
        bbRes = bbRes Or (FSO.FolderExists(sOutputFolder(i)) = False)
    Next i
    FodlersIsWrong = bbRes
End Function

Private Function SplitLongString(ByRef sLongString As String, ByVal dSizeKoef As Single, _
        ByVal iMaxRows As Integer, ByVal dMaxLen As Double) As Variant
    Dim aRes() As String, dCurrLen As Double, sCurrString As String, lPos As Long, i As Long
    sCurrString = sLongString
    ReDim aRes(1 To iMaxRows)
    
    For i = 1 To dMaxLen
        dCurrLen = Len(sCurrString) * dSizeKoef
        If dCurrLen > dMaxLen Then ' here we must make a dance with booben :)))
            lPos = Len(sCurrString)
            Do
                lPos = InStrRev(sCurrString, " ", lPos - 1, vbTextCompare)
                If lPos = 0 Then Exit Do
            Loop While lPos * dSizeKoef > dMaxLen
            If lPos < 3 Then aRes(i) = sCurrString: Exit For
            aRes(i) = Left$(sCurrString, lPos - 1)
            sCurrString = Mid$(sCurrString, lPos + 1, Len(sCurrString))
        Else
            aRes(i) = sCurrString
            Exit For
        End If
    Next i
    SplitLongString = aRes
End Function


Private Function ReadProjects() As Boolean
    Dim i As Integer
    On Error Resume Next
    MAX_PROJECT_NUM = 0
    
    If FSO.FileExists(App.Path & "\" & TXT_PROJECTS) Then
        Open App.Path & "\" & TXT_PROJECTS For Input As #1
            Input #1, MAX_PROJECT_NUM
        Close #1
    End If
    
    If MAX_PROJECT_NUM = 0 Then 'we dont have any projects added! Creating example
        MAX_PROJECT_NUM = 1
        ReDim sProjects(1 To 1)
        'ReDim bProjectIsFotoMeCompat(1 To 1)
        sProjects(1) = "PROJECT1"
        Open App.Path & "\" & TXT_PROJECTS For Output As #1
            Print #1, "1"
            Print #1, sProjects(1)
        Close #1
        FSO.CopyFile App.Path & "\blank.jpg", App.Path & "\PROJECT1.jpg", True
    End If
    
    Open App.Path & "\" & TXT_PROJECTS For Input As #1
        Input #1, MAX_PROJECT_NUM
        ReDim sProjects(1 To MAX_PROJECT_NUM)
        ReDim sINP(1 To MAX_PROJECT_NUM)
        ReDim sOUTP(1 To MAX_PROJECT_NUM)
        ReDim sInputFolder(1 To MAX_PROJECT_NUM)
        ReDim sOutputFolder(1 To MAX_PROJECT_NUM)
        ReDim bProjectIsFotoMeCompat(1 To MAX_PROJECT_NUM)
        For i = 1 To MAX_PROJECT_NUM
            Input #1, sProjects(i)
            If Trim$(sProjects(i)) = vbNullString Then sProjects(i) = "PROJECT" & CStr(i)
            bProjectIsFotoMeCompat(i) = False
            If Right$(sProjects(i), 1) = Chr$(149) Then
                bProjectIsFotoMeCompat(i) = True
                sProjects(i) = Left$(sProjects(i), Len(sProjects(i)) - 1)
            End If
        Next i
    Close #1
    If Err.Number <> 0 Then
        ReadProjects = False
    Else
        ReadProjects = True
    End If
    Err.Clear
    On Error GoTo 0
End Function

Private Function WriteProjects() As Boolean
    Dim i As Integer, iFN As Integer
    On Error Resume Next
    
    iFN = FreeFile
    Open App.Path & "\" & TXT_PROJECTS For Output As iFN
        Print #iFN, MAX_PROJECT_NUM
        For i = 1 To MAX_PROJECT_NUM
            If Trim$(sProjects(i)) = vbNullString Then sProjects(i) = "PROJECT" & CStr(i)
            If bProjectIsFotoMeCompat(i) Then
                Print #iFN, sProjects(i) & Chr$(149)
            Else
                Print #iFN, sProjects(i)
            End If
        Next i
    Close iFN
    
    If Err.Number <> 0 Then
        WriteProjects = False
    Else
        WriteProjects = True
    End If
    Err.Clear
    On Error GoTo 0
End Function

Private Function CleanStringFromWrongSymbols(ByVal sStringToAnalyze) As String
    Dim j As Integer, sRes As String, lLen As Long, sTmp As String
    lLen = Len(sStringToAnalyze)
    sRes = vbNullString
    For j = 1 To lLen
        sTmp = Mid$(sStringToAnalyze, j, 1)
        If ((Asc(sTmp) >= &H30) And (Asc(sTmp) <= &H39)) Or _
                ((Asc(sTmp) >= &H41) And (Asc(sTmp) <= &H5A)) Or _
                ((Asc(sTmp) >= &H61) And (Asc(sTmp) <= &H7A)) Then
            sRes = sRes & sTmp
        End If
    Next j
    CleanStringFromWrongSymbols = sRes
End Function

Private Sub RemoveProject(ByVal lIndex As Long)
    Dim sTmp As String, iii As Integer
    On Error Resume Next
    FSO.DeleteFile App.Path & "\" & sProjects(lIndex) & ".jpg", True
    FSO.DeleteFile App.Path & "\" & sProjects(lIndex) & "_preview.jpg", True
    For iii = lIndex To MAX_PROJECT_NUM - 1
        sProjects(iii) = sProjects(iii + 1)
        sInputFolder(iii) = sInputFolder(iii + 1)
        sOutputFolder(iii) = sOutputFolder(iii + 1)
        bProjectIsFotoMeCompat(iii) = bProjectIsFotoMeCompat(iii + 1)
    Next iii
    MAX_PROJECT_NUM = MAX_PROJECT_NUM - 1
    ReDim Preserve sProjects(1 To MAX_PROJECT_NUM)
    ReDim Preserve sINP(1 To MAX_PROJECT_NUM)
    ReDim Preserve sOUTP(1 To MAX_PROJECT_NUM)
    ReDim Preserve sInputFolder(1 To MAX_PROJECT_NUM)
    ReDim Preserve sOutputFolder(1 To MAX_PROJECT_NUM)
    ReDim Preserve bProjectIsFotoMeCompat(1 To MAX_PROJECT_NUM)
    Me.lstProjects.ListItems.Remove lIndex
    Err.Clear
    On Error GoTo 0
    WriteProjects
    WriteHotFolders
    Me.lstProjects.ListItems.Clear
    For iii = 1 To MAX_PROJECT_NUM
        Me.lstProjects.ListItems.Add iii, sProjects(iii), sProjects(iii)
    Next iii
    Me.lstProjects.ListItems(1).Selected = True
    Call lstProjects_ItemClick(Me.lstProjects.ListItems(1))
End Sub

'Private Function ParsePDF(ByVal sFileIN As String, _
                    ByVal xmlRoot As Variant, _
                    Optional ByRef in_JPG As Variant = False, _
                    Optional ByVal bProjectIsFotoMeCompatible As Boolean = False) As Variant
Private Function ParsePDF(ByVal sFileIN As String, _
                    Optional ByRef in_JPG As Variant = False, _
                    Optional ByVal bProjectIsFotoMeCompatible As Boolean = False) As Variant
    Dim sTempString As String, varArray As Variant, strVarArrayItem As String, bBuff() As Byte
    Dim ii As Long, lFileLength As Long, jj As Long, kk As Long, lStrings As Long
    Dim vResultArray() As String, iFileNumber As Integer
    Dim lBeginStream As Long, lEndStream As Long, lBeginLength As Long, lRealLength As Long
    Dim xrefs_old() As Long, xrefs_new() As Long, ll As Long
    Dim sDecode As String, childNode As Object, sFileOut As String, sBinding As String
    'next vars is used for ensuring that filename is present in XML
    'Dim bFileNameProcessed As Boolean, bFileNameFound As Boolean
    
    On Error Resume Next
    
    iFileNumber = FreeFile
    lFileLength = FileLen(sFileIN)

    ReDim bBuff(1 To lFileLength)
    Open sFileIN For Binary Access Read As iFileNumber
        Get #iFileNumber, , bBuff
    Close iFileNumber

FormSys.TrayIcon = Me.pictWorking

    'converting byte array to string
    sTempString = StrConv(bBuff, vbUnicode, 1049)
    
    'split file by vbLf
    varArray = Split(sTempString, vbCr)
    ii = 0: jj = 0: lStrings = UBound(varArray)
    ReDim vResultArray(1 To 1)
    'analize strings and collect all object data together
    strVarArrayItem = vbNullString
    Do While ii <= lStrings
            strVarArrayItem = varArray(ii)
            'if srting looks like "NNN 0 obj" then it is a begin of PDF object
            If (strVarArrayItem Like "? 0 obj") Or (strVarArrayItem Like "?? 0 obj") Or _
                (strVarArrayItem Like "??? 0 obj") Or (strVarArrayItem Like "???? 0 obj") Or _
                (strVarArrayItem = "xref") Or (ii < 3) Then
                jj = jj + 1
                ReDim Preserve vResultArray(1 To jj)
                vResultArray(jj) = strVarArrayItem
            Else 'else we must to append all strings until next object
                vResultArray(jj) = vResultArray(jj) & vbCr & strVarArrayItem
            End If
        ii = ii + 1
    Loop
    'after all we have JJ PDF objects
    lStrings = jj
    
    'trying to get all OLD xref addresses of each object
    varArray = Split(vResultArray(lStrings), vbCrLf)
    ReDim xrefs_old(1 To 1)
    kk = 0
    strVarArrayItem = vbNullString
    For ii = 0 To UBound(varArray)
        strVarArrayItem = varArray(ii)
        If (strVarArrayItem Like "########## *") Then
'        If Val(varArray(ii)) > 0 Then
            sTempString = Left$(strVarArrayItem, 10)
            If Val(sTempString) > 0 Then
                kk = kk + 1
                ReDim Preserve xrefs_old(1 To kk)
                xrefs_old(kk) = CLng(Val(sTempString))
            End If
        End If
    Next ii
    kk = kk + 1
    ReDim Preserve xrefs_old(1 To kk)
    xrefs_old(kk) = CLng(Val(varArray(ii - 3)))
    
    
    If bProjectIsFotoMeCompatible Then
        
        ' Here will be a part for FotoMe compatible projects
        vResultArray(7) = Replace(vResultArray(7), "          ", " ")
        
    Else
        'another one change - font size for ORDER_ID filed
        If Val(Me.cmbFont.Text) <> 15 Then
            vResultArray(7) = Replace(vResultArray(7), "3.937 0 Td" & vbLf & "[(order_id)]TJ", _
                            Format$(Val(Me.cmbFont.Text), "00") & _
                            " 0 0 " & Format$(Val(Me.cmbFont.Text), "00") & _
                            " 341.1033 578.3567 Tm" & vbLf & "[(order_id)]TJ", 1, 1)
            vResultArray(7) = Replace(vResultArray(7), "3.937 0 Td" & vbLf & "[(order_id)]TJ", _
                            Format$(Val(Me.cmbFont.Text), "00") & _
                            " 0 0 " & Format$(Val(Me.cmbFont.Text), "00") & _
                            " 341.1033 239.4861 Tm" & vbLf & "[(order_id)]TJ", 1, 1)
                            
        End If
    
    End If
    
    
'    vResultArray(7) = Replace(vResultArray(7), _
            "[(\037\036\035\034)53(\033)3(\032\031\030)-10(\027\026\025\024)-24(\023\022\021\020\017\016\r)-20(\f)-22(\013)3(\n)25(\t\b)15(\007\006\005)3(\004)81(\003\002\001\201\215\217\220)10(\235\240\255\200)43(\202)5(\203\204\205)-18(\206\207\210\211)-25(\212\213\214\216\221\222\223)-24(\224)-8(\225)11(\226)13(\227\230)10(\231\232\233\234)57(\236\237)-5(\241\242\243\244\245\246)]TJ", _
            "[( )]TJ")
    
'    vResultArray(7) = Replace(vResultArray(7), _
            "[(abc)6(def)3(ghijklmnopqrstuv)-20(wx)-19(yzABCDEFGHIJKLMNOPQRSTUVWX)-12(YZ)]TJ", _
            "[( )]TJ")
    
'    vResultArray(7) = Replace(vResultArray(7), _
            "(\253 1234567890-+=_`~!@#$%^&*\(\)[]{};:\247\273\\|/?,.<>" & Chr$(34) & ")Tj", _
            "[( )]TJ")
    
'    vResultArray(7) = Replace(vResultArray(7), _
            "[(\037\036)3(\035\034)43(\033\032\031\030)-17(\027\026\025\024)-25(\023\022\021\020\017\016\r)-27(\f)-33(\013)-3(\n)16(\t\b)15(\007\006\005\004)82(\003\002)-4(\001\201\215\217\220)-4(\235\240\255\200)41(\202)4(\203\204)4(\205)-28(\206\207\210\211)-18(\212\213\214\216\221\222\223)-30(\224)-11(\225)12(\226)18(\227\230)13(\231\232\233\234)61(\236\237)-8(\241\242)6(\243\244\245\246)]TJ", _
            "[( )]TJ")
    
'    vResultArray(7) = Replace(vResultArray(7), _
            "[(abcdef)8(ghijklmnopqrstuv)-27(wx)-23(yzABCDEFGHIJKLMNOPQRSTUVWX)-32(YZ)]TJ", _
            "[( )]TJ")
    
'    vResultArray(7) = Replace(vResultArray(7), _
            "(\253 1234567890-+=_`~!@#$%^&*\(\)[]{};:\247\273\\|/?,.<>" & Chr$(34) & ")Tj", _
            "[( )]TJ")
    
        sFileOut = vbNullString
        sBinding = vbNullString
        
        'bFileNameProcessed = False
        'bFileNameFound = False
                
        'For Each childNode In xmlRoot.childNodes
        For Each childNode In currNode.childNodes
                
                If Trim$(childNode.baseName) = "orderID" Then
                    sFileOut = App.Path & "\ORDER_" & childNode.Text
'                    bFileNameFound = True
                End If
                
                If Trim$(childNode.baseName) = "item_id_1" Then
                    sFileOut = sFileOut & "_itemID_" & Format$(Val(childNode.Text), "000")
'                    bFileNameFound = True
                End If
                
                
        
            If bProjectIsFotoMeCompatible Then
                
                If childNode.baseName = "item_Filename" Then
                    strVarArrayItem = vbNullString
                    strVarArrayItem = Trim$(childNode.getAttribute("value"))
                    If Err.Number = 0 Then
                        'strVarArrayItem = Split(strVarArrayItem, "_itemID_", , vbTextCompare)(0)
                        strVarArrayItem = Replace(strVarArrayItem, ".pdf", vbNullString, , , vbTextCompare)
                    Else
                        Err.Clear
                        strVarArrayItem = childNode.Text
                        'strVarArrayItem = Split(strVarArrayItem, "_itemID_", , vbTextCompare)(0)
                        strVarArrayItem = Replace(strVarArrayItem, ".pdf", vbNullString, , , vbTextCompare)
                    End If
                    If Len(Trim$(strVarArrayItem)) > 0 Then sFileOut = App.Path & "\" & strVarArrayItem
'                    bFileNameProcessed = True
                End If

                
            Else
                
                If childNode.baseName = "file_name" Then
                    strVarArrayItem = vbNullString
                    strVarArrayItem = childNode.getAttribute("value")
                    Err.Clear
                    If Len(Trim$(strVarArrayItem)) > 0 Then sFileOut = App.Path & "\" & strVarArrayItem
 '                   bFileNameProcessed = True
                End If
            
            End If
            

'            If childNode.baseName = "job_phb_binding" Then
'                Select Case UCase$(Trim$(childNode.getAttribute("value")))
                    'so many error may be in this one short word :(
                    'first line - one english letter (in 4 words - C  K  O  A)
                    'second line - two english letters (in 6 words - CK  CO CA KO KA OA)
                    'third line - three english letters (in 4 word - CKO CKA KOA COA)
                    'fourth line - four english letter (in 1 word - CKOA)
                    'last line - right russian word :)
'                    Case "CÊÎÁÀ", "ÑKÎÁÀ", "ÑÊOÁÀ", "ÑÊÎÁA", _
                         "CKÎÁÀ", "CÊOÁÀ", "CÊÎÁA", "ÑKOÁÀ", "ÑKÎÁA", "ÑÊOÁA", _
                         "CKOÁÀ", "CKÎÁA", "ÑKOÁA", "CÊOÁA", _
                         "CKOÁA", _
                         "ÑÊÎÁÀ"

'                        sBinding = "_SADDLE-1"
'                    Case Else
'                        sBinding = "_PERFECT-1"
'                End Select
'            End If

            'so many time spented to fucking SKOBA... And now it is not important :(  Life is a big surprise
            sBinding = "_label"

            If Len(Trim$(childNode.getAttribute("value"))) > 0 Then
                If Err.Number = 0 Then
                    sTempString = EncodeTextForPDF(Trim$(childNode.getAttribute("value")))
                Else
                    Err.Clear
                    If Len(Trim$(childNode.Text)) > 0 Then
                        If Err.Number = 0 Then
                            sTempString = EncodeTextForPDF(Trim$(childNode.Text))
                        Else
                            Err.Clear
                            sTempString = " "
                        End If
                    End If
                End If
            Else
                sTempString = " "
            End If
            If Err.Number <> 0 Then
                Err.Clear
                sTempString = " "
            End If
            
            
            'item_unitPrice|item_quantity|item_discountQuantity|item_subtotal
            'here we must replace fuckin UAH sign with nothing
            If bProjectIsFotoMeCompatible Then
                If Trim$(childNode.baseName) = "item_unitPrice" Or Trim$(childNode.baseName) = "item_discountQuantity" Or _
                    Trim$(childNode.baseName) = "item_subtotal" Then
                    'MsgBox Asc(Left$(sTempString, 1))
                    sTempString = Replace(sTempString, Chr$(63), vbNullString)
                    sTempString = Replace(sTempString, "ãðí", vbNullString)
                    sTempString = Replace(sTempString, "\200\222\214", vbNullString)
                End If
                'paymentType
                If Trim$(childNode.baseName) = "paymentType" Then
                    sTempString = Replace(sTempString, "InStore", EncodeTextForPDF("Ïðè ïîëó÷åíèè"))
                    sTempString = Replace(sTempString, "ContraReembolso", EncodeTextForPDF("Ïðè ïîëó÷åíèè"))
                    sTempString = Replace(sTempString, "BankTransfer", EncodeTextForPDF("Áàíêîâñêèé ïåðåâîä"))
                    sTempString = Replace(sTempString, "BankTranfer", EncodeTextForPDF("Áàíêîâñêèé ïåðåâîä"))
                    'some new items
                    sTempString = Replace(sTempString, "BANK_TRANSFER", EncodeTextForPDF("Áàíêîâñêèé ïåðåâîä"))
                    sTempString = Replace(sTempString, "bank_transfer", EncodeTextForPDF("Áàíêîâñêèé ïåðåâîä"))
                    sTempString = Replace(sTempString, "STORE_PAYMENT", EncodeTextForPDF("Ïðè ïîëó÷åíèè"))
                    sTempString = Replace(sTempString, "store_payment", EncodeTextForPDF("Ïðè ïîëó÷åíèè"))
                    sTempString = Replace(sTempString, "PAYMENT_DELIVERY", EncodeTextForPDF("Ïðè ïîëó÷åíèè"))
                    sTempString = Replace(sTempString, "payment_delivery", EncodeTextForPDF("Ïðè ïîëó÷åíèè"))
                    sTempString = Replace(sTempString, "MONEY.UA", EncodeTextForPDF("Îïëàòà îíëàéí"))
                    sTempString = Replace(sTempString, "money.ua", EncodeTextForPDF("Îïëàòà îíëàéí"))
                    sTempString = Replace(sTempString, "CREDIT CARD", EncodeTextForPDF("Îïëàòà îíëàéí"))
                    sTempString = Replace(sTempString, "credit card", EncodeTextForPDF("Îïëàòà îíëàéí"))
                    sTempString = Replace(sTempString, "CREDIT_CARD", EncodeTextForPDF("Îïëàòà îíëàéí"))
                    sTempString = Replace(sTempString, "credit_card", EncodeTextForPDF("Îïëàòà îíëàéí"))
                    'old items
                    sTempString = Replace(sTempString, "null", EncodeTextForPDF("Êðåäèòíàÿ êàðòà"))
                    sTempString = Replace(sTempString, "none", " ")
                End If
                If Trim$(childNode.baseName) = "deliveryType" Then
                    sTempString = Replace(sTempString, "StorePickUp", EncodeTextForPDF("Ñàìîâûâîç"))
                    sTempString = Replace(sTempString, "CustomerAddress", EncodeTextForPDF("Ñëóæáà äîñòàâêè"))
                End If
                If Trim$(childNode.baseName) = "store_name" Then
                    If Len(sTempString) > 4 Then
                        If Left$(sTempString, 4) = "0 - " Then
                            sTempString = Right$(sTempString, Len(sTempString) - 4)
                        End If
                    End If
                End If
            End If
            
            If UCase$(Trim$(sTempString)) = "NONE" Then sTempString = " "
            
            'here we doing (sick!) replacement childNames in PDF with its values :)))
            vResultArray(7) = Replace(vResultArray(7), Trim$(childNode.baseName), sTempString)
        Next childNode
        
        'here we must cleanup PDF from empty fields names - just replace them with space
        If bProjectIsFotoMeCompatible Then
            varArray = Split(sFiledsFotoMe, "|")
        Else
            varArray = Split(sFileds, "|")
        End If
        strVarArrayItem = vbNullString
        For ii = 0 To UBound(varArray)
            strVarArrayItem = varArray(ii)
            vResultArray(7) = Replace(vResultArray(7), strVarArrayItem, " ")
        Next ii
        
        If Not bProjectIsFotoMeCompatible Then
            '1 0 1 0.5 k - fucking new changes - make green text black :(
            vResultArray(7) = Replace(vResultArray(7), "1 0 1 0.5 k", "0 0 0 1 k")
        End If
        
        If Len(sFileOut) = 0 Then
            'sFileOut = Left$(sFileIN, Len(sFileIN) - 4)
            ParsePDF = False
            Err.Clear
            On Error GoTo 0
            FormSys.TrayIcon = "DEFAULT"
            Exit Function
        End If
        
        sFileOut = sFileOut & sBinding & ".pdf"

    
    If TypeName(in_JPG) <> "Boolean" Then
        sDecode = vbNullString
        If in_JPG(9) = True Then 'do invert image - fucking Adobe inverted format
            Select Case in_JPG(3)
                Case "DeviceCMYK"
                    sDecode = "/Decode [1.0 0.0 1.0 0.0 1.0 0.0 1.0 0.0]"
                Case "DeviceRGB"
                    sDecode = "/Decode [1.0 0.0 1.0 0.0 1.0 0.0]"
                Case "DeviceGray"
                    sDecode = "/Decode [1.0 0.0]"
            End Select
        End If
        'JPEG object in usual shablon have number 13, in FotoMe shablon - 12
        'So, let's do the trick - val of boolean is 0 or -1, add it to 13
        'vResultArray(8) = "13 0 obj" & vbCr & "<</BitsPerComponent " & CStr(in_JPG(4)) &
        vResultArray(8) = CStr(13 - -(bProjectIsFotoMeCompatible)) & " 0 obj" & vbCr & "<</BitsPerComponent " & CStr(in_JPG(4)) & _
                    "/ColorSpace/" & in_JPG(3) & "/Filter/" & in_JPG(5) & sDecode & _
                    "/Height " & CStr(in_JPG(2)) & "/Length " & CStr(in_JPG(8)) & _
                    "/Name/X/Subtype/Image/Type/XObject/Width " & CStr(in_JPG(1)) & _
                    ">>stream" & vbCrLf & in_JPG(6) & vbCrLf & "endstream" & vbCr & "endobj"
    End If
    
    'trying to get real length of each
    For ii = 1 To lStrings
        If (vResultArray(ii) Like "? 0 obj*") Or (vResultArray(ii) Like "?? 0 obj*") Or _
                (vResultArray(ii) Like "??? 0 obj*") Or (vResultArray(ii) Like "???? 0 obj*") Then
            lBeginLength = InStr(1, vResultArray(ii), "/Length ", vbBinaryCompare)
            If lBeginLength > 0 Then
                lBeginStream = InStr(lBeginLength + 7, vResultArray(ii), ">>stream", vbBinaryCompare) + 10
                lEndStream = InStr(lBeginStream + 8, vResultArray(ii), "endstream", vbBinaryCompare) - 2
                'kk is a digits quantity in old string (in "/Length 2345" kk is 4 etc)
                kk = Len(CStr(Val(Mid$(vResultArray(ii), lBeginLength + 7, Len(vResultArray(ii))))))
                'lRealLength  is a real length of a new stream
                lRealLength = lEndStream - lBeginStream
                vResultArray(ii) = Left$(vResultArray(ii), lBeginLength + 7) & CStr(lRealLength) & _
                            Mid$(vResultArray(ii), lBeginLength + 8 + kk, Len(vResultArray(ii)))
            End If
        End If
    Next ii
    
    'trying to get all NEW xref addresses of each object and correct XREF links
    ReDim xrefs_new(1 To UBound(xrefs_old))
    kk = 0
    For ii = 1 To jj
        If (vResultArray(ii) Like "? 0 obj*") Or (vResultArray(ii) Like "?? 0 obj*") Or _
                (vResultArray(ii) Like "??? 0 obj*") Or (vResultArray(ii) Like "???? 0 obj*") Or _
                (vResultArray(ii) Like "?xref*") Or (vResultArray(ii) Like "xref*") Then
            kk = CLng(Val(Split(vResultArray(ii), " ", 2)(0)))
            If kk > 4 Then kk = kk - 1
            If (vResultArray(ii) Like "?xref*") Or (vResultArray(ii) Like "xref*") Then
                kk = UBound(xrefs_new)
            End If
            xrefs_new(kk) = 0
            For ll = 1 To ii - 1
                xrefs_new(kk) = xrefs_new(kk) + Len(vResultArray(ll)) + 1
            Next ll
            If Left$(vResultArray(ii), 1) = vbLf Then xrefs_new(kk) = xrefs_new(kk) + 1
        End If
    Next ii
    'now we correct XREFs
    kk = UBound(xrefs_new)
    For ii = 1 To kk - 1
        vResultArray(UBound(vResultArray)) = Replace(vResultArray(UBound(vResultArray)), _
                    Format$(xrefs_old(ii), "0000000000"), Format$(xrefs_new(ii), "0000000000"))
    Next ii
    'at last we have to correct self XREF reference
    vResultArray(UBound(vResultArray)) = Replace(vResultArray(UBound(vResultArray)), _
                    "startxref" & vbCrLf & CStr(xrefs_old(kk)), "startxref" & vbCrLf & CStr(xrefs_new(kk)))

    iFileNumber = FreeFile
    Open sFileOut For Output As #iFileNumber
        For ii = 1 To jj
            Print #iFileNumber, vResultArray(ii) & vbCr;
        Next ii
    Close #iFileNumber
    
    Close
    
    FormSys.TrayIcon = "DEFAULT"
    If Err.Number <> 0 Then
        Err.Clear
        ParsePDF = False
    Else
        ParsePDF = sFileOut
    End If
    On Error GoTo 0
End Function


Private Function ParseJPG(pFileName As String) As Variant

Dim in_Bytes   As Long

Dim str_TChar  As String
Dim in_res     As Long

Dim sIMG       As Long
Dim inIMG

Dim in_PEnd     As Long
Dim in_idx      As Long
Dim str_SegmMk  As String
Dim in_SegmSz   As Long
Dim bChar       As Byte
Dim in_TmpColor As Long
Dim in_bpc      As Long

Dim ArrBFile()  As Byte

Dim ArrIMG As aIMG

    ' Extract info from a JPEG file
    inIMG = FreeFile

    sIMG = FileLen(pFileName)

    If sIMG < 250 Then
        MsgBox "File Image is non JPEG" & _
                vbCrLf & _
                "Cannot add image to PDF file.", vbCritical, "Univest Digital"
        ParseJPG = False
        Exit Function
    End If

    ArrIMG.in_8 = sIMG
    ArrIMG.in_9_ycck = False

    ReDim Preserve ArrBFile(1 To sIMG)
    Open pFileName For Binary Access Read As #inIMG
        Get #inIMG, , ArrBFile
    Close #inIMG

    in_PEnd = UBound(ArrBFile) - 1

    If IntAsHex(ArrBFile, 1) <> "FFD8" Or IntAsHex(ArrBFile, in_PEnd) <> "FFD9" Then
        MsgBox "Invalid JPEG marker" & _
                vbCrLf & _
                "Cannot add image to PDF file.", vbCritical, "Univest Digitall"
        ParseJPG = False
        Exit Function
    End If

    in_idx = 3
    Do While in_idx < in_PEnd
        str_SegmMk = IntAsHex(ArrBFile, in_idx)
        in_SegmSz = IntVal(ArrBFile, in_idx + 2)

        If str_SegmMk = "FFFF" Then
            Do While ArrBFile(in_idx + 1) = &HFF
                in_idx = in_idx + 1
            Loop
            in_SegmSz = IntVal(ArrBFile, in_idx + 2)
        End If

        Select Case str_SegmMk
            Case "FFE0"
                bChar = ArrBFile(in_idx + 11)
                If bChar = 0 Then
                    ArrIMG.in_7 = "Dots"
                ElseIf bChar = 1 Then
                    ArrIMG.in_7 = "Dots/inch (DPI)"
                ElseIf bChar = 2 Then
                    ArrIMG.in_7 = "Dots/cm"
                Else
                    MsgBox "Invalid resolution data in image!" & vbCrLf & _
                            "Cannot add image to PDF file.", vbCritical, "Univest Digitall"
                    ParseJPG = False
                    Exit Function
                End If
            Case "FFC0", "FFC1", "FFC2", "FFC3", "FFC5", "FFC6", "FFC7"
                ArrIMG.in_1 = IntVal(ArrBFile, in_idx + 7)
                ArrIMG.in_2 = IntVal(ArrBFile, in_idx + 5)

                in_TmpColor = ArrBFile(in_idx + 9) * 8

                If in_TmpColor = 8 Then
                    ArrIMG.in_3 = "DeviceGray"
                ElseIf in_TmpColor = 24 Then
                    ArrIMG.in_3 = "DeviceRGB"
                ElseIf in_TmpColor = 32 Then
                    ArrIMG.in_3 = "DeviceCMYK"
                Else
                    MsgBox "Invalid color mode in image! Must be CMYK, RGB or Gray." & vbCrLf & _
                            "Cannot add image to PDF file.", vbCritical, "Univest Digitall"
                    ParseJPG = False
                    Exit Function
                End If
            Case "FFEE" 'if last byte of this block equal to 2 then we have fucking Adobe CMYK-YCCK color transformed JPEG
                        'Direct inject this JPEG into PDF produce INVERTED image (sucks!)
                bChar = ArrBFile(in_idx + in_SegmSz + 1)
                If bChar = 2 Then 'sucks YCCK!
                    ArrIMG.in_9_ycck = True
                End If
        End Select

        in_idx = in_idx + in_SegmSz + 2
    Loop

    If ArrIMG.in_4 <> "" Then
        in_bpc = ArrIMG.in_4
    Else
        in_bpc = 8
        ArrIMG.in_4 = 8
    End If

    ArrIMG.in_5 = "DCTDecode"
    ArrIMG.in_6 = ""

    Open pFileName For Binary As #inIMG
        str_TChar = String(sIMG, " ")
        Get #inIMG, , str_TChar
        ArrIMG.in_6 = ArrIMG.in_6 & str_TChar
    Close #inIMG

    ParseJPG = Array(ArrIMG.in_1, _
                     ArrIMG.in_2, _
                     ArrIMG.in_3, _
                     in_bpc, ArrIMG.in_5, _
                     ArrIMG.in_6, _
                     ArrIMG.in_7, _
                     ArrIMG.in_8, _
                     ArrIMG.in_9_ycck)

End Function

Private Function IntAsHex(ArrBF As Variant, in_Index As Long) As String

    IntAsHex = Right("00" & Hex(ArrBF(in_Index)), 2) & _
                  Right("00" & Hex(ArrBF(in_Index + 1)), 2)

End Function

Private Function IntVal(ArrBF As Variant, in_idx As Long) As Long

    IntVal = CLng(ArrBF(in_idx)) * 256& + _
                CLng(ArrBF(in_idx + 1))

End Function

