VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmBalanceExpress 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   Icon            =   "frmBalanceExpress.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   11910
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   3045
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   12195
      _ExtentX        =   21511
      _ExtentY        =   5371
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin StockMarket.uctlDate uctlFromDate 
         Height          =   405
         Left            =   1980
         TabIndex        =   0
         Top             =   1020
         Width           =   2800
         _ExtentX        =   4948
         _ExtentY        =   714
      End
      Begin MSComctlLib.ProgressBar prgProgress 
         Height          =   195
         Left            =   1980
         TabIndex        =   2
         Top             =   1680
         Width           =   9075
         _ExtentX        =   16007
         _ExtentY        =   344
         _Version        =   393216
         Appearance      =   1
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   12195
         _ExtentX        =   21511
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin StockMarket.uctlTextBox txtPercent 
         Height          =   465
         Left            =   1980
         TabIndex        =   3
         Top             =   2040
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   820
      End
      Begin StockMarket.uctlDate uctlToDate 
         Height          =   405
         Left            =   7170
         TabIndex        =   1
         Top             =   1020
         Width           =   2800
         _ExtentX        =   4948
         _ExtentY        =   714
      End
      Begin VB.Label lblToDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   5880
         TabIndex        =   12
         Top             =   1080
         Width           =   1215
      End
      Begin Threed.SSCommand cmdStart 
         Height          =   525
         Left            =   7800
         TabIndex        =   4
         Top             =   2100
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmBalanceExpress.frx":27A2
         ButtonStyle     =   3
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   345
         Left            =   3720
         TabIndex        =   11
         Top             =   2160
         Width           =   1275
      End
      Begin VB.Label lblProgress 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   315
         Left            =   360
         TabIndex        =   10
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label lblPercent 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   330
         TabIndex        =   9
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label lblFromDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   90
         TabIndex        =   8
         Top             =   1080
         Width           =   1815
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   9495
         TabIndex        =   5
         Top             =   2100
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmBalanceExpress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_HasModify As Boolean

Public ID As Long
Public OKClick As Boolean
Public ShowMode As SHOW_MODE_TYPE
Public HeaderText As String
Private Sub cmdOK_Click()
   OKClick = True
   Unload Me
End Sub
Private Sub cmdStart_Click()
Dim Status As Boolean
Dim PartItemID As Long
   
   If Not VerifyDate(lblFromDate, uctlFromDate, False) Then
      Exit Sub
   End If
      
   Call glbDaily.StartTransaction
   
   Me.Enabled = False
          
   Status = UpdateFromExpress(uctlFromDate.ShowDate, uctlToDate.ShowDate)
   
   Me.Enabled = True

   If Status Then
      Call glbDaily.CommitTransaction
      glbErrorLog.LocalErrorMsg = "การอัฟเดดจาก EXPRESS เสร็จสมบูรณ์"
      glbErrorLog.ShowUserError
   Else
      Call glbDaily.RollbackTransaction
      glbErrorLog.LocalErrorMsg = "การอัฟเดดจาก EXPRESS มีข้อผิดพลาด"
      glbErrorLog.ShowUserError
   End If

   Call cmdOK_Click
   Exit Sub

End Sub

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
   End If

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = DUMMY_KEY Then
      glbErrorLog.LocalErrorMsg = Me.Name
      glbErrorLog.ShowUserError
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 116 Then
'      Call cmdSearch_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 115 Then
'      Call cmdClear_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 118 Then
'      Call cmdAdd_Click
      KeyCode = 0
'   ElseIf Shift = 0 And KeyCode = 117 Then
'      Call cmdDelete_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 113 Then
      Call cmdOK_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 114 Then
'      Call cmdEdit_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 121 Then
'      Call cmdPrint_Click
      KeyCode = 0
   End If

End Sub

Private Sub ResetStatus()
   prgProgress.Max = 100
   prgProgress.Min = 0
   prgProgress.Value = 0
   txtPercent.Text = 0
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   HeaderText = "นำเข้าเช็คจาก ระบบ EXPRESS V4"
   Me.Caption = HeaderText
   pnlHeader.Caption = HeaderText

   Me.Picture = LoadPicture(glbParameterObj.MainPicture)

   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   Call InitNormalLabel(lblFromDate, "จากวันที่เช็ค", RGB(255, 0, 0))
   Call InitNormalLabel(lblProgress, "ความคืบหน้า")
   Call InitNormalLabel(lblPercent, "เปอร์เซนต์")
   Call InitNormalLabel(Label1, "%")
   Call InitNormalLabel(lblToDate, "ถึงวันที่เช็ค")
   
   Call txtPercent.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   txtPercent.Enabled = False

   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
  ' cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdStart.Picture = LoadPicture(glbParameterObj.NormalButton1)

   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
  ' Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdStart, MapText("เริ่ม"))

   Call ResetStatus
End Sub

Private Sub cmdExit_Click()
   If Not ConfirmExit(m_HasModify) Then
      Exit Sub
   End If

   OKClick = False
   Unload Me
End Sub

Private Sub Form_Load()
   Call EnableForm(Me, False)
   m_HasActivate = False
   
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub
Public Function UpdateFromExpress(Optional FromDate As Date = -1, Optional ToDate As Date = -1) As Boolean
On Error GoTo ErrorHandler
Dim m_Rs As ADODB.Recordset
Dim IsOK As Boolean
Dim iCount As Long
Dim RecordCount As Long
Dim PERCENT As Double
Dim I As Long
Dim HasBegin As Boolean
Dim Result As Boolean
Dim Bt As CBkTrn
Dim Pcl As CPrintChequeLog
Dim Pos As Long
Dim Pn As CPayName
Dim TmColl As Collection
Dim TmNam As CBkTrn
   Set TmColl = New Collection
   Set m_Rs = New ADODB.Recordset
   Set Bt = New CBkTrn
   Bt.FROM_CHQDAT = FromDate
   Bt.TO_CHQDAT = ToDate
   Bt.JNLTRNTYP = "1" 'เช็คจ่ายเท่านั้น
   
   Call Bt.QueryData(1, m_Rs, iCount)
   
   '
   Set Pcl = New CPrintChequeLog
   HasBegin = True
   
   prgProgress.Min = 1
   prgProgress.Max = m_Rs.RecordCount
   
   I = 0
   While Not m_Rs.EOF
      I = I + 1
      Call Bt.PopulateFromRS(1, m_Rs)
      
      If CheckChequeNoCanPrint(Bt.CHQNUM) Then
      
         Pcl.CHEQUE_NO = Bt.CHQNUM
         Pcl.CHEQUE_AMOUNT = Bt.AMOUNT
         Pcl.CHEQUE_DATE = Bt.CHQDAT
         
         Set TmNam = GetObject("CBkTrn", TmColl, Trim(Bt.BNKNUM2), False)
         If TmNam Is Nothing Then
            'ใส่ ประเภทพิมพ์
            frmAddChequeType.AccNo = Bt.BNKNUM2
            frmAddChequeType.AccName = Bt.BNKNUM1
            Load frmAddChequeType
            frmAddChequeType.Show 1
            
            Unload frmAddChequeType
            
            Set TmNam = New CBkTrn
            TmNam.BNKNAM = frmAddChequeType.ChequeType
            TmNam.BNKNUM2 = Bt.BNKNUM2
            If Len(TmNam.BNKNAM) > 0 Then
               Pcl.CHEQUE_TYPE = TmNam.BNKNAM
               Call TmColl.Add(TmNam, Trim(Bt.BNKNUM2))
            Else
               UpdateFromExpress = False
               Exit Function
            End If
            Set frmAddChequeType = Nothing
         Else
            Pcl.CHEQUE_TYPE = TmNam.BNKNAM
         End If
         
         Pcl.PAYEE_NAME = Bt.SUPNAM
         If Len(Pcl.PAYEE_NAME) <= 0 Then
            'ให้ตัดเอา ข้อความที่อยู่ข้างหน้า * มาเท่านั้น
            Pos = InStr(1, Bt.REMARK, "*")
            If Pos > 0 Then
               Pcl.PAYEE_NAME = Mid(Bt.REMARK, 1, Pos - 1)
            Else
               Pcl.PAYEE_NAME = Bt.REMARK
            End If
         End If
         
         Pcl.AddEditMode = SHOW_ADD
         Call Pcl.AddEditData
         
         Set Pn = GetObject("CPayName", glbPayeeName, Pcl.PAYEE_NAME, False)
         If Pn Is Nothing Then
            Set Pn = New CPayName
            Pn.PAY_NAME = Pcl.PAYEE_NAME
            Call glbPayeeName.Add(Pn, Pcl.PAYEE_NAME)
            
            Pn.AddEditMode = SHOW_ADD
            Call Pn.AddEditData
         End If
         
      End If
      
      prgProgress.Value = I
      txtPercent.Text = MyDiff(I, m_Rs.RecordCount) * 100
      Me.Refresh
      
      m_Rs.MoveNext
   Wend

   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   
   prgProgress.Value = prgProgress.Max
   txtPercent.Text = "100"
   
   Set Pcl = Nothing
   
   Set m_Rs = Nothing
   
   UpdateFromExpress = True
   Exit Function

ErrorHandler:
   If HasBegin Then
   End If

   glbErrorLog.LocalErrorMsg = "Runtime error."
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.RoutineName = "UpdateCapitalMovement"
   glbErrorLog.ModuleName = "frmBalanceExpress"
   glbErrorLog.LocalErrorMsg = "Eror"
   glbErrorLog.ShowErrorLog (LOG_MSGBOX)

   Set Pcl = Nothing
   
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing

   UpdateFromExpress = False
End Function
