VERSION 5.00
Begin VB.Form frmExample 
   Caption         =   "�˺� Ȩ�ý� ����(����)��꼭 ���� API SDK"
   ClientHeight    =   11295
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12315
   LinkTopic       =   "Form1"
   ScaleHeight     =   11295
   ScaleWidth      =   12315
   StartUpPosition =   2  'ȭ�� ���
   Begin VB.Frame Frame6 
      Caption         =   "Ȩ�ý� ����(����)��꼭 ���� ���� API"
      Height          =   7575
      Left            =   240
      TabIndex        =   23
      Top             =   3480
      Width           =   11775
      Begin VB.Frame Frame10 
         Caption         =   "�ΰ����"
         Height          =   2415
         Left            =   9000
         TabIndex        =   40
         Top             =   360
         Width           =   2600
         Begin VB.CommandButton btnGetCertificatePopUpURL 
            Caption         =   "���������� ��� URL"
            Height          =   410
            Left            =   120
            TabIndex        =   44
            Top             =   1280
            Width           =   2295
         End
         Begin VB.CommandButton btnGetFlatRateState 
            Caption         =   "������ ���� ���� Ȯ��"
            Height          =   410
            Left            =   120
            TabIndex        =   43
            Top             =   800
            Width           =   2295
         End
         Begin VB.CommandButton btnGetFlatRatePopUPURL 
            Caption         =   "������ ���� ��û URL"
            Height          =   410
            Left            =   120
            TabIndex        =   42
            Top             =   300
            Width           =   2295
         End
         Begin VB.CommandButton btnGetCertificateExpireDate 
            Caption         =   "���������� �������� Ȯ��"
            Height          =   410
            Left            =   120
            TabIndex        =   41
            Top             =   1760
            Width           =   2295
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "����(����)��꼭 ������ ��ȸ"
         Height          =   1935
         Left            =   4800
         TabIndex        =   35
         Top             =   360
         Width           =   3975
         Begin VB.CommandButton btnGetXML 
            Caption         =   "������ ��ȸ - XML"
            Height          =   410
            Left            =   1800
            TabIndex        =   39
            Top             =   1200
            Width           =   1935
         End
         Begin VB.CommandButton btnGetTaxinvoice 
            Caption         =   "������ ��ȸ"
            Height          =   410
            Left            =   240
            TabIndex        =   38
            Top             =   1200
            Width           =   1455
         End
         Begin VB.TextBox txtNtsconfirmNum 
            Height          =   300
            Left            =   1560
            TabIndex        =   37
            Top             =   300
            Width           =   2295
         End
         Begin VB.Label Label6 
            Caption         =   "('������� ��ȸ'�� ��ȯ�� ����(����)��꼭 ����û���ι�ȣ�� �Է��Ͻñ� �ٶ��ϴ�.)"
            Height          =   375
            Left            =   120
            TabIndex        =   45
            Top             =   720
            Width           =   3735
         End
         Begin VB.Label Label5 
            Caption         =   "����û���ι�ȣ : "
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.ListBox taxinvoiceInfo 
         Height          =   4020
         Left            =   240
         TabIndex        =   33
         Top             =   3000
         Width           =   11295
      End
      Begin VB.Frame Frame8 
         Caption         =   "����/���� ������� ��ȸ"
         Height          =   1935
         Left            =   2280
         TabIndex        =   29
         Top             =   360
         Width           =   2415
         Begin VB.CommandButton btnSummary 
            Caption         =   "���� ��� ������� ��ȸ"
            Height          =   410
            Left            =   120
            TabIndex        =   31
            Top             =   840
            Width           =   2175
         End
         Begin VB.CommandButton btnSearch 
            Caption         =   "���� ��� ��ȸ"
            Height          =   410
            Left            =   120
            TabIndex        =   30
            Top             =   300
            Width           =   2175
         End
      End
      Begin VB.TextBox txtJobID 
         Height          =   300
         Left            =   2000
         TabIndex        =   28
         Top             =   2560
         Width           =   2175
      End
      Begin VB.Frame Frame7 
         Caption         =   "����/���� ���� ����"
         Height          =   1900
         Left            =   120
         TabIndex        =   24
         Top             =   360
         Width           =   2055
         Begin VB.CommandButton btnRequestJob 
            Caption         =   "���� ��û"
            Height          =   375
            Left            =   120
            TabIndex        =   32
            Top             =   320
            Width           =   1815
         End
         Begin VB.CommandButton btnGetJobState 
            Caption         =   "���� ���� Ȯ��"
            Height          =   410
            Left            =   120
            TabIndex        =   26
            Top             =   800
            Width           =   1815
         End
         Begin VB.CommandButton btnListActiveJob 
            Caption         =   "���� ���� ��� Ȯ��"
            Height          =   410
            Left            =   120
            TabIndex        =   25
            Top             =   1300
            Width           =   1815
         End
      End
      Begin VB.Label Label4 
         Caption         =   "(�۾����̵�� '���� ��û' ȣ��� �����˴ϴ� )"
         Height          =   255
         Left            =   4300
         TabIndex        =   34
         Top             =   2620
         Width           =   4095
      End
      Begin VB.Label Label3 
         Caption         =   "�۾����̵� (jobID) :"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   2640
         Width           =   1695
      End
   End
   Begin VB.TextBox txtCorpNum 
      Height          =   315
      Left            =   2400
      TabIndex        =   0
      Text            =   "1234567890"
      Top             =   255
      Width           =   1935
   End
   Begin VB.TextBox txtUserID 
      Height          =   315
      Left            =   6120
      TabIndex        =   1
      Text            =   "testkorea"
      Top             =   240
      Width           =   1935
   End
   Begin VB.CommandButton btnCheckID 
      Caption         =   "ID �ߺ� Ȯ��"
      Height          =   410
      Left            =   480
      TabIndex        =   3
      Top             =   1870
      Width           =   1455
   End
   Begin VB.CommandButton btnListContact 
      Caption         =   "����� ��� ��ȸ"
      Height          =   410
      Left            =   4920
      TabIndex        =   9
      Top             =   1870
      Width           =   2055
   End
   Begin VB.CommandButton btnUpdateContact 
      Caption         =   "����� ���� ����"
      Height          =   410
      Left            =   4920
      TabIndex        =   10
      Top             =   2350
      Width           =   2055
   End
   Begin VB.CommandButton btnUpdateCorpInfo 
      Caption         =   "ȸ������ ����"
      Height          =   410
      Left            =   9720
      TabIndex        =   15
      Top             =   1870
      Width           =   1935
   End
   Begin VB.CommandButton btnPopbillURL_CHRG 
      Caption         =   " ����Ʈ ���� URL"
      Height          =   410
      Left            =   7320
      TabIndex        =   12
      Top             =   1870
      Width           =   2055
   End
   Begin VB.Frame Frame15 
      Caption         =   "ȸ������ ����"
      Height          =   1935
      Left            =   9600
      TabIndex        =   14
      Top             =   1035
      Width           =   2240
      Begin VB.CommandButton btnGetCorpInfo 
         Caption         =   "ȸ������ ��ȸ"
         Height          =   410
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " �˺� �⺻ API "
      Height          =   2535
      Left            =   240
      TabIndex        =   16
      Top             =   675
      Width           =   11775
      Begin VB.Frame Frame5 
         Caption         =   " �˺� �⺻ URL"
         Height          =   1935
         Left            =   6960
         TabIndex        =   20
         Top             =   360
         Width           =   2295
         Begin VB.CommandButton btnGetPopbillURL_LOGIN 
            Caption         =   " �˺� �α��� URL"
            Height          =   410
            Left            =   120
            TabIndex        =   11
            Top             =   360
            Width           =   2055
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "����� ����"
         Height          =   1935
         Left            =   4560
         TabIndex        =   19
         Top             =   360
         Width           =   2295
         Begin VB.CommandButton btnRegistContact 
            Caption         =   "����� �߰�"
            Height          =   410
            Left            =   120
            TabIndex        =   8
            Top             =   360
            Width           =   2055
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   " ����Ʈ ����"
         Height          =   1935
         Left            =   1920
         TabIndex        =   18
         Top             =   360
         Width           =   2535
         Begin VB.CommandButton btnGetChargeInfo 
            Caption         =   "�������� Ȯ��"
            Height          =   410
            Left            =   120
            TabIndex        =   7
            Top             =   1320
            Width           =   2295
         End
         Begin VB.CommandButton btnGetBalance 
            Caption         =   "�ܿ�����Ʈ Ȯ��"
            Height          =   410
            Left            =   120
            TabIndex        =   5
            Top             =   360
            Width           =   2295
         End
         Begin VB.CommandButton btnGetPartnerBalance 
            Caption         =   "��Ʈ�� �ܿ�����Ʈ Ȯ��"
            Height          =   410
            Left            =   120
            TabIndex        =   6
            Top             =   840
            Width           =   2295
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   " ȸ������"
         Height          =   1935
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   1695
         Begin VB.CommandButton btnJoinMember 
            Caption         =   "ȸ�� ����"
            Height          =   410
            Left            =   120
            TabIndex        =   4
            Top             =   1320
            Width           =   1455
         End
         Begin VB.CommandButton btnCheckIsMember 
            Caption         =   "���� ���� Ȯ��"
            Height          =   410
            Left            =   120
            TabIndex        =   2
            Top             =   360
            Width           =   1455
         End
      End
   End
   Begin VB.Label Label1 
      Caption         =   "�˺�ȸ�� ����ڹ�ȣ :"
      Height          =   180
      Left            =   360
      TabIndex        =   22
      Top             =   315
      Width           =   1920
   End
   Begin VB.Label Label2 
      Caption         =   "�˺�ȸ�� ���̵� : "
      Height          =   180
      Left            =   4560
      TabIndex        =   21
      Top             =   315
      Width           =   1500
   End
End
Attribute VB_Name = "frmExample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'��ũ���̵�
Private Const linkID = "TESTER"
'���Ű. ���⿡ �����Ͻñ� �ٶ��ϴ�.
Private Const SecretKey = "SwWxqU+0TErBXy/9TVjIPEnI0VTUMMSQZtJf3Ed8q3I="

Private htTaxinvoiceService As New PBHTTaxinvoiceService

Private Sub btnCheckID_Click()
    Dim Response As PBResponse
    
    Set Response = htTaxinvoiceService.CheckID(txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(htTaxinvoiceService.LastErrCode) + "] " + htTaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.Message)
End Sub

Private Sub btnCheckIsMember_Click()
    Dim Response As PBResponse
    
    Set Response = htTaxinvoiceService.CheckIsMember(txtCorpNum.Text, linkID)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(htTaxinvoiceService.LastErrCode) + "] " + htTaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.Message)
End Sub

Private Sub btnGetBalance_Click()
    Dim balance As Double
    
    balance = htTaxinvoiceService.GetBalance(txtCorpNum.Text)
    
    If balance < 0 Then
        
        MsgBox ("[" + CStr(htTaxinvoiceService.LastErrCode) + "] " + htTaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "�ܿ�����Ʈ : " + CStr(balance)
End Sub

Private Sub btnGetCertificateExpireDate_Click()
    Dim expireDate As String
    
    expireDate = htTaxinvoiceService.GetCertificateExpireDate(txtCorpNum.Text)
    
    If expireDate = "" Then
        
        MsgBox ("[" + CStr(htTaxinvoiceService.LastErrCode) + "] " + htTaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "������������ : " + expireDate
End Sub

Private Sub btnGetCertificatePopUpURL_Click()
    Dim url As String
    
    url = htTaxinvoiceService.GetCertificatePopUpURL(txtCorpNum.Text, txtUserID.Text)
    
    If url = "" Then
         MsgBox ("[" + CStr(htTaxinvoiceService.LastErrCode) + "] " + htTaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub

Private Sub btnGetChargeInfo_Click()
    Dim ChargeInfo As PBChargeInfo
    
    Set ChargeInfo = htTaxinvoiceService.GetChargeInfo(txtCorpNum.Text)
     
    If ChargeInfo Is Nothing Then
        MsgBox ("[" + CStr(htTaxinvoiceService.LastErrCode) + "] " + htTaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    
    tmp = tmp + "unitCost ([������]�����׿��) : " + ChargeInfo.unitCost + vbCrLf
    tmp = tmp + "chargeMethod (��������) : " + ChargeInfo.chargeMethod + vbCrLf
    tmp = tmp + "rateSystem (��������) : " + ChargeInfo.rateSystem + vbCrLf
    
    MsgBox tmp
End Sub

Private Sub btnGetCorpInfo_Click()
    Dim CorpInfo As PBCorpInfo
    
    Set CorpInfo = htTaxinvoiceService.GetCorpInfo(txtCorpNum.Text, txtUserID.Text)
     
    If CorpInfo Is Nothing Then
        MsgBox ("[" + CStr(htTaxinvoiceService.LastErrCode) + "] " + htTaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    
    tmp = tmp + "ceoname : " + CorpInfo.ceoname + vbCrLf
    tmp = tmp + "corpName : " + CorpInfo.corpName + vbCrLf
    tmp = tmp + "addr : " + CorpInfo.addr + vbCrLf
    tmp = tmp + "bizType : " + CorpInfo.bizType + vbCrLf
    tmp = tmp + "bizClass : " + CorpInfo.bizClass + vbCrLf
    
    MsgBox tmp
End Sub

Private Sub btnGetFlatRatePopUPURL_Click()
    Dim url As String
    
    url = htTaxinvoiceService.GetFlatRatePopUpURL(txtCorpNum.Text, txtUserID.Text)
    
    If url = "" Then
         MsgBox ("[" + CStr(htTaxinvoiceService.LastErrCode) + "] " + htTaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
    
End Sub

Private Sub btnGetFlatRateState_Click()
    Dim flatRateInfo As PBHTTaxinvoiceFlatRate
    
    Set flatRateInfo = htTaxinvoiceService.GetFlatRateState(txtCorpNum.Text)
     
    If flatRateInfo Is Nothing Then
        MsgBox ("[" + CStr(htTaxinvoiceService.LastErrCode) + "] " + htTaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    
    tmp = tmp + "referencdeID (����ڹ�ȣ) : " + flatRateInfo.referenceID + vbCrLf
    tmp = tmp + "contractDT (������ ���� �����Ͻ�) : " + flatRateInfo.contractDT + vbCrLf
    tmp = tmp + "useEndDate (������ ���� ������) : " + flatRateInfo.useEndDate + vbCrLf
    tmp = tmp + "baseDate (�ڵ����� ������) : " + CStr(flatRateInfo.baseDate) + vbCrLf
    tmp = tmp + "state (������ ���� ����) : " + CStr(flatRateInfo.state) + vbCrLf
    tmp = tmp + "closeRequestYN (���� ������û ����) : " + CStr(flatRateInfo.closeRequestYN) + vbCrLf
    tmp = tmp + "useRestrictYN (���� ������� ����) : " + CStr(flatRateInfo.useRestrictYN) + vbCrLf
    tmp = tmp + "closeOnExpired (���񽺸���� �������� ) : " + CStr(flatRateInfo.closeOnExpired) + vbCrLf
    tmp = tmp + "unPaidYN (�̼��� ���� ����) : " + CStr(flatRateInfo.unPaidYN) + vbCrLf
    
    MsgBox tmp
End Sub

Private Sub btnGetJobState_Click()
    Dim jobInfo As PBHTTaxinvoiceJobState
    
    Set jobInfo = htTaxinvoiceService.GetJobState(txtCorpNum.Text, txtJobID.Text)
     
    If jobInfo Is Nothing Then
        MsgBox ("[" + CStr(htTaxinvoiceService.LastErrCode) + "] " + htTaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    
    tmp = tmp + "jobID(�۾����̵�) : " + jobInfo.jobID + vbCrLf
    tmp = tmp + "jobState(��������) : " + CStr(jobInfo.jobState) + vbCrLf
    tmp = tmp + "queryType(��������) : " + jobInfo.queryType + vbCrLf
    tmp = tmp + "queryDateType(��������) : " + jobInfo.queryDateType + vbCrLf
    tmp = tmp + "queryStDate(��������) : " + jobInfo.queryStDate + vbCrLf
    tmp = tmp + "queryEnDate(��������) : " + jobInfo.queryEnDate + vbCrLf
    tmp = tmp + "errorCode(�����ڵ�) : " + CStr(jobInfo.errorCode) + vbCrLf
    tmp = tmp + "errorReason(�����޽���) : " + jobInfo.errorReason + vbCrLf
    tmp = tmp + "jobStartDT(�۾� �����Ͻ�) : " + jobInfo.jobStartDT + vbCrLf
    tmp = tmp + "jobEndDT(�۾� �����Ͻ�) : " + jobInfo.jobEndDT + vbCrLf
    tmp = tmp + "collectCount(��������) : " + CStr(jobInfo.collectCount) + vbCrLf
    tmp = tmp + "regDT(���� ��û�Ͻ�) : " + jobInfo.regDT + vbCrLf
    
    MsgBox tmp
End Sub

Private Sub btnGetPartnerBalance_Click()
    Dim balance As Double
    
    balance = htTaxinvoiceService.GetPartnerBalance(txtCorpNum.Text)
    
    If balance < 0 Then
        MsgBox ("[" + CStr(htTaxinvoiceService.LastErrCode) + "] " + htTaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "�ܿ�����Ʈ : " + CStr(balance)
End Sub

Private Sub btnGetPopbillURL_LOGIN_Click()
    Dim url As String
    
    url = htTaxinvoiceService.GetPopbillURL(txtCorpNum.Text, txtUserID.Text, "LOGIN")
    
    If url = "" Then
         MsgBox ("[" + CStr(htTaxinvoiceService.LastErrCode) + "] " + htTaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub

Private Sub btnGetTaxinvoice_Click()
    Dim taxinvoiceInfo As PBHTTaxinvoice
    Dim i As Integer
    
    Set taxinvoiceInfo = htTaxinvoiceService.GetTaxinvoice(txtCorpNum.Text, txtNtsconfirmNum.Text)
     
    If taxinvoiceInfo Is Nothing Then
        MsgBox ("[" + CStr(htTaxinvoiceService.LastErrCode) + "] " + htTaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    '����(����)��꼭 �׸� ���� �ڼ��� ������ [�˺� Ȩ�ý� ����(����)��꼭 ���� API �����Ŵ��� > 4.1.2 ������ Ȯ��]�� �����Ͻñ� �ٶ��ϴ�
    
    tmp = "========����(����)��꼭 ����=======" + vbCrLf
    tmp = tmp + "writeDate : " + taxinvoiceInfo.writeDate + vbCrLf
    tmp = tmp + "issueDT : " + taxinvoiceInfo.issueDT + vbCrLf
    tmp = tmp + "invoiceType : " + taxinvoiceInfo.invoiceType + vbCrLf
    tmp = tmp + "taxType : " + taxinvoiceInfo.taxType + vbCrLf
    tmp = tmp + "taxTotal : " + taxinvoiceInfo.taxTotal + vbCrLf
    tmp = tmp + "supplyCostTotal : " + taxinvoiceInfo.supplyCostTotal + vbCrLf
    tmp = tmp + "totalAmount : " + taxinvoiceInfo.totalAmount + vbCrLf
    tmp = tmp + "purposeType : " + taxinvoiceInfo.purposeType + vbCrLf
    tmp = tmp + "cash : " + taxinvoiceInfo.cash + vbCrLf
    tmp = tmp + "chkBill : " + taxinvoiceInfo.chkBill + vbCrLf
    tmp = tmp + "credit : " + taxinvoiceInfo.credit + vbCrLf
    tmp = tmp + "note : " + taxinvoiceInfo.note + vbCrLf
    tmp = tmp + "remark1 : " + taxinvoiceInfo.remark1 + vbCrLf
    tmp = tmp + "remark2 : " + taxinvoiceInfo.remark2 + vbCrLf
    tmp = tmp + "remark3 : " + taxinvoiceInfo.remark3 + vbCrLf
    tmp = tmp + "ntsconfirmNum : " + taxinvoiceInfo.ntsconfirmNum + vbCrLf + vbCrLf
    
    tmp = tmp + "========������ ����=======" + vbCrLf
    tmp = tmp + "invoicerCorpNum : " + taxinvoiceInfo.invoicerCorpNum + vbCrLf
    tmp = tmp + "invoicerMgtKey : " + taxinvoiceInfo.invoicerMgtKey + vbCrLf
    tmp = tmp + "invoicerTaxRegID : " + taxinvoiceInfo.invoicerTaxRegID + vbCrLf
    tmp = tmp + "invoicerCorpName : " + taxinvoiceInfo.invoicerCorpName + vbCrLf
    tmp = tmp + "invoicerCEOName : " + taxinvoiceInfo.invoicerCEOName + vbCrLf
    tmp = tmp + "invoicerAddr : " + taxinvoiceInfo.invoicerAddr + vbCrLf
    tmp = tmp + "invoicerBizType : " + taxinvoiceInfo.invoicerBizType + vbCrLf
    tmp = tmp + "invoicerBizClass : " + taxinvoiceInfo.invoicerBizClass + vbCrLf
    tmp = tmp + "invoicerContactName : " + taxinvoiceInfo.invoicerContactName + vbCrLf
    tmp = tmp + "invoicerDeptName : " + taxinvoiceInfo.invoicerDeptName + vbCrLf
    tmp = tmp + "invoicerTEL : " + taxinvoiceInfo.invoicerTEL + vbCrLf
    tmp = tmp + "invoicerEmail : " + taxinvoiceInfo.invoicerEmail + vbCrLf + vbCrLf
    
    tmp = tmp + "========���޹޴��� ����=======" + vbCrLf
    tmp = tmp + "invoiceeCorpNum : " + taxinvoiceInfo.invoiceeCorpNum + vbCrLf
    tmp = tmp + "invoiceeType : " + taxinvoiceInfo.invoiceeType + vbCrLf
    tmp = tmp + "invoiceeMgtKey : " + taxinvoiceInfo.invoiceeMgtKey + vbCrLf
    tmp = tmp + "invoiceeTaxRegID : " + taxinvoiceInfo.invoiceeTaxRegID + vbCrLf
    tmp = tmp + "invoiceeCorpName : " + taxinvoiceInfo.invoiceeCorpName + vbCrLf
    tmp = tmp + "invoiceeCEOName : " + taxinvoiceInfo.invoiceeCEOName + vbCrLf
    tmp = tmp + "invoiceeAddr : " + taxinvoiceInfo.invoiceeAddr + vbCrLf
    tmp = tmp + "invoiceeBizType : " + taxinvoiceInfo.invoiceeBizType + vbCrLf
    tmp = tmp + "invoiceeBizClass : " + taxinvoiceInfo.invoiceeBizClass + vbCrLf
    tmp = tmp + "invoiceeContactName1 : " + taxinvoiceInfo.invoiceeContactName1 + vbCrLf
    tmp = tmp + "invoiceeDeptName1 : " + taxinvoiceInfo.invoiceeDeptName1 + vbCrLf
    tmp = tmp + "invoiceeTEL1 : " + taxinvoiceInfo.invoiceeTEL1 + vbCrLf
    tmp = tmp + "invoiceeEmail1 : " + taxinvoiceInfo.invoiceeEmail1 + vbCrLf
    tmp = tmp + "invoiceeContactName2 : " + taxinvoiceInfo.invoiceeContactName2 + vbCrLf
    tmp = tmp + "invoiceeDeptName2 : " + taxinvoiceInfo.invoiceeDeptName2 + vbCrLf
    tmp = tmp + "invoiceeTEL2 : " + taxinvoiceInfo.invoiceeTEL2 + vbCrLf
    tmp = tmp + "invoiceeEmail2 : " + taxinvoiceInfo.invoiceeEmail2 + vbCrLf + vbCrLf
        
        
    tmp = tmp + "========����(����)��꼭 ǰ��迭========" + vbCrLf
    Dim detailInfo As PBHTTaxinvoiceDetail
    
    For Each detailInfo In taxinvoiceInfo.detailList
        tmp = tmp + "serialNum : " + CStr(detailInfo.serialNum) + vbCrLf
        tmp = tmp + "purchaseDT : " + detailInfo.purchaseDT + vbCrLf
        tmp = tmp + "itemName : " + detailInfo.itemName + vbCrLf
        tmp = tmp + "spec : " + detailInfo.spec + vbCrLf
        tmp = tmp + "qty : " + detailInfo.qty + vbCrLf
        tmp = tmp + "unitCost : " + detailInfo.unitCost + vbCrLf
        tmp = tmp + "supplyCost : " + detailInfo.supplyCost + vbCrLf
        tmp = tmp + "tax : " + detailInfo.tax + vbCrLf
        tmp = tmp + "remark : " + detailInfo.remark + vbCrLf + vbCrLf
    Next
    
    MsgBox (tmp)
    
End Sub

Private Sub btnGetXML_Click()
    Dim taxinvoiceXML As PBHTTaxinvoiceXML
    Dim i As Integer
    
    Set taxinvoiceXML = htTaxinvoiceService.GetXML(txtCorpNum.Text, txtNtsconfirmNum.Text)
     
    If taxinvoiceXML Is Nothing Then
        MsgBox ("[" + CStr(htTaxinvoiceService.LastErrCode) + "] " + htTaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    
    tmp = tmp + "ResultCode (�����ڵ�) : " + CStr(taxinvoiceXML.ResultCode) + vbCrLf
    tmp = tmp + "Message (����û���ι�ȣ) : " + taxinvoiceXML.Message + vbCrLf
    tmp = tmp + "retObject (XML����) : " + taxinvoiceXML.retObject + vbCrLf
    
    MsgBox (tmp)
End Sub

Private Sub btnJoinMember_Click()
    Dim joinData As New PBJoinForm
    Dim Response As PBResponse
    
    joinData.linkID = linkID '��ũ ���̵�
    joinData.CorpNum = "1231212312" '����ڹ�ȣ "-" ����.
    joinData.ceoname = "��ǥ�ڼ���"
    joinData.corpName = "ȸ����ȣ"
    joinData.addr = "�ּ�"
    joinData.bizType = "����"
    joinData.bizClass = "����"
    joinData.id = "userid"      '6�� �̻� 20�� �̸�.
    joinData.pwd = "pwd_must_be_long_enough"    '6�� �̻� 20�� �̸�.
    joinData.ContactName = "����ڼ���"
    joinData.ContactTEL = "02-999-9999"
    joinData.ContactHP = "010-1234-5678"
    joinData.ContactFAX = "02-999-9998"
    joinData.ContactEmail = "test@test.com"
    
    Set Response = htTaxinvoiceService.JoinMember(joinData)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(htTaxinvoiceService.LastErrCode) + "] " + htTaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox (Response.Message)
End Sub

Private Sub btnListActiveJob_Click()
    Dim jobList As Collection
        
    Set jobList = htTaxinvoiceService.ListActiveJob(txtCorpNum.Text)
     
    If jobList Is Nothing Then
        MsgBox ("[" + CStr(htTaxinvoiceService.LastErrCode) + "] " + htTaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    
    tmp = tmp + "�۾����̵�(jobID)�� ��ȿ�ð��� 1�ð��Դϴ�" + vbCrLf + vbCrLf
    tmp = tmp + "jobID | jobState | queryType | queryDateType | queryStDate | queryEnDate | errorCode | errorReason | jobStartDT | jobEndDT | collectCount | regDT " + vbCrLf
    
    Dim info As PBHTTaxinvoiceJobState
    
    For Each info In jobList
        tmp = tmp + CStr(info.jobID) + " | "
        tmp = tmp + CStr(info.jobState) + " | "
        tmp = tmp + info.queryType + " | "
        tmp = tmp + info.queryDateType + " | "
        tmp = tmp + info.queryStDate + " | "
        tmp = tmp + info.queryEnDate + " | "
        tmp = tmp + CStr(info.errorCode) + " | "
        tmp = tmp + info.errorReason + " | "
        tmp = tmp + info.jobStartDT + " | "
        tmp = tmp + info.jobEndDT + " | "
        tmp = tmp + CStr(info.collectCount) + " | "
        tmp = tmp + info.regDT
        tmp = tmp + vbCrLf
    Next
    
    MsgBox tmp
    
    If jobList.count > 0 Then
        txtJobID.Text = jobList.Item(1).jobID
    End If
       
End Sub

Private Sub btnListContact_Click()
    Dim resultList As Collection
        
    Set resultList = htTaxinvoiceService.ListContact(txtCorpNum.Text, txtUserID.Text)
     
    If resultList Is Nothing Then
        MsgBox ("[" + CStr(htTaxinvoiceService.LastErrCode) + "] " + htTaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    
    tmp = "id | email | hp | personName | searchAllAllowYN | tel | fax | mgrYN | regDT " + vbCrLf
    
    Dim info As PBContactInfo
    
    For Each info In resultList
        tmp = tmp + info.id + " | " + info.email + " | " + info.hp + " | " + info.personName + " | " + CStr(info.searchAllAllowYN) _
                + info.tel + " | " + info.fax + " | " + CStr(info.mgrYN) + " | " + info.regDT + vbCrLf
    Next
    
    MsgBox tmp
End Sub

Private Sub btnPopbillURL_CHRG_Click()
    Dim url As String
    
    url = htTaxinvoiceService.GetPopbillURL(txtCorpNum.Text, txtUserID.Text, "CHRG")
    
    If url = "" Then
         MsgBox ("[" + CStr(htTaxinvoiceService.LastErrCode) + "] " + htTaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub

Private Sub btnRegistContact_Click()
    Dim joinData As New PBContactInfo
    Dim Response As PBResponse
    
    '����� ���̵�
    joinData.id = "testkorea_20151007"
    
    '����� ��й�ȣ
    joinData.pwd = "test@test.com"
    
    '����ڸ�
    joinData.personName = "����ڸ�"
    
    '����ó
    joinData.tel = "070-1234-1234"
    
    '�޴�����ȣ
    joinData.hp = "010-1234-1234"
    
    '�̸��� �ּ�
    joinData.email = "test@test.com"
    
    '�ѽ���ȣ
    joinData.fax = "070-1234-1234"
    
    '��ü��ȸ ����, true-ȸ����ȸ, false-������ȸ
    joinData.searchAllAllowYN = True
    
    '������ ���ѿ���
    joinData.mgrYN = False
        
    Set Response = htTaxinvoiceService.RegistContact(txtCorpNum.Text, joinData, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(htTaxinvoiceService.LastErrCode) + "] " + htTaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.Message)
End Sub

Private Sub btnRequestJob_Click()
    Dim jobID As String
    Dim DType As String
    Dim SDate As String
    Dim EDate As String
    Dim tiType As KeyType
    
    '����(����)��꼭 ����, SELL-����, BUY-����, TURSTEE-����Ź
    tiType = SELL
    
    '��������, W-�ۼ�����, I-��������, S-��������
    DType = "W"
    
    '��������, ǥ������(yyyyMMdd)
    SDate = "20160501"
    
    '��������, ǥ������(yyyyMMdd)
    EDate = "20160701"
        
        
    '�۾����̵�(jobID)�� ��ȿ�ð��� 1�ð��Դϴ�.
    jobID = htTaxinvoiceService.RequestJob(txtCorpNum.Text, tiType, DType, SDate, EDate)
    
    If jobID = "" Then
         MsgBox ("[" + CStr(htTaxinvoiceService.LastErrCode) + "] " + htTaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "jobID(�۾����̵�) : " + jobID + vbCrLf
    
    
    txtJobID.Text = jobID
    
End Sub

Private Sub btnSearch_Click()
    Dim SearchList As PBHTTaxinvoiceSearch
    Dim tiType As New Collection
    Dim taxType As New Collection
    Dim purposeType As New Collection
    Dim taxRegIDType As String
    Dim taxRegID As String
    Dim taxRegIDYN As String
    Dim page As Integer
    Dim perPage As Integer
    Dim order As String
    Dim tmp As String
    Dim listboxRow As String
    
    
    '�������� �迭, N-�Ϲ�, M-����
    tiType.Add "N"
    tiType.Add "M"
    
    '�������� �迭, T-����, N-�鼼, Z-����
    taxType.Add "T"
    taxType.Add "N"
    taxType.Add "Z"
    
    '����/û�� �迭, R-����, C-û��, N-����
    purposeType.Add "R"
    purposeType.Add "C"
    purposeType.Add "N"
    
    '��������ȣ ����� ����, S-������, B-���޹޴���, T-��Ź��
    taxRegIDType = "S"
    
    '��������ȣ �޸�(,)�� �����Ͽ� ���� ex) 0001,0002
    taxRegID = ""
    
    '������� ����, ����-��ü��ȸ, 0-��������ȣ ����, 1-��������ȣ ��ȸ
    taxRegIDYN = ""
    
    '������ ��ȣ
    page = 1
    
    '�������� �˻�����, �ִ� 1000��
    perPage = 20
    
    '���� ����, D-��������, A-��������
    order = "D"
        
        
    'Search ȣ��
    Set SearchList = htTaxinvoiceService.Search(txtCorpNum.Text, txtJobID.Text, tiType, taxType, purposeType, taxRegIDType, taxRegID, taxRegIDYN, page, perPage, order)
    
        
    If SearchList Is Nothing Then
        MsgBox ("[" + CStr(htTaxinvoiceService.LastErrCode) + "] " + htTaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = "code : " + CStr(SearchList.code) + vbCrLf
    tmp = tmp + "message : " + SearchList.Message + vbCrLf
    tmp = tmp + "total : " + CStr(SearchList.total) + vbCrLf
    tmp = tmp + "perPage : " + CStr(SearchList.perPage) + vbCrLf
    tmp = tmp + "pageNum : " + CStr(SearchList.pageNum) + vbCrLf
    tmp = tmp + "pageCount : " + CStr(SearchList.pageCount) + vbCrLf + vbCrLf
    
    taxinvoiceInfo.Clear
    
    taxinvoiceInfo.AddItem "�ۼ����� | �������� | �������� | �ŷ�ó | ��Ϲ�ȣ | �������� | ���ް��� | �������� | ����û���ι�ȣ", 0
    
    Dim tiInfo As PBHTTaxinvoiceAbbr
           
    For Each tiInfo In SearchList.list
        ' �߰����� ����(����)��꼭 �׸��� [Ȩ�ý� ����(����)��꼭 ���� API �����Ŵ��� > 4.1.���䱸������] �� �����Ͻñ� �ٶ��ϴ�.'
        listboxRow = ""
        listboxRow = tiInfo.writeDate + " | "
        listboxRow = listboxRow + tiInfo.issueDate + " | "
        listboxRow = listboxRow + tiInfo.sendDate + " | "
        listboxRow = listboxRow + tiInfo.invoiceeCorpName + " | "
        listboxRow = listboxRow + tiInfo.invoiceeCorpNum + " | "
        listboxRow = listboxRow + tiInfo.taxType + " | "
        listboxRow = listboxRow + tiInfo.supplyCostTotal + " | "
        
        If tiInfo.modifyYN Then
            listboxRow = listboxRow + "���� | "
        Else
            listboxRow = listboxRow + "�Ϲ� | "
        End If
        
        listboxRow = listboxRow + tiInfo.ntsconfirmNum
        
        taxinvoiceInfo.AddItem listboxRow, taxinvoiceInfo.ListCount
        
    Next
            
    MsgBox (tmp)
    
End Sub

Private Sub btnSummary_Click()
    Dim summaryInfo As PBHTTaxinvoiceSummary
    Dim tiType As New Collection
    Dim taxType As New Collection
    Dim purposeType As New Collection
    Dim taxRegIDType As String
    Dim taxRegID As String
    Dim taxRegIDYN As String
    Dim tmp As String
    
    
    '�������� �迭, N-�Ϲ�, M-����
    tiType.Add "N"
    tiType.Add "M"
    
    '�������� �迭, T-����, N-�鼼, Z-����
    taxType.Add "T"
    taxType.Add "N"
    taxType.Add "Z"
    
    '����/û�� �迭, R-����, C-û��, N-����
    purposeType.Add "R"
    purposeType.Add "C"
    purposeType.Add "N"
    
    '��������ȣ ����� ����, S-������, B-���޹޴���, T-��Ź��
    taxRegIDType = "S"
    
    '��������ȣ, �޸�(,)�� �����Ͽ� ���� ex) 0001,0002
    taxRegID = ""
    
    '������� ����, ����-��ü��ȸ, 0-��������ȣ ����, 1-��������ȣ ��ȸ
    taxRegIDYN = ""
      
        
        
    'Summary ȣ��
    Set summaryInfo = htTaxinvoiceService.Summary(txtCorpNum.Text, txtJobID.Text, tiType, taxType, purposeType, taxRegIDType, taxRegID, taxRegIDYN)
    
        
    If summaryInfo Is Nothing Then
        MsgBox ("[" + CStr(htTaxinvoiceService.LastErrCode) + "] " + htTaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = "count (��������Ǽ�) : " + CStr(summaryInfo.count) + vbCrLf
    tmp = tmp + "supplyCostTotal (���ް��� �հ�) : " + CStr(summaryInfo.supplyCostTotal) + vbCrLf
    tmp = tmp + "taxTotal (���� �հ�) : " + CStr(summaryInfo.taxTotal) + vbCrLf
    tmp = tmp + "amountTotal (�հ� �ݾ�) : " + CStr(summaryInfo.amountTotal) + vbCrLf
       
            
    MsgBox (tmp)
End Sub

Private Sub btnUpdateContact_Click()
    Dim joinData As New PBContactInfo
    Dim Response As PBResponse
    
    '����ڸ�
    joinData.personName = "����ڸ�_����"
    
    '����ó
    joinData.tel = "070-1234-1234"
    
    '�޴�����ȣ
    joinData.hp = "010-1234-1234"
    
    '�̸��� �ּ�
    joinData.email = "test@test.com"
    
    '�ѽ���ȣ
    joinData.fax = "070-1234-1234"
    
    '��ü��ȸ����, True-ȸ����ȸ, False-������ȸ
    joinData.searchAllAllowYN = True
    
    '������ ���ѿ���
    joinData.mgrYN = False
                
    Set Response = htTaxinvoiceService.UpdateContact(txtCorpNum.Text, joinData, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(htTaxinvoiceService.LastErrCode) + "] " + htTaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.Message)
End Sub

Private Sub btnUpdateCorpInfo_Click()
    Dim CorpInfo As New PBCorpInfo
    Dim Response As PBResponse
    
    '��ǥ�� ����
    CorpInfo.ceoname = "��ǥ��"
    
    '��ȣ��
    CorpInfo.corpName = "��ȣ"
    
    '�ּ�
    CorpInfo.addr = "����Ư����"
    
    '����
    CorpInfo.bizType = "����"
    
    '����
    CorpInfo.bizClass = "����"
    
    Set Response = htTaxinvoiceService.UpdateCorpInfo(txtCorpNum.Text, CorpInfo, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(htTaxinvoiceService.LastErrCode) + "] " + htTaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.Message)
End Sub

Private Sub Form_Load()
    '��� �ʱ�ȭ
    htTaxinvoiceService.Initialize linkID, SecretKey
    
    '����ȯ�� ������ True(�׽�Ʈ��), False(�����)
    htTaxinvoiceService.IsTest = True
End Sub

