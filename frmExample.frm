VERSION 5.00
Begin VB.Form frmExample 
   Caption         =   "�˺� Ȩ�ý� ����(����)��꼭 ���Ը��� API SDK"
   ClientHeight    =   11280
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   17670
   LinkTopic       =   "Form1"
   ScaleHeight     =   11280
   ScaleWidth      =   17670
   StartUpPosition =   2  'ȭ�� ���
   Begin VB.Frame Frame6 
      Caption         =   "Ȩ�ý� ����(����)��꼭 ���� ���� API"
      Height          =   7575
      Left            =   240
      TabIndex        =   20
      Top             =   3480
      Width           =   17175
      Begin VB.Frame Frame13 
         Caption         =   "Ȩ�ý� �������� ���"
         Height          =   2415
         Left            =   11760
         TabIndex        =   48
         Top             =   360
         Width           =   5175
         Begin VB.CommandButton btnDeleteDeptUser 
            Caption         =   "�μ������ ������� ����"
            Height          =   375
            Left            =   2640
            TabIndex        =   55
            Top             =   1800
            Width           =   2295
         End
         Begin VB.CommandButton btnCheckLoginDeptUser 
            Caption         =   "�μ������ �α��� �׽�Ʈ"
            Height          =   375
            Left            =   2640
            TabIndex        =   54
            Top             =   1320
            Width           =   2295
         End
         Begin VB.CommandButton btnCheckDeptUser 
            Caption         =   "�μ������ ������� Ȯ��"
            Height          =   375
            Left            =   2640
            TabIndex        =   53
            Top             =   800
            Width           =   2295
         End
         Begin VB.CommandButton btnRegistDeptUser 
            Caption         =   "�μ������ �������"
            Height          =   375
            Left            =   2640
            TabIndex        =   52
            Top             =   300
            Width           =   2295
         End
         Begin VB.CommandButton btnCheckCertValidation 
            Caption         =   "���������� �α��� �׽�Ʈ"
            Height          =   375
            Left            =   120
            TabIndex        =   51
            Top             =   1320
            Width           =   2295
         End
         Begin VB.CommandButton btnGetCertificateExpireDate 
            Caption         =   "���������� �������� Ȯ��"
            Height          =   410
            Left            =   120
            TabIndex        =   50
            Top             =   800
            Width           =   2295
         End
         Begin VB.CommandButton btnGetCertificatePopUpURL 
            Caption         =   "Ȩ�ý����� �������� URL"
            Height          =   410
            Left            =   120
            TabIndex        =   49
            Top             =   300
            Width           =   2295
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "�ΰ����"
         Height          =   2415
         Left            =   9000
         TabIndex        =   37
         Top             =   360
         Width           =   2600
         Begin VB.CommandButton btnGetFlatRateState 
            Caption         =   "������ ���� ���� Ȯ��"
            Height          =   410
            Left            =   120
            TabIndex        =   39
            Top             =   800
            Width           =   2295
         End
         Begin VB.CommandButton btnGetFlatRatePopUPURL 
            Caption         =   "������ ���� ��û URL"
            Height          =   410
            Left            =   120
            TabIndex        =   38
            Top             =   300
            Width           =   2295
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "����(����)��꼭 ������ ��ȸ"
         Height          =   2415
         Left            =   4800
         TabIndex        =   32
         Top             =   360
         Width           =   3975
         Begin VB.CommandButton btnGetPopUpURL 
            Caption         =   "���ݰ�꼭 ���� �˾�"
            Height          =   410
            Left            =   1920
            TabIndex        =   47
            Top             =   1800
            Width           =   1935
         End
         Begin VB.CommandButton btnGetXML 
            Caption         =   "������ ��ȸ - XML"
            Height          =   410
            Left            =   1920
            TabIndex        =   36
            Top             =   1200
            Width           =   1980
         End
         Begin VB.CommandButton btnGetTaxinvoice 
            Caption         =   "������ ��ȸ"
            Height          =   410
            Left            =   120
            TabIndex        =   35
            Top             =   1200
            Width           =   1735
         End
         Begin VB.TextBox txtNtsconfirmNum 
            Height          =   300
            Left            =   1560
            TabIndex        =   34
            Top             =   300
            Width           =   2295
         End
         Begin VB.Label Label6 
            Caption         =   "('������� ��ȸ'�� ��ȯ�� ����(����)��꼭 ����û���ι�ȣ�� �Է��Ͻñ� �ٶ��ϴ�.)"
            Height          =   375
            Left            =   120
            TabIndex        =   40
            Top             =   720
            Width           =   3735
         End
         Begin VB.Label Label5 
            Caption         =   "����û���ι�ȣ : "
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.ListBox taxinvoiceInfo 
         Height          =   4020
         Left            =   240
         TabIndex        =   30
         Top             =   3360
         Width           =   16695
      End
      Begin VB.Frame Frame8 
         Caption         =   "����/���� ������� ��ȸ"
         Height          =   2415
         Left            =   2280
         TabIndex        =   26
         Top             =   360
         Width           =   2415
         Begin VB.CommandButton btnSummary 
            Caption         =   "���� ��� ������� ��ȸ"
            Height          =   410
            Left            =   120
            TabIndex        =   28
            Top             =   840
            Width           =   2175
         End
         Begin VB.CommandButton btnSearch 
            Caption         =   "���� ��� ��ȸ"
            Height          =   410
            Left            =   120
            TabIndex        =   27
            Top             =   300
            Width           =   2175
         End
      End
      Begin VB.TextBox txtJobID 
         Height          =   300
         Left            =   2000
         TabIndex        =   25
         Top             =   2925
         Width           =   2175
      End
      Begin VB.Frame Frame7 
         Caption         =   "����/���� ���� ����"
         Height          =   2385
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   2055
         Begin VB.CommandButton btnRequestJob 
            Caption         =   "���� ��û"
            Height          =   375
            Left            =   120
            TabIndex        =   29
            Top             =   320
            Width           =   1815
         End
         Begin VB.CommandButton btnGetJobState 
            Caption         =   "���� ���� Ȯ��"
            Height          =   410
            Left            =   120
            TabIndex        =   23
            Top             =   800
            Width           =   1815
         End
         Begin VB.CommandButton btnListActiveJob 
            Caption         =   "���� ���� ��� Ȯ��"
            Height          =   410
            Left            =   120
            TabIndex        =   22
            Top             =   1300
            Width           =   1815
         End
      End
      Begin VB.Label Label4 
         Caption         =   "(�۾����̵�� '���� ��û' ȣ��� �����˴ϴ� )"
         Height          =   255
         Left            =   4305
         TabIndex        =   31
         Top             =   2985
         Width           =   4095
      End
      Begin VB.Label Label3 
         Caption         =   "�۾����̵� (jobID) :"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   3000
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
      Top             =   1920
      Width           =   1575
   End
   Begin VB.CommandButton btnListContact 
      Caption         =   "����� ��� ��ȸ"
      Height          =   410
      Left            =   9960
      TabIndex        =   7
      Top             =   1920
      Width           =   2055
   End
   Begin VB.CommandButton btnUpdateContact 
      Caption         =   "����� ���� ����"
      Height          =   410
      Left            =   9960
      TabIndex        =   8
      Top             =   2400
      Width           =   2055
   End
   Begin VB.CommandButton btnUpdateCorpInfo 
      Caption         =   "ȸ������ ����"
      Height          =   410
      Left            =   14880
      TabIndex        =   12
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Frame Frame15 
      Caption         =   "ȸ������ ����"
      Height          =   1935
      Left            =   14760
      TabIndex        =   11
      Top             =   1080
      Width           =   2240
      Begin VB.CommandButton btnGetCorpInfo 
         Caption         =   "ȸ������ ��ȸ"
         Height          =   410
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " �˺� �⺻ API "
      Height          =   2535
      Left            =   240
      TabIndex        =   13
      Top             =   720
      Width           =   17175
      Begin VB.Frame Frame12 
         Caption         =   "��Ʈ�ʰ��� ����Ʈ"
         Height          =   1935
         Left            =   6960
         TabIndex        =   42
         Top             =   360
         Width           =   2535
         Begin VB.CommandButton btnGetPartnerURL_CHRG 
            Caption         =   "����Ʈ ���� URL"
            Height          =   410
            Left            =   120
            TabIndex        =   46
            Top             =   840
            Width           =   2295
         End
         Begin VB.CommandButton btnGetPartnerBalance 
            Caption         =   "��Ʈ�� �ܿ�����Ʈ Ȯ��"
            Height          =   410
            Left            =   120
            TabIndex        =   44
            Top             =   360
            Width           =   2295
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "�������� ����Ʈ"
         Height          =   1935
         Left            =   4560
         TabIndex        =   41
         Top             =   360
         Width           =   2295
         Begin VB.CommandButton btnGetChargeURL 
            Caption         =   " ����Ʈ ���� URL"
            Height          =   410
            Left            =   120
            TabIndex        =   45
            Top             =   840
            Width           =   2055
         End
         Begin VB.CommandButton btnGetBalance 
            Caption         =   "�ܿ�����Ʈ Ȯ��"
            Height          =   410
            Left            =   120
            TabIndex        =   43
            Top             =   360
            Width           =   2055
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   " �˺� �⺻ URL"
         Height          =   1935
         Left            =   12000
         TabIndex        =   17
         Top             =   360
         Width           =   2295
         Begin VB.CommandButton btnGetAccessURL 
            Caption         =   " �˺� �α��� URL"
            Height          =   410
            Left            =   120
            TabIndex        =   9
            Top             =   360
            Width           =   2055
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "����� ����"
         Height          =   1935
         Left            =   9600
         TabIndex        =   16
         Top             =   360
         Width           =   2295
         Begin VB.CommandButton btnRegistContact 
            Caption         =   "����� �߰�"
            Height          =   410
            Left            =   120
            TabIndex        =   6
            Top             =   360
            Width           =   2055
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   " ����Ʈ ����"
         Height          =   1935
         Left            =   2160
         TabIndex        =   15
         Top             =   360
         Width           =   2175
         Begin VB.CommandButton btnGetChargeInfo 
            Caption         =   "�������� Ȯ��"
            Height          =   410
            Left            =   120
            TabIndex        =   5
            Top             =   360
            Width           =   1935
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   " ȸ������"
         Height          =   1935
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   1935
         Begin VB.CommandButton btnJoinMember 
            Caption         =   "ȸ�� ����"
            Height          =   410
            Left            =   120
            TabIndex        =   4
            Top             =   1320
            Width           =   1575
         End
         Begin VB.CommandButton btnCheckIsMember 
            Caption         =   "���� ���� Ȯ��"
            Height          =   410
            Left            =   120
            TabIndex        =   2
            Top             =   360
            Width           =   1575
         End
      End
   End
   Begin VB.Label Label1 
      Caption         =   "�˺�ȸ�� ����ڹ�ȣ :"
      Height          =   180
      Left            =   360
      TabIndex        =   19
      Top             =   315
      Width           =   1920
   End
   Begin VB.Label Label2 
      Caption         =   "�˺�ȸ�� ���̵� : "
      Height          =   180
      Left            =   4560
      TabIndex        =   18
      Top             =   315
      Width           =   1500
   End
End
Attribute VB_Name = "frmExample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=========================================================================
'
' �˺� Ȩ�ý� ���ڼ��ݰ�꼭 ���Ը��� ��ȸ API VB 6.0 SDK Example
'
' - VB6 SDK ����ȯ�� ������� �ȳ� : http://blog.linkhub.co.kr/569/
' - ������Ʈ ���� : 2019-02-11
' - ���� ������� ����ó : 1600-8536 / 070-4304-2991
' - ���� ������� �̸��� : code@linkhub.co.kr
'
' <�׽�Ʈ �������� �غ����>
' 1) 30, 33�� ���ο� ����� ��ũ���̵�(LinkID)�� ���Ű(SecretKey)��
'    ��ũ��� ���Խ� ���Ϸ� �߱޹��� ���������� �����Ͽ� �����մϴ�.
' 2) �˺� ���߿� ����Ʈ(test.popbill.com)�� ����ȸ������ �����մϴ�.
' 3) Ȩ�ý� �������񽺸� �̿��ϱ� ���� �˺��� ���������� ��� �մϴ�. (��������� �μ������ ���� / ���������� ���� ����� �ֽ��ϴ�.)
'    - �˺��α��� > [Ȩ�ý�����] > [ȯ�漳��] > [���� ����] �޴����� [Ȩ�ý� �μ������ ���] Ȥ��
'      [Ȩ�ý� ���������� ���]�� ���� ���������� ����մϴ�.
'    - Ȩ�ý����� ���� ���� �˾� URL(GetCertificatePopUpURL API) ��ȯ�� URL�� ���� �Ͽ�
'      [Ȩ�ý� �μ������ ���] Ȥ�� [Ȩ�ý� ���������� ���]�� ���� ���������� ����մϴ�.
'=========================================================================

Option Explicit

'=========================================================================
' - ��������(��ũ���̵�, ���Ű)�� ��Ʈ���� ����ȸ���� �ĺ��ϴ�
'   ������ ���Ǵ� ������ ������� �ʵ��� �����Ͻñ� �ٶ��ϴ�.
' - ����� ��ȯ���Ŀ��� ��������(��ũ���̵�, ���Ű)�� ������� �ʽ��ϴ�.
'=========================================================================

'��ũ���̵�
Private Const LinkID = "TESTER"

'���Ű. ���⿡ �����Ͻñ� �ٶ��ϴ�.
Private Const SecretKey = "SwWxqU+0TErBXy/9TVjIPEnI0VTUMMSQZtJf3Ed8q3I="

'Ȩ�ý� ���ڼ��ݰ�꼭 ���� Ŭ���� ����
Private htTaxinvoiceService As New PBHTTaxinvoiceService

'=========================================================================
' ��Ʈ���� ����ȸ������ ���Ե� ����ڹ�ȣ���� Ȯ���մϴ�.
' - LinkID�� ���������� �����Ǿ� �ִ� ��ũ���̵� ���Դϴ�.
'=========================================================================
Private Sub btnCheckIsMember_Click()
    Dim Response As PBResponse
    
    Set Response = htTaxinvoiceService.CheckIsMember(txtCorpNum.Text, LinkID)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(htTaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + htTaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.Message)
End Sub

'=========================================================================
' �˺� ȸ�����̵� �ߺ����θ� Ȯ���մϴ�.
'=========================================================================
Private Sub btnCheckID_Click()
    Dim Response As PBResponse
    
    Set Response = htTaxinvoiceService.CheckID(txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(htTaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + htTaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.Message)
End Sub

'=========================================================================
' �˺� ����ȸ�� ������ ��û�մϴ�.
'=========================================================================
Private Sub btnJoinMember_Click()
    Dim joinData As New PBJoinForm
    Dim Response As PBResponse
    
    '��ũ ���̵�
    joinData.LinkID = LinkID
    
    '����ڹ�ȣ, '-'����, 10�ڸ�
    joinData.CorpNum = "1234567890"
    
    '��ǥ�ڼ���, �ִ� 30��
    joinData.ceoname = "��ǥ�ڼ���"
    
    '��ȣ��, �ִ� 70��
    joinData.corpName = "ȸ����ȣ"
    
    '�ּ�, �ִ� 300��
    joinData.addr = "�ּ�"
    
    '����, �ִ� 40��
    joinData.bizType = "����"
    
    '����, �ִ� 40��
    joinData.bizClass = "����"
    
    '���̵�, 6���̻� 20�� �̸�
    joinData.id = "userid"
    
    '��й�ȣ, 6���̻� 20�� �̸�
    joinData.pwd = "pwd_must_be_long_enough"
    
    '����ڸ�, �ִ� 30��
    joinData.ContactName = "����ڼ���"
    
    '����� ����ó, �ִ� 20��
    joinData.ContactTEL = "02-999-9999"
    
    '����� �޴�����ȣ, �ִ� 20��
    joinData.ContactHP = "010-1234-5678"
    
    '����� �ѽ���ȣ, �ִ� 20��
    joinData.ContactFAX = "02-999-9998"
    
    '����� ����, �ִ� 70��
    joinData.ContactEmail = "test@test.com"
    
    Set Response = htTaxinvoiceService.JoinMember(joinData)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(htTaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + htTaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.Message)
End Sub

'=========================================================================
' ����ȸ���� Ȩ�ý� ���ڼ��ݰ�꼭 ���� API ���� ���������� Ȯ���մϴ�.
'=========================================================================
Private Sub btnGetChargeInfo_Click()
    Dim ChargeInfo As PBChargeInfo
    Dim tmp As String
    
    Set ChargeInfo = htTaxinvoiceService.GetChargeInfo(txtCorpNum.Text)
     
    If ChargeInfo Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(htTaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + htTaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = tmp + "unitCost (�����״ܰ�) : " + ChargeInfo.unitCost + vbCrLf
    tmp = tmp + "chargeMethod (��������) : " + ChargeInfo.chargeMethod + vbCrLf
    tmp = tmp + "rateSystem (��������) : " + ChargeInfo.rateSystem + vbCrLf
    
    MsgBox tmp
End Sub

'=========================================================================
' �˺��� �α��ε� �˺� URL�� ��ȯ�մϴ�.
' - ������å�� ���� ��ȯ�� URL�� 30���� ��ȿ�ð��� �����ϴ�.
'=========================================================================
Private Sub btnGetAccessURL_Click()
    Dim url As String
           
    url = htTaxinvoiceService.GetAccessURL(txtCorpNum.Text, txtUserID.Text)
    
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(htTaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + htTaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' ����ȸ���� ����ڸ� �űԷ� ����մϴ�.
'=========================================================================
Private Sub btnRegistContact_Click()
    Dim joinData As New PBContactInfo
    Dim Response As PBResponse
    
    '����� ���̵�, 6�� �̻� 50�� �̸�
    joinData.id = "testkorea"
    
    '��й�ȣ, 6�� �̻� 20�� �̸�
    joinData.pwd = "test@test.com"
    
    '����ڸ�, �ִ� 100��
    joinData.personName = "����ڸ�"
    
    '����� ����ó, �ִ� 20��
    joinData.tel = "070-1234-1234"
    
    '����� �޴�����ȣ, �ִ� 20��
    joinData.hp = "010-1234-1234"
    
    '����� �ѽ���,�ִ� 20��
    joinData.fax = "070-1234-1234"
    
    '����� �����ּ�, �ִ� 100��
    joinData.email = "test@test.com"
    
    'ȸ����ȸ ���ѿ���, True-ȸ����ȸ / False-������ȸ
    joinData.searchAllAllowYN = True
    
    '������ ����, True-������ / False-�����
    joinData.mgrYN = False
        
    Set Response = htTaxinvoiceService.RegistContact(txtCorpNum.Text, joinData)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(htTaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + htTaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.Message)
End Sub

'=========================================================================
' ����ȸ���� ����� ����� Ȯ���մϴ�.
'=========================================================================
Private Sub btnListContact_Click()
    Dim resultList As Collection
    Dim tmp As String
    Dim info As PBContactInfo
    
    Set resultList = htTaxinvoiceService.ListContact(txtCorpNum.Text)
     
    If resultList Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(htTaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + htTaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = "id(���̵�) | personName(����) | email(�̸���) | hp(�޴�����ȣ) |  fax(�ѽ���ȣ) | tel(����ó) | " _
         + "regDT(����Ͻ�) | searchAllAllowYN(ȸ����ȸ ���ѿ���) | mgrYN(������ ����) | state(����) " + vbCrLf
    
    For Each info In resultList
        tmp = tmp + info.id + " | " + info.personName + " | " + info.email + " | " + info.hp + " | " + info.fax _
        + info.tel + " | " + info.regDT + " | " + CStr(info.searchAllAllowYN) + " | " + CStr(info.mgrYN) + " | " + CStr(info.state) + vbCrLf
    Next
    
    MsgBox tmp
End Sub

'=========================================================================
' ����ȸ���� ����� ������ �����մϴ�.
'=========================================================================
Private Sub btnUpdateContact_Click()
    Dim joinData As New PBContactInfo
    Dim Response As PBResponse
    
    '����� ���̵�
    joinData.id = txtUserID.Text
    
    '����� ����, �ִ� 100��
    joinData.personName = "����ڸ�_����"
    
    '����� ����ó, �ִ� 20��
    joinData.tel = "070-1234-1234"
    
    '����� �޴�����ȣ, �ִ� 20��
    joinData.hp = "010-1234-1234"
        
    '����� �ѽ���ȣ, �ִ� 20��
    joinData.fax = "070-1234-1234"
    
    '����� �̸���, �ִ� 100��
    joinData.email = "test@test.com"

    'ȸ����ȸ ���ѿ���, True-ȸ����ȸ / False-������ȸ
    joinData.searchAllAllowYN = True
    
    '������ ����, True-������ / False-�����
    joinData.mgrYN = False
                
    Set Response = htTaxinvoiceService.UpdateContact(txtCorpNum.Text, joinData, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(htTaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + htTaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.Message)
End Sub

'=========================================================================
' ����ȸ���� ȸ�������� Ȯ���մϴ�.
'=========================================================================
Private Sub btnGetCorpInfo_Click()
    Dim CorpInfo As PBCorpInfo
    Dim tmp As String
    
    Set CorpInfo = htTaxinvoiceService.GetCorpInfo(txtCorpNum.Text)
     
    If CorpInfo Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(htTaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + htTaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = tmp + "ceoname (��ǥ�ڸ�) : " + CorpInfo.ceoname + vbCrLf
    tmp = tmp + "corpName (��ȣ) : " + CorpInfo.corpName + vbCrLf
    tmp = tmp + "addr (�ּ�) : " + CorpInfo.addr + vbCrLf
    tmp = tmp + "bizType (����) : " + CorpInfo.bizType + vbCrLf
    tmp = tmp + "bizClass (����) : " + CorpInfo.bizClass + vbCrLf
    
    MsgBox tmp
End Sub

'=========================================================================
' ����ȸ���� ȸ�������� �����մϴ�
'=========================================================================
Private Sub btnUpdateCorpInfo_Click()
    Dim CorpInfo As New PBCorpInfo
    Dim Response As PBResponse
    
    '��ǥ�ڸ�, �ִ� 100��
    CorpInfo.ceoname = "��ǥ��"
    
    '��ȣ, �ִ� 200��
    CorpInfo.corpName = "��ȣ"
    
    '�ּ�, �ִ� 300��
    CorpInfo.addr = "����Ư����"
    
    '����, �ִ� 100��
    CorpInfo.bizType = "����"
    
    '����, �ִ� 100��
    CorpInfo.bizClass = "����"
    
    Set Response = htTaxinvoiceService.UpdateCorpInfo(txtCorpNum.Text, CorpInfo)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(htTaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + htTaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.Message)
End Sub

'=========================================================================
' ����ȸ���� �ܿ�����Ʈ�� Ȯ���մϴ�.
' - ���ݹ���� ��Ʈ�ʰ����� ��� ��Ʈ�� �ܿ�����Ʈ(GetPartnerBalance API)
'   �� ���� Ȯ���Ͻñ� �ٶ��ϴ�.
'=========================================================================
Private Sub btnGetBalance_Click()
    Dim balance As Double
    
    balance = htTaxinvoiceService.GetBalance(txtCorpNum.Text)
    
    If balance < 0 Then
        MsgBox ("�����ڵ� : " + CStr(htTaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + htTaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "����ȸ�� �ܿ�����Ʈ : " + CStr(balance)
End Sub

'=========================================================================
' ����ȸ�� ����Ʈ ���� URL�� ��ȯ�մϴ�.
' - URL ������å�� ���� ��ȯ�� URL�� 30���� ��ȿ�ð��� �����ϴ�.
'=========================================================================
Private Sub btnGetChargeURL_Click()
    Dim url As String
           
    url = htTaxinvoiceService.GetChargeURL(txtCorpNum.Text, txtUserID.Text)
    
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(htTaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + htTaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' ��Ʈ���� �ܿ�����Ʈ�� Ȯ���մϴ�.
' - ���ݹ���� ���������� ��� ����ȸ�� �ܿ�����Ʈ(GetBalance API)
'   �� ���� Ȯ���Ͻñ� �ٶ��ϴ�.
'=========================================================================
Private Sub btnGetPartnerBalance_Click()
    Dim balance As Double
    
    balance = htTaxinvoiceService.GetPartnerBalance(txtCorpNum.Text)
    
    If balance < 0 Then
        MsgBox ("�����ڵ� : " + CStr(htTaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + htTaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "��Ʈ�� �ܿ�����Ʈ : " + CStr(balance)
End Sub

'=========================================================================
' ��Ʈ�� ����Ʈ ���� URL�� ��ȯ�մϴ�.
' - URL ������å�� ���� ��ȯ�� URL�� 30���� ��ȿ�ð��� �����ϴ�.
'=========================================================================
Private Sub btnGetPartnerURL_CHRG_Click()
    Dim url As String
           
    url = htTaxinvoiceService.GetPartnerURL(txtCorpNum.Text, "CHRG")
    
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(htTaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + htTaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
'����(����)��꼭 ����/���� ���� ������ ��û�մϴ�
'- Ȩ�ý����� ���μ����� "[Ȩ�ý�����(���ڼ��ݰ�꼭) API �����Ŵ���] >
'  1.1. Ȩ�ý�����(���ڼ��ݰ�꼭) API ����" �� �����Ͻñ� �ٶ��ϴ�.
'- ���� ��û�� ��ȯ���� �۾����̵�(JobID)�� ��ȿ�ð��� 1�ð� �Դϴ�.
'=========================================================================
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
    SDate = "20180601"
    
    '��������, ǥ������(yyyyMMdd)
    EDate = "20190201"
    
    jobID = htTaxinvoiceService.RequestJob(txtCorpNum.Text, tiType, DType, SDate, EDate)
    
    If jobID = "" Then
        MsgBox ("�����ڵ� : " + CStr(htTaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + htTaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "jobID(�۾����̵�) : " + jobID + vbCrLf
    
    txtJobID.Text = jobID
End Sub

'=========================================================================
'���� ��û ���¸� Ȯ���մϴ�.
'- �����׸� ���� ������ "[Ȩ�ý����� (���ڼ��ݰ�꼭) API �����Ŵ���] >
'  3.1.2. GetJobState(���� ���� Ȯ��)" �� �����Ͻñ� �ٶ��ϴ�.
'=========================================================================
Private Sub btnGetJobState_Click()
    Dim jobInfo As PBHTTaxinvoiceJobState
    Dim tmp As String
    
    Set jobInfo = htTaxinvoiceService.GetJobState(txtCorpNum.Text, txtJobID.Text)
     
    If jobInfo Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(htTaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + htTaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
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

'=========================================================================
' ���� ��û�ǵ鿡 ���� ���� ����� Ȯ���մϴ�.
' - ���� ��û �۾����̵�(JobID)�� ��ȿ�ð��� 1�ð� �Դϴ�.
' - �����׸� ���� ������ "[Ȩ�ý����� (���ڼ��ݰ�꼭��꼭) API �����Ŵ���] >
'   3.1.3. ListActiveJob(���� ���� ��� Ȯ��)" �� �����Ͻñ� �ٶ��ϴ�.
'=========================================================================
Private Sub btnListActiveJob_Click()
    Dim jobList As Collection
    Dim tmp As String
    Dim info As PBHTTaxinvoiceJobState
    
    Set jobList = htTaxinvoiceService.ListActiveJob(txtCorpNum.Text)
     
    If jobList Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(htTaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + htTaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = tmp + "�۾����̵�(jobID)�� ��ȿ�ð��� 1�ð��Դϴ�" + vbCrLf + vbCrLf
    tmp = tmp + "jobID(�۾����̵�) | jobState(��������) | queryType(��������) | queryDateType(��������) | queryStDate(��������) | queryEnDate(��������) |" _
            + "errorCode(�����ڵ�) | errorReason(�����޽���) | jobStartDT(�۾� �����Ͻ�) | jobEndDT(�۾� �����Ͻ�) | collectCount(��������) | regDT(���� ��û�Ͻ�) " + vbCrLf
    
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

'=========================================================================
' ���ڼ��ݰ�꼭 ����/���� ������ ���� ����� ��ȸ�մϴ�.
' - �����׸� ���� ������ "[Ȩ�ý����� (���ڼ��ݰ�꼭��꼭) API �����Ŵ���] >
'   3.2.1. Search(���� ��� ��ȸ)" �� �����Ͻñ� �ٶ��ϴ�.
'=========================================================================
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
    
    '��������ȣ, �⺻�� ��1��
    page = 1
    
    '�������� �˻�����, �⺻�� 500, �ִ� 1000
    perPage = 20
    
    '���� ����, D-��������, A-��������
    order = "D"
        
    Set SearchList = htTaxinvoiceService.Search(txtCorpNum.Text, txtJobID.Text, tiType, taxType, purposeType, taxRegIDType, taxRegID, taxRegIDYN, page, perPage, order)
    
        
    If SearchList Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(htTaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + htTaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = "code (�����ڵ�) : " + CStr(SearchList.code) + vbCrLf
    tmp = tmp + "message (����޽���) : " + SearchList.Message + vbCrLf
    tmp = tmp + "total (�� �˻���� �Ǽ�) : " + CStr(SearchList.total) + vbCrLf
    tmp = tmp + "perPage (�������� �˻�����) : " + CStr(SearchList.perPage) + vbCrLf
    tmp = tmp + "pageNum (������ ��ȣ) : " + CStr(SearchList.pageNum) + vbCrLf
    tmp = tmp + "pageCount (������ ����) : " + CStr(SearchList.pageCount) + vbCrLf + vbCrLf
    
    taxinvoiceInfo.Clear
        
    taxinvoiceInfo.AddItem "ntsconfirmNum (����û���ι�ȣ) | writeDate (�ۼ�����) | issueDate (��������) | sendDate (��������) | taxType (��������) ", 0
    taxinvoiceInfo.AddItem "purposeType (����/û��) | supplyCostTotal (���ް��� �հ�) | taxTotal (���� �հ�) | totalAmount (�հ�ݾ�) ", 1
    taxinvoiceInfo.AddItem "remark1 (���) | invoiceType (����/����) | modifyYN (���� ���ڼ��ݰ�꼭 ����) | orgNTSConfirmNum (���� ���ڼ��ݰ�꼭 ����û���ι�ȣ) ", 2
    taxinvoiceInfo.AddItem "purchaseDate (�ŷ�����) | itemName (ǰ��) | spec (�԰�) | qty (����) | unitCost (�ܰ�) | supplyCost (���ް���) ", 3
    taxinvoiceInfo.AddItem "tax (����) | remark (���) | invoicerCorpNum (������ ����ڹ�ȣ) | invoicerTaxRegID (������ ��������ȣ) | invoicerCorpName (������ ��ȣ) ", 4
    taxinvoiceInfo.AddItem "invoicerCEOName (������ ��ǥ�ڼ���) | invoicerEmail (������ ����� �̸���) | invoiceeCorpNum (���޹޴��� ����ڹ�ȣ) ", 5
    taxinvoiceInfo.AddItem "invoiceeType (���޹޴��� ����) | invoiceeTaxRegID (���޹޴��� ��������ȣ) | invoiceeCorpName (���޹޴��� ��ȣ) ", 6
    taxinvoiceInfo.AddItem "invoiceeCEOName (���޹޴��� ��ǥ�� ����) | invoiceeEmail1 (���޹޴��� ����� �̸���) | invoiceeEmail2 (���޹޴��� ASP �������� �̸���)"

    Dim tiInfo As PBHTTaxinvoiceAbbr
           
    For Each tiInfo In SearchList.list
        listboxRow = ""
        listboxRow = tiInfo.ntsconfirmNum + " | "
        listboxRow = listboxRow + tiInfo.writeDate + " | "
        listboxRow = listboxRow + tiInfo.issueDate + " | "
        listboxRow = listboxRow + tiInfo.sendDate + " | "
        listboxRow = listboxRow + tiInfo.taxType + " | "
        listboxRow = listboxRow + tiInfo.purposeType + " | "
        listboxRow = listboxRow + tiInfo.supplyCostTotal + " | "
        listboxRow = listboxRow + tiInfo.taxTotal + " | "
        listboxRow = listboxRow + tiInfo.totalAmount + " | "
        listboxRow = listboxRow + tiInfo.remark1 + " | "
        listboxRow = listboxRow + tiInfo.invoiceType + " | "
        
        If tiInfo.modifyYN Then
            listboxRow = listboxRow + "���� | "
        Else
            listboxRow = listboxRow + "�Ϲ� | "
        End If
        
        listboxRow = listboxRow + tiInfo.orgNTSConfirmNum + " | "
        listboxRow = listboxRow + tiInfo.purchaseDate + " | "
        listboxRow = listboxRow + tiInfo.itemName + " | "
        listboxRow = listboxRow + tiInfo.spec + " | "
        listboxRow = listboxRow + tiInfo.qty + " | "
        listboxRow = listboxRow + tiInfo.unitCost + " | "
        listboxRow = listboxRow + tiInfo.supplyCost + " | "
        listboxRow = listboxRow + tiInfo.tax + " | "
        listboxRow = listboxRow + tiInfo.remark + " | "
        listboxRow = listboxRow + tiInfo.invoicerCorpNum + " | "
        listboxRow = listboxRow + tiInfo.invoicerTaxRegID + " | "
        listboxRow = listboxRow + tiInfo.invoicerCorpName + " | "
        listboxRow = listboxRow + tiInfo.invoicerCEOName + " | "
        listboxRow = listboxRow + tiInfo.invoicerEmail + " | "
        listboxRow = listboxRow + tiInfo.invoiceeCorpNum + " | "
        listboxRow = listboxRow + tiInfo.invoiceeType + " | "
        listboxRow = listboxRow + tiInfo.invoiceeTaxRegID + " | "
        listboxRow = listboxRow + tiInfo.invoiceeCorpName + " | "
        listboxRow = listboxRow + tiInfo.invoiceeCEOName + " | "
        listboxRow = listboxRow + tiInfo.invoiceeEmail1 + " | "
        listboxRow = listboxRow + tiInfo.invoiceeEmail2
        
        taxinvoiceInfo.AddItem listboxRow, taxinvoiceInfo.ListCount
    Next
  
    MsgBox (tmp)
End Sub

'=========================================================================
' ���ڼ��ݰ�꼭 ����/���� ������ ���� ��� ��������� ��ȸ�մϴ�.
' - �����׸� ���� ������ "[Ȩ�ý����� (���ڼ��ݰ�꼭��꼭) API �����Ŵ���] >
'   3.2.2. Summary(���� ��� ������� ��ȸ)" �� �����Ͻñ� �ٶ��ϴ�.
'=========================================================================
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
        
    
    Set summaryInfo = htTaxinvoiceService.Summary(txtCorpNum.Text, txtJobID.Text, tiType, taxType, purposeType, taxRegIDType, taxRegID, taxRegIDYN)
        
    If summaryInfo Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(htTaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + htTaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = "count (������� �Ǽ�) : " + CStr(summaryInfo.count) + vbCrLf
    tmp = tmp + "supplyCostTotal (���ް��� �հ�) : " + CStr(summaryInfo.supplyCostTotal) + vbCrLf
    tmp = tmp + "taxTotal (���� �հ�) : " + CStr(summaryInfo.taxTotal) + vbCrLf
    tmp = tmp + "amountTotal (�հ� �ݾ�) : " + CStr(summaryInfo.amountTotal) + vbCrLf
       
    MsgBox (tmp)
End Sub

'=========================================================================
' ���ڼ��ݰ�꼭 1���� �������� Ȯ���մϴ�.
' - �����׸� ���� ������ "[Ȩ�ý����� (���ڼ��ݰ�꼭��꼭) API �����Ŵ���] >
'   4.1.2. GetTaxinvoice �������� ����" �� �����Ͻñ� �ٶ��ϴ�.
'=========================================================================
Private Sub btnGetTaxinvoice_Click()
    Dim taxinvoiceInfo As PBHTTaxinvoice
    Dim i As Integer
    Dim tmp As String
    
    Set taxinvoiceInfo = htTaxinvoiceService.GetTaxinvoice(txtCorpNum.Text, txtNtsconfirmNum.Text)
     
    If taxinvoiceInfo Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(htTaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + htTaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = "========����(����)��꼭 ����=======" + vbCrLf
    tmp = tmp + "writeDate(�ۼ�����) : " + taxinvoiceInfo.writeDate + vbCrLf
    tmp = tmp + "issueDT(�����Ͻ�) : " + taxinvoiceInfo.issueDT + vbCrLf
    tmp = tmp + "invoiceType(���ڼ��ݰ�꼭 ����) : " + taxinvoiceInfo.invoiceType + vbCrLf
    tmp = tmp + "taxType(��������) : " + taxinvoiceInfo.taxType + vbCrLf
    tmp = tmp + "taxTotal(�����հ�) : " + taxinvoiceInfo.taxTotal + vbCrLf
    tmp = tmp + "supplyCostTotal(���ް��� �հ�) : " + taxinvoiceInfo.supplyCostTotal + vbCrLf
    tmp = tmp + "totalAmount(�հ�ݾ�) : " + taxinvoiceInfo.totalAmount + vbCrLf
    tmp = tmp + "purposeType(����/û��) : " + taxinvoiceInfo.purposeType + vbCrLf
    tmp = tmp + "cash(����) : " + taxinvoiceInfo.cash + vbCrLf
    tmp = tmp + "chkBill(��ǥ) : " + taxinvoiceInfo.chkBill + vbCrLf
    tmp = tmp + "credit(�ܻ�) : " + taxinvoiceInfo.credit + vbCrLf
    tmp = tmp + "note(����) : " + taxinvoiceInfo.note + vbCrLf
    tmp = tmp + "remark1(���1) : " + taxinvoiceInfo.remark1 + vbCrLf
    tmp = tmp + "remark2(���2) : " + taxinvoiceInfo.remark2 + vbCrLf
    tmp = tmp + "remark3(���3) : " + taxinvoiceInfo.remark3 + vbCrLf
    tmp = tmp + "ntsconfirmNum(����û ���ι�ȣ) : " + taxinvoiceInfo.ntsconfirmNum + vbCrLf + vbCrLf
    
    tmp = tmp + "========������ ����=======" + vbCrLf
    tmp = tmp + "invoicerCorpNum(����ڹ�ȣ) : " + taxinvoiceInfo.invoicerCorpNum + vbCrLf
    tmp = tmp + "invoicerMgtKey(������ȣ) : " + taxinvoiceInfo.invoicerMgtKey + vbCrLf
    tmp = tmp + "invoicerTaxRegID(������� ��ȣ) : " + taxinvoiceInfo.invoicerTaxRegID + vbCrLf
    tmp = tmp + "invoicerCorpName(��ȣ) : " + taxinvoiceInfo.invoicerCorpName + vbCrLf
    tmp = tmp + "invoicerCEOName(����) : " + taxinvoiceInfo.invoicerCEOName + vbCrLf
    tmp = tmp + "invoicerAddr(�ּ�) : " + taxinvoiceInfo.invoicerAddr + vbCrLf
    tmp = tmp + "invoicerBizType(����) : " + taxinvoiceInfo.invoicerBizType + vbCrLf
    tmp = tmp + "invoicerBizClass(����) : " + taxinvoiceInfo.invoicerBizClass + vbCrLf
    tmp = tmp + "invoicerContactName(����� ����) : " + taxinvoiceInfo.invoicerContactName + vbCrLf
    tmp = tmp + "invoicerTEL(����� ����ó) : " + taxinvoiceInfo.invoicerTEL + vbCrLf
    tmp = tmp + "invoicerEmail(����� �̸���) : " + taxinvoiceInfo.invoicerEmail + vbCrLf + vbCrLf
    
    tmp = tmp + "========���޹޴��� ����=======" + vbCrLf
    tmp = tmp + "invoiceeCorpNum(����ڹ�ȣ) : " + taxinvoiceInfo.invoiceeCorpNum + vbCrLf
    tmp = tmp + "invoiceeType(���޹޴��� ����) : " + taxinvoiceInfo.invoiceeType + vbCrLf
    tmp = tmp + "invoiceeMgtKey(������ȣ : " + taxinvoiceInfo.invoiceeMgtKey + vbCrLf
    tmp = tmp + "invoiceeTaxRegID(������� ��ȣ) : " + taxinvoiceInfo.invoiceeTaxRegID + vbCrLf
    tmp = tmp + "invoiceeCorpName(��ȣ) : " + taxinvoiceInfo.invoiceeCorpName + vbCrLf
    tmp = tmp + "invoiceeCEOName(����) : " + taxinvoiceInfo.invoiceeCEOName + vbCrLf
    tmp = tmp + "invoiceeAddr(�ּ�) : " + taxinvoiceInfo.invoiceeAddr + vbCrLf
    tmp = tmp + "invoiceeBizType(����) : " + taxinvoiceInfo.invoiceeBizType + vbCrLf
    tmp = tmp + "invoiceeBizClass(����) : " + taxinvoiceInfo.invoiceeBizClass + vbCrLf
    tmp = tmp + "invoiceeContactName1(��)����� ����) : " + taxinvoiceInfo.invoiceeContactName1 + vbCrLf
    tmp = tmp + "invoiceeTEL1(��)����� ����ó) : " + taxinvoiceInfo.invoiceeTEL1 + vbCrLf
    tmp = tmp + "invoiceeEmail1(��)����� �̸���) : " + taxinvoiceInfo.invoiceeEmail1 + vbCrLf
        
    tmp = tmp + "========����(����)��꼭 ǰ��迭========" + vbCrLf
    Dim detailInfo As PBHTTaxinvoiceDetail
    
    For Each detailInfo In taxinvoiceInfo.detailList
        tmp = tmp + "serialNum(�Ϸù�ȣ) : " + CStr(detailInfo.serialNum) + vbCrLf
        tmp = tmp + "purchaseDT(�ŷ�����) : " + detailInfo.purchaseDT + vbCrLf
        tmp = tmp + "itemName(ǰ���) : " + detailInfo.itemName + vbCrLf
        tmp = tmp + "spec(�԰�) : " + detailInfo.spec + vbCrLf
        tmp = tmp + "qty(����) : " + detailInfo.qty + vbCrLf
        tmp = tmp + "unitCost(�ܰ�) : " + detailInfo.unitCost + vbCrLf
        tmp = tmp + "supplyCost(���ް���) : " + detailInfo.supplyCost + vbCrLf
        tmp = tmp + "tax(����) : " + detailInfo.tax + vbCrLf
        tmp = tmp + "remark(���) : " + detailInfo.remark + vbCrLf + vbCrLf
    Next
    
    MsgBox (tmp)
End Sub

'=========================================================================
' XML ������ ���ڼ��ݰ�꼭 �������� Ȯ���մϴ�.
' - �����׸� ���� ������ "[Ȩ�ý����� (���ڼ��ݰ�꼭��꼭) API �����Ŵ���] >
'   3.2.4. GetXML(������ Ȯ�� - XML)" �� �����Ͻñ� �ٶ��ϴ�.
'=========================================================================
Private Sub btnGetXML_Click()
    Dim taxinvoiceXML As PBHTTaxinvoiceXML
    Dim i As Integer
    Dim tmp As String
    
    Set taxinvoiceXML = htTaxinvoiceService.GetXML(txtCorpNum.Text, txtNtsconfirmNum.Text)
    
    If taxinvoiceXML Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(htTaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + htTaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = tmp + "ResultCode (��û�� ���� ���� �����ڵ�) : " + CStr(taxinvoiceXML.ResultCode) + vbCrLf
    tmp = tmp + "Message (����û���ι�ȣ) : " + taxinvoiceXML.Message + vbCrLf
    tmp = tmp + "retObject (XML����) : " + taxinvoiceXML.retObject + vbCrLf
    
    MsgBox (tmp)
End Sub

'=========================================================================
' Ȩ�ý� ���ڼ��ݰ�꼭 �˾� URL�� ��ȯ�մϴ�.
' - ��ȯ�� URL�� ������å���� ���� 30���� ��ȿ�ð��� �����ϴ�.
'=========================================================================
Private Sub btnGetPopUpURL_Click()
    Dim url As String
    
    url = htTaxinvoiceService.GetPopUpURL(txtCorpNum.Text, txtNtsconfirmNum.Text)
    
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(htTaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + htTaxinvoiceService.LastErrMessage)
        Exit Sub
    End If

    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' ������ ��û �˾� URL�� ��ȯ�մϴ�.
' - ������å�� ���� ��ȯ�� URL�� 30���� ��ȿ�ð��� �����ϴ�.
'=========================================================================
Private Sub btnGetFlatRatePopUpURL_Click()
    Dim url As String
    
    url = htTaxinvoiceService.GetFlatRatePopUpURL(txtCorpNum.Text)
    
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(htTaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + htTaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' ����ȸ���� ������ ���� �̿���¸� Ȯ���մϴ�.
'=========================================================================
Private Sub btnGetFlatRateState_Click()
    Dim flatRateInfo As PBHTTaxinvoiceFlatRate
    Dim tmp As String
    
    Set flatRateInfo = htTaxinvoiceService.GetFlatRateState(txtCorpNum.Text)
     
    If flatRateInfo Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(htTaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + htTaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
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

'=========================================================================
' Ȩ�ý����� ���������� ���� URL�� ��ȯ�մϴ�.
' ������Ŀ��� �μ������/���������� ���� ����� �ֽ��ϴ�.
' - ��ȯ�� URL�� ������å�� ���� 30���� ��ȿ�ð��� �����ϴ�.
'=========================================================================
Private Sub btnGetCertificatePopUpURL_Click()
    Dim url As String
    
    url = htTaxinvoiceService.GetCertificatePopUpURL(txtCorpNum.Text)
    
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(htTaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + htTaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' �˺��� ��ϵ� Ȩ�ý� ������������ �������ڸ� Ȯ���մϴ�.
'=========================================================================
Private Sub btnGetCertificateExpireDate_Click()
    Dim expireDate As String
    
    expireDate = htTaxinvoiceService.GetCertificateExpireDate(txtCorpNum.Text)
    
    If expireDate = "" Then
        MsgBox ("�����ڵ� : " + CStr(htTaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + htTaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "������������ : " + expireDate
End Sub

'=========================================================================
' �˺��� ��ϵ� ������������ Ȩ�ý� �α����� �׽�Ʈ�Ѵ�.
'=========================================================================
Private Sub btnCheckCertValidation_Click()
    Dim Response As PBResponse
    
    Set Response = htTaxinvoiceService.CheckCertValidation(txtCorpNum.Text)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(htTaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + htTaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.Message)
End Sub

'=========================================================================
' Ȩ�ý� ���ڼ��ݰ�꼭 �μ������ ������ ����մϴ�.
'=========================================================================
Private Sub btnRegistDeptUser_Click()
    Dim Response As PBResponse
    Dim DeptUserID As String
    Dim DeptUserPWD As String
    
    'Ȩ�ý����� ������ ���ڼ��ݰ�꼭 �μ������ ���̵�
    DeptUserID = "userid_test"
    
    'Ȩ�ý����� ������ ���ڼ��ݰ�꼭 �μ������ ��й�ȣ
    DeptUserPWD = "passwd_test"
    
    Set Response = htTaxinvoiceService.RegistDeptUser(txtCorpNum.Text, DeptUserID, DeptUserPWD)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(htTaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + htTaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.Message)
End Sub

'=========================================================================
' �˺��� ��ϵ� ���ڼ��ݰ�꼭 �μ������ ���̵� Ȯ���մϴ�.
'=========================================================================
Private Sub btnCheckDeptUser_Click()
    Dim Response As PBResponse
    
    Set Response = htTaxinvoiceService.CheckDeptUser(txtCorpNum.Text)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(htTaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + htTaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.Message)
End Sub

'=========================================================================
' �˺��� ��ϵ� ���ڼ��ݰ�꼭 �μ������ ���������� �̿��Ͽ� Ȩ�ý� �α����� �׽�Ʈ�մϴ�.
'=========================================================================
Private Sub btnCheckLoginDeptUser_Click()
    Dim Response As PBResponse
    
    Set Response = htTaxinvoiceService.CheckLoginDeptUser(txtCorpNum.Text)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(htTaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + htTaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.Message)
End Sub

'=========================================================================
'  �˺��� ��ϵ� ���ڼ��ݰ�꼭 �μ������ ���������� �����մϴ�.
'=========================================================================
Private Sub btnDeleteDeptUser_Click()
    Dim Response As PBResponse
    
    Set Response = htTaxinvoiceService.DeleteDeptUser(txtCorpNum.Text)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(htTaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + htTaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.Message)
End Sub

Private Sub Form_Load()
    '��� �ʱ�ȭ
    htTaxinvoiceService.Initialize LinkID, SecretKey
    
    '����ȯ�� ������ True(���߿�), False(�����)
    htTaxinvoiceService.IsTest = True
End Sub

