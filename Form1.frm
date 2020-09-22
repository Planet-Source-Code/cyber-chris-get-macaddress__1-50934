VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Get the Macaddress"
   ClientHeight    =   255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2775
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   255
   ScaleWidth      =   2775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Label lblMac 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'(c) Copyright by CC
'   Email: cyber_chris235@gmx.net
'
'Please mail me, when you want to use my code!

Option Explicit
Private Const NCBASTAT                       As Long = &H33
Private Const NCBNAMSZ                       As Integer = 16
Private Const HEAP_ZERO_MEMORY               As Long = &H8
Private Const HEAP_GENERATE_EXCEPTIONS       As Long = &H4
Private Const NCBRESET                       As Long = &H32
Private Type NCB
    ncb_command                                As Byte
    ncb_retcode                                As Byte
    ncb_lsn                                    As Byte
    ncb_num                                    As Byte
    ncb_buffer                                 As Long
    ncb_length                                 As Integer
    ncb_callname                               As String * NCBNAMSZ
    ncb_name                                   As String * NCBNAMSZ
    ncb_rto                                    As Byte
    ncb_sto                                    As Byte
    ncb_post                                   As Long
    ncb_lana_num                               As Byte
    ncb_cmd_cplt                               As Byte
    ncb_reserve(9)                             As Byte
    ncb_event                                  As Long
End Type
Private Type ADAPTER_STATUS
    adapter_address(5)                         As Byte
    rev_major                                  As Byte
    reserved0                                  As Byte
    adapter_type                               As Byte
    rev_minor                                  As Byte
    duration                                   As Integer
    frmr_recv                                  As Integer
    frmr_xmit                                  As Integer
    iframe_recv_err                            As Integer
    xmit_aborts                                As Integer
    xmit_success                               As Long
    recv_success                               As Long
    iframe_xmit_err                            As Integer
    recv_buff_unavail                          As Integer
    t1_timeouts                                As Integer
    ti_timeouts                                As Integer
    Reserved1                                  As Long
    free_ncbs                                  As Integer
    max_cfg_ncbs                               As Integer
    max_ncbs                                   As Integer
    xmit_buf_unavail                           As Integer
    max_dgram_size                             As Integer
    pending_sess                               As Integer
    max_cfg_sess                               As Integer
    max_sess                                   As Integer
    max_sess_pkt_size                          As Integer
    name_count                                 As Integer
End Type
Private Type NAME_BUFFER
    name                                       As String * NCBNAMSZ
    name_num                                   As Integer
    name_flags                                 As Integer
End Type
Private Type ASTAT
    adapt                                      As ADAPTER_STATUS
    NameBuff(30)                               As NAME_BUFFER
End Type
Private Declare Function Netbios Lib "netapi32.dll" (pncb As NCB) As Byte
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, _
                                                                     ByVal hpvSource As Long, _
                                                                     ByVal cbCopy As Long)
Private Declare Function GetProcessHeap Lib "kernel32" () As Long
Private Declare Function HeapAlloc Lib "kernel32" (ByVal hHeap As Long, _
                                                   ByVal dwFlags As Long, _
                                                   ByVal dwBytes As Long) As Long
Private Declare Function HeapFree Lib "kernel32" (ByVal hHeap As Long, _
                                                  ByVal dwFlags As Long, _
                                                  lpMem As Any) As Long

Private Sub Form_Load()

  
  Dim bRet    As Byte
  Dim myNcb   As NCB
  Dim myASTAT As ASTAT
  Dim pASTAT  As Long
    myNcb.ncb_command = NCBRESET
    bRet = Netbios(myNcb)
    With myNcb
        .ncb_command = NCBASTAT
        .ncb_lana_num = 0
        .ncb_callname = "* "
        .ncb_length = Len(myASTAT)
        pASTAT = HeapAlloc(GetProcessHeap(), HEAP_GENERATE_EXCEPTIONS Or HEAP_ZERO_MEMORY, .ncb_length)
    End With
    If pASTAT = 0 Then
        Exit Sub
    End If
    myNcb.ncb_buffer = pASTAT
    bRet = Netbios(myNcb)
    CopyMemory myASTAT, myNcb.ncb_buffer, Len(myASTAT)
    lblMac.Caption = HexEx(myASTAT.adapt.adapter_address(0)) & "-" & HexEx(myASTAT.adapt.adapter_address(1)) & "-" & HexEx(myASTAT.adapt.adapter_address(2)) & "-" & HexEx(myASTAT.adapt.adapter_address(3)) & "-" & HexEx(myASTAT.adapt.adapter_address(4)) & "-" & HexEx(myASTAT.adapt.adapter_address(5))
    Call HeapFree(GetProcessHeap(), 0, pASTAT)

End Sub

Private Function HexEx(ByVal B As Long) As String
 
  Dim aa As String

    aa = Hex$(B)
    If Len(aa) < 2 Then
        aa = "0" & aa
    End If
    HexEx = aa

End Function

