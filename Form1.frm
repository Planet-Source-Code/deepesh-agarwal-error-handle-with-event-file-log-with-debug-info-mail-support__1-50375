VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "VBCustomErrorTrapDemo"
   ClientHeight    =   900
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   900
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Generate Error For Test >>>>"
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=========================================================================================
'  VBErrorTrapDemo
'  How to Use:-
'  (1)Add Errfrm.frm, ModError.Bas, ModMSMail.bas _
'  Sysinfo.bas and ErrBitmap.res to your project.
'  (2)Add a refrence to Microsoft Scripting Runtime _
'  from Project --> references
'  (3)Add RichtextBox and winsock controls to your toolbox
'   from Project --> Components
'=========================================================================================
'  Coded By: Deepesh Agarwal
'  Published Date: 29/09/2003
'  WebSite: http://www.deepeshagarwal.tk
'  E-mail: agarwal_deepesh@indiatimes.com
'  Visit my site for Free-Software's like:
'  1). The-AdPolice - Blocks 17000+ adservers to save bandwidth
'  2). Dr. System -  Schedule Computer Maintainence - A must for every computer user
'  3). Service Controller XP (A Must For XP User) - Start,Stop,Pause and change startup type of 2000/XP services with recommended settings for different system config.
'   And Many More........
'=========================================================================================


Private Sub Command1_Click()
    On Error GoTo Oops_Error
    'generate an error for testing
    Err.Raise Number:=7
Oops_Error:
    'calling our error handler from the ModError module
    ModError.ErrorHandler
End Sub 'Command1_Click()


