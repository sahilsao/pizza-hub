VERSION 5.00
Begin VB.Form frmNewCustomer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Customer"
   ClientHeight    =   3600
   ClientLeft      =   8130
   ClientTop       =   3840
   ClientWidth     =   5505
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   5505
   Begin VB.TextBox txtComment 
      DataField       =   "Comment"
      DataMember      =   "comCustomers"
      DataSource      =   "denPizza"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1260
      MaxLength       =   10
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   2655
      Width           =   4095
   End
   Begin VB.TextBox txtPhoneNumber 
      DataField       =   "PhoneNumber"
      DataMember      =   "comCustomers"
      DataSource      =   "denPizza"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1260
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   120
      Width           =   4095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2895
      TabIndex        =   13
      Top             =   3135
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   1335
      TabIndex        =   12
      Top             =   3135
      Width           =   1215
   End
   Begin VB.TextBox txtZip 
      DataField       =   "Zip"
      DataMember      =   "comCustomers"
      DataSource      =   "denPizza"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1260
      MaxLength       =   10
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   2280
      Width           =   4095
   End
   Begin VB.TextBox txtState 
      DataField       =   "State"
      DataMember      =   "comCustomers"
      DataSource      =   "denPizza"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1260
      MaxLength       =   20
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   1920
      Width           =   4095
   End
   Begin VB.TextBox txtCity 
      DataField       =   "City"
      DataMember      =   "comCustomers"
      DataSource      =   "denPizza"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1260
      MaxLength       =   40
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   1560
      Width           =   4095
   End
   Begin VB.TextBox txtAddress 
      DataField       =   "Address"
      DataMember      =   "comCustomers"
      DataSource      =   "denPizza"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1260
      MaxLength       =   40
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   1200
      Width           =   4095
   End
   Begin VB.TextBox txtLastName 
      DataField       =   "LastName"
      DataMember      =   "comCustomers"
      DataSource      =   "denPizza"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1260
      MaxLength       =   40
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   840
      Width           =   4095
   End
   Begin VB.TextBox txtFirstName 
      DataField       =   "FirstName"
      DataMember      =   "comCustomers"
      DataSource      =   "denPizza"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1260
      MaxLength       =   40
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   480
      Width           =   4095
   End
   Begin VB.Label Label2 
      Caption         =   "Comment:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   90
      TabIndex        =   17
      Top             =   2655
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Telephone:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   60
      TabIndex        =   15
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "First Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   60
      TabIndex        =   5
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   "Last Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   60
      TabIndex        =   4
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label9 
      Caption         =   "Address:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   60
      TabIndex        =   3
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label11 
      Caption         =   "City:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   60
      TabIndex        =   2
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label13 
      Caption         =   "State:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   60
      TabIndex        =   1
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label15 
      Caption         =   "Zip:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   60
      TabIndex        =   0
      Top             =   2250
      Width           =   1335
   End
End
Attribute VB_Name = "frmNewCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdcancel_Click()
'Cancel new customer addition and return to order form
denPizza.rscomCustomers.CancelUpdate
denPizza.rscomOrders.CancelUpdate
Unload frmNewCustomer
frmPizza.medPhone.Text = "(___) ___-____"
frmPizza.medPhone.SetFocus
End Sub

Private Sub cmdSave_Click()
'Save entered information and return to order form
Dim Desc As String, AllOK As Boolean
AllOK = True
If txtFirstName.Text = "" Then AllOK = False
If txtLastName.Text = "" Then AllOK = False
If txtAddress.Text = "" Then AllOK = False
If txtCity.Text = "" Then AllOK = False
If txtState.Text = "" Then AllOK = False
If txtZip.Text = "" Then AllOK = False
If Not (AllOK) Then
  MsgBox "All text boxes require an entry.", vbOKOnly + vbInformation, "Information Missing"
  txtFirstName.SetFocus
  Exit Sub
End If
'Rebind fields
denPizza.rscomCustomers.Update
Set txtFirstName.DataSource = denPizza
Set txtFirstName.DataSource = denPizza
Set txtLastName.DataSource = denPizza
Set txtAddress.DataSource = denPizza
Set txtCity.DataSource = denPizza
Set txtState.DataSource = denPizza
Set txtZip.DataSource = denPizza
Set txtFirstName.DataSource = denPizza
Set txtFirstName.DataSource = denPizza
Set txtLastName.DataSource = denPizza
Set txtAddress.DataSource = denPizza
Set txtCity.DataSource = denPizza
Set txtState.DataSource = denPizza
Set txtZip.DataSource = denPizza
Set txtComment.DataSource = denPizza
NewCustomer = True
Unload frmNewCustomer
End Sub

Private Sub Form_Activate()
denPizza.rscomCustomers.AddNew
txtPhoneNumber.Text = frmPizza.medPhone.Text
txtFirstName.SetFocus
End Sub

Private Sub txtAddress_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then txtCity.SetFocus
End Sub

Private Sub txtCity_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then txtState.SetFocus
End Sub

Private Sub txtFirstName_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then txtLastName.SetFocus
End Sub


Private Sub txtLastName_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then txtAddress.SetFocus
End Sub


Private Sub txtState_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then txtZip.SetFocus
End Sub

Private Sub txtZip_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then cmdSave.SetFocus
End Sub
