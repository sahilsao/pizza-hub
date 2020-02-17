VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmPizza 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pizza Order"
   ClientHeight    =   4365
   ClientLeft      =   1140
   ClientTop       =   1515
   ClientWidth     =   9765
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4365
   ScaleWidth      =   9765
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSpecial 
      Enabled         =   0   'False
      Height          =   495
      Left            =   3420
      MaxLength       =   50
      MultiLine       =   -1  'True
      TabIndex        =   36
      Top             =   2910
      Width           =   5355
   End
   Begin VB.Frame fraDeliver 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3465
      TabIndex        =   34
      Top             =   1920
      Width           =   1455
      Begin VB.CheckBox chkDeliver 
         Caption         =   "Deliver"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   270
         TabIndex        =   35
         Top             =   180
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdComplete 
      Caption         =   "Order &Complete"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5445
      TabIndex        =   27
      Top             =   3495
      Width           =   1935
   End
   Begin VB.Frame Frame4 
      Height          =   2595
      Left            =   420
      TabIndex        =   16
      Top             =   1200
      Width           =   2955
      Begin VB.TextBox txtComment 
         DataField       =   "Comment"
         DataMember      =   "comCustomers"
         DataSource      =   "denPizza"
         Height          =   585
         Left            =   120
         MaxLength       =   50
         MultiLine       =   -1  'True
         TabIndex        =   26
         Top             =   1860
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.TextBox txtZip 
         DataField       =   "Zip"
         DataMember      =   "comCustomers"
         DataSource      =   "denPizza"
         Height          =   285
         Left            =   2040
         TabIndex        =   25
         Top             =   1560
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtState 
         DataField       =   "State"
         DataMember      =   "comCustomers"
         DataSource      =   "denPizza"
         Height          =   285
         Left            =   1560
         TabIndex        =   24
         Top             =   1560
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox txtCity 
         DataField       =   "City"
         DataMember      =   "comCustomers"
         DataSource      =   "denPizza"
         Height          =   285
         Left            =   120
         TabIndex        =   23
         Top             =   1560
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.TextBox txtAddress 
         DataField       =   "Address"
         DataMember      =   "comCustomers"
         DataSource      =   "denPizza"
         Height          =   285
         Left            =   120
         TabIndex        =   22
         Top             =   1260
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.TextBox txtLastName 
         DataField       =   "LastName"
         DataMember      =   "comCustomers"
         DataSource      =   "denPizza"
         Height          =   285
         Left            =   1080
         TabIndex        =   21
         Top             =   960
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox txtFirstName 
         DataField       =   "FirstName"
         DataMember      =   "comCustomers"
         DataSource      =   "denPizza"
         Height          =   285
         Left            =   120
         TabIndex        =   20
         Top             =   960
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "&Find Customer"
         Height          =   315
         Left            =   660
         TabIndex        =   19
         Top             =   600
         Width           =   1575
      End
      Begin MSMask.MaskEdBox medPhone 
         Bindings        =   "Pizza.frx":0000
         Height          =   375
         Left            =   1080
         TabIndex        =   17
         Top             =   180
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "(###) ###-####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label2 
         Caption         =   "Phone:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   315
         TabIndex        =   18
         Top             =   225
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   450
      TabIndex        =   13
      Top             =   525
      Width           =   2925
      Begin VB.Label lblDate 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "OrderDate"
         DataMember      =   "comOrders"
         DataSource      =   "denPizza"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1035
         TabIndex        =   14
         Top             =   180
         Width           =   1365
      End
      Begin VB.Label Label1 
         Caption         =   "Date:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   405
         TabIndex        =   15
         Top             =   225
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7425
      TabIndex        =   12
      Top             =   3495
      Width           =   1335
   End
   Begin VB.CommandButton cmdBuild 
      Caption         =   "&Build Pizza"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3465
      TabIndex        =   11
      Top             =   3495
      Width           =   1935
   End
   Begin VB.Frame fraToppings 
      Caption         =   "Toppings"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   5100
      TabIndex        =   4
      Top             =   600
      Width           =   3675
      Begin VB.CheckBox chkTop 
         Caption         =   "Pepperoni"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   1920
         TabIndex        =   33
         Top             =   300
         Width           =   1695
      End
      Begin VB.CheckBox chkTop 
         Caption         =   "Salami"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   1920
         TabIndex        =   32
         Top             =   660
         Width           =   1335
      End
      Begin VB.CheckBox chkTop 
         Caption         =   "Sausage"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   1920
         TabIndex        =   31
         Top             =   960
         Width           =   1695
      End
      Begin VB.CheckBox chkTop 
         Caption         =   "Ground Beef"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   9
         Left            =   1920
         TabIndex        =   30
         Top             =   1200
         Width           =   1515
      End
      Begin VB.CheckBox chkTop 
         Caption         =   "Shrimp"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   1920
         TabIndex        =   29
         Top             =   1560
         Width           =   1515
      End
      Begin VB.CheckBox chkTop 
         Caption         =   "Anchovies"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   1920
         TabIndex        =   28
         Top             =   1860
         Width           =   1335
      End
      Begin VB.CheckBox chkTop 
         Caption         =   "Tomatoes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   180
         TabIndex        =   10
         Top             =   1860
         Width           =   1335
      End
      Begin VB.CheckBox chkTop 
         Caption         =   "Green Peppers"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   180
         TabIndex        =   9
         Top             =   1560
         Width           =   1815
      End
      Begin VB.CheckBox chkTop 
         Caption         =   "Onions"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   180
         TabIndex        =   8
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CheckBox chkTop 
         Caption         =   "Black Olives"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   180
         TabIndex        =   7
         Top             =   960
         Width           =   1695
      End
      Begin VB.CheckBox chkTop 
         Caption         =   "Mushrooms"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   6
         Top             =   660
         Width           =   1335
      End
      Begin VB.CheckBox chkTop 
         Caption         =   "Extra Cheese"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   180
         TabIndex        =   5
         Top             =   300
         Width           =   1695
      End
   End
   Begin VB.Frame fraSize 
      Caption         =   "Size"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   3480
      TabIndex        =   0
      Top             =   600
      Width           =   1455
      Begin VB.OptionButton optSize 
         Caption         =   "Large"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   180
         TabIndex        =   3
         Top             =   900
         Width           =   1035
      End
      Begin VB.OptionButton optSize 
         Caption         =   "Medium"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   2
         Top             =   600
         Width           =   1095
      End
      Begin VB.OptionButton optSize 
         Caption         =   "Small"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   1
         Top             =   300
         Width           =   915
      End
   End
   Begin VB.Label lblHead 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ABC Pizzas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3840
      TabIndex        =   38
      Top             =   0
      Width           =   2010
   End
   Begin VB.Label Label3 
      Caption         =   "Special Info:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3465
      TabIndex        =   37
      Top             =   2595
      Width           =   1620
   End
End
Attribute VB_Name = "frmPizza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim OrderNumber As Integer
Dim PizzaSize As Integer
Dim TotalCost As Single, NumberPizzas As Integer
Dim ToppingCost(11) As Single
Dim SizeCost(2) As Single
Dim DeliveryCost As Single
Dim TaxRate As Single
Private Const NumberToppings = 12

Private Sub PrintReceipt()
Dim I As Integer, J As Integer, S As String
'Print receipt using printer object
Printer.Scale (0, 0)-(8.5, 11#)
Printer.FontName = "Arial"
Printer.FontSize = 12
Printer.CurrentX = 1: Printer.CurrentY = 1
Printer.FontBold = True
Printer.Print "Pizza Order Number " + Format(OrderNumber, "0000")
Printer.FontBold = False
Printer.Print
Printer.CurrentX = 1
Printer.Print txtFirstName.Text + " " + txtLastName.Text
Printer.CurrentX = 1
Printer.Print txtAddress.Text
Printer.CurrentX = 1
Printer.Print txtCity.Text + ", " + txtState.Text + " " + txtZip.ToolTipText
Printer.CurrentX = 1
Printer.Print medPhone.Text
Printer.Print
Printer.CurrentX = 1
Printer.Print "Ordered" + Str(NumberPizzas) + " Pizza(s):"
denPizza.rscomPizzas.Open
denPizza.rscomPizzas.Find "OrderNumber = '" + Trim(Str(OrderNumber)) + "'", 0, adSearchForward
For I = 1 To NumberPizzas
  Printer.Print
  Printer.CurrentX = 1
  Printer.Print optSize(Val(denPizza.rscomPizzas.Fields("Size"))).Caption + " Pizza - ";
  If Val(denPizza.rscomPizzas.Fields("Delivery")) = vbChecked Then
    Printer.Print "Delivered"
  Else
    Printer.Print "To Be Picked Up"
  End If
  S = denPizza.rscomPizzas.Fields("Toppings")
  For J = 0 To NumberToppings - 1
    If Val(Mid(S, J + 1, 1)) = 1 Then
      Printer.CurrentX = 1
      Printer.Print chkTop(J).Caption
    End If
  Next J
  Printer.CurrentX = 1
  Printer.Print "Cost: $" + Format(denPizza.rscomPizzas.Fields("cost"), "0.00")
  If I <> NumberPizzas Then denPizza.rscomPizzas.MoveNext
Next I
denPizza.rscomPizzas.Close
Printer.Print
Printer.CurrentX = 1
Printer.Print "Total Cost: $" + Format(TotalCost, "0.00")
Printer.EndDoc
End Sub

Private Sub cmdBuild_Click()
'This procedure builds a message box that displays your pizza type
Dim Message As String
Dim I As Integer
Dim ToppingString As String
Dim Cost As Integer
Dim Rtn As Integer
If PizzaSize < 0 Then
  MsgBox "You must choose a size", vbOKOnly + vbInformation, "Error"
  Exit Sub
End If
Message = Message + optSize(PizzaSize).Caption + " pizza" + vbCr
Cost = SizeCost(PizzaSize)
ToppingString = ""
For I = 0 To NumberToppings - 1
  If chkTop(I).Value = vbChecked Then
    Message = Message + chkTop(I).Caption + vbCr
    ToppingString = ToppingString + "1"
    Cost = Cost + ToppingCost(I)
  Else
    ToppingString = ToppingString + "0"
  End If
Next I
If ToppingString = String(12, "0") Then
  Message = Message + "Cheese only" + vbCr
End If
If txtSpecial.Text <> "" Then
  Message = Message + txtSpecial.Text + vbCr
End If
If chkDeliver.Value = vbChecked Then
  Message = Message + "To be delivered" + vbCr
  Cost = Cost + DeliveryCost
Else
  Message = Message + "For pickup" + vbCr
End If
Message = Message + "Cost is $" + Format(Cost, "0.00")
Rtn = MsgBox(Message + vbCr + vbCr + "Is this correct?", vbYesNo + vbQuestion, "Your Pizza")
'if ok add to database and order; if not just exit
Select Case Rtn
Case vbYes
'Add pizza to pizzas table
  denPizza.rscomPizzas.Open
  denPizza.rscomPizzas.AddNew
  denPizza.rscomPizzas.Fields("OrderNumber") = OrderNumber
  denPizza.rscomPizzas.Fields("Delivery") = Str(chkDeliver.Value)
  denPizza.rscomPizzas.Fields("Size") = Str(PizzaSize)
  denPizza.rscomPizzas.Fields("Toppings") = ToppingString
  denPizza.rscomPizzas.Fields("Special") = txtSpecial.Text
  denPizza.rscomPizzas.Fields("Cost") = Cost
  denPizza.rscomPizzas.Update
  denPizza.rscomPizzas.Close
  NumberPizzas = NumberPizzas + 1
  TotalCost = TotalCost + Cost
  Cost = 0
  optSize(PizzaSize).Value = False
  PizzaSize = -1
  For I = 0 To NumberToppings - 1
    chkTop(I).Value = vbUnchecked
  Next I
  txtSpecial.Text = ""
  chkDeliver.Value = vbUnchecked
  cmdComplete.Enabled = True
Case vbNo
  Exit Sub
End Select
End Sub

Private Sub cmdComplete_Click()
Dim Rtn As Integer
TotalCost = (1 + TaxRate / 100) * TotalCost
Rtn = MsgBox("Order includes" + Str(NumberPizzas) + " pizza(s)" + vbCr + "Total cost is $" + Format(TotalCost, "0.00") + " (including tax)" + vbCr + vbCr + "Would you like a printed receipt?", vbYesNo + vbQuestion, "Order Complete")
If Rtn = vbYes Then
'Print receipt
  PrintReceipt
End If
'Write order number to database - get ready for new customer
denPizza.rscomOrders.Fields("OrderNumber") = OrderNumber
denPizza.rscomOrders.Fields("PhoneNumber") = medPhone.Text
denPizza.rscomOrders.Fields("OrderDate") = lblDate.Caption
denPizza.rscomOrders.Fields("TotalCost") = TotalCost
denPizza.rscomOrders.Update
OrderNumber = OrderNumber + 1
frmPizza.Caption = "Pizza Order #" + Format(OrderNumber, "0000")
cmdFind.Enabled = True
fraSize.Enabled = False
fraDeliver.Enabled = False
fraToppings.Enabled = False
txtSpecial.Enabled = False
cmdBuild.Enabled = False
cmdComplete.Enabled = False
medPhone.Text = "(___) ___-____"
medPhone.Enabled = True
txtFirstName.Visible = False
txtLastName.Visible = False
txtAddress.Visible = False
txtCity.Visible = False
txtState.Visible = False
txtZip.Visible = False
txtComment.Visible = False
medPhone.SetFocus
End Sub
Private Sub cmdExit_Click()
Open App.Path + "\pizza.ini" For Output As #1
Write #1, OrderNumber
Close 1
Unload Me
End Sub







Private Sub cmdFind_Click()
Dim L As Integer
'Check phone number validity (at least that it has enough characters)
medPhone.PromptInclude = False
L = Len(medPhone.Text)
medPhone.PromptInclude = True
If L <> 10 Then
  MsgBox "Phone number requires 10 digits.", vbExclamation + vbOKOnly, "Phone Number Error"
  medPhone.SetFocus
  Exit Sub
End If
  
denPizza.rscomCustomers.MoveFirst
denPizza.rscomCustomers.Find "PhoneNumber = '" + medPhone.Text + "'", 0, adSearchForward
If Not (denPizza.rscomCustomers.EOF) Then
  cmdFind.Enabled = False
  fraSize.Enabled = True
  fraDeliver.Enabled = True
  fraToppings.Enabled = True
  txtSpecial.Enabled = True
  cmdBuild.Enabled = True
  cmdComplete.Enabled = False
  medPhone.Enabled = False
  txtFirstName.Visible = True
  txtLastName.Visible = True
  txtAddress.Visible = True
  txtCity.Visible = True
  txtState.Visible = True
  txtZip.Visible = True
  txtComment.Visible = True
  TotalCost = 0
  NumberPizzas = 0
Else
  frmNewCustomer.Show vbModal
End If
End Sub

Private Sub NewOrder()
denPizza.rscomOrders.AddNew
lblDate.Caption = Format(Now, "mm/dd/yy")
medPhone.SetFocus
End Sub

Private Sub Form_Activate()
If Not (NewCustomer) Then
  Call NewOrder
Else
  NewCustomer = False
  cmdFind_Click
End If
End Sub

Private Sub Form_Load()
Open App.Path + "\pizza.ini" For Input As #1
Input #1, OrderNumber
Close 1
frmPizza.Caption = "Pizza Order #" + Format(OrderNumber, "0000")
NewCustomer = False
PizzaSize = -1
'Topping cost information (veggies 50 cents, meats 1 dollar)
ToppingCost(0) = 0.5
ToppingCost(1) = 0.5
ToppingCost(2) = 0.5
ToppingCost(3) = 0.5
ToppingCost(4) = 0.5
ToppingCost(5) = 0.5
ToppingCost(6) = 1#
ToppingCost(7) = 1#
ToppingCost(8) = 1#
ToppingCost(9) = 1#
ToppingCost(10) = 1#
ToppingCost(11) = 1#
'Size cost
SizeCost(0) = 4#
SizeCost(1) = 6#
SizeCost(2) = 8#
'Delivery cost
DeliveryCost = 1.5
'Tax rate
TaxRate = 8.9
'Open database
denPizza.conPizza.Open "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=C:\Users\SAHIL\Downloads\pro_0\pro\PizzaDB.mdb"
'Open customers and orders recordset
denPizza.rscomCustomers.Open "SELECT * FROM Customers"
denPizza.rscomOrders.Open "SELECT * FROM Orders"
'Bind controls
Set txtFirstName.DataSource = denPizza
Set txtFirstName.DataSource = denPizza
Set txtLastName.DataSource = denPizza
Set txtAddress.DataSource = denPizza
Set txtCity.DataSource = denPizza
Set txtState.DataSource = denPizza
Set txtZip.DataSource = denPizza
Set txtComment.DataSource = denPizza
End Sub


Private Sub medPhone_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
  cmdFind_Click
End If
End Sub


Private Sub optSize_Click(Index As Integer)
PizzaSize = Index
End Sub




