VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "realObject.Show"
      Height          =   975
      Left            =   1320
      TabIndex        =   1
      Top             =   1680
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Set realObject = frmHello"
      Height          =   975
      Left            =   1200
      TabIndex        =   0
      Top             =   240
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Code created by Rodolfo González Ruiz
'San Salvador, El Salvador, Central America
'rodolfo@flashmail.com
'24/10/2000

'Use the following code as you wish... yet, vote for it if you
'liked it please... thx.

Dim realObject As Object

Private Sub Command1_Click()
    Dim objScriptInMemory As New ScriptControl
    
    'sets the language to execute or evaluate the scripts on
    objScriptInMemory.Language = "VBS"
    
    'add an object to the script ran by the scriptcontrol
        'this would be the same as executing a 'set frmHello = XX' statement,
        'where XX were a valid object, active in the script.
        'Yet in this case we are passing the script an object
        'of our own code to it, and giving the frmHelloInScript variable
        'a reference to it.
    objScriptInMemory.AddObject "frmHelloInScript", frmHello
    
    'Execute a statement in the script (set the macro substitution variable, 'scriptObject' in this case)
    objScriptInMemory.ExecuteStatement "set scriptObject = frmHelloInScript"
    
    'Assign the value of the Script's variable (an object in this case)
        'to a variable in our code
    Set realObject = objScriptInMemory.Eval("scriptObject")
    
    'Release our variable 'frmHello' from memory (notice through a WATCH
        'that the variable maintains its class frmHello, yet its reference
        'to an instance of that object is destroyed)
    Set frmHello = Nothing
    
    'Because we used DIM to create an instance of the ScriptControl
    'inside this Sub, the object and all of its code is destroyed
    'after the following End Sub.
End Sub

Private Sub Command2_Click()
    'Use the object in our code.
    realObject.Show
End Sub
