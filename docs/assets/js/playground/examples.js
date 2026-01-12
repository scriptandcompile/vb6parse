/**
 * VB6Parse Playground - Example Code Snippets
 * 
 * This module contains example VB6 code snippets for each file type.
 * Users can select these from the Examples dropdown to quickly test the parser.
 * 
 * TODO: Add more realistic examples from real VB6 projects
 * TODO: Consider loading examples from external files for larger samples
 */

export const examples = {
    'simple-module': {
        name: 'Simple Module',
        fileType: 'module',
        code: `VERSION 1.0 CLASS
Attribute VB_Name = "MathUtils"

Option Explicit

' Simple mathematical utilities
Public Function Add(ByVal a As Long, ByVal b As Long) As Long
    Add = a + b
End Function

Public Function Multiply(ByVal x As Double, ByVal y As Double) As Double
    Multiply = x * y
End Function

Public Sub PrintResult(ByVal result As Double)
    Debug.Print "Result: " & result
End Sub
`
    },
    
    'class-with-properties': {
        name: 'Class with Properties',
        fileType: 'class',
        code: `VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Person"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Private m_Name As String
Private m_Age As Integer

' Name property
Public Property Get Name() As String
    Name = m_Name
End Property

Public Property Let Name(ByVal value As String)
    m_Name = value
End Property

' Age property
Public Property Get Age() As Integer
    Age = m_Age
End Property

Public Property Let Age(ByVal value As Integer)
    If value >= 0 And value <= 150 Then
        m_Age = value
    Else
        Err.Raise vbObjectError + 1, "Person", "Invalid age"
    End If
End Property

Public Function GetInfo() As String
    GetInfo = m_Name & " is " & m_Age & " years old"
End Function
`
    },
    
    'form-with-controls': {
        name: 'Form with Controls',
        fileType: 'form',
        code: `VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Calculator"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnCalculate 
      Caption         =   "Calculate"
      Height          =   495
      Left            =   1560
      TabIndex        =   3
      Top             =   2520
      Width           =   1575
   End
   Begin VB.TextBox txtResult 
      Height          =   495
      Left            =   1440
      TabIndex        =   2
      Top             =   1800
      Width           =   1815
   End
   Begin VB.TextBox txtNumber2 
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Top             =   1080
      Width           =   1815
   End
   Begin VB.TextBox txtNumber1 
      Height          =   495
      Left            =   1440
      TabIndex        =   0
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Result:"
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Number 2:"
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Number 1:"
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   480
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub btnCalculate_Click()
    Dim num1 As Double
    Dim num2 As Double
    
    num1 = Val(txtNumber1.Text)
    num2 = Val(txtNumber2.Text)
    
    txtResult.Text = CStr(num1 + num2)
End Sub

Private Sub Form_Load()
    Me.Caption = "Simple Calculator"
End Sub
`
    },
    
    'project-file': {
        name: 'Project File',
        fileType: 'project',
        code: `Type=Exe
Form=Form1.frm
Reference=*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#..\\..\\..\\Windows\\SysWOW64\\stdole2.tlb#OLE Automation
Module=MathUtils; MathUtils.bas
Class=Person; Person.cls
IconForm="Form1"
Startup="Form1"
HelpFile=""
Title="MyVB6App"
ExeName32="MyVB6App.exe"
Command32=""
Name="MyVB6App"
HelpContextID="0"
CompatibleMode="0"
MajorVer=1
MinorVer=0
RevisionVer=0
AutoIncrementVer=0
ServerSupportFiles=0
VersionCompanyName="Example Corp"
CompilationType=0
OptimizationType=0
FavorPentiumPro(tm)=0
CodeViewDebugInfo=0
NoAliasing=0
BoundsCheck=0
OverflowCheck=0
FlPointCheck=0
FDIVCheck=0
UnroundedFP=0
StartMode=0
Unattended=0
Retained=0
ThreadPerObject=0
MaxNumberOfThreads=1
`
    }
};

/**
 * Get an example by ID
 * @param {string} exampleId - The example identifier
 * @returns {object|null} The example object or null if not found
 */
export function getExample(exampleId) {
    return examples[exampleId] || null;
}

/**
 * Get all example IDs
 * @returns {string[]} Array of example IDs
 */
export function getExampleIds() {
    return Object.keys(examples);
}

/**
 * Get examples by file type
 * @param {string} fileType - 'module', 'class', 'form', or 'project'
 * @returns {object[]} Array of examples for the given type
 */
export function getExamplesByType(fileType) {
    return Object.entries(examples)
        .filter(([_, example]) => example.fileType === fileType)
        .map(([id, example]) => ({ id, ...example }));
}

/**
 * TODO: Future enhancements
 * - Add more complex examples (error handling, API calls, database access)
 * - Load examples from JSON file for easier maintenance
 * - Add descriptions and learning points for each example
 * - Support user-contributed examples
 * - Add syntax error examples for testing error handling
 */
