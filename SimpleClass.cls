VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SimpleClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type udtType
    value1              As Long
    value2              As String
    value3              As Single
End Type

Private lngVariable1    As Long
Private strVariable2    As String
Private sngVariable3    As Single

Private udtMyType       As udtType

'''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''

' adding your variables to a propertybag
' is some paperwork, but also easy to read
Public Function Clone() As SimpleClass
    Dim clsNewClass As SimpleClass
    Dim props       As PropertyBag

    ' the clone
    Set clsNewClass = New SimpleClass

    ' PropertyBag which carries all the vars of this class
    Set props = New PropertyBag

    props.WriteProperty "var1", lngVariable1
    props.WriteProperty "var2", strVariable2
    props.WriteProperty "var3", sngVariable3

    ' call the Clone Constructor of the new instance
    ' with a PropertyBag reference and the UDT
    CallCloneConstructor clsNewClass, _
                         VarPtr(props), _
                         VarPtr(udtMyType)

    Set Clone = clsNewClass
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get MyUDTvalue1() As Long
    MyUDTvalue1 = udtMyType.value1
End Property

Public Property Let MyUDTvalue1(ByVal lngVal As Long)
    udtMyType.value1 = lngVal
End Property

Public Property Get MyUDTvalue2() As String
    MyUDTvalue2 = udtMyType.value2
End Property

Public Property Let MyUDTvalue2(ByVal strText As String)
    udtMyType.value2 = strText
End Property

Public Property Get MyVariable() As Long
    MyVariable = lngVariable1
End Property

Public Property Let MyVariable(ByVal lngVar As Long)
    lngVariable1 = lngVar
End Property

Public Property Get MyString() As String
    MyString = strVariable2
End Property

Public Property Let MyString(ByVal strText As String)
    strVariable2 = strText
End Property

Public Property Get MySingle() As Single
    MySingle = sngVariable3
End Property

Public Property Let MySingle(ByVal sngVal As Single)
    sngVariable3 = sngVal
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' The Clone Constructor
'''
''' Can look like anything, as long as you pass the correct
''' parameters to it in your Clone method!
'''
''' This method has to be the last in this interface!
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub CloneConstructor(props As PropertyBag, udt As udtType)
    lngVariable1 = props.ReadProperty("var1")
    strVariable2 = props.ReadProperty("var2")
    sngVariable3 = props.ReadProperty("var3")
    udtMyType = udt
End Sub
