VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSearchItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarFileName As String 'local copy
Private mvarSize As String 'local copy
Private mvarFileType As String 'local copy
Private mvarPath As String 'local copy
Private mvarTrackName As String 'local copy
Private mvarAttributes As String 'local copy
Public Property Let Attributes(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Attributes = 5
    mvarAttributes = vData
End Property


Public Property Get Attributes() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Attributes
    Attributes = mvarAttributes
End Property



Public Property Let TrackName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TrackName = 5
    mvarTrackName = vData
End Property


Public Property Get TrackName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TrackName
    TrackName = mvarTrackName
End Property



Public Property Let Path(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Path = 5
    mvarPath = vData
End Property


Public Property Get Path() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Path
    Path = mvarPath
End Property



Public Property Let FileType(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FileType = 5
    mvarFileType = vData
End Property


Public Property Get FileType() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FileType
    FileType = mvarFileType
End Property



Public Property Let Size(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Size = 5
    mvarSize = vData
End Property


Public Property Get Size() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Size
    Size = mvarSize
End Property



Public Property Let FileName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FileName = 5
    mvarFileName = vData
End Property


Public Property Get FileName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FileName
    FileName = mvarFileName
End Property



