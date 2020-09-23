VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTrackItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarTrackName As String 'local copy
Private mvarCategory As String 'local copy
Private mvarCreated As String 'local copy
Private mvarNumRecords As Long 'local copy
Private mvarTrackSize As String 'local copy
Private mvarBasePath As String 'local copy
Private mvarComment As String 'local copy
Public Property Let Comment(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Comment = 5
    mvarComment = vData
End Property


Public Property Get Comment() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Comment
    Comment = mvarComment
End Property



Public Property Let BasePath(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.BasePath = 5
    mvarBasePath = vData
End Property


Public Property Get BasePath() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BasePath
    BasePath = mvarBasePath
End Property



Public Property Let TrackSize(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TrackSize = 5
    mvarTrackSize = vData
End Property


Public Property Get TrackSize() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TrackSize
    TrackSize = mvarTrackSize
End Property



Public Property Let NumRecords(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NumRecords = 5
    mvarNumRecords = vData
End Property


Public Property Get NumRecords() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NumRecords
    NumRecords = mvarNumRecords
End Property



Public Property Let Created(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Created = 5
    mvarCreated = vData
End Property


Public Property Get Created() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Created
    Created = mvarCreated
End Property



Public Property Let Category(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Category = 5
    mvarCategory = vData
End Property


Public Property Get Category() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Category
    Category = mvarCategory
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



