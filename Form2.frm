VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   4410
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   4410
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command1 
      Caption         =   "Click Me"
      Height          =   495
      Left            =   1740
      TabIndex        =   0
      Top             =   3630
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Unfortunately Collections do not expose the Keys of Items. So I made two little functions which
'return an Item's Key by Index or Index by Key [Key = ItemKey(Index, Collection)]           and
'                                              [Index = ItemIndex(Key, Collection)].

Private Declare Sub PokeLong Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, Optional ByVal Length As Long = 4)

Private SampleCol   As Collection

Private Sub Command1_Click()

  Dim i         As Long

    Cls
    On Error Resume Next
        Print "Item Count "; Tab; SampleCol.Count; ""
        Print
        For i = 0 To SampleCol.Count + 1
            Print "Key("; i; ")"; Tab;

            Print ItemKey(i, SampleCol),            'get key by index

            Print Err.Description
            Err.Clear
        Next i
        
        Print
        Print "Key ""8th is form2"" is at Index "; ItemIndex("8th is form2", SampleCol)
        Print "First ""<no key>"" is at Index    "; ItemIndex(vbNullString, SampleCol)

        SampleCol.Remove 1
        If Err = 0 Then
            Print
            Print "Removed Item(1)"
        End If
    On Error GoTo 0

End Sub

Private Sub Form_Load()

    Set SampleCol = New Collection
    With SampleCol
        .Add 1, Key:="1st"
        .Add 2, Key:="2nd"
        .Add 3 '3rd has no key
        .Add Command1, Key:="4th is the Button on this Form"
        .Add 5, Key:="5th before 2nd", Before:=2
        .Add 6, Key:="6th after 3rd", After:=3
        .Add SampleCol, Key:="7th is the Collection itself"
        .Add Me, Key:="8th is Form2"
    End With 'SAMPLECOL

End Sub

Private Function ItemKey(ByVal Index As Long, Coll As Collection) As String

  'optimized get collection sKey by index
  'Private Declare Sub PokeLong Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, Optional ByVal Length As Long = 4)

  Dim i     As Long
  Dim Ptr   As Long
  Dim sKey  As String

    If Coll Is Nothing Then                             'oops!
        Err.Raise 91                                    'No object
      Else 'NOT COLL...
        Select Case Index
          Case Is < 1, Is > Coll.Count                  'oops!
            Err.Raise 9                                 'Index out of range
          Case Is <= Coll.Count / 2                     'walk items upwards from first
            PokeLong Ptr, ByVal ObjPtr(Coll) + 24       'first Ptr
            For i = 2 To Index
                PokeLong Ptr, ByVal Ptr + 24            'next Ptr
            Next i
          Case Else                                     'walk items downwards from last
            PokeLong Ptr, ByVal ObjPtr(Coll) + 28       'last Ptr
            For i = Coll.Count - 1 To Index Step -1
                PokeLong Ptr, ByVal Ptr + 20            'prev Ptr
            Next i
        End Select
        i = StrPtr(sKey)                                'save StrPtr
        PokeLong ByVal VarPtr(sKey), ByVal Ptr + 16     'replace StrPtr by that from collection sKey (which is null if there ain't no sKey)
        ItemKey = sKey                                  'now copy it to function value
        PokeLong ByVal VarPtr(sKey), i                  'and finally restore original StrPtr
    End If

End Function

Private Function ItemIndex(ByVal Key As String, Coll As Collection, Optional ByVal Compare As VbCompareMethod = vbTextCompare) As Long

  'get collection index by key
  'Private Declare Sub PokeLong Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, Optional ByVal Length As Long = 4)

  Dim Ptr   As Long
  Dim sKey  As String
  Dim aKey  As Long

    If Coll Is Nothing Then                             'oops!
        Err.Raise 91                                    'No object
      Else 'NOT COLL...
        If Coll.Count Then
            aKey = StrPtr(sKey)                         'save StrPtr
            PokeLong Ptr, ByVal ObjPtr(Coll) + 24       'first Ptr
            ItemIndex = 1                               'walk items upwards from first
            Do
                PokeLong ByVal VarPtr(sKey), ByVal Ptr + 16
                If StrComp(Key, sKey, Compare) = 0 Then 'equal
                    Exit Do                             'found
                End If
                ItemIndex = ItemIndex + 1               'next Index
                PokeLong Ptr, ByVal Ptr + 24            'next Ptr
            Loop Until Ptr = 0                          'end of chain
            PokeLong ByVal VarPtr(sKey), aKey           'restore original StrPtr
        End If
        If Ptr = 0 Then
            ItemIndex = -1                              'key not found
        End If
    End If

End Function

':) Ulli's VB Code Formatter V2.23.12 (2007-Mrz-07 23:47)  Decl: 7  Code: 155  Total: 162 Lines
':) CommentOnly: 10 (6,2%)  Commented: 43 (26,5%)  Empty: 28 (17,3%)  Max Logic Depth: 5
