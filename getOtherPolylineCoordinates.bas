Attribute VB_Name = "getOtherPolylineCoordinates"
'Option Explicit
''' CODE VBA NHAP TOA DO POLYLINE TU AUTOCAD
''' TAC GIA: HA NGUYEN
''' FACEBOOK.COM/G.TECHNICAL

Sub GetCoordinates()
    Dim rox, coy As Double
    Dim sname As String
    Dim Lastcolumn As Integer
    'rox = 2: coy = 14
    'sname = "Polyline"
    shName = Range("A2").Value
    'rox = Range("M1").Value
    'coy = Range("N1").Value
    'Checking whether "Master" sheet already exists in the workbook
    'For Each Source In ActiveWorkbook.Worksheets 'ThisWorkbook.Worksheets
     '   If Source.Name = shName Then
      '      MsgBox "Sheet already exist"
            'Exit Sub
       ' Else
        '    MsgBox "CAN TONG HOP SO LIEU'"
         '   Exit Sub
        'End If
    'Next
    On Error Resume Next
    Set ws = ActiveWorkbook.Sheets(shName)
    With ws
        Lastcolumn = .Cells(3, Columns.Count).End(xlToLeft).Column
        rox = 4
        coy = Lastcolumn + 1
        .Cells(3, coy).Value = "Toa do X"
        .Cells(3, coy + 1).Value = "Toa do Y"
    End With
    MsgBox "- Select Polyline/Polylines from AutoCAD for And Press Enter - ", vbInformation, "Select  Polyline"
    Call getOtherPolylineCoordinatesFromAutoCAD(rox, coy, shName)
End Sub

Private Sub getOtherPolylineCoordinatesFromAutoCAD(ByVal cx As Integer, ByVal cy As Integer, ByVal sname As String)
    Set NewDC = GetObject(, "AutoCAD.Application")
    Set A2Kdwg = NewDC.ActiveDocument
    Dim Selection As AcadSelectionSet
    Dim poly As AcadLWPolyline
    Dim Obj As AcadEntity
    Dim Bound As Double
    Dim x, y As Double
    Dim rows, i, scount As Integer
    '---Search Object from SelectionSet and Delete If Found ----''
    For i = 0 To A2Kdwg.SelectionSets.Count - 1
        If A2Kdwg.SelectionSets.Item(i).Name = "AcDbPolyline" Then
            ''-- Delete Object Name from AutoCAD SelectionSet ---''
            A2Kdwg.SelectionSets.Item(i).Delete
            Exit For
        End If
    Next i
    ''-- Add Object to AutoCad SelectionSet ----''
    Set Selection = A2Kdwg.SelectionSets.Add("AcDbPolyline")
    ''-- Select Object from AutoCad Screen ---'''
    Selection.SelectOnScreen
    ''-- Get Coordinates of Object if Object name is ACadPolyline--''
    rows = cx
    For Each Obj In Selection
        If Obj.ObjectName = "AcDbPolyline" Then
            ''- Set Obj as Polyline--''
            Set poly = Obj
            On Error Resume Next
            ''-- Set Size of Coordinates Like array Size--''
            Bound = UBound(poly.Coordinates)
            '' Starting Index of Excel Row to insert Coordinates --''
            rows = rows
            ''-- Display Coordinates one by one to Excel Columns --'''
            For i = 0 To Bound
                ''-- Set Coordinates into Variables--'''
                x = Round(poly.Coordinates(i), 3)
                y = Round(poly.Coordinates(i + 1), 3)
                ''-- Set Coordinates into Excel Columns --'''
                Worksheets(sname).Cells(rows, cy) = Round(x, 3)
                Worksheets(sname).Cells(rows, cy + 1) = Round(y, 4)
                ''- Increment variable for Excel Rows ---''
                rows = rows + 1
                ''--- Increment Counter variable to get Next point of Polyline --'''
                i = i + 1
            Next
        Else
            MsgBox "--- This is not a Polyline --- ", vbInformation, "Please Select a Polyline"
        End If
        rows = rows + 1
    Next Obj
End Sub
