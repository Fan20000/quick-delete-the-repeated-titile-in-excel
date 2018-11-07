# quick-delete-the-repeated-titile-in-excel

Sub 删除相同行()
    Application.ScreenUpdating = False
    Dim arr, d As Object, i&, s$, rng As Range
    Set d = CreateObject("scripting.dictionary")
    arr = ActiveSheet.UsedRange.Value            '设指定区域如：=Range("a5:k" & Range("a65536").End(xlUp).Row)
    For i = 1 To UBound(arr)
        s = Join(Application.Index(arr, i), "/")  '设指定列数据重复如:s=arr(i, 2) & "/" & arr(i, 4) & "/" & arr(i, 5) & "/" & arr(i, 6) & "/" & arr(i, 7)
          If Not d.Exists(s) Then
             d(s) = ""
          Else
             If rng Is Nothing Then Set rng = Cells(i, 1) Else Set rng = Union(rng, Cells(i, 1))
         End If
     Next
     If Not rng Is Nothing Then rng.EntireRow.Delete
     For i = UBound(arr) To 1 Step -1    '删除中间的空行
         Set rng = Cells(i, 1).Resize(1, UBound(arr, 2))
         If Application.CountA(rng) = 0 Then Rows(i).Delete
     Next
     For j = UBound(arr, 2) To 1 Step -1 '删除中间的空列
         Set rng = Cells(1, j).Resize(UBound(arr), 1)
         If Application.CountA(rng) = 0 Then Columns(j).Delete
     Next
    Application.ScreenUpdating = True
End Sub
