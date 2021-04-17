Sub save()
Cells(50, 50).Select
For i = 1 To 4
    For j = 1 To 4
        If Cells(i, j) <> "" Then
            Cells(9 + i, j) = Cells(i, j) * 1
        Else
            Cells(9 + i, j).ClearContents
        End If
    Next j
Next i

Cells(11, 10) = Cells(2, 10) * 1
Cells(13, 10) = Cells(4, 10) * 1

End Sub

Sub backup()
ActiveSheet.Unprotect "Tkdlqjrj"
For i = 1 To 4
    For j = 1 To 4
        If Cells(9 + i, j) <> "" Then
            Cells(i, j) = Cells(9 + i, j) * 1
        Else
            Cells(i, j).ClearContents
        End If
    Next j
Next i

Cells(2, 10) = Cells(11, 10) * 1
Cells(4, 10) = Cells(13, 10) * 1

Call color
ActiveSheet.Protect "Tkdlqjrj", False, True
End Sub

Sub color()
For i = 1 To 4
    For j = 1 To 4
        Cells(i, j).Font.color = RGB(249, 246, 242)
        If Cells(i, j) = "" Then
            Cells(i, j).Interior.color = RGB(205, 193, 180)
            Cells(i, j).Font.color = RGB(205, 193, 180)
            
        ElseIf Cells(i, j) >= 2 And Cells(i, j) < 4 Then
            Cells(i, j).Interior.color = RGB(238, 228, 218)
            Cells(i, j).Font.color = RGB(119, 110, 101)
            
        ElseIf Cells(i, j) < 8 Then
            Cells(i, j).Interior.color = RGB(237, 224, 200)
            Cells(i, j).Font.color = RGB(119, 110, 101)
            
        ElseIf Cells(i, j) < 16 Then
            Cells(i, j).Interior.color = RGB(242, 177, 121)
            
        ElseIf Cells(i, j) < 32 Then
            Cells(i, j).Interior.color = RGB(245, 149, 99)
            
        ElseIf Cells(i, j) < 64 Then
            Cells(i, j).Interior.color = RGB(246, 124, 95)
            
        ElseIf Cells(i, j) < 128 Then
            Cells(i, j).Interior.color = RGB(246, 94, 59)
            
        ElseIf Cells(i, j) < 256 Then
            Cells(i, j).Interior.color = RGB(237, 207, 114)
            
        ElseIf Cells(i, j) < 512 Then
            Cells(i, j).Interior.color = RGB(237, 204, 97)
            
        ElseIf Cells(i, j) < 1024 Then
            Cells(i, j).Interior.color = RGB(237, 200, 80)
            
        ElseIf Cells(i, j) < 2048 Then
            Cells(i, j).Interior.color = RGB(237, 197, 64)
            
        ElseIf Cells(i, j) < 4096 Then
            Cells(i, j).Interior.color = RGB(237, 194, 46)
            
        ElseIf Cells(i, j) >= 4096 Then
            Cells(i, j).Interior.color = RGB(60, 58, 50)
            Cells(i, j).Font.color = RGB(247, 244, 240)
        Else
            Cells(i, j).Interior.color = RGB(205, 193, 180)
            Cells(i, j).Font.color = RGB(205, 193, 180)
        
        End If
    Next j
Next i

End Sub

Sub gameover()

For i = 1 To 4
    For j = 1 To 4
        If Cells(i, j) <> "" Then
            check1 = check1 + 1
        End If
    Next j
Next i

If check1 = 16 Then
    For i = 1 To 4
        For j = 1 To 3
            If Cells(i, j) = Cells(i, j + 1) Then
                check2 = 1
                Exit For
            End If
        Next j
    Next i
    If check2 = 0 Then
        For i = 1 To 4
            For j = 1 To 3
                If Cells(j, i) = Cells(j + 1, i) Then
                    check2 = 1
                    Exit For
                End If
            Next j
        Next i
    End If
End If

If check1 = 16 And check2 = 0 Then
    Call save
    MsgBox "Game Over!!" & Chr(10) & Chr(13) & "최종 점수 : " & Cells(2, 10) & "점", 64, "2048 게임"
    If Cells(2, 10) >= Cells(4, 10) Then
        MsgBox "최고 점수를 달성하였습니다!!", 64, "2048 게임"
    End If
    Call reset
End If

End Sub

Sub random(move, num)
Randomize

Call color

i = 0
j = 0

Dim arr(3)
arr(0) = 2
arr(1) = 2
arr(2) = 2
arr(3) = 4

Dim row(16)
Dim col(16)
Dim count

If move = 1 Then
    For a = 1 To num
        count = 0
        For i = 1 To 4
            For j = 1 To 4
                If Cells(i, j) = "" Then
                    col(count) = i
                    row(count) = j
                    count = count + 1
                End If
            Next j
        Next i
        
        r = Int(count * Rnd)
        i = col(r)
        j = row(r)
        
        Cells(i, j) = arr(Int(Rnd * 4)) * 1
        Cells(i, j).Interior.color = RGB(130, 230, 255)
        Cells(i, j).Font.color = RGB(119, 110, 101)
    Next a
        
End If

End Sub

Sub reset()
ActiveSheet.Unprotect "Tkdlqjrj"
Application.ScreenUpdating = False
Range(Cells(1, 1), Cells(4, 4)).Interior.color = RGB(205, 193, 180)
Range(Cells(1, 1), Cells(4, 4)).Font.color = RGB(205, 193, 180)
Range(Cells(1, 1), Cells(4, 4)).ClearContents
Range(Cells(10, 1), Cells(13, 4)).ClearContents

Cells(2, 10) = 0
Call random(1, 2)
Call save
Application.ScreenUpdating = True
ActiveSheet.Protect "Tkdlqjrj", False, True
End Sub

Sub 상()
ActiveSheet.Unprotect "Tkdlqjrj"
Application.ScreenUpdating = False
Call save

move = 0

For j = 1 To 4
    For i = 2 To 4
        If Cells(i, j) <> "" Then
            For k = i - 1 To 1 Step -1
                If Cells(k, j) = "" Then
                    Cells(k, j) = Cells(k + 1, j) * 1
                    move = 1
                    Cells(k + 1, j).ClearContents
                Else
                    Exit For
                End If
            Next k
        End If
    Next i
    
    For i = 1 To 3
        If Cells(i, j) <> "" And Cells(i, j) = Cells(i + 1, j) Then
            move = 1
            score = 8
            If Cells(i, j) >= 1024 Then
                score = 16
            End If
            
            Cells(2, 10) = Cells(2, 10) + Cells(i, j) * score
            
            
            Cells(i, j) = Cells(i + 1, j) * 2
            Cells(i + 1, j).ClearContents
            Cells(i + 1, j).Interior.color = RGB(205, 193, 180)
            
            If Cells(i, j) = 2048 Then
                MsgBox "2048이 만들어졌습니다!", 64, "2048 게임"
            End If
            
            For k = i + 1 To 3
                If Cells(k + 1, j) = "" Then
                        Cells(k, j).ClearContents
                    Else
                        Cells(k, j) = Cells(k + 1, j) * 1
                End If
                Cells(k + 1, j).ClearContents
            Next k
        End If
    Next i
Next j

Call random(move, 1)
If Cells(2, 10) >= Cells(4, 10) Then
    Cells(4, 10) = Cells(2, 10) * 1
    ActiveWorkbook.save
End If
Call gameover
Application.ScreenUpdating = True
ActiveSheet.Protect "Tkdlqjrj", False, True
End Sub

Sub 하()
ActiveSheet.Unprotect "Tkdlqjrj"
Application.ScreenUpdating = False
Call save

move = 0

For j = 1 To 4
    For i = 2 To 4
        If Cells(5 - i, j) <> "" Then
            For k = i - 1 To 1 Step -1
                If Cells(5 - k, j) = "" Then
                    Cells(5 - k, j) = Cells(5 - (k + 1), j) * 1
                    move = 1
                    Cells(5 - (k + 1), j).ClearContents
                Else
                    Exit For
                End If
            Next k
        End If
    Next i
    
    For i = 1 To 3
        If Cells(5 - i, j) <> "" And Cells(5 - i, j) = Cells(5 - (i + 1), j) Then
            move = 1
            score = 8
            If Cells(5 - i, j) >= 1024 Then
                score = 16
            End If
            
            Cells(2, 10) = Cells(2, 10) + Cells(5 - i, j) * score
            
            Cells(5 - i, j) = Cells(5 - (i + 1), j) * 2
            Cells(5 - (i + 1), j).ClearContents
            Cells(5 - (i + 1), j).Interior.color = RGB(205, 193, 180)
            
            If Cells(5 - i, j) = 2048 Then
                MsgBox "2048이 만들어졌습니다!", 64, "2048 게임"
            End If
            
            For k = i + 1 To 3
                If Cells(5 - (k + 1), j) = "" Then
                        Cells(5 - k, j).ClearContents
                    Else
                        Cells(5 - k, j) = Cells(5 - (k + 1), j) * 1
                End If
                Cells(5 - (k + 1), j).ClearContents
            Next k
        End If
    Next i
Next j

Call random(move, 1)
If Cells(2, 10) >= Cells(4, 10) Then
    Cells(4, 10) = Cells(2, 10) * 1
    ActiveWorkbook.save
End If
Call gameover
Application.ScreenUpdating = True
ActiveSheet.Protect "Tkdlqjrj", False, True
End Sub

Sub 좌()
ActiveSheet.Unprotect "Tkdlqjrj"
Application.ScreenUpdating = False
Call save

move = 0

For j = 1 To 4
    For i = 2 To 4
        If Cells(j, i) <> "" Then
            For k = i - 1 To 1 Step -1
                If Cells(j, k) = "" Then
                    Cells(j, k) = Cells(j, k + 1) * 1
                    move = 1
                    Cells(j, k + 1).ClearContents
                Else
                    Exit For
                End If
            Next k
        End If
    Next i
    
    For i = 1 To 3
        If Cells(j, i) <> "" And Cells(j, i) = Cells(j, i + 1) Then
            move = 1
            score = 8
            If Cells(j, i) >= 1024 Then
                score = 16
            End If
            
            Cells(2, 10) = Cells(2, 10) + Cells(j, i) * score
            
            Cells(j, i) = Cells(j, i + 1) * 2
            Cells(j, i + 1).ClearContents
            Cells(j, i + 1).Interior.color = RGB(205, 193, 180)
            
            If Cells(j, i) = 2048 Then
                MsgBox "2048이 만들어졌습니다!", 64, "2048 게임"
            End If
            
            For k = i + 1 To 3
                
                If Cells(j, k + 1) = "" Then
                        Cells(j, k).ClearContents
                    Else
                        Cells(j, k) = Cells(j, k + 1) * 1
                End If
                
                Cells(j, k + 1).ClearContents
            Next k
        End If
    Next i
Next j

Call random(move, 1)
If Cells(2, 10) >= Cells(4, 10) Then
    Cells(4, 10) = Cells(2, 10) * 1
    ActiveWorkbook.save
End If
Call gameover
Application.ScreenUpdating = True
ActiveSheet.Protect "Tkdlqjrj", False, True
End Sub

Sub 우()
ActiveSheet.Unprotect "Tkdlqjrj"
Application.ScreenUpdating = False
Call save

move = 0


For j = 1 To 4
    For i = 2 To 4
        If Cells(j, 5 - i) <> "" Then
            For k = i - 1 To 1 Step -1
                If Cells(j, 5 - k) = "" Then
                    
                    Cells(j, 5 - k) = Cells(j, 5 - (k + 1)) * 1
                    move = 1
                    Cells(j, 5 - (k + 1)).ClearContents
                Else
                    Exit For
                End If
            Next k
        End If
    Next i
    
    For i = 1 To 3
        If Cells(j, 5 - i) <> "" And Cells(j, 5 - i) = Cells(j, 5 - (i + 1)) Then
            move = 1
            score = 8
            If Cells(j, 5 - i) >= 1024 Then
                score = 16
            End If
            
            Cells(2, 10) = Cells(2, 10) + Cells(j, 5 - i) * score
            If Cells(2, 10) >= Cells(4, 10) Then
                Cells(4, 10) = Cells(2, 10) * 1
                ActiveWorkbook.save
            End If
            
            Cells(j, 5 - i) = Cells(j, 5 - (i + 1)) * 2
            Cells(j, 5 - (i + 1)).ClearContents
            Cells(j, 5 - (i + 1)).Interior.color = RGB(205, 193, 180)
            
            If Cells(j, 5 - i) = 2048 Then
                MsgBox "2048이 만들어졌습니다!", 64, "2048 게임"
            End If
            
            For k = i + 1 To 3
                If Cells(j, 5 - (k + 1)) = "" Then
                        Cells(j, 5 - k).ClearContents
                    Else
                        Cells(j, 5 - k) = Cells(j, 5 - (k + 1)) * 1
                End If
                Cells(j, 5 - (k + 1)).ClearContents
            Next k
        End If
    Next i
Next j

Call random(move, 1)
If Cells(2, 10) >= Cells(4, 10) Then
    Cells(4, 10) = Cells(2, 10) * 1
    ActiveWorkbook.save
End If
Call gameover
Application.ScreenUpdating = True
ActiveSheet.Protect "Tkdlqjrj", False, True
End Sub

Sub keyset()
    Application.OnKey "{UP}", "상"
    Application.OnKey "{DOWN}", "하"
    Application.OnKey "{LEFT}", "좌"
    Application.OnKey "{RIGHT}", "우"
    
    Application.OnKey "{F5}", "reset"
    
    Application.OnKey "{BS}", "backup"
End Sub

Sub unkeyset()
    Application.OnKey "{UP}"
    Application.OnKey "{DOWN}"
    Application.OnKey "{LEFT}"
    Application.OnKey "{RIGHT}"
    
    Application.OnKey "{F5}"
    
    Application.OnKey "{BS}"
End Sub
