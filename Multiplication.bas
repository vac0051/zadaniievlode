Attribute VB_Name = "Module1"
Sub MultiplyNumbers()
    Dim S As Variant, SS As Variant
    Dim S1 As String, S2 As String
    Dim SS1 As String, SS2 As String
    Dim I As Integer, J As Integer
    Dim D1 As Integer, D2 As Integer
    Dim DD1 As Integer, DD2 As Integer
    Dim G As Integer, V As Integer
    Dim a As Integer, b As Integer
    Dim E As Integer, D As Integer
    Dim f As Integer, ff As Integer, fff As Integer
    Dim x As Integer

    Rem Находим числа десятков и единиц каждого перемножаемого числа из ячеек A3 и B3
    S = Cells(3, 1).Value
    SS = Cells(3, 2).Value
    
    If S = "" Or SS = "" Then Exit Sub
    
    S1 = Mid(S, 1, 1)
    S2 = Mid(S, 2, 1)
    SS1 = Mid(SS, 1, 1)
    SS2 = Mid(SS, 2, 1)

    Rem Примитивно очищаем таблицу проверки счёта
    For I = 0 To 17
        For J = 0 To 17
            Cells(2 + I, 6 + J).Value = ""
        Next J
    Next I

    Rem Определяем количество красных и синих горизонтальных прямых линий G
    D1 = CInt(S1)
    D2 = CInt(S2)
    G = D1 + D2

    Rem Определяем количество красных и синих вертикальных прямых линий V
    DD1 = CInt(SS1)
    DD2 = CInt(SS2)
    V = DD1 + DD2

    Rem Провести горизонтальные красные(1) прямые линии
    Rem (количество прямых линий совпадает с числом десятков в первом двузначном числе)
    a = 0
    For J = 0 To D1 - 1
        a = a + 1
        For I = 0 To V - 1
            Cells(2 + J, 6 + I).Value = 1
        Next I
    Next J

    Rem Ниже нарисовать горизонтальные синие(2) прямые линии
    Rem (число таких прямых линий равно количеству единиц в первом двузначном числе)
    For J = 0 To D2 - 1
        For I = 0 To V - 1
            Cells(2 + a + J, 6 + I).Value = 2
        Next I
    Next J

    Rem Провести вертикальные красные(1) прямые линии
    Rem (число красных прямых линий - это количество десятков во втором двузначном числе)
    Rem параллельно мы складываем 1 с имеющимися в ячейках числами
    b = 0
    For J = 0 To DD1 - 1
        b = b + 1
        For I = 0 To G - 1
            E = Val(Cells(2 + I, 6 + J).Value)
            Cells(2 + I, 6 + J).Value = E + 1
        Next I
    Next J

    Rem Правее нарисовать синие вертикальные прямые линии
    Rem (количество таких прямых линий совпадает с числом единиц второго двузначного числа)
    For J = 0 To DD2 - 1
        For I = 0 To G - 1
            E = Val(Cells(2 + I, 6 + b + J).Value)
            Cells(2 + I, 6 + b + J).Value = E + 2
        Next I
    Next J

    Rem Считываем данные из таблицы и считаем результат умножения
    f = 0
    ff = 0
    fff = 0

    For I = 0 To G - 1
        For J = 0 To V - 1
            D = Val(Cells(2 + I, 6 + J).Value)
            If D = 2 Then f = f + 1
            If D = 3 Then ff = ff + 1
            If D = 4 Then fff = fff + 1
        Next J
    Next I

    Rem выкладываем результат в ячейку C3
    x = f * 100 + ff * 10 + fff
    Cells(3, 3).Value = x
End Sub
