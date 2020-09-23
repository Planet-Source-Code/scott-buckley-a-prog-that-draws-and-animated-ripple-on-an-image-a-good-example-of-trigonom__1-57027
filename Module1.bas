Attribute VB_Name = "Module1"
Public Function Within(inp, min, max)
    inp2 = inp
    If inp2 < min Then inp2 = min
    If inp2 > max Then inp2 = max
    Within = inp2
End Function
