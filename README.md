' Эльвин Ахмедов 08.02.2025

Sub ReplaceCyrToLat()
    Dim ws As Worksheet
    Dim cell As Range
    Dim Target As String, Source As String
    Dim i As Integer
    
    ' Устанавливаем активный лист
    Set ws = ActiveSheet
    
    ' Определяем целевые символы (латиница)
    Target = ChrW(65) + ChrW(66) + ChrW(67) + ChrW(199) + ChrW(68) + ChrW(69) + ChrW(399) + ChrW(70) + _
             ChrW(71) + ChrW(286) + ChrW(72) + ChrW(88) + ChrW(73) + ChrW(304) + ChrW(74) + ChrW(75) + _
             ChrW(81) + ChrW(76) + ChrW(77) + ChrW(78) + ChrW(79) + ChrW(214) + ChrW(80) + ChrW(82) + _
             ChrW(83) + ChrW(350) + ChrW(84) + ChrW(85) + ChrW(220) + ChrW(86) + ChrW(89) + ChrW(90) + _
             ChrW(97) + ChrW(98) + ChrW(99) + ChrW(231) + ChrW(100) + ChrW(101) + ChrW(601) + ChrW(102) + _
             ChrW(103) + ChrW(287) + ChrW(104) + ChrW(120) + ChrW(305) + ChrW(105) + ChrW(106) + _
             ChrW(107) + ChrW(113) + ChrW(108) + ChrW(109) + ChrW(110) + ChrW(111) + ChrW(246) + _
             ChrW(112) + ChrW(114) + ChrW(115) + ChrW(351) + ChrW(116) + ChrW(117) + ChrW(252) + _
             ChrW(118) + ChrW(121) + ChrW(122)
             
    ' Определяем исходные символы (кириллица)
    Source = ChrW(1040) + ChrW(1041) + ChrW(1066) + ChrW(1063) + ChrW(1044) + ChrW(1045) + _
             ChrW(1071) + ChrW(1060) + ChrW(1069) + ChrW(1068) + ChrW(1065) + ChrW(1061) + _
             ChrW(1067) + ChrW(1048) + ChrW(1046) + ChrW(1050) + ChrW(1043) + ChrW(1051) + _
             ChrW(1052) + ChrW(1053) + ChrW(1054) + ChrW(1070) + ChrW(1055) + ChrW(1056) + _
             ChrW(1057) + ChrW(1064) + ChrW(1058) + ChrW(1059) + ChrW(1062) + ChrW(1042) + _
             ChrW(1049) + ChrW(1047) + ChrW(1072) + ChrW(1073) + ChrW(1098) + ChrW(1095) + _
             ChrW(1076) + ChrW(1077) + ChrW(1103) + ChrW(1092) + ChrW(1101) + ChrW(1100) + _
             ChrW(1097) + ChrW(1093) + ChrW(1099) + ChrW(1080) + ChrW(1078) + ChrW(1082) + _
             ChrW(1075) + ChrW(1083) + ChrW(1084) + ChrW(1085) + ChrW(1086) + ChrW(1102) + _
             ChrW(1087) + ChrW(1088) + ChrW(1089) + ChrW(1096) + ChrW(1090) + ChrW(1091) + _
             ChrW(1094) + ChrW(1074) + ChrW(1081) + ChrW(1079)

    ' Перебираем все ячейки с текстом
    For Each cell In ws.UsedRange
        If Not IsEmpty(cell.Value) Then
            ' Проходим по каждому символу и заменяем
            For i = 1 To Len(Source)
                cell.Value = Replace(cell.Value, Mid(Source, i, 1), Mid(Target, i, 1))
            Next i
        End If
    Next cell
    
    MsgBox "Замена букв завершена!", vbInformation
End Sub
