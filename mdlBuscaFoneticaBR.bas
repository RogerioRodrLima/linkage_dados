Attribute VB_Name = "mdlBuscaFoneticaBR"
Function BuscaFoneticaBR(Palavra)

'Converte para maiusculo
Palavra = UCase(Palavra)

'Remove acentos
Palavra = Replace(Palavra, "Á", "A")
Palavra = Replace(Palavra, "À", "A")
Palavra = Replace(Palavra, "Â", "A")
Palavra = Replace(Palavra, "Ã", "A")

Palavra = Replace(Palavra, "É", "E")
Palavra = Replace(Palavra, "È", "E")
Palavra = Replace(Palavra, "Ê", "E")

Palavra = Replace(Palavra, "Í", "I")
Palavra = Replace(Palavra, "Ì", "I")
Palavra = Replace(Palavra, "Î", "I")

Palavra = Replace(Palavra, "Ó", "O")
Palavra = Replace(Palavra, "Ò", "O")
Palavra = Replace(Palavra, "Ô", "O")
Palavra = Replace(Palavra, "Õ", "O")

Palavra = Replace(Palavra, "Ú", "U")
Palavra = Replace(Palavra, "Ù", "U")
Palavra = Replace(Palavra, "Û", "U")

'Eliminamos hifen
Palavra = Replace(Palavra, "-", "")

'Substitui apostrofo por espaço
Palavra = Replace(Palavra, "'", " ")

'Remove artigos, preposições e conjunções das frases
Palavra = Replace(Palavra, " A ", " ")
Palavra = Replace(Palavra, " E ", " ")
Palavra = Replace(Palavra, " O ", " ")
Palavra = Replace(Palavra, " NA ", " ")
Palavra = Replace(Palavra, " NO ", " ")
Palavra = Replace(Palavra, " NAS ", " ")
Palavra = Replace(Palavra, " NOS ", " ")
Palavra = Replace(Palavra, " DA ", " ")
Palavra = Replace(Palavra, " DE ", " ")
Palavra = Replace(Palavra, " DO ", " ")
Palavra = Replace(Palavra, " DAS ", " ")
Palavra = Replace(Palavra, " DOS ", " ")
Palavra = Replace(Palavra, " EM ", " ")

PalavraLimpa = ""

'Eliminamos todas as letras em duplicidade e espaços
TamPalavra = Len(Palavra)

For i = 1 To TamPalavra
    Letra = Mid(Palavra, i, 1)
    LetraSeguinte = Mid(Palavra, i + 1, 1)
    
    If LetraSeguinte <> Letra Then
        PalavraLimpa = PalavraLimpa & Letra
    End If
Next i

'Seta a Palavra com a PalavraLimpa (sem duplicidades)
Palavra = PalavraLimpa

'Divide a Frase
Frase = Split(Palavra, " ")

    x = 0

    For x = 0 To UBound(Frase) Step 1
    
        Palavra = Frase(x)
    
        'Substituimos Y por I
        Palavra = Replace(Palavra, "Y", "I")
        
        'Substituimos PH por F
        Palavra = Replace(Palavra, "PH", "F")
        
        'Substituimos CHR por CR
        Palavra = Replace(Palavra, "CHR", "CR")
        
        'Substituimos CHIO por QUIO
        Palavra = Replace(Palavra, "CHIO", "QUIO")
        
        'Substituimos CÇ por S
        Palavra = Replace(Palavra, "CÇ", "S")
        
        'Substituimos CT por T
        Palavra = Replace(Palavra, "CT", "T")
        Palavra = Replace(Palavra, "CS", "S")
        
        'Substituimos mb...mz, nb...nz tirando o n ou m
        Palavra = Replace(Palavra, "MB", "B")
        Palavra = Replace(Palavra, "MC", "C")
        Palavra = Replace(Palavra, "MÇ", "S")
        Palavra = Replace(Palavra, "MD", "D")
        Palavra = Replace(Palavra, "MF", "F")
        Palavra = Replace(Palavra, "MG", "G")
        Palavra = Replace(Palavra, "MJ", "J")
        Palavra = Replace(Palavra, "MK", "K")
        Palavra = Replace(Palavra, "ML", "L")
        Palavra = Replace(Palavra, "MP", "P")
        Palavra = Replace(Palavra, "MQ", "Q")
        Palavra = Replace(Palavra, "MR", "R")
        Palavra = Replace(Palavra, "MS", "S")
        Palavra = Replace(Palavra, "MT", "T")
        Palavra = Replace(Palavra, "MV", "V")
        Palavra = Replace(Palavra, "VW", "W")
        Palavra = Replace(Palavra, "MX", "X")
        Palavra = Replace(Palavra, "MZ", "Z")
        
        Palavra = Replace(Palavra, "NB", "B")
        Palavra = Replace(Palavra, "NC", "C")
        Palavra = Replace(Palavra, "NÇ", "S")
        Palavra = Replace(Palavra, "ND", "D")
        Palavra = Replace(Palavra, "NF", "F")
        Palavra = Replace(Palavra, "NG", "G")
        Palavra = Replace(Palavra, "NJ", "J")
        Palavra = Replace(Palavra, "NK", "K")
        Palavra = Replace(Palavra, "NL", "L")
        Palavra = Replace(Palavra, "NQ", "Q")
        Palavra = Replace(Palavra, "NP", "P")
        Palavra = Replace(Palavra, "NR", "R")
        Palavra = Replace(Palavra, "NS", "S")
        Palavra = Replace(Palavra, "NT", "T")
        Palavra = Replace(Palavra, "NV", "V")
        Palavra = Replace(Palavra, "VW", "W")
        Palavra = Replace(Palavra, "NX", "X")
        Palavra = Replace(Palavra, "NZ", "Z")
        
        'Substituimos sb...sz tirando o s
        Palavra = Replace(Palavra, "SB", "B")
        Palavra = Replace(Palavra, "SC", "C")
        Palavra = Replace(Palavra, "SÇ", "C")
        Palavra = Replace(Palavra, "SD", "D")
        Palavra = Replace(Palavra, "SF", "F")
        Palavra = Replace(Palavra, "SG", "G")
        Palavra = Replace(Palavra, "SJ", "J")
        Palavra = Replace(Palavra, "SK", "K")
        Palavra = Replace(Palavra, "SL", "L")
        Palavra = Replace(Palavra, "SM", "M")
        Palavra = Replace(Palavra, "SN", "N")
        Palavra = Replace(Palavra, "SP", "P")
        Palavra = Replace(Palavra, "SQ", "Q")
        Palavra = Replace(Palavra, "SR", "R")
        Palavra = Replace(Palavra, "ST", "T")
        Palavra = Replace(Palavra, "SV", "V")
        Palavra = Replace(Palavra, "SW", "W")
        Palavra = Replace(Palavra, "SX", "X")
        Palavra = Replace(Palavra, "SZ", "Z")
        
        'Substituimos BR...ZR e BL...BZ por B...Z
        Palavra = Replace(Palavra, "BR", "B")
        Palavra = Replace(Palavra, "BL", "B")
        Palavra = Replace(Palavra, "CR", "C")
        Palavra = Replace(Palavra, "CL", "C")
        Palavra = Replace(Palavra, "DR", "D")
        Palavra = Replace(Palavra, "DL", "D")
        Palavra = Replace(Palavra, "FR", "F")
        Palavra = Replace(Palavra, "FL", "F")
        Palavra = Replace(Palavra, "GR", "G")
        Palavra = Replace(Palavra, "GL", "G")
        Palavra = Replace(Palavra, "KR", "K")
        Palavra = Replace(Palavra, "KL", "K")
        Palavra = Replace(Palavra, "MR", "M")
        Palavra = Replace(Palavra, "PR", "P")
        Palavra = Replace(Palavra, "PL", "P")
        Palavra = Replace(Palavra, "TR", "T")
        Palavra = Replace(Palavra, "TL", "T")
        Palavra = Replace(Palavra, "VR", "V")
        Palavra = Replace(Palavra, "VL", "V")
        Palavra = Replace(Palavra, "WR", "W")
        Palavra = Replace(Palavra, "WL", "W")
        Palavra = Replace(Palavra, "ZR", "Z")
        Palavra = Replace(Palavra, "ZL", "Z")
        
        'Substituimos palavras com consoante muda por sua similar fonetica
        strCons1 = Split("B C D F G P Q S T")
        strCons2 = Split("B C Ç D F G J M N P Q S T V W X Z")
        
        For intLoop1 = 0 To UBound(strCons1)
            For intLoop2 = 0 To UBound(strCons2)
                strJuncao = strCons1(intLoop1) & strCons2(intLoop2)
                If strCons1(intLoop1) = "G" Or strCons1(intLoop1) = "Q" Then VogalAcresc = "UI" Else VogalAcresc = "I"
                strJuncaoSubst = strCons1(intLoop1) & VogalAcresc & strCons2(intLoop2)
                If InStr(Palavra, strJuncao) Then
                    Palavra = Replace(Palavra, strJuncao, strJuncaoSubst)
                End If
            Next intLoop2
        Next intLoop1

        
        'Substituimos GR, MG, NG, RG por G
        Palavra = Replace(Palavra, "MG", "G")
        Palavra = Replace(Palavra, "RG", "G")
        
        'Substituimos GE, GI, RJ, MJ, NJ por J
        Palavra = Replace(Palavra, "GE", "JE")
        Palavra = Replace(Palavra, "GI", "JI")
        Palavra = Replace(Palavra, "RJ", "J")
        Palavra = Replace(Palavra, "GU", "J")
        Palavra = Replace(Palavra, "MJ", "J")
        Palavra = Replace(Palavra, "NJ", "J")
        Palavra = Replace(Palavra, "GR", "G")
        Palavra = Replace(Palavra, "GL", "G")
        
        'Substituimos CE, CI e CH por S
        Palavra = Replace(Palavra, "CE", "SE")
        Palavra = Replace(Palavra, "CI", "SI")
        Palavra = Replace(Palavra, "CH", "S")
        
        'Substituimos Q, QU, CA, CO, CU, C por K;
        Palavra = Replace(Palavra, "QU", "K")
        Palavra = Replace(Palavra, "Q", "K")
        Palavra = Replace(Palavra, "CA", "KA")
        Palavra = Replace(Palavra, "CO", "KO")
        Palavra = Replace(Palavra, "CU", "KU")
        Palavra = Replace(Palavra, "CK", "K")
        Palavra = Replace(Palavra, "C", "K")
        
        'Substituimos LH por L
        Palavra = Replace(Palavra, "LH", "L")
        Palavra = Replace(Palavra, "RM", "SM")
        
        'Substituimos N, RM, GM, MD, SM e Terminação AO por AM
        Palavra = Replace(Palavra, "N", "M")
        Palavra = Replace(Palavra, "RM", "M")
        Palavra = Replace(Palavra, "GM", "M")
        Palavra = Replace(Palavra, "MD", "M")
        Palavra = Replace(Palavra, "SM", "M")
        
        If Right(Palavra, 2) = "AO" Then
            TamPalavra = Len(Palavra)
            Palavra = Left(Palavra, TamPalavra - 2) & "AM"
        End If
        
        'Substituimos AO por AM PARA PEGAR FRASES
        Palavra = Replace(Palavra, "AO", "AM")
        
        'Substituimos AL, EL, IL, OL por AU, EU, IU, OU, SE VIER SEGUIDA DE CONSOANTE
        strCons3 = Split("B C Ç D F G J K M N P Q R S T V W X Z")
        Contador = 0
        
        For intLoop3 = 0 To UBound(strCons3)
            If InStr(Palavra, "AL" & strCons3(intLoop3)) <> 0 _
            Or InStr(Palavra, "EL" & strCons3(intLoop3)) <> 0 _
            Or InStr(Palavra, "IL" & strCons3(intLoop3)) <> 0 _
            Or InStr(Palavra, "OL" & strCons3(intLoop3)) <> 0 Then
                Contador = Contador + 1
                ConsAchada = strCons3(intLoop3)
            End If
        Next intLoop3
        
        
        If Contador <> 0 Then
            Palavra = Replace(Palavra, "AL" & ConsAchada, "AU" & ConsAchada)
            Palavra = Replace(Palavra, "EL" & ConsAchada, "EU" & ConsAchada)
            Palavra = Replace(Palavra, "IL" & ConsAchada, "IU" & ConsAchada)
            Palavra = Replace(Palavra, "OL" & ConsAchada, "OU" & ConsAchada)
        End If
        
        'Substituimos as terminações AL, EL, IL, OL por AU, EU, IU, OU
        TamPalavra = Len(Palavra)
        If Right(Palavra, 2) = "AL" Then Palavra = Left(Palavra, TamPalavra - 2) & "AU"
        If Right(Palavra, 2) = "EL" Then Palavra = Left(Palavra, TamPalavra - 2) & "EU"
        If Right(Palavra, 2) = "IL" Then Palavra = Left(Palavra, TamPalavra - 2) & "IU"
        If Right(Palavra, 2) = "OL" Then Palavra = Left(Palavra, TamPalavra - 2) & "OU"
            
        'Substituimos NH por N;
        Palavra = Replace(Palavra, "NH", "N")
        
        'Substituimos PR por P;
        Palavra = Replace(Palavra, "PR", "P")
        
        'Substituimos Ç, X, TS, C, Z, RS por S
        Palavra = Replace(Palavra, "Ç", "S")
        Palavra = Replace(Palavra, "X", "S")
        'Palavra = Replace(Palavra, "TS", "S")
        Palavra = Replace(Palavra, "C", "S")
        Palavra = Replace(Palavra, "Z", "S")
        Palavra = Replace(Palavra, "RS", "S")
        
        
        'Substituimos LT, TR, CT, RT, ST por T
        Palavra = Replace(Palavra, "TR", "T")
        Palavra = Replace(Palavra, "TL", "T")
        Palavra = Replace(Palavra, "LT", "T")
        Palavra = Replace(Palavra, "RT", "T")
        Palavra = Replace(Palavra, "ST", "T")
        
        'Substituimos U por O
        Palavra = Replace(Palavra, "U", "O")
        
        'Substituimos W por V;
        Palavra = Replace(Palavra, "W", "V")
        
        'Substituimos L por R
        Palavra = Replace(Palavra, "L", "R")
        
        'Eliminamos H
        Palavra = Replace(Palavra, "H", "")
        
        'Substituimos I por E
        Palavra = Replace(Palavra, "I", "E")
        
        'Eliminamos as terminações S, Z, R, R, M, N, AO e L
        TamPalavra = Len(Palavra)
        
        Fim1 = Right(Palavra, 1)
        Fim2 = Right(Palavra, 2)
        
        If Fim2 = "AO" Then
            Palavra = Left(Palavra, TamPalavra - 2)
        End If
        
        If Fim1 = "S" Or Fim1 = "Z" Or Fim1 = "R" Or Fim1 = "M" Or Fim1 = "N" Or Fim1 = "L" Or Fim1 = "T" Or Fim1 = "U" Or Fim1 = "D" Then
            Palavra = Left(Palavra, TamPalavra - 1)
        End If
        
        If Fim1 = "K" Then
            Palavra = Replace(Palavra, "K", "KE")
        End If
        
        'Elimina as vogais
        'Palavra = Replace(Palavra, "A", "")
        'Palavra = Replace(Palavra, "E", "")
        'Palavra = Replace(Palavra, "I", "")
        'Palavra = Replace(Palavra, "O", "")
        'Palavra = Replace(Palavra, "U", "")
        
        PalavraLimpa = ""
        
        'Eliminamos todas as letras em duplicidade
        TamPalavra = Len(Palavra)
        
        For i = 1 To TamPalavra
            Letra = Mid(Palavra, i, 1)
            LetraSeguinte = Mid(Palavra, i + 1, 1)
            
            If LetraSeguinte <> Letra Then
                PalavraLimpa = PalavraLimpa & Letra
            End If
        Next i
        
        'Seta a Palavra com a PalavraLimpa (sem duplicidades)
        Palavra = PalavraLimpa
        
        PalavraFinal = PalavraFinal & IIf(x <> 0, " ", "") & Palavra
    
    Next x

'Retorna o resultado da função
BuscaFoneticaBR = Trim(PalavraFinal)

End Function
