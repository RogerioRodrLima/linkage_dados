Attribute VB_Name = "mdlCalcSimilar"
Function SimilaridadeStrings(str1, str2)

'Retira espa��s e coloca em letras maiusculas
str1 = Trim(UCase(str1))
str2 = Trim(UCase(str2))

'Verifica se ambas as strings s�o iguais,
'caso positivo, retorna 100% para a fun��o e sai do procedimento
If str1 = str2 Then SimilaridadeStrings = 1: Exit Function

'Pega a quantidade de caracteres das strings
NumCaractStr1 = Len(str1)
NumCaractStr2 = Len(str2)

'Verifica qual a maior string e seta as variaveis
If NumCaractStr1 >= NumCaractStr2 Then
    NumCaractStrMaior = NumCaractStr1
    strMaior = str1
    strMenor = str2
Else
    NumCaractStrMaior = NumCaractStr2
    strMaior = str2
    strMenor = str1
End If

'Reseta a variavel SeqCaractIguais (Sequ�ncia de Caracteres Iguais)
'Essa variavel receber� o quantidade m�xima de caracteres identicos entre as strings strMaior e strMenor
SeqCaractIguais = 0

'Variavel que receber� o quantidade de caracteres da strMaior para ser dividida no fim e encontrado a porcentagem
TotalCaracter = NumCaractStrMaior

'------------------------------------------------------------------------------------------------------------------
'ESSE PROCEDIMENTO EXECUTA O DESMEMBRAMENTO DAS STRINGS EM SEQUENCIAS DO MAIOR PARA O MENOR
'PARA ENCONTRAR A MAIOR SEUQENCIA DE CARACTERES IDENTICAS ENTRE AS DUAS STRINGS
'------------------------------------------------------------------------------------------------------------------

'Faz um looping do Maior Numero de caracteres para o menor para total de caracteres da Fun��o MID
For i = NumCaractStrMaior To 3 Step -1
    'S� executa enqaunto a quantidade de caracteres que sobram nas strings for maior que 2
    If Len(strMaior) >= 2 Then
        'Faz um looping para deslocar a posi��o inicial da contagem para a fun��o MID
        For strInicial = 1 To NumCaractStrMaior
            'Esta variavel recebe uma Substring da strMaior para verifica��o se a mesma � identica em ambas as strings iniciais
            StrMaiorParcial = Mid(strMaior, strInicial, i - 1)
            'Verifica se a Substring anterior � menor que o numero de caracteres analisado,
            'caso seja, reseta a variavel strMaiorPaarcial com o valor armazenado na strMaior
            'e sai do la�o For para reiniciar com uma quantidade de caracter N-1
            If Len(StrMaiorParcial) < i - 1 Then StrMaiorParcial = strMaior: Exit For
            'Debug.Print strMaior & " - " & strMenor & " - " & StrMaiorParcial
            'Se a maior SubString da strMaior for encontrada na StrMenor,
            'Retira a Substring das duas Strings Iniciais,
            'Seta a variavel NumCaractStrMaior com o novo comprimento da strMaior e
            'Seta a variavel SeqCaractIguais com o valor ja armazenado nela somado ao numero de caracteres da Substring
            'para ser dividido depois pelo total da string maior e sai do la�o for reiniciando o processo
            If InStr(strMenor, StrMaiorParcial) > 0 Then
                strMaior = Replace(strMaior, StrMaiorParcial, "")
                strMenor = Replace(strMenor, StrMaiorParcial, "")
                NumCaractStrMaior = Len(strMaior)
                SeqCaractIguais = SeqCaractIguais + Len(StrMaiorParcial)
                'Debug.Print strMaior & " - " & strMenor & " - " & NumCaractStrMaior
                Exit For
            End If
        Next strInicial
    Else
        Exit For
    End If
Next i
'------------------------------------------------------------------------------------------------------------------

'Divide a soma dos caracteres iguais entre as duas strings e divide pelo
'total de caracteres da strMaior
SimilaridadeStrings = SeqCaractIguais / TotalCaracter

End Function

Function CalcSimilar(str1, str2)

'Calculo da Similaridade entre as Strings Completas
SimilarCompleteString = SimilaridadeStrings(str1, str2)


'Decomposi��o do nome em array
ArrayStr1 = Split(str1, " ")
ArrayStr2 = Split(str2, " ")

'Extra��o da Primeira Palavra da String 1
FirstString1 = ArrayStr1(0)
Debug.Print FirstString1

'Extra��o da Primeira Palavra da String 2
FirstString2 = ArrayStr2(0)
Debug.Print FirstString2

'Calculo da Similaridade entre as Primeiras Palavras da Strings
SimilarFirstString = SimilaridadeStrings(FirstString1, FirstString2)
Debug.Print SimilarFirstString

'Extra��o da Ultima Palavra da String 1
For i = LBound(ArrayStr1) To UBound(ArrayStr1)
    LastString1 = ArrayStr1(i)
Next i
Debug.Print LastString1

'Extra��o da Ultima Palavra da String 2
For i = LBound(ArrayStr2) To UBound(ArrayStr2)
    LastString2 = ArrayStr2(i)
Next i
Debug.Print LastString2

'Calculo da Similaridade entre as Ultimas Palavras da Strings
SimilarLastString = SimilaridadeStrings(LastString1, LastString2)
Debug.Print SimilarLastString


CalcSimilar = (SimilarCompleteString * 6 + SimilarFirstString * 3 + SimilarLastString) / 10

End Function


