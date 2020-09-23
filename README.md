<div align="center">

## DiferenciaEnFechas \-\-\>Difference between Dates


</div>

### Description

Calculate the difference between dates and return it in Age Format

Ex. xx Years, yy Months, zz Days

It works for both dates future and past.

Ex.

A)My age

DiferenciaEnFechas(Now,MyBornDate)

B)Next year(01/01/2000)

DiferenciaEnFechas(12/08/1999,01/01/2000)-->Futuro: 0 Años,4 Meses,20 Dias
 
### More Info
 
'1.- pdFechaBase As Date --> Is the base date (Start point)

'2.- pdFecha As Date --> Is the date that you want to know the difference

'Return a String (in Spanish)

'Ex. DiferenciaEnFechas(12/08/1999,01/01/2000)

'Return ---> Futuro: 0 Años,4 Meses,20 Dias

'You can translate to English:

'Futuro = Future

'Hoy = Today

'Pasado = Past

'Año/Años = Year/Years

'Mes/Meses = Month/Months

'Día/Dias = Day/Days


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Ricardo Ortiz](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/ricardo-ortiz.md)
**Level**          |Unknown
**User Rating**    |4.7 (14 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Math/ Dates](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/math-dates__1-37.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/ricardo-ortiz-diferenciaenfechas-difference-between-dates__1-2960/archive/master.zip)





### Source Code

```
Function DiferenciaEnFechas(pdFechaBase As Date, pdFecha As Date) As String
'******************************************************
'* Autor : Ricardo Ortiz
'* Ultima Modificación: 17/08/1999
'******************************************************
Dim dFechaAux As Date
Dim iYear As Integer, iMes As Integer, iDia As Integer
Dim iYearFinal As Integer
Dim iMesFinal As Integer
Dim iDiaFinal As Integer
Dim sTiempo As String, sAux As String
  iDia = DatePart("d", pdFecha)
  iMes = Month(pdFechaBase)
  iYear = Year(pdFechaBase)
  dFechaAux = DateSerial(iYear, iMes, iDia)
  iDiaFinal = DateDiff("d", dFechaAux, pdFechaBase)
  iMes = DateDiff("m", pdFecha, pdFechaBase)
  Select Case iMes
   Case Is > 0  'Pasado
     iYearFinal = iMes \ 12
     iMesFinal = iMes Mod 12
     If iDiaFinal < 0 Then
      If Month(dFechaAux) <> Month(pdFechaBase) Then 'Caso Raro
        iDiaFinal = 31 - (DatePart("d", DateAdd("d", -1, DateSerial(iYear, Month(dFechaAux), 1))))
        dFechaAux = DateAdd("m", -1, dFechaAux)
        dFechaAux = DateAdd("d", -iDiaFinal, dFechaAux)
      Else                      'Caso Normal
        dFechaAux = DateAdd("m", -1, dFechaAux)
      End If
      iDiaFinal = DateDiff("d", dFechaAux, pdFechaBase)
      If iMesFinal > 0 Then
        iMesFinal = iMesFinal - 1
      Else
        If iYearFinal > 0 Then
         iYearFinal = iYearFinal - 1
         iMesFinal = 11
        End If
      End If
     End If
     sTiempo = "Pasado: "
   Case Is = 0
     iYearFinal = 0
     iMesFinal = 0
     If iDiaFinal < 0 Then    'Futuro
      iDiaFinal = DateDiff("d", pdFechaBase, dFechaAux)
      sTiempo = "Futuro: "
     ElseIf iDiaFinal = 0 Then  'HOY
      sTiempo = "HOY: "
     Else             'Pasado
      sTiempo = "Pasado: "
     End If
   Case Else     'Futuro
     iMes = DateDiff("m", pdFechaBase, pdFecha)
     iYearFinal = iMes \ 12
     iMesFinal = iMes Mod 12
     If iDiaFinal > 0 Then
      dFechaAux = DateAdd("m", 1, dFechaAux)
      iDiaFinal = DateDiff("d", pdFechaBase, dFechaAux)
      If iMesFinal > 0 Then
        iMesFinal = iMesFinal - 1
      Else
        If iYearFinal > 0 Then
         iYearFinal = iYearFinal - 1
         iMesFinal = 11
        End If
      End If
     Else
      iDiaFinal = DateDiff("d", pdFechaBase, dFechaAux)
     End If
     sTiempo = "Futuro: "
  End Select
  sAux = Str(iYearFinal)
  If iYearFinal = 1 Then
   sAux = sAux & " Año, "
  Else
   sAux = sAux & " Años, "
  End If
  sAux = sAux & Str(iMesFinal)
  If iMesFinal = 1 Then
   sAux = sAux & " Mes, "
  Else
   sAux = sAux & " Meses, "
  End If
  sAux = sAux & Str(iDiaFinal)
  If iDiaFinal = 1 Then
   sAux = sAux & " Día"
  Else
   sAux = sAux & " Dias"
  End If
  DiferenciaEnFechas = sTiempo & sAux
End Function
```

