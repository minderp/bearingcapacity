Attribute VB_Name = "StructuralCapacity"
Dim VersionStringG As String
Dim initializedState As String
Public Function VersionG()
VersionStringG = "Version 0.02, 2018-01-20"
VersionG = VersionStringG
End Function

Public Sub testfunction()
Dim xPosition, xEnd, b, EI, MeModul As Double
Dim xFi_Range, Fi_Range As Range

xPosition = 0    'Auswertestelle
xEnd = 8#        'm, Länge
b = 2#            'm, Breite
EI = 5625# * 1000#   'kPa für starres Fundament
'EI = 13.33 * 1000 'kPa für schlaffes Fundament

MeModul = 50# * 1000#   ' kPa
'die Ranges kann ich im Testfile nicht dynamisch definieren.
'sie verweisen immer auf die unten angegebenen Zellen

Set xFi_Range = Workbooks(1).Sheets(1).Range("F6:F8")
Set Fi_Range = Workbooks(1).Sheets(1).Range("G6:G8")

'testaufruf für STV
wert_STV = Moment(xPosition, xEnd, b, EI, MeModul, xFi_Range, xFi_Range, True)
'testaufruf für BMV
wert_BMV = Moment(xPosition, xEnd, b, EI, MeModul, xFi_Range, Fi_Range, False)
wert2_BMV = Biegelinie(xPosition, xEnd, b, EI, MeModul, xFi_Range, Fi_Range)
'wert3_BMV = Max_Moment(xEnd, b, EI, MeModul, xFi_Range, Fi_Range)
'wert3_BMV = Min_Moment(xEnd, b, EI, MeModul, xFi_Range, Fi_Range)
'endetest

I = 1
End Sub
Public Function Moment(ByVal xPosition As Variant, _
ByVal xEnd As Variant, _
ByVal b As Variant, _
ByVal EI As Variant, _
ByVal MeModul As Variant, _
ByVal xFi_Range As Range, _
ByVal Fi_Range As Range, _
Optional ByVal ModeSTV As Boolean = False _
) As Variant
Attribute Moment.VB_Description = "Berechnet das Moment im Fundament an der angegebenen Stelle in kNm/m'"
Attribute Moment.VB_ProcData.VB_Invoke_Func = " \n14"

'Input parameter
'Position, an der die Beanspruchung ausgewertet werden soll (xPosition)
'Länge des Balken (xEnd) ‘Annahme xStart=0
'Steifigkeit des Balkens (Breite, EI)
'Steifigkeit des Bodens (ME)
'Position der Kräfte (xFi) Vektor (min 1)
'Betrag der Kräfte (Fi) Vektor (min 1)
'Modus ob zwingend STV oder auch BMV (true=immer STV)
'
'Ausgabe ist in kNm/m
'
'(Option: Prüfung auf zulässigen Input, Ausgabe von Fehlermeldungen Exit Function)
'--- Inputdaten bereinigen-------------------------------------------------------
Dim xF() As Variant ' declare an unallocated array.
Dim F() As Variant  ' declare an unallocated array.
Dim nloads As Integer

'--- Check Loads for consistency and define the vectors dynamically
'n=1 or more?
If xFi_Range.Rows.Count = 1 And xFi_Range.Columns.Count = 1 Then
 ReDim xF(1 To 1) As Variant
 xF(1) = xFi_Range.Cells(1, 1).Value
Else
 xF = Application.Transpose(xFi_Range)
End If

If Fi_Range.Rows.Count = 1 And Fi_Range.Columns.Count = 1 Then
 ReDim F(1 To 1) As Variant
 F(1) = Fi_Range.Cells(1, 1).Value
Else
 F = Application.Transpose(Fi_Range)
End If
'--- equal size of input vectors
nxF = UBound(xF)
nF = UBound(F)

If nxF <> nF Then
  Moment = "Different range sizes"
  Exit Function
 Else
 nloads = nxF
End If
'aling orientation
If xFi_Range.Columns.Count > xFi_Range.Rows.Count Then
  xF = Application.Transpose(Application.Transpose(xFi_Range))
End If
If Fi_Range.Columns.Count > Fi_Range.Rows.Count Then
  F = Application.Transpose(Application.Transpose(Fi_Range))
End If
''--- Inputdaten provisorisch bereinigt -------------------
'----------------------------------------------------------
' en Bettungsmodul ks Aus ME und Breite d berechnen (Nach Näherungsformel Lang)
ks = MeModul / (fShape(b, xEnd) * b)
'(Option: ks aus Cc, Breite und Mächtigkeit der Schicht berechnen (Boussinesq+Layers))
'--------------------------------
'Fallunterscheidung für Verfahren
'Das Steifigkeitsverhältnis (Elastische Länge L) im Vergleich zur Fundamentlänge dient als Grundlage.
'Für xEnd/L<2 ist das STV eine ausreichend gute Approximation, ansonsten ist das BMV zu verwenden.
' ! Wichtiger Hinweis: Für den Testbetrieb wird zurzeit nur wenn Mode=true STV verwendet, sonst immer BMV. ----------------------
L = (4 * EI / (ks * b)) ^ 0.25
If ModeSTV Then
    Moment = STV_M(xPosition, xEnd, xF, F, nloads) / b
Else
    Moment = BMV_M(xPosition, xEnd, xF, F, nloads, L) / (b) 'eher schlaff. Bettungsmodulverfahren benutzen
    'Bei zu schlaffen Fundamenten ist auch das Bettungsmodulverfahren schlecht.
    'Dann sollten eigentlich die Lasten direkt auf den Baugrund gegeben werden (z.B. Boussinesq mit Einzellaten)
    'Zur Bestimmung der inneren Beanspruchung reicht es aber aus (die Momente werden sehr klein)
End If

End Function 'Moment() Ausgabe ist in kNm/m
Public Function Max_Moment(ByVal xEnd As Variant, _
ByVal b As Variant, _
ByVal EI As Variant, _
ByVal MeModul As Variant, _
ByVal xFi_Range As Range, _
ByVal Fi_Range As Range, _
Optional ByVal xBereich_Start, Optional ByVal xBereich_Ende)

'Input parameter
'Länge des Balken (xEnd) ‘Annahme xStart=0
'Steifigkeit des Balkens (Breite, EI)
'Steifigkeit des Bodens (ME)
'Position der Kräfte (xFi) Vektor (min 1)
'Betrag der Kräfte (Fi) Vektor (min 1)
'Bereiche über die ausgewertet werden soll
'
'Ausgabe ist in kNm/m
'
'Define characteristic points xchar= (0, xFi, xEnd) within interwall
Dim I As Integer, xchar() As Double
Const NumElements As Integer = 101
Dim Increment As Double

'provisorisch
    Increment = xEnd / (NumElements - 1)
    ReDim xchar(1 To NumElements)
    For I = 1 To NumElements
        xchar(I) = (I - 1) * Increment
    Next I
'TODO Limit auf Bereich
'ToDO einfügen von xFi, da dort oft min/max auftrit
'initialize
M = Moment(xchar(1), xEnd, b, EI, MeModul, xFi_Range, Fi_Range)
'Check Value at specific points and compare to current max
For Each x In xchar
Mnew = Moment(x, xEnd, b, EI, MeModul, xFi_Range, Fi_Range)
    If Mnew > M Then
        M = Mnew
    End If
Next x
'return function value
Max_Moment = M
End Function 'Max_Moment
Public Function Min_Moment(ByVal xEnd As Variant, _
ByVal b As Variant, _
ByVal EI As Variant, _
ByVal MeModul As Variant, _
ByVal xFi_Range As Range, _
ByVal Fi_Range As Range, _
Optional ByVal xBereich_Start, Optional ByVal xBereich_Ende)
'Input parameter
'Länge des Balken (xEnd) ‘Annahme xStart=0
'Steifigkeit des Balkens (Breite, EI)
'Steifigkeit des Bodens (ME)
'Position der Kräfte (xFi) Vektor (min 1)
'Betrag der Kräfte (Fi) Vektor (min 1)
'Bereiche über die ausgewertet werden soll
'
'Ausgabe ist in kNm/m
'
'Define characteristic points xchar= (0, xFi, xEnd) within interwall
Dim I As Integer, xchar() As Double
Const NumElements As Integer = 101
Dim Increment As Double
'provisorisch
    Increment = xEnd / (NumElements - 1)
    ReDim xchar(1 To NumElements)
    For I = 1 To NumElements
        xchar(I) = (I - 1) * Increment
    Next I
'TODO Limit auf Bereich
'ToDO einfügen von xFi, da dort oft min/max auftrit
'initialize
M = Moment(xchar(1), xEnd, b, EI, MeModul, xFi_Range, Fi_Range)
'Check Value at specific points and compare to current max
For Each x In xchar
Mnew = Moment(x, xEnd, b, EI, MeModul, xFi_Range, Fi_Range)
    If Mnew < M Then
        M = Mnew
    End If
Next x
'return function value
Min_Moment = M
End Function      'Max_Moment
Public Function Biegelinie(ByVal xPosition As Variant, _
ByVal xEnd As Variant, _
ByVal b As Variant, _
ByVal EI As Variant, _
ByVal MeModul As Variant, _
ByVal xFi_Range As Range, _
ByVal Fi_Range As Range) As Variant
Attribute Biegelinie.VB_Description = "Berechnet die Verschiebung im Fundament resp die Setzung des Baugrunds an der angegebenen Stelle in m"
Attribute Biegelinie.VB_ProcData.VB_Invoke_Func = " \n14"
'
'Input parameter
'Position, an der die Beanspruchung ausgewertet werden soll (xPosition)
'Länge des Balken (xEnd) ‘Annahme xStart=0
'Steifigkeit des Balkens (Breite, EI)
'Steifigkeit des Bodens (ME)
'Position der Kräfte (xFi) Vektor (min 1)
'Betrag der Kräfte (Fi) Vektor (min 1)
'
'Ausgabe ist in m

'(Option: Prüfung auf zulässigen Input, Ausgabe von Fehlermeldungen Exit Function)
'--- Inputdaten bereinigen-------------------------------------------------------
Dim xF() As Variant ' declare an unallocated array.
Dim F() As Variant  ' declare an unallocated array.
Dim nloads As Integer

'--- Check Loads for consistency and define the vectors dynamically
'n=1 or more?
If xFi_Range.Rows.Count = 1 And xFi_Range.Columns.Count = 1 Then
 ReDim xF(1 To 1) As Variant
 xF(1) = xFi_Range.Cells(1, 1).Value
Else
 xF = Application.Transpose(xFi_Range)
End If

If Fi_Range.Rows.Count = 1 And Fi_Range.Columns.Count = 1 Then
 ReDim F(1 To 1) As Variant
 F(1) = Fi_Range.Cells(1, 1).Value
Else
 F = Application.Transpose(Fi_Range)
End If
'--- equal size of input vectors
nxF = UBound(xF)
nF = UBound(F)

If nxF <> nF Then
  Biegelinie = "Error: Different range sizes"
  Exit Function
 Else
 nloads = nxF
End If
'aling orientation
If xFi_Range.Columns.Count > xFi_Range.Rows.Count Then
  xF = Application.Transpose(Application.Transpose(xFi_Range))
End If
If Fi_Range.Columns.Count > Fi_Range.Rows.Count Then
  F = Application.Transpose(Application.Transpose(Fi_Range))
End If
''--- Inputdaten provisorisch bereinigt -------------------
'----------------------------------------------------------
'
'Aus ME und Breite den Bettungsmodul ks berechnen (Mit Boussinesq+Annahmen)
ks = MeModul / (fShape(b, xEnd) * b)
'(Option: aus Cc, Breite und Mächtigkeit der Schicht ks berechnen (Boussinesq+Layers))
'--------------------------------
'Fallunterscheidung für Verfahren
'Das Spannungstrapezverfahren liefert keine Setzung nur die Biegelinie des Fundaments.
'Deshalb wird in allen Fällen das BMV verwendet.
L = (4 * EI / (ks * b)) ^ 0.25
If xEnd < 2 * L Then
    'Option Biegelinie des Balkens unter Trapezbelastung ausgeben
    'Biegelinie = STV_y((xPosition, xEnd, xF, F, nloads) / (EI) -----Option
    Biegelinie = BMV_Y(xPosition, xEnd, xF, F, nloads, L, EI)
Else
    Biegelinie = BMV_Y(xPosition, xEnd, xF, F, nloads, L, EI)
    
End If
End Function 'Biegelinie()
Public Function Querkraft(ByVal xPosition As Variant, _
ByVal xEnd As Variant, _
ByVal b As Variant, _
ByVal EI As Variant, _
ByVal MeModul As Variant, _
ByVal xFi_Range As Range, _
ByVal Fi_Range As Range) As Variant
Attribute Querkraft.VB_Description = "Berechnet die Querkraft im Fundament an der angegebenen Stelle in kN/m"
Attribute Querkraft.VB_ProcData.VB_Invoke_Func = " \n14"
'
'Input parameter
'Position, an der die Beanspruchung ausgewertet werden soll (xPosition)
'Länge des Balken (xEnd) ‘Annahme xStart=0
'Steifigkeit des Balkens (Breite, EI)
'Steifigkeit des Bodens (ME)
'Position der Kräfte (xFi) Vektor (min 1)
'Betrag der Kräfte (Fi) Vektor (min 1)
'
'Ausgabe ist in m

'(Option: Prüfung auf zulässigen Input, Ausgabe von Fehlermeldungen Exit Function)
'--- Inputdaten bereinigen-------------------------------------------------------
Dim xF() As Variant ' declare an unallocated array.
Dim F() As Variant  ' declare an unallocated array.
Dim nloads As Integer

'--- Check Loads for consistency and define the vectors dynamically
'n=1 or more?
If xFi_Range.Rows.Count = 1 And xFi_Range.Columns.Count = 1 Then
 ReDim xF(1 To 1) As Variant
 xF(1) = xFi_Range.Cells(1, 1).Value
Else
 xF = Application.Transpose(xFi_Range)
End If

If Fi_Range.Rows.Count = 1 And Fi_Range.Columns.Count = 1 Then
 ReDim F(1 To 1) As Variant
 F(1) = Fi_Range.Cells(1, 1).Value
Else
 F = Application.Transpose(Fi_Range)
End If
'--- equal size of input vectors
nxF = UBound(xF)
nF = UBound(F)

If nxF <> nF Then
  Querkraft = "Error: Different range sizes"
  Exit Function
 Else
 nloads = nxF
End If
'aling orientation
If xFi_Range.Columns.Count > xFi_Range.Rows.Count Then
  xF = Application.Transpose(Application.Transpose(xFi_Range))
End If
If Fi_Range.Columns.Count > Fi_Range.Rows.Count Then
  F = Application.Transpose(Application.Transpose(Fi_Range))
End If
''--- Inputdaten provisorisch bereinigt -------------------
'----------------------------------------------------------
'
'Aus ME und Breite den Bettungsmodul ks berechnen (Mit Boussinesq+Annahmen)
ks = MeModul / (fShape(b, xEnd) * b)
'(Option: aus Cc, Breite und Mächtigkeit der Schicht ks berechnen (Boussinesq+Layers))
'--------------------------------
'Fallunterscheidung für Verfahren
'Das Spannungstrapezverfahren liefert keine Setzung nur die Biegelinie des Fundaments.
'Deshalb wird in allen Fällen das BMV verwendet.
L = (4 * EI / (ks * b)) ^ 0.25
If xEnd < 2 * L Then
    'Option Biegelinie des Balkens unter Trapezbelastung ausgeben
    Querkraft = BMV_Q(xPosition, xEnd, xF, F, nloads, L)
Else
    Querkraft = BMV_Q(xPosition, xEnd, xF, F, nloads, L)
    
End If
End Function 'Biegelinie()

Private Function STV_M(xPosition, xEnd, xF, F, nloads) 'Spannungstrapezverfahren
'Das Spannungstapezverfahren setzt eine trapezförmige Sohlpressung an,
'die im Gleichgewicht mit den angreifenden Lasten ist. Am einfachsten
'berechnet sich die Baugrundreaktion aus der Resultierenden aller Lasten
'und der Exzentrizität. Das Verfahren vernachlässigt die Verträglichkeits-
'bedingung von Baugrundverformung und Fundamentverformung.
'Gleichgewicht ist aber eingehalten und “unrealistische” Spannungsspitzen
'am Fundamentrand kommen nicht vor. Die Lösung ist relativ nah an den
'komplexeren Lösungen für eher starre Fundamente.

'---Resultierende und Exzentrizität:
Dim R As Double: R = 0 'Reaction force
Dim xR As Double: xR = 0 'Position of reaction force
Dim e, q0, qx, qend As Double

For I = 1 To nloads 'Load
    If F(I) = 0 Then
     xR = xR
    Else
     xR = (xR * R + (xF(I) * F(I))) / (R + F(I))
    End If
    R = R + F(I)
Next I
e = xR - xEnd / 2
' ---Sohlpressung = F(R, e)
q0 = R / (xEnd) * (1 - 6 * e / xEnd)  'pressure at coordinate x=0
qend = R / (xEnd) * (1 + 6 * e / xEnd) 'pressure at coordinate x=xEnd
'---(Später: Option Mit Fuge, d.h. Keine Zugspannung)-------------------------------Option Fuge
'- Superposition
'Moment infolge Sohldruck
qx = q0 + (qend - q0) * xPosition / xEnd
M = -xPosition * xPosition / 2 * (q0 + (qx - q0) / 3)
'Moment infolge Einzellasten
For j = 1 To nloads 'Load
    If xPosition > xF(j) Then
        M = M + F(j) * (xPosition - xF(j))
    End If
Next
STV_M = M
End Function 'STV_M ausgabe in kNm
Private Function BMV_M(xPosition, xEnd, xF, F, nloads, Lelast)
'Berechnung des Moments nach dem Bettungsmodulverfahren.
'je nachdem ob die Last nahe am Rand ist müssen die Randbedingungen berücksichtigt werden.
'zudem muss für fast symmetrische Laststellung wegen Stellenauslöschung eine Näherungslösung verwendet werden
'- Superposition
'Moment infolge Einzellasten (inklusive Reaktion pro Einzellast)
Dim M As Double
M = 0
For j = 1 To nloads 'Load
    Debug.Print "Load"; j
     dist1 = 0 + 2 * Lelast
     dist2 = xEnd - 2 * Lelast
     If xEnd < 4.8 * Lelast * 2 Then
        condition_finite = True
     Else
        condition_finite = False
     End If
     If xF(j) = 0 Or xF(j) = xEnd Then
        condition_edge = True
     Else
        condition_edge = False
     End If

     If condition_finite = False Then

        If condition_edge = True Then
            If xF(j) = 0 Then
                Debug.Print "semi inifinite beam at edge x = 0"
                State = SetCoefficientsInfiniteBeamEdge(xEnd - xF(j), xEnd, Lelast, A1, A2, A3, A4, B1, B2, B3, B4)
                a = -(xPosition - xF(j)) / Lelast
                Debug.Print a
            Else
                Debug.Print "semi inifinite beam at edge x= xEnd"
                State = SetCoefficientsInfiniteBeamEdge(xF(j), xEnd, Lelast, A1, A2, A3, A4, B1, B2, B3, B4)
                a = ((xPosition - xF(j)) / Lelast)
            End If
        Else
            If xF(j) < dist1 Then
                Debug.Print "semi_infinite beam one border x=0"
                State = SetCoefficientsSemiFiniteBeam(xEnd - xF(j), xEnd, Lelast, A1, A2, A3, A4, B1, B2, B3, B4)
                a = -(xPosition - xF(j)) / Lelast
            Else
              If xF(j) > dist2 Then
                Debug.Print "semi_infinite beam one border x=xEnd"
                State = SetCoefficientsSemiFiniteBeam(xF(j), xEnd, Lelast, A1, A2, A3, A4, B1, B2, B3, B4)
                a = (xPosition - xF(j)) / Lelast
              Else
                Debug.Print "infinite beam"
                State = SetCoefficientsInfiniteBeam(xF(j), xEnd, Lelast, A1, A2, A3, A4, B1, B2, B3, B4)
                a = (xPosition - xF(j)) / Lelast
              End If
            End If
        End If
     Else 'finite beam
        If condition_edge = True Then
            If xF(j) = 0 Then
                Debug.Print "finite beam at edge x=0"
                State = SetCoefficientsFiniteBeamEdge(xF(j), xEnd, Lelast, A1, A2, A3, A4, B1, B2, B3, B4)
                a = (xPosition - xF(j)) / Lelast
            Else
                Debug.Print "finite beam at edge x=xEnd"
                State = SetCoefficientsFiniteBeamEdge(xEnd - xF(j), xEnd, Lelast, A1, A2, A3, A4, B1, B2, B3, B4)
                a = -(xPosition - xF(j)) / Lelast
            End If
        Else
                Debug.Print "finite beam"
                State = SetCoefficientsFiniteBeam(xF(j), xEnd, Lelast, A1, A2, A3, A4, B1, B2, B3, B4)
                a = (xPosition - xF(j)) / Lelast
        End If
     End If

    If xPosition < 0 Or xPosition > xEnd Then
        'dy = 0
        dM = 0
        'dQ = 0
    Else
     If a <= 0 Then
        'dY = F(j) / (EI) * (B1 * Exp(a) * Cos(a) + B2 * Exp(a) * Sin(a) + B3 * Exp(-a) * Cos(a) + B4 * Exp(-a) * Sin(a))
        dM = -2 * F(j) / Lelast ^ 2 * (B2 * Exp(a) * Cos(a) - B1 * Exp(a) * Sin(a) - B4 * Exp(-a) * Cos(a) + B3 * Exp(-a) * Sin(a))
        'dQ = -2 * F(j) / Lelast ^ 3 * ((B2 - B1) * Exp(a) * Cos(a) - (B1 + B2) * Exp(a) * Sin(a) + (B3 + B4) * Exp(-a) * Cos(a) + (B4 - B3) * Exp(-a) * Sin(a))
     Else
        'dY = F(j) / (EI) * (A1 * Exp(a) * Cos(a) + A2 * Exp(a) * Sin(a) + A3 * Exp(-a) * Cos(a) + A4 * Exp(-a) * Sin(a))
        dM = -2 * F(j) / Lelast ^ 2 * (A2 * Exp(a) * Cos(a) - A1 * Exp(a) * Sin(a) - A4 * Exp(-a) * Cos(a) + A3 * Exp(-a) * Sin(a))
        'dQ = -2 * F(j) / Lelast ^ 3 * ((A2 - A1) * Exp(a) * Cos(a) - (A1 + A2) * Exp(a) * Sin(a) + (A3 + A4) * Exp(-a) * Cos(a) + (A4 - A3) * Exp(-a) * Sin(a))
     End If
    End If
     M = M + dM
Next j
BMV_M = M
End Function 'BMV_M Ausgabe in kNm
Private Function BMV_Y(xPosition, xEnd, xF, F, nloads, Lelast, EI)
'Berechnung die Verschiebung nach dem Bettungsmodulverfahren.
'je nachdem ob die Last nahe am Rand ist müssen die Randbedingungen berücksichtigt werden.
'zudem muss für fast symmetrische Laststellung wegen Stellenauslöschung eine Näherungslösung verwendet werden
'für die Sohlpressung kann gemäss der definitiondes BMV die Verschiebung mit dem Bettungsmodul multipliziert werden
'- Superposition
'Biegelinie infolge Einzellasten

Dim Y As Double
Y = 0
For j = 1 To nloads 'Load
    Debug.Print "Load"; j
     dist1 = 0 + 2 * Lelast
     dist2 = xEnd - 2 * Lelast
     If xEnd < 4.8 * Lelast * 2 Then
        condition_finite = True
     Else
        condition_finite = False
     End If
     If xF(j) = 0 Or xF(j) = xEnd Then
        condition_edge = True
     Else
        condition_edge = False
     End If

     If condition_finite = False Then

        If condition_edge = True Then
            If xF(j) = 0 Then
                Debug.Print "semi inifinite beam at edge x = 0"
                State = SetCoefficientsInfiniteBeamEdge(xF(j), xEnd, Lelast, A1, A2, A3, A4, B1, B2, B3, B4)
                a = (xPosition - xF(j)) / Lelast
                Debug.Print
            Else
                Debug.Print "semi inifinite beam at edge x= xEnd"
                State = SetCoefficientsInfiniteBeamEdge(xF(j), xEnd, Lelast, A1, A2, A3, A4, B1, B2, B3, B4)
                a = (xPosition - xF(j)) / Lelast
            End If
        Else
            If xF(j) < dist1 Then
                Debug.Print "semi_infinite beam one border x=0"
                State = SetCoefficientsSemiFiniteBeam(xEnd - xF(j), xEnd, Lelast, A1, A2, A3, A4, B1, B2, B3, B4)
                a = -(xPosition - xF(j)) / Lelast
            Else
              If xF(j) > dist2 Then
                Debug.Print "semi_infinite beam one border x=xEnd"
                State = SetCoefficientsSemiFiniteBeam(xF(j), xEnd, Lelast, A1, A2, A3, A4, B1, B2, B3, B4)
                a = (xPosition - xF(j)) / Lelast
              Else
                Debug.Print "infinite beam"
                State = SetCoefficientsInfiniteBeam(xF(j), xEnd, Lelast, A1, A2, A3, A4, B1, B2, B3, B4)
                a = (xPosition - xF(j)) / Lelast
              End If
            End If
        End If
     Else 'finite beam
        If condition_edge = True Then
            If xF(j) = 0 Then
                Debug.Print "finite beam at edge x=0"
                State = SetCoefficientsFiniteBeamEdge(xF(j), xEnd, Lelast, A1, A2, A3, A4, B1, B2, B3, B4)
                a = (xPosition - xF(j)) / Lelast
            Else
                Debug.Print "finite beam at edge x=xEnd"
                State = SetCoefficientsFiniteBeamEdge(xEnd - xF(j), xEnd, Lelast, A1, A2, A3, A4, B1, B2, B3, B4)
                a = -(xPosition - xF(j)) / Lelast
            End If
        Else
                Debug.Print "finite beam"
                State = SetCoefficientsFiniteBeam(xF(j), xEnd, Lelast, A1, A2, A3, A4, B1, B2, B3, B4)
                a = (xPosition - xF(j)) / Lelast
        End If
     End If

    If xPosition < 0 Or xPosition > xEnd Then
        dY = 0
        'dM = 0
        'dQ = 0
    Else
     If a <= 0 Then
        dY = F(j) / (EI) * (B1 * Exp(a) * Cos(a) + B2 * Exp(a) * Sin(a) + B3 * Exp(-a) * Cos(a) + B4 * Exp(-a) * Sin(a))
        'dM = -2 * F(j) / Lelast ^ 2 * (B2 * Exp(a) * Cos(a) - B1 * Exp(a) * Sin(a) - B4 * Exp(-a) * Cos(a) + B3 * Exp(-a) * Sin(a))
        'dQ = -2 * F(j) / Lelast ^ 3 * ((B2 - B1) * Exp(a) * Cos(a) - (B1 + B2) * Exp(a) * Sin(a) + (B3 + B4) * Exp(-a) * Cos(a) + (B4 - B3) * Exp(-a) * Sin(a))
     Else
        dY = F(j) / (EI) * (A1 * Exp(a) * Cos(a) + A2 * Exp(a) * Sin(a) + A3 * Exp(-a) * Cos(a) + A4 * Exp(-a) * Sin(a))
        'dM = -2 * F(j) / Lelast ^ 2 * (A2 * Exp(a) * Cos(a) - A1 * Exp(a) * Sin(a) - A4 * Exp(-a) * Cos(a) + A3 * Exp(-a) * Sin(a))
        'dQ = -2 * F(j) / Lelast ^ 3 * ((A2 - A1) * Exp(a) * Cos(a) - (A1 + A2) * Exp(a) * Sin(a) + (A3 + A4) * Exp(-a) * Cos(a) + (A4 - A3) * Exp(-a) * Sin(a))
     End If
    End If
     Y = Y + dY
Next j
BMV_Y = Y

End Function 'BMV_Y Ausgabe in m
Private Function BMV_Q(xPosition, xEnd, xF, F, nloads, Lelast)
'Berechnung des Querkraft nach dem Bettungsmodulverfahren.
'je nachdem ob die Last nahe am Rand ist müssen die Randbedingungen berücksichtigt werden.
'zudem muss für fast symmetrische Laststellung wegen Stellenauslöschung eine Näherungslösung verwendet werden
'- Superposition
'Querkraft infolge Einzellasten (inklusive Reaktion pro Einzellast)
Dim Q As Double
Q = 0
For j = 1 To nloads 'Load
    Debug.Print "Load"; j
     dist1 = 0 + 2 * Lelast
     dist2 = xEnd - 2 * Lelast
     If xEnd < 4.8 * Lelast * 2 Then
        condition_finite = True
     Else
        condition_finite = False
     End If
     If xF(j) = 0 Or xF(j) = xEnd Then
        condition_edge = True
     Else
        condition_edge = False
     End If

     If condition_finite = False Then

        If condition_edge = True Then
            If xF(j) = 0 Then
                Debug.Print "semi inifinite beam at edge x = 0"
                State = SetCoefficientsInfiniteBeamEdge(xF(j), xEnd, Lelast, A1, A2, A3, A4, B1, B2, B3, B4)
                a = (xPosition - xF(j)) / Lelast
                Debug.Print
            Else
                Debug.Print "semi inifinite beam at edge x= xEnd"
                State = SetCoefficientsInfiniteBeamEdge(xF(j), xEnd, Lelast, A1, A2, A3, A4, B1, B2, B3, B4)
                'B1 = -B1
                a = (xPosition - xF(j)) / Lelast
            End If
        Else
            If xF(j) < dist1 Then
                Debug.Print "semi_infinite beam one border x=0"
                State = SetCoefficientsSemiFiniteBeam(xEnd - xF(j), xEnd, Lelast, A1, A2, A3, A4, B1, B2, B3, B4)
                A1 = -A1
                A2 = -A2
                A3 = -A3
                A4 = -A4
                B1 = -B1
                B2 = -B2
                B3 = -B3
                B4 = -B4
                a = -(xPosition - xF(j)) / Lelast
            Else
              If xF(j) > dist2 Then
                Debug.Print "semi_infinite beam one border x=xEnd"
                State = SetCoefficientsSemiFiniteBeam(xF(j), xEnd, Lelast, A1, A2, A3, A4, B1, B2, B3, B4)
                a = (xPosition - xF(j)) / Lelast
              Else
                Debug.Print "infinite beam"
                State = SetCoefficientsInfiniteBeam(xF(j), xEnd, Lelast, A1, A2, A3, A4, B1, B2, B3, B4)
                a = (xPosition - xF(j)) / Lelast
              End If
            End If
        End If
     Else 'finite beam
        If condition_edge = True Then
            If xF(j) = 0 Then
                Debug.Print "finite beam at edge x=0"
                State = SetCoefficientsFiniteBeamEdge(xF(j), xEnd, Lelast, A1, A2, A3, A4, B1, B2, B3, B4)
                a = (xPosition - xF(j)) / Lelast
            Else
                Debug.Print "finite beam at edge x=xEnd"
                State = SetCoefficientsFiniteBeamEdge(xEnd - xF(j), xEnd, Lelast, A1, A2, A3, A4, B1, B2, B3, B4)
                A1 = -A1
                A2 = -A2
                A3 = -A3
                A4 = -A4
                B1 = -B1
                B2 = -B2
                B3 = -B3
                B4 = -B4
                a = -(xPosition - xF(j)) / Lelast
            End If
        Else
                Debug.Print "finite beam"
                State = SetCoefficientsFiniteBeam(xF(j), xEnd, Lelast, A1, A2, A3, A4, B1, B2, B3, B4)
                a = (xPosition - xF(j)) / Lelast
        End If
     End If

    If xPosition < 0 Or xPosition > xEnd Then
        'y = 0
        'dM = 0
        dQ = 0
    Else
     If a <= 0 Then
        'dY = F(j) / (EI) * (B1 * Exp(a) * Cos(a) + B2 * Exp(a) * Sin(a) + B3 * Exp(-a) * Cos(a) + B4 * Exp(-a) * Sin(a))
        'dM = -2 * F(j) / Lelast ^ 2 * (B2 * Exp(a) * Cos(a) - B1 * Exp(a) * Sin(a) - B4 * Exp(-a) * Cos(a) + B3 * Exp(-a) * Sin(a))
        dQ = -2 * F(j) / Lelast ^ 3 * ((B2 - B1) * Exp(a) * Cos(a) - (B1 + B2) * Exp(a) * Sin(a) + (B3 + B4) * Exp(-a) * Cos(a) + (B4 - B3) * Exp(-a) * Sin(a))
     Else
        'dY = F(j) / (EI) * (A1 * Exp(a) * Cos(a) + A2 * Exp(a) * Sin(a) + A3 * Exp(-a) * Cos(a) + A4 * Exp(-a) * Sin(a))
        'dM = -2 * F(j) / Lelast ^ 2 * (A2 * Exp(a) * Cos(a) - A1 * Exp(a) * Sin(a) - A4 * Exp(-a) * Cos(a) + A3 * Exp(-a) * Sin(a))
        dQ = -2 * F(j) / Lelast ^ 3 * ((A2 - A1) * Exp(a) * Cos(a) - (A1 + A2) * Exp(a) * Sin(a) + (A3 + A4) * Exp(-a) * Cos(a) + (A4 - A3) * Exp(-a) * Sin(a))
     End If
    End If
     Q = Q + dQ
Next j
BMV_Q = Q
     
End Function
' Helperfunctions
'-------------------------------------------------------------------------------------------------------
Private Function fShape(ByVal width, ByVal length)
'[Bild 11.10 aus LHAP]
fShape = 1.8 'provisorisch
End Function
Private Function SetCoefficientsFiniteBeam(ByVal ForcePosition, ByVal PositionEnd, ByVal Lelast, _
ByRef A1, ByRef A2, ByRef A3, ByRef A4, ByRef B1, ByRef B2, ByRef B3, ByRef B4) As Boolean
    aL = (PositionEnd - ForcePosition) / Lelast
    bL = (0 - ForcePosition) / Lelast
    'Differential equation
    ' y''''+4*y/L^4=0
    ' general solution
    ' ya=A1*exp(a)*cos(a)+A2*exp(a)*sin(a)+A3*exp(-a)*cos(-a)+A4*exp(a)*sin(a) for a>=0
    ' yb=B1*exp(a)*cos(a)+B2*exp(a)*sin(a)+B3*exp(-a)*cos(-a)+B4*exp(a)*sin(a) for a<=0
    'boundary conditions
    'see documentation
    Q = -1 * Lelast ^ 3 / (2 * 4)
    'Coefficients
    Ca_1 = Exp(aL) * Sin(aL)
    Ca_2 = Exp(aL) * Cos(aL)
    Ca_3 = Exp(-aL) * Sin(aL)
    Ca_4 = Exp(-aL) * Cos(aL)
    Cb_1 = Exp(bL) * Sin(bL)
    Cb_2 = Exp(bL) * Cos(bL)
    Cb_3 = Exp(-bL) * Sin(bL)
    Cb_4 = Exp(-bL) * Cos(bL)
    D3 = (2 * Cb_1 * Cb_3 - Cb_1 * Cb_4 + Cb_2 * Cb_3) / (Cb_1 * Cb_1 + Cb_2 * Cb_2)
    D4 = (-Cb_1 * Cb_3 - 2 * Cb_1 * Cb_4 - Cb_2 * Cb_4) / (Cb_1 * Cb_1 + Cb_2 * Cb_2)
    E3 = (Ca_3 - Ca_1 * Cb_3 / Cb_1 - D3 * (Ca_2 - Ca_1 * Cb_2 / Cb_1)) * Cb_1
    E4 = (-Ca_4 + Ca_1 * Cb_4 / Cb_1 - D4 * (Ca_2 - Ca_1 * Cb_2 / Cb_1)) * Cb_1
    E5 = (Ca_1 + Ca_2 + Ca_3 - Ca_4) * Cb_1
    F3 = (-2 * Ca_3 + Ca_4 - Ca_2 * Cb_3 / Cb_1 - D3 * (-Ca_1 - Ca_2 * Cb_2 / Cb_1)) * Cb_1
    F4 = (Ca_3 + 2 * Ca_4 + Ca_2 * Cb_4 / Cb_1 - D4 * (-Ca_1 - Ca_2 * Cb_2 / Cb_1)) * Cb_1
    'Solution
    B4 = -Q * ((-Ca_1 + Ca_2 - Ca_3 + 3 * Ca_4) * Cb_1 - E5 / E3 * F3) / (F4 - F3 * E4 / E3) 'eq5.9
    B3 = -Q * E5 / E3 - E4 * B4 / E3 'eq4.8
    B2 = 0 - (D3 * B3 + D4 * B4)                'eq7.3
    B1 = 0 - (-Cb_2 / Cb_1 * B2 - Cb_3 / Cb_1 * B3 + Cb_4 / Cb_1 * B4) 'eq6.1
    A4 = Q - (-B4)                              'eq8.3
    A3 = 0 - (-A4 - B3 + B4)                    'eq3
    A2 = 0 - (-2 * A3 + A4 - B2 + 2 * B3 - B4)  'eq2
    A1 = 0 - (1 * A3 - B1 - B3)                 'eq1
    SetCoefficientsFiniteBeam = True
End Function

Private Function SetCoefficientsFiniteBeamEdge(ByVal ForcePosition, ByVal PositionEnd, ByVal Lelast, _
ByRef A1, ByRef A2, ByRef A3, ByRef A4, ByRef B1, ByRef B2, ByRef B3, ByRef B4) As Boolean
'Force at x=0
    aL = (PositionEnd - ForcePosition) / Lelast '->positive
    bL = (0 - ForcePosition) / Lelast           '=0
    'Differential equation
    ' y''''+4*y/L^4=0
    ' general solution
    ' ya=A1*exp(a)*cos(a)+A2*exp(a)*sin(a)+A3*exp(-a)*cos(-a)+A4*exp(a)*sin(a) for a>=0
    ' yb=B1*exp(a)*cos(a)+B2*exp(a)*sin(a)+B3*exp(-a)*cos(-a)+B4*exp(a)*sin(a) for a<=0
    'boundary conditions
    'see documentation
    Q = -1 * Lelast ^ 3 / (2#)
    'Coefficients
    Ca_1 = Exp(aL) * Sin(aL)
    Ca_2 = Exp(aL) * Cos(aL)
    Ca_3 = Exp(-aL) * Sin(aL)
    Ca_4 = Exp(-aL) * Cos(aL)
    'Cb_1 = Exp(bL) * Sin(bL)
    'Cb_2 = Exp(bL) * Cos(bL)
    'Cb_3 = Exp(-bL) * Sin(bL)
    'Cb_4 = Exp(-bL) * Cos(bL)
    KD_3 = -Ca_1 + Ca_3
    KD_4 = -2 * Ca_1 + Ca_2 - Ca_4
    KE_3 = Ca_4 - Ca_2 - 2 * Ca_3
    KE_4 = -Ca_1 - 2 * Ca_2 + Ca_3 + 2 * Ca_4
    B4 = 0
    B3 = 0
    B2 = 0
    B1 = 0
    A4 = -Q * (Ca_2 - Ca_1 * KE_3 / KD_3) / (KE_4 - KD_4 * KE_3 / KD_3)
    A3 = -Q * Ca_1 / KD_3 - A4 * KD_4 / KD_3
    A2 = A4
    A1 = -Q + A3 + 2 * A4
    SetCoefficientsFiniteBeamEdge = True
End Function

Private Function SetCoefficientsInfiniteBeam(ByVal ForcePosition, ByVal PositionEnd, ByVal Lelast, _
ByRef A1, ByRef A2, ByRef A3, ByRef A4, ByRef B1, ByRef B2, ByRef B3, ByRef B4) As Boolean
    'Differential equation
    ' y''''+4*y/L^4=0
    ' general solution
    ' ya=A1*exp(a)*cos(a)+A2*exp(a)*sin(a)+A3*exp(-a)*cos(-a)+A4*exp(a)*sin(a) for a>=0
    ' yb=B1*exp(a)*cos(a)+B2*exp(a)*sin(a)+B3*exp(-a)*cos(-a)+B4*exp(a)*sin(a) for a<=0
    'boundary conditions
    'see documentation
    B4 = 0
    B3 = 0
    B2 = 1 * Lelast ^ 3 / 8
    B1 = -1 * Lelast ^ 3 / 8
    A4 = -1 * Lelast ^ 3 / 8
    A3 = -1 * Lelast ^ 3 / 8
    A2 = 0
    A1 = 0
    SetCoefficientsInfiniteBeam = True
End Function
Private Function SetCoefficientsInfiniteBeamEdge(ByVal ForcePosition, ByVal PositionEnd, ByVal Lelast, _
ByRef A1, ByRef A2, ByRef A3, ByRef A4, ByRef B1, ByRef B2, ByRef B3, ByRef B4) As Boolean
'Force at x=0 or x=xEnd
    'Differential equation
    ' y''''+4*y/L^4=0
    ' general solution
    ' ya=A1*exp(a)*cos(a)+A2*exp(a)*sin(a)+A3*exp(-a)*cos(-a)+A4*exp(a)*sin(a) for a>=0  for force at x=x0
    ' yb=B1*exp(a)*cos(a)+B2*exp(a)*sin(a)+B3*exp(-a)*cos(-a)+B4*exp(a)*sin(a) for a<=0  for force at x=xend
    'boundary conditions
    'see documentation
    B4 = 0
    B3 = 0
    B2 = 0
    B1 = -1 * Lelast ^ 3 / 2
    A4 = 0
    A3 = -1 * Lelast ^ 3 / 2
    A2 = 0
    A1 = 0
    SetCoefficientsInfiniteBeamEdge = True
End Function
Private Function SetCoefficientsSemiFiniteBeam(ByVal ForcePosition, ByVal PositionEnd, ByVal Lelast, _
ByRef A1, ByRef A2, ByRef A3, ByRef A4, ByRef B1, ByRef B2, ByRef B3, ByRef B4) As Boolean
'force close to PositionEnd
    aL = (PositionEnd - ForcePosition) / Lelast
    bL = (0 - ForcePosition) / Lelast
    'Differential equation
    ' y''''+4*y/L^4=0
    ' general solution
    ' ya=A1*exp(a)*cos(a)+A2*exp(a)*sin(a)+A3*exp(-a)*cos(-a)+A4*exp(a)*sin(a) for a>=0
    ' yb=B1*exp(a)*cos(a)+B2*exp(a)*sin(a)+B3*exp(-a)*cos(-a)+B4*exp(a)*sin(a) for a<=0
    'boundary conditions
    'see documentation
    Q = 1 * Lelast ^ 3 / (2 * 4)
    'Coefficients
    Ca_1 = Exp(aL) * Sin(aL)
    Ca_2 = Exp(aL) * Cos(aL)
    Ca_3 = Exp(-aL) * Sin(aL)
    Ca_4 = Exp(-aL) * Cos(aL)
'    Cb_1 = Exp(bL) * Sin(bL)
'    Cb_2 = Exp(bL) * Cos(bL)
'    Cb_3 = Exp(-bL) * Sin(bL)
'    Cb_4 = Exp(-bL) * Cos(bL)
    Konst_4_R = (Ca_1 + Ca_2 + Ca_3 - Ca_4) / Ca_1
    Konst_5_2 = Ca_1 + Ca_2 * Ca_2 / Ca_1
    Konst_5_R = (-Ca_1 + Ca_2 - Ca_3 + 3 * Ca_4 - Konst_4_R * Ca_2)
    
    B4 = 0
    B3 = 0
    B2 = -Q * Konst_5_R / Konst_5_2
    B1 = -Q * Konst_4_R + Ca_2 / Ca_1 * B2
    A4 = -Q
    A3 = A4
    A2 = 0 - (-2 * A3 + 1 * A4 - 1 * B2)
    A1 = 0 - (A3 - B1)
    SetCoefficientsSemiFiniteBeam = True
End Function

Sub DescribeFunction1()
   Dim FuncName As String
   Dim FuncDesc As String
   Dim Category As String
   Dim ArgDesc(1 To 8) As String

   FuncName = "Moment"
   FuncDesc = "Berechnet das Moment im Fundament an der angegebenen Stelle in kNm/m'"
   ArgDesc(1) = "Koordinate, wo das Moment bestimmt werden soll in m, x=0 am Fundamentrand"
   ArgDesc(2) = "Länge des Fundaments in m"
   ArgDesc(3) = "Breite des Fundament in m"
   ArgDesc(4) = "Steifigkeit des Fundaments EI in kNm2"
   ArgDesc(5) = "ME-Modul des Baugrunds in kN/m2"
   ArgDesc(6) = "Bereich mit den Ortskoordinaten der einwirkenden Kräfte in m, x=0 am Fundamentrand"
   ArgDesc(7) = "Bereich mit den Beträgen der einwirkenden Kräfte in kN"
   ArgDesc(8) = "Optional, Modus WAHR/true->Spannungstrapezverfahren, sonst BMV [Default=FALSCH/false]"

 
    Category = 14 '14=user defined
    
   Application.MacroOptions _
      Macro:=FuncName, _
      Description:=FuncDesc, _
      Category:=Category, _
      ArgumentDescriptions:=ArgDesc
      
End Sub
Sub DescribeFunction2()
   Dim FuncName As String
   Dim FuncDesc As String
   Dim Category As String
   Dim ArgDesc(1 To 7) As String

   FuncName = "Biegelinie"
   FuncDesc = "Berechnet die Verschiebung im Fundament resp die Setzung des Baugrunds an der angegebenen Stelle in m"
   ArgDesc(1) = "x-Koordinate, wo die Verschiebung bestimmt werden soll in m, x=0 am Fundamentrand"
   ArgDesc(2) = "Länge des Fundaments in m"
   ArgDesc(3) = "Breite des Fundament in m"
   ArgDesc(4) = "Steifigkeit des Fundaments EI in kNm2"
   ArgDesc(5) = "ME-Modul des Baugrunds in kN/m2"
   ArgDesc(6) = "Bereich mit den Ortskoordinaten der einwirkenden Kräfte in m, x=0 am Fundamentrand"
   ArgDesc(7) = "Bereich mit den Beträgen der einwirkenden Kräfte in kN"
   
 
    Category = 14 '14=user defined
    
   Application.MacroOptions _
      Macro:=FuncName, _
      Description:=FuncDesc, _
      Category:=Category, _
      ArgumentDescriptions:=ArgDesc
      
End Sub
Sub DescribeFunction3()
   Dim FuncName As String
   Dim FuncDesc As String
   Dim Category As String
   Dim ArgDesc(1 To 8) As String

   FuncName = "Querkraft"
   FuncDesc = "Berechnet die Querkraft im Fundament an der angegebenen Stelle in kN/m"
   ArgDesc(1) = "Koordinate, wo die Querkraft bestimmt werden soll in. x=0 am Fundamentrand"
   ArgDesc(2) = "Länge des Fundaments in m"
   ArgDesc(3) = "Breite des Fundament in m"
   ArgDesc(4) = "Steifigkeit des Fundaments EI in kNm2"
   ArgDesc(5) = "ME-Modul des Baugrunds in kN/m2"
   ArgDesc(6) = "Bereich mit den Ortskoordinaten der einwirkenden Kräfte in m, x=0 am Fundamentrand"
   ArgDesc(7) = "Bereich mit den Beträgen der einwirkenden Kräfte in kN"
   ArgDesc(8) = "Optional, Modus WAHR/true->Spannungstrapezverfahren, sonst BMV [Default=FALSCH/false]"

 
    Category = 14 '14=user defined
    
   Application.MacroOptions _
      Macro:=FuncName, _
      Description:=FuncDesc, _
      Category:=Category, _
      ArgumentDescriptions:=ArgDesc
      
End Sub

Function Initialize()

If initializedState <> "Beschreibung hinzugefügt" Then
    DescribeFunction1
    DescribeFunction2
    DescribeFunction3
    'DescribeFunction4
    MsgBox ("Makros initialisiert")
    initializedState = "Beschreibung hinzugefügt"
End If
Initialize = VersionG
End Function




