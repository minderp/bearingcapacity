Attribute VB_Name = "BearingCapacityModul"
Dim VersionStringG As String
Dim initializedState As String
Public Function VersionG()
VersionStringG = "Version 0.1.1, 2018-02-01"
VersionG = VersionStringG
End Function

Function Grundbruch_Rechteck(c, phi, gamma, q_soil, t_soil, b, a As Variant, _
Optional omega = 0, Optional eB = 0, Optional eA = 0, _
Optional beta = 0, Optional alpha = 0, Optional Fresb = 0)
Attribute Grundbruch_Rechteck.VB_Description = "Berechnet die zul�ssige Bodenpressung Rd,N in [kN] aufgrund der angegebenen Bodenparametern, Fundamentgeometrie und Belastungsrichtung"
Attribute Grundbruch_Rechteck.VB_ProcData.VB_Invoke_Func = " \n14"
'   FuncName = "Grundbruch_Rechteck"
'   FuncDesc = "Berechnet die zul�ssige Bodenpressung Rd,N in [kN] aufgrund der angegebenen "& _
'              "Bodenparametern, Fundamentgeometrie und Belastungsrichtung"
'   ArgDesc(1) = "Koh�sion in kPa"
'   ArgDesc(2) = "Reibungswinkel in Grad "
'   ArgDesc(3) = "Effektives Bodengewicht unter Fundament kN/m3 (\gamma f�r trockener Boden, \gamma' bei Grundwasser bis Sohle)�
'   ArgDesc(4) = "Auflast neben Fundament auf Niveau Sohle (inkl. Bodengewicht) in kPa, (\gamma t +q)"
'   ArgDesc(5) = "Einbindetiefe (Abstand OKT zur Sohle) in m"
'   ArgDesc(6) = "Breite des Fundaments in Versagensrichtung in m"
'   ArgDesc(7) = "L�nge des Fundaments quer zur Versagensrichtung in m"
'   ArgDesc(8) = "Abweichung der Kraftrichtung zur Vertikalen in Grad [Default=0]"
'   ArgDesc(9) = "Exzentrizit�t der Resultierenden in Versagensrichtung in m [Default=0]"
'   ArgDesc(10) = "Exzentrizit�t der Resultierenden quer zur Versagensrichtung in m [Default=0]"
'   ArgDesc(11) = "Gel�ndeneigung in Grad [Default=0]"
'   ArgDesc(12) = "Sohlneigung in Grad [Default=0]"
'   ArgDesc(13) = "Betrag der resultierenden Einwirkung in Versagensebene, n�tig falls c>0"
Const PI As Double = 3.14159265358979
'
'Normparameter
gamma_phi = 1.2
gamma_c = 1.5
gamma_g = 1#
'Designparameter
cd = c / gamma_c 'Cohesion design [kPa]
phid = Atn(Tan(phi / 180 * PI) / gamma_phi) 'Friction angle design [rad]
gammad = gamma / gamma_g 'density design [kN/m3]
'--------------------------------------------------------------------------------
'Inputparameter f�r die Tragf�higkeitsformel auf die Konventionen vorbereiten:
'------ Fundamentneigung, B�schungsneigung und Lastneigung k�nnen nur in "b"-Richtung ber�cksichtigt werden.
'------ alpha, beta und omega sind nur eingeschr�nkt g�ltig. F�r werte ausserhalb des
'------ G�ltigkeitsbereichs werden folgende konservativen Annahmen getroffen
'------ Auch die Exzentrizit�t infolge Moment oder Verschobener Resultierenden wird nur f�r ung�nstige Kombinationen ber�cksichtigt.
alpha = Application.Max(0, alpha) 'nicht definiert f�r negative Werte, Konservativer cutoff bei 0
beta = Application.Max(0, beta)   'nicht definiert f�r negative Werte, Konservativer cutoff bei 0
'Abfangen von negativen oder entgegensetzten Effekten (Lastneigung, Exzentrizit�t)
If omega <= 0 And eB <= 0 Then 'beide negativ oder null->Versagen in andere Richtung
 omega = Abs(omega)
 eB = Abs(eB)
Else
 If (omega < 0 And eB > 0) Or (omega > 0 And eB < 0) Then 'Lastneigung ist Exzentrizit�t entgegengesetzt
 'im Fall von g�nstiger Lastneigung wird sie vernachl�ssigt (omega=0). Das Vorzeichen der Exzentrizit�t wird an die Konvention angepasst.
  X1 = Grundbruch_Rechteck(c, phi, gamma, q_soil, t_soil, b, a, 0, Abs(eB), eA, beta, alpha, Fresb) 'Achtung Rekursion
  'im Fall von g�nstiger Exzentrizit�t wird sie vernachl�ssigt (eB=0). Das Vorzeichen der Lastneigung wird an die Konvention angepasst.
  X2 = Grundbruch_Rechteck(c, phi, gamma, q_soil, t_soil, b, a, Abs(omega), 0, eA, beta, alpha, Fresb) 'Achtung Rekursion
  Grundbruch_Rechteck = Application.Min(X1, X2)
  Exit Function
 End If
 'der verbleibende Fall, wo beide Werte positiv sind, erf�llt die Konvention ohne Anpassung.
End If
'Reduzierte Fundamentfl�che f�r exzentrisch belastete Sohle. (Ausschliessen von negativer Fl�che)
beff = Application.Max(0, b - 2 * Abs(eB))
'Abfangen Streifenfundament
If a = "Streifen" Then
    aeff = 1
    Else
    aeff = Application.Max(0, a - 2 * Abs(eA)) 'in der Richtung senkrecht zum Mechanismus wird die Exzentrizit�t ber�cksichtigt (immer ung�nstig).
End If
If beff * aeff = 0 Then
    Grundbruch_Rechteck = "Fehler, Fundament hat effektive Fl�che von 0"
    Exit Function
End If
'�berpr�fen von kombinierter Wirkung von Koh�sion bei gleichzeitiger Lastneigung, der Betrag muss definiert sein.
If c > 0 Then
 If Fresb = 0 And (omega - alpha <> 0) Then
  Grundbruch_Rechteck = "F�r c>0 muss infolge Lastneigung zur Sohle ein Kraftbetrag angegeben werden."
  Exit Function
 End If
 R = Fresb 'mit Koh�sion wird der Betragben�tigt falls die Kraft nicht senkrecht auf die Sohle wirkt.
Else
 R = 1 'ohne Koh�sion ist nur die Richtung massgebend
End If
N = R * Cos(omega / 180 * PI - alpha / 180 * PI) 'Normalkraft auf geneigte Sohle
T = R * Sin(omega / 180 * PI - alpha / 180 * PI) 'Tangentialkraft auf geneigte Sohle
'
'--------------------------------------------------------------------------------
'Alle folgenden Faktoren werden nach Lang et al. (8.Auflage) Kapitel 9 berechnet.
'Sie bilden die Grundlage zur erweiterten Tragf�higkeitsformel nach Terzaghi.
'--------------------------------------------------------------------------------
'Tragfaehigkeitsfaktoren
Nq = Exp(PI * Tan(phid)) * Tan(PI / 4 + 0.5 * phid) ^ 2
Ngamma = 1.8 * (Nq - 1) * Tan(phid)
Nc = (Nq - 1) / Tan(phid)
'---------------------------------------------------------------------------------
'Korrekturfaktoren
'Form
If a = "Streifen" Then
    sc = 1
    sq = 1
    sgamma = 1
 Else
    sc = 1 + beff / aeff * Nq / Nc
    sq = 1 + beff / aeff * Tan(phid)
    sgamma = 1 - 0.4 * beff / aeff
End If
'Tiefe
dc = 1 + 0.007 * (Atn(t_soil / beff) * 180 / PI)                                 'Atn von Radiant umrechnen in Grad
dq = 1 + 0.035 * Tan(phid) * (1 - Sin(phid)) ^ 2 * Atn(t_soil / beff) * 180 / PI 'Atn von Radiant umrechnen in Grad
dgamma = 1
'
'dc = 1 + 0.4 * Atn(t_soil / beff)                                 'Gleichung angepasst auf Radiant
'dq = 1 + 2 * Tan(phid) * (1 - Sin(phid)) ^ 2 * Atn(t_soil / beff) 'Gleichung angepasst auf Radiant
'dgamma = 1
'
'Lastneigung
'fuer c<>0 muss auch der Kraftbetrag bekannt sein.
iq = (1 - 0.5 * T / (N + beff * aeff * cd / Tan(phid))) ^ 5
igamma = (1 - (0.7 - alpha / 450) * T / (N + beff * aeff * cd / Tan(phid))) ^ 5
'igamma = (1 - (0.7 - 2 * alpha / (5 * PI)) * T / (N + beff * aeff * cd / Tan(phid))) ^ 5 'Gleichung angepasst auf alpha in Radiant
ic = iq - (1 - iq) / (Nq - 1)
'
'Gelaendeneigung
gc = 1 - beta / 147
gq = (1 - 0.5 * Tan(beta / 180 * PI)) ^ 5
ggamma = gq
'
'Fundamentneigung
fc = 1 - alpha / 147
fq = Exp(-0.035 * alpha * Tan(phid))      'Gleichung wie Doku f�r alpha in grad
fgamma = Exp(-0.047 * alpha * Tan(phid))  'Gleichung wie Doku f�r alpha in grad
'
'fc = 1 - alpha / (0.817 * PI)            'f�r alpha in Radiant
'fq = Exp(-2.00 * alpha * Tan(phid))      'f�r alpha in Radiant
'fgamma = Exp(-2.7 * alpha  * Tan(phid))  'f�r alpha in Radiant
'
sigmaf = cd * Nc * sc * dc * ic * gc * fc + _
        (q_soil) * Nq * sq * dq * iq * gq * fq + _
        0.5 * gammad * beff * Ngamma * sgamma * dgamma * igamma * ggamma * fgamma
Grundbruch_Rechteck = sigmaf * beff * aeff
End Function
Function Grundbruch_H_Rechteck(c, phi, gamma, q_soil, t_soil, b, a, Ed_z, _
Optional eB = 0, Optional eA = 0, _
Optional beta = 0, Optional alpha = 0)
Attribute Grundbruch_H_Rechteck.VB_Description = "Berechnet die zul�ssige Horizontalkraft RT,d f�r eine gegebene Vertikalkraft Ed,z in [kN] aufgrund der angegebenen Bodenparametern, Fundamentgeometrie und Belastungsrichtung"
Attribute Grundbruch_H_Rechteck.VB_ProcData.VB_Invoke_Func = " \n14"
'Gesucht ist die maximal zul�ssig Horizontalkraft f�r eine gegebene Vertikalkraft.
'Dies ist ein Grundbruch- und kein Gleitnachweis.
'Die Berechnung liefert nur eine N�herung.

'   FuncName = "Grundbruch_H_Rechteck"
'   FuncDesc = "Berechnet die zul�ssige Horizontalkraft RT,d f�r eine gegebene Vertikalkraft Ed,z in [kN] " & _
'              "aufgrund der angegebenen Bodenparametern, Fundamentgeometrie und Belastungsrichtung"
'   ArgDesc(1) = "Koh�sion in kPa"
'   ArgDesc(2) = "Reibungswinkel in Grad "
'   ArgDesc(3) = "Effektives Bodengewicht unter Fundament kN/m3 (\gamma f�r trockener Boden, \gamma' bei Grundwasser bis Sohle)"
'   ArgDesc(4) = "Auflast neben Fundament auf Niveau Sohle (inkl. Bodengewicht) in kPa (\gamma t +q)"
'   ArgDesc(5) = "Einbindetiefe (Abstand OKT zur Sohle) in m"
'   ArgDesc(6) = "Breite des Fundaments in Versagensrichtung in m"
'   ArgDesc(7) = "L�nge des Fundaments quer zur Versagensrichtung in m"
'   ArgDesc(8) = "Betrag der Vertikalkraft Ed,z in kN"
'   ArgDesc(9) = "Optional, Exzentrizit�t der Resultierenden in Versagensrichtung in m [Default=0]"
'   ArgDesc(10) = "Optional, Exzentrizit�t der Resultierenden quer zur Versagensrichtung in m [Default=0]"
'   ArgDesc(11) = "Optional, Gel�ndeneigung in Grad [Default=0]"
'   ArgDesc(12) = "Optional, Sohlneigung in Grad [Default=0]"


'Iterationsalgorithmus
' Ez0-> omega0=10�, Rz1=Grundbruch(omega0),Rh1 auf FS=1
' omega1=arctan(Rhi/Rz(i-1)) Rzi+1=Grundbruch(omegai), Rhi+1=Rzi+1*tan(omegai)
Const PI As Double = 3.14159265358979

'startwerte
Rzold = Ed_z
omegai = 10
Rzi = Grundbruch_Rechteck(c, phi, gamma, q_soil, t_soil, b, a, omegai, eB, eA, beta, alpha, Ed_z)
Rhi = Rzi * Tan(omegai * PI / 180)
'iteration
For I = 1 To 20
    omegai = Atn(Rhi / Ed_z) * 180 / PI
        Rzi = Grundbruch_Rechteck(c, phi, gamma, q_soil, t_soil, b, a, omegai, eB, eA, beta, alpha, Ed_z)
    Rhi = Rzi * Tan(omegai / 180 * PI)
Next I
Grundbruch_H_Rechteck = Rhi
End Function

Function Grundbruch_Streifen(c, phi, gamma, q_soil, t_soil, b, _
Optional omega = 0, Optional eB = 0, _
Optional beta = 0, Optional alpha = 0, Optional Fresb = 0)
Attribute Grundbruch_Streifen.VB_Description = "Berechnet die zul�ssige Bodenpressung Rd,N in [kN/m] aufgrund der angegebenenBodenparametern, Fundamentgeometrie und Belastungsrichtung f�r ein unendlich langes Streifen Fundament"
Attribute Grundbruch_Streifen.VB_ProcData.VB_Invoke_Func = " \n14"
'   FuncName = "Grundbruch_Streifen"
'   FuncDesc = "Berechnet die zul�ssige Bodenpressung Rd,N in [kN/m] aufgrund der angegebenen" & _
'              "Bodenparametern, Fundamentgeometrie und Belastungsrichtung f�r ein unendlich langes Streifen Fundament"
'   ArgDesc(1) = "Koh�sion in kPa"
'   ArgDesc(2) = "Reibungswinkel in Grad "
'   ArgDesc(3) = "Effektives Bodengewicht unter Fundament kN/m3 (\gamma f�r trockener Boden, \gamma' bei Grundwasser bis Sohle)"
'   ArgDesc(4) = "Auflast neben Fundament auf Niveau Sohle (inkl. Bodengewicht) in kPa"
'   ArgDesc(5) = "Einbindetiefe (Abstand OKT zur Sohle) in m"
'   ArgDesc(6) = "Breite des Fundaments in Versagensrichtung in m"
'   ArgDesc(7) = "Optional, Abweichung der Kraftrichtung zur Vertikalen in Grad [Default=0]"
'   ArgDesc(8) = "Optional, Exzentrizit�t der Resultierenden in Versagensrichtung in m [Default=0]"
'   ArgDesc(9) = "Optional, Gel�ndeneigung in Grad [Default=0]"
'   ArgDesc(10) = "Optional, Sohlneigung in Grad [Default=0]"
'   ArgDesc(11) = "Optional, Betrag der resultierenden Einwirkung in kN/min Versagensebene, n�tig falls c>0"
Grundbruch_Streifen = Grundbruch_Rechteck(c, phi, gamma, q_soil, t_soil, b, "Streifen", omega, eB, 0, beta, alpha, Fresb)
End Function
Function Grundbruch_H_Streifen(c, phi, gamma, q_soil, t_soil, b, Ed_z, _
Optional eB = 0, Optional beta = 0, Optional alpha = 0)
Attribute Grundbruch_H_Streifen.VB_Description = "Berechnet die zul�ssige Horizontalkraft RT,d f�r eine gegebene Vertikalkraft Ed,z in [kN] aufgrund der angegebenen Bodenparametern, Fundamentgeometrie und Belastungsrichtung"
Attribute Grundbruch_H_Streifen.VB_ProcData.VB_Invoke_Func = " \n14"
'Gesucht ist die maximal zul�ssig Horizontalkraft f�r eine gegebene Vertikalkraft.
'Dies ist ein Grundbruch- und kein Gleitnachweis.
'Die Berechnung liefert nur eine N�herung.

'   FuncName = "Grundbruch_H_Streifen"
'   FuncDesc = "Berechnet die zul�ssige Horizontalkraft RT,d f�r eine gegebene Vertikalkraft Ed,z "& _
'              "in [kN] aufgrund der angegebenen Bodenparametern, Fundamentgeometrie und Belastungsrichtung"
'   ArgDesc(1) = "Koh�sion in kPa"
'   ArgDesc(2) = "Reibungswinkel in Grad "
'   ArgDesc(3) = "Effektives  Bodengewicht unter Fundament kN/m3 (\gamma f�r trockener Boden, \gamma' bei Grundwasser bis Sohle)"
'   ArgDesc(4) = "Auflast neben Fundament auf Niveau Sohle (inkl. Bodengewicht) in kPa (\gamma t +q)"
'   ArgDesc(5) = "Einbindetiefe (Abstand OKT zur Sohle) in m"
'   ArgDesc(6) = "Breite des Fundaments in Versagensrichtung in m"
'   ArgDesc(7) = "Betrag der Vertikalkraft Ed,z in kN"
'   ArgDesc(8) = "Optional, Exzentrizit�t der Resultierenden in Versagensrichtung in m [Default=0]"
'   ArgDesc(9) = "Optional, Gel�ndeneigung in Grad [Default=0]"
'   ArgDesc(10) = "Optional, Sohlneigung in Grad [Default=0]"

'Iterationsalgorithmus
' Ez0-> omega0=10�, Rz1=Grundbruch(omega0),Rh1 auf FS=1
' omega1=arctan(Rhi/Rz(i-1)) Rzi+1=Grundbruch(omegai), Rhi+1=Rzi+1*tan(omegai)
Const PI As Double = 3.14159265358979
'startwerte
Rzold = Ed_z
omegai = 10
Rzi = Grundbruch_Streifen(c, phi, gamma, q_soil, t_soil, b, omegai, eB, beta, alpha, Ed_z)
Rhi = Rzi * Tan(omegai * PI / 180)
'iteration
For I = 1 To 20
    omegai = Atn(Rhi / Ed_z) * 180 / PI
        Rzi = Grundbruch_Streifen(c, phi, gamma, q_soil, t_soil, b, omegai, eB, beta, alpha, Ed_z)
    Rhi = Rzi * Tan(omegai / 180 * PI)
Next I
Grundbruch_H_Streifen = Rhi
End Function
Sub DescribeFunction1()
   Dim FuncName As String
   Dim FuncDesc As String
   Dim Category As String
   Dim ArgDesc(1 To 13) As String

   FuncName = "Grundbruch_Rechteck"
   FuncDesc = "Berechnet die zul�ssige Bodenpressung Rd,N in [kN] aufgrund der angegebenen " & _
              "Bodenparametern, Fundamentgeometrie und Belastungsrichtung"
   ArgDesc(1) = "Koh�sion in kPa"
   ArgDesc(2) = "Reibungswinkel in Grad "
   ArgDesc(3) = "Effektives Bodengewicht unter Fundament kN/m3 (\gamma f�r trockener Boden, \gamma' bei Grundwasser bis Sohle)�"
   ArgDesc(4) = "Auflast neben Fundament auf Niveau Sohle (inkl. Bodengewicht) in kPa, (\gamma t +q)"
   ArgDesc(5) = "Einbindetiefe (Abstand OKT zur Sohle) in m"
   ArgDesc(6) = "Breite des Fundaments in Versagensrichtung in m"
   ArgDesc(7) = "L�nge des Fundaments quer zur Versagensrichtung in m"
   ArgDesc(8) = "Abweichung der Kraftrichtung zur Vertikalen in Grad [Default=0]"
   ArgDesc(9) = "Exzentrizit�t der Resultierenden in Versagensrichtung in m [Default=0]"
   ArgDesc(10) = "Exzentrizit�t der Resultierenden quer zur Versagensrichtung in m [Default=0]"
   ArgDesc(11) = "Gel�ndeneigung in Grad [Default=0]"
   ArgDesc(12) = "Sohlneigung in Grad [Default=0]"
   ArgDesc(13) = "Betrag der resultierenden Einwirkung in Versagensebene, n�tig falls c>0"
   
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
   Dim ArgDesc(1 To 12) As String
   
   FuncName = "Grundbruch_H_Rechteck"
   FuncDesc = "Berechnet die zul�ssige Horizontalkraft RT,d f�r eine gegebene Vertikalkraft Ed,z in [kN] " & _
              "aufgrund der angegebenen Bodenparametern, Fundamentgeometrie und Belastungsrichtung"
   ArgDesc(1) = "Koh�sion in kPa"
   ArgDesc(2) = "Reibungswinkel in Grad "
   ArgDesc(3) = "Effektives Bodengewicht unter Fundament kN/m3 (\gamma f�r trockener Boden, \gamma' bei Grundwasser bis Sohle)"
   ArgDesc(4) = "Auflast neben Fundament auf Niveau Sohle (inkl. Bodengewicht) in kPa (\gamma t +q)"
   ArgDesc(5) = "Einbindetiefe (Abstand OKT zur Sohle) in m"
   ArgDesc(6) = "Breite des Fundaments in Versagensrichtung in m"
   ArgDesc(7) = "L�nge des Fundaments quer zur Versagensrichtung in m"
   ArgDesc(8) = "Betrag der Vertikalkraft Ed,z in kN"
   ArgDesc(9) = "Optional, Exzentrizit�t der Resultierenden in Versagensrichtung in m [Default=0]"
   ArgDesc(10) = "Optional, Exzentrizit�t der Resultierenden quer zur Versagensrichtung in m [Default=0]"
   ArgDesc(11) = "Optional, Gel�ndeneigung in Grad [Default=0]"
   ArgDesc(12) = "Optional, Sohlneigung in Grad [Default=0]"
   
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
   Dim ArgDesc(1 To 11) As String
 
   FuncName = "Grundbruch_Streifen"
   FuncDesc = "Berechnet die zul�ssige Bodenpressung Rd,N in [kN/m] aufgrund der angegebenen" & _
              "Bodenparametern, Fundamentgeometrie und Belastungsrichtung f�r ein unendlich langes Streifen Fundament"
   ArgDesc(1) = "Koh�sion in kPa"
   ArgDesc(2) = "Reibungswinkel in Grad "
   ArgDesc(3) = "Effektives Bodengewicht unter Fundament kN/m3 (\gamma f�r trockener Boden, \gamma' bei Grundwasser bis Sohle)"
   ArgDesc(4) = "Auflast neben Fundament auf Niveau Sohle (inkl. Bodengewicht) in kPa"
   ArgDesc(5) = "Einbindetiefe (Abstand OKT zur Sohle) in m"
   ArgDesc(6) = "Breite des Fundaments in Versagensrichtung in m"
   ArgDesc(7) = "Optional, Abweichung der Kraftrichtung zur Vertikalen in Grad [Default=0]"
   ArgDesc(8) = "Optional, Exzentrizit�t der Resultierenden in Versagensrichtung in m [Default=0]"
   ArgDesc(9) = "Optional, Gel�ndeneigung in Grad [Default=0]"
   ArgDesc(10) = "Optional, Sohlneigung in Grad [Default=0]"
   ArgDesc(11) = "Optional, Betrag der resultierenden Einwirkung in kN/min Versagensebene, n�tig falls c>0"
    
    Category = 14 '14=user defined
    
   Application.MacroOptions _
      Macro:=FuncName, _
      Description:=FuncDesc, _
      Category:=Category, _
      ArgumentDescriptions:=ArgDesc
End Sub

Sub DescribeFunction4()
   Dim FuncName As String
   Dim FuncDesc As String
   Dim Category As String
   Dim ArgDesc(1 To 10) As String

   FuncName = "Grundbruch_H_Streifen"
   FuncDesc = "Berechnet die zul�ssige Horizontalkraft RT,d f�r eine gegebene Vertikalkraft Ed,z in [kN] aufgrund der angegebenen Bodenparametern, Fundamentgeometrie und Belastungsrichtung"
   ArgDesc(1) = "Koh�sion in kPa"
   ArgDesc(2) = "Reibungswinkel in Grad "
   ArgDesc(3) = "Effektives Bodengewicht unter Fundament kN/m3 (\gamma f�r trockener Boden, \gamma' bei Grundwasser bis Sohle)"
   ArgDesc(4) = "Auflast neben Fundament auf Niveau Sohle (inkl. Bodengewicht) in kPa (\gamma t +q)"
   ArgDesc(5) = "Einbindetiefe (Abstand OKT zur Sohle) in m"
   ArgDesc(6) = "Breite des Fundaments in Versagensrichtung in m"
   ArgDesc(7) = "Betrag der Vertikalkraft Ed,z in kN"
   ArgDesc(8) = "Optional, Exzentrizit�t der Resultierenden in Versagensrichtung in m [Default=0]"
   ArgDesc(9) = "Optional, Gel�ndeneigung in Grad [Default=0]"
   ArgDesc(10) = "Optional, Sohlneigung in Grad [Default=0]"
    
   Category = 14 '14=user defined
    
   Application.MacroOptions _
      Macro:=FuncName, _
      Description:=FuncDesc, _
      Category:=Category, _
      ArgumentDescriptions:=ArgDesc
      
End Sub

Function Initialize()
'ThisWorkbook needs a on-open Event that calls this function
    'Sub Workbook_Open()
    'value = BearingCapacityModul.Initialize()
    'End Sub
If initializedState <> "Beschreibung hinzugef�gt" Then
    BearingCapacityModul.DescribeFunction1
    BearingCapacityModul.DescribeFunction2
    BearingCapacityModul.DescribeFunction3
    BearingCapacityModul.DescribeFunction4
    initializedState = "Beschreibung hinzugef�gt"
    someVal = ActiveWorkbook.Worksheets(2).Cells(1, 7).Value
    If Left(someVal, 4) = "Vers" Then
     ActiveWorkbook.Worksheets(2).Cells(1, 7).Value = VersionG
     '     ActiveWorkbook.Worksheets(Rechteckfundament).Cells(1, 7).Value = "=initialize()"
    End If
    someVal = ActiveWorkbook.Worksheets(1).Cells(1, 7).Value
    If Left(someVal, 4) = "Vers" Then
     ActiveWorkbook.Worksheets(1).Cells(1, 7).Value = VersionG
     '     ActiveWorkbook.Worksheets("Streifenfundament").Cells(1, 7).Value = "=initialize()"
    End If
    MsgBox ("Makros initialisiert")
End If
Initialize = VersionG
End Function

