Attribute VB_Name = "BearingCapacityModul"
Function Grundbruch(c, phi, gamma, q_soil, t_soil, B, Optional L = 0, Optional omega = 0, Optional eB = 0, Optional eL = 0, Optional beta = 0, Optional alpha = 0, Optional Fresb = 0)
Attribute Grundbruch.VB_Description = "Berechnet die zulässige Bodenpressung in [kN] aufgrund der angegebenen Bodenparametern, Fundamentgeometrie und Belastungsrichtung"
Attribute Grundbruch.VB_ProcData.VB_Invoke_Func = " \n1"
'   FuncDesc = "Berechnet die zulässige Bodenpressung in [kN] aufgrund der angegebenen Bodenparametern, Fundamentgeometrie und Belastungsrichtung"
'   ArgDesc(1) = "Kohäsion in kPa"
'   ArgDesc(2) = "Reibungswinkel in Grad "
'   ArgDesc(3) = "Bodengewicht unter Fundament kN/m3”
'   ArgDesc(4) = "Auflast neben Fundament auf Niveau Sohle (inkl. Bodengewicht) in kPa"
'   ArgDesc(5) = "Einbindetiefe (Abstand OKT zur Sohle) in m"
'   ArgDesc(6) = "Breite des Fundaments in Versagensrichtung in m"
'   ArgDesc(8) = "Abweichung der Kraftrichtung zur Vertikalen in Grad [Default=0]"
'   ArgDesc(7) = "Länge des Fundaments quer zur Versagensrichtung in m [Default=0]"
'   ArgDesc(9) = "Exzentrizität der Resultierenden in Versagensrichtung in m [Default=0]"
'   ArgDesc(10) = "Exzentrizität der Resultierenden quer zur Versagensrichtung in m [Default=0]"
'   ArgDesc(11) = "Geländeneigung in Grad [Default=0]"
'   ArgDesc(12) = "Sohlneigung in Grad [Default=0]"
'   ArgDesc(13) = "Betrag der resultierenden Einwirkung in Versagensebene, nötig falls c>0"
Const PI As Double = 3.14159265358979
'
'Normparameter
gamma_phi = 1.2
gamma_c = 1.5
gamma_g = 1#
cd = c / gamma_c 'Cohesion design [kPa]
phid = Atn(Tan(phi / 180 * PI) / gamma_phi) 'Friction angle design [rad]
gammad = gamma / gamma_g 'density design [kN/m3]
'Inputparameter

'------ Fundamentneigung, Böschungsneigung und Lastneigung können nur in "b"-Richtung berücksichtigt werden.
'------ alpha, beta und omega sind nur eingeschränkt gültig. Für werte ausserhalb des
'------ Gültigkeitsbereichs werden folgende konservativen Annahmen getroffen
'------ Auch die Exzentrizität infolge Moment oder Verschobener Resultierenden ist nur für ungünstige Kombinationen definiert.
beff = Application.Max(0, B - 2 * Abs(eB))
leff = Application.Max(0, L - 2 * Abs(eL))
If L = 0 Then 'unendliches Streifenfundament-> Output ist in [kN/m]
leff = 1
End If
alpha = Application.Max(0, alpha)
beta = Application.Max(0, beta)
omega = Application.Max(0, omega)

If c > 0 Then
 If Fresb = 0 And omega - alpha = 0 Then
  Grundbau = "Für c>0 muss infolge Lastneigung zur Sohle ein Kraftbetrag angegeben werden."
 End If
 R = Fresb
Else
 R = 1
End If
N = R * Cos(omega / 180 * PI - alpha / 180 * PI) 'Normalkraft auf geneigte Sohle
T = R * Sin(omega / 180 * PI - alpha / 180 * PI) 'Tangentialkraft auf geneigte Sohle
'Tragfaehigkeitsfaktoren
Nq = Exp(PI * Tan(phid)) * Tan(PI / 4 + 0.5 * phid) ^ 2
Ng = 1.8 * (Nq - 1) * Tan(phid)
Nc = (Nq - 1) / Tan(phid)
'Korrekturfaktoren
'Form
If L = 0 Then
sc = 1
sq = 1
sg = 1
Else: 'Use leff
sc = 1 + beff / leff * Nq / Ng
sq = 1 + beff / leff * Tan(phid)
sg = 1 - 0.4 * beff / leff
End If
'Tiefe
dc = 1 + 0.007 * 180 / PI * Atn(t_soil / beff)
dq = 1 + 0.035 * 180 / PI * Tan(phid) * (1 - Sin(phid)) ^ 2 * Atn(t_soil / beff)
dg = 1
'Lastneigung
'Use leff2, fuer c<>0 muss auch der Kraftbetrag bekannt sein.
iq = (1 - 0.5 * T / (N + beff * leff * cd / Tan(phid))) ^ 5
ig = (1 - (0.7 - alpha / 450) * T / (N + beff * leff * cd / Tan(phid))) ^ 5
ic = 1 - (1 - iq) / (Nq - 1)
'Gelaendeneigung
gc = 1 - beta / 147
gq = (1 - 0.5 * Tan(beta / 180 * PI)) ^ 5
gg = gq
'Fundamentneigung
bdc = 1 - alpha / 147
bdq = Exp(-2.005 * alpha / 180 * PI * Tan(phid))
bdg = Exp(-2.693 * alpha / 180 * PI * Tan(phid))

sigma_f = cd * Nc * sc * dc * ic * gc * bdc + (q_soil) * Nq * sq * dq * iq * gq * bdq + 0.5 * gammad * Ng * sg * dg * ig * gg * bdg
Grundbruch = sigma_f * beff * leff
End Function
Sub DescribeFunction1()
   Dim FuncName As String
   Dim FuncDesc As String
   Dim Category As String
   Dim ArgDesc(1 To 13) As String
    FuncName = "Grundbruch"
    
   FuncDesc = "Berechnet die zulässige Bodenpressung in [kN] aufgrund der angegebenen Bodenparametern, Fundamentgeometrie und Belastungsrichtung"
   ArgDesc(1) = "Kohäsion in kPa"
   ArgDesc(2) = "Reibungswinkel in Grad "
   ArgDesc(3) = "Bodengewicht unter Fundament kN/m3”"
   ArgDesc(4) = "Auflast neben Fundament auf Niveau Sohle (inkl. Bodengewicht) in kPa"
   ArgDesc(5) = "Einbindetiefe (Abstand OKT zur Sohle) in m"
   ArgDesc(6) = "Breite des Fundaments in Versagensrichtung in m"
   ArgDesc(8) = "Abweichung der Kraftrichtung zur Vertikalen in Grad [Default=0]"
   ArgDesc(7) = "Länge des Fundaments quer zur Versagensrichtung in m [Default=0]"
   ArgDesc(9) = "Exzentrizität der Resultierenden in Versagensrichtung in m [Default=0]"
   ArgDesc(10) = "Exzentrizität der Resultierenden quer zur Versagensrichtung in m [Default=0]"
   ArgDesc(11) = "Geländeneigung in Grad [Default=0]"
   ArgDesc(12) = "Sohlneigung in Grad [Default=0]"
   ArgDesc(13) = "Betrag der resultierenden Einwirkung in Versagensebene, nötig falls c>0"
    
    
    Category = 1 '1=??
   Application.MacroOptions _
      Macro:=FuncName, _
      Description:=FuncDesc, _
      Category:=Category, _
      ArgumentDescriptions:=ArgDesc
End Sub

Function initialize(ist)
 If ist <> "Beschreibung hinzugefügt" Then
  DescribeFunction1
  initialize = "Beschreibung hinzugefügt"
 Else
  initialize = "Beschreibung bereits vorhanden"
 End If
End Function
