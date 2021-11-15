Attribute VB_Name = "SistemaCombate"
'Argentum Online 0.11.6
'Copyright (C) 2002 M�rquez Pablo Ignacio
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 n�mero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'C�digo Postal 1900
'Pablo Ignacio M�rquez
'
'Dise�o y correcci�n del modulo de combate por
'Gerardo Saiz, gerardosaiz@yahoo.com
'

Option Explicit

Public Const MAXDISTANCIAARCO As Byte = 18
Public Const MAXDISTANCIAMAGIA As Byte = 18

Function ModificadorEvasion(ByVal clase As eClass) As Single

Select Case clase
    Case eClass.Warrior
        ModificadorEvasion = 1
    Case eClass.Hunter
        ModificadorEvasion = 0.9
    Case eClass.Paladin
        ModificadorEvasion = 0.9
    Case eClass.Bandit
        ModificadorEvasion = 0.9
    Case eClass.Assasin
        ModificadorEvasion = 1.1
    Case eClass.Pirat
        ModificadorEvasion = 0.9
    Case eClass.Thief
        ModificadorEvasion = 1.1
    Case eClass.Bard
        ModificadorEvasion = 1.1
    Case eClass.Mage
        ModificadorEvasion = 0.4
    Case eClass.Druid
        ModificadorEvasion = 0.75
    Case Else
        ModificadorEvasion = 0.8
End Select
End Function

Function ModificadorPoderAtaqueArmas(ByVal clase As eClass) As Single
Select Case UCase$(clase)
    Case eClass.Warrior
        ModificadorPoderAtaqueArmas = 1
    Case eClass.Paladin
        ModificadorPoderAtaqueArmas = 0.9
    Case eClass.Hunter
        ModificadorPoderAtaqueArmas = 0.8
    Case eClass.Assasin
        ModificadorPoderAtaqueArmas = 0.85
    Case eClass.Pirat
        ModificadorPoderAtaqueArmas = 0.8
    Case eClass.Thief
        ModificadorPoderAtaqueArmas = 0.75
    Case eClass.Bandit
        ModificadorPoderAtaqueArmas = 0.7
    Case eClass.Cleric
        ModificadorPoderAtaqueArmas = 0.75
    Case eClass.Bard
        ModificadorPoderAtaqueArmas = 0.7
    Case eClass.Druid
        ModificadorPoderAtaqueArmas = 0.65
    Case eClass.Fisher
        ModificadorPoderAtaqueArmas = 0.6
    Case eClass.Lumberjack
        ModificadorPoderAtaqueArmas = 0.6
    Case eClass.Miner
        ModificadorPoderAtaqueArmas = 0.6
    Case eClass.Blacksmith
        ModificadorPoderAtaqueArmas = 0.6
    Case eClass.Carpenter
        ModificadorPoderAtaqueArmas = 0.6
    Case Else
        ModificadorPoderAtaqueArmas = 0.5
End Select
End Function

Function ModificadorPoderAtaqueProyectiles(ByVal clase As eClass) As Single
Select Case UCase$(clase)
    Case eClass.Warrior
        ModificadorPoderAtaqueProyectiles = 0.8
    Case eClass.Hunter
        ModificadorPoderAtaqueProyectiles = 1
    Case eClass.Paladin
        ModificadorPoderAtaqueProyectiles = 0.75
    Case eClass.Assasin
        ModificadorPoderAtaqueProyectiles = 0.75
    Case eClass.Pirat
        ModificadorPoderAtaqueProyectiles = 0.75
    Case eClass.Thief
        ModificadorPoderAtaqueProyectiles = 0.8
    Case eClass.Bandit
        ModificadorPoderAtaqueProyectiles = 0.75
    Case eClass.Cleric
        ModificadorPoderAtaqueProyectiles = 0.7
    Case eClass.Bard
        ModificadorPoderAtaqueProyectiles = 0.7
    Case eClass.Druid
        ModificadorPoderAtaqueProyectiles = 0.75
    Case eClass.Fisher
        ModificadorPoderAtaqueProyectiles = 0.65
    Case eClass.Lumberjack
        ModificadorPoderAtaqueProyectiles = 0.7
    Case eClass.Miner
        ModificadorPoderAtaqueProyectiles = 0.65
    Case eClass.Blacksmith
        ModificadorPoderAtaqueProyectiles = 0.65
    Case eClass.Carpenter
        ModificadorPoderAtaqueProyectiles = 0.7
    Case Else
        ModificadorPoderAtaqueProyectiles = 0.5
End Select
End Function

Function ModicadorDa�oClaseArmas(ByVal clase As eClass) As Single
Select Case UCase$(clase)
    Case eClass.Warrior
        ModicadorDa�oClaseArmas = 1.1
    Case eClass.Paladin
        ModicadorDa�oClaseArmas = 0.95
    Case eClass.Hunter
        ModicadorDa�oClaseArmas = 0.9
    Case eClass.Assasin
        ModicadorDa�oClaseArmas = 0.9
    Case eClass.Thief
        ModicadorDa�oClaseArmas = 0.8
    Case eClass.Pirat
        ModicadorDa�oClaseArmas = 0.8
    Case eClass.Bandit
        ModicadorDa�oClaseArmas = 1
    Case eClass.Cleric
        ModicadorDa�oClaseArmas = 0.8
    Case eClass.Bard
        ModicadorDa�oClaseArmas = 0.75
    Case eClass.Druid
        ModicadorDa�oClaseArmas = 0.7
    Case eClass.Fisher
        ModicadorDa�oClaseArmas = 0.6
    Case eClass.Lumberjack
        ModicadorDa�oClaseArmas = 0.7
    Case eClass.Miner
        ModicadorDa�oClaseArmas = 0.75
    Case eClass.Blacksmith
        ModicadorDa�oClaseArmas = 0.75
    Case eClass.Carpenter
        ModicadorDa�oClaseArmas = 0.7
    Case Else
        ModicadorDa�oClaseArmas = 0.5
End Select
End Function

Function ModicadorDa�oClaseWrestling(ByVal clase As eClass) As Single
'Pablo (ToxicWaste): Esto en proxima versi�n habr� que balancearlo para cada clase
'Hoy por hoy est� solo hecho para el bandido.
Select Case UCase$(clase)
    Case eClass.Warrior
        ModicadorDa�oClaseWrestling = 0.4
    Case eClass.Paladin
        ModicadorDa�oClaseWrestling = 0.4
    Case eClass.Hunter
        ModicadorDa�oClaseWrestling = 0.4
    Case eClass.Assasin
        ModicadorDa�oClaseWrestling = 0.4
    Case eClass.Thief
        ModicadorDa�oClaseWrestling = 0.4
    Case eClass.Pirat
        ModicadorDa�oClaseWrestling = 0.4
    Case eClass.Bandit
        ModicadorDa�oClaseWrestling = 1.1
    Case eClass.Cleric
        ModicadorDa�oClaseWrestling = 0.4
    Case eClass.Bard
        ModicadorDa�oClaseWrestling = 0.4
    Case eClass.Druid
        ModicadorDa�oClaseWrestling = 0.4
    Case eClass.Fisher
        ModicadorDa�oClaseWrestling = 0.4
    Case eClass.Lumberjack
        ModicadorDa�oClaseWrestling = 0.4
    Case eClass.Miner
        ModicadorDa�oClaseWrestling = 0.4
    Case eClass.Blacksmith
        ModicadorDa�oClaseWrestling = 0.4
    Case eClass.Carpenter
        ModicadorDa�oClaseWrestling = 0.4
    Case Else
        ModicadorDa�oClaseWrestling = 0.4
End Select
End Function


Function ModicadorDa�oClaseProyectiles(ByVal clase As eClass) As Single
Select Case clase
    Case eClass.Hunter
        ModicadorDa�oClaseProyectiles = 1.1
    Case eClass.Warrior
        ModicadorDa�oClaseProyectiles = 0.9
    Case eClass.Paladin
        ModicadorDa�oClaseProyectiles = 0.8
    Case eClass.Assasin
        ModicadorDa�oClaseProyectiles = 0.8
    Case eClass.Thief
        ModicadorDa�oClaseProyectiles = 0.75
    Case eClass.Pirat
        ModicadorDa�oClaseProyectiles = 0.75
    Case eClass.Bandit
        ModicadorDa�oClaseProyectiles = 0.8
    Case eClass.Cleric
        ModicadorDa�oClaseProyectiles = 0.7
    Case eClass.Bard
        ModicadorDa�oClaseProyectiles = 0.7
    Case eClass.Druid
        ModicadorDa�oClaseProyectiles = 0.75
    Case eClass.Fisher
        ModicadorDa�oClaseProyectiles = 0.6
    Case eClass.Lumberjack
        ModicadorDa�oClaseProyectiles = 0.7
    Case eClass.Miner
        ModicadorDa�oClaseProyectiles = 0.6
    Case eClass.Blacksmith
        ModicadorDa�oClaseProyectiles = 0.6
    Case eClass.Carpenter
        ModicadorDa�oClaseProyectiles = 0.7
    Case Else
        ModicadorDa�oClaseProyectiles = 0.5
End Select
End Function

Function ModEvasionDeEscudoClase(ByVal clase As eClass) As Single

Select Case clase
    Case eClass.Warrior
        ModEvasionDeEscudoClase = 1
    Case eClass.Hunter
        ModEvasionDeEscudoClase = 0.8
    Case eClass.Paladin
        ModEvasionDeEscudoClase = 1
    Case eClass.Assasin
        ModEvasionDeEscudoClase = 0.8
    Case eClass.Thief
        ModEvasionDeEscudoClase = 0.7
    Case eClass.Bandit
        ModEvasionDeEscudoClase = 2
    Case eClass.Pirat
        ModEvasionDeEscudoClase = 0.75
    Case eClass.Cleric
        ModEvasionDeEscudoClase = 0.85
    Case eClass.Bard
        ModEvasionDeEscudoClase = 0.8
    Case eClass.Druid
        ModEvasionDeEscudoClase = 0.75
    Case eClass.Fisher
        ModEvasionDeEscudoClase = 0.7
    Case eClass.Lumberjack
        ModEvasionDeEscudoClase = 0.7
    Case eClass.Miner
        ModEvasionDeEscudoClase = 0.7
    Case eClass.Blacksmith
        ModEvasionDeEscudoClase = 0.7
    Case eClass.Carpenter
        ModEvasionDeEscudoClase = 0.7
    Case Else
        ModEvasionDeEscudoClase = 0.6
End Select

End Function
Function Minimo(ByVal a As Single, ByVal b As Single) As Single
If a > b Then
    Minimo = b
    Else: Minimo = a
End If
End Function

Function MinimoInt(ByVal a As Integer, ByVal b As Integer) As Integer
If a > b Then
    MinimoInt = b
    Else: MinimoInt = a
End If
End Function

Function Maximo(ByVal a As Single, ByVal b As Single) As Single
If a > b Then
    Maximo = a
    Else: Maximo = b
End If
End Function

Function MaximoInt(ByVal a As Integer, ByVal b As Integer) As Integer
If a > b Then
    MaximoInt = a
    Else: MaximoInt = b
End If
End Function


Function PoderEvasionEscudo(ByVal UserIndex As Integer) As Long

PoderEvasionEscudo = (UserList(UserIndex).Stats.UserSkills(eSkill.Defensa) * _
ModEvasionDeEscudoClase(UserList(UserIndex).clase)) / 2

End Function

Function PoderEvasion(ByVal UserIndex As Integer) As Long
    Dim lTemp As Long
     With UserList(UserIndex)
       lTemp = (.Stats.UserSkills(eSkill.Tacticas) + _
          .Stats.UserSkills(eSkill.Tacticas) / 33 * .Stats.UserAtributos(eAtributos.Agilidad)) * _
          ModificadorEvasion(.clase)
       
        PoderEvasion = (lTemp + (2.5 * Maximo(CInt(.Stats.ELV) - 12, 0)))
    End With
End Function

Function PoderAtaqueArma(ByVal UserIndex As Integer) As Long
Dim PoderAtaqueTemp As Long

If UserList(UserIndex).Stats.UserSkills(eSkill.Armas) < 31 Then
    PoderAtaqueTemp = (UserList(UserIndex).Stats.UserSkills(eSkill.Armas) * _
    ModificadorPoderAtaqueArmas(UserList(UserIndex).clase))
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Armas) < 61 Then
    PoderAtaqueTemp = ((UserList(UserIndex).Stats.UserSkills(eSkill.Armas) + _
    UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad)) * _
    ModificadorPoderAtaqueArmas(UserList(UserIndex).clase))
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Armas) < 91 Then
    PoderAtaqueTemp = ((UserList(UserIndex).Stats.UserSkills(eSkill.Armas) + _
    (2 * UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad))) * _
    ModificadorPoderAtaqueArmas(UserList(UserIndex).clase))
Else
   PoderAtaqueTemp = ((UserList(UserIndex).Stats.UserSkills(eSkill.Armas) + _
   (3 * UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad))) * _
   ModificadorPoderAtaqueArmas(UserList(UserIndex).clase))
End If

PoderAtaqueArma = (PoderAtaqueTemp + (2.5 * Maximo(CInt(UserList(UserIndex).Stats.ELV) - 12, 0)))
End Function

Function PoderAtaqueProyectil(ByVal UserIndex As Integer) As Long
Dim PoderAtaqueTemp As Long

If UserList(UserIndex).Stats.UserSkills(eSkill.Proyectiles) < 31 Then
    PoderAtaqueTemp = (UserList(UserIndex).Stats.UserSkills(eSkill.Proyectiles) * _
    ModificadorPoderAtaqueProyectiles(UserList(UserIndex).clase))
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Proyectiles) < 61 Then
        PoderAtaqueTemp = ((UserList(UserIndex).Stats.UserSkills(eSkill.Proyectiles) + _
        UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad)) * _
        ModificadorPoderAtaqueProyectiles(UserList(UserIndex).clase))
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Proyectiles) < 91 Then
        PoderAtaqueTemp = ((UserList(UserIndex).Stats.UserSkills(eSkill.Proyectiles) + _
        (2 * UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad))) * _
        ModificadorPoderAtaqueProyectiles(UserList(UserIndex).clase))
Else
       PoderAtaqueTemp = ((UserList(UserIndex).Stats.UserSkills(eSkill.Proyectiles) + _
      (3 * UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad))) * _
      ModificadorPoderAtaqueProyectiles(UserList(UserIndex).clase))
End If

PoderAtaqueProyectil = (PoderAtaqueTemp + (2.5 * Maximo(CInt(UserList(UserIndex).Stats.ELV) - 12, 0)))

End Function

Function PoderAtaqueWrestling(ByVal UserIndex As Integer) As Long
Dim PoderAtaqueTemp As Long

If UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) < 31 Then
    PoderAtaqueTemp = (UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) * _
    ModificadorPoderAtaqueArmas(UserList(UserIndex).clase))
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) < 61 Then
        PoderAtaqueTemp = ((UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) + _
        UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad)) * _
        ModificadorPoderAtaqueArmas(UserList(UserIndex).clase))
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) < 91 Then
        PoderAtaqueTemp = ((UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) + _
        (2 * UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad))) * _
        ModificadorPoderAtaqueArmas(UserList(UserIndex).clase))
Else
       PoderAtaqueTemp = ((UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) + _
       (3 * UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad))) * _
       ModificadorPoderAtaqueArmas(UserList(UserIndex).clase))
End If

PoderAtaqueWrestling = (PoderAtaqueTemp + (2.5 * Maximo(CInt(UserList(UserIndex).Stats.ELV) - 12, 0)))

End Function


Public Function UserImpactoNpc(ByVal UserIndex As Integer, ByVal NpcIndex As Integer) As Boolean
Dim PoderAtaque As Long
Dim Arma As Integer
Dim proyectil As Boolean
Dim ProbExito As Long

Arma = UserList(UserIndex).Invent.WeaponEqpObjIndex
If Arma = 0 Then proyectil = False Else proyectil = ObjData(Arma).proyectil = 1

If Arma > 0 Then 'Usando un arma
    If proyectil Then
        PoderAtaque = PoderAtaqueProyectil(UserIndex)
    Else
        PoderAtaque = PoderAtaqueArma(UserIndex)
    End If
Else 'Peleando con pu�os
    PoderAtaque = PoderAtaqueWrestling(UserIndex)
End If


ProbExito = Maximo(10, Minimo(90, 50 + ((PoderAtaque - Npclist(NpcIndex).PoderEvasion) * 0.4)))

UserImpactoNpc = (RandomNumber(1, 100) <= ProbExito)

If UserImpactoNpc Then
    If Arma <> 0 Then
       If proyectil Then
            Call SubirSkill(UserIndex, Proyectiles)
       Else
            Call SubirSkill(UserIndex, Armas)
       End If
    Else
        Call SubirSkill(UserIndex, Wrestling)
    End If
End If


End Function

Public Function NpcImpacto(ByVal NpcIndex As Integer, ByVal UserIndex As Integer) As Boolean
'*************************************************
'Author: Unknown
'Last modified: 03/15/2006
'Revisa si un NPC logra impactar a un user o no
'03/15/2006 Maraxus - Evit� una divisi�n por cero que eliminaba NPCs
'*************************************************
Dim Rechazo As Boolean
Dim ProbRechazo As Long
Dim ProbExito As Long
Dim UserEvasion As Long
Dim NpcPoderAtaque As Long
Dim PoderEvasioEscudo As Long
Dim SkillTacticas As Long
Dim SkillDefensa As Long

UserEvasion = PoderEvasion(UserIndex)
NpcPoderAtaque = Npclist(NpcIndex).PoderAtaque
PoderEvasioEscudo = PoderEvasionEscudo(UserIndex)

SkillTacticas = UserList(UserIndex).Stats.UserSkills(eSkill.Tacticas)
SkillDefensa = UserList(UserIndex).Stats.UserSkills(eSkill.Defensa)

'Esta usando un escudo ???
If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then UserEvasion = UserEvasion + PoderEvasioEscudo

ProbExito = Maximo(10, Minimo(90, 50 + ((NpcPoderAtaque - UserEvasion) * 0.4)))

NpcImpacto = (RandomNumber(1, 100) <= ProbExito)

' el usuario esta usando un escudo ???
If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then
    If Not NpcImpacto Then
        If SkillDefensa + SkillTacticas > 0 Then  'Evitamos divisi�n por cero
            ProbRechazo = Maximo(10, Minimo(90, 100 * (SkillDefensa / (SkillDefensa + SkillTacticas))))
            Rechazo = (RandomNumber(1, 100) <= ProbRechazo)
            If Rechazo = True Then
                'Se rechazo el ataque con el escudo
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_ESCUDO))
                Call WriteBlockedWithShieldUser(UserIndex)
                Call SubirSkill(UserIndex, Defensa)
            End If
        End If
    End If
End If
End Function

Public Function CalcularDa�o(ByVal UserIndex As Integer, Optional ByVal NpcIndex As Integer = 0) As Long
Dim Da�oArma As Long, Da�oUsuario As Long, Arma As ObjData, ModifClase As Single
Dim proyectil As ObjData
Dim Da�oMaxArma As Long

''sacar esto si no queremos q la matadracos mate el Dragon si o si
Dim matoDragon As Boolean
matoDragon = False


If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
    Arma = ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex)
    
    
    ' Ataca a un npc?
    If NpcIndex > 0 Then
        
        'Usa la mata Dragones?
        If UserList(UserIndex).Invent.WeaponEqpObjIndex = EspadaMataDragonesIndex Then ' Usa la mataDragones?
            ModifClase = ModicadorDa�oClaseArmas(UserList(UserIndex).clase)
            
            If Npclist(NpcIndex).NPCtype = DRAGON Then 'Ataca Dragon?
                Da�oArma = RandomNumber(Arma.MinHIT, Arma.MaxHIT)
                Da�oMaxArma = Arma.MaxHIT
                matoDragon = True ''sacar esto si no queremos q la matadracos mate el Dragon si o si
            Else ' Sino es Dragon da�o es 1
                Da�oArma = 1
                Da�oMaxArma = 1
            End If
        Else ' da�o comun
           If Arma.proyectil = 1 Then
                ModifClase = ModicadorDa�oClaseProyectiles(UserList(UserIndex).clase)
                Da�oArma = RandomNumber(Arma.MinHIT, Arma.MaxHIT)
                Da�oMaxArma = Arma.MaxHIT
                If Arma.Municion = 1 Then
                    proyectil = ObjData(UserList(UserIndex).Invent.MunicionEqpObjIndex)
                    Da�oArma = Da�oArma + RandomNumber(proyectil.MinHIT, proyectil.MaxHIT)
                    Da�oMaxArma = Arma.MaxHIT
                End If
           Else
                ModifClase = ModicadorDa�oClaseArmas(UserList(UserIndex).clase)
                Da�oArma = RandomNumber(Arma.MinHIT, Arma.MaxHIT)
                Da�oMaxArma = Arma.MaxHIT
           End If
        End If
    
    Else ' Ataca usuario
        If UserList(UserIndex).Invent.WeaponEqpObjIndex = EspadaMataDragonesIndex Then
            ModifClase = ModicadorDa�oClaseArmas(UserList(UserIndex).clase)
            Da�oArma = 1 ' Si usa la espada mataDragones da�o es 1
            Da�oMaxArma = 1
        Else
           If Arma.proyectil = 1 Then
                ModifClase = ModicadorDa�oClaseProyectiles(UserList(UserIndex).clase)
                Da�oArma = RandomNumber(Arma.MinHIT, Arma.MaxHIT)
                Da�oMaxArma = Arma.MaxHIT
                
                If Arma.Municion = 1 Then
                    proyectil = ObjData(UserList(UserIndex).Invent.MunicionEqpObjIndex)
                    Da�oArma = Da�oArma + RandomNumber(proyectil.MinHIT, proyectil.MaxHIT)
                    Da�oMaxArma = Arma.MaxHIT
                End If
           Else
                ModifClase = ModicadorDa�oClaseArmas(UserList(UserIndex).clase)
                Da�oArma = RandomNumber(Arma.MinHIT, Arma.MaxHIT)
                Da�oMaxArma = Arma.MaxHIT
           End If
        End If
    End If
Else
    'Pablo (ToxicWaste)
    ModifClase = ModicadorDa�oClaseWrestling(UserList(UserIndex).clase)
    Da�oArma = RandomNumber(1, 3) 'Hacemos que sea "tipo" una daga el ataque de Wrestling
    Da�oMaxArma = 3
End If

Da�oUsuario = RandomNumber(UserList(UserIndex).Stats.MinHIT, UserList(UserIndex).Stats.MaxHIT)

''sacar esto si no queremos q la matadracos mate el Dragon si o si
If matoDragon Then
    CalcularDa�o = Npclist(NpcIndex).Stats.MinHP + Npclist(NpcIndex).Stats.def
Else
    CalcularDa�o = (((3 * Da�oArma) + ((Da�oMaxArma / 5) * Maximo(0, (UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) - 15))) + Da�oUsuario) * ModifClase)
End If

End Function

Public Sub UserDa�oNpc(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)

Dim da�o As Long

da�o = CalcularDa�o(UserIndex, NpcIndex)

'esta navegando? si es asi le sumamos el da�o del barco
If UserList(UserIndex).flags.Navegando = 1 And UserList(UserIndex).Invent.BarcoObjIndex > 0 Then _
        da�o = da�o + RandomNumber(ObjData(UserList(UserIndex).Invent.BarcoObjIndex).MinHIT, ObjData(UserList(UserIndex).Invent.BarcoObjIndex).MaxHIT)

If UserList(UserIndex).flags.Montado = 1 And UserList(UserIndex).Invent.MonturaObjIndex > 0 Then _
        da�o = da�o + RandomNumber(ObjData(UserList(UserIndex).Invent.MonturaObjIndex).MinHIT, ObjData(UserList(UserIndex).Invent.MonturaObjIndex).MaxHIT)

da�o = da�o - Npclist(NpcIndex).Stats.def

If da�o < 0 Then da�o = 0

'[KEVIN]
Call WriteUserHitNPC(UserIndex, da�o)
Call CalcularDarExp(UserIndex, NpcIndex, da�o)
Npclist(NpcIndex).Stats.MinHP = Npclist(NpcIndex).Stats.MinHP - da�o
'[/KEVIN]

If Npclist(NpcIndex).Stats.MinHP > 0 Then
    'Trata de apu�alar por la espalda al enemigo
    If PuedeApu�alar(UserIndex) Then
       Call DoApu�alar(UserIndex, NpcIndex, 0, da�o)
       Call SubirSkill(UserIndex, Apu�alar)
    End If
    'trata de dar golpe cr�tico
    Call DoGolpeCritico(UserIndex, NpcIndex, 0, da�o)
    
End If

 
If Npclist(NpcIndex).Stats.MinHP <= 0 Then
        
        ' Si era un Dragon perdemos la espada mataDragones
        If Npclist(NpcIndex).NPCtype = DRAGON Then
            'Si tiene equipada la matadracos se la sacamos
            If UserList(UserIndex).Invent.WeaponEqpObjIndex = EspadaMataDragonesIndex Then
                Call QuitarObjetos(EspadaMataDragonesIndex, 1, UserIndex)
            End If
            If Npclist(NpcIndex).Stats.MaxHP > 100000 Then Call LogDesarrollo(UserList(UserIndex).Name & " mat� un drag�n")
        End If
        
        
        ' Para que las mascotas no sigan intentando luchar y
        ' comiencen a seguir al amo
        
        Dim j As Integer
        For j = 1 To MAXMASCOTAS
            If UserList(UserIndex).MascotasIndex(j) > 0 Then
                If Npclist(UserList(UserIndex).MascotasIndex(j)).TargetNPC = NpcIndex Then Npclist(UserList(UserIndex).MascotasIndex(j)).TargetNPC = 0
                Npclist(UserList(UserIndex).MascotasIndex(j)).Movement = TipoAI.SigueAmo
            End If
        Next j
        
        Call MuereNpc(NpcIndex, UserIndex)
End If

End Sub


Public Sub NpcDa�o(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)

Dim da�o As Integer, Lugar As Integer, absorbido As Integer, npcfile As String
Dim antda�o As Integer, defbarco As Integer
Dim Obj As ObjData



da�o = RandomNumber(Npclist(NpcIndex).Stats.MinHIT, Npclist(NpcIndex).Stats.MaxHIT)
antda�o = da�o

If UserList(UserIndex).flags.Navegando = 1 And UserList(UserIndex).Invent.BarcoObjIndex > 0 Then
    Obj = ObjData(UserList(UserIndex).Invent.BarcoObjIndex)
    defbarco = RandomNumber(Obj.MinDef, Obj.MaxDef)
End If

Dim defMontura As Integer

If UserList(UserIndex).flags.Montado = 1 And UserList(UserIndex).Invent.MonturaObjIndex > 0 Then
    Obj = ObjData(UserList(UserIndex).Invent.MonturaObjIndex)
    defMontura = RandomNumber(Obj.MinDef, Obj.MaxDef)
End If

Lugar = RandomNumber(1, 6)


Select Case Lugar
  Case PartesCuerpo.bCabeza
        'Si tiene casco absorbe el golpe
        If UserList(UserIndex).Invent.CascoEqpObjIndex > 0 Then
           Obj = ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex)
           absorbido = RandomNumber(Obj.MinDef, Obj.MaxDef)
           absorbido = absorbido + defbarco
           da�o = da�o - absorbido
           If da�o < 1 Then da�o = 1
        End If
  Case Else
        'Si tiene armadura absorbe el golpe
        If UserList(UserIndex).Invent.ArmourEqpObjIndex > 0 Then
           Dim Obj2 As ObjData
           Obj = ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex)
           If UserList(UserIndex).Invent.EscudoEqpObjIndex Then
                Obj2 = ObjData(UserList(UserIndex).Invent.EscudoEqpObjIndex)
                absorbido = RandomNumber(Obj.MinDef + Obj2.MinDef, Obj.MaxDef + Obj2.MaxDef)
           Else
                absorbido = RandomNumber(Obj.MinDef, Obj.MaxDef)
           End If
           absorbido = absorbido + defbarco
           da�o = da�o - absorbido
           If da�o < 1 Then da�o = 1
        End If
End Select

Call WriteNPCHitUser(UserIndex, Lugar, da�o)

If UserList(UserIndex).flags.Privilegios And PlayerType.User Then UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MinHP - da�o

If UserList(UserIndex).flags.Meditando Then
    If da�o > Fix(UserList(UserIndex).Stats.MinHP / 100 * UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia) * UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) / 100 * 12 / (RandomNumber(0, 5) + 7)) Then
        UserList(UserIndex).flags.Meditando = False
        Call WriteMeditateToggle(UserIndex)
        Call WriteConsoleMsg(UserIndex, "Dejas de meditar.", FontTypeNames.FONTTYPE_INFO)
        UserList(UserIndex).Char.FX = 0
        UserList(UserIndex).Char.loops = 0
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.CharIndex, 0, 0))
    End If
End If

'Muere el usuario
If UserList(UserIndex).Stats.MinHP <= 0 Then

    Call WriteNPCKillUser(UserIndex) ' Le informamos que ha muerto ;)
    
    If Npclist(NpcIndex).MaestroUser > 0 Then
        Call AllFollowAmo(Npclist(NpcIndex).MaestroUser)
    Else
        'Al matarlo no lo sigue mas
        If Npclist(NpcIndex).Stats.Alineacion = 0 Then
                    Npclist(NpcIndex).Movement = Npclist(NpcIndex).flags.OldMovement
                    Npclist(NpcIndex).Hostile = Npclist(NpcIndex).flags.OldHostil
                    Npclist(NpcIndex).flags.AttackedBy = vbNullString
        End If
    End If
    
    
    Call UserDie(UserIndex)

End If

End Sub



Public Sub CheckPets(ByVal NpcIndex As Integer, ByVal UserIndex As Integer, Optional ByVal CheckElementales As Boolean = True)

Dim j As Integer
For j = 1 To MAXMASCOTAS
    If UserList(UserIndex).MascotasIndex(j) > 0 Then
       If UserList(UserIndex).MascotasIndex(j) <> NpcIndex Then
        If CheckElementales Or (Npclist(UserList(UserIndex).MascotasIndex(j)).Numero <> ELEMENTALFUEGO And Npclist(UserList(UserIndex).MascotasIndex(j)).Numero <> ELEMENTALTIERRA) Then
            If Npclist(UserList(UserIndex).MascotasIndex(j)).TargetNPC = 0 Then Npclist(UserList(UserIndex).MascotasIndex(j)).TargetNPC = NpcIndex
            'Npclist(UserList(UserIndex).MascotasIndex(j)).Flags.OldMovement = Npclist(UserList(UserIndex).MascotasIndex(j)).Movement
            Npclist(UserList(UserIndex).MascotasIndex(j)).Movement = TipoAI.NpcAtacaNpc
        End If
       End If
    End If
Next j

End Sub
Public Sub AllFollowAmo(ByVal UserIndex As Integer)
Dim j As Integer
For j = 1 To MAXMASCOTAS
    If UserList(UserIndex).MascotasIndex(j) > 0 Then
        Call FollowAmo(UserList(UserIndex).MascotasIndex(j))
    End If
Next j
End Sub

Public Function NpcAtacaUser(ByVal NpcIndex As Integer, ByVal UserIndex As Integer) As Boolean

If UserList(UserIndex).flags.AdminInvisible = 1 Then Exit Function
If (Not UserList(UserIndex).flags.Privilegios And PlayerType.User) <> 0 And Not UserList(UserIndex).flags.AdminPerseguible Then Exit Function

' El npc puede atacar ???
If Npclist(NpcIndex).CanAttack = 1 Then
    NpcAtacaUser = True
    Call CheckPets(NpcIndex, UserIndex, False)

    If Npclist(NpcIndex).Target = 0 Then Npclist(NpcIndex).Target = UserIndex

    If UserList(UserIndex).flags.AtacadoPorNpc = 0 And _
       UserList(UserIndex).flags.AtacadoPorUser = 0 Then UserList(UserIndex).flags.AtacadoPorNpc = NpcIndex
Else
    NpcAtacaUser = False
    Exit Function
End If

Npclist(NpcIndex).CanAttack = 0

If Npclist(NpcIndex).flags.Snd1 > 0 Then
    Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayWave(Npclist(NpcIndex).flags.Snd1))
End If

If NpcImpacto(NpcIndex, UserIndex) Then
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_IMPACTO))
    
    If UserList(UserIndex).flags.Meditando = False Then
        If UserList(UserIndex).flags.Navegando = 0 Or UserList(UserIndex).flags.Montado = 0 Then
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.CharIndex, FXSANGRE, 0))
        End If
    End If
    
    Call NpcDa�o(NpcIndex, UserIndex)
    Call WriteUpdateHP(UserIndex)
    '�Puede envenenar?
    If Npclist(NpcIndex).Veneno = 1 Then Call NpcEnvenenarUser(UserIndex)
Else
    Call WriteNPCSwing(UserIndex)
End If



'-----Tal vez suba los skills------
Call SubirSkill(UserIndex, Tacticas)

'Controla el nivel del usuario
Call CheckUserLevel(UserIndex)

End Function

Function NpcImpactoNpc(ByVal Atacante As Integer, ByVal Victima As Integer) As Boolean
Dim PoderAtt As Long, PoderEva As Long, dif As Long
Dim ProbExito As Long

PoderAtt = Npclist(Atacante).PoderAtaque
PoderEva = Npclist(Victima).PoderEvasion
ProbExito = Maximo(10, Minimo(90, 50 + _
            ((PoderAtt - PoderEva) * 0.4)))
NpcImpactoNpc = (RandomNumber(1, 100) <= ProbExito)


End Function

Public Sub NpcDa�oNpc(ByVal Atacante As Integer, ByVal Victima As Integer)
Dim da�o As Integer
Dim ANpc As npc, DNpc As npc
ANpc = Npclist(Atacante)

da�o = RandomNumber(ANpc.Stats.MinHIT, ANpc.Stats.MaxHIT)
Npclist(Victima).Stats.MinHP = Npclist(Victima).Stats.MinHP - da�o

If Npclist(Victima).Stats.MinHP < 1 Then
        
        If LenB(Npclist(Atacante).flags.AttackedBy) <> 0 Then
            Npclist(Atacante).Movement = Npclist(Atacante).flags.OldMovement
            Npclist(Atacante).Hostile = Npclist(Atacante).flags.OldHostil
        Else
            Npclist(Atacante).Movement = Npclist(Atacante).flags.OldMovement
        End If
        
        Call FollowAmo(Atacante)
        
        Call MuereNpc(Victima, Npclist(Atacante).MaestroUser)
End If

End Sub

Public Sub NpcAtacaNpc(ByVal Atacante As Integer, ByVal Victima As Integer, Optional ByVal cambiarMOvimiento As Boolean = True)

' El npc puede atacar ???
If Npclist(Atacante).CanAttack = 1 Then
       Npclist(Atacante).CanAttack = 0
        If cambiarMOvimiento Then
            Npclist(Victima).TargetNPC = Atacante
            Npclist(Victima).Movement = TipoAI.NpcAtacaNpc
        End If
Else
    Exit Sub
End If

If Npclist(Atacante).flags.Snd1 > 0 Then
    Call SendData(SendTarget.ToNPCArea, Atacante, PrepareMessagePlayWave(Npclist(Atacante).flags.Snd1))
End If

If NpcImpactoNpc(Atacante, Victima) Then
    
    If Npclist(Victima).flags.Snd2 > 0 Then
        Call SendData(SendTarget.ToNPCArea, Victima, PrepareMessagePlayWave(Npclist(Victima).flags.Snd2))
    Else
        Call SendData(SendTarget.ToNPCArea, Victima, PrepareMessagePlayWave(SND_IMPACTO2))
    End If

    If Npclist(Atacante).MaestroUser > 0 Then
        Call SendData(SendTarget.ToNPCArea, Atacante, PrepareMessagePlayWave(SND_IMPACTO))
    Else
        Call SendData(SendTarget.ToNPCArea, Victima, PrepareMessagePlayWave(SND_IMPACTO))
    End If
    Call NpcDa�oNpc(Atacante, Victima)
    
Else
    If Npclist(Atacante).MaestroUser > 0 Then
        Call SendData(SendTarget.ToNPCArea, Atacante, PrepareMessagePlayWave(SND_SWING))
    Else
        Call SendData(SendTarget.ToNPCArea, Victima, PrepareMessagePlayWave(SND_SWING))
    End If
End If

End Sub

Public Sub UsuarioAtacaNpc(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)

If UserList(UserIndex).flags.Privilegios And PlayerType.Consejero Then Exit Sub

If Distancia(UserList(UserIndex).Pos, Npclist(NpcIndex).Pos) > MAXDISTANCIAARCO Then
   Call WriteConsoleMsg(UserIndex, "Est�s muy lejos para disparar.", FontTypeNames.FONTTYPE_FIGHT)
   Exit Sub
End If

If Npclist(NpcIndex).MaestroUser <> 0 Then
    If UserList(Npclist(NpcIndex).MaestroUser).Faccion.Alineacion = UserList(UserIndex).Faccion.Alineacion And UserList(UserIndex).Faccion.Alineacion <> e_Alineacion.Neutro Then
        Call WriteConsoleMsg(UserIndex, "No puedes atacar a usuarios de tu faccion.", FontTypeNames.FONTTYPE_WARNING)
        Exit Sub
    End If
End If

If Npclist(NpcIndex).EsRey Then
    If UserList(UserIndex).Faccion.Alineacion = e_Alineacion.Neutro Then
        Call WriteConsoleMsg(UserIndex, "Debes pertenecer a una faccion para atacar a este npc.", FontTypeNames.FONTTYPE_FIGHT)
        Exit Sub
    ElseIf UserList(UserIndex).Faccion.Alineacion = Castillo(Npclist(NpcIndex).EsRey).LeaderFaccion Then
        Call WriteConsoleMsg(UserIndex, "No puedes atacar al rey de tu castillo.", FontTypeNames.FONTTYPE_FIGHT)
        Exit Sub
    Else
        CastleUnderAttack Npclist(NpcIndex).EsRey
    End If
End If

'Revisa que el Rey pretoriano no est� solo.
If esPretoriano(NpcIndex) = 4 Then
    If Npclist(NpcIndex).Pos.X < 50 Then
        If pretorianosVivos(1) > 0 Then
            Call WriteConsoleMsg(UserIndex, "Debes matar al resto del ejercito antes de atacar al rey!", FontTypeNames.FONTTYPE_FIGHT)
            Exit Sub
        End If
    Else
        If pretorianosVivos(2) > 0 Then
            Call WriteConsoleMsg(UserIndex, "Debes matar al resto del ejercito antes de atacar al rey!", FontTypeNames.FONTTYPE_FIGHT)
            Exit Sub
        End If
    End If
End If

Call NPCAtacado(NpcIndex, UserIndex)
Call CheckPets(NpcIndex, UserIndex)

If UserImpactoNpc(UserIndex, NpcIndex) Then
    
    If Npclist(NpcIndex).flags.Snd2 > 0 Then
        Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayWave(Npclist(NpcIndex).flags.Snd2))
    Else
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_IMPACTO2))
    End If
    
    Call UserDa�oNpc(UserIndex, NpcIndex)
   
Else
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_SWING))
    Call WriteUserSwing(UserIndex)
End If

End Sub

Public Sub UsuarioAtaca(ByVal UserIndex As Integer)

'If UserList(UserIndex).flags.PuedeAtacar = 1 Then
'Check bow's interval
If Not IntervaloPermiteUsarArcos(UserIndex, False) Then Exit Sub

'Nos fijamos que no corte los intervalos.
'If Not UserList(UserIndex).LAC.LPegar.Puedo Then Exit Sub

'Check Spell-Magic interval
If Not IntervaloPermiteMagiaGolpe(UserIndex) Then
    'Check Attack interval
    If Not IntervaloPermiteAtacar(UserIndex) Then
        Exit Sub
    End If
End If

'Quitamos stamina
If UserList(UserIndex).Stats.MinSta >= 10 Then
    Call QuitarSta(UserIndex, RandomNumber(1, 10))
Else
    Call WriteConsoleMsg(UserIndex, "Estas muy cansado para luchar.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If
 
'UserList(UserIndex).flags.PuedeAtacar = 0

Dim AttackPos As WorldPos
AttackPos = UserList(UserIndex).Pos
Call HeadtoPos(UserList(UserIndex).Char.Heading, AttackPos)
   
'Exit if not legal
If AttackPos.X < XMinMapSize Or AttackPos.X > XMaxMapSize Or AttackPos.Y <= YMinMapSize Or AttackPos.Y > YMaxMapSize Then
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_SWING))
    Exit Sub
End If
    
Dim Index As Integer
Index = MapData(AttackPos.Map, AttackPos.X, AttackPos.Y).UserIndex
    
'Look for user
If Index > 0 Then
    Call UsuarioAtacaUsuario(UserIndex, Index)
    Call WriteUpdateUserStats(UserIndex)
    Call WriteUpdateUserStats(Index)
    Exit Sub
End If
    
'Look for NPC
If MapData(AttackPos.Map, AttackPos.X, AttackPos.Y).NpcIndex > 0 Then
    
    If Npclist(MapData(AttackPos.Map, AttackPos.X, AttackPos.Y).NpcIndex).Attackable Then
            
        If Npclist(MapData(AttackPos.Map, AttackPos.X, AttackPos.Y).NpcIndex).MaestroUser > 0 And _
            MapInfo(Npclist(MapData(AttackPos.Map, AttackPos.X, AttackPos.Y).NpcIndex).Pos.Map).Pk = False Then
                Call WriteConsoleMsg(UserIndex, "No pod�s atacar mascotas en zonas seguras", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub
        End If

        Call UsuarioAtacaNpc(UserIndex, MapData(AttackPos.Map, AttackPos.X, AttackPos.Y).NpcIndex)
            
    Else
        Call WriteConsoleMsg(UserIndex, "No pod�s atacar a este NPC", FontTypeNames.FONTTYPE_FIGHT)
    End If
        
    Call WriteUpdateUserStats(UserIndex)
        
    Exit Sub
End If
    
Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_SWING))
Call WriteUpdateUserStats(UserIndex)


If UserList(UserIndex).Counters.Trabajando Then _
    UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando - 1
    
If UserList(UserIndex).Counters.Ocultando Then _
    UserList(UserIndex).Counters.Ocultando = UserList(UserIndex).Counters.Ocultando - 1

End Sub

Public Function UsuarioImpacto(ByVal atacanteindex As Integer, ByVal victimaindex As Integer) As Boolean

Dim ProbRechazo As Long
Dim Rechazo As Boolean
Dim ProbExito As Long
Dim PoderAtaque As Long
Dim UserPoderEvasion As Long
Dim UserPoderEvasionEscudo As Long
Dim Arma As Integer
Dim proyectil As Boolean
Dim SkillTacticas As Long
Dim SkillDefensa As Long

SkillTacticas = UserList(victimaindex).Stats.UserSkills(eSkill.Tacticas)
SkillDefensa = UserList(victimaindex).Stats.UserSkills(eSkill.Defensa)

Arma = UserList(atacanteindex).Invent.WeaponEqpObjIndex
If Arma > 0 Then
    proyectil = ObjData(Arma).proyectil = 1
Else
    proyectil = False
End If

'Calculamos el poder de evasion...
UserPoderEvasion = PoderEvasion(victimaindex)

If UserList(victimaindex).Invent.EscudoEqpObjIndex > 0 Then
   UserPoderEvasionEscudo = PoderEvasionEscudo(victimaindex)
   UserPoderEvasion = UserPoderEvasion + UserPoderEvasionEscudo
Else
    UserPoderEvasionEscudo = 0
End If

'Esta usando un arma ???
If UserList(atacanteindex).Invent.WeaponEqpObjIndex > 0 Then
    
    If proyectil Then
        PoderAtaque = PoderAtaqueProyectil(atacanteindex)
    Else
        PoderAtaque = PoderAtaqueArma(atacanteindex)
    End If
    ProbExito = Maximo(10, Minimo(90, 50 + _
                ((PoderAtaque - UserPoderEvasion) * 0.4)))
   
Else
    PoderAtaque = PoderAtaqueWrestling(atacanteindex)
    ProbExito = Maximo(10, Minimo(90, 50 + _
                ((PoderAtaque - UserPoderEvasion) * 0.4)))
    
End If
UsuarioImpacto = (RandomNumber(1, 100) <= ProbExito)

' el usuario esta usando un escudo ???
If UserList(victimaindex).Invent.EscudoEqpObjIndex > 0 Then
    
    'Fallo ???
    If UsuarioImpacto = False Then
      ProbRechazo = Maximo(10, Minimo(90, 100 * (SkillDefensa / (SkillDefensa + SkillTacticas))))
      Rechazo = (RandomNumber(1, 100) <= ProbRechazo)
      If Rechazo = True Then
      'Se rechazo el ataque con el escudo
              Call SendData(SendTarget.ToPCArea, atacanteindex, PrepareMessagePlayWave(SND_ESCUDO))
              
              Call WriteBlockedWithShieldOther(atacanteindex)
              Call WriteBlockedWithShieldUser(victimaindex)
              
              Call SubirSkill(victimaindex, Defensa)
      End If
    End If
End If
    
If UsuarioImpacto Then
   If Arma > 0 Then
           If Not proyectil Then
                  Call SubirSkill(atacanteindex, Armas)
           Else
                  Call SubirSkill(atacanteindex, Proyectiles)
           End If
   Else
        Call SubirSkill(atacanteindex, Wrestling)
   End If
End If

Call FlushBuffer(victimaindex)
End Function

Public Sub UsuarioAtacaUsuario(ByVal atacanteindex As Integer, ByVal victimaindex As Integer)

If Not PuedeAtacar(atacanteindex, victimaindex) Then Exit Sub

If Distancia(UserList(atacanteindex).Pos, UserList(victimaindex).Pos) > MAXDISTANCIAARCO Then
   Call WriteConsoleMsg(atacanteindex, "Est�s muy lejos para disparar.", FontTypeNames.FONTTYPE_FIGHT)
   Exit Sub
End If


Call UsuarioAtacadoPorUsuario(atacanteindex, victimaindex)

If UsuarioImpacto(atacanteindex, victimaindex) Then
    Call SendData(SendTarget.ToPCArea, atacanteindex, PrepareMessagePlayWave(SND_IMPACTO))
    
    If UserList(victimaindex).flags.Navegando = 0 Or UserList(victimaindex).flags.Montado = 0 Then
        Call SendData(SendTarget.ToPCArea, victimaindex, PrepareMessageCreateFX(UserList(victimaindex).Char.CharIndex, FXSANGRE, 0))
    End If
    
    Call UserDa�oUser(atacanteindex, victimaindex)
    'Pablo (ToxicWaste): Guantes de Hurto del Bandido en acci�n
    If UserList(atacanteindex).clase = eClass.Bandit Then Call DoHurtar(atacanteindex, victimaindex)
    'y ahora, el ladr�n puede llegar a paralizar con el golpe.
    If UserList(atacanteindex).clase = eClass.Thief Then Call DoHandInmo(atacanteindex, victimaindex)
    
Else
    Call SendData(SendTarget.ToPCArea, atacanteindex, PrepareMessagePlayWave(SND_SWING))
    Call WriteUserSwing(atacanteindex)
    Call WriteUserAttackedSwing(victimaindex, atacanteindex)
End If

If UserList(atacanteindex).clase = eClass.Thief Then Call Desarmar(atacanteindex, victimaindex)

End Sub

Public Sub UserDa�oUser(ByVal atacanteindex As Integer, ByVal victimaindex As Integer)

Dim da�o As Long, antda�o As Integer
Dim Lugar As Integer, absorbido As Long
Dim defbarco As Integer
Dim defmontu As Integer
Dim MontuDamage As Integer

Dim Obj As ObjData

da�o = CalcularDa�o(atacanteindex)
antda�o = da�o

Call UserEnvenena(atacanteindex, victimaindex)

If UserList(atacanteindex).flags.Navegando = 1 And UserList(atacanteindex).Invent.BarcoObjIndex > 0 Then
     Obj = ObjData(UserList(atacanteindex).Invent.BarcoObjIndex)
     da�o = da�o + RandomNumber(Obj.MinHIT, Obj.MaxHIT)
End If

If UserList(victimaindex).flags.Navegando = 1 And UserList(victimaindex).Invent.BarcoObjIndex > 0 Then
     Obj = ObjData(UserList(victimaindex).Invent.BarcoObjIndex)
     defbarco = RandomNumber(Obj.MinDef, Obj.MaxDef)
End If

If UserList(atacanteindex).flags.Montado = 1 And UserList(atacanteindex).Invent.MonturaObjIndex > 0 Then
     Obj = ObjData(UserList(atacanteindex).Invent.MonturaObjIndex)
     If Obj.MinHIT > 0 Then
        If RandomNumber(1, 5) = 5 Then
            MontuDamage = RandomNumber(Obj.MinHIT, Obj.MaxHIT)
            da�o = da�o + MontuDamage
        End If
    End If
End If

If UserList(victimaindex).flags.Montado = 1 And UserList(victimaindex).Invent.MonturaObjIndex > 0 Then
     Obj = ObjData(UserList(victimaindex).Invent.MonturaObjIndex)
     defmontu = RandomNumber(Obj.MinDef, Obj.MaxDef)
End If

Dim Resist As Byte
If UserList(atacanteindex).Invent.WeaponEqpObjIndex > 0 Then
    Resist = ObjData(UserList(atacanteindex).Invent.WeaponEqpObjIndex).Refuerzo
End If

Lugar = RandomNumber(1, 6)

Select Case Lugar
  
  Case PartesCuerpo.bCabeza
        'Si tiene casco absorbe el golpe
        If UserList(victimaindex).Invent.CascoEqpObjIndex > 0 Then
           Obj = ObjData(UserList(victimaindex).Invent.CascoEqpObjIndex)
           absorbido = RandomNumber(Obj.MinDef, Obj.MaxDef)
           absorbido = absorbido + defbarco - Resist
           da�o = da�o - absorbido
           If da�o < 0 Then da�o = 1
        End If
  Case Else
        'Si tiene armadura absorbe el golpe
        If UserList(victimaindex).Invent.ArmourEqpObjIndex > 0 Then
           Obj = ObjData(UserList(victimaindex).Invent.ArmourEqpObjIndex)
           Dim Obj2 As ObjData
           If UserList(victimaindex).Invent.EscudoEqpObjIndex Then
                Obj2 = ObjData(UserList(victimaindex).Invent.EscudoEqpObjIndex)
                absorbido = RandomNumber(Obj.MinDef + Obj2.MinDef, Obj.MaxDef + Obj2.MaxDef)
           Else
                absorbido = RandomNumber(Obj.MinDef, Obj.MaxDef)
           End If
           absorbido = absorbido + defbarco - Resist
           da�o = da�o - absorbido
           If da�o < 0 Then da�o = 1
        End If
        
        'Tiene montura?, tambien absorbe el golpe.
        If UserList(victimaindex).Invent.MonturaObjIndex > 0 Then
            If RandomNumber(1, 7) = 5 Then
                da�o = da�o - RandomNumber(ObjData(UserList(victimaindex).Invent.MonturaObjIndex).MinDef, ObjData(UserList(victimaindex).Invent.MonturaObjIndex).MaxDef)
            End If
        End If
        
End Select

Call WriteUserHittedUser(atacanteindex, Lugar, UserList(victimaindex).Char.CharIndex, da�o)
Call WriteUserHittedByUser(victimaindex, Lugar, UserList(atacanteindex).Char.CharIndex, da�o)
If MontuDamage > 0 Then
    Call WriteConsoleMsg(victimaindex, "La montura de tu atacante de ha pegado por " & MontuDamage & ".", FontTypeNames.FONTTYPE_FIGHT)
    Call WriteConsoleMsg(atacanteindex, "Tu montura le ha pegado a la victima por " & MontuDamage & ".", FontTypeNames.FONTTYPE_FIGHT)
End If

UserList(victimaindex).Stats.MinHP = UserList(victimaindex).Stats.MinHP - da�o

If UserList(atacanteindex).flags.Hambre = 0 And UserList(atacanteindex).flags.Sed = 0 Then
        'Si usa un arma quizas suba "Combate con armas"
        If UserList(atacanteindex).Invent.WeaponEqpObjIndex > 0 Then
            If ObjData(UserList(atacanteindex).Invent.WeaponEqpObjIndex).proyectil Then
                'es un Arco. Sube Armas a Distancia
                Call SubirSkill(atacanteindex, Proyectiles)
            Else
                'Sube combate con armas.
                Call SubirSkill(atacanteindex, Armas)
            End If
        Else
        'sino tal vez lucha libre
                Call SubirSkill(atacanteindex, Wrestling)
        End If
        
        Call SubirSkill(victimaindex, Tacticas)
        
        'Trata de apu�alar por la espalda al enemigo
        If PuedeApu�alar(atacanteindex) Then
                Call DoApu�alar(atacanteindex, 0, victimaindex, da�o)
                Call SubirSkill(atacanteindex, Apu�alar)
        End If
        'e intenta dar un golpe cr�tico [Pablo (ToxicWaste)]
        Call DoGolpeCritico(atacanteindex, 0, victimaindex, da�o)
End If

If UserList(victimaindex).Stats.MinHP <= 0 Then
    'Store it!

    Call Statistics.StoreFrag(atacanteindex, victimaindex)

    Call ContarMuerte(victimaindex, atacanteindex)
    ' Para que las mascotas no sigan intentando luchar y
    ' comiencen a seguir al amo
    Dim j As Integer
    For j = 1 To MAXMASCOTAS
        If UserList(atacanteindex).MascotasIndex(j) > 0 Then
            If Npclist(UserList(atacanteindex).MascotasIndex(j)).Target = victimaindex Then Npclist(UserList(atacanteindex).MascotasIndex(j)).Target = 0
            Call FollowAmo(UserList(atacanteindex).MascotasIndex(j))
        End If
    Next j
    
    Call ActStats(victimaindex, atacanteindex)

Else
    'Est� vivo - Actualizamos el HP
    Call WriteUpdateHP(victimaindex)
End If

'Controla el nivel del usuario
Call CheckUserLevel(atacanteindex)

Call FlushBuffer(victimaindex)
End Sub

Sub UsuarioAtacadoPorUsuario(ByVal attackerIndex As Integer, ByVal VictimIndex As Integer)
'***************************************************
'Autor: Unknown
'Last Modification: 03/09/06 Nacho
'Usuario deja de meditar
'***************************************************
    If TriggerZonaPelea(attackerIndex, VictimIndex) = TRIGGER6_PERMITE Or UserList(attackerIndex).flags.EnDuelo Or UserList(attackerIndex).Pos.Map = MAPATORNEO Or IsInCastle(attackerIndex) Then Exit Sub
    
    Dim EraCriminal As Boolean
    
    If UserList(VictimIndex).flags.Meditando Then
        UserList(VictimIndex).flags.Meditando = False
        Call WriteMeditateToggle(VictimIndex)
        Call WriteConsoleMsg(VictimIndex, "Dejas de meditar.", FontTypeNames.FONTTYPE_INFO)
        UserList(VictimIndex).Char.FX = 0
        UserList(VictimIndex).Char.loops = 0
        Call SendData(SendTarget.ToPCArea, VictimIndex, PrepareMessageCreateFX(UserList(VictimIndex).Char.CharIndex, 0, 0))
    End If
    
    Call AllMascotasAtacanUser(attackerIndex, VictimIndex)
    Call AllMascotasAtacanUser(VictimIndex, attackerIndex)
    
    Call FlushBuffer(VictimIndex)
End Sub

Sub AllMascotasAtacanUser(ByVal victim As Integer, ByVal Maestro As Integer)
'Reaccion de las mascotas
Dim iCount As Integer

For iCount = 1 To MAXMASCOTAS
    If UserList(Maestro).MascotasIndex(iCount) > 0 Then
            Npclist(UserList(Maestro).MascotasIndex(iCount)).flags.AttackedBy = UserList(victim).Name
            Npclist(UserList(Maestro).MascotasIndex(iCount)).Movement = TipoAI.NPCDEFENSA
            Npclist(UserList(Maestro).MascotasIndex(iCount)).Hostile = 1
    End If
Next iCount

End Sub

Public Function PuedeAtacar(ByVal attackerIndex As Integer, ByVal VictimIndex As Integer) As Boolean
'***************************************************
'Autor: Unknown
'Last Modification: 24/01/2007
'Returns true if the AttackerIndex is allowed to attack the VictimIndex.
'24/01/2007 Pablo (ToxicWaste) - Ordeno todo y agrego situacion de Defensa en ciudad Armada y Caos.
'***************************************************
Dim T As eTrigger6
Dim rank As Integer
'MUY importante el orden de estos "IF"...

'Estas muerto no podes atacar
If UserList(attackerIndex).flags.Muerto = 1 Then
    Call WriteConsoleMsg(attackerIndex, "No pod�s atacar porque estas muerto", FontTypeNames.FONTTYPE_INFO)
    PuedeAtacar = False
    Exit Function
End If

'No podes atacar a alguien muerto
If UserList(VictimIndex).flags.Muerto = 1 Then
    Call WriteConsoleMsg(attackerIndex, "No pod�s atacar a un espiritu", FontTypeNames.FONTTYPE_INFO)
    PuedeAtacar = False
    Exit Function
End If

'Estamos en una Arena? o un trigger zona segura?
T = TriggerZonaPelea(attackerIndex, VictimIndex)

If T = eTrigger6.TRIGGER6_PERMITE Then
    PuedeAtacar = True
    Exit Function
ElseIf T = eTrigger6.TRIGGER6_PROHIBE Then
    PuedeAtacar = False
    Exit Function
ElseIf T = eTrigger6.TRIGGER6_AUSENTE Then
    'Si no estamos en el Trigger 6 entonces es imposible atacar un gm
    If Not UserList(VictimIndex).flags.Privilegios And PlayerType.User Then
        If UserList(VictimIndex).flags.AdminInvisible = 0 Then Call WriteConsoleMsg(attackerIndex, "El ser es demasiado poderoso", FontTypeNames.FONTTYPE_WARNING)
        PuedeAtacar = False
        Exit Function
    End If
End If

'Consejeros no pueden atacar
'If UserList(attackerIndex).flags.Privilegios And PlayerType.Consejero Then
'    PuedeAtacar = False
'    Exit Sub
'End If

'Estas queriendo atacar a un GM?
rank = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero

If (UserList(VictimIndex).flags.Privilegios And rank) > (UserList(attackerIndex).flags.Privilegios And rank) Then
    If UserList(VictimIndex).flags.AdminInvisible = 0 Then Call WriteConsoleMsg(attackerIndex, "El ser es demasiado poderoso", FontTypeNames.FONTTYPE_WARNING)
    PuedeAtacar = False
    Exit Function
End If

'Atacar a uno de tu misma faccion?
If UserList(attackerIndex).Faccion.Alineacion <> e_Alineacion.Neutro And UserList(VictimIndex).Faccion.Alineacion = UserList(attackerIndex).Faccion.Alineacion Then
    Call WriteConsoleMsg(attackerIndex, "Los soldados del Ejercito Real tienen prohibido atacar ciudadanos.", FontTypeNames.FONTTYPE_WARNING)
    PuedeAtacar = False
    Exit Function
End If

'Estas en un Mapa Seguro?
If MapInfo(UserList(VictimIndex).Pos.Map).Pk = False Then
    Call WriteConsoleMsg(attackerIndex, "Esta es una zona segura, aqui no podes atacar otros usuarios.", FontTypeNames.FONTTYPE_WARNING)
    PuedeAtacar = False
    Exit Function
End If

'Estas atacando desde un trigger seguro? o tu victima esta en uno asi?
If MapData(UserList(VictimIndex).Pos.Map, UserList(VictimIndex).Pos.X, UserList(VictimIndex).Pos.Y).trigger = eTrigger.ZONASEGURA Or _
    MapData(UserList(attackerIndex).Pos.Map, UserList(attackerIndex).Pos.X, UserList(attackerIndex).Pos.Y).trigger = eTrigger.ZONASEGURA Then
    Call WriteConsoleMsg(attackerIndex, "No podes pelear aqui.", FontTypeNames.FONTTYPE_WARNING)
    PuedeAtacar = False
    Exit Function
End If

PuedeAtacar = True

End Function

Public Function PuedeAtacarNPC(ByVal attackerIndex As Integer, ByVal NpcIndex As Integer) As Boolean
'***************************************************
'Autor: Unknown Author (Original version)
'Returns True if AttackerIndex can attack the NpcIndex
'Last Modification: 24/01/2007
'24/01/2007 Pablo (ToxicWaste) - Orden y correcci�n de ataque sobre una mascota y guardias
'***************************************************

'Estas muerto?
If UserList(attackerIndex).flags.Muerto = 1 Then
    Call WriteConsoleMsg(attackerIndex, "No pod�s atacar porque estas muerto", FontTypeNames.FONTTYPE_INFO)
    PuedeAtacarNPC = False
    Exit Function
End If

'Es el NPC mascota de alguien?
If Npclist(NpcIndex).MaestroUser > 0 Then
    'De un cudadanos y sos Armada?
    If Not PuedeAtacar(attackerIndex, Npclist(NpcIndex).MaestroUser) Then
        PuedeAtacarNPC = False
        Exit Function
    End If
End If

'Sos consejero? no podes atacar nunca.
If UserList(attackerIndex).flags.Privilegios And PlayerType.Consejero Then
    PuedeAtacarNPC = False
    Exit Function
End If

'Es el Rey Preatoriano?
If esPretoriano(NpcIndex) = 4 Then
    If Npclist(NpcIndex).Pos.X < 50 Then
        If pretorianosVivos(1) > 0 Then
            Call WriteConsoleMsg(attackerIndex, "Debes matar al resto del ejercito antes de atacar al rey!", FontTypeNames.FONTTYPE_FIGHT)
            PuedeAtacarNPC = False
            Exit Function
        End If
    Else
        If pretorianosVivos(2) > 0 Then
            Call WriteConsoleMsg(attackerIndex, "Debes matar al resto del ejercito antes de atacar al rey!", FontTypeNames.FONTTYPE_FIGHT)
            PuedeAtacarNPC = False
            Exit Function
        End If
    End If
End If
Debug.Print "0 -- 0"

If Npclist(NpcIndex).EsRey Then
    If Not UserList(attackerIndex).GuildIndex > 0 Then
        Debug.Print "3"
        Call WriteConsoleMsg(attackerIndex, "Debes pertenecer a un clan para atacar a este npc.", FontTypeNames.FONTTYPE_FIGHT)
        PuedeAtacarNPC = False
        Exit Function
    End If
End If
If Npclist(NpcIndex).EsRey Then
    If UserList(attackerIndex).Faccion.Alineacion = Castillo(Npclist(NpcIndex).EsRey).LeaderFaccion Then
        Call WriteConsoleMsg(attackerIndex, "No podes atacar al rey de tu castillo.", FontTypeNames.FONTTYPE_FIGHT)
        PuedeAtacarNPC = False
        Exit Function
    ElseIf UserList(attackerIndex).Faccion.Alineacion = e_Alineacion.Neutro Then
        Call WriteConsoleMsg(attackerIndex, "Debes pertenecer a una faccion para atacar a este npc.", FontTypeNames.FONTTYPE_FIGHT)
        PuedeAtacarNPC = False
        Exit Function
    End If
End If

PuedeAtacarNPC = True
End Function

Sub CalcularDarExp(ByVal UserIndex As Integer, ByVal NpcIndex As Integer, ByVal ElDa�o As Long)
'***************************************************
'Autor: Nacho (Integer)
'Last Modification: 03/09/06 Nacho
'Reescribi gran parte del Sub
'Ahora, da toda la experiencia del npc mientras este vivo.
'***************************************************
Dim ExpaDar As Long

'[Nacho] Chekeamos que las variables sean validas para las operaciones
If ElDa�o <= 0 Then ElDa�o = 0
If Npclist(NpcIndex).Stats.MaxHP <= 0 Then Exit Sub
If ElDa�o > Npclist(NpcIndex).Stats.MinHP Then ElDa�o = Npclist(NpcIndex).Stats.MinHP

If ElDa�o = Npclist(NpcIndex).Stats.MinHP Then
    ExpaDar = Npclist(NpcIndex).GiveEXP
Else
    ExpaDar = CLng((ElDa�o) * (Npclist(NpcIndex).GiveEXP / Npclist(NpcIndex).Stats.MaxHP))
End If

'[Nacho] Le damos la exp al user
If ExpaDar > 0 Then
    If ClanPoseeMapa(UserList(UserIndex).GuildIndex, Npclist(NpcIndex).Pos.Map) Then
        UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.Exp + ExpaDar * 1.1
    Else
        UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.Exp + ExpaDar
    End If
    If UserList(UserIndex).Stats.Exp > MAXEXP Then _
        UserList(UserIndex).Stats.Exp = MAXEXP
    Call WriteConsoleMsg(UserIndex, "Has ganado " & ExpaDar & " puntos de experiencia.", FontTypeNames.FONTTYPE_FIGHT)
    
    Call CheckUserLevel(UserIndex)
End If

End Sub

Public Function TriggerZonaPelea(ByVal Origen As Integer, ByVal Destino As Integer) As eTrigger6
'TODO: Pero que rebuscado!!
'Nigo:  Te lo redise�e, pero no te borro el TODO para que lo revises.
On Error GoTo errhandler
    Dim tOrg As eTrigger
    Dim tDst As eTrigger
    
    tOrg = MapData(UserList(Origen).Pos.Map, UserList(Origen).Pos.X, UserList(Origen).Pos.Y).trigger
    tDst = MapData(UserList(Destino).Pos.Map, UserList(Destino).Pos.X, UserList(Destino).Pos.Y).trigger
    
    If tOrg = eTrigger.ZONAPELEA Or tDst = eTrigger.ZONAPELEA Then
        If tOrg = tDst Then
            TriggerZonaPelea = TRIGGER6_PERMITE
        Else
            TriggerZonaPelea = TRIGGER6_PROHIBE
        End If
    Else
        TriggerZonaPelea = TRIGGER6_AUSENTE
    End If

Exit Function
errhandler:
    TriggerZonaPelea = TRIGGER6_AUSENTE
    LogError ("Error en TriggerZonaPelea - " & Err.description)
End Function

Sub UserEnvenena(ByVal atacanteindex As Integer, ByVal victimaindex As Integer)
Dim ArmaObjInd As Integer, ObjInd As Integer
Dim num As Long

ArmaObjInd = UserList(atacanteindex).Invent.WeaponEqpObjIndex
ObjInd = 0

If ArmaObjInd > 0 Then
    If ObjData(ArmaObjInd).proyectil = 0 Then
        ObjInd = ArmaObjInd
    Else
        ObjInd = UserList(atacanteindex).Invent.MunicionEqpObjIndex
    End If
    
    If ObjInd > 0 Then
        If (ObjData(ObjInd).Envenena = 1) Then
            num = RandomNumber(1, 100)
            
            If num < 60 Then
                UserList(victimaindex).flags.Envenenado = 1
                Call WriteConsoleMsg(victimaindex, UserList(atacanteindex).Name & " te ha envenenado!!", FontTypeNames.FONTTYPE_FIGHT)
                Call WriteConsoleMsg(atacanteindex, "Has envenenado a " & UserList(victimaindex).Name & "!!", FontTypeNames.FONTTYPE_FIGHT)
            End If
        End If
    End If
End If

Call FlushBuffer(victimaindex)
End Sub
