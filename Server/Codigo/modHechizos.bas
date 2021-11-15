Attribute VB_Name = "modHechizos"
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

Option Explicit

Public Const HELEMENTAL_FUEGO As Integer = 26
Public Const HELEMENTAL_TIERRA As Integer = 28
Public Const SUPERANILLO As Integer = 700

Sub NpcLanzaSpellSobreUser(ByVal NpcIndex As Integer, ByVal UserIndex As Integer, ByVal Spell As Integer)

If Npclist(NpcIndex).CanAttack = 0 Then Exit Sub
If UserList(UserIndex).flags.invisible = 1 Or UserList(UserIndex).flags.Oculto = 1 Then Exit Sub

Npclist(NpcIndex).CanAttack = 0
Dim da�o As Integer

If Hechizos(Spell).SubeHP = 1 Then

    da�o = RandomNumber(Hechizos(Spell).MinHP, Hechizos(Spell).MaxHP)
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(Hechizos(Spell).WAV))
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.CharIndex, Hechizos(Spell).FXgrh, Hechizos(Spell).loops))

    UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MinHP + da�o
    If UserList(UserIndex).Stats.MinHP > UserList(UserIndex).Stats.MaxHP Then UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MaxHP
    
    Call WriteConsoleMsg(UserIndex, Npclist(NpcIndex).Name & " te ha quitado " & da�o & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
    Call WriteUpdateUserStats(UserIndex)

ElseIf Hechizos(Spell).SubeHP = 2 Then
    
    If UserList(UserIndex).flags.Privilegios And PlayerType.User Then
    
        da�o = RandomNumber(Hechizos(Spell).MinHP, Hechizos(Spell).MaxHP)
        
        If UserList(UserIndex).Invent.CascoEqpObjIndex > 0 Then
            da�o = da�o - RandomNumber(ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).DefensaMagicaMin, ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).DefensaMagicaMax)
        End If
        
        If UserList(UserIndex).Invent.AnilloEqpObjIndex > 0 Then
            da�o = da�o - RandomNumber(ObjData(UserList(UserIndex).Invent.AnilloEqpObjIndex).DefensaMagicaMin, ObjData(UserList(UserIndex).Invent.AnilloEqpObjIndex).DefensaMagicaMax)
        End If
        
        If da�o < 0 Then da�o = 0
        
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(Hechizos(Spell).WAV))
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.CharIndex, Hechizos(Spell).FXgrh, Hechizos(Spell).loops))
    
        UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MinHP - da�o
        
        Call WriteConsoleMsg(UserIndex, Npclist(NpcIndex).Name & " te ha quitado " & da�o & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
        Call WriteUpdateUserStats(UserIndex)
        
        'Muere
        If UserList(UserIndex).Stats.MinHP < 1 Then
            UserList(UserIndex).Stats.MinHP = 0
            Call UserDie(UserIndex)
            '[Barrin 1-12-03]
            If Npclist(NpcIndex).MaestroUser > 0 Then
                'Store it!
                Call Statistics.StoreFrag(Npclist(NpcIndex).MaestroUser, UserIndex)
                Call ContarMuerte(UserIndex, Npclist(NpcIndex).MaestroUser)
                Call ActStats(UserIndex, Npclist(NpcIndex).MaestroUser)
            End If
            '[/Barrin]
        End If
    
    End If
    
End If

If Hechizos(Spell).Paraliza = 1 Then
     If UserList(UserIndex).flags.Paralizado = 0 Then
          Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(Hechizos(Spell).WAV))
          Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.CharIndex, Hechizos(Spell).FXgrh, Hechizos(Spell).loops))
          
            If UserList(UserIndex).Invent.AnilloEqpObjIndex = SUPERANILLO Then
                Call WriteConsoleMsg(UserIndex, " Tu anillo rechaza los efectos del hechizo.", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub
            End If
          UserList(UserIndex).flags.Paralizado = 1
          UserList(UserIndex).Counters.Paralisis = IntervaloParalizado
          
          Call WriteParalizeOK(UserIndex)
     End If
     
     
End If


End Sub


Sub NpcLanzaSpellSobreNpc(ByVal NpcIndex As Integer, ByVal TargetNPC As Integer, ByVal Spell As Integer)
'solo hechizos ofensivos!

If Npclist(NpcIndex).CanAttack = 0 Then Exit Sub
Npclist(NpcIndex).CanAttack = 0

Dim da�o As Integer

If Hechizos(Spell).SubeHP = 2 Then
    
        da�o = RandomNumber(Hechizos(Spell).MinHP, Hechizos(Spell).MaxHP)
        Call SendData(SendTarget.ToNPCArea, TargetNPC, PrepareMessagePlayWave(Hechizos(Spell).WAV))
        Call SendData(SendTarget.ToNPCArea, TargetNPC, PrepareMessageCreateFX(Npclist(TargetNPC).Char.CharIndex, Hechizos(Spell).FXgrh, Hechizos(Spell).loops))
        
        Npclist(TargetNPC).Stats.MinHP = Npclist(TargetNPC).Stats.MinHP - da�o
        
        'Muere
        If Npclist(TargetNPC).Stats.MinHP < 1 Then
            Npclist(TargetNPC).Stats.MinHP = 0
            If Npclist(NpcIndex).MaestroUser > 0 Then
                Call MuereNpc(TargetNPC, Npclist(NpcIndex).MaestroUser)
            Else
                Call MuereNpc(TargetNPC, 0)
            End If
        End If
    
End If
    
End Sub



Function TieneHechizo(ByVal i As Integer, ByVal UserIndex As Integer) As Boolean

On Error GoTo errhandler
    
    Dim j As Integer
    For j = 1 To MAXUSERHECHIZOS
        If UserList(UserIndex).Stats.UserHechizos(j) = i Then
            TieneHechizo = True
            Exit Function
        End If
    Next

Exit Function
errhandler:

End Function

Sub AgregarHechizo(ByVal UserIndex As Integer, ByVal Slot As Integer)
Dim hIndex As Integer
Dim j As Integer
hIndex = ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).HechizoIndex

If Not TieneHechizo(hIndex, UserIndex) Then
    'Buscamos un slot vacio
    For j = 1 To MAXUSERHECHIZOS
        If UserList(UserIndex).Stats.UserHechizos(j) = 0 Then Exit For
    Next j
        
    If UserList(UserIndex).Stats.UserHechizos(j) <> 0 Then
        Call WriteConsoleMsg(UserIndex, "No tenes espacio para mas hechizos.", FontTypeNames.FONTTYPE_INFO)
    Else
        UserList(UserIndex).Stats.UserHechizos(j) = hIndex
        Call UpdateUserHechizos(False, UserIndex, CByte(j))
        'Quitamos del inv el item
        Call QuitarUserInvItem(UserIndex, CByte(Slot), 1)
    End If
Else
    Call WriteConsoleMsg(UserIndex, "Ya tenes ese hechizo.", FontTypeNames.FONTTYPE_INFO)
End If

End Sub
            
Sub DecirPalabrasMagicas(ByVal S As String, ByVal UserIndex As Integer)
On Error Resume Next
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead(S, UserList(UserIndex).Char.CharIndex, vbCyan))
    Exit Sub
End Sub

Function PuedeLanzar(ByVal UserIndex As Integer, ByVal HechizoIndex As Integer) As Boolean

If UserList(UserIndex).flags.Muerto = 0 Then
    Dim wp2 As WorldPos
    wp2.Map = UserList(UserIndex).flags.TargetMap
    wp2.X = UserList(UserIndex).flags.TargetX
    wp2.Y = UserList(UserIndex).flags.TargetY
    
    If Hechizos(HechizoIndex).NeedStaff > 0 Then
        If UserList(UserIndex).clase = eClass.Mage Then
            If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
                If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).StaffPower < Hechizos(HechizoIndex).NeedStaff Then
                    Call WriteConsoleMsg(UserIndex, "Tu B�culo no es lo suficientemente poderoso para que puedas lanzar el conjuro.", FontTypeNames.FONTTYPE_INFO)
                    PuedeLanzar = False
                    Exit Function
                End If
            Else
                Call WriteConsoleMsg(UserIndex, "No puedes lanzar este conjuro sin la ayuda de un b�culo.", FontTypeNames.FONTTYPE_INFO)
                PuedeLanzar = False
                Exit Function
            End If
        End If
    End If
        
    If UserList(UserIndex).Stats.MinMAN >= Hechizos(HechizoIndex).ManaRequerido Then
        If UserList(UserIndex).Stats.UserSkills(eSkill.Magia) >= Hechizos(HechizoIndex).MinSkill Then
            If UserList(UserIndex).Stats.MinSta >= Hechizos(HechizoIndex).StaRequerido Then
                PuedeLanzar = True
            Else
                Call WriteConsoleMsg(UserIndex, "Est�s muy cansado para lanzar este hechizo.", FontTypeNames.FONTTYPE_INFO)
                PuedeLanzar = False
            End If
                
        Else
            Call WriteConsoleMsg(UserIndex, "No tenes suficientes puntos de magia para lanzar este hechizo.", FontTypeNames.FONTTYPE_INFO)
            PuedeLanzar = False
        End If
    Else
            Call WriteConsoleMsg(UserIndex, "No tenes suficiente mana.", FontTypeNames.FONTTYPE_INFO)
            PuedeLanzar = False
    End If
Else
   Call WriteConsoleMsg(UserIndex, "No podes lanzar hechizos porque estas muerto.", FontTypeNames.FONTTYPE_INFO)
   PuedeLanzar = False
End If

End Function

Sub HechizoTerrenoEstado(ByVal UserIndex As Integer, ByRef b As Boolean)
Dim PosCasteadaX As Integer
Dim PosCasteadaY As Integer
Dim PosCasteadaM As Integer
Dim H As Integer
Dim TempX As Integer
Dim TempY As Integer


    PosCasteadaX = UserList(UserIndex).flags.TargetX
    PosCasteadaY = UserList(UserIndex).flags.TargetY
    PosCasteadaM = UserList(UserIndex).flags.TargetMap
    
    H = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
    
    If Hechizos(H).RemueveInvisibilidadParcial = 1 Then
        b = True
        For TempX = PosCasteadaX - 8 To PosCasteadaX + 8
            For TempY = PosCasteadaY - 8 To PosCasteadaY + 8
                If InMapBounds(PosCasteadaM, TempX, TempY) Then
                    If MapData(PosCasteadaM, TempX, TempY).UserIndex > 0 Then
                        'hay un user
                        If UserList(MapData(PosCasteadaM, TempX, TempY).UserIndex).flags.invisible = 1 And UserList(MapData(PosCasteadaM, TempX, TempY).UserIndex).flags.AdminInvisible = 0 Then
                            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(MapData(PosCasteadaM, TempX, TempY).UserIndex).Char.CharIndex, Hechizos(H).FXgrh, Hechizos(H).loops))
                        End If
                    End If
                End If
            Next TempY
        Next TempX
    
        Call InfoHechizo(UserIndex)
    End If

End Sub

Sub HechizoInvocacion(ByVal UserIndex As Integer, ByRef b As Boolean)

If UserList(UserIndex).NroMacotas >= MAXMASCOTAS Then Exit Sub

'No permitimos se invoquen criaturas en zonas seguras
If MapInfo(UserList(UserIndex).Pos.Map).Pk = False Or MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = eTrigger.ZONASEGURA Then
    Call WriteConsoleMsg(UserIndex, "En zona segura no puedes invocar criaturas.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If

Dim H As Integer, j As Integer, ind As Integer, Index As Integer
Dim TargetPos As WorldPos


TargetPos.Map = UserList(UserIndex).flags.TargetMap
TargetPos.X = UserList(UserIndex).flags.TargetX
TargetPos.Y = UserList(UserIndex).flags.TargetY

H = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
    
    
For j = 1 To Hechizos(H).cant
    
    If UserList(UserIndex).NroMacotas < MAXMASCOTAS Then
        ind = SpawnNpc(Hechizos(H).NumNpc, TargetPos, True, False)
        If ind > 0 Then
            UserList(UserIndex).NroMacotas = UserList(UserIndex).NroMacotas + 1
            
            Index = FreeMascotaIndex(UserIndex)
            
            UserList(UserIndex).MascotasIndex(Index) = ind
            UserList(UserIndex).MascotasType(Index) = Npclist(ind).Numero
            
            Npclist(ind).MaestroUser = UserIndex
            Npclist(ind).Contadores.TiempoExistencia = IntervaloInvocacion
            Npclist(ind).GiveGlD = 0
            
            Call FollowAmo(ind)
        End If
            
    Else
        Exit For
    End If
    
Next j


Call InfoHechizo(UserIndex)
b = True


End Sub

Sub HandleHechizoTerreno(ByVal UserIndex As Integer, ByVal uh As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 01/10/07
'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
'Antes de procesar cualquier hechizo chequea de que este en modo de combate el
'usuario
'***************************************************

Dim b As Boolean

Select Case Hechizos(uh).Tipo
    Case TipoHechizo.uInvocacion '
        Call HechizoInvocacion(UserIndex, b)
    Case TipoHechizo.uEstado
        Call HechizoTerrenoEstado(UserIndex, b)
    
End Select

If b Then
    Call SubirSkill(UserIndex, Magia)
    'If Hechizos(uh).Resis = 1 Then Call SubirSkill(UserList(UserIndex).Flags.TargetUser, Resis)
    UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MinMAN - Hechizos(uh).ManaRequerido
    If UserList(UserIndex).Stats.MinMAN < 0 Then UserList(UserIndex).Stats.MinMAN = 0
    UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - Hechizos(uh).StaRequerido
    If UserList(UserIndex).Stats.MinSta < 0 Then UserList(UserIndex).Stats.MinSta = 0
    Call WriteUpdateUserStats(UserIndex)
End If


End Sub

Sub HandleHechizoUsuario(ByVal UserIndex As Integer, ByVal uh As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 01/10/07
'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
'Antes de procesar cualquier hechizo chequea de que este en modo de combate el
'usuario
'***************************************************

Dim b As Boolean
Select Case Hechizos(uh).Tipo
    Case TipoHechizo.uEstado ' Afectan estados (por ejem : Envenenamiento)
       Call HechizoEstadoUsuario(UserIndex, b)
    Case TipoHechizo.uPropiedades ' Afectan HP,MANA,STAMINA,ETC
       Call HechizoPropUsuario(UserIndex, b)
End Select
If b Then
    Call SubirSkill(UserIndex, Magia)
    'If Hechizos(uh).Resis = 1 Then Call SubirSkill(UserList(UserIndex).Flags.TargetUser, Resis)
    UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MinMAN - Hechizos(uh).ManaRequerido
    If UserList(UserIndex).Stats.MinMAN < 0 Then UserList(UserIndex).Stats.MinMAN = 0
    UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - Hechizos(uh).StaRequerido
    If UserList(UserIndex).Stats.MinSta < 0 Then UserList(UserIndex).Stats.MinSta = 0
    Call WriteUpdateUserStats(UserIndex)
    Call WriteUpdateUserStats(UserList(UserIndex).flags.TargetUser)
    UserList(UserIndex).flags.TargetUser = 0
End If

End Sub

Sub HandleHechizoNPC(ByVal UserIndex As Integer, ByVal uh As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 01/10/07
'Last Modified By: Agustin Andreucci (Blizzard)
'Antes de procesar cualquier hechizo chequea de que este en modo de combate el
'usuario
'Antes de procesar hechizo se fija si puede atacar al npc. (Fuertes)
'***************************************************

Dim b As Boolean

Select Case Hechizos(uh).Tipo
    Case TipoHechizo.uEstado ' Afectan estados (por ejem : Envenenamiento)
        Call HechizoEstadoNPC(UserList(UserIndex).flags.TargetNPC, uh, b, UserIndex)
    Case TipoHechizo.uPropiedades ' Afectan HP,MANA,STAMINA,ETC
        Call HechizoPropNPC(uh, UserList(UserIndex).flags.TargetNPC, UserIndex, b)
End Select

If b Then
    Call SubirSkill(UserIndex, Magia)
    UserList(UserIndex).flags.TargetNPC = 0
    UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MinMAN - Hechizos(uh).ManaRequerido
    If UserList(UserIndex).Stats.MinMAN < 0 Then UserList(UserIndex).Stats.MinMAN = 0
    UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - Hechizos(uh).StaRequerido
    If UserList(UserIndex).Stats.MinSta < 0 Then UserList(UserIndex).Stats.MinSta = 0
    Call WriteUpdateUserStats(UserIndex)
End If

End Sub


Sub LanzarHechizo(Index As Integer, UserIndex As Integer)

Dim uh As Integer
Dim exito As Boolean

uh = UserList(UserIndex).Stats.UserHechizos(Index)

If PuedeLanzar(UserIndex, uh) Then
    Select Case Hechizos(uh).Target
        Case TargetType.uUsuarios
            If UserList(UserIndex).flags.TargetUser > 0 Then
                If Abs(UserList(UserList(UserIndex).flags.TargetUser).Pos.Y - UserList(UserIndex).Pos.Y) <= RANGO_VISION_Y Then
                    Call HandleHechizoUsuario(UserIndex, uh)
                Else
                    Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos para lanzar este hechizo.", FontTypeNames.FONTTYPE_WARNING)
                End If
            Else
                Call WriteConsoleMsg(UserIndex, "Este hechizo actua solo sobre usuarios.", FontTypeNames.FONTTYPE_INFO)
            End If
        
        Case TargetType.uNPC
            If UserList(UserIndex).flags.TargetNPC > 0 Then
                If Abs(Npclist(UserList(UserIndex).flags.TargetNPC).Pos.Y - UserList(UserIndex).Pos.Y) <= RANGO_VISION_Y Then
                    Call HandleHechizoNPC(UserIndex, uh)
                Else
                    Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos para lanzar este hechizo.", FontTypeNames.FONTTYPE_WARNING)
                End If
            Else
                Call WriteConsoleMsg(UserIndex, "Este hechizo solo afecta a los npcs.", FontTypeNames.FONTTYPE_INFO)
            End If
        
        Case TargetType.uUsuariosYnpc
            If UserList(UserIndex).flags.TargetUser > 0 Then
                If Abs(UserList(UserList(UserIndex).flags.TargetUser).Pos.Y - UserList(UserIndex).Pos.Y) <= RANGO_VISION_Y Then
                    Call HandleHechizoUsuario(UserIndex, uh)
                Else
                    Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos para lanzar este hechizo.", FontTypeNames.FONTTYPE_WARNING)
                End If
            ElseIf UserList(UserIndex).flags.TargetNPC > 0 Then
                If Abs(Npclist(UserList(UserIndex).flags.TargetNPC).Pos.Y - UserList(UserIndex).Pos.Y) <= RANGO_VISION_Y Then
                    Call HandleHechizoNPC(UserIndex, uh)
                Else
                    Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos para lanzar este hechizo.", FontTypeNames.FONTTYPE_WARNING)
                End If
            Else
                Call WriteConsoleMsg(UserIndex, "Target invalido.", FontTypeNames.FONTTYPE_INFO)
            End If
        
        Case TargetType.uTerreno
            Call HandleHechizoTerreno(UserIndex, uh)
    End Select
    
End If

If UserList(UserIndex).Counters.Trabajando Then _
    UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando - 1

If UserList(UserIndex).Counters.Ocultando Then _
    UserList(UserIndex).Counters.Ocultando = UserList(UserIndex).Counters.Ocultando - 1
    
End Sub

Sub HechizoEstadoUsuario(ByVal UserIndex As Integer, ByRef b As Boolean)
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 24/01/2007
'Handles the Spells that afect the Stats of an User
'24/01/2007 Pablo (ToxicWaste) - Invisibilidad no permitida en Mapas con InviSinEfecto
'26/01/2007 Pablo (ToxicWaste) - Cambios que permiten mejor manejo de ataques en los rings.
'26/01/2007 Pablo (ToxicWaste) - Revivir no permitido en Mapas con ResuSinEfecto
'***************************************************


Dim H As Integer, tU As Integer

H = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)

tU = UserList(UserIndex).flags.TargetUser

If Hechizos(H).Invisibilidad = 1 Then
   
    If UserList(tU).flags.Muerto = 1 Then
        Call WriteConsoleMsg(UserIndex, "�Est� muerto!", FontTypeNames.FONTTYPE_INFO)
        b = False
        Exit Sub
    End If
    
    'No usar invi mapas InviSinEfecto
    If MapInfo(UserList(tU).Pos.Map).InviSinEfecto > 0 Then
        Call WriteConsoleMsg(UserIndex, "�La invisibilidad no funciona aqu�!", FontTypeNames.FONTTYPE_INFO)
        b = False
        Exit Sub
    End If
    
    'Para poder tirar invi a un pk en el ring
    If (TriggerZonaPelea(UserIndex, tU) <> TRIGGER6_PERMITE) Then
        If Faccion(tU) = e_Alineacion.Caos And Faccion(UserIndex) = e_Alineacion.Real Then
            Call WriteConsoleMsg(UserIndex, "No puedes ayudar a miembros de la faccion del Caos.", FontTypeNames.FONTTYPE_INFO)
            b = False
            Exit Sub
        Else
            b = True
        End If
    End If
    
   
    UserList(tU).flags.invisible = 1
    Call SendData(SendTarget.ToPCArea, tU, PrepareMessageSetInvisible(UserList(tU).Char.CharIndex, True))

    Call InfoHechizo(UserIndex)
    b = True
End If

If Hechizos(H).Mimetiza = 1 Then
    If UserList(tU).flags.Muerto = 1 Then
        Exit Sub
    End If
    
    If UserList(tU).flags.Navegando = 1 Then
        Exit Sub
    End If
    
    If UserList(tU).flags.Montado = 1 Then
        Exit Sub
    End If
    
    If UserList(UserIndex).flags.Montado = 1 Then
        Exit Sub
    End If
    
    If UserList(UserIndex).flags.Navegando = 1 Then
        Exit Sub
    End If
    
    If Not UserList(tU).flags.Privilegios And PlayerType.User Then
        Exit Sub
    End If
    
    If UserList(UserIndex).flags.Mimetizado = 1 Then
        Call WriteConsoleMsg(UserIndex, "Ya te encuentras transformado. El hechizo no ha tenido efecto", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    If UserList(UserIndex).flags.AdminInvisible = 1 Then Exit Sub
    
    'copio el char original al mimetizado
    
    With UserList(UserIndex)
        .CharMimetizado.body = .Char.body
        .CharMimetizado.Head = .Char.Head
        .CharMimetizado.CascoAnim = .Char.CascoAnim
        .CharMimetizado.ShieldAnim = .Char.ShieldAnim
        .CharMimetizado.WeaponAnim = .Char.WeaponAnim
        
        .flags.Mimetizado = 1
        
        'ahora pongo local el del enemigo
        .Char.body = UserList(tU).Char.body
        .Char.Head = UserList(tU).Char.Head
        .Char.CascoAnim = UserList(tU).Char.CascoAnim
        .Char.ShieldAnim = UserList(tU).Char.ShieldAnim
        .Char.WeaponAnim = UserList(tU).Char.WeaponAnim
    
        Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
    End With
   
   Call InfoHechizo(UserIndex)
   b = True
End If


If Hechizos(H).Envenena = 1 Then
        If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
        If UserIndex <> tU Then
            Call UsuarioAtacadoPorUsuario(UserIndex, tU)
        End If
        UserList(tU).flags.Envenenado = 1
        Call InfoHechizo(UserIndex)
        b = True
End If

If Hechizos(H).CuraVeneno = 1 Then
    'Para poder tirar curar veneno a un pk en el ring
    If (TriggerZonaPelea(UserIndex, tU) <> TRIGGER6_PERMITE) Then
        If PuedeAtacar(UserIndex, tU) Then
            b = True
        End If
    End If
        
    UserList(tU).flags.Envenenado = 0
    Call InfoHechizo(UserIndex)
    b = True
End If

If Hechizos(H).Maldicion = 1 Then
        If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
        If UserIndex <> tU Then
            Call UsuarioAtacadoPorUsuario(UserIndex, tU)
        End If
        UserList(tU).flags.Maldicion = 1
        Call InfoHechizo(UserIndex)
        b = True
End If

If Hechizos(H).RemoverMaldicion = 1 Then
        UserList(tU).flags.Maldicion = 0
        Call InfoHechizo(UserIndex)
        b = True
End If

If Hechizos(H).Bendicion = 1 Then
        UserList(tU).flags.Bendicion = 1
        Call InfoHechizo(UserIndex)
        b = True
End If

If Hechizos(H).Paraliza = 1 Or Hechizos(H).Inmoviliza = 1 Then
     If UserList(tU).flags.Paralizado = 0 Then
            If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
            
            If UserIndex <> tU Then
                Call UsuarioAtacadoPorUsuario(UserIndex, tU)
            End If
            
            Call InfoHechizo(UserIndex)
            b = True
            If UserList(tU).Invent.AnilloEqpObjIndex = SUPERANILLO Then
                Call WriteConsoleMsg(tU, " Tu anillo rechaza los efectos del hechizo.", FontTypeNames.FONTTYPE_FIGHT)
                Call WriteConsoleMsg(UserIndex, " �El hechizo no tiene efecto!", FontTypeNames.FONTTYPE_FIGHT)
                Call FlushBuffer(tU)
                Exit Sub
            End If
            
            UserList(tU).flags.Paralizado = 1
            UserList(tU).Counters.Paralisis = IntervaloParalizado
            
            Call WriteParalizeOK(tU)
            Call FlushBuffer(tU)

            
    End If
End If


If Hechizos(H).RemoverParalisis = 1 Then
    If UserList(tU).flags.Paralizado = 1 Then
        'Para poder tirar remo a un pk en el ring
        If (TriggerZonaPelea(UserIndex, tU) <> TRIGGER6_PERMITE) Then
            If Faccion(tU) = e_Alineacion.Caos And Faccion(UserIndex) = e_Alineacion.Real Then
                Call WriteConsoleMsg(UserIndex, "No puedes ayudar a miembros de la faccion del Caos.", FontTypeNames.FONTTYPE_INFO)
                b = False
                Exit Sub
            End If
        End If
        
        UserList(tU).flags.Paralizado = 0
        'no need to crypt this
        Call WriteParalizeOK(tU)
        Call InfoHechizo(UserIndex)
        b = True
    End If
End If

If Hechizos(H).RemoverEstupidez = 1 Then
    If UserList(tU).flags.Estupidez = 1 Then
        'Para poder tirar remo estu a un pk en el ring
        If (TriggerZonaPelea(UserIndex, tU) <> TRIGGER6_PERMITE) Then
            If Not PuedeAtacar(UserIndex, tU) Then
                b = True
            End If
        End If
    
        UserList(tU).flags.Estupidez = 0
        'no need to crypt this
        Call WriteDumbNoMore(tU)
        Call FlushBuffer(tU)
        Call InfoHechizo(UserIndex)
        b = True
    End If
End If


If Hechizos(H).Revivir = 1 Then
    If UserList(tU).flags.Muerto = 1 Then
    
        'No usar resu en mapas con ResuSinEfecto
        If MapInfo(UserList(tU).Pos.Map).ResuSinEfecto > 0 Then
            Call WriteConsoleMsg(UserIndex, "�Revivir no est� permitido aqui! Retirate de la Zona si deseas utilizar el Hechizo.", FontTypeNames.FONTTYPE_INFO)
            b = False
            Exit Sub
        End If
        
        'No podemos resucitar si nuestra barra de energ�a no est� llena. (GD: 29/04/07)
        If UserList(UserIndex).Stats.MaxSta <> UserList(UserIndex).Stats.MinSta Then
            Call WriteConsoleMsg(UserIndex, "No puedes resucitar si no tienes tu barra de energ�a llena.", FontTypeNames.FONTTYPE_INFO)
            b = False
            Exit Sub
        End If
        
        'revisamos si necesita vara
        If UserList(UserIndex).clase = eClass.Mage Then
            If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
                If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).StaffPower < Hechizos(H).NeedStaff Then
                    Call WriteConsoleMsg(UserIndex, "Necesitas un mejor b�culo para este hechizo", FontTypeNames.FONTTYPE_INFO)
                    b = False
                    Exit Sub
                End If
            End If
        ElseIf UserList(UserIndex).clase = eClass.Bard Then
            If UserList(UserIndex).Invent.AnilloEqpObjIndex <> LAUDMAGICO Then
                Call WriteConsoleMsg(UserIndex, "Necesitas un instrumento m�gico para devolver la vida", FontTypeNames.FONTTYPE_INFO)
                b = False
                Exit Sub
            End If
        End If
        
        'Para poder tirar revivir a un pk en el ring
        If (TriggerZonaPelea(UserIndex, tU) <> TRIGGER6_PERMITE) Then
            If Faccion(tU) = e_Alineacion.Caos And Faccion(UserIndex) = e_Alineacion.Real Then
                Call WriteConsoleMsg(UserIndex, "No puedes ayudar a miembros de la faccion del Caos.", FontTypeNames.FONTTYPE_INFO)
                b = False
                Exit Sub
            End If
        End If
        
        
        'Pablo Toxic Waste (GD: 29/04/07)
        UserList(tU).Stats.MinAGU = 0
        UserList(tU).flags.Sed = 1
        UserList(tU).Stats.MinHam = 0
        UserList(tU).flags.Hambre = 1
        Call WriteUpdateHungerAndThirst(tU)
        Call InfoHechizo(UserIndex)
        UserList(tU).Stats.MinMAN = 0
        UserList(tU).Stats.MinSta = 0
        Dim aux As Double
        aux = UserList(tU).Stats.ELV / 100
        aux = UserList(UserIndex).Stats.MaxHP * aux
        'Solo saco vida si es User. no quiero que exploten GMs por ahi.
        If UserList(UserIndex).flags.Privilegios And PlayerType.User Then
            UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MinHP - aux
        End If
        If (UserList(UserIndex).Stats.MinHP <= 0) Then
            Call UserDie(UserIndex)
            Call WriteConsoleMsg(UserIndex, "El esfuerzo de Resucitar fue demasiado grande", FontTypeNames.FONTTYPE_INFO)
            b = False
        Else
            Call WriteConsoleMsg(UserIndex, "El esfuerzo de resucitar te ha debilitado", FontTypeNames.FONTTYPE_INFO)
            b = True
        End If
        
        Call RevivirUsuario(tU)
    Else
        b = False
    End If

End If

If Hechizos(H).Ceguera = 1 Then
        If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
        If UserIndex <> tU Then
            Call UsuarioAtacadoPorUsuario(UserIndex, tU)
        End If
        UserList(tU).flags.Ceguera = 1
        UserList(tU).Counters.Ceguera = IntervaloParalizado / 3

        Call WriteBlind(tU)
        Call FlushBuffer(tU)
        Call InfoHechizo(UserIndex)
        b = True
End If

If Hechizos(H).Estupidez = 1 Then
        If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
        If UserIndex <> tU Then
            Call UsuarioAtacadoPorUsuario(UserIndex, tU)
        End If
        If UserList(tU).flags.Estupidez = 0 Then
            UserList(tU).flags.Estupidez = 1
            UserList(tU).Counters.Ceguera = IntervaloParalizado
        End If
        Call WriteDumb(tU)
        Call FlushBuffer(tU)

        Call InfoHechizo(UserIndex)
        b = True
End If

End Sub

Sub RevisoAtaqueNPC(ByVal NpcIndex As Integer, ByVal UserIndex As Integer, ByRef b As Boolean, ByRef ExitSub As Boolean)
'***************************************************
'Autor: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'Finds out if the UserIndex can attack the NpcIndex
'***************************************************
    
    'Es guardia caos y lo quiere atacar un caos?
    If Npclist(NpcIndex).NPCtype = eNPCType.Guardiascaos & UserList(UserIndex).Faccion.Alineacion = e_Alineacion.Caos Then
        Call WriteConsoleMsg(UserIndex, "No puedes atacar Guardias del Caos siendo Legionario", FontTypeNames.FONTTYPE_WARNING)
        b = False
        ExitSub = True
        Exit Sub
    End If
    'Es guardia Real?
    If Npclist(NpcIndex).NPCtype = eNPCType.GuardiaReal Then
        If Faccion(UserIndex) = e_Alineacion.Real Then
            Call WriteConsoleMsg(UserIndex, "No atacar Guardias Reales siendo un Soldado Real.", FontTypeNames.FONTTYPE_INFO)
            b = False
            ExitSub = True
            Exit Sub
        End If
    End If
    If Npclist(NpcIndex).MaestroUser > 0 Then 'Es mascota?
        'Puede atacar mascota?
        If Not PuedeAtacar(UserIndex, Npclist(NpcIndex).MaestroUser) Then
            Call WriteConsoleMsg(UserIndex, "No puedes atacar mascotas de miembros de tu faccion.", FontTypeNames.FONTTYPE_WARNING)
            b = False
            ExitSub = True
            Exit Sub
        End If
    End If

    Call NPCAtacado(NpcIndex, UserIndex)

End Sub

Sub HechizoEstadoNPC(ByVal NpcIndex As Integer, ByVal hIndex As Integer, ByRef b As Boolean, ByVal UserIndex As Integer)
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 26/01/2007
'Handles the Spells that afect the Stats of an NPC
'26/01/2007 Pablo (ToxicWaste) - Modificaciones por funcionamiento en los Rings y ataque a guardias
'***************************************************
Dim ExitSub As Boolean

If Hechizos(hIndex).Invisibilidad = 1 Then
   Call InfoHechizo(UserIndex)
   Npclist(NpcIndex).flags.invisible = 1
   b = True
End If

If Hechizos(hIndex).Envenena = 1 Then
   If Npclist(NpcIndex).Attackable = 0 Then
        Call WriteConsoleMsg(UserIndex, "No podes atacar a ese npc.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
   End If
   
   ExitSub = False
   Call RevisoAtaqueNPC(NpcIndex, UserIndex, b, ExitSub)
   If ExitSub = True Then Exit Sub
        
   Call InfoHechizo(UserIndex)
   Npclist(NpcIndex).flags.Envenenado = 1
   b = True
End If

If Hechizos(hIndex).CuraVeneno = 1 Then
   Call InfoHechizo(UserIndex)
   Npclist(NpcIndex).flags.Envenenado = 0
   b = True
End If

If Hechizos(hIndex).Maldicion = 1 Then
   If Npclist(NpcIndex).Attackable = 0 Then
        Call WriteConsoleMsg(UserIndex, "No podes atacar a ese npc.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
   End If
    
    Call InfoHechizo(UserIndex)
    Npclist(NpcIndex).flags.Maldicion = 1
    b = True
End If

If Hechizos(hIndex).RemoverMaldicion = 1 Then
   Call InfoHechizo(UserIndex)
   Npclist(NpcIndex).flags.Maldicion = 0
   b = True
End If

If Hechizos(hIndex).Bendicion = 1 Then
   Call InfoHechizo(UserIndex)
   Npclist(NpcIndex).flags.Bendicion = 1
   b = True
End If

If Hechizos(hIndex).Paraliza = 1 Then
    If Npclist(NpcIndex).flags.AfectaParalisis = 0 Then
        
        ExitSub = False
        Call RevisoAtaqueNPC(NpcIndex, UserIndex, b, ExitSub)
        If ExitSub = True Then Exit Sub
        
        Call InfoHechizo(UserIndex)
        Npclist(NpcIndex).flags.Paralizado = 1
        Npclist(NpcIndex).flags.Inmovilizado = 0
        Npclist(NpcIndex).Contadores.Paralisis = IntervaloParalizado
        b = True
    Else
        Call WriteConsoleMsg(UserIndex, "El npc es inmune a este hechizo.", FontTypeNames.FONTTYPE_FIGHT)
    End If
End If

'[Barrin 16-2-04]
If Hechizos(hIndex).RemoverParalisis = 1 Then
   If Npclist(NpcIndex).flags.Paralizado = 1 And Npclist(NpcIndex).MaestroUser = UserIndex Then
            Call InfoHechizo(UserIndex)
            Npclist(NpcIndex).flags.Paralizado = 0
            Npclist(NpcIndex).Contadores.Paralisis = 0
            b = True
   Else
      Call WriteConsoleMsg(UserIndex, "Este hechizo solo afecta NPCs que tengan amo.", FontTypeNames.FONTTYPE_WARNING)
   End If
End If
'[/Barrin]
 
If Hechizos(hIndex).Inmoviliza = 1 Then
    If Npclist(NpcIndex).flags.AfectaParalisis = 0 Then
        
        ExitSub = False
        Call RevisoAtaqueNPC(NpcIndex, UserIndex, b, ExitSub)
        If ExitSub = True Then Exit Sub
        
        Npclist(NpcIndex).flags.Inmovilizado = 1
        Npclist(NpcIndex).flags.Paralizado = 0
        Npclist(NpcIndex).Contadores.Paralisis = IntervaloParalizado
        Call InfoHechizo(UserIndex)
        b = True
    Else
        Call WriteConsoleMsg(UserIndex, "El npc es inmune a este hechizo.", FontTypeNames.FONTTYPE_FIGHT)
    End If
End If

End Sub

Sub HechizoPropNPC(ByVal hIndex As Integer, ByVal NpcIndex As Integer, ByVal UserIndex As Integer, ByRef b As Boolean)
Dim da�o As Long

'Salud
If Hechizos(hIndex).SubeHP = 1 Then
    da�o = RandomNumber(Hechizos(hIndex).MinHP, Hechizos(hIndex).MaxHP)
    da�o = da�o + Porcentaje(da�o, 3 * UserList(UserIndex).Stats.ELV)
    
    Call InfoHechizo(UserIndex)
    Npclist(NpcIndex).Stats.MinHP = Npclist(NpcIndex).Stats.MinHP + da�o
    If Npclist(NpcIndex).Stats.MinHP > Npclist(NpcIndex).Stats.MaxHP Then _
        Npclist(NpcIndex).Stats.MinHP = Npclist(NpcIndex).Stats.MaxHP
    Call WriteConsoleMsg(UserIndex, "Has curado " & da�o & " puntos de salud a la criatura.", FontTypeNames.FONTTYPE_FIGHT)
    b = True
ElseIf Hechizos(hIndex).SubeHP = 2 Then
    
    If Npclist(NpcIndex).Attackable = 0 Then
        Call WriteConsoleMsg(UserIndex, "No podes atacar a ese npc.", FontTypeNames.FONTTYPE_INFO)
        b = False
        Exit Sub
    End If
    
    If Not PuedeAtacarNPC(UserIndex, NpcIndex) Then
        b = False
        Exit Sub
    End If
    
    da�o = RandomNumber(Hechizos(hIndex).MinHP, Hechizos(hIndex).MaxHP)
    da�o = da�o + Porcentaje(da�o, 3 * UserList(UserIndex).Stats.ELV)

    If Hechizos(hIndex).StaffAffected Then
        If UserList(UserIndex).clase = eClass.Mage Then
            If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
                da�o = (da�o * (ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).StaffDamageBonus + 70)) / 100
                'Aumenta da�o segun el staff-
                'Da�o = (Da�o* (80 + BonifB�culo)) / 100
            Else
                da�o = da�o * 0.7 'Baja da�o a 70% del original
            End If
        End If
    End If
    If UserList(UserIndex).Invent.AnilloEqpObjIndex = LAUDMAGICO Then
        da�o = da�o * 1.04  'laud magico de los bardos
    End If

    If UserList(UserIndex).Invent.WeaponEqpObjIndex = VaraMataDragonesIndex Then
        If Npclist(NpcIndex).NPCtype = eNPCType.DRAGON Then
            da�o = Npclist(NpcIndex).Stats.MinHP + Npclist(NpcIndex).Stats.def
            Call QuitarObjetos(VaraMataDragonesIndex, 1, UserIndex)
        Else
            da�o = 1
        End If
    End If
    
    Call InfoHechizo(UserIndex)
    b = True
    Call NPCAtacado(NpcIndex, UserIndex)
    If Npclist(NpcIndex).flags.Snd2 > 0 Then
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(Npclist(NpcIndex).flags.Snd2))
    End If
    
    Npclist(NpcIndex).Stats.MinHP = Npclist(NpcIndex).Stats.MinHP - da�o
    Call WriteConsoleMsg(UserIndex, "Le has causado " & da�o & " puntos de da�o a la criatura!", FontTypeNames.FONTTYPE_FIGHT)
    Call CalcularDarExp(UserIndex, NpcIndex, da�o)

    If Npclist(NpcIndex).Stats.MinHP < 1 Then
        Npclist(NpcIndex).Stats.MinHP = 0
        Call MuereNpc(NpcIndex, UserIndex)
    Else
        If Npclist(UserList(UserIndex).flags.TargetNPC).EsRey Then
            CastleUnderAttack Npclist(UserList(UserIndex).flags.TargetNPC).EsRey
        End If
        Call CheckPets(UserList(UserIndex).flags.TargetNPC, UserIndex)
    End If
End If

End Sub

Sub InfoHechizo(ByVal UserIndex As Integer)


    Dim H As Integer
    
    H = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
    
    Call DecirPalabrasMagicas(Hechizos(H).PalabrasMagicas, UserIndex)
    
    If UserList(UserIndex).flags.TargetUser > 0 Then
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserList(UserIndex).flags.TargetUser).Char.CharIndex, Hechizos(H).FXgrh, Hechizos(H).loops))
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(Hechizos(H).WAV)) 'Esta linea faltaba. Pablo (ToxicWaste)
    ElseIf UserList(UserIndex).flags.TargetNPC > 0 Then
        Call SendData(SendTarget.ToNPCArea, UserList(UserIndex).flags.TargetNPC, PrepareMessageCreateFX(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex, Hechizos(H).FXgrh, Hechizos(H).loops))
        Call SendData(SendTarget.ToNPCArea, UserList(UserIndex).flags.TargetNPC, PrepareMessagePlayWave(Hechizos(H).WAV))
    End If
    
    If UserList(UserIndex).flags.TargetUser > 0 Then
        If UserIndex <> UserList(UserIndex).flags.TargetUser Then
            If UserList(UserIndex).showName Then
                Call WriteConsoleMsg(UserIndex, Hechizos(H).HechizeroMsg & " " & UserList(UserList(UserIndex).flags.TargetUser).Name, FontTypeNames.FONTTYPE_FIGHT)
            Else
                Call WriteConsoleMsg(UserIndex, Hechizos(H).HechizeroMsg & " alguien.", FontTypeNames.FONTTYPE_FIGHT)
            End If
            Call WriteConsoleMsg(UserList(UserIndex).flags.TargetUser, UserList(UserIndex).Name & " " & Hechizos(H).TargetMsg, FontTypeNames.FONTTYPE_FIGHT)
        Else
            Call WriteConsoleMsg(UserIndex, Hechizos(H).PropioMsg, FontTypeNames.FONTTYPE_FIGHT)
        End If
    ElseIf UserList(UserIndex).flags.TargetNPC > 0 Then
        Call WriteConsoleMsg(UserIndex, Hechizos(H).HechizeroMsg & " " & "la criatura.", FontTypeNames.FONTTYPE_FIGHT)
    End If

End Sub

Sub HechizoPropUsuario(ByVal UserIndex As Integer, ByRef b As Boolean)

Dim H As Integer
Dim da�o As Integer
Dim tempChr As Integer
    
    

H = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)


tempChr = UserList(UserIndex).flags.TargetUser
      
      
'Hambre
If Hechizos(H).SubeHam = 1 Then
    
    Call InfoHechizo(UserIndex)
    
    da�o = RandomNumber(Hechizos(H).MinHam, Hechizos(H).MaxHam)
    
    UserList(tempChr).Stats.MinHam = UserList(tempChr).Stats.MinHam + da�o
    If UserList(tempChr).Stats.MinHam > UserList(tempChr).Stats.MaxHam Then _
        UserList(tempChr).Stats.MinHam = UserList(tempChr).Stats.MaxHam
    
    If UserIndex <> tempChr Then
        Call WriteConsoleMsg(UserIndex, "Le has restaurado " & da�o & " puntos de hambre a " & UserList(tempChr).Name, FontTypeNames.FONTTYPE_FIGHT)
        Call WriteConsoleMsg(tempChr, UserList(UserIndex).Name & " te ha restaurado " & da�o & " puntos de hambre.", FontTypeNames.FONTTYPE_FIGHT)
    Else
        Call WriteConsoleMsg(UserIndex, "Te has restaurado " & da�o & " puntos de hambre.", FontTypeNames.FONTTYPE_FIGHT)
    End If
    
    Call WriteUpdateHungerAndThirst(tempChr)
    b = True
    
ElseIf Hechizos(H).SubeHam = 2 Then
    If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
    If UserIndex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)
    Else
        Exit Sub
    End If
    
    Call InfoHechizo(UserIndex)
    
    da�o = RandomNumber(Hechizos(H).MinHam, Hechizos(H).MaxHam)
    
    UserList(tempChr).Stats.MinHam = UserList(tempChr).Stats.MinHam - da�o
    
    If UserList(tempChr).Stats.MinHam < 0 Then UserList(tempChr).Stats.MinHam = 0
    
    If UserIndex <> tempChr Then
        Call WriteConsoleMsg(UserIndex, "Le has quitado " & da�o & " puntos de hambre a " & UserList(tempChr).Name, FontTypeNames.FONTTYPE_FIGHT)
        Call WriteConsoleMsg(tempChr, UserList(UserIndex).Name & " te ha quitado " & da�o & " puntos de hambre.", FontTypeNames.FONTTYPE_FIGHT)
    Else
        Call WriteConsoleMsg(UserIndex, "Te has quitado " & da�o & " puntos de hambre.", FontTypeNames.FONTTYPE_FIGHT)
    End If
    
    Call WriteUpdateHungerAndThirst(tempChr)
    
    b = True
    
    If UserList(tempChr).Stats.MinHam < 1 Then
        UserList(tempChr).Stats.MinHam = 0
        UserList(tempChr).flags.Hambre = 1
    End If
    
End If

'Sed
If Hechizos(H).SubeSed = 1 Then
    
    Call InfoHechizo(UserIndex)
    
    da�o = RandomNumber(Hechizos(H).MinSed, Hechizos(H).MaxSed)
    
    UserList(tempChr).Stats.MinAGU = UserList(tempChr).Stats.MinAGU + da�o
    If UserList(tempChr).Stats.MinAGU > UserList(tempChr).Stats.MaxAGU Then _
        UserList(tempChr).Stats.MinAGU = UserList(tempChr).Stats.MaxAGU
         
    If UserIndex <> tempChr Then
      Call WriteConsoleMsg(UserIndex, "Le has restaurado " & da�o & " puntos de sed a " & UserList(tempChr).Name, FontTypeNames.FONTTYPE_FIGHT)
      Call WriteConsoleMsg(tempChr, UserList(UserIndex).Name & " te ha restaurado " & da�o & " puntos de sed.", FontTypeNames.FONTTYPE_FIGHT)
    Else
      Call WriteConsoleMsg(UserIndex, "Te has restaurado " & da�o & " puntos de sed.", FontTypeNames.FONTTYPE_FIGHT)
    End If
    
    b = True
    
ElseIf Hechizos(H).SubeSed = 2 Then
    
    If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
    If UserIndex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)
    End If
    
    Call InfoHechizo(UserIndex)
    
    da�o = RandomNumber(Hechizos(H).MinSed, Hechizos(H).MaxSed)
    
    UserList(tempChr).Stats.MinAGU = UserList(tempChr).Stats.MinAGU - da�o
    
    If UserIndex <> tempChr Then
        Call WriteConsoleMsg(UserIndex, "Le has quitado " & da�o & " puntos de sed a " & UserList(tempChr).Name, FontTypeNames.FONTTYPE_FIGHT)
        Call WriteConsoleMsg(tempChr, UserList(UserIndex).Name & " te ha quitado " & da�o & " puntos de sed.", FontTypeNames.FONTTYPE_FIGHT)
    Else
        Call WriteConsoleMsg(UserIndex, "Te has quitado " & da�o & " puntos de sed.", FontTypeNames.FONTTYPE_FIGHT)
    End If
    
    If UserList(tempChr).Stats.MinAGU < 1 Then
            UserList(tempChr).Stats.MinAGU = 0
            UserList(tempChr).flags.Sed = 1
    End If
    
    b = True
End If

' <-------- Agilidad ---------->
If Hechizos(H).SubeAgilidad = 1 Then
    
    Call InfoHechizo(UserIndex)
    da�o = RandomNumber(Hechizos(H).MinAgilidad, Hechizos(H).MaxAgilidad)
    
    UserList(tempChr).flags.DuracionEfecto = 1200
    UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) + da�o
    If UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) > MinimoInt(MAXATRIBUTOS, UserList(tempChr).Stats.UserAtributosBackUP(Agilidad) * 2) Then _
        UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = MinimoInt(MAXATRIBUTOS, UserList(tempChr).Stats.UserAtributosBackUP(Agilidad) * 2)
    UserList(tempChr).flags.TomoPocion = True
    b = True
    
ElseIf Hechizos(H).SubeAgilidad = 2 Then
    
    If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
    If UserIndex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)
    End If
    
    Call InfoHechizo(UserIndex)
    
    UserList(tempChr).flags.TomoPocion = True
    da�o = RandomNumber(Hechizos(H).MinAgilidad, Hechizos(H).MaxAgilidad)
    UserList(tempChr).flags.DuracionEfecto = 700
    UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) - da�o
    If UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) < MINATRIBUTOS Then UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = MINATRIBUTOS
    b = True
    
    Call WriteUpdateStrengthAgility(tempChr)
    
End If

' <-------- Fuerza ---------->
If Hechizos(H).SubeFuerza = 1 Then
    
    Call InfoHechizo(UserIndex)
    da�o = RandomNumber(Hechizos(H).MinFuerza, Hechizos(H).MaxFuerza)
    
    UserList(tempChr).flags.DuracionEfecto = 1200

    UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) = UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) + da�o
    If UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) > MinimoInt(MAXATRIBUTOS, UserList(tempChr).Stats.UserAtributosBackUP(Fuerza) * 2) Then _
        UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) = MinimoInt(MAXATRIBUTOS, UserList(tempChr).Stats.UserAtributosBackUP(Fuerza) * 2)
    
    UserList(tempChr).flags.TomoPocion = True
    b = True
    
    Call WriteUpdateStrengthAgility(tempChr)
    
ElseIf Hechizos(H).SubeFuerza = 2 Then

    If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
    If UserIndex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)
    End If
    
    Call InfoHechizo(UserIndex)
    
    UserList(tempChr).flags.TomoPocion = True
    
    da�o = RandomNumber(Hechizos(H).MinFuerza, Hechizos(H).MaxFuerza)
    UserList(tempChr).flags.DuracionEfecto = 700
    UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) = UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) - da�o
    If UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) < MINATRIBUTOS Then UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) = MINATRIBUTOS
    b = True
    
End If

'Salud
If Hechizos(H).SubeHP = 1 Then
    da�o = RandomNumber(Hechizos(H).MinHP, Hechizos(H).MaxHP)
    da�o = da�o + Porcentaje(da�o, 3 * UserList(UserIndex).Stats.ELV)
    
    Call InfoHechizo(UserIndex)

    UserList(tempChr).Stats.MinHP = UserList(tempChr).Stats.MinHP + da�o
    If UserList(tempChr).Stats.MinHP > UserList(tempChr).Stats.MaxHP Then _
        UserList(tempChr).Stats.MinHP = UserList(tempChr).Stats.MaxHP
    
    If UserIndex <> tempChr Then
        Call WriteConsoleMsg(UserIndex, "Le has restaurado " & da�o & " puntos de vida a " & UserList(tempChr).Name, FontTypeNames.FONTTYPE_FIGHT)
        Call WriteConsoleMsg(tempChr, UserList(UserIndex).Name & " te ha restaurado " & da�o & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
    Else
        Call WriteConsoleMsg(UserIndex, "Te has restaurado " & da�o & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
    End If
    
    b = True
ElseIf Hechizos(H).SubeHP = 2 Then
    
    If UserIndex = tempChr Then
        Call WriteConsoleMsg(UserIndex, "No podes atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
        Exit Sub
    End If
    
    da�o = RandomNumber(Hechizos(H).MinHP, Hechizos(H).MaxHP)
    
    da�o = da�o + Porcentaje(da�o, 3 * UserList(UserIndex).Stats.ELV)
    
    If Hechizos(H).StaffAffected Then
        If UserList(UserIndex).clase = eClass.Mage Then
            If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
                da�o = (da�o * (ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).StaffDamageBonus + 70)) / 100
            Else
                da�o = da�o * 0.7 'Baja da�o a 70% del original
            End If
        End If
    End If
    
    If UserList(UserIndex).Invent.AnilloEqpObjIndex = LAUDMAGICO Then
        da�o = da�o * 1.04  'laud magico de los bardos
    End If
    
    'cascos antimagia
    If (UserList(tempChr).Invent.CascoEqpObjIndex > 0) Then
        da�o = da�o - RandomNumber(ObjData(UserList(tempChr).Invent.CascoEqpObjIndex).DefensaMagicaMin, ObjData(UserList(tempChr).Invent.CascoEqpObjIndex).DefensaMagicaMax)
    End If
    
    'anillos
    If (UserList(tempChr).Invent.AnilloEqpObjIndex > 0) Then
        da�o = da�o - RandomNumber(ObjData(UserList(tempChr).Invent.AnilloEqpObjIndex).DefensaMagicaMin, ObjData(UserList(tempChr).Invent.AnilloEqpObjIndex).DefensaMagicaMax)
    End If
    
    If da�o < 0 Then da�o = 0
    
    If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
    If UserIndex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)
    End If
    
    Call InfoHechizo(UserIndex)
    
    UserList(tempChr).Stats.MinHP = UserList(tempChr).Stats.MinHP - da�o
    
    Call WriteConsoleMsg(UserIndex, "Le has quitado " & da�o & " puntos de vida a " & UserList(tempChr).Name, FontTypeNames.FONTTYPE_FIGHT)
    Call WriteConsoleMsg(tempChr, UserList(UserIndex).Name & " te ha quitado " & da�o & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
    
    'Si Lanza hechizo a otro usuario y le quita vida, entonces ya no es invisible!
    UserList(UserIndex).flags.invisible = 0
    UserList(UserIndex).Counters.Invisibilidad = 0
    If UserList(UserIndex).flags.Oculto = 0 Then
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(UserList(UserIndex).Char.CharIndex, False))
        Call WriteConsoleMsg(UserIndex, "�Has vuelto a ser visible!", FontTypeNames.FONTTYPE_INFO)
    End If
    
    'Muere
    If UserList(tempChr).Stats.MinHP < 1 Then
        'Store it!
        Call Statistics.StoreFrag(UserIndex, tempChr)
        
        Call ContarMuerte(tempChr, UserIndex)
        UserList(tempChr).Stats.MinHP = 0
        Call ActStats(tempChr, UserIndex)
        'GRAVE ERROR?
        'Call UserDie(tempChr)
    End If
    
    b = True
End If

'Mana
If Hechizos(H).SubeMana = 1 Then
    
    Call InfoHechizo(UserIndex)
    UserList(tempChr).Stats.MinMAN = UserList(tempChr).Stats.MinMAN + da�o
    If UserList(tempChr).Stats.MinMAN > UserList(tempChr).Stats.MaxMAN Then _
        UserList(tempChr).Stats.MinMAN = UserList(tempChr).Stats.MaxMAN
    
    If UserIndex <> tempChr Then
        Call WriteConsoleMsg(UserIndex, "Le has restaurado " & da�o & " puntos de mana a " & UserList(tempChr).Name, FontTypeNames.FONTTYPE_FIGHT)
        Call WriteConsoleMsg(tempChr, UserList(UserIndex).Name & " te ha restaurado " & da�o & " puntos de mana.", FontTypeNames.FONTTYPE_FIGHT)
    Else
        Call WriteConsoleMsg(UserIndex, "Te has restaurado " & da�o & " puntos de mana.", FontTypeNames.FONTTYPE_FIGHT)
    End If
    
    b = True
    
ElseIf Hechizos(H).SubeMana = 2 Then
    If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
    If UserIndex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)
    End If
    
    Call InfoHechizo(UserIndex)
    
    If UserIndex <> tempChr Then
        Call WriteConsoleMsg(UserIndex, "Le has quitado " & da�o & " puntos de mana a " & UserList(tempChr).Name, FontTypeNames.FONTTYPE_FIGHT)
        Call WriteConsoleMsg(tempChr, UserList(UserIndex).Name & " te ha quitado " & da�o & " puntos de mana.", FontTypeNames.FONTTYPE_FIGHT)
    Else
        Call WriteConsoleMsg(UserIndex, "Te has quitado " & da�o & " puntos de mana.", FontTypeNames.FONTTYPE_FIGHT)
    End If
    
    UserList(tempChr).Stats.MinMAN = UserList(tempChr).Stats.MinMAN - da�o
    If UserList(tempChr).Stats.MinMAN < 1 Then UserList(tempChr).Stats.MinMAN = 0
    b = True
    
End If

'Stamina
If Hechizos(H).SubeSta = 1 Then
    Call InfoHechizo(UserIndex)
    UserList(tempChr).Stats.MinSta = UserList(tempChr).Stats.MinSta + da�o
    If UserList(tempChr).Stats.MinSta > UserList(tempChr).Stats.MaxSta Then _
        UserList(tempChr).Stats.MinSta = UserList(tempChr).Stats.MaxSta
    If UserIndex <> tempChr Then
        Call WriteConsoleMsg(UserIndex, "Le has restaurado " & da�o & " puntos de vitalidad a " & UserList(tempChr).Name, FontTypeNames.FONTTYPE_FIGHT)
        Call WriteConsoleMsg(tempChr, UserList(UserIndex).Name & " te ha restaurado " & da�o & " puntos de vitalidad.", FontTypeNames.FONTTYPE_FIGHT)
    Else
        Call WriteConsoleMsg(UserIndex, "Te has restaurado " & da�o & " puntos de vitalidad.", FontTypeNames.FONTTYPE_FIGHT)
    End If
    b = True
ElseIf Hechizos(H).SubeMana = 2 Then
    If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
    If UserIndex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)
    End If
    
    Call InfoHechizo(UserIndex)
    
    If UserIndex <> tempChr Then
        Call WriteConsoleMsg(UserIndex, "Le has quitado " & da�o & " puntos de vitalidad a " & UserList(tempChr).Name, FontTypeNames.FONTTYPE_FIGHT)
        Call WriteConsoleMsg(tempChr, UserList(UserIndex).Name & " te ha quitado " & da�o & " puntos de vitalidad.", FontTypeNames.FONTTYPE_FIGHT)
    Else
        Call WriteConsoleMsg(UserIndex, "Te has quitado " & da�o & " puntos de vitalidad.", FontTypeNames.FONTTYPE_FIGHT)
    End If
    
    UserList(tempChr).Stats.MinSta = UserList(tempChr).Stats.MinSta - da�o
    
    If UserList(tempChr).Stats.MinSta < 1 Then UserList(tempChr).Stats.MinSta = 0
    b = True
End If

Call FlushBuffer(tempChr)

End Sub

Sub UpdateUserHechizos(ByVal UpdateAll As Boolean, ByVal UserIndex As Integer, ByVal Slot As Byte)

'Call LogTarea("Sub UpdateUserHechizos")

Dim LoopC As Byte

'Actualiza un solo slot
If Not UpdateAll Then

    'Actualiza el inventario
    If UserList(UserIndex).Stats.UserHechizos(Slot) > 0 Then
        Call ChangeUserHechizo(UserIndex, Slot, UserList(UserIndex).Stats.UserHechizos(Slot))
    Else
        Call ChangeUserHechizo(UserIndex, Slot, 0)
    End If

Else

'Actualiza todos los slots
For LoopC = 1 To MAXUSERHECHIZOS

        'Actualiza el inventario
        If UserList(UserIndex).Stats.UserHechizos(LoopC) > 0 Then
            Call ChangeUserHechizo(UserIndex, LoopC, UserList(UserIndex).Stats.UserHechizos(LoopC))
        Else
            Call ChangeUserHechizo(UserIndex, LoopC, 0)
        End If

Next LoopC

End If

End Sub

Sub ChangeUserHechizo(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal Hechizo As Integer)

'Call LogTarea("ChangeUserHechizo")

UserList(UserIndex).Stats.UserHechizos(Slot) = Hechizo


If Hechizo > 0 And Hechizo < NumeroHechizos + 1 Then
    
    Call WriteChangeSpellSlot(UserIndex, Slot)

Else

    Call WriteChangeSpellSlot(UserIndex, Slot)

End If


End Sub


Public Sub DesplazarHechizo(ByVal UserIndex As Integer, ByVal Dire As Integer, ByVal CualHechizo As Integer)

If (Dire <> 1 And Dire <> -1) Then Exit Sub
If Not (CualHechizo >= 1 And CualHechizo <= MAXUSERHECHIZOS) Then Exit Sub

Dim TempHechizo As Integer

If Dire = 1 Then 'Mover arriba
    If CualHechizo = 1 Then
        Call WriteConsoleMsg(UserIndex, "No puedes mover el hechizo en esa direccion.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    Else
        TempHechizo = UserList(UserIndex).Stats.UserHechizos(CualHechizo)
        UserList(UserIndex).Stats.UserHechizos(CualHechizo) = UserList(UserIndex).Stats.UserHechizos(CualHechizo - 1)
        UserList(UserIndex).Stats.UserHechizos(CualHechizo - 1) = TempHechizo

        'Prevent the user from casting other spells than the one he had selected when he hitted "cast".
        If UserList(UserIndex).flags.Hechizo > 0 Then
            UserList(UserIndex).flags.Hechizo = UserList(UserIndex).flags.Hechizo - 1
        End If
    End If
Else 'mover abajo
    If CualHechizo = MAXUSERHECHIZOS Then
        Call WriteConsoleMsg(UserIndex, "No puedes mover el hechizo en esa direccion.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    Else
        TempHechizo = UserList(UserIndex).Stats.UserHechizos(CualHechizo)
        UserList(UserIndex).Stats.UserHechizos(CualHechizo) = UserList(UserIndex).Stats.UserHechizos(CualHechizo + 1)
        UserList(UserIndex).Stats.UserHechizos(CualHechizo + 1) = TempHechizo

        'Prevent the user from casting other spells than the one he had selected when he hitted "cast".
        If UserList(UserIndex).flags.Hechizo > 0 Then
            UserList(UserIndex).flags.Hechizo = UserList(UserIndex).flags.Hechizo + 1
        End If
    End If
End If
End Sub

