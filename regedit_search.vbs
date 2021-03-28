```vbs
' Recherche récursivement une chaine dans les valeurs des clés du registre
'
' syntaxe : cls && cscript /nologo regedit_search.vbs "chaine_a_rechercher"
'  En arg1 la chaine a rechercher
' exemple : cscript /nologo regedit_search.vbs "\Macromed\Flash"


' Testé AVEC succès sur Windows 7 Pro 8GB de vRAM en 10 minutes (voir [2])

' Testé SANS succès sur Windows XP Pro (Mémoire insuffisante - voir [1])

'Option Explicit

' Read registry value according to type : https://www.sysadmins.lv/retired-msft-blogs/alejacma/how-to-read-a-registry-key-and-its-values-vbscript.aspx
' Array the easiest way : https://www.tutorialspoint.com/vbscript/vbscript_foreach_loop.htm
' Dynamically eval var : https://www.w3schools.com/asp/func_eval.asp

' ideas ...
' Split a\b\c to get a : https://www.w3schools.com/asp/func_split.asp
' String managing : https://support.smartbear.com/testcomplete/docs/scripting/working-with/strings/vbscript.html
' Pointers and eval : https://riteshbawaskar.wordpress.com/2009/08/25/pointers-in-vbscript/

' TODO : implémenter le cas où le type est REG_FULL_RESOURCE_DESCRIPTOR (id 9) (voir dessous)

' TODO : implementer le cas ou le type est REG_QWORD (id 11) (https://docs.microsoft.com/en-us/previous-versions/windows/desktop/regprov/enumvalues-method-in-class-stdregprov)

' TODO : implementer la gestion de la ruche HKEY_DYN_DATA (id &H80000006) (https://docs.microsoft.com/en-us/previous-versions/windows/desktop/regprov/enumvalues-method-in-class-stdregprov) 

' TODO : trouver une solution pour remettre "valeur non définie" pour une valeur par défaut (pas pour faire joli, mais car je ne sais pas si ça peut poser problème)

' TODO : accélérer le traitement avec du parallèlisme (multithreading) : voir ici https://www.itprotoday.com/programming-languages/how-multi-thread-vbscript-scripts

' TODO : corriger erreur "Mémoire insuffisante" pour oReg.EnumValues (voir [1])

' TODO : quand on trouve une valeur, afficher comme avec reg query pour pouvoir directement en faire un fichier REG


' Types de valeurs
' https://docs.microsoft.com/en-us/previous-versions/windows/desktop/regprov/enumvalues-method-in-class-stdregprov
' https://docs.microsoft.com/en-us/previous-versions/windows/it-pro/windows-xp/bb490984(v=technet.10)?redirectedfrom=MSDN
' https://www.chemtable.com/blog/en/types-of-registry-data.htm
' https://docs.microsoft.com/en-us/dotnet/api/microsoft.win32.registryvaluekind?view=net-5.0
' https://www.vulgarisation-informatique.com/base-registres-cles-fonctions.php
'  REG_FULL_RESOURCE_DESCRIPTOR : Ce type de données qui ne s'applique qu'à Windows XP contient des tableaux imbriqués stockant une liste de ressources correspondant à un matériel ou à un pilote
'  exemple : cas de la valeur "Configuration Data" de la clé HKEY_LOCAL_MACHINE\HARDWARE\DESCRIPTION\System
'  => apparemment non traite : https://bytes.com/topic/c-sharp/answers/251872-reading-registry-value-reg_full_resource_descriptor-returns-exception
'  mais en C quelqu'un propose une solution : https://stackoverflow.com/questions/50418343/which-struct-is-used-to-extract-information-from-a-registry-value-data-of-type-r
' id REG_FULL_RESOURCE_DESCRIPTOR : http://www.binaryworld.net/Main/ApiDetail.aspx?ApiId=47223
Const REG_SZ						= 1
Const REG_EXPAND_SZ					= 2
Const REG_BINARY					= 3
Const REG_DWORD						= 4
Const REG_MULTI_SZ					= 7
Const REG_FULL_RESOURCE_DESCRIPTOR	= 9
'Const REG_QWORD						= 11
' autres types non implémentés ici :
' REG_DWORD_BIG_ENDIAN
' REG_DWORD_LITTLE_ENDIAN
' REG_LINK

Const FIELD_NAME					= 0
Const FIELD_ID						= 1

Const REG_MULTI_SZ_SEPARATOR		= " | "
Const DQUOTE 						= """"

If WScript.Arguments.Count = 0 Then
 WScript.Echo DQUOTE & "Chaine_a_rechercher" & DQUOTE & " non trouvee !" &_
  vbCrLf &_
   "syntaxe : cls && cscript /nologo regedit_search.vbs " &_
   DQUOTE & "chaine_a_rechercher" & DQUOTE
  WScript.Quit 1
End If

stringToSearch = WScript.Arguments ( 0 )

strComputer = "."

Set oReg = GetObject ( "winmgmts:{impersonationLevel=impersonate}!\\" & _ 
  strComputer & "\root\default:StdRegProv" )
  
' Dim hives, hive, key
' Final
hives = Array ( _
 Array ( "HKEY_CLASSES_ROOT", 	&H80000000 ), _
 Array ( "HKEY_CURRENT_USER", 	&H80000001 ), _
 Array ( "HKEY_LOCAL_MACHINE", 	&H80000002 ), _
 Array ( "HKEY_USERS", 			&H80000003 ), _
 Array ( "HKEY_CURRENT_CONFIG", &H80000005 ), _
 Array ( "HKEY_DYN_DATA", 		&H80000006 ) _
)

' Debug
' https://stackoverflow.com/questions/17663365/visual-basic-scripting-dynamic-array/17664578#17664578
' hives = Array()
' ' ReDim Preserve hives ( UBound ( hives ) + 1 ) : hives ( UBound ( hives ) ) = Array ( "HKEY_CLASSES_ROOT",	&H80000000 )
' ' ReDim Preserve hives ( UBound ( hives ) + 1 ) : hives ( UBound ( hives ) ) = Array ( "HKEY_CURRENT_USER",	&H80000001 )
' ReDim Preserve hives ( UBound ( hives ) + 1 ) : hives ( UBound ( hives ) ) = Array ( "HKEY_LOCAL_MACHINE",	&H80000002 )
' ' ReDim Preserve hives ( UBound ( hives ) + 1 ) : hives ( UBound ( hives ) ) = Array ( "HKEY_USERS",			&H80000003 )
' ' ReDim Preserve hives ( UBound ( hives ) + 1 ) : hives ( UBound ( hives ) ) = Array ( "HKEY_CURRENT_CONFIG",	&H80000005 )
' ' ReDim Preserve hives ( UBound ( hives ) + 1 ) : hives ( UBound ( hives ) ) = Array ( "HKEY_DYN_DATA",			&H80000006 )  
 
Sub searchInValues ( hive, strKeyPath )
 'WScript.Echo "> " & hive ( FIELD_NAME ) & "\" & strKeyPath
 
 oReg.EnumValues hive ( FIELD_ID ), strKeyPath, values, types
 
 If IsArray ( types ) Then
  'WScript.Echo "Nb de valeurs: " & UBound ( values ) - 1 & ", Nb de types: " & UBound ( types ) - 1
  For i = LBound ( types ) To UBound ( types )
   value = values ( i )
   'WScript.Echo "val num " & i & " = " & value
   Select Case types ( i )
    Case REG_SZ
     val_type = "REG_SZ"
     oReg.GetStringValue hive ( FIELD_ID ), strKeyPath, value, data
    Case REG_EXPAND_SZ
     val_type = "REG_EXPAND_SZ"
     oReg.GetExpandedStringValue hive ( FIELD_ID ), strKeyPath, value, data
    Case REG_BINARY
	 val_type = "REG_BINARY"
     oReg.GetBinaryValue hive ( FIELD_ID ), strKeyPath, value, bytes
     data = ""
     For Each val_byte In bytes
      data = data & Hex ( val_byte ) & " "
     Next
    Case REG_DWORD
	 val_type = "REG_DWORD"
     oReg.GetDWORDValue hive ( FIELD_ID ), strKeyPath, value, data
    Case REG_MULTI_SZ
	 val_type = "REG_MULTI_SZ"
     oReg.GetMultiStringValue hive ( FIELD_ID ), strKeyPath, value, strings
     data = ""
     For Each str In strings
	  If data <> "" Then
	   data = data & REG_MULTI_SZ_SEPARATOR
	  End If
      data = data & str
     Next	
    Case REG_FULL_RESOURCE_DESCRIPTOR
	 val_type = "REG_FULL_RESOURCE_DESCRIPTOR"
	 ' GetStringValue n'est pas la procédure appropriée à ce type de valeur, 
	 '  mais commme il n'existe pas de procédure appropriée (en tout cas 
	 '  je ne l'ai pas trouvée), on peut au moins, par cette astuce, 
	 '  récupérer le nom de la valeur.
	 oReg.GetStringValue hive ( FIELD_ID ), strKeyPath, value, data
	 data = "<unknown>"
   End Select
   'If types ( i ) = REG_SZ Then WScript.Echo "<" & data & ">" End If
   If value = "" Then
    value = "(Par defaut)"
   End If
   If IsNull ( data ) Or data = "" Then
    ' Nota : une valeur par défaut sera détectée que si ses données ont été 
	'  définies (cad modifiées). Si depuis regedit pour les données il est 
	'  indiqué "valeur non définie" elle ne sera pas détectée, s'il n'y a rien
	'  d'affiché (les données sont vides) elle sera détectée vide.
    'data = "<vide>" : WScript.Echo hive ( FIELD_NAME ) & "\" & strKeyPath & " : "  & value & " [" & val_type & "]" & " = " & data	
   Else
    If InStr ( data, stringToSearch ) Then
     WScript.Echo "# " & hive ( FIELD_NAME ) & "\" & strKeyPath & " : "  &_
      value & " [" & val_type & "]" & " = " & data
    End If
    'WScript.Echo hive ( FIELD_NAME ) & "\" & strKeyPath & " : "  & value & _
    ' " [" & val_type & "]" & " = " & data
   End If
  Next
 Else
  'WScript.Echo "Non ce n'est pas un tableau ; il n'y a pas de valeurs ici"
 End If
End Sub
 
Sub scanKeys ( hive, strKeyPath )
 'WScript.Echo "strKeyPath recue: <" & strKeyPath & ">"
 
 oReg.EnumKey hive ( FIELD_ID ), strKeyPath, arrSubKeys
 
 searchInValues hive, strKeyPath
 
 If IsArray ( arrSubKeys ) Then
  For Each subkey In arrSubKeys
   ' If key path empty => HKLM\<empty>\... => HKLM\\... => impossible !
   If strKeyPath <> "" Then
    subkey = strKeyPath & "\" & subkey
   End If
   'WScript.Echo "strKeyPath envoye: <" & subkey & ">"
   scanKeys hive, subkey
  Next
 End If
End Sub

For Each hive In hives
 WScript.Echo "Ruche: " & hive ( FIELD_NAME ) & " (&" & Hex ( hive ( FIELD_ID ) ) & ")"
 
 strKeyPath = ""
 'strKeyPath = "HARDWARE"
 'strKeyPath = "HARDWARE\DESCRIPTION"
 'strKeyPath = "HARDWARE\DESCRIPTION\System"
 'strKeyPath = "HARDWARE\DESCRIPTION\System\CentralProcessor"
 'strKeyPath = "HARDWARE\DESCRIPTION\System\CentralProcessor\0"
 'WScript.Echo "strKeyPath envoye: <" & strKeyPath & ">"
 scanKeys hive, strKeyPath
 Set hive = Nothing
Next

Set oReg = Nothing
Set hives = Nothing

WScript.Quit 0 
```