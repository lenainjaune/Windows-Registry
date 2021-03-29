```vbs
' Recherche récursivement une chaine dans les valeurs des clés du registre
'
' syntaxe : cls && cscript /nologo regedit_search.vbs "chaine_a_rechercher"
'  En arg1 la chaine a rechercher
' exemple : cscript /nologo regedit_search.vbs "\Macromed\Flash"


' Testé AVEC succès sur Windows 7 Pro 8GB de vRAM en 10 minutes (voir [2])

' Testé SANS succès sur Windows XP Pro (Mémoire insuffisante - voir [1])


' Read registry value according to type : https://www.sysadmins.lv/retired-msft-blogs/alejacma/how-to-read-a-registry-key-and-its-values-vbscript.aspx
' Array the easiest way : https://www.tutorialspoint.com/vbscript/vbscript_foreach_loop.htm
' Dynamically eval var : https://www.w3schools.com/asp/func_eval.asp

' ideas ...
' Split a\b\c to get a : https://www.w3schools.com/asp/func_split.asp
' String managing : https://support.smartbear.com/testcomplete/docs/scripting/working-with/strings/vbscript.html
' Pointers and eval : https://riteshbawaskar.wordpress.com/2009/08/25/pointers-in-vbscript/


' TODO : implémenter le cas où le type est REG_FULL_RESOURCE_DESCRIPTOR (id 9) (voir dessous)

' TODO : implementer le cas ou le type est REG_QWORD (id 11) (https://docs.microsoft.com/en-us/previous-versions/windows/desktop/regprov/enumvalues-method-in-class-stdregprov)

' TODO : implementer la gestion de la rootKey HKEY_DYN_DATA (id &H80000006) (https://docs.microsoft.com/en-us/previous-versions/windows/desktop/regprov/enumvalues-method-in-class-stdregprov) 

' TODO : trouver une solution pour remettre "valeur non définie" pour une valeur par défaut (pas pour faire joli, mais car je ne sais pas si ça peut poser problème)

' TODO : accélérer le traitement avec du parallèlisme (multithreading) : voir ici https://www.itprotoday.com/programming-languages/how-multi-thread-vbscript-scripts

' TODO : corriger erreur Windows XP "Mémoire insuffisante" pour registry.EnumValues (voir [1])

' TODO : nettoyer les tests de debug pour les finaliser

' TODO : debug avec niveau de verbosité (1 affiche la progression, 2 affiche presque tout, 3 affiche tout

' TODO : choisir une clé racine au lieu de tout scanner pertinent ou pas ? En arg2 ?





'---------------------------------- DECLARE ------------------------------------

'Option Explicit

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

Const WSH_STATE_RUNNING				= 0
Const WSH_STATE_FINISHED			= 1
Const WSH_STATE_FAILED				= 2

Const REG_MULTI_SZ_SEPARATOR		= " | "
Const DQUOTE 						= """"




'-------------------------------- FUNCTIONS ------------------------------------


Sub echoOut ( rootKeyName, keyPath, value, valueType, valueData )
 If value = "" Then
  value = "(Par defaut)"
 End If
 
 WScript.Echo rootKeyName & "\" & keyPath &_
  " : "  & value & " [" & valueType & "]" & " = " & valueData
End Sub
  
 
Sub searchInValues ( rootKey, keyPath )
 'WScript.Echo "> " & rootKey ( FIELD_NAME ) & "\" & keyPath
 
 registry.EnumValues rootKey ( FIELD_ID ), keyPath, values, types
 
 If IsArray ( types ) Then
  'WScript.Echo "Nb de values: " & UBound ( values ) - 1 & ", Nb de types: " & UBound ( types ) - 1
  For i = LBound ( types ) To UBound ( types )
   value = values ( i )
   'WScript.Echo "val num " & i & " = " & value
   Select Case types ( i )
    Case REG_SZ
     valueType = "REG_SZ"
     registry.GetStringValue rootKey ( FIELD_ID ), keyPath, value, valueData
    Case REG_EXPAND_SZ
     valueType = "REG_EXPAND_SZ"
     registry.GetExpandedStringValue rootKey ( FIELD_ID ), keyPath, value, _
	  valueData
    Case REG_BINARY
	 valueType = "REG_BINARY"
     registry.GetBinaryValue rootKey ( FIELD_ID ), keyPath, value, bytes
     valueData = ""
     For Each val_byte In bytes
      valueData = valueData & Hex ( val_byte ) & " "
     Next
    Case REG_DWORD
	 valueType = "REG_DWORD"
     registry.GetDWORDValue rootKey ( FIELD_ID ), keyPath, value, valueData
    Case REG_MULTI_SZ
	 valueType = "REG_MULTI_SZ"
     registry.GetMultiStringValue rootKey ( FIELD_ID ), keyPath, value, strings
     valueData = ""
     For Each str In strings
	  If valueData <> "" Then
	   valueData = valueData & REG_MULTI_SZ_SEPARATOR
	  End If
      valueData = valueData & str
     Next	
    Case REG_FULL_RESOURCE_DESCRIPTOR
	 valueType = "REG_FULL_RESOURCE_DESCRIPTOR"
	 ' GetStringValue n'est pas la procédure appropriée à ce type de valeur, 
	 '  mais commme il n'existe pas de procédure appropriée (en tout cas 
	 '  je ne l'ai pas trouvée), on peut au moins, par cette astuce, 
	 '  récupérer le nom de la valeur.
	 registry.GetStringValue rootKey ( FIELD_ID ), keyPath, value, valueData
	 valueData = "<unknown>"
   End Select
   'If types ( i ) = REG_SZ Then WScript.Echo "<" & valueData & ">" End If
   If IsNull ( valueData ) Or valueData = "" Then
    ' Nota : une valeur par défaut sera détectée que si ses données ont été 
	'  définies (c.a.d. modifiées). Si depuis regedit pour les données il est 
	'  indiqué "valeur non définie" elle ne sera pas détectée, s'il n'y a rien
	'  d'affiché (les données sont vides) elle sera détectée vide.
    'echoOut rootKey ( FIELD_NAME ), keyPath, value, valueType, "<vide>"
   Else
    If InStr ( valueData, stringToSearch ) Then
	 WScript.Echo "# " & rootKey ( FIELD_NAME ) & "\" & keyPath & " : "  &_
      value & " [" & valueType & "]" & " = " & valueData
    End If
	'echoOut rootKey ( FIELD_NAME ), keyPath, value, valueType, valueData
   End If
  Next
 Else
  'WScript.Echo "Non ce n'est pas un tableau ; il n'y a pas de valeurs ici"
 End If
End Sub

 
Sub scanKeys ( rootKey, keyPath )
 'WScript.Echo "keyPath recu: <" & keyPath & ">"
 
 registry.EnumKey rootKey ( FIELD_ID ), keyPath, subKeys
 
 searchInValues rootKey, keyPath
 
 If IsArray ( subKeys ) Then
  For Each subKey In subKeys
   ' If keyPath empty => HKLM\<empty>\... => HKLM\\... => impossible !
   If keyPath <> "" Then
    subKey = keyPath & "\" & subKey
   End If
   'WScript.Echo "keyPath envoye: <" & subKey & ">"
   scanKeys rootKey, subKey
  Next
 End If
End Sub


'----------------------------------- MAIN --------------------------------------

If WScript.Arguments.Count = 0 Then
 WScript.Echo DQUOTE & "Chaine_a_rechercher" & DQUOTE & " non trouvee !" &_
  vbCrLf &_
   "syntaxe : cls && cscript /nologo regedit_search.vbs " &_
   DQUOTE & "chaine_a_rechercher" & DQUOTE
 WScript.Quit 1
End If

stringToSearch = WScript.Arguments ( 0 )
'WScript.Echo stringToSearch

computer = "."

Set registry = GetObject ( "winmgmts:{impersonationLevel=impersonate}!\\" & _ 
  computer & "\root\default:StdRegProv" )
    
' Clés racine  
' Full
rootKeys = Array ( _
 Array ( "HKEY_CLASSES_ROOT", 	&H80000000 ), _
 Array ( "HKEY_CURRENT_USER", 	&H80000001 ), _
 Array ( "HKEY_LOCAL_MACHINE", 	&H80000002 ), _
 Array ( "HKEY_USERS", 			&H80000003 ), _
 Array ( "HKEY_CURRENT_CONFIG", &H80000005 ), _
 Array ( "HKEY_DYN_DATA", 		&H80000006 ) _
)

' Debug
' https://stackoverflow.com/questions/17663365/visual-basic-scripting-dynamic-array/17664578#17664578
' rootKeys = Array()
' ' ReDim Preserve rootKeys ( UBound ( rootKeys ) + 1 ) : rootKeys ( UBound ( rootKeys ) ) = Array ( "HKEY_CLASSES_ROOT",	&H80000000 )
' ' ReDim Preserve rootKeys ( UBound ( rootKeys ) + 1 ) : rootKeys ( UBound ( rootKeys ) ) = Array ( "HKEY_CURRENT_USER",	&H80000001 )
' ReDim Preserve rootKeys ( UBound ( rootKeys ) + 1 ) : rootKeys ( UBound ( rootKeys ) ) = Array ( "HKEY_LOCAL_MACHINE",	&H80000002 )
' ' ReDim Preserve rootKeys ( UBound ( rootKeys ) + 1 ) : rootKeys ( UBound ( rootKeys ) ) = Array ( "HKEY_USERS",			&H80000003 )
' ' ReDim Preserve rootKeys ( UBound ( rootKeys ) + 1 ) : rootKeys ( UBound ( rootKeys ) ) = Array ( "HKEY_CURRENT_CONFIG",	&H80000005 )
' ' ReDim Preserve rootKeys ( UBound ( rootKeys ) + 1 ) : rootKeys ( UBound ( rootKeys ) ) = Array ( "HKEY_DYN_DATA",			&H80000006 )  

For Each rootKey In rootKeys
 WScript.Echo "; " & rootKey ( FIELD_NAME ) & " (&" & Hex ( rootKey ( FIELD_ID ) ) & ")"
 
 keyPath = ""
 'keyPath = "HARDWARE"
 'keyPath = "HARDWARE\DESCRIPTION"
 'keyPath = "HARDWARE\DESCRIPTION\System"
 'keyPath = "HARDWARE\DESCRIPTION\System\CentralProcessor"
 'keyPath = "HARDWARE\DESCRIPTION\System\CentralProcessor\0"
 'WScript.Echo "keyPath envoye: <" & keyPath & ">"
 scanKeys rootKey, keyPath
 Set rootKey = Nothing
Next

Set registry = Nothing
Set rootKeys = Nothing

WScript.Quit 0



'----------------------------------- NOTES -------------------------------------

' [1] voir aussi [1']
' Testé SANS succès sur Windows XP Pro (Mémoire insuffisante - voir [1])
' 4GB de vRAM mais Windows ne voit que 3GB => limitation technique
'  (https://answers.microsoft.com/en-us/windows/forum/windows_xp-windows_install/xp-32-see-only-3gb-ram-bios-and-pc-wizzard-see/bf65bbf4-5703-40b7-8d1f-cdea75702a22)
'  Ceci étant, la mémoire n'avait pas l'air d'être sollicité plus que ça ; j'opte plutôt pour un problème de gestion de mémoire avec VBScript
' C:\Documents and Settings\root>cscript /nologo Bureau\regedit_search.vbs "BOCHS"
' # HKEY_LOCAL_MACHINE\HARDWARE\DESCRIPTION\System : SystemBiosVersion [REG_MULTI_SZ] = BOCHS  - 1
' { C:\Documents and Settings\user\Bureau\regedit_search.vbs(133, 2) Erreur d'exécution Microsoft VBScript: Mémoire insuffisante: 'registry.EnumValues'
' test 1
' contexte du script :
' keyPath = ""
' juste la rootKey HKLM : ReDim Preserve rootKeys ( UBound ( rootKeys ) + 1 ) : rootKeys ( UBound ( rootKeys ) ) = Array ( "HKEY_LOCAL_MACHINE",	&H80000002 )
' test 2 : reboot + même test => mêmes symptômes
' test 3 : j'ai activé l'affichage des valeurs trouvées => message équivalent
' C:\Documents and Settings\user\Bureau\regedit_search.vbs(201, 5) (null): Espace insuffisant pour traiter cette commande.
' ligne 201 j'ai la commande d'affichage des valeurs trouvée :
'    WScript.Echo rootKey ( FIELD_NAME ) & "\" & keyPath & " : "  & value & _
'     " [" & valueType & "]" & " = " & valueData

' [1']
' 28/03/2021
' 21:12
' rootKey: HKEY_CLASSES_ROOT (&80000000)
' # HKEY_CLASSES_ROOT\CLSID\{1171A62F-05D2-11D1-83FC-00A0C9089C5A}\InprocServer32 :  [REG_SZ] = C:\WINDOWS\system32\Macromed\Flash\Flash6.ocx
' # HKEY_CLASSES_ROOT\CLSID\{D27CDB6E-AE6D-11cf-96B8-444553540000}\InprocServer32 :  [REG_SZ] = C:\WINDOWS\system32\Macromed\Flash\Flash6.ocx
' # HKEY_CLASSES_ROOT\CLSID\{D27CDB70-AE6D-11cf-96B8-444553540000}\InprocServer32 :  [REG_SZ] = C:\WINDOWS\system32\Macromed\Flash\Flash6.ocx
' rootKey: HKEY_CURRENT_USER (&80000001)
' rootKey: HKEY_LOCAL_MACHINE (&80000002)
' # HKEY_LOCAL_MACHINE\SOFTWARE\Classes\CLSID\{1171A62F-05D2-11D1-83FC-00A0C9089C5A}\InprocServer32 :  [REG_SZ] = C:\WINDOWS\system32\Macromed\Flash\Flash6.ocx
' # HKEY_LOCAL_MACHINE\SOFTWARE\Classes\CLSID\{D27CDB6E-AE6D-11cf-96B8-444553540000}\InprocServer32 :  [REG_SZ] = C:\WINDOWS\system32\Macromed\Flash\Flash6.ocx
' # HKEY_LOCAL_MACHINE\SOFTWARE\Classes\CLSID\{D27CDB70-AE6D-11cf-96B8-444553540000}\InprocServer32 :  [REG_SZ] = C:\WINDOWS\system32\Macromed\Flash\Flash6.ocx
' # HKEY_LOCAL_MACHINE\SOFTWARE\Macromedia\FlashPlayerPlugin : PlayerPath [REG_SZ] = C:\WINDOWS\system32\Macromed\Flash\NPSWF32.dll
' # HKEY_LOCAL_MACHINE\SOFTWARE\Macromedia\FlashPlayerPlugin : UninstallerPath [REG_SZ] = C:\WINDOWS\system32\Macromed\Flash\FlashUtil10m_Plugin.exe
' # HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Adobe Flash Player Plugin : UninstallString [REG_SZ] = C:\WINDOWS\system32\Macromed\Flash\FlashUtil10m_Plugin.exe -maintain plugin
' # HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Adobe Flash Player Plugin : DisplayIcon [REG_SZ] = C:\WINDOWS\system32\Macromed\Flash\FlashUtil10m_Plugin.exe
' # HKEY_LOCAL_MACHINE\SOFTWARE\MozillaPlugins\@adobe.com/FlashPlayer : Path [REG_SZ] = C:\WINDOWS\system32\Macromed\Flash\NPSWF32.dll
' # HKEY_LOCAL_MACHINE\SOFTWARE\MozillaPlugins\@adobe.com/FlashPlayer : XPTPath [REG_SZ] = C:\WINDOWS\system32\Macromed\Flash\flashplayer.xpt
' C:\Documents and Settings\user\Bureau\regedit_search.vbs(128, 2) Erreur d'exécution Microsoft VBScript: Mémoire insuffisante: 'registry.EnumValues'
' 28/03/2021
' 21:18


' [2] (voir aussi [2'] pour pouvoir réimporter)
' Testé AVEC succès sur Windows 7 Pro 8GB de vRAM en 10 minutes (voir [2])
' C:\Users\user>cls && date /t & time /t & cscript /nologo Desktop\regedit_search.vbs "\Macromed\Flash" & date /t & time /t
' 28/03/2021
' 20:44
' rootKey: HKEY_CLASSES_ROOT (&80000000)
' rootKey: HKEY_CURRENT_USER (&80000001)
' rootKey: HKEY_LOCAL_MACHINE (&80000002)
' # HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Macromedia\FlashPlayerPlugin : PlayerPath [REG_SZ] = C:\Windows\SysWOW64\Macromed\Flash\NPSWF32.dll
' # HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Macromedia\FlashPlayerPlugin : UninstallerPath [REG_SZ] = C:\Windows\SysWOW64\Macromed\Flash\FlashUtil10m_Plugin.exe
' # HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Adobe Flash Player Plugin : UninstallString [REG_SZ] = C:\Windows\SysWOW64\Macromed\Flash\FlashUtil10m_Plugin.exe -maintain plugin
' # HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Adobe Flash Player Plugin : DisplayIcon [REG_SZ] = C:\Windows\SysWOW64\Macromed\Flash\FlashUtil10m_Plugin.exe
' # HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\MozillaPlugins\@adobe.com/FlashPlayer : Path [REG_SZ] = C:\Windows\SysWOW64\Macromed\Flash\NPSWF32.dll
' # HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\MozillaPlugins\@adobe.com/FlashPlayer : XPTPath [REG_SZ] = C:\Windows\SysWOW64\Macromed\Flash\flashplayer.xpt
' rootKey: HKEY_USERS (&80000003)
' rootKey: HKEY_CURRENT_CONFIG (&80000005)
' rootKey: HKEY_DYN_DATA (&80000006)
' 28/03/2021
' 20:54



' [2']
' Liste les clés, valeurs qui peuvent directement être réimportées dans regedit
' C:\Users\user>cls && date /t & time /t & cscript /nologo Desktop\regedit_search.vbs "\Macromed\Flash" & date /t & time /t
' 29/03/2021
' 11:19
' ; rootKey: HKEY_CLASSES_ROOT (&80000000)
' ; rootKey: HKEY_CURRENT_USER (&80000001)
' ; rootKey: HKEY_LOCAL_MACHINE (&80000002)
' HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Macromedia\FlashPlayerPlugin
'   PlayerPath    REG_SZ    C:\Windows\SysWOW64\Macromed\Flash\NPSWF32.dll
'
' HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Macromedia\FlashPlayerPlugin
'   UninstallerPath    REG_SZ    C:\Windows\SysWOW64\Macromed\Flash\FlashUtil10m_Plugin.exe
'
' HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Adobe Flash Player Plugin
'   UninstallString    REG_SZ    C:\Windows\SysWOW64\Macromed\Flash\FlashUtil10m_Plugin.exe -maintain plugin
'
' HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Adobe Flash Player Plugin
'   DisplayIcon    REG_SZ    C:\Windows\SysWOW64\Macromed\Flash\FlashUtil10m_Plugin.exe
'
' HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\MozillaPlugins\@adobe.com/FlashPlayer
'   Path    REG_SZ    C:\Windows\SysWOW64\Macromed\Flash\NPSWF32.dll
'
' HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\MozillaPlugins\@adobe.com/FlashPlayer
'   XPTPath    REG_SZ    C:\Windows\SysWOW64\Macromed\Flash\flashplayer.xpt
'
' ; rootKey: HKEY_USERS (&80000003)
' ; rootKey: HKEY_CURRENT_CONFIG (&80000005)
' ; rootKey: HKEY_DYN_DATA (&80000006)
' 29/03/2021
' 11:29

```
