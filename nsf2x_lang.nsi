# The translated strings used by the NSF2X isntaller

!insertmacro MUI_LANGUAGE English
!insertmacro MUI_LANGUAGE French

LangString 	COMMANDLINE_HELP		${LANG_English}	"\
      /ALL : Extract application for all users$\n\
      /SHORTCUT : Install desktop shortcut$\n\
      /S : Silent install$\n\
      /D=%directory% : Specify destination directory$\n"
LangString 	COMMANDLINE_HELP		${LANG_French}	"\
      /ALL : Extraire l'application pour tous les utilisateurs$\n\
      /SHORTCUT : Installer la raccourci sur le bureau$\n\
      /S : Installation silencieuse$\n\
      /D=%répertoire% : Spécifier le répertoire destination$\n"

LangString 	HAVE_CLICKTORUN		${LANG_English}	"\
You appear to have a $\"Click To Run$\" version of Outlook installed. This will interfere with \
the conversion to Outlook PST files. You have three choices$\r$\n\
$\r$\n\
1. Allow NSF2X to patch the registry. In this case the NSF2X installer must run with administrator \
privileges. This fix will also cause issues if you are running multiple versions of Outlook. \
After patching, your $\"Click To Run$\" Outlook client will work correctly, but older versions \
of Outlook might fail in unexplained manners. See$\r$\n\
$\r$\n\
https://blogs.msdn.microsoft.com/stephen_griffin/2014/04/21/outlook-2013-click-to-run-and-com-interfaces/$\r$\n\
$\r$\n\
for more information$\r$\n\
2. Continue the installation knowing that the conversion to PST will not be possible$\r$\n\
3. Cancel the installation of NSF2X$\r$\n\
$\r$\n\
Do you wish to let NSF2X patch the registry ?"
LangString 	HAVE_CLICKTORUN		${LANG_French}	"\
Vous semblez d'avoir un version $\"Click To Run$\" d'Outlook installé. La conversion vers \
des fichiers PST d'Outlook va être impacté. Vous avez 3 choix$\r$\n\
$\r$\n\
1. Permettre NSF2X à modifier le registre. L'installer de NSF2X doit avoir les privilèges \
d'un administrateur. Cette solution pourrait poser des problèmes si vous avez plusieurs versions \
d'Outlook installés. Après modification du registre le client $\"Click To Run$\" Outlook \
fonctionnera correctement, mais les autres version d'Outlook pourrait échouer de manière \
imprévisible. Voir$\r$\n\
$\r$\n\
https://blogs.msdn.microsoft.com/stephen_griffin/2014/04/21/outlook-2013-click-to-run-and-com-interfaces/$\r$\n\
$\r$\n\
pour plus d'information$\r$\n\
2. Continuer l'installation sachant que la conversion vers des fichiers PST ne serait pas \
possible$\r$\n\
3. Annuler l'installation de NSF2X$\r$\n\
$\r$\n\
Est-ce que vous voulez permettre NSF2X à modifier le registre ?"

LangString 	FAIL_COPY_ICONV		${LANG_English}	"Failed to Copy $\"CLSID_IConverterSession$\" registry value. Aborting"
LangString 	FAIL_COPY_ICONV		${LANG_French}	"Le copie du registre $\"CLSID_IConverterSession$\" à échouer. Abandonner"

LangString 	FAIL_COPY_MIME		${LANG_English}	"Failed to Copy $\"CLSID_IMimeMessage$\" registry value. Aborting"
LangString 	FAIL_COPY_MIME		${LANG_French}	"Le copie du registre $\"CLSID_IMimeMessage$\" à échouer. Abandonner"

LangString 	ADMIN_RIGHTS		${LANG_English}	"You need administrator rights to install ${NAME}"
LangString 	ADMIN_RIGHTS		${LANG_French}	"Vous avez besoin les droits d'un administrateur pour l'installation de ${NAME}"

LangString 	INSTALL_MODE		${LANG_English}	"Install Mode"
LangString 	INSTALL_MODE		${LANG_French}	"Mode d'installation"

LangString 	INSTALL_CHOOSE		${LANG_English}	"Choose how you want to install ${NAME}."
LangString 	INSTALL_CHOOSE		${LANG_French}	"Choisir comment vous voulez installer ${NAME}."

LangString 	INSTALL_ADMIN		${LANG_English}	"Install for all users (requires administrator privileges)"
LangString 	INSTALL_ADMIN		${LANG_French}	"Installer pour tous les utilisateurs (les droits d'un administrateur sont nécessaire)"

LangString 	INSTALL_SHORTCUT		${LANG_English}	"Create desktop shortcut"
LangString 	INSTALL_SHORTCUT		${LANG_French}	"Créer un raccourci sur le bureau"

LangString 	ALREADY_INSTALLED		${LANG_English}	"NSF2X version ${VERSION} is already installed. Launch the uninstaller?"
LangString 	ALREADY_INSTALLED		${LANG_French}	"La version ${VERSION} de NSF2X est déjà installé. Désinstaller ?"

LangString 	COMMANDLINE_HELP2		${LANG_English}	"\
      /ALL : Remove application for all users$\n\
      /USER : Remove application for the current user$\n\
      /S : Silent$\n"
LangString 	COMMANDLINE_HELP2		${LANG_French}	"\
      /ALL : Supprimer l'application pour tous les utilisateurs$\n\
      /USER : Supprimer l'application pour l'utilisateur courant$\n\
      /S : Silencieuse$\n"

LangString 	UNINSTALL_REGISTRY	${LANG_English}	"Do you wish to let NSF2X remove the modifications it made to the registry to support $\"Click To Run$\" ?"
LangString 	UNINSTALL_REGISTRY	${LANG_French}	"Est-ce que vous voulez laisser NSF2X supprimer les modifications qu’il a fait au registre afin de $\"Click To Run$\" ?"