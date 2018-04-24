# The translated strings used by the NSF2X isntaller

!insertmacro MUI_LANGUAGE English
!insertmacro MUI_LANGUAGE French
!insertmacro MUI_LANGUAGE German

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
LangString 	COMMANDLINE_HELP		${LANG_German}	"\
      /ALL : Anwendung für alle Benutzer extrahieren$\n\
      /SHORTCUT : Desktop-Verknüpfung installieren$\n\
      /S : Stille Installation$\n\
      /D=%pfad% : Zielverzeichnis angeben$\n"

LangString 	ADMIN_RIGHTS		${LANG_English}	"You need administrator rights to install ${NAME}"
LangString 	ADMIN_RIGHTS		${LANG_French}	"Vous avez besoin les droits d'un administrateur pour l'installation de ${NAME}"
LangString 	ADMIN_RIGHTS		${LANG_German}	"Sie benötigen Administratorrechte für die Installation von ${NAME}"

LangString 	INSTALL_MODE		${LANG_English}	"Install Mode"
LangString 	INSTALL_MODE		${LANG_French}	"Mode d'installation"
LangString 	INSTALL_MODE		${LANG_German}	"Installationsmodus"

LangString 	INSTALL_CHOOSE		${LANG_English}	"Choose how you want to install ${NAME}."
LangString 	INSTALL_CHOOSE		${LANG_French}	"Choisir comment vous voulez installer ${NAME}."
LangString 	INSTALL_CHOOSE		${LANG_German}	"Wählen Sie, wie Sie ${NAME} installieren möchten."

LangString 	INSTALL_ADMIN		${LANG_English}	"Install for all users (requires administrator privileges)"
LangString 	INSTALL_ADMIN		${LANG_French}	"Installer pour tous les utilisateurs (les droits d'un administrateur sont nécessaire)"
LangString 	INSTALL_ADMIN		${LANG_German}	"Installation für alle Benutzer (erfordert Administratorrechte)"

LangString 	INSTALL_SHORTCUT		${LANG_English}	"Create desktop shortcut"
LangString 	INSTALL_SHORTCUT		${LANG_French}	"Créer un raccourci sur le bureau"
LangString 	INSTALL_SHORTCUT		${LANG_German}	"Desktopverknüpfung erstellen

LangString 	ALREADY_INSTALLED		${LANG_English}	"NSF2X version ${VERSION} is already installed. Launch the uninstaller?"
LangString 	ALREADY_INSTALLED		${LANG_French}	"La version ${VERSION} de NSF2X est déjà installé. Désinstaller?"
LangString 	ALREADY_INSTALLED		${LANG_German}	"NSF2X-Version ${VERSION} ist bereits installiert. Das Deinstallationsprogramm starten?"

LangString 	COMMANDLINE_HELP2		${LANG_English}	"\
      /ALL : Remove application for all users$\n\
      /USER : Remove application for the current user$\n\
      /S : Silent$\n"
LangString 	COMMANDLINE_HELP2		${LANG_French}	"\
      /ALL : Supprimer l'application pour tous les utilisateurs$\n\
      /USER : Supprimer l'application pour l'utilisateur courant$\n\
      /S : Silencieuse$\n"
LangString 	COMMANDLINE_HELP2		${LANG_German}	"\
      /ALL : Anwendung für alle Benutzer entfernen$\n\
      /USER : Anwendung für aktuellen Benutzer entfernen$\n\
      /S : Still$\n"

LangString 	UNINSTALL_REGISTRY	${LANG_English}	"Do you wish to let ${NAME} remove the modifications it made to the registry to support $\"Click To Run$\" ?"
LangString 	UNINSTALL_REGISTRY	${LANG_French}	"Est-ce que vous voulez laisser ${NAME} supprimer les modifications qu’il a fait au registre afin de supporter $\"Click To Run$\" ?"
LangString 	UNINSTALL_REGISTRY	${LANG_German}	"Möchten Sie zulassen, dass ${NAME} die Änderungen an der Registrierung um $\"Click To Run $\" zu unterstützen entfernt?"
