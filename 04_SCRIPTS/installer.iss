; Script de Inno Setup para EqualityMomentum
; Crea un instalador profesional con identidad corporativa

#define MyAppName "EqualityMomentum"
#define MyAppVersion "1.0.1"
#define MyAppPublisher "EqualityMomentum"
#define MyAppURL "https://github.com/juanfranartes/EqualityMomentum"
#define MyAppExeName "EqualityMomentum.exe"

[Setup]
; Información de la aplicación
AppId={{A1B2C3D4-E5F6-7890-ABCD-EF1234567890}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
AllowNoIcons=yes
LicenseFile=..\LICENSE.txt
; Icono del instalador (si existe)
; SetupIconFile=..\00_DOCUMENTACION\isotipo.ico
OutputDir=..\Instaladores
OutputBaseFilename=EqualityMomentum_Setup_v{#MyAppVersion}
Compression=lzma
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=admin
; Colores corporativos (usar imágenes predeterminadas)
; WizardImageFile=compiler:wizmodernimage-is.bmp
; WizardSmallImageFile=compiler:wizmodernsmallimage-is.bmp

[Languages]
Name: "spanish"; MessagesFile: "compiler:Languages\Spanish.isl"
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked
Name: "quicklaunchicon"; Description: "{cm:CreateQuickLaunchIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked; OnlyBelowVersion: 6.1; Check: not IsAdminInstallMode

[Files]
; Ejecutable principal
Source: "dist\EqualityMomentum\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion
; Todos los archivos de la carpeta dist
Source: "dist\EqualityMomentum\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
; Archivos de configuración
Source: "config.json"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\version.json"; DestDir: "{app}"; Flags: ignoreversion
; Documentación
Source: "..\README.md"; DestDir: "{app}"; Flags: ignoreversion isreadme
; Isotipo
Source: "..\00_DOCUMENTACION\isotipo.jpg"; DestDir: "{app}\00_DOCUMENTACION"; Flags: ignoreversion

[Dirs]
; Crear carpetas de usuario en Documentos
Name: "{userdocs}\EqualityMomentum"; Permissions: users-full
Name: "{userdocs}\EqualityMomentum\Datos"; Permissions: users-full
Name: "{userdocs}\EqualityMomentum\Resultados"; Permissions: users-full
Name: "{userdocs}\EqualityMomentum\Informes"; Permissions: users-full
Name: "{userdocs}\EqualityMomentum\Logs"; Permissions: users-full

[Icons]
; Acceso directo en el menú inicio
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\Manual de Usuario"; Filename: "{app}\MANUAL_USUARIO.pdf"; Check: FileExists(ExpandConstant('{app}\MANUAL_USUARIO.pdf'))
Name: "{group}\Carpeta de Datos"; Filename: "{userdocs}\EqualityMomentum"
Name: "{group}\{cm:UninstallProgram,{#MyAppName}}"; Filename: "{uninstallexe}"
; Acceso directo en el escritorio
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon
; Acceso rápido
Name: "{userappdata}\Microsoft\Internet Explorer\Quick Launch\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: quicklaunchicon

[Run]
; Ejecutar la aplicación después de instalar (opcional)
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent

[Code]
// Código Pascal para personalización adicional

procedure CurPageChanged(CurPageID: Integer);
begin
  // Personalizar texto según la página
  if CurPageID = wpWelcome then
  begin
    WizardForm.WelcomeLabel2.Caption :=
      'Este asistente le guiará en la instalación de ' + ExpandConstant('{#MyAppName}') + ' ' +
      ExpandConstant('{#MyAppVersion}') + ' en su equipo.' + #13#10#13#10 +
      'EqualityMomentum es un sistema profesional de gestión de registros retributivos ' +
      'que le permite procesar datos y generar informes de manera sencilla.' + #13#10#13#10 +
      'Se recomienda cerrar todas las demás aplicaciones antes de continuar.';
  end;
end;

procedure CurStepChanged(CurStep: TSetupStep);
var
  ResultCode: Integer;
begin
  if CurStep = ssPostInstall then
  begin
    // Crear archivo de configuración inicial si no existe
    if not FileExists(ExpandConstant('{userdocs}\EqualityMomentum\config_user.json')) then
    begin
      SaveStringToFile(ExpandConstant('{userdocs}\EqualityMomentum\config_user.json'),
        '{ "first_run": true }', False);
    end;
  end;
end;

function InitializeSetup(): Boolean;
var
  ResultCode: Integer;
begin
  Result := True;

  // Verificar si hay una versión anterior instalada
  if RegKeyExists(HKEY_LOCAL_MACHINE, 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{#MyAppName}_is1') then
  begin
    if MsgBox('Se detectó una versión anterior de ' + ExpandConstant('{#MyAppName}') + '.' + #13#10#13#10 +
              '¿Desea continuar con la actualización?', mbConfirmation, MB_YESNO) = IDNO then
    begin
      Result := False;
    end;
  end;
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
var
  mRes: Integer;
begin
  case CurUninstallStep of
    usUninstall:
      begin
        mRes := MsgBox('¿Desea eliminar también los datos de usuario en Documentos?' + #13#10 +
                       '(Si selecciona NO, sus datos y configuración se conservarán)',
                       mbConfirmation, MB_YESNO or MB_DEFBUTTON2);
        if mRes = IDYES then
        begin
          // Eliminar carpeta de usuario
          DelTree(ExpandConstant('{userdocs}\EqualityMomentum'), True, True, True);
        end;
      end;
  end;
end;
