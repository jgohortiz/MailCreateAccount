#*****************************************************************************
# Nombre del Programa: Functions.ps1
# Versión: 1.0
# Autor: José Guillermo Ortiz Hernández
# Fecha de creación: 2025-01-09
#
# Descripción:
# ----------------------------------------------------------------------------
# Este script contiene las funciones de uso común
#
# Historial de cambios:
# Versión | Fecha       | Autor                          | Descripción
# 1.0     | 2025-01-06  | José Guillermo Ortiz Hernández | Versión inicial
#
# Este programa está protegido por las leyes de derechos de autor.
# Se permite el uso, modificación y distribución bajo los términos de
# licencia GPL.
#
#*****************************************************************************/

# Función de prueba
function Test {
	Clear-Host
	$laAsciiArt = @(" .....                                                ........                                                    "," ..*%#=:....                                      ...:=*%#:...                                                    ","  ..#%%%%#*=...                                ..-*#%%%%%=..                                                      ","   .+%%%%%##%%*-..                        ...:+#%###%%%%%...                                                      ","   .=%%%%%%=.=#%%#=..                   ..:##%#+:-#%%%%%#...                                                      ","   .+%%%%%%%:..:#%%%+..               ..-#%%#=...#%%%%%%#...                                                      ","   .+%%%%%%%#....+%%%#-...           .:*%%%*....+%%%%%%%%...                                                      ","   .*%%%%%%%%*....=##%%-...        ..:#%%#*....=%%%%%%%%%...                                                      ","   .*%%%%%%%%%+....+++%#-..        ..%%#-*....-%%%%%%%%%%...                                                      ","   .=#.-#%%%%%%#:...#:*%*..        .+%#:=:...*%%%%%%#+.**...                                                      ","   .:#-..*%%%%%*.....-+%#:.        .*%*.=....=%%%%%#:.:#-..                                                       ","    .+#...-#%%%%+.....+%*.        ..+%*.....=#%%%#+...+*...                                                       ","    .:#*....#%%%%=...:##-.        ...#%:...:#%%%%-...=#-...                                                       ","     .-%-....*%%%=....-:...       ....=....:#%%#:...:#*..                                                         ","     ..##%=..+%%*..........       ..........+%%#..-##%:..                                                         ","      .+%%%*##*-..                         ..:*#%*#%%%...                                                         ","      .*%%%*...                              ....=#%%%..                                                          ","     .:%%%%*...                               ...=#%%%=....                                                       ","    ..#%%#:....                               .....*%%#:...                                                       ","    .+#:-......   ...   ..           . ..     ......::*#...                                                       ","    :#......     ..::..+:..        ..+-.:++-...  ......*=..                                                       ","   ....#%=..   ...:-+#%%#+...      :*%%%#=:....  ..:##-.....                                                      ","   ...#*:...  ........-#%%+..     .*%%*..........%%%=.#:....               ....... .--:                 --- .---  ","    .....    ..:*%+-:....:....   ..--....-*%#=..#%%%%-.....               -#%%%%%= :#%+                 #%# .#%*  ","                ..*%%%=...         .. ..#%%+...-%%=   ....... ...... ....:%%*   .: :#%+.....    .....   #%# .#%*  ","                  .:*%%-..           ..#%+....+%%%%%=.+#%%%%#*..*%#:.*%*:.#%%*=.   :#%##%%%#. :*%###%+. #%# .#%*  ","                   ..:#*...          .-#:...   -%%=  +%%:  .#%* .+%%*%*.  .=##%%#- :#%*  =%%-.*%*:.:##= #%# .#%*  ","                   ....:........... ........   -%%=  *%#.   #%#  :#%%#:     ..:#%%::#%+  -%%-.*%#+++++- #%# .#%*  ","                      .......:==:..  .=...     -%%=  :#%#-:#%#-.=%#==%%=.:#*--*%%#.:#%+  -%%-.+%#=..:=  #%# .#%*  ","                      ..*:.-#%%%%%+..*+...     :++:   .-+##+-..-++-..-++-.=+##*=:. .++-  :++: .:=*#*+-  ++= .++=  ","                       ..+..:*#%#-..+=....                                                                        ","                       ...-+**#%##*-...                                                                           ","                           ..........                                                                             ")
	$laAsciiArt | ForEach-Object { Write-Host $_ -ForegroundColor Green }
}

function LoadConectModuleExchangeOnline {
	Write-Host "Cargando modulo y conexion ExchangeOnline, espere ..."
	Install-Module -Name ExchangeOnlineManagement
	Import-Module -Name ExchangeOnlineManagement

	try {
		Get-Mailbox -ResultSize 1 | Out-Null
	} catch {
		Connect-ExchangeOnline -ShowProgress $true
	}
}

function LoadConectModuleMSOnline {
	Write-Host "Cargando modulo y conexion MSOnline, espere ..."
	Install-Module MSOnline
	Import-Module -Name MSOnline

	try {
		Get-MsolCompanyInformation -ErrorAction Stop -WarningAction SilentlyContinue | Out-Null
	} catch {
		Connect-MsolService
	}
}

function LoadConectModuleMSGraph {
	Write-Host "Cargando modulo y conexion Microsoft.Graph, espere ..."
	Install-Module Microsoft.Graph
	Connect-MgGraph -Scopes "User.ReadWrite.All", "Organization.Read.All", "Directory.AccessAsUser.All", "UserAuthenticationMethod.ReadWrite.All" -NoWelcome
}

function SetPhoneMobileNumber {
    param (
        [string]$tcPhoneMobileNumber
    )
	$lcPhoneMobileNumber = "$tcPhoneMobileNumber"
	if ($lcPhoneMobileNumber.Length -gt 10) {
		$lcPhoneMobileNumber = $lcPhoneMobileNumber.Substring(0, 10)
	}
	$lcPhoneMobileNumber = "+57 $lcPhoneMobileNumber"

	return $lcPhoneMobileNumber
}

# Función para genera un ID del proceso por lote
function GetProcessLoteUUID {
    param (
        [string]$tcText,
		[switch]$Replace,
        [switch]$ToLower,
		[switch]$ToUpper,
        [switch]$Time
    )

    $lcUUID = [guid]::NewGuid().ToString()
	if ($Time) {
		$lcTimeStamp = Get-Date -Format "yyMMdd:HHmmss"
	} else {
		$lcTimeStamp = $lcUUID.Substring(0, 13)
	}
	$lcUUID  = $lcUUID.Substring(14, 9)

    return "LT-$lcTimeStamp-$lcUUID"
}

# Función para escribir en consola y en el log
function LogMessage {
    param (
        [string]$tcFilePathLog,
		[string]$tcMessage
    )

    $lcTimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $lclogMessage = "$lcTimeStamp - $tcMessage"

    # Escribir en el archivo de log
    $lclogMessage | Out-File -FilePath $tcFilePathLog -Append
}

# Función para escribir en CSV
function CsvOutput {
    param (
        [string]$tcFilePathLog,
		[string]$tcRow
    )

    # Escribir en el archivo de log
    $tcRow | Out-File -FilePath $tcFilePathLog -Append
}

# Función para validar si una cuenta existe en ExchangeOnline
function getAccountIdentyName {
    param (
        [string]$tcEmail
    )

	$lcDisplayName = ""
    try {
        $loMailbox = Get-Mailbox -Identity $tcEmail -ErrorAction Stop
        $lcDisplayName = $loMailbox.DisplayName
    } catch {
        $lcDisplayName = ""
    }
	$lcDisplayName = $($lcDisplayName).Trim().ToUpper()

	return $lcDisplayName;
}

# Función para validar si una cuenta existe en ExchangeOnline
function ChekAccountExchangeOnlineExist {
    param (
        [string]$tcEmail,
		[string]$tcFilePathLog
    )

    try {

        $loMailbox = Get-Mailbox -Identity $tcEmail -ErrorAction Stop
		LogMessage -tcFilePathLog $tcFilePathLog -tcMessage "El email $tcEmail esta en uso por $($loMailbox.DisplayName)"
        return $true
    } catch {
        return $false
    }
}

# Función para validar si el nombre de usuario está en uso
function ChekDisplayNameAvailability {
    param (
        [string]$tcGivenName,
		[string]$tcSurname,
		[string]$tcDisplayName
    )
	$llAvailability = $True
    $loUser = Get-MgUser -All | Where-Object {($_.DisplayName -eq '$tcDisplayName') -or ($_.GivenName -like '$tcGivenName*' -and $_.Surname -like '$tcSurname*')}
	if (-not ($loUser -eq $null)) {
		$tcDisplayName = $tcDisplayName.Trim() -replace '\s+', ' '
		$tcDisplayNameAux = $loUser.DisplayName.Trim() -replace '\s+', ' '

		if (-not ($tcDisplayNameAux -eq $lcDisplayName)) {
			$llAvailability = $False
		}
	}

	return $llAvailability;
}

#Funciona para pedirle al usuario si desea continuar
function CheckContinue {
    param (
        [switch]$Exit
    )	
	$lcConfirmacion = ""
	$llContinue = $False
	
	while ($lcConfirmacion -notmatch "^S|N$"){
		$lcConfirmacion = Read-Host -Prompt "Continuar con ejecucion del script?(S/N)"
	}
	$lcConfirmacion = $lcConfirmacion.Trim().ToUpper()
	if ($lcConfirmacion -eq "S") {
		$llContinue = $True
	}else{
		Write-Host "Ejecucion cancelada por el usuario."
		if ($Exit){
			Exit
		}
	}
	
	return $llContinue
}

# Función para darle formato puntual a un string
# https://www.ascii-code.com/es
function GetTextWithFormat {
    param (
        [string]$tcText,
		[switch]$Replace,
        [switch]$ToLower,
		[switch]$ToUpper,
        [switch]$Trim
    )

	if ($Replace) {
		$replacements = @(
			@{Pattern = @([char]0xF1); Replacement = 'n'},
			@{Pattern = @([char]0xE0;[char]0xE1;[char]0xE2;[char]0xE3;[char]0xE4;[char]0xE5); Replacement = 'a'},
			@{Pattern = @([char]0xE8;[char]0xE9;[char]0xEA;[char]0xEB); Replacement = 'e'},
			@{Pattern = @([char]0xEC;[char]0xED;[char]0xEE;[char]0xEF); Replacement = 'i'},
			@{Pattern = @([char]0xF2;[char]0xF3;[char]0xF4;[char]0xF5;[char]0xF6); Replacement = 'o'},
			@{Pattern = @([char]0xF9;[char]0xFA;[char]0xFB;[char]0xFC); Replacement = 'u'}
			@{Pattern = @([char]0xD1); Replacement = 'N'}
			@{Pattern = @([char]0xC0;[char]0xC1;[char]0xC2;[char]0xC3;[char]0xC4;[char]0xC5); Replacement = 'A'},
			@{Pattern = @([char]0xC8;[char]0xC9;[char]0xCA;[char]0xCB); Replacement = 'E'},
			@{Pattern = @([char]0xCC;[char]0xCD;[char]0xCE;[char]0xCF); Replacement = 'I'},
			@{Pattern = @([char]0xD3;[char]0xD4;[char]0xD5;[char]0xD6); Replacement = 'O'},
			@{Pattern = @([char]0xD9;[char]0xDA;[char]0xDB;[char]0xDC); Replacement = 'U'}
		)
		foreach ($replacement in $replacements) {
			foreach ($character in $replacement.Pattern) {
				$tcText = $tcText -replace $character, $replacement.Replacement
			}
		}
	}
	if ($Clean) {
		$tcText = $tcText -replace '[^a-zA-Z]', ''
	}
    if ($Trim) {
        $tcText = $tcText.Trim()
    }

    if ($ToUpper) {
        $tcText = $tcText.ToUpper()
    }

    if ($ToLower) {
        $tcText = $tcText.ToLower()
    }

    return $tcText
}

#Función para generar el la cuenta de e-mail
function GetAccountIdentyMail {
    param (
        [string]$tcTipo,
		[string]$tcDisplayName,
		[string]$tcNombres,
        [string]$tcApellidos,
		[string]$tcFilePathLog,
		[string]$tcDominio
    )

	
	$tcTipo = GetTextWithFormat -tcText $tcTipo -Clean -Trim -ToLower
	$tcDisplayName = GetTextWithFormat -tcText $tcDisplayName -Trim -ToUpper
	$tcNombres = GetTextWithFormat -tcText $tcNombres -Replace -Clean -Trim -ToLower
	$tcApellidos = GetTextWithFormat -tcText $tcApellidos -Replace -Clean -Trim -ToLower
	$lnNombres = ($tcNombres -split ' ').Count
	$lnApellidos = ($tcApellidos -split ' ').Count
	$lcEmail = ""

	# Para estudiantes
	if($tcTipo -eq "e"){
		# Primera opción de nombre
		# --> Método 1
		$tcPrimerNombre = $tcNombres.Split(' ')[0]
		$tcLetraPrimerNombre = $tcPrimerNombre.Substring(0, 1)
		$tcPrimerApellido = $tcApellidos.Split(' ')[0]
		$lcEmail = "$tcLetraPrimerNombre$tcPrimerApellido$tcDominio"
		# -- Fin método 1

		#Segunda opción de nombre (En caso de estar en uso la anterior)
		$llCheckNext = $True
		$llAccountExist = ChekAccountExchangeOnlineExist -tcEmail $lcEmail -tcFilePathLog $tcFilePathLog
		if($llAccountExist -eq $True){
			$lcDisplayName = getAccountIdentyName -tcEmail $lcEmail
			if(-not($tcDisplayName -eq $lcDisplayName)){
				if($lnNombres -ge 2){
					# --> Método 2
					LogMessage -tcFilePathLog $tcFilePathLog -tcMessage "Generando mail metodo alternativo 2 para $tcNombres $tcApellidos"
					$tcPrimerNombre = $tcNombres.Split(' ')[0]
					$tcSegundoNombre = $tcNombres.Split(' ')[1]
					$tcLetraPrimerNombre = $tcPrimerNombre.Substring(0, 1)
					$tcLetraSegundoNombre = $tcSegundoNombre.Substring(0, 1)
					$tcPrimerApellido = $tcApellidos.Split(' ')[0]
					$lcEmail = "$tcLetraPrimerNombre$tcLetraSegundoNombre$tcPrimerApellido$tcDominio"
					LogMessage -tcFilePathLog $tcFilePathLog -tcMessage "Se genero el mail $lcEmail para $tcNombres $tcApellidos"
					# -- Fin método 2
				}
			} else {
				$llCheckNext = $False
			}
		}

		#Tercera opción de nombre (En caso de estar en uso la anterior)
		if($llCheckNext -eq $True){
			$llAccountExist = ChekAccountExchangeOnlineExist -tcEmail $lcEmail -tcFilePathLog $tcFilePathLog
			if($llAccountExist -eq $True){
				$lcDisplayName = getAccountIdentyName -tcEmail $lcEmail
				if(-not($tcDisplayName -eq $lcDisplayName)){
					if($lnApellidos -ge 2){
						# --> Método 3
						LogMessage -tcFilePathLog $tcFilePathLog -tcMessage "Generando mail metodo alternativo 3 para $tcNombres $tcApellidos"
						$tcPrimerNombre = $tcNombres.Split(' ')[0]
						$tcLetraPrimerNombre = $tcPrimerNombre.Substring(0, 1)
						$tcPrimerApellido = $tcApellidos.Split(' ')[0]
						$tcSegundoApellido = $tcApellidos.Split(' ')[1]
						$tcLetraSegundoApellido = $tcSegundoApellido.Substring(0, 1)
						$lcEmail = "$tcLetraPrimerNombre$tcPrimerApellido$tcLetraSegundoApellido$tcDominio"
						LogMessage -tcFilePathLog $tcFilePathLog -tcMessage "Se genero el mail $lcEmail para $tcNombres $tcApellidos"
						# -- Fin método 3
					}
				}
			} else {
				$llCheckNext = $False
			}
		}
	}
	LogMessage -tcFilePathLog $tcFilePathLog -tcMessage "Se genero para $tcDisplayName de tipo $tcTipo el nombre mail $lcEmail"

    return $lcEmail
}

function GetSelectedFile{
    param (
        [string]$tcType
    )

    Add-Type -AssemblyName System.Windows.Forms
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "$tcType"
    $openFileDialog.Title = "Seleccionar"
    $openFileDialog.Multiselect = $false

    if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return "$($openFileDialog.FileName)"
    } else {
        return ""
    }
}

function GetSelectedFolder {
    param (
        [string]$tcFile
    )
	
    Add-Type -AssemblyName System.Windows.Forms
    $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderDialog.Description = "Seleccione una carpeta"
    $dialogResult = $folderDialog.ShowDialog()

    if ($dialogResult -eq [System.Windows.Forms.DialogResult]::OK) {
		$tcFile = GetTextWithFormat -tcText $tcFile -Clean -Trim -ToLower
        return "$($folderDialog.SelectedPath)\$tcFile"
    }
    else {
        return ""
    }
}