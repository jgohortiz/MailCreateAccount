#*****************************************************************************
# Nombre del Programa: MailCreateAccount
# Versión: 1.0
# Autor: José Guillermo Ortiz Hernández
# Fecha de creación: 2025-01-09
#
# Descripción:
# ----------------------------------------------------------------------------
# Este script tiene como objetivo crear las cuentas de correo a partir de
# lista de usuarios en un archivo (CSV)
#
# Uso:
# ----------------------------------------------------------------------------
# .\Main.ps1
#
# Requisitos:
# ----------------------------------------------------------------------------
# - Habilitar la ejecución de scripts
# 		Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
# 		Set-Executionpolicy unrestricted -scope CurrentUser
#
# Historial de cambios:
# Versión | Fecha       | Autor                          | Descripción
# 1.0     | 2025-01-06  | José Guillermo Ortiz Hernández | Versión inicial
#
# Este programa está protegido por las leyes de derechos de autor.
# Se permite el uso, modificación y distribución bajo los términos de
# licencia GPL.
#
# https://learn.microsoft.com/en-us/microsoft-365/enterprise/create-user-accounts-with-microsoft-365-powershell?view=o365-worldwide
# https://learn.microsoft.com/en-us/powershell/module/microsoft.graph.identity.signins/update-mguserauthenticationmethod?view=graph-powershell-1.0
# https://learn.microsoft.com/en-us/powershell/module/microsoft.graph.identity.signins/get-mguserauthenticationmethod?view=graph-powershell-1.0
# https://learn.microsoft.com/en-us/powershell/module/microsoft.graph.identity.signins/new-mguserauthenticationphonemethod?view=graph-powershell-1.0&viewFallbackFrom=graph-powershell-beta
# https://www.easy365manager.com/get-mguser-filter-example/
# https://www.nakivo.com/es/blog/how-to-connect-office-365-exchange-online-powershell/
# https://regex101.com/
#
# PS1 a EXE
# https://github.com/MScholtes/PS2EXE
# Install-Module -Name ps2exe -Force -Scope CurrentUser
# Invoke-ps2exe .\Main.ps1 .\MailCreateAccount.exe
#
#*****************************************************************************/
Add-Type -AssemblyName System.Windows.Forms


# Cargando funciones desde el modulo
Import-Module ./Modules/Functions.psm1
Test

# ---> INICIO ENTORNO
# Cargando módulos y conexiones
LoadConectModuleExchangeOnline
LoadConectModuleMSOnline
LoadConectModuleMSGraph

# Constantes
$START_PROCESS = Get-Date -Format "yyyyMMdd-HHmmss"
$PROCESS_LOTE_UUID = GetProcessLoteUUID -Time
$FILE_XAML_SOURCE = ".\Views\MainWindow.xaml"

$CONFIG_ACCOUNT = @{
	Dominio = "@domain.com"
	City = "CUNDINAMARCA"
	Country = "COLOMBIA"
	Department = "AREA"
	Office = "EMPRESA"
	PhoneNumber   = "(+57) 9999999999"
	PostalCode = "111711"
	State = "BOGOTA D.C."
	StreetAddress = "CALLE 99 No. 9-99"
	UsageLocation = "CO"
}
$CONFIG_ACCOUNT | Format-Table -AutoSize

$CONFIG_SCRIPT = @{
	Attempts = 10
	FileCsvSource = ".\Resources\Input\MailCreateAccounts.csv"
	FileCsvOutput = ".\Resources\Output\MailCreateAccount-$START_PROCESS.csv"
	FileLogOutput = ".\Resources\Output\MailCreateAccount-$START_PROCESS.log"
	FileXAMLSource = ".\Views\MainWindow.xaml"
}
$CONFIG_SCRIPT | Format-Table -AutoSize
# ---> FIN ENTORNO

# ---> INICIO CARGA DE FORMULARIO
$lcInputXML = Get-Content -Path $FILE_XAML_SOURCE -Raw
$lcInputXML = $lcInputXML -replace 'mc:Ignorable="d"','' -replace "x:N",'N' -replace '^<Win.*', '<Window'
[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
[xml]$XAML = $lcInputXML

#Leer y crear el render de XAML
$loReader=(New-Object System.Xml.XmlNodeReader $XAML)
try{
    $loForm=[Windows.Markup.XamlReader]::Load( $loReader )
} catch {
    Write-Warning "Unable to parse XML, with error: $($Error[0])`n Ensure that there are NO SelectionChanged or TextChanged properties in your textboxes (PowerShell cannot process them)"
    throw
}

# Cargando objetos
$XAML.SelectNodes("//*[@Name]") | %{try {Set-Variable -Name "WPF$($_.Name)" -Value $loForm.FindName($_.Name) -ErrorAction Stop} catch{throw} }
# ---> FIN CARGA DE FORMULARIO

 
# ---> INICIO CARGA DE MÉTODOS
$WPFtxtFileCsvSource.Add_TextChanged({ 
	$WPFgrdDatos.Itemssource = ""
	if (Test-Path $WPFtxtFileCsvSource.Text) {
		$WPFgrdDatos.Itemssource = Import-Csv -Path $WPFtxtFileCsvSource.Text
		if (-not ($WPFgrdDatos.Itemssource -eq $null)) {
			$WPFgrdDatos.Itemssource = $WPFgrdDatos.Itemssource
		}
	}
})
$WPFbtnFileCsvSource.Add_Click({ 
	$WPFtxtFileCsvSource.Text = GetSelectedFile -tcType "Archivos CSV (*.csv)|*.csv" 
})
$WPFbtnPathCsvOut.Add_Click({
	$WPFtxtPathCsvOut.Text = GetSelectedFolder -tcFile "MailCreateAccount-$CONFIG_SCRIPT.StartProcess.csv" 
})
$WPFbtnPathLogOut.Add_Click({
	$WPFtxtPathLogOut.Text = GetSelectedFolder -tcFile "MailCreateAccount-$CONFIG_SCRIPT.StartProcess.log"
})
$WPFbtnProcess.Add_Click({
	$WPFbtnProcess.IsEnabled = $false 
	
	# Tipo de correos a crear
	if ($WPFcmbTipo.SelectedIndex -eq 0){
		$lcTypeAccount = "e" # {e:= Estudiante, m:= Maestro}
		$lcAccountJobTitle = "ESTUDIANTE"
	} else {
		$lcTypeAccount = "m"
		$lcAccountJobTitle = "DOCENTE"
	}
	
	$lnRows = $WPFgrdDatos.Itemssource.Count
	$lnCols = ($WPFgrdDatos.Itemssource | Select-Object -First 1 | Get-Member -MemberType Properties).Count
	$lnRow = 1;
	
	if ($lnCols -eq 6) {
		if ($lnRows -gt 0) {	
			
			#ShowProgress -tnPercentComplete 0 -tcStatus "Iniciando" -tcCurrentOperation "Procesando el CSV"	-tnId 1 -toProgressBar $WPFpgbProgressBar
			$WPFstbiPrincipalState.Content = "Iniciando"
			$WPFstbiSecondState.Content = "Procesando el CSV" 
			$WPFpgbProgressBar.value = 0
			[System.Windows.Forms.Application]::DoEvents()

			# --> Inicio recorrido del CSV
			CsvOutput -tcFilePathLog $WPFtxtPathCsvOut.Text -tcRow "LOTE,CEDULA,NOMBRES,APELLIDOS,EMAIL PERSONAL,ALIAS,MOVIL,ESTADO,EMAIL INSTITUCIONAL,NOMBRE COMPLETO,TIPO,ID,CREADO,ESTADO CUENTA"
			foreach ($loRow in $WPFgrdDatos.Itemssource) {
				LogMessage -tcFilePathLog $WPFtxtPathLogOut.Text -tcMessage "# INICIO $lnRow / $lnRows"

				# Cargando las columnas de la fila al array laRow
				$laRow = @("cedula", "nombres", "apellidos", "mailpersonal", "celular", "observacion")
				$lnCol = 0

				foreach ($loCol in $loRow.PSObject.Properties) {
					if ($lnCol -lt $laRow.Length) {
						$laRow[$lnCol]=$($loCol.Value)
						$lnCol++
					}
				}

				# ----> INICIO CREACIÓN USUARIO
				LogMessage -tcFilePathLog $WPFtxtPathLogOut.Text -tcMessage "## INICIO PREPARACION ESCENARIO PARA LA CREACION DE LA CUENTA DE $($laRow[1]) $($laRow[2])"

				if(-not($laRow[0] -eq "cedula" -and $laRow[1] -eq "nombres")){

					# ----> PREPARANDO LA INFORMACIÓN PARA LA CREACIÓN USUARIO

					# Propiedades básicas
					$lcCustomAttribute1 = "$PROCESS_LOTE_UUID"
					$lcCustomAttribute2 = "$($laRow[0])"
					$lcCustomAttribute3 = "$($laRow[3])"
					$lcFirstName = "$($laRow[1])" -replace '\s+', ' '
					$lcLastName = "$($laRow[2])" -replace '\s+', ' '
					$lcPhoneMobileNumber = SetPhoneMobileNumber -tcPhoneMobileNumber "$($laRow[4])"
					$lcDisplayName = "$lcFirstName $lcLastName"
					$lcName = "$lcDisplayName"
					$lcPassword = (ConvertTo-SecureString -String "CC$lcCustomAttribute2." -AsPlainText -Force)
					$laAlternateMobiles = @("$lcPhoneMobileNumber ")
					$lcMail = GetAccountIdentyMail -tcTipo $lcTypeAccount -tcDisplayName $lcDisplayName -tcNombres $lcFirstName -tcApellidos $lcLastName -tcFilePathLog $WPFtxtPathLogOut.Text -tcDominio "$($CONFIG_ACCOUNT.Dominio)"
					$lcAlias = $lcMail.Split('@')[0]

					# Propiedades estáticas
					$laStrongAuthenticationRequirements = @()
					$lcCity = "$($CONFIG_ACCOUNT.City)"
					$lcCountry = "$($CONFIG_ACCOUNT.Country)"
					$lcDepartment = "$($CONFIG_ACCOUNT.Department)"
					$lcJobTitle = $lcAccountJobTitle
					$lcOffice = "$($CONFIG_ACCOUNT.Office)"
					$lcPhoneNumber = "$($CONFIG_ACCOUNT.PhoneNumber)"
					$lcPostalCode = "$($CONFIG_ACCOUNT.PostalCode)"
					$lcState = "$($CONFIG_ACCOUNT.State)"
					$lcStreetAddress = "$($CONFIG_ACCOUNT.StreetAddress)"
					$lcUsageLocation = "$($CONFIG_ACCOUNT.UsageLocation)"
					$SMS = New-Object -TypeName Microsoft.Online.Administration.StrongAuthenticationMethod
					$SMS.IsDefault = $true
					$SMS.MethodType = "OneWaySMS"
					$laStrongAuthenticationMethods = @($SMS)

					$lnProgress = ($lnRow/ $lnRows) * 100
					#ShowProgress -tnPercentComplete $lnProgress -tcStatus "Fila $lnRow de $lnRows" -tcCurrentOperation "$($lcDisplayName) - $($lcMail)" -tnId 1
					
					$WPFstbiPrincipalState.Content = "$($lcDisplayName) - $($lcMail)"
					$WPFstbiSecondState.Content = "Fila $lnRow de $lnRows" 
					$WPFpgbProgressBar.value = [int]$lnProgress
					[System.Windows.Forms.Application]::DoEvents()
					
					
					LogMessage -tcFilePathLog $WPFtxtPathLogOut.Text -tcMessage "## FIN PREPARACION ESCENARIO PARA LA CREACION DE LA CUENTA DE $($lcDisplayName) - $($lcMail)"
					# ----> FIN DE LA PREPARACION  LA INFORMACIÓN PARA LA CREACIÓN USUARIO

					# ---> Inicio creación/actualización de cuenta, propiedades y método SSPR
					LogMessage -tcFilePathLog $WPFtxtPathLogOut.Text -tcMessage "## INICIO DEL PROCESO DEL REGISTRO DE $($lcDisplayName) - CUENTA A CREAR $($lcMail)"
					if(-not($lcMail -eq "")){

						# Creación o Actualización de la cuenta
						$llCreateAccount = $False
						$llUpdateAccount = $False
						$lcDisplayNameAux = getAccountIdentyName -tcEmail $lcMail -tcFilePathLog $tcFilePathLog
						if($lcDisplayNameAux -eq ""){
							LogMessage -tcFilePathLog $WPFtxtPathLogOut.Text -tcMessage "Crear la cuenta $lcMail para $lcDisplayName"
							$llCreateAccount = $True
						} elseif ($lcDisplayNameAux -eq $lcDisplayName){
							LogMessage -tcFilePathLog $WPFtxtPathLogOut.Text -tcMessage "Actualizar la cuenta $lcMail para $lcDisplayName"
							$llUpdateAccount = $True
						}

						# Procesando la creación y actualización
						if($llCreateAccount -eq $True -or $llUpdateAccount -eq $True){

							# Creación de la cuenta
							if($llCreateAccount -eq $True){
								# Validando si el nombre esta disponible
								$llDisplayNameAvailability = ChekDisplayNameAvailability -$tcGivenName "$lcFirstName" -tcSurname "$lcLastName" -tcDisplayName "$lcDisplayName"
								if($llDisplayNameAvailability -eq $True) {
									try{
										LogMessage -tcFilePathLog $WPFtxtPathLogOut.Text -tcMessage "Se creara la cuenta $lcMail para $lcDisplayName"
										New-Mailbox -Name "$lcName" -Password $lcPassword -MicrosoftOnlineServicesID "$lcMail" -Alias "$lcAlias" -FirstName "$lcFirstName" -LastName "$lcLastName" -DisplayName "$lcDisplayName" -ResetPasswordOnNextLogon $True -ErrorAction Stop #-ErrorAction SilentlyContinue
									} catch {
										LogMessage -tcFilePathLog $WPFtxtPathLogOut.Text -tcMessage "No se pudo crear la cuena $lcMail para $lcDisplayName, error $_.ScriptStackTrace"
									}
										
								} else {
									Write-Host "*** El nombre de usuario $lcDisplayName ya esta registrado "
									LogMessage -tcFilePathLog $WPFtxtPathLogOut.Text -tcMessage "*** El nombre de usuario $lcDisplayName ya esta registrado "
								}
							}
							
							# Esperando a que se cree el usuario
							$loUser = $null
							$lnCheck = 0
							while (($loUser -eq $null) -and ($lnCheck -lt $CONFIG_SCRIPT.Attempts)) {
								$loUser = Get-MgUser -All -Filter "UserPrincipalName eq '$lcMail'"
								Start-Sleep -Seconds 5
								$lnCheck++
								Write-Host "Buscando a $lcMail ... Intento $lnCheck"
							}
							$loUser = $null
							
							# Verificando la existencia
							$lcDisplayNameAux = getAccountIdentyName -tcEmail $lcMail -tcFilePathLog $tcFilePathLog
							if(-not($lcDisplayNameAux -eq "")){
								if($lcDisplayNameAux -eq $lcDisplayName){

									# Obteniendo el usuario
									try{
										LogMessage -tcFilePathLog $WPFtxtPathLogOut.Text -tcMessage "Se actualizara la cuenta $lcMail para $lcDisplayName"
										$loUser = Get-MgUser -All -Filter "UserPrincipalName eq '$lcMail'"
									} catch {
										LogMessage -tcFilePathLog $WPFtxtPathLogOut.Text -tcMessage "No se pudo obtener el usuario de $lcMail para $lcDisplayName"
									}

									# Asegúrate de que el usuario existe
									if (-not ($loUser -eq $null)) {
										if ($loUser.DisplayName -eq $lcDisplayName -and $loUser.mail -eq $lcMail) {

											# Asignando licencia
											LogMessage -tcFilePathLog $WPFtxtPathLogOut.Text -tcMessage "Asignando la licencia $lcLicenses a $lcMail asignada a $lcDisplayNameAux"
											$lcLicense = "STANDARDWOFFPACK_FACULTY"
											$loLicense = Get-MgSubscribedSku | Where-Object { $_.SkuPartNumber -eq $lcLicense }
											Update-MgUser -UserId $loUser.Id -JobTitle "$lcJobTitle" -UsageLocation "$lcUsageLocation" -Country "$lcCountry" | Out-Null
											Set-MgUserLicense -UserId $loUser.Id -AddLicenses @{SkuId = ($loLicense.SkuId) } -RemoveLicenses @() | Out-Null

											# Establecer propiedades
											LogMessage -tcFilePathLog $WPFtxtPathLogOut.Text -tcMessage "Estableciendo propiedades a $lcMail asignada a $lcDisplayNameAux"
											Set-MsolUser -UserPrincipalName "$lcMail" -Country "$lcCountry" -City "$lcCity" -Office "$lcOffice" -PhoneNumber "$lcPhoneNumber" -State "$lcState" -StreetAddress "$lcStreetAddress" -Title "$lcJobTitle" -UsageLocation "$lcUsageLocation" -PostalCode "$lcPostalCode" -Department "$lcDepartment" -MobilePhone "$lcPhoneMobileNumber" -AlternateMobilePhones $laAlternateMobiles -StrongAuthenticationMethods $laStrongAuthenticationMethods -StrongAuthenticationRequirements $laStrongAuthenticationRequirements | Out-Null
											Set-Mailbox -Identity "$lcMail" -CustomAttribute1 $lcCustomAttribute1 -CustomAttribute2 $lcCustomAttribute2 -CustomAttribute3 $lcCustomAttribute3

											# Métodos de autenticación para realizan el autoservicio de restablecimiento de contraseña (SSPR)
											$loAuthenticationMethod = Get-MgUserAuthenticationMethod -UserId "$lcMail" | Where-Object {$_.AdditionalProperties.phoneType -eq 'mobile' }
											if (-not ($loAuthenticationMethod -eq $null)) {
												if (-not ($loAuthenticationMethod.id -eq $null)) {
													LogMessage -tcFilePathLog $WPFtxtPathLogOut.Text -tcMessage "Actualizado  el metodo de autenticacion para realizan el autoservicio de restablecimiento de clave (SSPR) a $lcMail asignada a $lcDisplayNameAux"
													Update-MgUserAuthenticationPhoneMethod -PhoneAuthenticationMethodId $loAuthenticationMethod.id -UserId "$lcMail" -PhoneNumber "$lcPhoneMobileNumber" -PhoneType "mobile" -ErrorAction SilentlyContinue | Out-Null
												} else {
													LogMessage -tcFilePathLog $WPFtxtPathLogOut.Text -tcMessage "Creado  el metodo de autenticación para realizan el autoservicio de restablecimiento de contraseña (SSPR) a $lcMail asignada a $lcDisplayNameAux"
													New-MgUserAuthenticationPhoneMethod -PhoneAuthenticationMethodId $loAuthenticationMethod.id -UserId "$lcMail" -PhoneNumber "$lcPhoneMobileNumber" -PhoneType "mobile" -ErrorAction SilentlyContinue | Out-Null
												}
											}
										} else {
											LogMessage -tcFilePathLog $WPFtxtPathLogOut.Text -tcMessage "*** No se puede procesar la cuenta $lcMail, no hay concordancia, DisplayName y Mail"
										}
									} else {
										LogMessage -tcFilePathLog $WPFtxtPathLogOut.Text -tcMessage "*** El identificador de usuario para la cuenta $lcMail, No existe"
									}
								} else {
									LogMessage -tcFilePathLog $WPFtxtPathLogOut.Text -tcMessage "*** No se puede procesar la cuenta $lcMail, no hay concordancia, asignada $lcDisplayNameAux esperada $lcDisplayName"
								}
							} else {
								LogMessage -tcFilePathLog $WPFtxtPathLogOut.Text -tcMessage "*** No se puede procesar la cuenta $lcMail, no fue creada"
							}
						} else {
							LogMessage -tcFilePathLog $WPFtxtPathLogOut.Text -tcMessage "*** No se puede crear la cuenta $lcMail ya que esta en uso por $lcDisplayNameAux"
						}

					} else {
						LogMessage -tcFilePathLog $WPFtxtPathLogOut.Text -tcMessage "*** No se obtuvo mail para $lcDisplayName"
					}
					# ---> Fin creación/actualización de cuenta, propiedades y método SSPR

					# ---> Escribiendo en el CSV de proceso
					#cedula,nombres,apellidos,mailpersonal,celular,alias,creado,email,Nombre,cargo,ID
					$lcRow = "$lcCustomAttribute1,$lcCustomAttribute2,$lcFirstName,$lcLastName,$lcCustomAttribute3,$lcAlias"
					$lcRowAccountnfo = ""

					$lcDisplayNameAux = getAccountIdentyName -tcEmail $lcMail -tcFilePathLog $tcFilePathLog
					if(-not ($lcDisplayNameAux -eq "")){
						if($lcDisplayNameAux -eq $lcDisplayName){
							$loUser = Get-MgUser -All -Property UserPrincipalName,DisplayName,UserType,JobTitle,Id,CreatedDateTime,LastModifiedDateTime,AccountEnabled,Mail -Filter "UserPrincipalName eq '$lcMail'"
							if (-not ($loUser -eq $null)) {
								if ($loUser.DisplayName -eq $lcDisplayName -and $loUser.mail -eq $lcMail) {
									$lcRowAccountnfo = "CREADO"
								} else {
									$lcRowAccountnfo = "OMITIDO-EXISTE"
								}
								$lcRowAccountnfo = "$lcRowAccountnfo,$($loUser.UserPrincipalName),$($loUser.DisplayName),$($loUser.JobTitle),$($loUser.Id),$($loUser.CreatedDateTime),activa-$($loUser.AccountEnabled)"
							} else {
								$lcRowAccountnfo = "CREADO-INACCESIBLE"
							}
						} else {
							$loUser = Get-MgUser -All -Property UserPrincipalName,DisplayName,UserType,JobTitle,Id,CreatedDateTime,LastModifiedDateTime,AccountEnabled,Mail -Filter "UserPrincipalName eq '$lcMail'"
							if (-not ($loUser -eq $null)) {
								$lcRowAccountnfo = "OMITIDO-EXISTE"
								$lcRowAccountnfo = "$lcRowAccountnfo,$($loUser.UserPrincipalName),$($loUser.DisplayName),$($loUser.JobTitle),$($loUser.Id),$($loUser.CreatedDateTime),activa-$($loUser.AccountEnabled)"
							} else {
								$lcRowAccountnfo = "OMITIDO-INACCESIBLE"
							}
						}
					} else {
						$lcRowAccountnfo = "OMITIDO"
					}
					$lcRow = "$($lcRow),$lcPhoneMobileNumber,$lcRowAccountnfo"
					LogMessage -tcFilePathLog $WPFtxtPathLogOut.Text -tcMessage "Escribiendo en el CSV $lcRow"
					Write-Host $lcRow -ForegroundColor Black -BackgroundColor DarkGray
					CsvOutput -tcFilePathLog $WPFtxtPathCsvOut.Text -tcRow $lcRow
					# ---> Fin escribiendo en el CSV de proceso

					LogMessage -tcFilePathLog $WPFtxtPathLogOut.Text -tcMessage "## FIN DEL PROCESO DEL REGISTRO DE $($lcDisplayName) - CUENTA A CREAR $($lcMail)"
				} else {
					LogMessage -tcFilePathLog $WPFtxtPathLogOut.Text -tcMessage "* Cabecera de CSV"
				}
				# ----> FIN CREACIÓN USUARIO

				LogMessage -tcFilePathLog $WPFtxtPathLogOut.Text -tcMessage "# FIN $lnRow / $lnRows"
				LogMessage -tcFilePathLog $WPFtxtPathLogOut.Text -tcMessage "*"
				LogMessage -tcFilePathLog $WPFtxtPathLogOut.Text -tcMessage "*"

				$lnRow++

			}
			# --> Fin recorrido del CSV
			#ShowProgress -tnPercentComplete 100 -tcStatus "Completado" -tcCurrentOperation "Procesado" -tnId 1
			$WPFstbiPrincipalState.Content = "Completado"
			$WPFstbiSecondState.Content = "Procesado"
			$WPFpgbProgressBar.value = 100
			
			[System.Windows.MessageBox]::Show((New-Object System.Windows.Window -Property @{TopMost = $True}),"Proceso terminado","Terminado",0,[System.Windows.MessageBoxImage]::Information)
	
		} else {
			[System.Windows.MessageBox]::Show((New-Object System.Windows.Window -Property @{TopMost = $True}),"El archivo $($WPFtxtFileCsvSource.Text) no contiene información","Error",0,[System.Windows.MessageBoxImage]::Error)
		}
	} else {
		[System.Windows.MessageBox]::Show((New-Object System.Windows.Window -Property @{TopMost = $True}),"El archivo $($WPFtxtFileCsvSource.Text) no contiene las columnas requeridas","Error",0,[System.Windows.MessageBoxImage]::Error)
		
	}	
	$WPFbtnProcess.IsEnabled = $True
})
# ---> FIN CARGA DE MÉTODOS

	
# ---> INICIO ESTABLECIMIENTO PROPIEDADES
$WPFtxtFileCsvSource.Text = "$($CONFIG_SCRIPT.FileCsvSource)"
$WPFtxtPathCsvOut.Text = "$($CONFIG_SCRIPT.FileCsvOutput)"
$WPFtxtPathLogOut.Text = "$($CONFIG_SCRIPT.FileLogOutput)"
$WPFtxtUUID.Text = "$PROCESS_LOTE_UUID"
# ---> FIN ESTABLECIMIENTO PROPIEDADES

# Mostrado el formulario
$loForm.ShowDialog() | out-null

# Cerrando conexiones
#Disconnect-ExchangeOnline -Confirm:$false
