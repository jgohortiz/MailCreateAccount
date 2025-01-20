# MailCreateAccount
Este script tiene como objetivo crear las cuentas de correo en Microsoft 365 a partir de un archivo CSV

# Comentarios

Esta es una demostración sencilla de como se pueden crear cuentas de correo en Microsfot 365 mediante el uso de Powershell:
- En este script se crea una GUI mediante la definición de una vista con xaml
- Se usan funciones definidas en un psm1
- Se utlizan elementos de ExchangeOnline, MSOnline y Microsoft.Graph


![Screenshot](https://github.com/jgohortiz/MailCreateAccount/blob/main/preview.png)

## Instalación
 - Habilitar la ejecución de scripts
 		Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
 		Set-Executionpolicy unrestricted -scope CurrentUser

## Configuración
Configuration in `src/app/agentConfigs/simpleExample.ts`
```powershell
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

$CONFIG_SCRIPT = @{
	Attempts = 10
	FileCsvSource = ".\Resources\Input\MailCreateAccounts.csv"
	FileCsvOutput = ".\Resources\Output\MailCreateAccount-$START_PROCESS.csv"
	FileLogOutput = ".\Resources\Output\MailCreateAccount-$START_PROCESS.log"
	FileXAMLSource = ".\Views\MainWindow.xaml"
}
```

### Próximos pasos
Ejecute el script 
PS C:\Scripts\MailCreateAccount> .\Main.ps1

### Definiendo tus propios modelos
En la función GetAccountIdentyMail del modulo Functions.psm1 existen métodos propuestos para la asignación de la cuenta.
