<#
.SYNOPSIS
  Create Exchange MailBox account
.DESCRIPTION
  Creates a Exchange MailBox from txt information user.
.OUTPUTS
  none
.NOTES
  Version:        1.0
  Author:         Meest
  Creation Date:  05 october 2022
  Purpose/Change: Initial script development
#>

# Function to create a new 16 characters random password.
function NewCredencial {
        Add-Type -AssemblyName 'System.Web'
        return [System.Web.Security.Membership]::GeneratePassword(16, 1)
}

# Function to check if OU exist or not.
function OuExist{

        $Path = "OU=$tenant,OU=Tenants,DC=lab,DC=local" # SPECIFIES YOUR PATH

        $ou_exists = [adsi]::Exists("LDAP://$Path")

        if (-not $ou_exists) { # If not exist throw error and go arround to addOneUser function.
                Write-Host('[!] Supplied Path does not exist. Try Again!')
                pause
                Clear-Host
                addOneUser
        } else { # If exists write in debug mode that this path exists.
                Write-Debug "Path Exists: $Path"
        }
}

# Function to check if user exist or not in OU.
function checkUser{
        $ErrorActionPreference = "SilentlyContinue" # Silent errors from cmdlets

        $Path = "OU=$tenant,OU=Tenants,DC=lab,DC=local" # SPECIFIES YOUR PATH
        
        $ADUser = Get-ADUser -Filter {UserPrincipalName -eq $email} -SearchBase $Path -SearchScope OneLevel #Search if user by UserPrincipalName exists.
       
        if (-not $ADUser) { # If not exists write on debug mode that this user not exists.
                Write-Debug ('[+] Supplied User does not exist.')
        } else { # If exist thow an alert and go arround to addOneUser function.
                Write-Host ('[!] User Exists, try again!')
                pause
                Clear-Host
                addOneUser
        }
        
}
# Function to add new User if OU exist and if user don't exist into OU.
function userAdd {
        Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn
        $c = NewCredencial # Get randomized credential.
        $secureString = ConvertTo-SecureString $c -AsPlainText -Force # Convert credential to a Secure String.

        # Try to create user, if fails throw an error to exit.
        try{
                New-Mailbox -Name "$name $surname | $tenant" -Alias "$tenant_$name$surname" -OrganizationalUnit "lab.local/Tenants/$tenant" -UserPrincipalName "$email" -SamAccountName "$name$surname" -FirstName "$name" -LastName "$surname" -Password $secureString -ResetPasswordOnNextLogon $false
                Set-Mailbox "$email" -CustomAttribute1 "$tenant"
        } catch {
                Throw("[!] ERROR Creando el usuario")
        }
        Clear-Host
        # Write to screen the resume information for the new user.
        Write-Host "================ Resumen de la operacion ================"
        Write-Host "Nombre: $name"
        Write-Host "Apellido: $surname"
        Write-Host "Email: $email"
        Write-Host "Password: $c"
        pause
        
}
# Function that user write the information for the new Mailbox
function addOneUser{
        Clear-host
        Write-Host "================ Ha entrado en el modo de creacion de buzones individuales ================"
        $name= Read-Host "[+] Especifique el nombre del usuario"
        if(!($name -eq '')){
                $surname = Read-Host "[+]Especifique el apellido del usuario"
                if(!($surname -eq '')){
                        $email = Read-Host "[+]Especifique el email del usuario"
                        if(!($email -eq '')){
                                $tenant = Read-Host "[+] Especifique el cliente"
                                OuExist($tenant)
                                checkUser($email,$tenant)
                                userAdd($name,$surname,$email,$tenant)
                        }
                }
        }else{
                Write-Host "[!] ERROR: Ha introducido algun parametro incorrecto"
                pause
                addOneUser
        }
}

# Function to add Users from CSV file (Under Construction)
function addMassiveUser{
        Write-Host "Opcion massive User"
}

# Main menu with all options, if you want to quit press "q".
function showMenu{
        param ([string]$title = 'Menu de creacion de buzon Exchange')

        Clear-host
        Write-Host "================ $title ================"

        Write-Host "1: Agregar usuario individual."
        Write-Host "2: Agregar usuarios masivo."
        Write-Host "Q: Salir."
}

do
{
        showMenu
        $selection = Read-Host "Elija la opcion que quiere llevar a cabo"

        switch($selection)
        {
                '1' {
                        addOneUser
                }
                '2'{
                        addMassiveUser
                }

        }
}
until ($selection -eq 'q')