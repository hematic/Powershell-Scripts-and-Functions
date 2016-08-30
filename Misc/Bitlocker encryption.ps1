#Function to set up logging for the script.
function Log 
{ 
    param
    (
        [string]$strMessage
    )

        $LogDir = [environment]::GetEnvironmentVariable("TEMP")
        $Logfile = "\BitLocker.txt"
        $Path = $logdir + $logfile
        [string]$strDate = get-date
        add-content -path $Path -value ($strDate + "`t:`t"+ $strMessage)
}

#Function to control Error parsing.
function errDescription($intErr)
{
    Switch($intErr)
    {
        0          { "The method was successful." }
        2147942487 { "The OwnerAuth parameter is not valid." }
        2147943755 { "Connection Failed : Cannot save recovery information to the network. The computer has been configured to store recovery information to Active Directory Domain Services. A network connection is required to continue." }
        2150105089 { "The provided owner authorization value cannot fulfill the request." }
        2150105096 { "An endorsement key pair already exists on this TPM." }
        2150105099 { "An owner cannot be installed on this TPM." }
        2150105108 { "An owner already exists on the TPM." }
        2150105123 { "No endorsement key can be found on the TPM." }
        2150107139 { "The TPM is defending against dictionary attacks and is in a time-out period. For more information, see the ResetAuthLockOut method." }
        2150171392 { "A hardware failure occurred. Consult your computer manufacturer for more information." }
        2150171393 { "The user rejected the requested TPM operation." }
        2150171394 { "A BIOS failure occurred while running the TPM operation." }
        2150694922 { "Cannot save recovery information to the network. The computer has been configured to store recovery information to Active Directory Domain Services. For instructions on how to set up Active Directory, see BitLocker Drive Encryption Configuration Guide: Backing Up BitLocker and TPM Recovery Information to Active Directory." }2150694912 { "FVE_E_LOCKED_VOLUME : The volume is locked."}
        2150121480 { "No compatible TPM is found on this computer."}
        2150694947 { "The TPM cannot secure the volume's encryption key because the volume does not contain the currently running operating system."}
        2150694958 { "No encryption key exists for the volume"}
        2150694957 { "The provided encryption method does not match that of the partially or fully encrypted volume"}
        2150694942 { "The volume cannot be encrypted because this computer is configured to be part of a server cluster"}
        2150694912 { "The volume is locked."}
        2150694956 { "No key protectors of the type Numerical Password are specified. The Group Policy requires a backup of recovery information to Active Directory Domain Services. To add at least one key protector of that type, use the ProtectKeyWithNumericalPassword method."}
        2150694936 { "You must initialize the Trusted Platform Module (TPM) before you can use BitLocker Drive Encryption."}
        2150694960 { "A bootable CD/DVD is found in this computer. Remove the CD/DVD and restart the computer."}
        2150694961 { "A key protector of this type already exists."}
        DEFAULT    { "Unknown error code : " + $intErr}
    }
}

Log("********************")
Log("* SCRIPT Begins    *")
Log("********************")

$tpm = Get-WmiObject -Class Win32_TPM -EnableAllPrivileges -Namespace "root\CIMV2\Security\MicrosoftTpm"
$IsActivated = $tpm.IsActivated_InitialValue
$IsEnabled = $tpm.IsEnabled_InitialValue
$IsOwned = $tpm.IsOwned_InitialValue 

IF ($IsActivated -and $IsEnabled){Log("TPM is activated and enabled")}

IF ($IsOwned -eq $false)
{  
            Log ("No owner assigned, creating owner information")
            
            #Generate 128 character random passphrase.

            [Reflection.Assembly]::LoadWithPartialName(“System.Web”)
            $RandomPassphrase = [System.Web.Security.Membership]::GeneratePassword(128,32)
            
            # Translate random passphrase into 20-byte owner authorization 

            $tpm.ConvertToOwnerAuth($RandomPassphrase) | FOREACH {$OwnerAuth = $_.OwnerAuth ; $ReturnValue = $_.ReturnValue}
            Log ("Translate random passphrase into 20-byte owner authorization : " + (ErrDescription($ReturnValue)))
         
            # Attempt to take Ownership of TPM
            $tpm.TakeOwnership($OwnerAuth) | FOREACH {$ReturnValue = $_.ReturnValue}
            Log ("Attempt to take ownership of TPM : " + (ErrDescription($ReturnValue)))
                       
        }

ELSE{Log ("Owner already exists....continuing")}
        
# Get Win32_EncryptableVolume class
$EncryptableDrive = gwmi -namespace root\CIMV2\Security\MicrosoftVolumeEncryption -query "SELECT * FROM Win32_EncryptableVolume Where DriveLetter='C:'"

# Secures the volume's encryption key by using the Trusted Platform Module (TPM)
$EncryptableDrive.ProtectKeyWithTPM() | FOREACH {$ReturnValue = $_.ReturnValue}
Log ("Protecting Encryption Key with TPM : " + (ErrDescription($ReturnValue)))
    
        
IF ($ReturnValue -eq 0 -or $ReturnValue -eq 2150694961)
{
    Log("Secured encrypted volume master key with TPM")
    $EncryptableDrive.ProtectKeyWithNumericalPassword() | FOREACH {$ReturnValue = $_.ReturnValue}

    IF ($ReturnValue -eq 0)
    {
        Log("All BitLocker pre-requisites confirmed, now attempting to encrypt")
        $EncryptableDrive.Encrypt()| FOREACH {$ReturnValue = $_.ReturnValue}
    }
    
    Log ("Drive encryption status : " + (ErrDescription($ReturnValue)))
    
    IF ($ReturnValue -eq 0){Log("Encryption has begun")}

    ELSE{Log("Error encrypting drive.")}
}
    
ELSE{Log("Error protecting encryption key")}

ELSE
{  
    Log ("Pre-requisites not met! See information below")
    Log ("TPM Enabled   : " + [string]$IsEnabled)
    Log ("TPM Activated : " + [string]$IsActivated)
}
        
Log("********************")
Log("* SCRIPT ENDS      *")
Log("********************")