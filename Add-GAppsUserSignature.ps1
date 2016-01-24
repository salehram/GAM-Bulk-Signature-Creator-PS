# This script is written by: Saleh Ramadan (salehram@gmail.com) - http://www.is-linux.com
# The script is provided AS IS and I cannot provide directed support other than fixing bugs and maintaining the stability of the script.
# I accept NO liabilty or responsiblity for the improper use of this script.
#
# This script will extract the following data from AD and attach them to the HTML signature body:
#     1. Full name
#     2. Email address
#     3. Phone number 1 (mobile)
#     4. Phone number 2 (land line)
#     5. Extension for phone number 2
#     6. Street address
#     7. City
#     8. Country
#     9. Organization name
#
# Args:
# 1. User source
#     a. All users -> will run GAM to get a list with all users.
#     b. CSV list
#     c. Single user
# 2. Information source
#     a. AD -> will run Get-ADUser to get the information and populate them.
#     b. CSV file with columns of information mentioned above
#
# Initializing parameters:

[CmdletBinding()]
Param(
    [Parameter(Mandatory=$True,Position=1)]
    [string]$UserSource,

    [Parameter(Mandatory=$True)]
    [string]$InfoSource
)
Write-Host "Checking parameters..." -BackgroundColor Black -ForegroundColor Yellow
#
# Checking to analyse the parameter to see if it is a CSV flag
if ($UserSource -match '(?i)^CSV:'){
    $csvFileRawPath = $UserSource
    $UserSource = "CSV"
}
#
# Checking to analyse the parameter to see if it is an EMAIL flag
if ($UserSource -match '(?i)^EMAIL:'){
    $fullEamilParameter = $UserSource
    $UserSource = "EMAIL"
}
#
# Validating the parameters
switch ($UserSource)
{
    "AD"{ #Read from AD
        Write-Host "User list set to read from Active Directory." -BackgroundColor Black -ForegroundColor Green
        # Getting users array
        # Getting users info
        # Executing GAM with users info on signature data
        #
        # 1. Getting the BaseDN to use it with the search query
        $LocalDomain = Read-Host "Please enter the local domain name where you are executing the script: "
        #
        # Clearing already used variables
        $searchBase = $null
        $tempStr = $null
        $DNPathRAW = $null
        $arrayCount = 0
        $counter = 0
        #
        # Reading the local domain value and extract it into array
        $tempStr = $LocalDomain.split(".")
        $arrayCount = $tempStr.Count
        Write-Host "Array count = $arrayCount"
        #
        # Creating the default BaseDN from the above array
        do{
            $DNPathRAW = $DNPathRAW + "DC=" + $tempStr[$counter].ToString() + ","
            $counter++
            # FOR DEBUG
            #Write-Host $counter
            #Write-Host $arrayCount
            #
        } while ($counter -lt $arrayCount)
        #
        # Refining the final BaseDN value by removing the extra comma at the end of the string
        $DefaultsearchBase = $DNPathRAW.Substring(0,$DNPathRAW.Length-1)
        #
        # Getting user input to define the BaseDN
        write-host "Please provide the BaseDN value to use it with the search. If you don't know what the BaseDN value is, please see the below examples:"
        Write-Host "The BaseDN for all users in the domain '$LocalDomain' is: '$DefaultsearchBase'"
        Write-Host "Enter search base DN to look for users objects in (type 'base' to use the default DN for the domain):"
        $BaseDN = Read-Host "BaseDN value: "
        switch ($BaseDN){
            "base"{
                #
                # Using the default BaseDN value
                $searchBase = $DefaultsearchBase
                #
                # The final BaseDN value
                Write-Host "Using BaseDN: $searchBase" -BackgroundColor Black -ForegroundColor Yellow
                Write-Host "Searching active directory with BaseDN $searchBase..." -BackgroundColor Black -ForegroundColor Yellow
            }
            default{
                Write-Host "BaseDN: $BaseDN" -BackgroundColor Yellow -ForegroundColor Black
                $searchBase=$BaseDN
            }
        }
        #
        # Listing AD users to show what we will work on
        write-host "Listing AD users to show what we will work on" -BackgroundColor Black -ForegroundColor Yellow
        get-adUser -LDAPFilter "(name=*)" -searchBase "$searchBase" -Properties * | Select-Object mail, DisplayName, City, Company, Department, Fax, HomePhone, MobilePhone, OfficePhone, POBox, PostalCode, State, Country, StreetAddress, Title | ft -AutoSize
        get-adUser -LDAPFilter "(name=*)" -searchBase "$searchBase" -Properties * | Select-Object mail, DisplayName, City, Company, Department, Fax, HomePhone, MobilePhone, OfficePhone, POBox, PostalCode, State, Country, StreetAddress, Title | Export-Csv userInfo.csv -notype
    }
    "CSV"{ #CSV file source
        Write-Host "User list set to read from CSV file" -BackgroundColor Black -ForegroundColor Green
        #
        # Extracting the CSV file path
        $finalFilePath=$csvFileRawPath -Split "csv:"
        write-host "CSV file to use: $finalFilePath" -BackgroundColor Black -ForegroundColor Yellow
    }
    "EMAIL"{ #EMAIL file source
        Write-Host "User list set to read from a single email address" -BackgroundColor Black -ForegroundColor Green
        write-host "Getting the email address..." -BackgroundColor black -ForegroundColor Yellow
        $fullEamilParameter
    }
    default{ #Something else, we will validate it and assume it is an email address

    }
}
#
# Information source
switch ($InfoSource)
{
    "AD"{
        #
        # Before, we need to check if the user source was set to AD, because this will make it easier for us
        switch ($UserSource){
            "AD"{
                #
                # Read the CSV file and execute GAM with HTML signature code
                # Reading the signature text file
                $userData = Import-Csv userInfo.csv
                foreach ($Entry in $userData) {
                    $FullName = $Entry.DisplayName
                    $EmailAddr = $Entry.mail
                    $CityName = $Entry.City
                    $CompanyName = $Entry.Company
                    $DepartmentName = $Entry.Department
                    $FaxNo = $Entry.Fax
                    $HomePhoneNo = $Entry.HomePhone
                    $MobilePhoneNo = $Entry.MobilePhone
                    $OfficePhoneNo = $Entry.OfficePhone
                    $PoBoxNo = $Entry.POBox
                    $PostalCodeNo = $Entry.PostalCode
                    $StateName = $Entry.State
                    $CountryName = $Entry.Country
                    $StreetAddrDesc = $Entry.StreetAddress
                    $TitleName = $Entry.Title
                    write-host "Working on $FullName" -BackgroundColor Yellow -ForegroundColor black
                    #
                    # Reading the content of the original signature file, and replacing variables in a run-time signature file
                    (Get-Content signature.txt).Replace('$DisplayName',$FullName) | Set-Content signature-runtime.txt
                    (Get-Content signature-runtime.txt).Replace('$Title',$TitleName) | Set-Content signature-runtime.txt
                    (Get-Content signature-runtime.txt).Replace('$Company',$CompanyName) | Set-Content signature-runtime.txt
                    (Get-Content signature-runtime.txt).Replace('$Department',$DepartmentName) | Set-Content signature-runtime.txt
                    (Get-Content signature-runtime.txt).Replace('$HomePhone',$HomePhoneNo) | Set-Content signature-runtime.txt
                    (Get-Content signature-runtime.txt).Replace('$mail',$EmailAddr) | Set-Content signature-runtime.txt
                    (Get-Content signature-runtime.txt).Replace('$POBox',$PoBoxNo) | Set-Content signature-runtime.txt
                    (Get-Content signature-runtime.txt).Replace('$City',$CityName) | Set-Content signature-runtime.txt
                    (Get-Content signature-runtime.txt).Replace('$PostalCode',$PostalCodeNo) | Set-Content signature-runtime.txt
                    (Get-Content signature-runtime.txt).Replace('$State',$StateName) | Set-Content signature-runtime.txt
                    (Get-Content signature-runtime.txt).Replace('$Country',$CountryName) | Set-Content signature-runtime.txt
                    (Get-Content signature-runtime.txt).Replace('$Fax',$FaxNo) | Set-Content signature-runtime.txt
                    (Get-Content signature-runtime.txt).Replace('$MobilePhone',$MobilePhoneNo) | Set-Content signature-runtime.txt
                    (Get-Content signature-runtime.txt).Replace('$OfficePhone',$OfficePhoneNo) | Set-Content signature-runtime.txt
                    (Get-Content signature-runtime.txt).Replace('$StreetAddress',$StreetAddrDesc) | Set-Content signature-runtime.txt
                    #
                    # Assigning the signature text to the variable
                    $signatureTxt = Get-Content signature-runtime.txt
                    Write-Host "Signature text we will apply is:" -BackgroundColor Black -ForegroundColor Green
                    Write-Host $signatureTxt
                    #
                    # Invoking GAM tool to read the signature text
                    Write-Host "Invoking GAM for user $EmailAddr" -BackgroundColor Yellow -ForegroundColor Black
                    $currentWD=Convert-Path .
                    $sigFilePath = "$currentWD\signature-runtime.txt"
                    .\GAM\gam.exe user $EmailAddr signature file $sigFilePath
                }
            }
            "CSV"{
                #
                # User source was set to CSV
                # We need to read the CSV file, inquiry AD for required info, and invoke GAM for each user
            }
            "EMAIL"{
                #
                # User source was set to single email address
            }
        }
    } # End of "AD" case
}
