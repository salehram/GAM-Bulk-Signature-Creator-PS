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
        # End of listing active directory users
    }
    "CSV"{ #CSV file source
        Write-Host "User list set to read from CSV file" -BackgroundColor Black -ForegroundColor Green
        #
        # Extracting the CSV file path
        $finalCSVFilePath=$csvFileRawPath -Split "csv:"
        $csvFN = $finalCSVFilePath[1]
        write-host "CSV file to use: $csvFN" -BackgroundColor Black -ForegroundColor Yellow
    }
    "EMAIL"{ #EMAIL file source
        Write-Host "User list set to read from a single email address" -BackgroundColor Black -ForegroundColor Green
        write-host "Getting the email address..." -BackgroundColor black -ForegroundColor Yellow
        $SingleEmailAddress=$fullEamilParameter -split "email:"
        $UserEmail=$SingleEmailAddress[1]
    }
    default{ # Nothing has been spicified, or we have wrong parameters passed to the command.

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
                # Reading the user info file
                $userData = Import-Csv userInfo.csv
                foreach ($Entry in $userData) {
                    # Variables section
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
                    $IPPhone = $Entry.IPPhone
                    # End of variables section
                    write-host "Working on $FullName" -BackgroundColor Yellow -ForegroundColor Red
                    #
                    # Reading the content of the original signature file, and replacing variables in a run-time signature file
                    # Replacing variables section
                    (Get-Content signature.txt) -Replace "_DisplayName",$FullName | Out-File signature-runtime.txt -Encoding utf8
                    (Get-Content signature-runtime.txt) -Replace "_Title",$TitleName | Out-File signature-runtime.txt -Encoding utf8
                    (Get-Content signature-runtime.txt) -Replace "_Company",$CompanyName | Out-File signature-runtime.txt -Encoding utf8
                    (Get-Content signature-runtime.txt) -Replace "_Department",$DepartmentName | Out-File signature-runtime.txt -Encoding utf8
                    (Get-Content signature-runtime.txt) -Replace "_HomePhone",$HomePhoneNo | Out-File signature-runtime.txt -Encoding utf8
                    (Get-Content signature-runtime.txt) -Replace "_mail",$EmailAddr | Out-File signature-runtime.txt -Encoding utf8
                    (Get-Content signature-runtime.txt) -Replace "_POBox",$PoBoxNo | Out-File signature-runtime.txt -Encoding utf8
                    (Get-Content signature-runtime.txt) -Replace "_City",$CityName | Out-File signature-runtime.txt -Encoding utf8
                    (Get-Content signature-runtime.txt) -Replace "_PostalCode",$PostalCodeNo | Out-File signature-runtime.txt -Encoding utf8
                    (Get-Content signature-runtime.txt) -Replace "_State",$StateName | Out-File signature-runtime.txt -Encoding utf8
                    (Get-Content signature-runtime.txt) -Replace "_Country",$CountryName | Out-File signature-runtime.txt -Encoding utf8
                    (Get-Content signature-runtime.txt) -Replace "_Fax",$FaxNo | Out-File signature-runtime.txt -Encoding utf8
                    (Get-Content signature-runtime.txt) -Replace "_MobilePhone",$MobilePhoneNo | Out-File signature-runtime.txt -Encoding utf8
                    (Get-Content signature-runtime.txt) -Replace "_OfficePhone",$OfficePhoneNo | Out-File signature-runtime.txt -Encoding utf8
                    (Get-Content signature-runtime.txt) -Replace "_StreetAddress",$StreetAddrDesc | Out-File signature-runtime.txt -Encoding utf8
                    (Get-Content signature-runtime.txt) -Replace "_IPPhone",$IPPhone | Out-File signature-runtime.txt -Encoding utf8
                    # End of replacing variables section
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
                $userCSVSource = Import-Csv $finalCSVFilePath[1]
                foreach ($Entry in $userCSVSource) {
                    #
                    # Getting the user information from active directory, and storing them into a text file
                    $userID=$Entry.EmailAddress
                    # Listing AD users to show what we will work on
                    Get-ADUser -Filter {mail -eq $userID} -Properties * | Select-Object mail, DisplayName, City, Company, Department, Fax, HomePhone, MobilePhone, OfficePhone, POBox, PostalCode, State, Country, StreetAddress, Title | Export-Csv userInfo.csv -notype
                    # End of listing active directory users
                    #
                    # Invoking GAM with the user information and signature file
                    $userData = Import-Csv userInfo.csv
                    foreach ($Entry in $userData) {
                        # Variables section
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
                        $IPPhone = $Entry.IPPhone
                        # End of variables section
                        write-host "Working on $FullName" -BackgroundColor Yellow -ForegroundColor Red
                        #
                        # Reading the content of the original signature file, and replacing variables in a run-time signature file
                        # Replacing variables section
                        (Get-Content signature.txt) -Replace "_DisplayName",$FullName | Out-File signature-runtime.txt -Encoding utf8
                        (Get-Content signature-runtime.txt) -Replace "_Title",$TitleName | Out-File signature-runtime.txt -Encoding utf8
                        (Get-Content signature-runtime.txt) -Replace "_Company",$CompanyName | Out-File signature-runtime.txt -Encoding utf8
                        (Get-Content signature-runtime.txt) -Replace "_Department",$DepartmentName | Out-File signature-runtime.txt -Encoding utf8
                        (Get-Content signature-runtime.txt) -Replace "_HomePhone",$HomePhoneNo | Out-File signature-runtime.txt -Encoding utf8
                        (Get-Content signature-runtime.txt) -Replace "_mail",$EmailAddr | Out-File signature-runtime.txt -Encoding utf8
                        (Get-Content signature-runtime.txt) -Replace "_POBox",$PoBoxNo | Out-File signature-runtime.txt -Encoding utf8
                        (Get-Content signature-runtime.txt) -Replace "_City",$CityName | Out-File signature-runtime.txt -Encoding utf8
                        (Get-Content signature-runtime.txt) -Replace "_PostalCode",$PostalCodeNo | Out-File signature-runtime.txt -Encoding utf8
                        (Get-Content signature-runtime.txt) -Replace "_State",$StateName | Out-File signature-runtime.txt -Encoding utf8
                        (Get-Content signature-runtime.txt) -Replace "_Country",$CountryName | Out-File signature-runtime.txt -Encoding utf8
                        (Get-Content signature-runtime.txt) -Replace "_Fax",$FaxNo | Out-File signature-runtime.txt -Encoding utf8
                        (Get-Content signature-runtime.txt) -Replace "_MobilePhone",$MobilePhoneNo | Out-File signature-runtime.txt -Encoding utf8
                        (Get-Content signature-runtime.txt) -Replace "_OfficePhone",$OfficePhoneNo | Out-File signature-runtime.txt -Encoding utf8
                        (Get-Content signature-runtime.txt) -Replace "_StreetAddress",$StreetAddrDesc | Out-File signature-runtime.txt -Encoding utf8
                        (Get-Content signature-runtime.txt) -Replace "_IPPhone",$IPPhone | Out-File signature-runtime.txt -Encoding utf8
                        # End of replacing variables section
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
                        .\gam\gam.exe user $EmailAddr signature file $sigFilePath
                    }
                }
            }
            "EMAIL"{
                #
                # User source was set to single email address
                # We will just need to query the AD based on the email we have, then complete the rest process as normal
                # Listing AD users to show what we will work on
                Get-ADUser -Filter {mail -eq $UserEmail} -Properties * | Select-Object mail, DisplayName, City, Company, Department, Fax, HomePhone, MobilePhone, OfficePhone, POBox, PostalCode, State, Country, StreetAddress, Title | Export-Csv userInfo.csv -notype
                # End of listing active-directory users
                #
                # Invoking GAM with the user information and signature file
                $userData = Import-Csv userInfo.csv
                foreach ($Entry in $userData) {
                    # Variables section
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
                    $IPPhone = $Entry.IPPhone
                    # End of variables section
                    write-host "Working on $FullName" -BackgroundColor Yellow -ForegroundColor Red
                    #
                    # Reading the content of the original signature file, and replacing variables in a run-time signature file
                    # Replacing variables section
                    (Get-Content signature.txt) -Replace "_DisplayName",$FullName | Out-File signature-runtime.txt -Encoding utf8
                    (Get-Content signature-runtime.txt) -Replace "_Title",$TitleName | Out-File signature-runtime.txt -Encoding utf8
                    (Get-Content signature-runtime.txt) -Replace "_Company",$CompanyName | Out-File signature-runtime.txt -Encoding utf8
                    (Get-Content signature-runtime.txt) -Replace "_Department",$DepartmentName | Out-File signature-runtime.txt -Encoding utf8
                    (Get-Content signature-runtime.txt) -Replace "_HomePhone",$HomePhoneNo | Out-File signature-runtime.txt -Encoding utf8
                    (Get-Content signature-runtime.txt) -Replace "_mail",$EmailAddr | Out-File signature-runtime.txt -Encoding utf8
                    (Get-Content signature-runtime.txt) -Replace "_POBox",$PoBoxNo | Out-File signature-runtime.txt -Encoding utf8
                    (Get-Content signature-runtime.txt) -Replace "_City",$CityName | Out-File signature-runtime.txt -Encoding utf8
                    (Get-Content signature-runtime.txt) -Replace "_PostalCode",$PostalCodeNo | Out-File signature-runtime.txt -Encoding utf8
                    (Get-Content signature-runtime.txt) -Replace "_State",$StateName | Out-File signature-runtime.txt -Encoding utf8
                    (Get-Content signature-runtime.txt) -Replace "_Country",$CountryName | Out-File signature-runtime.txt -Encoding utf8
                    (Get-Content signature-runtime.txt) -Replace "_Fax",$FaxNo | Out-File signature-runtime.txt -Encoding utf8
                    (Get-Content signature-runtime.txt) -Replace "_MobilePhone",$MobilePhoneNo | Out-File signature-runtime.txt -Encoding utf8
                    (Get-Content signature-runtime.txt) -Replace "_OfficePhone",$OfficePhoneNo | Out-File signature-runtime.txt -Encoding utf8
                    (Get-Content signature-runtime.txt) -Replace "_StreetAddress",$StreetAddrDesc | Out-File signature-runtime.txt -Encoding utf8
                    (Get-Content signature-runtime.txt) -Replace "_IPPhone",$IPPhone | Out-File signature-runtime.txt -Encoding utf8
                    # End of replacing variables section
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
                    .\gam\gam.exe user $EmailAddr signature file $sigFilePath
                }
            }
        }
    } # End of "AD" case
    # Start of "CSV" case
    "CSV"{
        # If we get this parameter as information source, then this will automatically override the other parameter for user source, as this parameter will point to a csv file
        # that contains a field with the user email address, which can be used as an ID for the user we want to assign the signature for him...
        # We will not make AD queries. We will only invoke GAM with the data of the csv file we have.
        #
        # Invoking GAM with the user information and signature file
        $userData = Import-Csv userInfo.csv
        foreach ($Entry in $userData) {
            # Variables section
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
            $IPPhone = $Entry.IPPhone
            # End of variables section
            write-host "Working on $FullName" -BackgroundColor Yellow -ForegroundColor Red
            #
            # Reading the content of the original signature file, and replacing variables in a run-time signature file
            # Replacing variables section
            (Get-Content signature.txt) -Replace "_DisplayName",$FullName | Out-File signature-runtime.txt -Encoding utf8
            (Get-Content signature-runtime.txt) -Replace "_Title",$TitleName | Out-File signature-runtime.txt -Encoding utf8
            (Get-Content signature-runtime.txt) -Replace "_Company",$CompanyName | Out-File signature-runtime.txt -Encoding utf8
            (Get-Content signature-runtime.txt) -Replace "_Department",$DepartmentName | Out-File signature-runtime.txt -Encoding utf8
            (Get-Content signature-runtime.txt) -Replace "_HomePhone",$HomePhoneNo | Out-File signature-runtime.txt -Encoding utf8
            (Get-Content signature-runtime.txt) -Replace "_mail",$EmailAddr | Out-File signature-runtime.txt -Encoding utf8
            (Get-Content signature-runtime.txt) -Replace "_POBox",$PoBoxNo | Out-File signature-runtime.txt -Encoding utf8
            (Get-Content signature-runtime.txt) -Replace "_City",$CityName | Out-File signature-runtime.txt -Encoding utf8
            (Get-Content signature-runtime.txt) -Replace "_PostalCode",$PostalCodeNo | Out-File signature-runtime.txt -Encoding utf8
            (Get-Content signature-runtime.txt) -Replace "_State",$StateName | Out-File signature-runtime.txt -Encoding utf8
            (Get-Content signature-runtime.txt) -Replace "_Country",$CountryName | Out-File signature-runtime.txt -Encoding utf8
            (Get-Content signature-runtime.txt) -Replace "_Fax",$FaxNo | Out-File signature-runtime.txt -Encoding utf8
            (Get-Content signature-runtime.txt) -Replace "_MobilePhone",$MobilePhoneNo | Out-File signature-runtime.txt -Encoding utf8
            (Get-Content signature-runtime.txt) -Replace "_OfficePhone",$OfficePhoneNo | Out-File signature-runtime.txt -Encoding utf8
            (Get-Content signature-runtime.txt) -Replace "_StreetAddress",$StreetAddrDesc | Out-File signature-runtime.txt -Encoding utf8
            (Get-Content signature-runtime.txt) -Replace "_IPPhone",$IPPhone | Out-File signature-runtime.txt -Encoding utf8
            # End of replacing variable section
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
            .\gam\gam.exe user $EmailAddr signature file $sigFilePath
        }
    } # End of "CSV" case
}
