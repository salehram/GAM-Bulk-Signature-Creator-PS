GAM-Bulk-Signature-Creator-PS
=============================

GAM-Bulk-Signature-Creator-PS is a small powershell script that work with GAM to push customized signatures to users in your Google Apps for Work/Education domain.

Prerequisites
-------------

 - GAM must be fully configured and already working (please check GAM's documentation to get it working).
 - AD structure to be existent and working in case of using AD as a source for user or information.
 - Admin access to your Google Apps for Work/Education console.

Usage Instructions
------------------

    .\Add-GAppsUserSignature.ps1 -UserSource [AD,CSV:<CSVFilePath>,EMAIL:<EmailAddress>] -InfoSource [AD,CSV]

The script requires two parameters to be passed when starting the script, in addition to two more while executing the script.

The first 2 are:

 - **-UserSource**: this parameter sets the source from where the script will read the list of users to apply the signature on them.
	 - It can accept 3 arguments:
		 - *AD*, sets the source of information to be pulled from Active-Directory, this will require supplying the local domain name and the baseDN where we want to perform the search.
		 - *CSV*, example: CSV:c:\users\user1\documents\userlist.csv (**the file should only contain one column, with a header called 'EmailAddress', and all users' emails should be listed under this header**), this can be used for having a list of specific users we want to edit their signature without using AD.
		 - *Email*, example: EMAIL:someone@domain.com, this is going to perform the signature edit for this specific user only.
 - **-InfoSource**: this parameter sets the source from where the script will read the signature information to apply it on users we have acquired in the previous parameter.
	 - It can accept 2 arguments:
		 - *AD*, sets the information source to be pulled from Active-Directory. We don't need to have a domain name or a baseDN here, as the user list should be already prepared using the first parameter.
		 - *CSV*, allows us to read the signature information from a pre-filled CSV file. This can be useful as there is no need to specify a CSV file name, as the file name is hardcoded with userInfo.csv, and the file will be sitting in the same script directory. This file is also auto-generated when using the AD as user source. This file, if it was to be made manually, should be formatted to have the following headers:

*Not all of these are mandatory, but it is really recommended to have them all filled*

    mail,DisplayName,City,Company,Department,Fax,HomePhone,MobilePhone,OfficePhone,POBox,PostalCode,State,Country,StreetAddress,Title.

When using AD as user source, the following two parameters will be required to supply to the script in order to be able to search in AD:

 - Local domain name.
 - The base DN where we want to perform the search.

The base DN value is where/what container we will perform the search in AD, it is basically the OU, but the value can be tricky to get/understand...

Let's assume you have the following structure in AD:

    doman.local -> CompanyNameOU -> Users -> Management
                |                         |-> Sales
                |                         |-> Finance
                |                         |-> IT
                -> Users


We want to do the search on the users that are in the Sales OU only. so the base DN will be:

    ou=Sales,ou=Users,ou=CompanyNameOU,dc=domain,dc=local

The above path, will only make us see the objects in the "Sales" OU.

When we want to perform the search on the whole AD structure, (eg. include the users that are existent in domain.local/Users container), then we will use the default base DN in this case, which is easy to write, but has the broadest scope in terms of searching, so it is not recommended to use it:

    dc=domain,dc=local

It is easier to follow an easy rule to quickly determine the path of a specific OU in AD schema:

 - All OUs are referenced as "ou" in schema.
 - When planning to find out the path for a specific OU, always start from the end of the path, in our example we started from "Sales" OU. in other words, we need to follow the hierarchy of the OU we want to search in it.
 - The domain part is always referenced as "dc (domain component)", and each dot (.) is considered a separator. In our example the domain is "domain.local", so that is the domain part and should be at the end of the OU path, and we will be having two segments (dc=domain,dc=local), so if we have something like (west.domain.local), we should get (dc=west,dc=domain,dc=local).
 - The format of the DN path, should always end with "dc" segments.
