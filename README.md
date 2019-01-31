PBKDF2 hashing algorithm implemented in VBscript for Classic ASP

## INSTALLATION:
Uses CryptoHelper (a Standalone password hasher for ASP.NET Core using a PBKDF2 implementation) by henkmollema
https://github.com/henkmollema/CryptoHelper

Make sure you have the lastest .NET Framework installed (tested on .NET Framework 4.7.2)
	
Open IIS, go to the applicaton pools and open the application pool being used by your 
Classic ASP application. Check the .NET CRL version
E.g: v4.0.30319
	
Navigate to the CRL folder
E.g: C:\Windows\Microsoft.NET\Framework64\v4.0.30319
	
Copy over: ClassicASP.PBKDF2.dll, Microsoft.AspNetCore.Cryptography.Internal.dll and Microsoft.AspNetCore.Cryptography.KeyDerivation.dll
	
Run CMD as administrator

Change the directory to your CRL folder
E.g: cd C:\Windows\Microsoft.NET\Framework64\v4.0.30319
	
Run the following command: RegAsm ClassicASP.PBKDF2.dll /tlb /codebase

## Usage

	Set PBKDF2 = Server.CreateObject("ClassicASP.PBKDF2")

	' Generate a hash with default parameters (iterations = 10000, algorithm = sha1, saltBytes = 16, keyLength = 32)
	PBKDF2.Hash("myPassword")
	
	' Custom parameters (algorithms supported are: sha1, sha256 and sha512)
	PBKDF2.Hash("myPassword", int iterations, string algorithm, int saltBytes, int keyLength)

	' Verify a hash (the PBKDF2 parameters are contained within the base64 hash string)
	PBKDF2.Verify("myPassword","AQAAAAIAAHUwAAAAEHhMDmEV5UBQt7PP/Da/ughVC2xxpluFxBi7tsseMgD/uVovJw+cY4xrXftimQbYng==") ' True / False

## Example output from PBKDF2.asp

	Password: myPassword
	
	PBKDF2 Hash: AQAAAAIAAHUwAAAAEHhMDmEV5UBQt7PP/Da/ughVC2xxpluFxBi7tsseMgD/uVovJw+cY4xrXftimQbYng==

	Time to execute: 0.0859s

	PBKDF2 Verified: True

	Time to execute: 0.0938s