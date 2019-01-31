<%

	' INSTALLATION:
	'************************************************************************
	
	' Uses CryptoHelper (a Standalone password hasher for ASP.NET Core using a PBKDF2 implementation) by henkmollema
	' https://github.com/henkmollema/CryptoHelper
	
	' Make sure you have the lastest .NET Framework installed (tested on .NET Framework 4.7.2)
	
	' Open IIS, go to the applicaton pools and open the application pool being used by your 
	' website. Check the .NET CRL version
	' E.g: v4.0.30319
	
	' Navigate to the CRL folder
	' E.g: C:\Windows\Microsoft.NET\Framework64\v4.0.30319
	
	' Copy over: 
	' ClassicASP.PBKDF2.dll,
	' Microsoft.AspNetCore.Cryptography.Internal.dll,
	' Microsoft.AspNetCore.Cryptography.KeyDerivation.dll
	
	' Run CMD as administrator
	
	' Change the directory to your CRL folder
	' E.g: cd C:\Windows\Microsoft.NET\Framework64\v4.0.30319
	
	' Run the following command: RegAsm ClassicASP.PBKDF2.dll /tlb /codebase
	
	
	' PBKDF2 IN VBSCRIPT:
	'************************************************************************
	
	' Defaults used by ClassicASP.PBKDF2 if none specified:
			
	' PBKDF2_iterations = 10000
	' PBKDF2_alg = sha1
	' PBKDF2_saltBytes = 16 bytes
	' PBKDF2_keyLength = 32 bytes
	
	Const PBKDF2_iterations = 30000
	Const PBKDF2_alg = "sha512" ' only sha1, sha256 and sha512 are supported
	Const PBKDF2_saltBytes = 16
	Const PBKDF2_keyLength = 32
		
	function PBKDF2_hash(ByVal password)
		
		Dim PBKDF2 : set PBKDF2 = server.CreateObject("ClassicASP.PBKDF2")
		
			PBKDF2_hash = 	PBKDF2.hash(_
					password,_
					PBKDF2_iterations,_
					PBKDF2_alg,_
					PBKDF2_saltBytes,_
					PBKDF2_keyLength)
		
		set PBKDF2 = nothing
		
	end function
	
	function PBKDF2_verify(ByVal password, ByVal PBKDF2Hash)
	
		Dim PBKDF2 : set PBKDF2 = server.CreateObject("ClassicASP.PBKDF2")
			
			PBKDF2_verify = PBKDF2.verify(password,PBKDF2Hash)
			
		set PBKDF2 = nothing
		
	end function
	
	Dim pb2_hash, start, testPassword
	
	testPassword = "myPassword"
	
	response.write "<p><b>Password:</b> " & testPassword & "</p>"
	
	start = Timer()
	
	pb2_hash = PBKDF2_hash(testPassword)
				
	response.write "<p><b>PBKDF2 Hash:</b> " & pb2_hash & "</p>"
	response.write "<p><b>Time to execute:</b> " & formatNumber(Timer()-start,4) & "s</p>"
	
	start = Timer()
	
	response.write "<p><b>PBKDF2 Verified:</b> " & PBKDF2_verify(testPassword,pb2_hash) & "</p>"
	response.write "<p><b>Time to execute:</b> " & formatNumber(Timer()-start,4) & "s</p>"
	
%>
