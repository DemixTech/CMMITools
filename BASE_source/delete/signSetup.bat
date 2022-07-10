Rem option 1 | signtool.exe sign /fd sha256 /v /n "Demix" /tr "http://ts.ssl.com" /td sha256 "..\BASE_install\setup.exe"

Rem option 2 | signtool.exe sign /fd sha256 /v /a /tr "http://ts.ssl.com" /td sha256 "..\BASE_install\setup.exe"

Rem /n | If you have more than one code signing USB tokens or certificates installed, 
Rem		| you can specify the certificate you want to use by including its Subject Name via the /n option.
Rem		| Looking at the Demix ssl EV certificate it is issued to "Demix"