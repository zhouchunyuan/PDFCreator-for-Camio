; PDFCreator.nsi
;
; It will install PDFCreator into the [D:\LGE_Report\Program] directory 
;
;--------------------------------


;--------------------------------

; The name of the installer
Name "PDF Creator for Nikon Camio"

; The file to write
OutFile "PDFCreator.exe"

; The default installation directory
InstallDir D:\LGE_Report\Program

;--------------------------------

; Pages

Page directory
Page instfiles

;--------------------------------

; The stuff to install
Section "" ;No components page, name is not important

  ; Set output path to the installation directory.
  SetOutPath $INSTDIR
  
  ; Put file there
  FILE /r "build\exe.win-amd64-2.7\*.*"
  
SectionEnd ; end the section