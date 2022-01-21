#!/usr/bin/env make

# CS = /cygdrive/c/Windows/Microsoft.NET/Framework64/v4.0.30319/csc.exe
# CS = /cygdrive/c/Windows/Microsoft.NET/Framework/v4.0.30319/csc.exe
# CS = /cygdrive/c/Windows/Microsoft.NET/Framework64/v3.5/csc.exe
# CS = /cygdrive/c/Windows/Microsoft.NET/Framework/v3.5/csc.exe
CS = "/cygdrive/c/Program Files (x86)/MSBuild/14.0/Bin/csc.exe"

DLL = "/r:./dll/Microsoft.Office.Interop.Word.dll"
DLL += "/r:./dll/System.IO.Compression.dll"
DLL += "/r:System.IO.Compression.ZipFile.dll" 
DLL += "/r:System.IO.Compression.FileSystem.dll"

.PHONY: EXE clean cleanall

EXE: docx2txt.cs
	$(CS) /target:exe /platform:x86 docx2txt.cs $(DLL)

clean:
	rm -rf docx2txt.exe

cleanall: clean
	rm -rf *.odt

