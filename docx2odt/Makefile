#!/usr/bin/env make

# CS = /cygdrive/c/Windows/Microsoft.NET/Framework64/v4.0.30319/csc.exe
# CS = /cygdrive/c/Windows/Microsoft.NET/Framework/v4.0.30319/csc.exe
# CS = /cygdrive/c/Windows/Microsoft.NET/Framework64/v3.5/csc.exe
# CS = /cygdrive/c/Windows/Microsoft.NET/Framework/v3.5/csc.exe
CS = "/cygdrive/c/Program Files (x86)/MSBuild/14.0/Bin/csc.exe"

.PHONY: EXE info clean

EXE: docx2odt.cs
	@echo "compiling EXE ..."
	# $(CS) /target:exe /platform:x86 docx2odt.cs
	$(CS) /target:exe /platform:x86 docx2odt.cs /r:Microsoft.Office.Interop.Word.dll /r:System.IO.Compression.dll /r:System.IO.Compression.ZipFile.dll /r:System.IO.Compression.FileSystem.dll

info:
	$(CS) /version

clean:
	@echo "cleaning up ..."
	rm -rf docx2odt.exe

