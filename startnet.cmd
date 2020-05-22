wpeinit
@for %%a in (C D E F G H I J K L M N O P Q R S T U V W X Y Z) do @if exist %%a:\ag-disk-imager\ set IMAGESDRIVE=%%a
@%IMAGESDRIVE%:
@cd ag-disk-imager
@powershell -executionpolicy bypass -nologo -noprofile -file .\ag-disk-imager.ps1