on run {input, parameters}
    -- Assuming the last item in 'input' is the selected folder
    set outputFolder to item (count of input) of input
    tell application "Microsoft PowerPoint" to launch
    repeat with i from 1 to (count of input) - 1
		set curFile to item i of input
		if (curFile as String ends with ".ppt") or (curFile as String ends with ".pptx") then
			set outputPath to my createPDFPathInFolder(curFile, outputFolder)
			my savePowerPointAsPDF(curFile, outputPath)
		else
			display dialog "Skipping file (not a .ppt or .pptx): " & curFile buttons {"OK"} default button 1
        end if
    end repeat
    tell application "Microsoft PowerPoint" to quit
end run

on createPDFPathInFolder(inputFile, folderPath)
	set fileInfo to info for inputFile
	set baseNameExt to name of (fileInfo) as string
	
	set dotIndex to (offset of "." in (reverse of characters of baseNameExt) as string) as integer
	if dotIndex is not 0 then
		set baseName to text 1 thru -(dotIndex + 1) of baseNameExt
	end if
	set pdfName to baseName & ".pdf"
	set pdfPath to (folderPath as string) & pdfName
	return pdfPath
end createPDFPathInFolder


on savePowerPointAsPDF(documentAlias, pdfPath)
	tell application "Microsoft PowerPoint"
		open documentAlias
		set pdfPath to my createEmptyFile(pdfPath)
		delay 1
		save active presentation in pdfPath as save as PDF
		delay 1
		close active presentation
	end tell
end savePowerPointAsPDF

on createEmptyFile(f)
	do shell script "touch " & quoted form of POSIX path of f
	return (POSIX path of f) as POSIX file
end createEmptyFile
