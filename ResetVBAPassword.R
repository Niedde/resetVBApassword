## ## ## ## ## ## ## ## ## ## ## ## ## ## ## ## ## ## ## ## ## ## ## ## ## ## ## ## ## ## ## ## actions on the xlsm file

# let the user choose the xlsm file
xlsmFile= file.choose()
pathToXlsmFile= dirname(xlsmFile)
nameOfXlsmFile= basename(xlsmFile)
nameOfXlsmFile_noExtension= substr(nameOfXlsmFile, 1, nchar(nameOfXlsmFile)- 5)

# make a copy of the xlsm file and save it as .zip
setwd(pathToXlsmFile)
nameOfZipFile= paste("copy_of_", nameOfXlsmFile_noExtension, ".zip", sep= "")
nameOfZipFile= gsub(" ", "_", nameOfZipFile) # btw if the name contains spaces, shell won't open it
file.copy(nameOfXlsmFile, paste(pathToXlsmFile, "/", nameOfZipFile, sep= ""))

# open the zip folder using command prompt
openZipFile= paste("start ", nameOfZipFile, sep= "")
shell(openZipFile)

####################################### open xl subfolder and find vbaProject.bin manually
####################################### copy vbaProject.bin into another folder manually
####################################### The second part is here to edit the vbaProject.bin file


## ## ## ## ## ## ## ## ## ## ## ## ## ## ## ## ## ## ## ## ## ## ## ## ## ## ## ## ## ## ## ## open and load the vbaProject.bin file

# let the user choose the vbaProject.bin file manually, then load it
pathInput= file.choose()

startTime= Sys.time()

binFile = file(pathInput,"rb")
fileProperties= file.info(pathInput)
sizeToUse= fileProperties$size
vbaProject= readBin(binFile, "raw", n= sizeToUse) 
vbaProject[1:1994] # display some content (in hexadecimal) for information

setwd(dirname(pathInput))


## ## ## ## ## ## ## ## ## ## ## ## ## ## ## ## ## ## ## ## ## ## ## ## ## ## ## ## ## ## ## ## some functions that will be needed to edit the binary file


stringToVector= function(aString) { # converts a string into a character vector
	for (i in 1:nchar(aString)) {
		result[i]= substr(aString, i, i)
	}
	return (result)
}

vectorToString= function(aCharVector) { # converts a character vector into a string
	if(length(aCharVector)== 0) {result= ""}
	else { 
		result= ""
		for (i in 1:length(aCharVector)) {
			result= paste(result, aCharVector[i], sep= "")
		}
	}
	return (result)
}

`%isPartof%`= function(smallSequence, bigSequence) { # returns TRUE if the small sequence is included in the bigger one
	smallerLength= length(smallSequence)
	biggerLength= length(bigSequence)

	if(smallerLength== 0) {return (TRUE)}
	if(smallerLength> biggerLength) {return (FALSE)}
	else {
		for(firstIndex in 1:(biggerLength- smallerLength+ 1)) {
			if (prod(bigSequence[firstIndex: (firstIndex+ smallerLength- 1)]== smallSequence)== 1) {return (TRUE)} else {next}
		}
	}
	return (FALSE)
}

findSequence= function(smallSequence, bigSequence) { # if the small sequence is included in the bigger one, returns the index of the first occurrence of the small sequence
	smallerLength= length(smallSequence)
	biggerLength= length(bigSequence)

	if(smallerLength== 0) {return (1)}
	if(smallerLength> biggerLength) {return ("nope")}
	else {
		for(firstIndex in 1:(biggerLength- smallerLength+ 1)) {
			if (prod(bigSequence[firstIndex: (firstIndex+ smallerLength- 1)]== smallSequence)== 1) {return (firstIndex)} else {next}
		}
	}
	return ("nope")
}

replaceSequence= function(whatToReplace, whatInstead, whereToReplace) {  # finds the sequence "whatToReplace", and replaces it by the sequence "whatInstead", in the bigger array "whereToReplace"
	if (length(whatToReplace)== 0) {return (whereToReplace)} # if sequence is empty, leave whereToReplace untouched
	if (findSequence(whatToReplace, whereToReplace)== "nope") {return (whereToReplace)} # if sequence not found, leave whereToReplace untouched
	
	else {	# cut the sequence in 3 parts : left part, part to replace, and right part. Left part or right part may be empty
		replacementStartIndex= findSequence(whatToReplace, whereToReplace)
		if(replacementStartIndex> 1) {leftPart= whereToReplace[1: (replacementStartIndex- 1)]} else {leftPart= c()}
		if(replacementStartIndex+ length(whatToReplace)<= length(whereToReplace)) {rightPart= whereToReplace[(replacementStartIndex+ length(whatToReplace)): length(whereToReplace)]} else {rightPart= c()}
		
		# rebuild the vector
		whereToReplace= c(leftPart, whatInstead, rightPart)		
	}

	return (whereToReplace)
}


## ## ## ## ## ## ## ## ## ## ## ## ## ## ## ## ## ## ## ## ## ## ## ## ## ## ## ## ## ## ## ## modify the DPB key in the binary file

# parameters to change in the original vbaProject.bin file
DPBtoReplace= 'DPB='
DPBtoPut= 'DPx='
 
# convert this to raw binary data
DPBtoReplace_raw= as.raw(utf8ToInt(DPBtoReplace))
DPBtoPut_raw= as.raw(utf8ToInt(DPBtoPut))

# replace the DPB key in the binary file
# The key may appear twice so we make 2 replacements
vbaProject= replaceSequence(DPBtoReplace_raw, DPBtoPut_raw, vbaProject)
vbaProject= replaceSequence(DPBtoReplace_raw, DPBtoPut_raw, vbaProject)

# save the file and close connections with files
outputFile= file(pathInput, "wb")
writeBin(vbaProject, outputFile)
closeAllConnections()

endTime= Sys.time()
endTime- startTime


####################################### put manually the modified vbaProject.bin file into the zip folder
####################################### manually rename the .zip file to .xlsm and open it
####################################### Excel doesn't recognize the security anymore so you get a bunch of error messages "Invalid key", "System error 40230" .. Don't panic
####################################### Click "OK" as many times as needed
####################################### Open Tools > VBA Project Properties > Protection and enter a new password
####################################### Save the file, close it and re-open it
