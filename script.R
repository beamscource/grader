
# load library to read Excel files
library(gdata)

# to output tables into PDF
library(gridExtra)

# where are the Excel files stored?
# USE slash at the end
mainDir = "C:/Users/Eugen/Desktop/Noten/"
# list of all files
fileList = dir(mainDir, pattern = '.xlsx')

# to save names of data tables for later use
dataTablesList = ''

# read all excel files 
for (i in 1: length(fileList)){
  
  # full path to the file
  fullpath = paste(mainDir, fileList[i], sep = "", collapse = NULL)
  # name for the data table
  dataName = paste(strsplit(fileList[i], '.xlsx'), sep = '')

  # append names of data tables for later use
  if (i == 1){
  	dataTablesList = dataName
  	} else {
  	dataTablesList = c(dataTablesList, dataName)
  }

  # read data and assign it to a table for current file
  assign(dataName, read.xls(fullpath, sheet = 1, header = TRUE))
  
}

# type_col = ''
# assign()
# for (i in dataTablesList){
# 	# code pupils' names as factor
# 	#i$name_col = as.factor(i$name_col)
# 	# code the grade type as factor
# 	#i$type_col = as.factor(i$type_col)
# }

# to save names of data tables with final grades for later use
finalGradesdataTablesList = ''

for (i in 1:length(dataTablesList)){

	currentTable = dataTablesList[i]
	currentStudents = levels(get(currentTable)$Vorname)

	for (j in 1:length(currentStudents)){

		name = currentStudents[j]
		
		test = mean(get(currentTable)[get(currentTable)$Vorname == name & get(currentTable)$Typ == 's',]$Punkte, na.rm = T)

		classExam = mean(get(currentTable)[get(currentTable)$Vorname == name & get(currentTable)$Typ == 'ka',]$Punkte, na.rm = T)

		writtenGrade = round(test*0.4 + classExam*0.6, 2)

		oralGrade = round(mean(get(currentTable)[get(currentTable)$Vorname == name & get(currentTable)$Typ == 'm',]$Punkte, na.rm = T), 2)

		finalGrade = round((oralGrade + writtenGrade)/2, 2)
		
		# save final grades into a table
		if (j == 1){
			
			# create entry
			finalGradesTable = matrix(c(name, oralGrade, writtenGrade, finalGrade), nrow = 1)
			# assign column names and convert into a frame
			colnames(finalGradesTable) = c("Vorname", "Muendlich", "Schriftlich", "Endnote")
			finalGradesTable = as.data.frame(finalGradesTable)

		} else {
			# create next data point
			dataPoint = matrix(c(name, oralGrade, writtenGrade, finalGrade), nrow = 1)
			colnames(dataPoint) = c("Vorname", "Muendlich", "Schriftlich", "Endnote")
			dataPoint = as.data.frame(dataPoint)

			# add the data point to the table with final grades
			finalGradesTable = rbind(finalGradesTable, dataPoint)
		}
	}

	# rename the data frame class specific
	assign(paste(currentTable, '_Endnoten', sep = ''), finalGradesTable)
	rm(finalGradesTable)

	# append names of data tables with final grades for later use
  	if (i == 1){
  		finalGradesdataTablesList = paste(currentTable, '_Endnoten', sep = '')
  		} else {
  		finalGradesdataTablesList = c(finalGradesdataTablesList, paste(currentTable, '_Endnoten', sep = ''))
  	}
  	
}

# output data
for (currentTable in finalGradesdataTablesList){

	# if more than 20 students in class
	if (nrow(get(currentTable)) > 20) {

		pdf(file = paste(mainDir, currentTable, '.pdf', sep = ''))
		grid.table(get(currentTable)[1:20,])
		plot.new()
		grid.table(get(currentTable)[21:nrow(get(currentTable)),])
		dev.off()
		
		} else {

		pdf(file = paste(mainDir, currentTable, '.pdf', sep = ''))
		grid.table(get(currentTable))
		dev.off()
	}
}


