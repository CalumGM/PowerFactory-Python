import powerfactory

app = powerfactory.GetApplication()
app.ClearOutputWindow()

# app.SetGuiUpdateEnabled(1)


'''
Get the SearchObject to work

Run a load flow then save it into results file...all in python
'''

active_project = app.GetActiveProject()
project_name = active_project.loc_name
study_cases = active_project.GetContents('Study Cases', 0)[0]
study_case = study_cases.GetContents('*.IntCase', 0)[0]
res = study_case.GetContents('Quasi-Dynamic Simulation AC', 0)

res = res[0]

#loading the results file
res.Load()

#Calling the first line in the project
lines = app.GetCalcRelevantObjects('*.ElmLne')
line = lines[0]

res.AddVariable(line, 'c:loading') 

number_of_rows = res.GetNumberOfRows()
number_of_columns = res.GetNumberOfColumns()

app.PrintPlain('ElmRes has {} rows and {} columns'.format(number_of_rows, number_of_columns))
obj = res.GetObject(8) # each variable(column) is an attribute of an object
app.PrintPlain('Object type of column 3 is {}'.format(obj))  # just a test

variables= []
for i in range(number_of_columns): # print all the variable names
	variable = res.GetVariable(i)
	variables.append(variable)
','.join(variables)
app.PrintPlain(variables)

row = []
for i in range(number_of_rows):  # print in table-ish format
	for j in range(number_of_columns):
		element = res.GetValue(i, j)[1] 
		row.append(element)
	','.join(row)
	app.PrintPlain(row)

