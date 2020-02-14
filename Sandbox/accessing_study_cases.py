import powerfactory
app = powerfactory.GetApplication()



#Get the study case folder and its contents
aFolder=app.GetProjectFolder('study')
aCases=aFolder.GetContents('*.IntCase')
aCases=aCases[0]
#Counts the number of study cases
iCases=len(aCases)
if iCases == 0:
	app.PrintPlain('There are no study cases')
else:
	app.PrintPlain('Number of cases: {}'.format(iCases))
app.PrintPlain(aCases)
for aCase in aCases:
	app.PrintPlain(aCase)
	app.PrintPlain(aCase.loc_name)
	aCase.Activate