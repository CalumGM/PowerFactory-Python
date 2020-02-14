# PowerFactory Importing
import sys
import powerfactory as pf


app = pf.GetApplication()  # Start PowerFactory in engine mode
app.PrintPlain('something works')
project = app.GetActiveProject()
ldf = app.GetFromStudyCase('ComLdf')
ldf.Execute()
lines = app.GetCalcRelevantObjects('*.ElmLne')  # Get the list of lines contained in the project
for line in lines:
    # line_type = line.ElmCabsys.GetLineCable()
    name = line.loc_name  # Get the name of the line
    value = line.GetAttribute('c:loading')
    app.PrintPlain('Loading of the Line: {} -- {}'.format(name, value))  # Print results

