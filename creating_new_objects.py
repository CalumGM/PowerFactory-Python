'''
Needs some special files. Just skipped this
'''
import powerfactory
app = powerfactory.GetApplication()

oFold = app.GetFromStudyCase('IntEVt')
app.PrintPlain(oFold)

Script = apPp.GetCurrentScript()
Contents = Script.GetContents()
SC_Event = Contents[0]

oFold.AddCopy(SC_Event)

EventSet = oFold.GetContents('*EvtShc')
app.PrintPlain(EventSet)
OEvent = EventSet[0][0]
OEvent.time = 1