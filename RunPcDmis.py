# example for running PC-DMIS with python
# works with python 2.7 & 3.7
# requires PC-DMIS installed on the system + and you may need to pip install pywin32
# PC-DMIS autmation documentation here: https://docs.hexagonmi.com/pcdmis/2020.2/en/helpcenter/mergedProjects/automationobjects/webframe.html
import win32com.client # uses the COM client to communicate

# connect - Note: this may open PC-DMIS if you don't have it open already
pcdmis = win32com.client.Dispatch('PCDLRN.Application') 

# example uses
pcdmis.Visible = True # make the window show

# load and run a program
program = 'C:\\Users\\Public\\Documents\\Hexagon\\PC-DMIS\\2020 R1\\test.PRG'
units = 1 # unittype: 0=mm, 1=inches
cmm = 'CMM1' # for multi-arm this could be CMM2, for example
probefile = '' # no need to set this for now
pcdmis.PartPrograms.CloseAll() # close any existing programs
pcdmis.PartPrograms.Add(program, units, cmm, probefile) # loads the program
didItRun = pcdmis.ActivePartProgram.EXECUTE # run it
print(didItRun)