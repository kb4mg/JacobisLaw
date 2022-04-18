# This script to compute values for the Maximum Power Transfer Theorem known as Jabobi's Law
# Author: Martin Buehring
# April , 2022
# v1

# Program parameters
sourceVoltage = 100  # in volts , eg 100 volts
sourceResistance = 50 # in Ohms
loadResistance = 0

loadResistanceStartVal = 10  #in Ohms, but never zero or infinity
loadResistanceEndVal = 200  # in Ohms
loadResistanceStepVal = 5  #in Ohms, to add and sweep the range

loopCurrent = 0.0
loadVoltage = 0.0
loadPower = 0.0
maxPower = 0.0

# Functions
def printBanner(pgname):
    # Use a breakpoint in the code line below to debug your script.
    print(f'{pgname}')  # Press Ctrl+F8 to toggle the breakpoint.

def powerVal(i,v):
    return (i * v)

###############################################
## MAIN PROGRAM
###############################################

if __name__ == '__main__':
    printBanner('MPTT Calculation')
    while loadResistance < loadResistanceEndVal:
        loadResistance = loadResistance + loadResistanceStepVal
        print('* At Load = ',loadResistance)
        loopCurrent = sourceVoltage/(loadResistance + sourceResistance)  # add load plus source resistances
        print(f' ','LoopCurrent:',end="")
        print('%.3f'%loopCurrent, 'A') # Format current to three decimal places
        loadVoltage = loopCurrent * loadResistance
        print(f' ','Load Voltage:',end="")
        print('%.3f'%loadVoltage, 'V')
        loadPower = powerVal(loopCurrent, loadVoltage)
        print(f' ','Load Power:', end="")
        print('%.3f'%loadPower, 'W with Load Resistance:', loadResistance,'\n')
        if (maxPower < loadPower):
            maxPower = loadPower
print('Max Power was ', maxPower)