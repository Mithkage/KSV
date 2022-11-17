import clr
clr.AddReference('ProtoGeometry')
from Autodesk.DesignScript.Geometry import *
#The inputs to this node will be stored as a list in the IN variables.

#Get inputs into Node

indexList = IN[0]

inCableReference = IN[1]
inCableFrom = IN[2]


#Initiate List Variables
cableReference = []
cableFrom = []


cableReference.append('File Name')
cableFrom.append('Document No*')

cableReference.append('Cable Reference')
cableFrom.append('SWB From')






#empty column
for i , indexValue in enumerate(indexList):
	cableReference.append(inCableReference[i])
	cableFrom.append(inCableFrom[i])


#combine lists into master list


#Assign your output to the OUT variable.

OUT = zip(cableReference, cableFrom)

