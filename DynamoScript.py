#The inputs to this node will be stored as a list in the IN variables.

#Get inputs into Node

indexList = IN[0]

inCableReference = IN[1]
inSwbFrom = IN[2]
inSwbTo = IN[3]
#inSwbType = IN[]
inSwbLoad = IN[4]
#inSwbLoadScope = IN[]
inSwbPF = IN[5]
inCableLength = IN[6]
inCableSizeActiveconductors = IN[7]
inCableSizeNeutralconductors = IN[8]
#inCableSizeEarthingconductor = IN[]
inActiveConductormaterial = IN[9]
#inOfPhases = IN[]
inCableType = IN[10]
inCableInsulation = IN[11]
inInstallationMethod = IN[12]
#inCableAdditionalDerating = IN[]
#inSwitchgearTripUnitType = IN[]
inSwitchgearManufacturer = IN[13]
#inBusType = IN[]
#inBusChassisRating = IN[]
#inUpstreamDiversity = IN[]
#inIsolatorType = IN[]
#inIsolatorRating = IN[]
inProtectiveDeviceRating = IN[14]
inProtectiveDeviceManufacturer = IN[15]
#inProtectiveDeviceType = IN[]
#inProtectiveDeviceModel = IN[]
#inProtectiveDeviceOCRTripUnit = IN[]
#inProtectiveDeviceTripSetting = IN[]


#Initiate List Variables
cableReference = []
swbFrom = []
swbTo = []
swbType = []
swbLoad = []
swbLoadScope = []
swbPF = []
cableLength = []
cableSizeActiveconductors = []
cableSizeNeutralconductors = []
cableSizeEarthingconductor = []
activeConductormaterial = []
ofPhases = []
cableType = []
cableInsulation = []
installationMethod = []
cableAdditionalDerating = []
switchgearTripUnitType = []
switchgearManufacturer = []
busType = []
busChassisRating = []
upstreamDiversity = []
isolatorType = []
isolatorRating = []
protectiveDeviceRating = []
protectiveDeviceManufacturer = []
protectiveDeviceType = []
protectiveDeviceModel = []
protectiveDeviceOCRTripUnit = []
protectiveDeviceTripSetting = []


#Add header row
cableReference.append('Cable Reference')
swbFrom.append('SWB From')
swbTo.append('SWB To')
swbType.append('SWB Type')
swbLoad.append('SWB Load')
swbLoadScope.append('SWB Load Scope')
swbPF.append('SWB PF')
cableLength.append('Cable Length')
cableSizeActiveconductors.append('Cable Size - Active conductors')
cableSizeNeutralconductors.append('Cable Size - Neutral conductors')
cableSizeEarthingconductor.append('Cable Size - Earthing conductor')
activeConductormaterial.append('Active Conductor material')
ofPhases.append('# of Phases')
cableType.append('Cable Type')
cableInsulation.append('Cable Insulation')
installationMethod.append('Installation Method')
cableAdditionalDerating.append('Cable Additional De-rating')
switchgearTripUnitType.append('Switchgear Trip Unit Type')
switchgearManufacturer.append('Switchgear Manufacturer')
busType.append('Bus Type')
busChassisRating.append('Bus/Chassis Rating (A)')
upstreamDiversity.append('Upstream Diversity')
isolatorType.append('Isolator Type')
isolatorRating.append('Isolator Rating (A)')
protectiveDeviceRating.append('Protective Device Rating (A)')
protectiveDeviceManufacturer.append('Protective Device Manufacturer')
protectiveDeviceType.append('Protective Device Type')
protectiveDeviceModel.append('Protective Device Model')
protectiveDeviceOCRTripUnit.append('Protective Device OCR/Trip Unit')
protectiveDeviceTripSetting.append('Protective Device Trip Setting (A)')



#Populate list with values 

#Populate list with values 
for i , indexValue in enumerate(indexList):
	cableReference.append(inCableReference[i])
	swbFrom.append(inSwbFrom[i])
	swbTo.append(inSwbTo[i])
	swbType.append('')
	swbLoad.append(inSwbLoad[i])
	swbLoadScope.append('Local')
	swbPF.append(inSwbPF)
	cableLength.append(inCableLength[i])
	cableSizeActiveconductors.append(inCableSizeActiveconductors[i])
	cableSizeNeutralconductors.append(inCableSizeNeutralconductors[i])
	cableSizeEarthingconductor.append('')
	activeConductormaterial.append(inActiveConductormaterial[i])
	ofPhases.append('RWB')
	if(inCableType[i] == "4C+E"): #Convert to PowerCAD values
		cableType.append('Multi')
	elif(inCableType[i] == "4x1C+E"):
		cableType.append('SDI')
	elif(inCableType[i] == "BUS DUCT"):
		cableType.append('BD')
	else:
		cableType.append(inCableType[i])
	cableInsulation.append(inCableInsulation[i])

#	installationMethod.append(inInstallationMethod[i])
	if(inInstallationMethod[i] == "LADDER, SPACED"): #Convert to PowerCAD values
		installationMethod.append('L')
	elif(inInstallationMethod[i] == "PERFORATED TRAY, TOUCHING"):
		installationMethod.append('PT')
	elif(inInstallationMethod[i] == "IN UNDERGROUND WIRING ENCLOSURE"):
		installationMethod.append('C')
	else:
		installationMethod.append(inInstallationMethod[i])
	cableAdditionalDerating.append('')
	switchgearTripUnitType.append('Electronic')
	switchgearManufacturer.append(inSwitchgearManufacturer)
	busType.append('Bus Bar')
	busChassisRating.append('')
	upstreamDiversity.append('STD')
	isolatorType.append('None')
	isolatorRating.append('')
	protectiveDeviceRating.append(inProtectiveDeviceRating[i])
	protectiveDeviceManufacturer.append(inProtectiveDeviceManufacturer)
	protectiveDeviceType.append('')
	protectiveDeviceModel.append('')
	protectiveDeviceOCRTripUnit.append('')
	protectiveDeviceTripSetting.append('')
	
#combine lists into master list

OUT = zip(cableReference, swbFrom, swbTo, swbType, swbLoad, swbLoadScope, swbPF, cableLength, cableSizeActiveconductors, cableSizeNeutralconductors, cableSizeEarthingconductor, activeConductormaterial, ofPhases, cableType, cableInsulation, installationMethod, cableAdditionalDerating, switchgearTripUnitType, switchgearManufacturer, busType, busChassisRating, upstreamDiversity, isolatorType, isolatorRating, protectiveDeviceRating, protectiveDeviceManufacturer, protectiveDeviceType, protectiveDeviceModel, protectiveDeviceOCRTripUnit, protectiveDeviceTripSetting)



