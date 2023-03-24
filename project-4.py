# 0_ICCop_Read(BiCS5)USGD_Killer_Param.py

Project	= "YQR-xxxxx"
Chip	= 0
Block	= 0x0351   		# Killer BLK(Addr. >= 0x0002)
Page	= 0x000

Add	= 0x1F			# Parameter Address
Mask	= 0x7F			# Bit Mask
Start	= 0x00			# Start DAC
End	= 0x7F			# End DAC
Step	= 0x08			# DAC Step
MDAC	= 0x00

Flag_A2h = 0			# Special SLC のON/OFF(1:ON 0:OFF)
Flag_1Eh = 1			# All Block Non Select(1:ON 0:OFF) 
Flag_2Ah = 0			# All Plane  Select(1:ON 0:OFF) 
Flag_3Eh = 0			# All CG Driver Select のON/OFF(1:ON 0:OFF)
Flag_4Eh = 0			# All CG Driver Non Select のON/OFF(1:ON 0:OFF)
Flag_5Eh = 0			# All CG Driver Disable のON/OFF(1:ON 0:OFF)
Flag_DBh = 0			# SG All Nonselect

HVSW 	= 0			# 1:HVSW disable, Other:Enable
BLS	= 1			# 1:BLS Disable, Other:Enable
HOLDFLAG= 0

if HOLDFLAG == 1:
	Start = End = MDAC

##### 必要な設定はここまで #####

CURRENT ="%s Chip%d BLK0x%04X ICCop\n" % (Project,Chip,Block)
CURRENT+="Option:A2h=%d, 1Eh=%d,3Eh=%d, 4Eh=%d, 5Eh=%d, DBh=%d, 2Ah=%d\n\n" % (Flag_A2h,Flag_1Eh,Flag_3Eh,Flag_4Eh,Flag_5Eh,Flag_DBh,Flag_2Ah)
CURRENT+="DAC , Icc(uA)\n"
PRINT("DAC , Icc(uA)") 

Row = ( Block << 9 ) | ( Page << 0 )
add3 =  Row & 0x0000ff  >> 0
add4 = (Row & 0x00ff00) >> 8
add5 = (Row & 0xff0000) >> 16

Block1=Block%2		# Killer BLKのPlaneをReadする(電流計測するはこのBlock)
Row1 = ( Block1 << 9 ) | ( Page << 0 )
add31 =  Row1 & 0x0000ff  >> 0
add41 = (Row1 & 0x00ff00) >> 8
add51 = (Row1 & 0xff0000) >> 16
##### ここまで
POWER_ON2(Chip,15)

for DAC in range(Start,End+1,Step):
	ADDSEQ("TEST_MODE")
	ADDSEQ("COM",0x55); ADDSEQ("ADD",0x03); ADDSEQ("DATA",Mask)				# Set Bit Mask
	ADDSEQ("COM",0x55); ADDSEQ("ADD",0x01); ADDSEQ("DATA",0x06)
	if Add >= 0x100:	ADDSEQ("COM",0x55); ADDSEQ("ADD",0xFF); ADDSEQ("DATA",0x01)
	ADDSEQ("COM",0x00); ADDSEQ("ADD",Add&0xFF); ADDSEQ("COM",0xEA); ADDSEQ("WAIT",1)
	ADDSEQ("COM",0x57); ADDSEQ("ADD",Add&0xFF); ADDSEQ("DATA",DAC)				# Set Parameter Dac
	if Add >= 0x100:	ADDSEQ("COM",0x55); ADDSEQ("ADD",0xFF); ADDSEQ("DATA",0x00)

	ADDSEQ("COM",0x00)
	ADDSEQ("ADD",0x00); ADDSEQ("ADD",0x00); ADDSEQ("ADD",add3); ADDSEQ("ADD",add4); ADDSEQ("ADD",add5); 
	ADDSEQ("COM",0x30); ADDSEQ("MATCH")							# Normal Read(Killer BLK)

	if Flag_1Eh == 1: ADDSEQ("COM", 0x1E)							# All Block non-select
	if Flag_2Ah == 1: ADDSEQ("COM", 0x2A)							# All Plane Select
	if Flag_3Eh == 1: ADDSEQ("COM", 0x3E)							# All CG Select
	if Flag_4Eh == 1: ADDSEQ("COM", 0x4E)							# All CG Non Select
	if Flag_5Eh == 1: ADDSEQ("COM", 0x5E)							# All CG Disable
	if Flag_DBh == 1: ADDSEQ("COM", 0xDB);	ADDSEQ("ADD",0x00)				# SG ALL Non Select
	if HVSW == 1:	ADDSEQ("COM", 0x22)							# HVSW disable
	if BLS == 1:	ADDSEQ("COM", 0xB8)							# BLS disable

	ADDSEQ("COM", 0x4F)									# External Timing Read	
	if Flag_A2h == 1: ADDSEQ("COM", 0xA2)							# Special SLC
	ADDSEQ("COM",0x00)
	ADDSEQ("ADD",0x00); ADDSEQ("ADD",0x00); ADDSEQ("ADD",add31); ADDSEQ("ADD",add41); ADDSEQ("ADD",add51); 
	ADDSEQ("COM",0xAE)									# Test Read(BLK 0x0 or 0x1)
	EXESEQ()

	if HOLDFLAG == 1:
		MSGBOX("Click")

	MEAS_VCC(3.3,3)
	ICC=float(GET_RBUF("CURRENT"))
	DATA ="%02Xh, %d" % (DAC,ICC)
	PRINT(DATA)
	CURRENT +="%s\n" % DATA

	FF_RESET()
POWER_OFF2()

if BLS == 1:
	LogFile = "%s_ICCop(Read_USGD-Killer)Chip%d_0x%04X_Page%02Xh(BLS)%02Xh依存.txt" % (Project,Chip,Block,Page,Add)
elif HVSW == 1:
	LogFile = "%s_ICCop(Read_USGD-Killer)Chip%d_0x%04X_Page%02Xh(HVSW)%02Xh依存.txt" % (Project,Chip,Block,Page,Add)
else:
	LogFile = "%s_ICCop(Read_USGD-Killer)Chip%d_0x%04X_Page%02Xh_%02Xh依存.txt" % (Project,Chip,Block,Page,Add)
fp = open(DataDir + "\\" + LogFile, "w")
fp.write( CURRENT )
fp.close()


