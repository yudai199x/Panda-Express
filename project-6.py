#########################################
# Iccs(BiCS5 512Gb D3 4Plane)2022/03/14 #
#########################################

#2020.12.22 MT技&五FG 内容変更(ExPBAに対応)
#2021.01.05 MT技 内容変更(MCM対応)
#2021.12.20 MT技 内容変更(4Plane対応)
#2022.03.15 MT技 内容変更(Defualt ParaSet追加、Mask方式、ParaLoopごとにPowerOffしてリセット)

PRINT_FLAG(0)




#入力箇所*******************************************************

Project = "xxxx"		#対象Project名

#Dansu	=4			# PKG内のNANDの枚数(単体ではCE数)を入力下さい
#MCM	=1			# PKG内のMCM数(単体ではCEあたりのChip数)を入力下さい

#MCM品の設定	# 4st4CE-BGA132 / 8st4CE-BGA132 / 16st8CE-BGA272 / 16st-X5C / stuck=CE
#Dansu	=4		#       4       /       4       /        8       /     8    /  1,2,4,8
#MCM 	=1		#       1       /       2       /        2       /     2    /  1 

Column	= 0x0000		#必要であれば入力下さい
Page	= 0x000			#必要であれば入力下さい
#Block	= 0x0000		#必要であれば入力下さい
Cbadd	= 0x1000		# MCM対応 -> 製品によって変更の必要有り(BiCS5 512Gb :0x1000,BiCS5 1Tb:0x2000)
Param	= "VDD"			#変数を切り替えて下さいALL:全計測,VDD,VDDA,VREF,VFOUR,VDDSA_PB0,VDDSA_PB1,VDDSA_PB2,VDDSA_PB3,IREF,
F_VDD	= 0x00			# VDD Default
F_VDDA	= 0x00			# VDDA Default
F_VREF	= 0x00			# VREF Default
F_VFOUR	= 0x00			# VFOUR Default
F_VDDSA	= 0x00			# VDDSA Default
F_IREF	= 0x00			# IREF Default
MASK	= 0x00			# MASK Bit
BDAC	= 0x00			#Begin_Dac
EDAC	= 0x1F			#End_Dac
SDAC	= 0x02			#Step_Dac
MDAC	= 0x00			#Measure_Dac

Deep_STBY=1			#Deep_STBY 1=enable , 0=disable
HOLDFLAG= 0			#発光解析用

#入力終り********************************************************


Block	= Cbadd*Chip
if Deep_STBY==1:
	PRINT("Deep_STBY-enable")
else:
	PRINT("Deep_STBY-disable")
	
#	CURRENT2="%s依存,ChipNo=,%s\n" % (Param,CE)			#★CURRENT2="%s依存,CE=,%s\n" % (Param,CE)から変更
CURRENT2="%s依存,CE=%s,Chip=%s\n" % (Param,CE,Chip)
PRINT(CURRENT2)
if Param == "VDD":				#0x00_1.60V , 0x10_2.00V 以降BurnIn Voltage
	Param2=["VDD"]
elif Param == "VDDA":				#0x00_1.90V , 0x07_2.60V
	Param2=["VDDA"]	
elif Param == "VREF":				#0x00_1.04V , 0x1F_1.35V
	Param2=["VREF"]
elif Param == "VFOUR":				#0x00_3.48V , 0x1F_4.31V
	Param2=["VFOUR"]
elif Param == "VDDSA_PB0":			#0x00_1.90V , 0x07_2.60V
	Param2=["VDDSA_PB0"]
	Block	= 0x0100 + Cbadd*Chip
elif Param == "VDDSA_PB1":			#0x00_1.90V , 0x07_2.60V
	Param2=["VDDSA_PB1"]
	Block	= 0x0101 + Cbadd*Chip
elif Param == "VDDSA_PB2":			#0x00_1.90V , 0x07_2.60V
	Param2=["VDDSA_PB2"]
	Block	= 0x0102 + Cbadd*Chip
elif Param == "VDDSA_PB3":			#0x00_1.90V , 0x07_2.60V
	Param2=["VDDSA_PB3"]
	Block	= 0x0103 + Cbadd*Chip
elif Param == "IREF":				#designmanual参照(501頁)
	Param2=["IREF"]
elif Param == "ALL":
	Param2=["VDD","VDDA","VREF","VFOUR","VDDSA_PB0","VDDSA_PB1","VDDSA_PB2","VDDSA_PB3","IREF"]

for Param3 in Param2:
	POWER_ON2(CE,63)						#★POWER_ON2(CE,PO2SET)から変更
	#SELECT_CHIP(CE)

	CURRENT3="%s依存\n" % (Param3)
	PRINT(CURRENT3)

	if Param3 == "VDD":			#0x00_1.60V , 0x10_2.00V 以降BurnIn Voltage	
		Para=0x024#D9	
		BDAC	= 0x00
		EDAC	= 0x10
		SDAC	= 0x01
		MASK	= 0x1F
	elif Param3 == "VDDA":			#0x00_1.90V , 0x07_2.60V
		Para=0x0DE
		BDAC	= 0x00
		EDAC	= 0x0F#1B
		SDAC	= 0x01	
		MASK	= 0x0F
	elif Param3 == "VREF":			#0x00_1.04V , 0x1F_1.35V
		Para=0x028#E2	
		BDAC	= 0x00
		EDAC	= 0x1F
		SDAC	= 0x02
		MASK	= 0x1F
	elif Param3 == "VFOUR":			#0x00_3.48V , 0x0F_4.31V
		Para=0x0DA#28	
		BDAC	= 0x00
		EDAC	= 0x0F#1F
		SDAC	= 0x02	
		MASK	= 0x0F
	elif Param3 == "VDDSA_PB0":		#0x00_1.90V , 0x07_2.60V
		Para=0x026			#VDDSA
		Block	= 0x0100 + Cbadd*Chip
		BDAC	= 0x00
		EDAC	= 0x1F
		SDAC	= 0x02
		MASK	= 0x1F
	elif Param3 == "VDDSA_PB1":		#0x00_1.90V , 0x07_2.60V
		Para=0x026			#VDDSA
		Block	= 0x0101 + Cbadd*Chip
		BDAC	= 0x00
		EDAC	= 0x1F
		SDAC	= 0x02
		MASK	= 0x1F
	elif Param3 == "VDDSA_PB2":		#0x00_1.90V , 0x07_2.60V
		Para=0x026			#VDDSA
		Block	= 0x0102 + Cbadd*Chip
		BDAC	= 0x00
		EDAC	= 0x1F
		SDAC	= 0x02
		MASK	= 0x1F
	elif Param3 == "VDDSA_PB3":		#0x00_1.90V , 0x07_2.60V
		Para=0x026			#VDDSA
		Block	= 0x0103 + Cbadd*Chip
		BDAC	= 0x00
		EDAC	= 0x1F
		SDAC	= 0x02
		MASK	= 0x1F
	elif Param3 == "IREF":			#designmanual参照(501頁)
		Para=0x07F#DB
		Block	= 0x0100 + Cbadd*Chip
		BDAC	= 0x00
		EDAC	= 0x3F#1F
		SDAC	= 0x02
		MASK	= 0x3F
	
	if HOLDFLAG == 1:
		BDAC = EDAC = HDAC

	#SET_PARA_BITMASK(0x55, 0x024, 0x1F, F_VDD)			# ROM値のときは行全体をコメントアウト
	#SET_PARA_BITMASK(0x55, 0x0DE, 0x0F, F_VDDA)		# ROM値のときは行全体をコメントアウト
	#SET_PARA_BITMASK(0x55, 0x028, 0x1F, F_VREF)		# ROM値のときは行全体をコメントアウト
	#SET_PARA_BITMASK(0x55, 0x0DA, 0x0F, F_VFOUR)		# ROM値のときは行全体をコメントアウト
	#SET_PARA_BITMASK(0x55, 0x026, 0x1F, F_VDDSA)		# ROM値のときは行全体をコメントアウト
	#SET_PARA_BITMASK(0x55, 0x07F, 0x3F, F_IREF)		# ROM値のときは行全体をコメントアウト

	for DAC in range(BDAC,EDAC+1,SDAC):
		FF_RESET();
		#CURRENT=""
		TEST_MODE()
		if Deep_STBY==1:
			COM(0x93)
		else:
			COM(0x94)

		#ADDSEQ("COM",0x55); ADDSEQ("ADD",Para); ADDSEQ("DATA",DAC)		
		#EXESEQ()	
		SET_PARA_BITMASK(0x55, Para, MASK, DAC)	

		DELAY(10)

		ADDSEQ("COM",0x00);
		ADDSEQ("ADDX5", Column, Page, Block)
		ADDSEQ("COM",0x30);
		EXESEQ()
	
		if HOLDFLAG == 1:
			MSGBOX("Click")
		
		if Param=="ALL":

			MEAS_VCC(3.3,1)
			CURRENT1   ="Iccs_%s,0x%02X, %s ,uA" % (Param3,DAC,VALUE("CURRENT"))
				#	CURRENT2 +="Iccs( DAC : 0x%02X )=%s uA\n" % (DAC,VALUE("CURRENT"))
				PRINT(CURRENT1)
			else:
				MEAS_VCC(3.3,1)
				CURRENT1   ="Iccs_%s,0x%02X, %s ,uA" % (Param,DAC,VALUE("CURRENT"))
				#	CURRENT2 +="Iccs( DAC : 0x%02X )=%s uA\n" % (DAC,VALUE("CURRENT"))
				PRINT(CURRENT1)

		POWER_OFF2();

LogFile="%s_Iccs(ChipNo%s)%s依存.txt" % (Project,CE,Param)		#★LogFile="%s_Iccs(CE%s)%s依存.txt" % (Project,CE,Param)
fp = open(DataDir + "\\" + LogFile, "w")
fp.write( CURRENT2 )
fp.close()

