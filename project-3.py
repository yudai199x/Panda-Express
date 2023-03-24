# 0_USGD_Open_SCR(Multi)Random2.py
# Killer BlockをRead後、Block 0/1のWL0 L/M/UをMulti Plane TLC Normal ReadでFBC
# FBC期待値はRandom seed
# 開始Blockから16Block連続Readする

Project	= "YQR-xxxxx"
Chip	= 0
Block_S	= 0x350			# Killer Block 開始アドレス(16Block Readします)
ADDX3	= 0			# 1:60h + ADDX3, 0:00h + ADDX5

FBC_Col	= 18336  		# Col数(max:18336)

Param	= 0			# 1:Parameter Set, Other:Not use
Add	= 0xd9
Mask	= 0x1F
Dac2	= 0x00			# Paramet評価のDAC値

File	= 0  			# 1:ログをファイル出力する, 0:画面表示のみ
##### 必要な設定はここまで #####
Block1	= 0x0000  		# Plane0 BLK
Block2	= 0x0001  		# Plane1 BLK
S_Page=0x00
E_Page=0x01
step=1

### Bitカウント用テーブル ###
FC_TBL = [0,1,1,2,1,2,2,3,1,2,2,3,2,3,3,4, \
	  1,2,2,3,2,3,3,4,2,3,3,4,3,4,4,5, \
	  1,2,2,3,2,3,3,4,2,3,3,4,3,4,4,5, \
	  2,3,3,4,3,4,4,5,3,4,4,5,4,5,5,6, \
	  1,2,2,3,2,3,3,4,2,3,3,4,3,4,4,5, \
	  2,3,3,4,3,4,4,5,3,4,4,5,4,5,5,6, \
	  2,3,3,4,3,4,4,5,3,4,4,5,4,5,5,6, \
	  3,4,4,5,4,5,5,6,4,5,5,6,5,6,6,7, \
	  1,2,2,3,2,3,3,4,2,3,3,4,3,4,4,5, \
	  2,3,3,4,3,4,4,5,3,4,4,5,4,5,5,6, \
	  2,3,3,4,3,4,4,5,3,4,4,5,4,5,5,6, \
	  3,4,4,5,4,5,5,6,4,5,5,6,5,6,6,7, \
	  2,3,3,4,3,4,4,5,3,4,4,5,4,5,5,6, \
	  3,4,4,5,4,5,5,6,4,5,5,6,5,6,6,7, \
	  3,4,4,5,4,5,5,6,4,5,5,6,5,6,6,7, \
	  4,5,5,6,5,6,6,7,5,6,6,7,6,7,7,8]

def READ(WLstr,LMU,PageL,Block1,Block2):
	FailBit=""
	PageB = PageL
	if LMU == "L":
		PreCOM = 0x01
	if LMU == "M":
		PreCOM = 0x02
	if LMU == "U":
		PreCOM = 0x03

	Badd1=bin(Block1)[2:].zfill(14)
	Badd2=bin(Block2)[2:].zfill(14)
	Padd=bin(WLstr)[2:].zfill(9)

	add3=int(Padd[1:9],2)
	add41=int(Badd1[7:14]+Padd[0:1],2)
	add51=int(Badd1[0:7],2)
	add42=int(Badd2[7:14]+Padd[0:1],2)
	add52=int(Badd2[0:7],2)

	if ADDX3==1:
		ADDSEQ("COM",PreCOM)
		ADDSEQ("COM",0x60); ADDSEQ("ADD",add3); ADDSEQ("ADD",add41); ADDSEQ("ADD",add51) 
		ADDSEQ("COM",PreCOM)
		ADDSEQ("COM",0x60); ADDSEQ("ADD",add3); ADDSEQ("ADD",add42); ADDSEQ("ADD",add52) 
		ADDSEQ("COM",0x30); ADDSEQ("MATCH"); EXESEQ()
	else:
		ADDSEQ("COM",PreCOM)
		ADDSEQ("COM",0x00)
		ADDSEQ("ADD",0x0); ADDSEQ("ADD",0x0); ADDSEQ("ADD",add3); ADDSEQ("ADD",add41); ADDSEQ("ADD",add51)
		ADDSEQ("COM",0x32); ADDSEQ("WAIT",1)
		ADDSEQ("COM",PreCOM)
		ADDSEQ("COM",0x00)
		ADDSEQ("ADD",0x0); ADDSEQ("ADD",0x0); ADDSEQ("ADD",add3); ADDSEQ("ADD",add42); ADDSEQ("ADD",add52)
		ADDSEQ("COM",0x30); ADDSEQ("MATCH"); EXESEQ()

	ADDSEQ("COM",0x05)
	ADDSEQ("ADD",0x00); ADDSEQ("ADD",0x00); ADDSEQ("ADD",add3); ADDSEQ("ADD",add41); ADDSEQ("ADD",add51) 
	ADDSEQ("COM",0xE0)
	ADDSEQ("WAIT",1)
	ADDSEQ("READ_P",0x00,FBC_Col)
	EXESEQ()
	READ_DATA1 = VALUE("READ_DATA").split(" ")

	ADDSEQ("COM",0x05)
	ADDSEQ("ADD",0x00); ADDSEQ("ADD",0x00); ADDSEQ("ADD",add3); ADDSEQ("ADD",add42); ADDSEQ("ADD",add52) 
	ADDSEQ("COM",0xE0)
	ADDSEQ("WAIT",1)
	ADDSEQ("READ_P",0x00,FBC_Col)
	EXESEQ()
	READ_DATA2 = VALUE("READ_DATA").split(" ")

	FBC1 = 0; total1=0; SECT1=[0]*16; max1=0; j1=0
	FBC2 = 0; total2=0; SECT2=[0]*16; max2=0; j2=0
	for Byte in range( len(READ_DATA1) ):
		Fail = ( (int(READ_DATA1[Byte],16) ^ int(EXP_DATA1[PageB][Byte],16)) & 0xFF)
		FBC1 += FC_TBL[Fail]
		if (Byte%1146==1145 or Byte==len(READ_DATA1)-1):
			SECT1[j1] = FBC1 
			total1 += FBC1
			FBC1 = 0
			j1 += 1
	for k in range(16):		
		if (max1 <= SECT1[k]):	max1 = SECT1[k]

	for Byte in range( len(READ_DATA2) ):
		Fail = ( (int(READ_DATA2[Byte],16) ^ int(EXP_DATA2[PageB][Byte],16)) & 0xFF)
		FBC2 += FC_TBL[Fail]
		if (Byte%1146==1145 or Byte==len(READ_DATA2)-1):
			SECT2[j2] = FBC2 
			total2 += FBC2
			FBC2 = 0
			j2 += 1
	for k in range(16):		
		if (max2 <= SECT2[k]):	max2 = SECT2[k]

	if PageB%3 == 0:
		WLstrB = PageB/3
		LMUB = "L"
	if PageB%3 == 1:
		WLstrB = (PageB-1)/3
		LMUB = "M"
	if PageB%3 == 2:
		WLstrB = (PageB-2)/3
		LMUB = "U"

	FailBit0 =" %3d|  %s  |%6d|%5d|%6d|%5d" % (WLstrB,LMUB,total1,max1,total2,max2)
	PRINT(FailBit0)
	FailBit ="%s\n" % FailBit0
	return FailBit

# FBC期待値取得

EXP_DATA1 = [0]*LOGICAL_PAGE
EXP_DATA2 = [0]*LOGICAL_PAGE
for Page in range(S_Page,E_Page,step):
	CRC32(0xAA, Page*3)
	EXP_DATA1[Page*3] = VALUE("CRC32_DATA").split(",")
	EXP_DATA2[Page*3] = VALUE("CRC32_DATA").split(",")
	CRC32(0xAA, Page*3+1)
	EXP_DATA1[Page*3+1] = VALUE("CRC32_DATA").split(",")
	EXP_DATA2[Page*3+1] = VALUE("CRC32_DATA").split(",")
	CRC32(0xAA, Page*3+2)
	EXP_DATA1[Page*3+2] = VALUE("CRC32_DATA").split(",")
	EXP_DATA2[Page*3+2] = VALUE("CRC32_DATA").split(",")

for Killer in range(Block_S,(Block_S+16),1):
	POWER_ON2(Chip,15)

	if Param == 1:
		SET_PARA_BITMASK(0x55,Add,Mask,Dac2)

	# Page FBC(Multi Plane Normal Read)
	PRINT("Chip%d" % Chip)

	FailBit ="Fail Bit Count BLK:%04Xh-%04Xh\n" % (Block1,Block2)
	FailBit+="Killer Block:0x%04X\n" % Killer
	PRINT(FailBit)

	Data1 ="    |     | FBC  |Seg  |FBC   |Seg  \n"
	Data1+="Page|L/M/U|Plane0|Worst|Plane1|Worst\n"
	Data1+="----+-----+------+-----+------+-----"

	PRINT("%s" % Data1)
	FailBit+="%s\n" % Data1

	PageK=0x0000
	BaddK=bin(Killer)[2:].zfill(14)
	PaddK=bin(PageK)[2:].zfill(9)
	add3K=int(PaddK[1:9],2)
	add4K=int(BaddK[7:14]+PaddK[0:1],2)
	add5K=int(BaddK[0:7],2)
	ADDSEQ("COM",0x01)
	ADDSEQ("COM",0x00)
	ADDSEQ("ADD",0x0); ADDSEQ("ADD",0x0); ADDSEQ("ADD",add3K); ADDSEQ("ADD",add4K); ADDSEQ("ADD",add5K)
	ADDSEQ("COM",0x30); ADDSEQ("MATCH")

	for Page3 in range(S_Page,E_Page,step):
		WLstr = Page3
		for LMU in ["L","M","U"]:

			if LMU == "L":
				PageL = WLstr*3
			elif LMU == "M":
				PageL = WLstr*3+1
			elif LMU == "U":
				PageL = WLstr*3+2
			result = READ(WLstr,LMU,PageL,Block1,Block2)		# Multi Plane Cache Read + Fail Bit Count
			FailBit+=result
	POWER_OFF2()

	import datetime
	Now = datetime.datetime.today()
	LogFile = "%s_USGD_Open_SCR(Multi)Chip%d_Killer-0x%04X.txt" % (Project,Chip,Killer)

	if File == 1:
		fp = open(DataDir + "\\" + LogFile, "w")		# ファイルの出力先指定
		fp.write( FailBit )					# ファイルに出力するデータ指定
		fp.close()	







