# T6RN1 SA SDL
# 23/02/14 PEgi isozumi

#-----------------------------
from TLIBC1 import *
import numpy as np
import datetime

dt = datetime.datetime.today()
t_stamp = "{:02d}{:02d}-{:02d}{:02d}".format(dt.month,dt.day,dt.hour,dt.minute)

def FBC_single(mpat, plane, send, f_sde, f_vddsa, f_vdd):
    mode_fbm = 0
    mode_fbc = 1
    ret = send-1

    pc = 0x000

    SET_REG("REG_PC", pc)
    SET_REG("REG_XH", 0x00)
    SET_REG("REG_YH", 0x00)
    SET_REG("REG_ZH", plane)

    SET_REG("REG_D1", f_sde)
    SET_REG("REG_D4B", f_vddsa)
    SET_REG("REG_D3B", f_vdd)

    SET_REG("REG_DT4", send)
    SET_REG("REG_DT5", ret)

    SET_REG("REG_ILR2", PAGESIZE-2)

    TLIBC1.START_MPAT_FBC(pc, mpat, 1) # wait 0
    PRINT("\nFBC(Dec) : {}\n".format(VALUE("READ_DATA")))

    TLIBC1.STOP_MPAT() # 必須


def str2int(data):
    return int(data)

def analog_sdl(mpat, plane, send, f_sde, f_vddsa, f_vdd, pxl_row, pxl_col):
    if pxl_col==512:
        c_offset = 10 * 2
    elif pxl_col==256:
        c_offset = 6 * 2
    else:
        c_offset = 5 * 2

    sdl_col = pxl_col + c_offset # Col Pixel in SDL

    if pxl_row==512: #Row Cnt / Loop
        sdl_row = 7
    elif pxl_row==256:
        sdl_row = 15
    else:
        sdl_row = 30

    r_offset = 2 * 2
    loop_cnt = (pxl_row+r_offset)//sdl_row+1 # Loop Cnt / Image

    PRINT("\nLoop Cnt : {}\n".format(loop_cnt))

    ret = send-1
    data_org = np.zeros(0, dtype=np.uint32)

    MSGBOX("レーザ始動後OK押下")

    pc = 0x020
    SET_REG("REG_PC", pc)
    SET_REG("REG_XH", 0x00)
    SET_REG("REG_YH", 0x00)
    SET_REG("REG_ZH", plane)

    SET_REG("REG_D1", f_sde)
    SET_REG("REG_D4B", f_vddsa)
    SET_REG("REG_D3B", f_vdd)

    SET_REG("REG_DT4", send)
    SET_REG("REG_DT5", ret)

    SET_REG("REG_ILR2", PAGESIZE-2)
    SET_REG("REG_ILR3", sdl_col-2)
    SET_REG("REG_ILR4", sdl_row-2)
    SET_REG("REG_ILR5", loop_cnt-2)

    TLIBC1.START_MPAT_FBC(pc, mpat, sdl_col*sdl_row) # wait 0
    str_data = VALUE("READ_DATA")
    PRINT(str_data)
    list_data = np.array(map(str2int, str_data.split(","))) #Python2
    PRINT(str(len(list_data)))
    data_org = np.append(data_org,list_data)

    for i in range(loop_cnt):
        TLIBC1.RESUME_MPAT_FBC(sdl_col*sdl_row) # wait 0 の次の行から再開。Wait 0 でPythonの処理に戻る。
        str_data = VALUE("READ_DATA")
        PRINT(str_data)
        list_data = np.array(map(str2int, str_data.split(","))) #Python2
        PRINT(str(len(list_data)))
        data_org = np.append(data_org,list_data)

    TLIBC1.STOP_MPAT() # 必須

    image_row = (loop_cnt+1)*sdl_row
    PRINT(str(image_row))


    data_ary = data_org.reshape([image_row, sdl_col])
    col_s = c_offset // 2
    col_e = -1*col_s
    row_s = r_offset // 2
    row_e = pxl_row+r_offset // 2

    result = data_ary[row_s:row_e, col_s:col_e]

    return result
#-----------------------------

def Single_DOUT(mpat, col, plane, send, f_sde, f_vddsa, f_vdd):
    mode_fbm = 0
    mode_fbc = 1

    ret = send-1

    pc = 0x040
    SET_REG("REG_PC", pc)
    SET_REG("REG_XH", col)
    SET_REG("REG_YH", 0x00)
    SET_REG("REG_ZH", plane)

    SET_REG("REG_D1", f_sde)
    SET_REG("REG_D4B", f_vddsa)
    SET_REG("REG_D3B", f_vdd)

    SET_REG("REG_DT4", send)
    SET_REG("REG_DT5", ret)

    MEAS_MPAT(mpat, mode_fbm)
    BasicLib.GET_DUC(2)
    str_data = GET_RBUF("READ_DATA")
    PRINT("\nRead Result(Hex) : {}\n".format(str_data))


def DIN_DOUT_SDL(pat, col, plane, send, f_sde, f_vddsa, f_vdd):
    ret = send-1

    pc = 0x060
    SET_REG("REG_PC", pc)
    SET_REG("REG_XH", col)
    SET_REG("REG_YH", 0x00)
    SET_REG("REG_ZH", plane)

    SET_REG("REG_D1", f_sde)
    SET_REG("REG_D4B", f_vddsa)
    SET_REG("REG_D3B", f_vdd)

    SET_REG("REG_DT4", send)
    SET_REG("REG_DT5", ret)

    TLIBC1.START_MPAT_ASYNC(pc, pat)
    BasicLib.GET_DUC(2)
    str_data = GET_RBUF("READ_DATA")
    PRINT("\nRead Result(Hex) : {}\n".format(str_data))

def SA_EMS(pat, send):
    ret = send-1

    pc = 0x090
    SET_REG("REG_PC", pc)

    SET_REG("REG_DT4", send)
    SET_REG("REG_DT5", ret)

    TLIBC1.START_MPAT_ASYNC(pc, pat)
    BasicLib.GET_DUC(2)
    str_data = GET_RBUF("READ_DATA")
    PRINT("\nRead Result(Hex) : {}\n".format(str_data))

#-----------------------------
cache = {\
    "XDL_Reset_FF" : 0x0c,
    "XDL_Reset_00" : 0x0d,

    "A2X": 0x01,
    "X2A": 0x02,

    "B2X": 0x03,
    "X2B": 0x04,

    "C2X": 0x05,
    "X2C": 0x06,

    "T2X": 0x07,
    "X2T": 0x08,
    }
#-----------------------------
### SETTING ###
# T6RN1 SA

pat_dir = "C:/OpenPT/PE_import/mpt/"
mpat = pat_dir + "RN1-SA-FBC.mpt"

Col = 0x32F6
Plane = 0

pxl_row = 256
pxl_col = 256

sdl_fg = 2		#２:XDL→ADL

f_sde = 0x00
f_vddsa = 0x00
f_vdd = 0x00

send = cache["X2A"]

#-----------------------------

analog_param = "SA-FBC"
log_dir = "C:\OpenPT\log\AnalogSDL\\"
sdl_cond = "AnalogSDL_"+analog_param+"-map_"+str(t_stamp)

#-----------------------------
TIMING_SET(40)

POWER_ON()
FF_RESET()
TEST_MODE()

if sdl_fg == 1:
    fg = OKCANCELBOX("Analog SDL Start")
    if fg == 1:
        analog_res = analog_sdl(mpat, Plane, send, f_sde, f_vddsa, f_vdd, pxl_row, pxl_col)
        np.savetxt(log_dir+"\\"+sdl_cond+".csv", analog_res, fmt = "%d", delimiter=",")

elif sdl_fg == 2:
        Single_DOUT(mpat, Col, Plane, send, f_sde, f_vddsa, f_vdd)

elif sdl_fg == 3:
    fg = OKCANCELBOX("SDL Start")
    if fg == 1:
        DIN_DOUT_SDL(mpat, Col, Plane, send, f_sde, f_vddsa, f_vdd)

else:
    FBC_single(mpat, Plane, send, f_sde, f_vddsa, f_vdd)

POWER_OFF()

