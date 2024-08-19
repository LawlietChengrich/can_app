#  zlgcan_demo.py
#
#  ~~~~~~~~~~~~
#
#  ZLGCAN USBCANFD Demo
#
#  ~~~~~~~~~~~~
#
#  ------------------------------------------------------------------
#  Author : guochuangjian    
#  Last change: 17.01.2019
#
#  Language: Python 3.6
#  ------------------------------------------------------------------
#
from zlgcan import *
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import threading
import time
import json

GRPBOX_WIDTH    = 200
MSGCNT_WIDTH    = 50
MSGID_WIDTH     = 80
MSGDIR_WIDTH    = 60
MSGINFO_WIDTH   = 100
MSGLEN_WIDTH    = 60
MSGDATA_WIDTH   = 320
MSGVIEW_WIDTH   = MSGCNT_WIDTH + MSGID_WIDTH + MSGDIR_WIDTH + MSGINFO_WIDTH + MSGLEN_WIDTH + MSGDATA_WIDTH
MSGVIEW_HEIGHT  = 500
SENDVIEW_HEIGHT = 250
REMOTE_WIN_WIDTH = 300
REMOTE_WIN_HEIGHT = 700

WIDGHT_WIDTH    = GRPBOX_WIDTH + MSGVIEW_WIDTH + 50
WIDGHT_HEIGHT   = MSGVIEW_HEIGHT + SENDVIEW_HEIGHT + 20

RM_DATA_HEAD_LEN = 5
MAX_RMDATA_LEN = 128
MPPT_CNT = 9
MAX_ONE_DATA_FRAME_LEN = 8
MAX_DISPLAY     = 1000
MAX_RCV_NUM     = 10
RM_MPPT_FRAME_CNT = 6
RM_BAT_FRAME_CNT = 3
RM_WING_FRAME_CNT = 2
DT_REMOTE_RETURN = 0b110
CANID_PRO_POS = 26
CANID_BUS_POS = 25
CANID_DT_POS = 20
CANID_DA_POS = 15

USBCANFD_TYPE    = (41, 42, 43)
USBCAN_XE_U_TYPE = (20, 21, 31)
USBCAN_I_II_TYPE = (3, 4)
###############################################################################
class PeriodSendThread(object):
    def __init__(self, period_func, args=[], kwargs={}):
        self._thread       = threading.Thread(target=self._run)
        self._function     = period_func
        self._args         = args
        self._kwargs       = kwargs
        self._period       = 0
        self._event        = threading.Event()
        self._period_event = threading.Event() 
        self._terminated   = False 
    
    def start(self):
        self._thread.start()

    def stop(self):
        self._terminated = True
        self._event.set()
        self._thread.join()

    def send_start(self, period):
        self._period = period
        self._event.set()

    def send_stop(self):
        self._period_event.set()

    def _run(self):
        while True:
            self._event.wait()
            self._event.clear()
            if self._terminated:
                break
            self._function(*self._args, **self._kwargs) 
            while not self._period_event.wait(self._period):
                self._function(*self._args, **self._kwargs)
            self._period_event.clear()
###############################################################################
class ZCAN_Demo(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("德华CAN通讯上位机")
        self.resizable(False, False)
        winw = WIDGHT_WIDTH
        winh = WIDGHT_HEIGHT
        scrw = self.winfo_screenwidth()
        scrh = self.winfo_screenheight()
        x = (scrw - winw)/2
        y = (scrh - winh)/2
        self.geometry('%dx%d+%d+%d' % (winw, winh, x, y))
        self.protocol("WM_DELETE_WINDOW", self.Form_OnClosing)

        self.DeviceInit()
        self.WidgetsInit()

        self._dev_info = None
        with open("./dev_info.json", "r") as fd:
            self._dev_info = json.load(fd)
        if self._dev_info == None:
            print("device info no exist!")
            return 

        self.DeviceInfoInit()
        self.ChnInfoUpdate(self._isOpen)

    def DeviceInit(self):
        self._zcan       = ZCAN() 
        self._dev_handle = INVALID_DEVICE_HANDLE 
        self._can_handle = INVALID_CHANNEL_HANDLE 

        self._isOpen = False
        self._isChnOpen = False

        #current device info
        self._is_canfd = False
        self._res_support = False

        #Transmit and receive count display
        self._tx_cnt = 0
        self._rx_cnt = 0
        self._view_cnt = 0

        #read can/canfd message thread
        self._read_thread = None
        self._terminated = False
        self._lock = threading.RLock()

        #period send var
        self._is_sending   = False
        self._id_increase  = False 
        self._send_num     = 1
        self._send_cnt     = 1
        self._is_canfd_msg = False
        self._send_msgs    = None
        self._send_thread  = None

    def WidgetsInit(self):
        self._dev_frame = tk.Frame(self)
        self._dev_frame.grid(row=0, column=0, padx=2, pady=2, sticky=tk.NSEW)

        # Device connect group
        self.gbDevConnect = tk.LabelFrame(self._dev_frame, height=100, width=GRPBOX_WIDTH, text="设备选择")
        self.gbDevConnect.grid_propagate(0)
        self.gbDevConnect.grid(row=0, column=0, padx=2, pady=2, sticky=tk.NE)
        self.DevConnectWidgetsInit()

        self.gbCANCfg = tk.LabelFrame(self._dev_frame, height=170, width=GRPBOX_WIDTH, text="通道配置")
        self.gbCANCfg.grid(row=1, column=0, padx=2, pady=2, sticky=tk.NSEW)
        self.gbCANCfg.grid_propagate(0)
        self.CANChnWidgetsInit()

        self.gbDevInfo = tk.LabelFrame(self._dev_frame, height=230, width=GRPBOX_WIDTH, text="设备信息")
        self.gbDevInfo.grid(row=2, column=0, padx=2, pady=2, sticky=tk.NSEW)
        self.gbDevInfo.grid_propagate(0)
        self.DevInfoWidgetsInit()

        self.gbMsgDisplay = tk.LabelFrame(height=MSGVIEW_HEIGHT, width=MSGVIEW_WIDTH + 12, text="报文显示")
        self.gbMsgDisplay.grid(row=0, column=1, padx=0, pady=6, sticky=tk.NSEW)
        self.gbMsgDisplay.grid_propagate(0)
        self.MsgDisplayWidgetsInit()

        self.gbMsgSend = tk.LabelFrame(heigh=SENDVIEW_HEIGHT, width=MSGVIEW_WIDTH + 12, text="报文发送")
        self.gbMsgSend.grid(row=2, column=1, padx=0, pady=0, sticky=tk.NSEW)
        self.gbMsgSend.grid_propagate(0)
        self.MsgSendWidgetsInit()

    def DeviceInfoInit(self):
        self.cmbDevType["value"] = tuple([dev_name for dev_name in self._dev_info])
        self.cmbDevType.current(6)

    def DevConnectWidgetsInit(self):
        #Device Type
        tk.Label(self.gbDevConnect, text="设备类型:").grid(row=0, column=0, sticky=tk.E)
        self.cmbDevType = ttk.Combobox(self.gbDevConnect, width=16, state="readonly")
        self.cmbDevType.grid(row=0, column=1, sticky=tk.E)

        #Device Index
        tk.Label(self.gbDevConnect, text="设备索引:").grid(row=1, column=0, sticky=tk.E)
        self.cmbDevIdx = ttk.Combobox(self.gbDevConnect, width=16, state="readonly")
        self.cmbDevIdx.grid(row=1, column=1, sticky=tk.E)
        self.cmbDevIdx["value"] = tuple([i for i in range(4)])
        self.cmbDevIdx.current(0)

        #Open/Close Device
        self.strvDevCtrl = tk.StringVar()
        self.strvDevCtrl.set("打开")
        self.btnDevCtrl = ttk.Button(self.gbDevConnect, textvariable=self.strvDevCtrl, command=self.BtnOpenDev_Click)
        self.btnDevCtrl.grid(row=2, column=0, columnspan=2, pady=2) 

    def CANChnWidgetsInit(self):
        #CAN Channel
        tk.Label(self.gbCANCfg, anchor=tk.W, text="CAN通道:").grid(row=0, column=0, sticky=tk.W)
        self.cmbCANChn = ttk.Combobox(self.gbCANCfg, width=12, state="readonly")
        self.cmbCANChn.grid(row=0, column=1, sticky=tk.E)

        #Work Mode
        tk.Label(self.gbCANCfg, anchor=tk.W, text="工作模式:").grid(row=1, column=0, sticky=tk.W)
        self.cmbCANMode = ttk.Combobox(self.gbCANCfg, width=12, state="readonly")
        self.cmbCANMode.grid(row=1, column=1, sticky=tk.E)

        #CAN Baudrate 
        tk.Label(self.gbCANCfg, anchor=tk.W, text="波特率:").grid(row=2, column=0, sticky=tk.W)
        self.cmbBaudrate = ttk.Combobox(self.gbCANCfg, width=12, state="readonly")
        self.cmbBaudrate.grid(row=2, column=1, sticky=tk.W)
        
        #CAN Data Baudrate 
        tk.Label(self.gbCANCfg, anchor=tk.W, text="数据域波特率:").grid(row=3, column=0, sticky=tk.W)
        self.cmbDataBaudrate = ttk.Combobox(self.gbCANCfg, width=12, state="readonly")
        self.cmbDataBaudrate.grid(row=3, column=1, sticky=tk.W)

        #resistance enable
        tk.Label(self.gbCANCfg, anchor=tk.W, text="终端电阻:").grid(row=4, column=0, sticky=tk.W)
        self.cmbResEnable = ttk.Combobox(self.gbCANCfg, width=12, state="readonly")
        self.cmbResEnable.grid(row=4, column=1, sticky=tk.W)

        #CAN Control
        self.strvCANCtrl = tk.StringVar()
        self.strvCANCtrl.set("打开")
        self.btnCANCtrl = ttk.Button(self.gbCANCfg, textvariable=self.strvCANCtrl, command=self.BtnOpenCAN_Click) 
        self.btnCANCtrl.grid(row=5, column=0, columnspan=2, padx=2, pady=2)

    def DevInfoWidgetsInit(self):
        #Hardware Version
        tk.Label(self.gbDevInfo, anchor=tk.W, text="硬件版本:").grid(row=0, column=0, sticky=tk.W)
        self.strvHwVer = tk.StringVar(value='')
        tk.Label(self.gbDevInfo, anchor=tk.W, textvariable=self.strvHwVer).grid(row=0, column=1, sticky=tk.W)

        #Firmware Version
        tk.Label(self.gbDevInfo, anchor=tk.W, text="固件版本:").grid(row=1, column=0, sticky=tk.W)
        self.strvFwVer = tk.StringVar(value='')
        tk.Label(self.gbDevInfo, anchor=tk.W, textvariable=self.strvFwVer).grid(row=1, column=1, sticky=tk.W)

        #Driver Version
        tk.Label(self.gbDevInfo, anchor=tk.W, text="驱动版本:").grid(row=2, column=0, sticky=tk.W)
        self.strvDrVer = tk.StringVar(value='')
        tk.Label(self.gbDevInfo, anchor=tk.W, textvariable=self.strvDrVer).grid(row=2, column=1, sticky=tk.W)

        #Interface Version
        tk.Label(self.gbDevInfo, anchor=tk.W, text="动态库版本:").grid(row=3, column=0, sticky=tk.W)
        self.strvInVer = tk.StringVar(value='')
        tk.Label(self.gbDevInfo, anchor=tk.W, textvariable=self.strvInVer).grid(row=3, column=1, sticky=tk.W)

        #CAN num
        tk.Label(self.gbDevInfo, anchor=tk.W, text="CAN通道数:").grid(row=4, column=0, sticky=tk.W)
        self.strvCANNum = tk.StringVar(value='')
        tk.Label(self.gbDevInfo, anchor=tk.W, textvariable=self.strvCANNum).grid(row=4, column=1, sticky=tk.W)

        #Device Serial
        tk.Label(self.gbDevInfo, anchor=tk.W, text="设备序列号:").grid(row=5, column=0, sticky=tk.W)
        self.strvSerial = tk.StringVar(value='')
        tk.Label(self.gbDevInfo, anchor=tk.W, textvariable=self.strvSerial).grid(row=6, column=0, columnspan=2, sticky=tk.W)
    
        #Hardware type
        tk.Label(self.gbDevInfo, anchor=tk.W, text="硬件类型:").grid(row=7, column=0, sticky=tk.W)
        self.strvHwType = tk.StringVar(value='')
        tk.Label(self.gbDevInfo, anchor=tk.W, textvariable=self.strvHwType).grid(row=8, column=0, columnspan=2, sticky=tk.W)
    
    def MsgDisplayWidgetsInit(self):
        self._msg_frame = tk.Frame(self.gbMsgDisplay, height=MSGVIEW_HEIGHT, width=WIDGHT_WIDTH-GRPBOX_WIDTH+10)
        self._msg_frame.pack(side=tk.TOP)
        
        self.treeMsg = ttk.Treeview(self._msg_frame, height=20, show="headings")
        self.treeMsg["columns"] = ("cnt", "id", "direction", "info", "len", "data")

        self.treeMsg.column("cnt",       anchor = tk.CENTER, width=MSGCNT_WIDTH)
        self.treeMsg.column("id",        anchor = tk.CENTER, width=MSGID_WIDTH)
        self.treeMsg.column("direction", anchor = tk.CENTER, width=MSGDIR_WIDTH)
        self.treeMsg.column("info",      anchor = tk.CENTER, width=MSGINFO_WIDTH)
        self.treeMsg.column("len",       anchor = tk.CENTER, width=MSGLEN_WIDTH)
        self.treeMsg.column("data", width=MSGDATA_WIDTH)

        self.treeMsg.heading("cnt", text="序号")
        self.treeMsg.heading("id", text="帧ID")
        self.treeMsg.heading("direction", text="方向")
        self.treeMsg.heading("info", text="帧信息")
        self.treeMsg.heading("len", text="长度")
        self.treeMsg.heading("data", text="数据")
        
        self.hbar = ttk.Scrollbar(self._msg_frame, orient=tk.HORIZONTAL, command=self.treeMsg.xview)
        self.hbar.pack(side=tk.BOTTOM, fill=tk.X)
        self.vbar = ttk.Scrollbar(self._msg_frame, orient=tk.VERTICAL, command=self.treeMsg.yview)
        self.vbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.treeMsg.configure(xscrollcommand=self.hbar.set, yscrollcommand=self.vbar.set)
        
        self.treeMsg.pack(side=tk.LEFT)
        self.treeMsg.selection_set()
        self.btnClrCnt = ttk.Button(self.gbMsgDisplay, width=10, text="清空", command=self.BtnClrCnt_Click) 
        self.btnClrCnt.pack(side=tk.RIGHT)

        self.strvRxCnt = tk.StringVar()
        self.strvRxCnt.set("0")
        tk.Label(self.gbMsgDisplay, anchor=tk.W, width=5, textvariable=self.strvRxCnt).pack(side=tk.RIGHT)
        tk.Label(self.gbMsgDisplay, width=10, text="接收帧数:").pack(sid=tk.RIGHT)

        self.strvTxCnt = tk.StringVar()
        self.strvTxCnt.set("0")
        tk.Label(self.gbMsgDisplay, anchor=tk.W, width=5, textvariable=self.strvTxCnt).pack(side=tk.RIGHT)
        tk.Label(self.gbMsgDisplay, width=10, text="发送帧数:").pack(side=tk.RIGHT)

    def CloseRemoteWin(self):
        if self.WinRemote != None:
            self.WinRemote.destroy()
            self.WinRemote = None

    def RemoteDataWindowCreate(self, tmt):
        if self.WinRemote == None:
            self.WinRemote = tk.Toplevel(self)
            self.WinRemote.protocol("WM_DELETE_WINDOW", self.CloseRemoteWin)
            winw = REMOTE_WIN_WIDTH
            winh = REMOTE_WIN_HEIGHT
            scrw = self.winfo_screenwidth()
            scrh = self.winfo_screenheight()
            x = (scrw - WIDGHT_WIDTH)/2 + WIDGHT_WIDTH
            y = (scrh - WIDGHT_HEIGHT)/2
            self.WinRemote.geometry('%dx%d+%d+%d' % (winw, winh, x, y))
            self.WinRemote.minsize(200,200)

        for widget in self.WinRemote.winfo_children():
            widget.destroy()

        tk.Label(self.WinRemote, width=12, anchor=tk.W, text=("遥测请求计数:")).grid(row=0, column=0, sticky=tk.W)
        tk.Label(self.WinRemote, width=12, anchor=tk.W, text=("正确指令计数:")).grid(row=1, column=0, sticky=tk.W)
        tk.Label(self.WinRemote, width=12, anchor=tk.W, text=("错误指令计数:")).grid(row=2, column=0, sticky=tk.W)
        tk.Label(self.WinRemote, width=12, anchor=tk.W, text=("最近执行指令:")).grid(row=3, column=0, sticky=tk.W)
        tk.Label(self.WinRemote, width=12, anchor=tk.W, text=("备份接收计数:")).grid(row=4, column=0, sticky=tk.W)
        tmt_value = []

        if tmt == 0:
            self.WinRemote.title("遥测MPPT数据")
            for i in range(0, 9):
                tk.Label(self.WinRemote, width=12, anchor=tk.W, text=("MPPT%d 电压(V):" % (i+1))).grid(row=i+RM_DATA_HEAD_LEN, column=0, sticky=tk.W)
                tk.Label(self.WinRemote, width=6, anchor=tk.W, text=("未知")).grid(row=i+RM_DATA_HEAD_LEN, column=1, sticky=tk.W)
                tk.Label(self.WinRemote, width=6, anchor=tk.W, text=("电流(A):")).grid(row=i+RM_DATA_HEAD_LEN, column=2, sticky=tk.W)
                tk.Label(self.WinRemote, width=6, anchor=tk.W, text=("未知")).grid(row=i+RM_DATA_HEAD_LEN, column=3, sticky=tk.W)
                tk.Label(self.WinRemote, width=4, anchor=tk.W, text=("状态:")).grid(row=i+RM_DATA_HEAD_LEN, column=4, sticky=tk.W)
                tk.Label(self.WinRemote, width=3, anchor=tk.W, text=("未知")).grid(row=i+RM_DATA_HEAD_LEN, column=5, sticky=tk.W)
        elif tmt == 1:
            self.WinRemote.title("遥测BAT数据")
            tmt_value = [
                "母线电压(V):", 
                "蓄电池电压(V):", 
                "负载总电流(A):", 
                "蓄电池电流(A):", 
                "放电欠压状态:",
                "自主加电状态:", 
                "放电状态:"]

            for i in range(0, len(tmt_value)):
                tk.Label(self.WinRemote, width=12, anchor=tk.W, text=(tmt_value[i])).grid(row=i+RM_DATA_HEAD_LEN, column=0, sticky=tk.W)
                tk.Label(self.WinRemote, width=12, anchor=tk.W, text=("未知")).grid(row=i+RM_DATA_HEAD_LEN, column=1, sticky=tk.W)
        elif tmt == 2:
            self.WinRemote.title("遥测WING数据")
            tmt_value = [
                "QVA:", 
                "QVB:", 
                "霍尔组件:", 
                "发射相控阵A:", 
                "发射相控阵B:",
                "接收相控阵:", 
                "飞轮X:",
                "飞轮Y:",
                "飞轮Z:",
                "飞轮S:",
                "帆板天线解锁:",
                ]

            for i in range(0, len(tmt_value)):
                tk.Label(self.WinRemote, width=12, anchor=tk.W, text=(tmt_value[i])).grid(row=i+RM_DATA_HEAD_LEN, column=0, sticky=tk.W)
                tk.Label(self.WinRemote, width=12, anchor=tk.W, text=("未知")).grid(row=i+RM_DATA_HEAD_LEN, column=1, sticky=tk.W)

    def MsgSendWidgetsInit(self):
        #Send Type
        tk.Label(self.gbMsgSend, anchor=tk.W, text="发送方式:").grid(row=0, column=0, sticky=tk.W)
        self.cmbSendType = ttk.Combobox(self.gbMsgSend, width=8, state="readonly")
        self.cmbSendType.grid(row = 0, column=1, sticky=tk.W) 
        self.cmbSendType["value"] = ("正常发送", "单次发送", "自发自收")
        self.cmbSendType.current(0)

        #CAN Type
        tk.Label(self.gbMsgSend, anchor=tk.W, text="帧类型:").grid(row=0, column=2, sticky=tk.W)
        self.cmbMsgType = ttk.Combobox(self.gbMsgSend, width=6, state="readonly")
        self.cmbMsgType.grid(row = 0, column=3, sticky=tk.W) 
        self.cmbMsgType["value"] = ("标准帧", "扩展帧")
        self.cmbMsgType.current(1)

        #CAN Format 
        tk.Label(self.gbMsgSend, anchor=tk.W, text="帧格式:").grid(row=0, column=4, sticky=tk.W)
        self.cmbMsgFormat = ttk.Combobox(self.gbMsgSend, width=6, state="readonly") 
        self.cmbMsgFormat.grid(row = 0, column=5, sticky=tk.W) 
        self.cmbMsgFormat["value"] = ("数据帧", "远程帧")
        self.cmbMsgFormat.bind("<<ComboboxSelected>>", self.CmbMsgFormatUpdate)
        self.cmbMsgFormat.current(0)

        #CANFD 
        tk.Label(self.gbMsgSend, anchor=tk.W, text="CAN类型:").grid(row=0, column=6, sticky=tk.W)
        self.cmbMsgCANFD = ttk.Combobox(self.gbMsgSend, width=10, state="readonly")
        self.cmbMsgCANFD.grid(row=0, column=7, sticky=tk.W) 
        self.cmbMsgCANFD["value"] = ("CAN", "CANFD", "CANFD BRS")
        self.cmbMsgCANFD.bind("<<ComboboxSelected>>", self.CmbMsgCANFDUpdate)
        self.cmbMsgCANFD.current(0)

        #CAN ID
        tk.Label(self.gbMsgSend, anchor=tk.W, text="帧ID(hex):").grid(row=1, column=0, sticky=tk.W)
        self.entryMsgID = tk.Entry(self.gbMsgSend, width=10, text="100")
        self.entryMsgID.grid(row=1, column=1, sticky=tk.W)
        self.entryMsgID.tmp_value = 0 
        #self.entryMsgID.insert(0, "100")

        #CAN Length 
        tk.Label(self.gbMsgSend, anchor=tk.W, text="长度:").grid(row=1, column=2, sticky=tk.W)
        self.cmbMsgLen = ttk.Combobox(self.gbMsgSend, width=6, state="readonly")
        self.cmbMsgLen["value"] = tuple([i for i in range(9)])
        self.cmbMsgLen.current(8) 
        self.cmbMsgLen.grid(row=1, column=3, sticky=tk.W) 

        #Data
        tk.Label(self.gbMsgSend, anchor=tk.W, text="数据(hex):").grid(row=1, column=4, sticky=tk.W)
        self.entryMsgData = tk.Entry(self.gbMsgSend, width=30)
        self.entryMsgData.grid(row = 1, column=5, columnspan=4, sticky=tk.W) 
        self.entryMsgData.insert(0, "FF 00 00 00 00 00 00 00")

        #send frame number
        tk.Label(self.gbMsgSend, anchor=tk.W, text="发送帧数:").grid(row=2, column=0, sticky=tk.W)
        self.entryMsgNum = tk.Entry(self.gbMsgSend, width=10)
        self.entryMsgNum.grid(row=2, column=1, sticky=tk.W) 
        self.entryMsgNum.insert(0, "1")

        #send frame cnt 
        tk.Label(self.gbMsgSend, anchor=tk.W, text="发送次数:").grid(row=2, column=2, sticky=tk.W)
        self.entryMsgCnt = tk.Entry(self.gbMsgSend, width=8)
        self.entryMsgCnt.grid(row=2, column=3, sticky=tk.W) 
        self.entryMsgCnt.insert(0, "1")

        #send frame period
        tk.Label(self.gbMsgSend, anchor=tk.W, text="发送间隔(ms):").grid(row=2, column=4, sticky=tk.W)
        self.entryMsgPeriod = tk.Entry(self.gbMsgSend, width=8)
        self.entryMsgPeriod.grid(row=2, column=5, sticky=tk.W) 
        self.entryMsgPeriod.insert(0, "1")

        #msg id add
        self.varIDInc = tk.IntVar()
        self.chkbtnIDInc = tk.Checkbutton(self.gbMsgSend, text="ID递增", variable=self.varIDInc)
        self.chkbtnIDInc.grid(row=2, column=6, columnspan=2, sticky=tk.W)

        #Send Butten
        self.strvSend = tk.StringVar()
        self.strvSend.set("发送")
        self.btnMsgSend = ttk.Button(self.gbMsgSend, textvariable=self.strvSend, command=self.BtnSendMsg_Click) 
        self.btnMsgSend.grid(row=7, column=6, padx=6, pady=2)
        self.btnMsgSend["state"] = tk.DISABLED

        tk.Label(self.gbMsgSend, anchor=tk.W, text="自定义部分:").grid(row=3, column=0, sticky=tk.W)

        #ID28-26 P 优先级
        tk.Label(self.gbMsgSend, anchor=tk.W, text="优先级(P):").grid(row=4, column=0, sticky=tk.W)
        self.cmbProvity = ttk.Combobox(self.gbMsgSend, width=6, state="readonly")
        self.cmbProvity.grid(row = 4, column=1, sticky=tk.W)
        self.cmbProvity.bind('<<ComboboxSelected>>', self.CanIdChangeEvent)
        self.cmbProvity["value"] = ("主节点", "从节点")
        self.cmbProvity.current(0)
        self.entryMsgID.tmp_value += 0b11<<CANID_PRO_POS

        #ID25 LT 总线
        tk.Label(self.gbMsgSend, anchor=tk.W, text="总线标志(LT):").grid(row=4, column=2, sticky=tk.W)
        self.cmbBusFlag = ttk.Combobox(self.gbMsgSend, width=6, state="readonly")
        self.cmbBusFlag.grid(row = 4, column=3, sticky=tk.W)
        self.cmbBusFlag.bind('<<ComboboxSelected>>', self.CanIdChangeEvent)
        self.cmbBusFlag["value"] = ("A总线", "B总线")
        self.cmbBusFlag.current(1)
        self.entryMsgID.tmp_value += 1<<CANID_BUS_POS

        #ID24 - 20 DT 数据类型
        tk.Label(self.gbMsgSend, anchor=tk.W, text="数据类型(DT):").grid(row=5, column=0, sticky=tk.W)
        self.cmbDataType = ttk.Combobox(self.gbMsgSend, width=8, state="readonly")
        self.cmbDataType.grid(row = 5, column=1, sticky=tk.W)
        self.cmbDataType.bind('<<ComboboxSelected>>', self.DataTypeChangeEvent)
        #self.cmbDataType.bind('<<ComboboxSelected>>', self.CanIdChangeEvent)
        #self.cmbDataType.bind('<<ComboboxSelected>>', self.TmtTypeChangeEvent)
        self.cmbDataType["value"] = ("遥测", "复位", "短控", "备份数据请求", "备份数据广播")
        self.cmbDataType.current(0)
        self.WinRemote = None
        self.RemoteDataWindowCreate(0)
        self.entryMsgID.tmp_value += 0<<CANID_DT_POS

        #指令码TMT


        tk.Label(self.gbMsgSend, anchor=tk.W, text="指令码(TMT):").grid(row=5, column=2, sticky=tk.W)
        self.cmbTmt = ttk.Combobox(self.gbMsgSend, width=8, state="readonly")
        self.cmbTmt.grid(row = 5, column=3, sticky=tk.W)

        self.cmbTmt.bind('<<ComboboxSelected>>', self.TmtTypeChangeEvent)
        self.cmbTmt["value"] = ("MPPT", "BAT", "WING")
        self.cmbTmt.current(0)

        #指令码参数
        tk.Label(self.gbMsgSend, anchor=tk.W, text="指令码参数:").grid(row=5, column=4, sticky=tk.W)
        self.cmbTmtPar = ttk.Combobox(self.gbMsgSend, width=14, state="readonly")
        self.cmbTmtPar.grid(row = 5, column=5, columnspan=2, sticky=tk.W)

        self.cmbTmtPar.bind('<<ComboboxSelected>>', self.TmtParChangeEvent)


        #ID19 - 15 DA 目的地址
        tk.Label(self.gbMsgSend, anchor=tk.W, text="目的地址(DA):").grid(row=6, column=0, sticky=tk.W)
        self.cmbDa = ttk.Combobox(self.gbMsgSend, width=6, state="readonly")
        self.cmbDa.grid(row = 6, column=1, sticky=tk.W)
        self.cmbDa.bind('<<ComboboxSelected>>', self.CanIdChangeEvent)
        self.cmbDa["value"] = ("星算", "电控主", "电控备", "广播")
        self.cmbDa.current(2)
        self.entryMsgID.tmp_value += 0b10001<<CANID_DA_POS

        #ID14 - 10 SA 源地址
        tk.Label(self.gbMsgSend, anchor=tk.W, text="源地址(SA):").grid(row=6, column=2, sticky=tk.W)
        self.cmbSa = ttk.Combobox(self.gbMsgSend, width=6, state="readonly")
        self.cmbSa.grid(row = 6, column=3, sticky=tk.W)

        self.cmbSa.bind('<<ComboboxSelected>>', self.CanIdChangeEvent)
        self.cmbSa["value"] = ("星算", "电控主", "电控备")
        self.cmbSa.current(0)

        #ID9 - 8 FT 源地址
        tk.Label(self.gbMsgSend, anchor=tk.W, text="单复帧(FT):").grid(row=6, column=4, sticky=tk.W)
        self.cmbFt = ttk.Combobox(self.gbMsgSend, width=6, state="readonly")
        self.cmbFt.grid(row = 6, column=5, sticky=tk.W)

        self.cmbFt.bind('<<ComboboxSelected>>', self.CanIdChangeEvent)
        self.cmbFt["value"] = ("单帧", "复首", "复中", "复尾")
        self.cmbFt.current(0)

        tk.Label(self.gbMsgSend, anchor=tk.W, text="帧计数(FC):").grid(row=6, column=6, sticky=tk.W)
        self.entryFc = tk.Entry(self.gbMsgSend, width=8)
        self.entryFc.grid(row=6, column=7, sticky=tk.W) 
        self.entryFc.bind('<KeyRelease>', self.CanIdChangeEvent)
        self.entryFc.insert(0, 1)
        self.entryMsgID.tmp_value += 1

        self.entryMsgID.tmp_value = "{:X}".format(self.entryMsgID.tmp_value)

        self.entryMsgID.insert(0, self.entryMsgID.tmp_value)

        self.Rmdata_tmt = 0
        self.Rmdata_cur_cnt = 0
        self.Rmdata_self = [0]*MAX_RMDATA_LEN

###############################################################################
### Function 
###############################################################################
    def __dlc2len(self, dlc):
        if dlc <= 8:
            return dlc
        elif dlc == 9:
            return 12
        elif dlc == 10:
            return 16
        elif dlc == 11:
            return 20
        elif dlc == 12:
            return 24
        elif dlc == 13:
            return 32
        elif dlc == 14:
            return 48
        else:
            return 64

    def CANMsg2View(self, msg, is_transmit=True):
        view = []
        view.append(str(self._view_cnt))
        self._view_cnt += 1 
        view.append(hex(msg.can_id)[2:].upper())
        view.append("发送" if is_transmit else "接收")

        str_info = ''
        str_info += 'EXT' if msg.eff else 'STD'
        if msg.rtr:
            str_info += ' RTR'
        view.append(str_info)
        view.append(str(msg.can_dlc))
        if msg.rtr:
            view.append('')
        else:
            view.append(''.join(format(msg.data[i],'02X') + ' ' for i in range(msg.can_dlc)))
        return view

    def CANFDMsg2View(self, msg, is_transmit=True):
        view = [] 
        view.append(str(self._view_cnt))
        self._view_cnt += 1 
        
        view.append(hex(msg.can_id)[2:])
        view.append("发送" if is_transmit else "接收")

        str_info = ''
        str_info += 'EXT' if msg.eff else 'STD'
        if msg.rtr:
            str_info += ' RTR'
        else:
            str_info += ' FD'
            if msg.brs:
                str_info += ' BRS'
            if msg.esi:
                str_info += ' ESI' 
        view.append(str_info)
        view.append(str(msg.len))
        if msg.rtr:
            view.append('')
        else:
            view.append(''.join(hex(msg.data[i])[2:] + ' ' for i in range(msg.len)))
        return view

    def ChnInfoUpdate(self, is_open):
        #通道信息获取
        cur_dev_info = self._dev_info[self.cmbDevType.get()]
        cur_chn_info = cur_dev_info["chn_info"]
        
        if is_open:
            # 通道 
            self.cmbCANChn["value"] = tuple([i for i in range(cur_dev_info["chn_num"])])
            self.cmbCANChn.current(0)

            # 工作模式
            self.cmbCANMode["value"] = ("正常模式", "只听模式")
            self.cmbCANMode.current(0)

            # 波特率
            self.cmbBaudrate["value"] = tuple([brt for brt in cur_chn_info["baudrate"].keys()])
            self.cmbBaudrate.current(len(self.cmbBaudrate["value"]) - 3)

            if cur_chn_info["is_canfd"] == True:
                # 数据域波特率 
                self.cmbDataBaudrate["value"] = tuple([brt for brt in cur_chn_info["data_baudrate"].keys()])
                self.cmbDataBaudrate.current(0)
                self.cmbDataBaudrate["state"] = "readonly"

            if cur_chn_info["sf_res"] == True:
                self.cmbResEnable["value"] = ("使能", "失能")
                self.cmbResEnable.current(0)
                self.cmbResEnable["state"] = "readonly"

            self.btnCANCtrl["state"] = tk.NORMAL
        else:
            self.cmbCANChn["state"] = tk.DISABLED
            self.cmbCANMode["state"] = tk.DISABLED
            self.cmbBaudrate["state"] = tk.DISABLED
            self.cmbDataBaudrate["state"] = tk.DISABLED
            self.cmbResEnable["state"] = tk.DISABLED
            
            self.cmbCANChn["value"] = ()
            self.cmbCANMode["value"] = ()
            self.cmbBaudrate["value"] = ()
            self.cmbDataBaudrate["value"] = ()
            self.cmbResEnable["value"] = ()

            self.btnCANCtrl["state"] = tk.DISABLED

    def ChnInfoDisplay(self, enable):
        if enable:
            self.cmbCANChn["state"] = "readonly"
            self.cmbCANMode["state"] = "readonly"
            self.cmbBaudrate["state"] = "readonly" 
            if self._is_canfd: 
                self.cmbDataBaudrate["state"] = "readonly" 
            if self._res_support: 
                self.cmbResEnable["state"] = "readonly"
        else:
            self.cmbCANChn["state"] = tk.DISABLED
            self.cmbCANMode["state"] = tk.DISABLED
            self.cmbBaudrate["state"] = tk.DISABLED
            self.cmbDataBaudrate["state"] = tk.DISABLED
            self.cmbResEnable["state"] = tk.DISABLED

    def RmDataMpptDisplay(self, self_data):
            rm_mppt_v = [0] * MPPT_CNT
            rm_mppt_i = [0] * MPPT_CNT
            rm_mppt_status = [0] * MPPT_CNT 
            
            for i in range(0, MPPT_CNT):
                rm_mppt_v[i] = self_data[2*i] + self_data[2*i+1]*0.01
                rm_mppt_i[i] = self_data[2*(i+MPPT_CNT)] + self_data[2*(i+MPPT_CNT)+1]*0.01

            for i in range(0,MPPT_CNT-1):
                rm_mppt_status[i] = (self_data[MPPT_CNT*4]>>(7-i))&0x1
                if rm_mppt_status[i] == 1:
                    rm_mppt_status[i] = "开"
                else:
                    rm_mppt_status[i] = "关"
            rm_mppt_status[MPPT_CNT-1] = self_data[MPPT_CNT*4+1]>>7 & 0x1
            if rm_mppt_status[MPPT_CNT-1] == 1:
                 rm_mppt_status[MPPT_CNT-1] = "开"
            else:
                rm_mppt_status[MPPT_CNT-1] = "关"

            for i in range(0, MPPT_CNT):
                tk.Label(self.WinRemote, width=6, anchor=tk.W, text=(str(rm_mppt_v[i]))).grid(row=i+RM_DATA_HEAD_LEN, column=1, sticky=tk.W)
                tk.Label(self.WinRemote, width=6, anchor=tk.W, text=(str(rm_mppt_i[i]))).grid(row=i+RM_DATA_HEAD_LEN, column=3, sticky=tk.W)
                tk.Label(self.WinRemote, width=3, anchor=tk.W, text=(rm_mppt_status[i])).grid(row=i+RM_DATA_HEAD_LEN, column=5, sticky=tk.W)

    def RmDataBatDisplay(self, self_data):
        rm_bat_v = [0] * 4
        rm_bat_status = [0] * 3

        for i in range(0, 3):
            rm_bat_v[i] = self_data[2*i] + self_data[2*i+1]*0.01

        rm_bat_v[3] = (self_data[6]&0x7f) + self_data[7]*0.01


        if (self_data[6]>>7)&0x1 == 1:
            rm_bat_v[3] = -rm_bat_v[3]
            tk.Label(self.WinRemote, width=6, anchor=tk.W, text=("放电")).grid(row=3+RM_DATA_HEAD_LEN, column=2, sticky=tk.W)
        else:
            tk.Label(self.WinRemote, width=6, anchor=tk.W, text=("充电")).grid(row=3+RM_DATA_HEAD_LEN, column=2, sticky=tk.W)

        for i in range(0,3):
            rm_bat_status[i] = (self_data[8]>>(7-i))&0x1
            if rm_bat_status[i] == 1:
                rm_bat_status[i] = "使能"
            else:
                rm_bat_status[i] = "禁止"

        for i in range(0, 4):
            tk.Label(self.WinRemote, width=6, anchor=tk.W, text=("{:.2f}".format(rm_bat_v[i]))).grid(row=i+RM_DATA_HEAD_LEN, column=1, sticky=tk.W)

        for i in range(0, 3):
            tk.Label(self.WinRemote, width=6, anchor=tk.W, text=(rm_bat_status[i])).grid(row=i+4+RM_DATA_HEAD_LEN, column=1, sticky=tk.W)

    def RmDataWingDisplay(self, self_data):
        rm_wing_status1 = [0] * 6
        rm_wing_status2 = [0] * 5

        for i in range(0,6):
            rm_wing_status1[i] = (self_data[0]>>(7-i))&0x1
            if rm_wing_status1[i] == 1:
                rm_wing_status1[i] = "接通"
            else:
                rm_wing_status1[i] = "断开"

        for i in range(0,4):
            rm_wing_status2[i] = (self_data[1]>>(7-i))&0x1
            if rm_wing_status2[i] == 1:
                rm_wing_status2[i] = "接通"
            else:
                rm_wing_status2[i] = "断开"

        if ((self_data[1]>>3)&0x1) == 1:
            rm_wing_status2[4] = "开"
        else:
            rm_wing_status2[4] = "关"

        for i in range(0, 6):
            tk.Label(self.WinRemote, width=6, anchor=tk.W, text=(rm_wing_status1[i])).grid(row=i+RM_DATA_HEAD_LEN, column=1, sticky=tk.W)

        for i in range(0, 5):
            tk.Label(self.WinRemote, width=6, anchor=tk.W, text=(rm_wing_status2[i])).grid(row=i+6+RM_DATA_HEAD_LEN, column=1, sticky=tk.W)

    def RmDataUpdata(self, msgs, msgs_num):
        if (msgs[0].frame.can_id & 0xff) == 1:
			#收到首包重置显示数据
            Rmdata_len = msgs[0].frame.data[0] + msgs[0].frame.data[1] *16
            self.Rmdata_tmt = msgs[0].frame.data[2]
            RmdataReqCnt = msgs[0].frame.data[3]
            RmdataCorrectCnt = (msgs[0].frame.data[4] >> 4) &0xf
            RmdataWrongCnt = msgs[0].frame.data[4]&0xf
            RmdataCmdNewest = msgs[0].frame.data[5]
            RmdataBackupCnt =  msgs[0].frame.data[6]
            tk.Label(self.WinRemote, width=6, anchor=tk.W, text=str(RmdataReqCnt)).grid(row=0, column=1, sticky=tk.W)
            tk.Label(self.WinRemote, width=6, anchor=tk.W, text=str(RmdataCorrectCnt)).grid(row=1, column=1, sticky=tk.W)
            tk.Label(self.WinRemote, width=6, anchor=tk.W, text=str(RmdataWrongCnt)).grid(row=2, column=1, sticky=tk.W)
            tk.Label(self.WinRemote, width=6, anchor=tk.W, text=str(RmdataCmdNewest)).grid(row=3, column=1, sticky=tk.W)
            tk.Label(self.WinRemote, width=6, anchor=tk.W, text=str(RmdataBackupCnt)).grid(row=4, column=1, sticky=tk.W)
            self.Rmdata_cur_cnt = 0
            for i in range(0, len(self.Rmdata_self)):
                self.Rmdata_self[i] = 0

		#分包接收
        if self.Rmdata_cur_cnt == 0:
            for i in range(0, msgs_num):
                if i == 0:
                    self.Rmdata_self[0] = msgs[0].frame.data[MAX_ONE_DATA_FRAME_LEN-1]
                else:
                    for j in range(0, MAX_ONE_DATA_FRAME_LEN):
                        self.Rmdata_self[1 + (i-1)*MAX_ONE_DATA_FRAME_LEN +j] = msgs[i].frame.data[j]
        else:
            for i in range(0, msgs_num):
                for j in range(0, MAX_ONE_DATA_FRAME_LEN):
                    self.Rmdata_self[1 + (i+(self.Rmdata_cur_cnt-1))*MAX_ONE_DATA_FRAME_LEN +j] = msgs[i].frame.data[j]

        self.Rmdata_cur_cnt += msgs_num

        if self.Rmdata_tmt == 0xff:
            if self.Rmdata_cur_cnt == RM_MPPT_FRAME_CNT:
                self.RmDataMpptDisplay(self.Rmdata_self)
        elif self.Rmdata_tmt == 0xfe:
            if self.Rmdata_cur_cnt == RM_BAT_FRAME_CNT:
                self.RmDataBatDisplay(self.Rmdata_self)
        elif self.Rmdata_tmt == 0xfd:
            if self.Rmdata_cur_cnt == RM_WING_FRAME_CNT:
                self.RmDataWingDisplay(self.Rmdata_self)

    def MsgReadThreadFunc(self):
        try:
            while not self._terminated:
                can_num = self._zcan.GetReceiveNum(self._can_handle, ZCAN_TYPE_CAN)
                canfd_num = self._zcan.GetReceiveNum(self._can_handle, ZCAN_TYPE_CANFD)
                if not can_num and not canfd_num:
                    time.sleep(0.005) #wait 5ms  
                    continue

                if can_num:
                    while can_num and not self._terminated:
                        read_cnt = MAX_RCV_NUM if can_num >= MAX_RCV_NUM else can_num
                        can_msgs, act_num = self._zcan.Receive(self._can_handle, read_cnt, MAX_RCV_NUM)
                        if act_num: 
                            #update data
                            can_cmd_ok = 1
                            for i in range(1, act_num-1):
                                if ((can_msgs[i].frame.can_id>>CANID_DT_POS)&0b11111) != ((can_msgs[i-1].frame.can_id>>CANID_DT_POS)&0b11111):
                                    can_cmd_ok = 0
                                    break
                            if can_cmd_ok != 0:
                                can_cmd_ok = (can_msgs[0].frame.can_id>>CANID_DT_POS)&0b11111
                                if can_cmd_ok == DT_REMOTE_RETURN:
                                    self.RmDataUpdata(can_msgs, act_num)

                            self._rx_cnt += act_num 
                            self.strvRxCnt.set(str(self._rx_cnt))
                            self.ViewDataUpdate(can_msgs, act_num, False, False)
                        else:
                            break
                        can_num -= act_num
                if canfd_num:
                    while canfd_num and not self._terminated:
                        read_cnt = MAX_RCV_NUM if canfd_num >= MAX_RCV_NUM else canfd_num
                        canfd_msgs, act_num = self._zcan.ReceiveFD(self._can_handle, read_cnt, MAX_RCV_NUM)
                        if act_num: 
                            #update data
                            self._rx_cnt += act_num 
                            self.strvRxCnt.set(str(self._rx_cnt))
                            self.ViewDataUpdate(canfd_msgs, act_num, True, False)
                        else:
                            break
                        canfd_num -= act_num
        except:
            print("Error occurred while read CAN(FD) data!")

    def ViewDataUpdate(self, msgs, msgs_num, is_canfd=False, is_send=True):
        with self._lock:
            if is_canfd:
                for i in range(msgs_num):
                    if len(self.treeMsg.get_children()) == MAX_DISPLAY:
                        self.treeMsg.delete(self.treeMsg.get_children()[0])
                    self.treeMsg.insert('', 'end', values=self.CANFDMsg2View(msgs[i].frame, is_send))
                    #focus section
                    child_id = self.treeMsg.get_children()[-1]
                    self.treeMsg.focus(child_id)
                    self.treeMsg.selection_set(child_id)
            else:
                for i in range(msgs_num):
                    if len(self.treeMsg.get_children()) == MAX_DISPLAY:
                        self.treeMsg.delete(self.treeMsg.get_children()[0])
                    self.treeMsg.insert('', 'end', values=self.CANMsg2View(msgs[i].frame, is_send))
                    #focus section
                    child_id = self.treeMsg.get_children()[-1]
                    self.treeMsg.focus(child_id)
                    self.treeMsg.selection_set(child_id)

    def PeriodSendIdUpdate(self, is_ext):
        self._cur_id += 1
        if is_ext:
            if self._cur_id > 0x1FFFFFFF:
                self._cur_id = 0
        else:
            if self._cur_id > 0x7FF:
                self._cur_id = 0

    def PeriodSendComplete(self):
        self._is_sending = False
        self.strvSend.set("发送")
        self._send_thread.send_stop()

    def PeriodSend(self):
        if self._is_canfd_msg: 
            ret = self._zcan.TransmitFD(self._can_handle, self._send_msgs, self._send_num)
        else:
            ret = self._zcan.Transmit(self._can_handle, self._send_msgs, self._send_num)
        
        #update transmit display
        self._tx_cnt += ret
        self.strvTxCnt.set(str(self._tx_cnt))
        self.ViewDataUpdate(self._send_msgs, ret, self._is_canfd_msg, True)
        
        if ret != self._send_num:
            self.PeriodSendComplete()
            messagebox.showerror(title="发送报文", message="发送失败！")
            return

        self._send_cnt -= 1
        if self._send_cnt:
            if self._id_increase:
                for i in range(self._send_num):
                    self._send_msgs[i].frame.can_id = self._cur_id
                    self.PeriodSendIdUpdate(self._send_msgs[i].frame.eff)
        else:
            self.PeriodSendComplete()

    def MsgSend(self, msg, is_canfd, num=1, cnt=1, period=0, id_increase=0):
        self._id_increase  = id_increase
        self._send_num     = num if num else 1
        self._send_cnt     = cnt if cnt else 1
        self._is_canfd_msg = is_canfd

        if is_canfd:    
            self._send_msgs = (ZCAN_TransmitFD_Data * self._send_num)()
        else:
            self._send_msgs = (ZCAN_Transmit_Data * self._send_num)()

        self._cur_id = msg.frame.can_id
        for i in range(self._send_num):
            self._send_msgs[i] = msg
            self._send_msgs[i].frame.can_id = self._cur_id
            self.PeriodSendIdUpdate(self._send_msgs[i].frame.eff)

        self._is_sending = True    
        self._send_thread.send_start(period * 0.001)
    def DevInfoRead(self):
        info = self._zcan.GetDeviceInf(self._dev_handle)
        if info != None:
            self.strvHwVer.set(info.hw_version)
            self.strvFwVer.set(info.fw_version)
            self.strvDrVer.set(info.dr_version)
            self.strvInVer.set(info.in_version)
            self.strvCANNum.set(str(info.can_num))
            self.strvSerial.set(info.serial)
            self.strvHwType.set(info.hw_type)

    def DevInfoClear(self):
        self.strvHwVer.set('')
        self.strvFwVer.set('')
        self.strvDrVer.set('')
        self.strvInVer.set('')
        self.strvCANNum.set('')
        self.strvSerial.set('')
        self.strvHwType.set('')
###############################################################################
### Event handers
###############################################################################
    def Form_OnClosing(self):
        if self._isOpen:
            self.btnDevCtrl.invoke()

        self.destroy()

    def BtnOpenDev_Click(self):
        if self._isOpen:
            #Close Channel 
            if self._isChnOpen:
                self.btnCANCtrl.invoke()

            #Close Device
            self._zcan.CloseDevice(self._dev_handle)

            self.DevInfoClear()
            self.strvDevCtrl.set("打开")
            self.cmbDevType["state"] = "readonly"
            self.cmbDevIdx["state"] = "readonly"
            self._isOpen = False
        else:
            self._cur_dev_info = self._dev_info[self.cmbDevType.get()]

            #Open Device
            self._dev_handle = self._zcan.OpenDevice(self._cur_dev_info["dev_type"], 
                                                     self.cmbDevIdx.current(), 0)
            if self._dev_handle == INVALID_DEVICE_HANDLE:
                #Open failed
                messagebox.showerror(title="打开设备", message="打开设备失败！")
                return 
            
            #Update Device Info Display
            self.DevInfoRead()

            self._is_canfd = self._cur_dev_info["chn_info"]["is_canfd"]
            self._res_support = self._cur_dev_info["chn_info"]["sf_res"]
            if self._is_canfd:
                self.cmbMsgCANFD["value"] = ("CAN", "CANFD", "CANFD BRS")
            else:
                self.cmbMsgCANFD["value"] = ("CAN")

            self.strvDevCtrl.set("关闭")
            self.cmbDevType["state"] = tk.DISABLED
            self.cmbDevIdx["state"] = tk.DISABLED
            self._isOpen = True 
        self.ChnInfoUpdate(self._isOpen)
        self.ChnInfoDisplay(self._isOpen)

    def BtnOpenCAN_Click(self):
        if self._isChnOpen:
            #wait read_thread exit
            self._terminated = True
            self._read_thread.join(0.1)

            #stop send thread
            self._send_thread.stop()

            #cancel send
            if self._is_sending:
                self.btnMsgSend.invoke()

            #Close channel
            self._zcan.ResetCAN(self._can_handle)
            self.strvCANCtrl.set("打开")
            self._isChnOpen = False
            self.btnMsgSend["state"] = tk.DISABLED
        else:
            #Initial channel
            if self._res_support: #resistance enable
                ip = self._zcan.GetIProperty(self._dev_handle)
                self._zcan.SetValue(ip, 
                                    str(self.cmbCANChn.current()) + "/initenal_resistance", 
                                    '1' if self.cmbResEnable.current() == 0 else '0')
                self._zcan.ReleaseIProperty(ip)

            #set usbcan-e-u baudrate
            if self._cur_dev_info["dev_type"] in USBCAN_XE_U_TYPE:
                ip = self._zcan.GetIProperty(self._dev_handle)
                self._zcan.SetValue(ip, 
                                    str(self.cmbCANChn.current()) + "/baud_rate", 
                                    self._cur_dev_info["chn_info"]["baudrate"][self.cmbBaudrate.get()])
                self._zcan.ReleaseIProperty(ip)

            #set usbcanfd clock 
            if self._cur_dev_info["dev_type"] in USBCANFD_TYPE:
                ip = self._zcan.GetIProperty(self._dev_handle)
                self._zcan.SetValue(ip, str(self.cmbCANChn.current()) + "/clock", "60000000")
                self._zcan.ReleaseIProperty(ip)
            
            chn_cfg = ZCAN_CHANNEL_INIT_CONFIG()
            chn_cfg.can_type = ZCAN_TYPE_CANFD if self._is_canfd else ZCAN_TYPE_CAN
            if self._is_canfd:
                chn_cfg.config.canfd.mode = self.cmbCANMode.current()
                chn_cfg.config.canfd.abit_timing = self._cur_dev_info["chn_info"]["baudrate"][self.cmbBaudrate.get()]
                chn_cfg.config.canfd.dbit_timing = self._cur_dev_info["chn_info"]["data_baudrate"][self.cmbDataBaudrate.get()]
            else:
                chn_cfg.config.can.mode = self.cmbCANMode.current()
                if self._cur_dev_info["dev_type"] in USBCAN_I_II_TYPE:
                    brt = self._cur_dev_info["chn_info"]["baudrate"][self.cmbBaudrate.get()]
                    chn_cfg.config.can.timing0 = brt["timing0"] 
                    chn_cfg.config.can.timing1 = brt["timing1"]
                    chn_cfg.config.can.acc_code = 0
                    chn_cfg.config.can.acc_mask = 0xFFFFFFFF

            self._can_handle = self._zcan.InitCAN(self._dev_handle, self.cmbCANChn.current(), chn_cfg)
            if self._can_handle == INVALID_CHANNEL_HANDLE:
                messagebox.showerror(title="打开通道", message="初始化通道失败!")
                return 

            ret = self._zcan.StartCAN(self._can_handle)
            if ret != ZCAN_STATUS_OK: 
                messagebox.showerror(title="打开通道", message="打开通道失败!")
                return 

            #start send thread
            self._send_thread = PeriodSendThread(self.PeriodSend)
            self._send_thread.start()

            #start receive thread
            self._terminated = False
            self._read_thread = threading.Thread(None, target=self.MsgReadThreadFunc)
            self._read_thread.start()

            self.strvCANCtrl.set("关闭")
            self._isChnOpen = True 
            self.btnMsgSend["state"] = tk.NORMAL
        self.ChnInfoDisplay(not self._isChnOpen)

    def BtnClrCnt_Click(self):
        self._tx_cnt = 0
        self._rx_cnt = 0
        self._view_cnt = 0
        self.strvTxCnt.set("0")
        self.strvRxCnt.set("0")
        # self.treeMsg
        for item in self.treeMsg.get_children():
            self.treeMsg.delete(item)

    def CmbMsgFormatUpdate(self, *args):
        if self.cmbMsgFormat.current() == 0: #Data Frame
            if self._is_canfd:
                self.cmbMsgCANFD["value"] = ("CAN", "CANFD", "CANFD BRS")
            else:
                self.cmbMsgCANFD["value"] = ("CAN")
                self.cmbMsgCANFD.current(0)
        else: #Remote Frame
            self.cmbMsgCANFD["value"] = ("CAN")
            self.cmbMsgCANFD.current(0)

    def CmbMsgCANFDUpdate(self, *args):
        tmp = self.cmbMsgLen.current()
        self.cmbMsgLen["value"] = tuple([self.__dlc2len(i) for i in range(16 if self.cmbMsgCANFD.current() else 9)]) 
        if tmp >= len(self.cmbMsgLen["value"]):
            self.cmbMsgLen.current(len(self.cmbMsgLen["value"]) - 1)            

    def TmtParChangeEvent(self, *args):
        #self.entryMsgData.delete(0, "end")
        if self.cmbDataType.current() == 2:
            self.entryMsgData.delete(0, "end")
            if self.cmbTmt.current() == 0:
                self.entryMsgData.insert(0, "01 00 00 00 00 00 00")
                #self.entryMsgData.insert(2, " " + "{:02X}".format(self.cmbTmtPar.current()))
            elif self.cmbTmt.current() == 1:
                self.entryMsgData.insert(0, "02 00 00 00 00 00 00")
            elif self.cmbTmt.current() == 2:
                self.entryMsgData.insert(0, "04 00 00 00 00 00 00")
            elif self.cmbTmt.current() == 3:
                self.entryMsgData.insert(0, "07 00 00 00 00 00 00")

            self.entryMsgData.insert(2, " " + "{:02X}".format(self.cmbTmtPar.current()))

    def TmtTypeChangeEvent(self, *args):
        self.entryMsgData.delete(0, "end")
        #self.cmbTmtPar.set("")
        if self.cmbDataType.current() == 0:
            self.RemoteDataWindowCreate(self.cmbTmt.current())
            if self.cmbTmt.current() == 0:
                self.entryMsgData.insert(0, "FF 00 00 00 00 00 00 00")
            elif self.cmbTmt.current() == 1:
                self.entryMsgData.insert(0, "FE 00 00 00 00 00 00 00")
            elif self.cmbTmt.current() == 2:
                self.entryMsgData.insert(0, "FD 00 00 00 00 00 00 00")
        elif self.cmbDataType.current() == 1:
            if self.cmbTmt.current() == 0:
                self.entryMsgData.insert(0, "F9 00 00 00 00 00 00 00")
            elif self.cmbTmt.current() == 1:
                self.entryMsgData.insert(0, "F8 00 00 00 00 00 00 00")
            elif self.cmbTmt.current() == 2:
                self.entryMsgData.insert(0, "F7 00 00 00 00 00 00 00")
        elif self.cmbDataType.current() == 2:
            if self.cmbTmt.current() == 0:
                self.entryMsgData.insert(0, "01 00 00 00 00 00 00 00")
                self.cmbTmtPar["value"] = (
                            "MPPT1控制关", 
                            "MPPT1控制开", 
                            "MPPT2控制关", 
                            "MPPT2控制开", 
                            "MPPT3控制关", 
                            "MPPT3控制开", 
                            "MPPT4控制关", 
                            "MPPT4控制开", 
                            "MPPT5控制关", 
                            "MPPT5控制开", 
                            "MPPT6控制关", 
                            "MPPT6控制开", 
                            "MPPT7控制关", 
                            "MPPT7控制开", 
                            "MPPT8控制关", 
                            "MPPT8控制开", 
                            "MPPT9控制关", 
                            "MPPT9控制开", 
                            )
            elif self.cmbTmt.current() == 1:
                self.entryMsgData.insert(0, "02 00 00 00 00 00 00 00")
                self.cmbTmtPar["value"] = (
                            "欠压保护禁止", 
                            "欠压保护使能", 
                            "自主加电禁止",
                            "自主加电使能",
                            "放电开关断开",
                            "放电开关接通",
                            "充电电流1档",
                            "充电电流2档",
                            "充电电压1档",
                            "充电电压2档",
                            "充电电压3档",
                            )
            elif self.cmbTmt.current() == 2:
                self.entryMsgData.insert(0, "04 00 00 00 00 00 00 00")
                self.cmbTmtPar["value"] = (
                            "QV前端A断开", 
                            "QV前端A接通", 
                            "QV前端B断开",
                            "QV前端B接通",
                            "霍尔组件断开",
                            "霍尔组件接通",
                            "发射相控阵A断开",
                            "发射相控阵A接通",
                            "发射相控阵B断开",
                            "发射相控阵B接通",
                            "接收相控阵断开",
                            "接收相控阵接通",
                            "飞轮X断开",
                            "飞轮X接通",
                            "飞轮Y断开",
                            "飞轮Y接通",
                            "飞轮Z断开",
                            "飞轮Z接通",
                            "飞轮S断开",
                            "飞轮S接通",
                            )
            elif self.cmbTmt.current() == 3:
                self.entryMsgData.insert(0, "07 00 00 00 00 00 00 00")
                self.cmbTmtPar["value"] = (
                            "帆板天线解锁配电断开", 
                            "帆板天线解锁配电接通", 
                            "解锁帆板",
                            "解锁举升机构",
                            "解锁天线A",
                            "解锁天线B",
                            )
            self.cmbTmtPar.current(0)
        elif self.cmbDataType.current() == 3:
            self.entryMsgData.insert(0, "F5 00 00 00 00 00 00 00")

    def DataTypeChangeEvent(self, *args):
        print(456)
        if isinstance(self.entryMsgID.tmp_value, str):
            self.entryMsgID.tmp_value = int(self.entryMsgID.tmp_value, 16)

        self.entryMsgID.tmp_value &= ~(0b11111<<20)
        self.entryMsgData.delete(0, "end")
        self.cmbTmtPar["value"] = {}
        self.cmbTmtPar.set("")
        self.CloseRemoteWin()

        if self.cmbDataType.current() == 0:
            self.cmbTmt["value"] = ("MPPT", "BAT", "WING")
            self.entryMsgData.insert(0, "FF 00 00 00 00 00 00 00")
            self.RemoteDataWindowCreate(0)
        elif self.cmbDataType.current() == 1:
            self.entryMsgID.tmp_value |= 0b1<<20
            self.cmbTmt["value"] = ("复位CANAB", "复位CANA", "复位CANB")
            self.entryMsgData.insert(0, "F9 00 00 00 00 00 00 00")
        elif self.cmbDataType.current() == 2:
            self.entryMsgID.tmp_value |= 0b10<<20
            self.cmbTmt["value"] = ("MPPT", "BAT", "WING1","WING2")
            self.cmbTmtPar["value"] = (
                                        "MPPT1控制关", 
                                        "MPPT1控制开", 
                                        "MPPT2控制关", 
                                        "MPPT2控制开", 
                                        "MPPT3控制关", 
                                        "MPPT3控制开", 
                                        "MPPT4控制关", 
                                        "MPPT4控制开", 
                                        "MPPT5控制关", 
                                        "MPPT5控制开", 
                                        "MPPT6控制关", 
                                        "MPPT6控制开", 
                                        "MPPT7控制关", 
                                        "MPPT7控制开", 
                                        "MPPT8控制关", 
                                        "MPPT8控制开", 
                                        "MPPT9控制关", 
                                        "MPPT9控制开", 
                                       )
            self.cmbTmtPar.current(0)
            self.entryMsgData.insert(0, "01 00 00 00 00 00 00 00")
        elif self.cmbDataType.current() == 3:
            self.entryMsgID.tmp_value |= 0b11<<20
            self.cmbTmt["value"] = ("获取96字节")
            self.entryMsgData.insert(0, "F5 00 00 00 00 00 00 00")
        elif self.cmbDataType.current() == 4:
            self.entryMsgID.tmp_value |= 0b10100<<20
            self.cmbTmt["value"] = ("广播96字节")

        self.cmbTmt.current(0)
        self.entryMsgID.tmp_value = "{:X}".format(self.entryMsgID.tmp_value)
        self.entryMsgID.delete(0, "end")
        self.entryMsgID.insert(0, self.entryMsgID.tmp_value )

    def CanIdChangeEvent(self, *args):
        print(123)
        if isinstance(self.entryMsgID.tmp_value, str):
            self.entryMsgID.tmp_value = int(self.entryMsgID.tmp_value, 16)
        self.entryMsgID.tmp_value &= (0b11111<<20)

        if self.cmbProvity.current() == 0:
            self.entryMsgID.tmp_value |= 0b11<<26
        elif self.cmbProvity.current() == 1:
            self.entryMsgID.tmp_value |= 0b110<<26

        if self.cmbBusFlag.current() == 1:
            self.entryMsgID.tmp_value |= 0b1<<25
        '''
        if self.cmbDataType.current() == 1:
            self.entryMsgID.tmp_value += 0b1<<20
        elif self.cmbDataType.current() == 2:
            self.entryMsgID.tmp_value += 0b10<<20
        elif self.cmbDataType.current() == 3:
            self.entryMsgID.tmp_value += 0b11<<20
        elif self.cmbDataType.current() == 4:
            self.entryMsgID.tmp_value += 0b10100<<20
        '''
        if self.cmbDa.current() == 1:
            self.entryMsgID.tmp_value |= 0b10000<<15
        elif self.cmbDa.current() == 2:
            self.entryMsgID.tmp_value |= 0b10001<<15
        elif self.cmbDa.current() == 3:
            self.entryMsgID.tmp_value |= 0b11111<<15

        if self.cmbSa.current() == 1:
            self.entryMsgID.tmp_value |= 0b10000<<10
        elif self.cmbSa.current() == 2:
            self.entryMsgID.tmp_value |= 0b10001<<10

        self.entryMsgID.tmp_value |= self.cmbFt.current()<<8

        self.entryMsgID.tmp_value |= (int)(self.entryFc.get())%0x100

        self.entryMsgID.tmp_value = "{:X}".format(self.entryMsgID.tmp_value)
        self.entryMsgID.delete(0, "end")
        self.entryMsgID.insert(0, self.entryMsgID.tmp_value )

    def BtnSendMsg_Click(self): 
        if not self._is_sending:
            is_canfd_msg = True if self.cmbMsgCANFD.current() > 0 else False
            if is_canfd_msg:
                msg = ZCAN_TransmitFD_Data()
            else:
                msg = ZCAN_Transmit_Data()

            msg.transmit_type = self.cmbSendType.current()
            try:
                msg.frame.can_id = int(self.entryMsgID.get(), 16)
            except:
                msg.frame.can_id = 0
            msg.frame.rtr = self.cmbMsgFormat.current()
            msg.frame.eff = self.cmbMsgType.current()

            if not is_canfd_msg:
                msg.frame.can_dlc = self.cmbMsgLen.current()
                msg_len = msg.frame.can_dlc
            else:
                msg.frame.brs = 1 if self.cmbMsgCANFD.current() == 2 else 0
                msg.frame.len = self.__dlc2len(self.cmbMsgLen.current())
                msg_len = msg.frame.len

            data = self.entryMsgData.get().split(' ')
            for i in range(msg_len):
                if i < len(data):
                    try:
                        msg.frame.data[i] = int(data[i], 16)
                    except:
                        msg.frame.data[i] = 0
                else:
                    msg.frame.data[i] = 0

            try:
                msg_num = int(self.entryMsgNum.get())
                msg_cnt = int(self.entryMsgCnt.get())
                period  = int(self.entryMsgPeriod.get())
            except:
                msg_num = 1
                msg_cnt = 1
                period  = 1
            self.MsgSend(msg, is_canfd_msg, msg_num, msg_cnt, period, self.varIDInc.get())
            self.strvSend.set("停止发送")
        else:
            self.PeriodSendComplete()

if __name__ == "__main__":
    demo = ZCAN_Demo()
    demo.mainloop()