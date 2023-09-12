using AWGTestCerter.Common;
using AWGTestCerter.Remote;
using CommonLibrary.Enum;
using CommonLibrary.EnumHelper;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows;
using System.Windows.Media;


namespace AWGTestCerter
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window, INotifyPropertyChanged
    {
        public MainWindow()
        {
            InitializeComponent();
            this.DataContext = this;
            Init();
        }
        private int _1652ASocketID = 0;
        private int _1652BSocketID = 1;
        private int _nrxSocketID = 4;
        private int _mxr604aSocketID = 5;
        private int _mso254aSocketID = 6;
        private int _fsw50SocketID = 7;
        private int _fsmr50SocketID = 8;

        private BackgroundWorker backgroundWorkerPowerCali;      //DC-2GHz幅度校准子任务
        /// <summary>
        /// 初始化
        /// </summary>
        public void Init()
        {
            ChannelComoBox = CHANNEL.CHA_1;
            Check_NRX = true;
            Check_Awg1652B = true;
            /*
            Check_MXR604A = true;
           
            Check_FSW50 = true;
            Check_FSMR50 = true;
            Check_MSO254A = true;
            */

            return;
        }

        /// <summary>
        /// 滤波器类型
        /// </summary>
        private CHANNEL _channelComoBox;

        /// <summary>
        /// 滤波器类型
        /// </summary>
        public CHANNEL ChannelComoBox
        {
            get { return _channelComoBox; }
            set
            {
                _channelComoBox = value;
                OnPropertyChanged("ChannelComoBox");
            }
        }

        public Dictionary<CHANNEL, string> ChannelComoBoxDictionary
        {
            get
            {
               return System.Enum.GetValues(typeof(CHANNEL)).Cast<CHANNEL>().ToDictionary(item => item, item => (EnumHelper.GetEnumDescription(item)).ToString());
            }
        }

        /// <summary>
        ///
        /// </summary>
        private bool _check_Awg1652B;

        /// <summary>
        ///  
        /// </summary>
        public bool Check_Awg1652B
        {
            get
            {
                return _check_Awg1652B;
            }
            set
            {
                _check_Awg1652B = value;            
                OnPropertyChanged("Check_Awg1652B");
            }
        }


        /// <summary>
        ///
        /// </summary>
        private bool _check_NRX;

        /// <summary>
        ///  
        /// </summary>
        public bool Check_NRX
        {
            get
            {
                return _check_NRX;
            }
            set
            {
                _check_NRX = value;
                if(_check_NRX == true)
                {
                    Check_MXR604A = false;
                    Check_MSO254A = false;
                    Check_FSW50 = false;
                    Check_FSMR50 = false;
                }
                OnPropertyChanged("Check_NRX");
            }
        }


        /// <summary>
        ///
        /// </summary>
        private bool _check_MXR604A;

        /// <summary>
        ///  
        /// </summary>
        public bool Check_MXR604A
        {
            get
            {
                return _check_MXR604A;
            }
            set
            {
                _check_MXR604A = value;
                if (_check_MXR604A == true)
                {
                    Check_NRX = false;
                    Check_MSO254A = false;
                    Check_FSW50 = false;
                    Check_FSMR50 = false;
                }
                OnPropertyChanged("Check_MXR604A");
            }
        }

        /// <summary>
        ///
        /// </summary>
        private bool _check_MSO254A;

        /// <summary>
        ///  
        /// </summary>
        public bool Check_MSO254A
        {
            get
            {
                return _check_MSO254A;
            }
            set
            {
                _check_MSO254A = value;
                if (_check_MSO254A == true)
                {
                    Check_NRX = false;
                    Check_MXR604A = false;
                    Check_FSW50 = false;
                    Check_FSMR50 = false;
                }
                OnPropertyChanged("Check_MSO254A");
            }
        }

        /// <summary>
        ///
        /// </summary>
        private bool _check_FSW50;

        /// <summary>
        ///  
        /// </summary>
        public bool Check_FSW50
        {
            get
            {
                return _check_FSW50;
            }
            set
            {
                _check_FSW50 = value;
                if (_check_FSW50 == true)
                {
                    Check_NRX = false;
                    Check_MXR604A = false;
                    Check_MSO254A = false;
                    Check_FSMR50 = false;
                }
                OnPropertyChanged("Check_FSW50");
            }
        }

        /// <summary>
        ///
        /// </summary>
        private bool _check_FSMR50;

        /// <summary>
        ///  
        /// </summary>
        public bool Check_FSMR50
        {
            get
            {
                return _check_FSMR50;
            }
            set
            {
                _check_FSMR50 = value;
                if (_check_FSMR50 == true)
                {
                    Check_NRX = false;
                    Check_MXR604A = false;
                    Check_MSO254A = false;
                    Check_FSW50 = false;
                }
                OnPropertyChanged("Check_FSMR50");
            }
        }

        /// <summary>
        /// 创建幅度校准子任务
        /// </summary>
        public void CreateBackgroundWorkerPowerCali()
        {
            backgroundWorkerPowerCali = new BackgroundWorker();
            backgroundWorkerPowerCali.WorkerReportsProgress = true;
            backgroundWorkerPowerCali.WorkerSupportsCancellation = true;
            backgroundWorkerPowerCali.DoWork += backgroundWorkerPowerCali_DoWork;
            backgroundWorkerPowerCali.ProgressChanged += backgroundWorkerPowerCali_ProgressChanged;
            backgroundWorkerPowerCali.RunWorkerCompleted += backgroundWorkerPowerCali_Completed;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="ReadData"></param>
        private void Clear(double[,] array)
        {
            int i, j;
            int row = array.GetLength(0);//获取维数，这里指行数
            int col = array.GetLength(1);//获取指定维度中的元素个数，这里也就是列数了
            for (i = 0; i < row; i++)
            {
                for (j = 0; j < col; j++)
                {
                    array[i, j] = 0 - 200.0;
                }
            }
            return;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="couple"></param>
        /// <param name="vpp"></param>
        /// <param name="maxVpp"></param>
        /// <param name="minVpp"></param>
        private void GetMaxMin_Vpp(int couple, double vpp, out double maxVpp, out double minVpp)
        {
            double deltaVpp = 0;
            maxVpp = 0;
            minVpp = 0;
            if (couple == TywCommon.COUPLETYPE_DC)
            {
                deltaVpp = 0.05 * vpp + 0.035;//5%,35mvpp
                maxVpp = vpp + deltaVpp;
                minVpp = vpp - deltaVpp;
            }
            if (couple == TywCommon.COUPLETYPE_AC)
            {
                if (vpp < 0.70001)
                {
                    deltaVpp = 0.05 * vpp + 0.035;//5%,35mvpp
                    maxVpp = vpp + deltaVpp;
                    minVpp = vpp - deltaVpp;
                }
                else
                {
                    deltaVpp = 0.05 * vpp + 0.05;//5%,50mvpp
                    maxVpp = vpp + deltaVpp;
                    minVpp = vpp - deltaVpp;
                }
            }
            return;
        }
        /// <summary>
        /// DC-2GHz幅度校准子任务
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void backgroundWorkerPowerCali_DoWork(object sender, DoWorkEventArgs e)
        {
            int who = (int)ChannelComoBox;
            double yScale = 0.0;
            double xScale = 0.0;
            bool []IsConnect = new bool [10];
            int i, j, FreqCount = 6, VppCount = 5;
            double[] DC_FreqSlot = new double[FreqCount];
            double[] DC_VppSlot = new double[VppCount];
            double[] AC_FreqSlot = new double[FreqCount];
            double[] AC_VppSlot = new double[VppCount];
            double[,] DC_ReadData = new double[FreqCount, VppCount];
            double[,] AC_ReadData = new double[FreqCount, VppCount];
            double maxValue = 10;
            double minValue = 0 - 50;
            double d, centerFreq = 0.0;
            string cmdStr = "";
            string saveFileName = @"D:\test.csv";
            int chgFreqSleepTime = 500;
            int time_out = 400;
            int Count = FreqCount * 2;
            BackgroundWorker _worker = sender as BackgroundWorker;

            Clear(DC_ReadData);
            Clear(AC_ReadData);

            DC_FreqSlot[0] = 10e6; DC_FreqSlot[1] = 100e6; DC_FreqSlot[2] = 500e6;
            DC_FreqSlot[3] = 1e9;  DC_FreqSlot[4] = 1.3e9; DC_FreqSlot[5] = 2e9;

            DC_VppSlot[0] = 0.05; DC_VppSlot[1] = 0.1; DC_VppSlot[2] = 0.5;
            DC_VppSlot[3] = 0.8;  DC_VppSlot[4] = 1.0;

            AC_FreqSlot[0] = 10e6; AC_FreqSlot[1] = 100e6; AC_FreqSlot[2] = 500e6;
            AC_FreqSlot[3] = 1e9;  AC_FreqSlot[4] = 1.3e9; AC_FreqSlot[5] = 2e9;

            AC_VppSlot[0] = 0.05; AC_VppSlot[1] = 0.5; AC_VppSlot[2] = 0.7;
            AC_VppSlot[3] = 1.0;  AC_VppSlot[4] = 1.5;

            if (_worker.CancellationPending == true)
            {
                e.Cancel = true;
            }
            else
            {
                
                if (Check_NRX == true)
                {
                    App.Current.Dispatcher.Invoke((Action)(() =>
                    {
                        //LAN连接建立:
                        IsConnect[_1652BSocketID] = MyClientSocketLAN.OnConnectSSocket(Awg1652B_Ip.Text, "4000",_1652BSocketID);
                        IsConnect[_nrxSocketID] = MyClientSocketLAN.OnConnectSSocket  (NRX_Ip.Text, NRX_Port.Text, _nrxSocketID);
                    }));
                    if (IsConnect[_1652BSocketID] == false)
                    {
                        MessageBox.Show("1652B连接失败!");
                        return;
                    }
                    if (IsConnect[_nrxSocketID] == false )
                    {
                        MessageBox.Show("NRX功率计仪器连接失败!");
                        return;
                    }
                    SetShowMsg("NRX读取功率中...", 1);
                    saveFileName = string.Format("D:\\AWG1652_CH{0}_dBm.csv", (who+1).ToString());
                    //发程控命令，初始化1652B                  
                    cmdStr = string.Format("AWGC:RF:CH{0} ON", (who + 1).ToString());
                    MyClientSocketLAN.OnSendMsg(cmdStr, _1652BSocketID);//射频通道开关
                    Thread.Sleep(500);
                    MyClientSocketLAN.OnSendMsg("AWGC:RUN ON", _1652BSocketID);//播放使能
                    Thread.Sleep(500);
                    cmdStr = string.Format("AWGC:CHA {0}", who.ToString());
                    MyClientSocketLAN.OnSendMsg(cmdStr, _1652BSocketID);//通道切换
                    MyClientSocketLAN.OnSendMsg("AWGC:MODE DDFS", _1652BSocketID);//模式切换
                    Thread.Sleep(500);
                    MyClientSocketLAN.OnSendMsg("DDFS:WAVE SIN", _1652BSocketID);//波形为正弦
                    Thread.Sleep(500);
                    MyClientSocketLAN.OnSendMsg("DDFS:SIN:FREQ 10MHz", _1652BSocketID);
                    Thread.Sleep(500);
                    MyClientSocketLAN.OnSendMsg("CHS:OUT:POW 0.5", _1652BSocketID);
                    Thread.Sleep(500);

                    //NRX功率计载波频率
                    MyClientSocketLAN.SetNRXCarrierFreq(50000000.0, _nrxSocketID);
                    Thread.Sleep(chgFreqSleepTime);                                   
                    //------------------： DC耦合
                    //------------------： DC耦合    
                    MyClientSocketLAN.OnSendMsg("CHS:OUT:COUP DC", _1652BSocketID);
                    MyClientSocketLAN.OnSendMsg("CHS:OUT:DCVOF 0.0", _1652BSocketID);
                    Thread.Sleep(300);
                    time_out = 400;
                    for (i = 0; i < FreqCount; i++)
                    {
                        centerFreq = DC_FreqSlot[i];
                        //1652B输出频率
                        cmdStr = string.Format("DDFS:SIN:FREQ {0}", centerFreq.ToString("0.000"));
                        MyClientSocketLAN.OnSendMsg(cmdStr, _1652BSocketID);//信号频率
                        Thread.Sleep(200);
                        //NRX功率计载波频率
                        MyClientSocketLAN.SetNRXCarrierFreq(centerFreq, _nrxSocketID);
                        Thread.Sleep(chgFreqSleepTime);

                        //发命令控制1652B输出幅度
                        for (j = 0; j < VppCount; j++)
                        {                         
                            cmdStr = string.Format("AWGC:CHA {0}", who.ToString());
                            MyClientSocketLAN.OnSendMsg(cmdStr, _1652BSocketID);//通道切换
                            Thread.Sleep(200);
                            cmdStr = string.Format("CHS:OUT:POW {0}", DC_VppSlot[j].ToString("0.000000"));
                            MyClientSocketLAN.OnSendMsg(cmdStr, _1652BSocketID);//发指令
                            Thread.Sleep(500);//等待回读
                            MyClientSocketLAN.GetNRXPower1(out d, time_out, _nrxSocketID);
                            if( i == 0 && j == 0)
                            {
                                Thread.Sleep(2000);
                                MyClientSocketLAN.GetNRXPower1(out d, time_out, _nrxSocketID);
                            }
                            DC_ReadData[i, j] = d;
                        }
                        _worker.ReportProgress(100 * i / Count);//显示校准进度        
                    }
                    //------------------： AC耦合
                    //------------------： AC耦合 
                    MyClientSocketLAN.OnSendMsg("CHS:OUT:COUP AC", _1652BSocketID);
                    MyClientSocketLAN.OnSendMsg("CHS:OUT:DCVOF 0.0", _1652BSocketID);
                    Thread.Sleep(300);
                    time_out = 400;
                    for (i = 0; i < FreqCount; i++)
                    {
                        centerFreq = DC_FreqSlot[i];
                        //1652B输出频率
                        cmdStr = string.Format("DDFS:SIN:FREQ {0}", centerFreq.ToString("0.000"));
                        MyClientSocketLAN.OnSendMsg(cmdStr, _1652BSocketID);
                        Thread.Sleep(200);
                        //NRX功率计载波频率
                        MyClientSocketLAN.SetNRXCarrierFreq(centerFreq, _nrxSocketID);
                        Thread.Sleep(chgFreqSleepTime);
                        //发命令控制1652B输出幅度
                        for (j = 0; j < VppCount; j++)
                        {
                            cmdStr = string.Format("AWGC:CHA {0}", who.ToString());
                            MyClientSocketLAN.OnSendMsg(cmdStr, _1652BSocketID);//通道切换
                            Thread.Sleep(200);
                            cmdStr = string.Format("CHS:OUT:POW {0}", AC_VppSlot[j].ToString("0.000000"));
                            MyClientSocketLAN.OnSendMsg(cmdStr, _1652BSocketID);//发指令
                            Thread.Sleep(500);//等待回读
                            MyClientSocketLAN.GetNRXPower1(out d, time_out, _nrxSocketID);
                            if (i == 0 && j == 0)
                            {
                                Thread.Sleep(2000);
                                MyClientSocketLAN.GetNRXPower1(out d, time_out, _nrxSocketID);
                            }
                            AC_ReadData[i, j] = d;
                        }
                        _worker.ReportProgress(100 * (i+ FreqCount) / Count);//显示校准进度       
                    }
                    ParseData.ParseData.DeleteFile(saveFileName);
                    using (FileStream fw = new FileStream(saveFileName, FileMode.Create))
                    {
                        StreamWriter sw = new StreamWriter(fw);//写.csv或.txt文件
                        cmdStr = string.Format("CoupleType,Freq(MHz),SetValue(Vpp/dBm),CH{0},MaxValue(dBm),MinValue(dBm)\r\n", (who+1).ToString());
                        sw.Write(cmdStr);
                        //DC
                        sw.Write("DC\r\n");
                        for (i = 0; i < FreqCount; i++)
                        {
                            for (j = 0; j < VppCount; j++)
                            {
                                GetMaxMin_Vpp(TywCommon.COUPLETYPE_DC, DC_VppSlot[j], out maxValue, out minValue);
                                maxValue = ConvertFunc.VppToDbm(maxValue);
                                minValue = ConvertFunc.VppToDbm(minValue);
                                cmdStr = string.Format(" ,{0}MHz,{1},{2},{3},{4}\r\n", (    DC_FreqSlot[i] / 1000000).ToString("0.0"),
                                                                                            DC_VppSlot[j].ToString() + "Vpp" + "(" + ConvertFunc.VppToDbm(DC_VppSlot[j]).ToString("0.00") + ")",
                                                                                            DC_ReadData[i, j].ToString("0.00"),
                                                                                            maxValue.ToString("0.00"),
                                                                                            minValue.ToString("0.00"));                                                                                           ; 
                                sw.Write(cmdStr);
                            }
                        }
                        //AC
                        sw.Write("AC\r\n");
                        for (i = 0; i < FreqCount; i++)
                        {
                            for (j = 0; j < VppCount; j++)
                            {
                                GetMaxMin_Vpp(TywCommon.COUPLETYPE_AC, AC_VppSlot[j], out maxValue, out minValue);
                                maxValue = ConvertFunc.VppToDbm(maxValue);
                                minValue = ConvertFunc.VppToDbm(minValue);
                                cmdStr = string.Format(" ,{0}MHz,{1},{2},{3},{4}\r\n", (      AC_FreqSlot[i] / 1000000).ToString("0.0"),
                                                                                              AC_VppSlot[j].ToString() + "Vpp"+"(" + ConvertFunc.VppToDbm(AC_VppSlot[j]).ToString("0.00") + "dBm)",
                                                                                              AC_ReadData[i, j].ToString("0.00"),
                                                                                              maxValue.ToString("0.00"),
                                                                                              minValue.ToString("0.00")); 
                                sw.Write(cmdStr);
                            }
                            _worker.ReportProgress(100 *(i + FreqCount) / Count);//显示校准进度  
                        }
                        sw.Close();//不要遗漏，否则会导致最后几个数据没有写入文件               
                        fw.Close();

                    }



                }


                if (Check_MXR604A == true)
                {
                    App.Current.Dispatcher.Invoke((Action)(() =>
                    {
                        //LAN连接建立:
                        IsConnect[_1652BSocketID] = MyClientSocketLAN.OnConnectSSocket(Awg1652B_Ip.Text, "4000", _1652BSocketID);
                        IsConnect[_mxr604aSocketID] = MyClientSocketLAN.OnConnectSSocket(MXR604A_Ip.Text, MXR604A_Port.Text, _mxr604aSocketID);
                    }));
                    if (IsConnect[_1652BSocketID] == false)
                    {
                        MessageBox.Show("1652B连接失败!");
                        return;
                    }
                    if (IsConnect[_mxr604aSocketID] == false)
                    {
                        MessageBox.Show("MXR604A示波器连接失败!");
                        return;
                    }
                    SetShowMsg("MXR604A读取中...", 1);
                    saveFileName = string.Format("D:\\AWG1652_CH{0}_Vpp.csv", (who + 1).ToString());
                    //发程控命令，初始化1652B                  
                    cmdStr = string.Format("AWGC:RF:CH{0} ON", (who + 1).ToString());
                    MyClientSocketLAN.OnSendMsg(cmdStr, _1652BSocketID);//射频通道开关
                    Thread.Sleep(500);
                    MyClientSocketLAN.OnSendMsg("AWGC:RUN ON", _1652BSocketID);//播放使能
                    Thread.Sleep(500);
                    cmdStr = string.Format("AWGC:CHA {0}", who.ToString());
                    MyClientSocketLAN.OnSendMsg(cmdStr, _1652BSocketID);//通道切换
                    MyClientSocketLAN.OnSendMsg("AWGC:MODE DDFS", _1652BSocketID);//模式切换
                    Thread.Sleep(500);
                    MyClientSocketLAN.OnSendMsg("DDFS:WAVE SIN", _1652BSocketID);//波形为正弦
                    Thread.Sleep(500);
                    MyClientSocketLAN.OnSendMsg("DDFS:SIN:FREQ 10MHz", _1652BSocketID);
                    Thread.Sleep(500);
                    MyClientSocketLAN.OnSendMsg("CHS:OUT:POW 0.5", _1652BSocketID);
                    Thread.Sleep(500);
                    //示波器初始化
                    MyClientSocketLAN.RstMXR604A(_mxr604aSocketID);
                    MyClientSocketLAN.OnSendMsg(":CHAN1:INP DC50", _mxr604aSocketID); //50欧姆阻抗
                    Thread.Sleep(1000);
                    MyClientSocketLAN.InitMXR604AOneChannel(1, _mxr604aSocketID);
                                                                            
                    //------------------： DC耦合
                    //------------------： DC耦合    
                    MyClientSocketLAN.OnSendMsg("CHS:OUT:COUP DC", _1652BSocketID);
                    MyClientSocketLAN.OnSendMsg("CHS:OUT:DCVOF 0.0", _1652BSocketID);
                    Thread.Sleep(300);
                    time_out = 400;
                    for (i = 0; i < FreqCount; i++)
                    {
                        centerFreq = DC_FreqSlot[i];
                        //1652B输出频率
                        cmdStr = string.Format("DDFS:SIN:FREQ {0}", centerFreq.ToString("0.000"));
                        MyClientSocketLAN.OnSendMsg(cmdStr, _1652BSocketID);//信号频率
                        Thread.Sleep(200);
                        //调整示波器
                        xScale = (1.0 / centerFreq) / 2.5;
                        MyClientSocketLAN.SetMXR604ATime(xScale, _mxr604aSocketID);//调整时基
                        //MyClientSocketLAN.GetMXR604ATime(out xScale, _mxr604aSocketID);
                        Thread.Sleep(500);                      
                        //发命令控制1652B输出幅度
                        for (j = 0; j < VppCount; j++)
                        {
                            if (i == 0 && j == 0)
                            {
                                MyClientSocketLAN.OnSendMsg("CHS:OUT:POW 0.05", _1652BSocketID);//发指令
                                Thread.Sleep(500);
                            }                           
                            cmdStr = string.Format("AWGC:CHA {0}", who.ToString());
                            MyClientSocketLAN.OnSendMsg(cmdStr, _1652BSocketID);//通道切换
                            Thread.Sleep(200);
                            cmdStr = string.Format("CHS:OUT:POW {0}", DC_VppSlot[j].ToString("0.000000"));
                            MyClientSocketLAN.OnSendMsg(cmdStr, _1652BSocketID);//发指令
                            if (j == 0) yScale = DC_VppSlot[j] / 4.0;
                            else        yScale = DC_VppSlot[j] / 6.0;
                            MyClientSocketLAN.SetMXR604AYScale(yScale, 1, _mxr604aSocketID);//调整垂直分辨率
                            Thread.Sleep(1000);//等待回读
                            MyClientSocketLAN.GetMXR604AVpp(out d, 1, _mxr604aSocketID);
                            if ( i== 0 && j == 0)
                            {
                                Thread.Sleep(2000);//等待回读
                                MyClientSocketLAN.GetMXR604AVpp(out d, 1, _mxr604aSocketID);
                                Thread.Sleep(2000);//等待回读
                                MyClientSocketLAN.GetMXR604AVpp(out d, 1, _mxr604aSocketID);
                            }
                            DC_ReadData[i, j] = d;
                        }
                        _worker.ReportProgress(100 * i / Count);//显示校准进度        
                    }

                    //------------------： AC耦合
                    //------------------： AC耦合    
                    MyClientSocketLAN.OnSendMsg("CHS:OUT:COUP AC", _1652BSocketID);
                    MyClientSocketLAN.OnSendMsg("CHS:OUT:DCVOF 0.0", _1652BSocketID);
                    Thread.Sleep(300);
                    time_out = 400;
                    for (i = 0; i < FreqCount; i++)
                    {
                        centerFreq = DC_FreqSlot[i];
                        //1652B输出频率
                        cmdStr = string.Format("DDFS:SIN:FREQ {0}", centerFreq.ToString("0.000"));
                        MyClientSocketLAN.OnSendMsg(cmdStr, _1652BSocketID);//信号频率
                        Thread.Sleep(200);
                        //调整示波器
                        xScale = (1.0 / centerFreq) / 2.5;
                        MyClientSocketLAN.SetMXR604ATime(xScale, _mxr604aSocketID);//调整时基
                        //MyClientSocketLAN.GetMXR604ATime(out xScale, _mxr604aSocketID);
                        Thread.Sleep(500);
                        //发命令控制1652B输出幅度
                        for (j = 0; j < VppCount; j++)
                        {
                            if (i == 0 && j == 0)
                            {
                                MyClientSocketLAN.OnSendMsg("CHS:OUT:POW 0.05", _1652BSocketID);//发指令
                                Thread.Sleep(500);
                            }
                            cmdStr = string.Format("AWGC:CHA {0}", who.ToString());
                            MyClientSocketLAN.OnSendMsg(cmdStr, _1652BSocketID);//通道切换
                            Thread.Sleep(200);
                            cmdStr = string.Format("CHS:OUT:POW {0}", AC_VppSlot[j].ToString("0.000000"));
                            MyClientSocketLAN.OnSendMsg(cmdStr, _1652BSocketID);//发指令
                            if(j == 0)    yScale = AC_VppSlot[j] / 4.0;
                            else          yScale = AC_VppSlot[j] / 6.0;
                            MyClientSocketLAN.SetMXR604AYScale(yScale, 1, _mxr604aSocketID);//调整垂直分辨率
                            Thread.Sleep(1000);//等待回读
                            MyClientSocketLAN.GetMXR604AVpp(out d, 1, _mxr604aSocketID);
                            if ( i== 0 && j == 0)
                            {
                                Thread.Sleep(2000);//等待回读
                                MyClientSocketLAN.GetMXR604AVpp(out d, 1, _mxr604aSocketID);
                                Thread.Sleep(2000);//等待回读
                                MyClientSocketLAN.GetMXR604AVpp(out d, 1, _mxr604aSocketID);
                            }
                            AC_ReadData[i, j] = d;
                        }
                        _worker.ReportProgress(100 * (i + FreqCount) / Count);//显示校准进度      
                    }

                    ParseData.ParseData.DeleteFile(saveFileName);
                    using (FileStream fw = new FileStream(saveFileName, FileMode.Create))
                    {
                        StreamWriter sw = new StreamWriter(fw);//写.csv或.txt文件
                        cmdStr = string.Format("CoupleType,Freq(MHz),SetValue(Vpp),CH{0},MaxValue(Vpp),MinValue(Vpp)\r\n", (who + 1).ToString());
                        sw.Write(cmdStr);
                        //DC
                        sw.Write("DC\r\n");
                        for (i = 0; i < FreqCount; i++)
                        {
                            for (j = 0; j < VppCount; j++)
                            {
                                GetMaxMin_Vpp(TywCommon.COUPLETYPE_DC, DC_VppSlot[j], out maxValue, out minValue);
                                cmdStr = string.Format(" ,{0}MHz,{1},{2},{3},{4}\r\n", (    DC_FreqSlot[i] / 1000000).ToString("0.0"),
                                                                                            DC_VppSlot[j].ToString("0.000"),
                                                                                            DC_ReadData[i, j].ToString("0.000"),
                                                                                            maxValue.ToString("0.000"),
                                                                                            minValue.ToString("0.000")); ;
                                sw.Write(cmdStr);
                            }
                        }
                        //AC
                        sw.Write("AC\r\n");
                        for (i = 0; i < FreqCount; i++)
                        {
                            for (j = 0; j < VppCount; j++)
                            {
                                GetMaxMin_Vpp(TywCommon.COUPLETYPE_AC, AC_VppSlot[j], out maxValue, out minValue);
                                cmdStr = string.Format(" ,{0}MHz,{1},{2},{3},{4}\r\n", (      AC_FreqSlot[i] / 1000000).ToString("0.0"),
                                                                                              AC_VppSlot[j].ToString("0.000"),
                                                                                              AC_ReadData[i, j].ToString("0.000"),
                                                                                              maxValue.ToString("0.000"),
                                                                                              minValue.ToString("0.000"));
                                sw.Write(cmdStr);
                            }
                            _worker.ReportProgress(100 * (i + FreqCount) / Count);//显示校准进度  
                        }
                        sw.Close();//不要遗漏，否则会导致最后几个数据没有写入文件               
                        fw.Close();

                    }

                }

            }
            return;
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void backgroundWorkerPowerCali_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            //Progress.Instance().ProgressValue = e.ProgressPercentage;
            SetPercentValue((double)e.ProgressPercentage);//界面显示校准进度
            return;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void backgroundWorkerPowerCali_Completed(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled == true)
            {
                SetShowMsg("Cancel!", 0);//取消!
            }
            else if (e.Error != null)
            {
                SetShowMsg("Error:" + e.Error.Message, 0);//错误:
            }
            else
            {
                SetShowMsg("");
            }
            // Close the AlertForm
            SetPercentValue(0);
            return;
        }

        /// 设置执行百分比
        /// </summary>
        /// <param name="percentV"></param>
        public void SetPercentValue(double percentV)
        {
            this.Dispatcher.BeginInvoke(new Action(() => {
                TXT_percent.Value = percentV;
            }));
            return;
        }

        /// <summary>
        /// 设置显示提示
        /// </summary>
        /// <param name="showMsg"></param>
        public void SetShowMsg(string showMsg, int _color = 0)
        {
            this.Dispatcher.BeginInvoke(new Action(() =>
            {
                labelResult.Content = showMsg;
                if (_color == 0)
                {
                    labelResult.Foreground = new SolidColorBrush(Colors.Red);
                }
                else if (_color == 1)
                {
                    labelResult.Foreground = new SolidColorBrush(Colors.Green);
                }
            }));
            return;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btn_Start_Click(object sender, RoutedEventArgs e)
        {
            //携带信息的事件
            var send = new object[4];
            send[0] = 10.5;
            send[1] = 'A';
            send[2] = 32;
            send[3] = "test string";

           
          
            CreateBackgroundWorkerPowerCali();
            backgroundWorkerPowerCali.RunWorkerAsync(send);
        

            return;
        }

        public event PropertyChangedEventHandler PropertyChanged;


        protected virtual void OnPropertyChanged(string propertyName)
        {
            var handler = PropertyChanged;
            if (handler != null) handler(this, new PropertyChangedEventArgs(propertyName));
        }

        
    }
}
