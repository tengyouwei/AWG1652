using System;
using System.Text;
using System.Net;
using System.Net.Sockets;
using System.Windows;
using System.Threading;
using AWGTestCerter.Common;

namespace AWGTestCerter.Remote
{
    /// <summary>
    /// 客户端Socket，用于校准
    /// </summary>
    public static partial class MyClientSocketLAN
    {
        public static int MSG_LENGTH = 32768;
        public static int MaxClientCount = 10;

        public static byte[] szBuffer = new byte[1024];//发送缓冲区
        public static byte[] rzBuffer = new byte[MSG_LENGTH];//接收缓冲区

        public static int m_SerPort = 4000;//服务器远程端口

        public static int  KeySightE4440X_Port = 5025;//是德科技E444x系列频谱仪
        public static int  TekMDO410X_Port = 4000;//泰克DPO7354/MDO4104系列示波器
        public static int  AV4051_Port = 5000;//中电仪器4051系列频谱仪

        public static int InstrumentType = TywCommon.SPECT_TYPE_FSW50;

        public static int WAIT_RESET = 2000;

        public static int WAIT_AUTOSET = 8000;
        public static int WAIT_WRITE = 300;
        public static int WAIT_WRITE_S = 50;

        public static int WAIT_READ = 100;
        public static int WAIT_MAXHOLD = 2000;
        public static int WAIT_CLEAR = 3000;

        public static Socket []_socketClient = new Socket[MaxClientCount];//可同时连接10个设备
        public static bool [] connectSucceed = new bool[MaxClientCount];//是否有连接标志
        public static int avtiveClientCount = 0;

               
        /// <summary>
        /// 向服务器端发送连接请求
        /// </summary>
        /// <param name="serPort">端口</param>
        /// <param name="IPAdressStr"></param>
        /// <returns></returns>
        public  static bool OnConnectSSocket(string IPAdressStr,string Port,int socketIndex)
        {
            //判断最大连接数是否超出
            if(socketIndex< 0 || socketIndex> MaxClientCount-1)
            {
                return false;
            }
            //1、创建Socket对象
            Socket socket = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
            _socketClient[socketIndex] = socket;
            //2、连接服务器,绑定IP 与 端口        
            try
            {
                IPEndPoint iPEndPoint = new IPEndPoint(IPAddress.Parse(IPAdressStr), int.Parse(Port));
                socket.Connect(iPEndPoint);
            }
            catch (Exception)
            {
                //MessageBox.Show("连接请求失败!", "提示");
                connectSucceed[socketIndex] = false;
                return false;
            }
            //3、启动接收消息线程
            //ThreadPool.QueueUserWorkItem(new WaitCallback(ReceiveServerMsg), socket);
            connectSucceed[socketIndex] = true;         
            return true;
        }

         /// <summary>
         /// 不断接收客户端信息子线程方法
         /// </summary>
         /// <param name="obj">参数Socke对象</param>
        public static bool ReceiveServerMsg(out string strRecv,out int len, int socketIndex)
        {
            len = 0;
            var proxSocket = _socketClient[socketIndex] as Socket;
            strRecv = "";
            if (proxSocket==null)
            {
                return false;
            }
            Thread.Sleep(WAIT_READ);//读前延时
            try
            {
                len = proxSocket.Receive(rzBuffer, 0, rzBuffer.Length, SocketFlags.None);
            }
            catch
            {
                //7、关闭Socket
                //异常退出
                try
                {
                    ServerExit(string.Format("服务端：{0}非正常退出", proxSocket.RemoteEndPoint.ToString()), proxSocket, socketIndex);
                }
                catch (Exception)
                {
                    //退出连接异常
                }
                return false;//让方法结束，终结当前客户端数据的异步线程，方法退出，即线程结束
            }
            if (len <= 0)//判断接收的字节数
            {
                //7、关闭Socket
                //小于0表示正常退出
                try
                {
                    ServerExit(string.Format("服务端：{0}正常退出", proxSocket.RemoteEndPoint.ToString()), proxSocket, socketIndex);
                }
                catch (Exception)
                {
                    //退出连接异常
                }
                return false;//让方法结束，终结当前客户端数据的异步线程，方法退出，即线程结束
            }
            //正常接收服务器端消息
            strRecv = Encoding.Default.GetString(rzBuffer, 0, len);
            //滤除转义字符 \r \n \t等
            strRecv = strRecv.Replace("\r", "");
            strRecv = strRecv.Replace("\n", "");
            strRecv = strRecv.Replace("\t", "");
            return true;
        }


        /// <summary>
        /// 关闭远程连接
        /// </summary>
        public static void ServerExit(string msg, Socket proxSocket, int socketIndex)
        {
            try
            {
                if (proxSocket.Connected)//如果是连接状态
                {
                    proxSocket.Shutdown(SocketShutdown.Both);//关闭连接
                    proxSocket.Close(100);//100秒超时间
                }
            }
            catch (Exception e)
            {
                //异常抛掉
                MessageBox.Show(e.ToString());
            }
            connectSucceed[socketIndex] = false;
            return;
        }


        /// <summary>
        /// 发送命令
        /// </summary>
        /// <param name="sendStr"></param>
        public static bool OnSendMsg(string sendStr, int socketIndex)
        {
            if(connectSucceed[socketIndex] ==false || sendStr==null || sendStr==string.Empty)
            {
                return false;
            }
            byte[] data = Encoding.Default.GetBytes(sendStr + "\r\n");
            //发送消息
            if (_socketClient[socketIndex] !=null)
            {
                _socketClient[socketIndex].Send(data, 0, data.Length, SocketFlags.None); 
                Thread.Sleep(WAIT_WRITE_S);
            }         
            return true;
        }


        //***************************************** 频谱仪 or 信号分析仪仪器 ******************************
        //***************************************** 频谱仪 or 信号分析仪仪器 ******************************
        //***************************************** 频谱仪 or 信号分析仪仪器 ******************************
        //***************************************** 频谱仪 or 信号分析仪仪器 ******************************
        /// <summary>
        /// 仪器复位
        /// </summary>
        /// <returns></returns>
         public static bool Reset(int socketIndex)
         {
	        if(!OnSendMsg ("*RST", socketIndex))
            {
                //MessageBox.Show("*RST命令发送失败!");
                return false;
            }
            else
            {
                Thread.Sleep(WAIT_RESET);
            }
            if (InstrumentType == TywCommon.SPECT_TYPE_FSMR50)
            {
                if (!OnSendMsg("INST:SEL SAN", socketIndex)) //FSMR50切换到频谱测量模式
                {
                    //MessageBox.Show("FSMR50切换频谱模式失败!");
                    return false;
                }
                else
                {
                    Thread.Sleep(WAIT_RESET);
                }                 
            }
		    return true;
         }

        /// <summary>
        /// 初始化
        /// </summary>
        /// <param name="rl">参考电平，dBm</param>
        /// <param name="sp">span，Hz，分辨率带宽</param>
        /// <param name="bExtRef">内外参考</param>
        /// <param name="coupleType">耦合方式</param>
        /// <returns></returns>
        public static bool Init(double r1, double sp, bool bExtRef, string coupleType,int socketIndex)
        {
            if (InstrumentType == TywCommon.SPECT_TYPE_E444X)
            {
                //E4440 N9030
                return InitE4440(r1, sp, bExtRef, coupleType, socketIndex);
            }
            else if (InstrumentType == TywCommon.SPECT_TYPE_FSW50)
            {
                //FSW50
                return InitFSW(r1, sp, bExtRef, coupleType, socketIndex);
            }
            else if (InstrumentType == TywCommon.SPECT_TYPE_AV4051X)
            {
                //4051
                return InitE4440(r1, sp, bExtRef, coupleType, socketIndex);
            }
            else if (InstrumentType == TywCommon.SPECT_TYPE_FSMR50)
            {
                //FSMR50测量接收机
                return InitFSMR50(r1, sp, bExtRef, coupleType, socketIndex);
            }
            return false;
        }

         /// <summary>
         /// E4440初始化
         /// </summary>
         /// <param name="r1"></param>
         /// <param name="sp"></param>
         /// <param name="bExtRef"></param>
         /// <param name="coupleType"></param>
         /// <returns></returns>
        public static bool InitE4440(double r1, double sp, bool bExtRef,string coupleType,int socketIndex)
        {
            string cmdStr="";
            //内外参考
            if (bExtRef)
	        {
                //外参考
		        if (!OnSendMsg ("ROSC:SOUR EXT", socketIndex))        return false;
			    if (!OnSendMsg ("ROSC:EXT:FREQ 10MHZ", socketIndex))  return false; 
	        }
            Thread.Sleep(WAIT_WRITE);
            //设置参考电平,小数点后6位
            cmdStr = string.Format("DISP:WIND:TRAC:Y:RLEV {0}dBm",r1.ToString("0.000000")); 
            if (!OnSendMsg (cmdStr, socketIndex))       return false;  Thread.Sleep(WAIT_WRITE);
            //设置SPAN   
            cmdStr = string.Format("FREQ:SPAN {0}Hz",sp.ToString("0.000000"));  
            if (!OnSendMsg (cmdStr, socketIndex))       return false;  Thread.Sleep(WAIT_WRITE);

            cmdStr=":CALC:MARK:MODE POS";
            if (!OnSendMsg (cmdStr, socketIndex))       return false;  Thread.Sleep(WAIT_WRITE);
            //耦合方式
            if (coupleType.ToUpper() != "AC" && coupleType.ToUpper()!="DC")
            {
                return  false;
            }
            cmdStr=string.Format(":INP:COUP {0}",coupleType.ToUpper());  
            if (!OnSendMsg (cmdStr, socketIndex))       return false;  
            return true;
        }

        /// <summary>
        /// 初始化
        /// </summary>
        /// <param name="r1"></param>
        /// <param name="sp"></param>
        /// <param name="bExtRef"></param>
        /// <param name="coupleType"></param>
        /// <returns></returns>
         public static bool InitFSW(double r1, double rBW, bool bExtRef,string coupleType,int socketIndex)
        {
             string cmdStr="";
             if (bExtRef)
	         {
                //外参考
		        if (!OnSendMsg ("ROSC:SOUR EXT", socketIndex))        return false;
			    if (!OnSendMsg ("ROSC:EXT:FREQ 10MHZ", socketIndex))  return false; 
	         }
             Thread.Sleep(WAIT_WRITE);
             //设置参考电平,小数点后6位
             cmdStr = string.Format("DISP:WIND:TRAC:Y:RLEV {0}dBm",r1.ToString("0.000000")); 
             if (!OnSendMsg (cmdStr, socketIndex))       return false;  Thread.Sleep(WAIT_WRITE);
             //设置分辨率带宽
             cmdStr = string.Format(":BAND {0}Hz",rBW.ToString("0.000000")); 
             if (!OnSendMsg (cmdStr, socketIndex))       return false;  Thread.Sleep(WAIT_WRITE);
              //耦合方式
             if (coupleType.ToUpper() != "AC" && coupleType.ToUpper() != "DC")
             {
                 return false;
             }
             cmdStr=string.Format(":INP:COUP {0}",coupleType.ToUpper());  
             if (!OnSendMsg (cmdStr, socketIndex))       return false;  
             return true;
         }

         /// <summary>
         /// 初始化
         /// </summary>
         /// <param name="r1"></param>
         /// <param name="sp"></param>
         /// <param name="bExtRef"></param>
         /// <param name="coupleType"></param>
         /// <returns></returns>
         public static bool InitFSMR50(double r1, double rBW, bool bExtRef, string coupleType , int socketIndex)
         {
             string cmdStr = "";
             if (bExtRef)
             {
                 //外参考
                 if (!OnSendMsg("ROSC:SOUR EXT", socketIndex)) return false;
                 if (!OnSendMsg("ROSC:EXT:FREQ 10MHZ", socketIndex)) return false;
             }
             Thread.Sleep(WAIT_WRITE);
             //设置参考电平,小数点后6位
             cmdStr = string.Format("DISP:WIND:TRAC:Y:RLEV {0}dBm", r1.ToString("0.000000"));
             if (!OnSendMsg(cmdStr, socketIndex)) return false; Thread.Sleep(WAIT_WRITE);
             //设置分辨率带宽
             cmdStr = string.Format(":BAND {0}Hz", rBW.ToString("0.000000"));
             if (!OnSendMsg(cmdStr, socketIndex)) return false; Thread.Sleep(WAIT_WRITE);
             //耦合方式
             if (coupleType.ToUpper() != "AC" && coupleType.ToUpper() != "DC")
             {
                 return false;
             }
             cmdStr = string.Format(":INP:COUP {0}", coupleType.ToUpper());
             if (!OnSendMsg(cmdStr, socketIndex)) return false;
             return true;
         }

         /// <summary>
         /// 设置显示单位时dBm还是V
         /// </summary>
         /// <param name="_uintString"></param>
         /// <returns></returns>
         public static bool SetUnit(string _uintString,int socketIndex)
         {
             string CommandStr = _uintString.ToUpper();//大写
             if (InstrumentType == TywCommon.SPECT_TYPE_E444X)
             {
                 //E4440
                 return SetE4440Unit(CommandStr, socketIndex);
             }
             else if (InstrumentType == TywCommon.SPECT_TYPE_FSW50)
             {
                 //FSW50
                 return SetFSWUnit(CommandStr, socketIndex);
             }
             else if (InstrumentType == TywCommon.SPECT_TYPE_AV4051X)
             {
                 //4051
                 return SetE4440Unit(CommandStr, socketIndex);
             }
             else if (InstrumentType == TywCommon.SPECT_TYPE_FSMR50)
             {
                 //FSMR50
                 return SetFSWUnit(CommandStr, socketIndex);
             }
             return false;
         }

         /// <summary>
         /// 设置E4440单位
         /// </summary>
         /// <param name="_uintString"></param>
         /// <returns></returns>
         public static bool SetE4440Unit(string _uintString, int socketIndex)
         {
             string cmdStr="";
             if(_uintString!="V" && _uintString!="DBM")
             {
                 return false;
             }
              cmdStr=string.Format(":UNIT:POW {0}",_uintString);
              if (!OnSendMsg(cmdStr, socketIndex)) return false; Thread.Sleep(WAIT_WRITE);
              return true;
         }
        /// <summary>
        /// 设置FSW单位
        /// </summary>
        /// <param name="_uintString"></param>
        /// <returns></returns>
         public static bool SetFSWUnit(string _uintString,int socketIndex)
         {
             string cmdStr = "";
             if (_uintString != "V" && _uintString != "DBM")
             {
                 return false;
             }
             cmdStr = string.Format(":UNIT:POW {0}", _uintString);
             if (!OnSendMsg(cmdStr, socketIndex)) return false; Thread.Sleep(WAIT_WRITE);
             return true;
         }

        /// <summary>
        /// 设置耦合方式
        /// </summary>
        /// <param name="bDC"></param>
        /// <returns></returns>
         public static bool SetACDC (string bDC,int socketIndex)
         {
             string CommandStr = bDC.ToUpper();//大写          
             if (InstrumentType == TywCommon.SPECT_TYPE_E444X)
             {
                 //E4440 N9030
                 return SetE4440ACDC(CommandStr, socketIndex);
             }
             else if (InstrumentType == TywCommon.SPECT_TYPE_FSW50)
             {
                 //FSW
                 return SetFSWACDC(CommandStr, socketIndex);
             }
             else if (InstrumentType == TywCommon.SPECT_TYPE_AV4051X)
             {
                 //4051
                 return SetE4440ACDC(CommandStr, socketIndex);
             }
             else if (InstrumentType == TywCommon.SPECT_TYPE_FSMR50)
             {
                 //FSMR50
                 return SetFSWACDC(CommandStr, socketIndex);
             }
             return false;
         }
        /// <summary>
        /// 设置E4440耦合方式
        /// </summary>
        /// <param name="bDC"></param>
        /// <returns></returns>
         public static bool SetE4440ACDC(string bDC,int socketIndex)
         {
             string cmdStr = "";
             if (bDC != "AC" && bDC != "DC")
             {
                 return false;
             }
             cmdStr = string.Format(":INP:COUP {0}", bDC);
             if (!OnSendMsg(cmdStr, socketIndex)) 
                 return false; 
             Thread.Sleep(WAIT_WRITE);
             return true;
         }
        /// <summary>
        /// FSW设置耦合方式
        /// </summary>
        /// <param name="bDC"></param>
        /// <returns></returns>
         public static bool SetFSWACDC (string bDC,int socketIndex)
         {
             string cmdStr = "";
             if (bDC != "AC" && bDC != "DC")
             {
                 return false;
             }
             cmdStr = string.Format(":INP:COUP {0}", bDC);
             if (!OnSendMsg(cmdStr, socketIndex))
                 return false;
             Thread.Sleep(WAIT_WRITE);
             return true;
         }

        /// <summary>
        /// 设置中心频率
        /// </summary>
        /// <param name="center"></param>
        /// <returns></returns>
         public static bool SetCenter(double center,int socketIndex)
         {
             if (InstrumentType == TywCommon.SPECT_TYPE_E444X)
             {
                 //E4440 N9030
                 return SetE4440Center(center, socketIndex);
             }
             else if (InstrumentType == TywCommon.SPECT_TYPE_FSW50)
             {
                 //FSW
                 return SetFSWCenter(center, socketIndex);
             }
             else if (InstrumentType == TywCommon.SPECT_TYPE_AV4051X)
             {
                 //4051
                 return SetE4440Center(center, socketIndex);
             }
             else if (InstrumentType == TywCommon.SPECT_TYPE_FSMR50)
             {
                 //FSMR50
                 return SetFSWCenter(center, socketIndex);
             }
             return false;
         }
         /// <summary>
         /// E4440中心频率设置
         /// </summary>
         /// <param name="freqStr">中心频率字符串</param>
         /// <returns></returns>
         public static bool SetE4440Center(double freqHz, int socketIndex)
         {
             double d = 0;
             int len = 0;
             string cmdStr = "";
             cmdStr = string.Format("FREQ:CENT {0}Hz", freqHz.ToString("0.000000"));
             //发中心频率设置
             if (!OnSendMsg(cmdStr, socketIndex))
                 return false;
             Thread.Sleep(WAIT_WRITE);
             //查询中心频率设置
             if (!OnSendMsg("FREQ:CENT?", socketIndex))
                 return false;
             Thread.Sleep(WAIT_WRITE);
             //读中心频率
             if(ReceiveServerMsg(out cmdStr,out len, socketIndex))
             {
                 d = ConvertFunc.ConvertStringToFrequency(cmdStr);
                 //误差小于500Hz认为中心频率设置有效
                 if(Math.Abs(freqHz-d)<500.00)
                 {
                     return true;
                 }else
                 {
                     return false;
                 }
             }
             else
             {
                 return false;
             }            
         }

         /// <summary>
         /// FSW中心频率设置
         /// </summary>
         /// <param name="freqStr">中心频率字符串</param>
         /// <returns></returns>
         public static bool SetFSWCenter(double freqHz, int socketIndex)
         {
             double d = 0;
             int len = 0;
             string cmdStr = "";
             cmdStr = string.Format("FREQ:CENT {0}Hz", freqHz.ToString("0.000000"));
             //发中心频率设置
             if (!OnSendMsg(cmdStr, socketIndex))
                 return false;
             Thread.Sleep(WAIT_WRITE);
             //查询中心频率设置
             if (!OnSendMsg("FREQ:CENT?", socketIndex))
                 return false;
             Thread.Sleep(WAIT_WRITE);
             //读中心频率
             if (ReceiveServerMsg(out cmdStr, out len, socketIndex))
             {
                 d = ConvertFunc.ConvertStringToFrequency(cmdStr);
                 //误差小于500Hz认为中心频率设置有效
                 if (Math.Abs(freqHz - d) < 500.00)
                 {
                     return true;
                 }
                 else
                 {
                     return false;
                 }
             }
             else
             {
                 return false;
             }
         }

        /// <summary>
        /// 设置接收机Span
        /// </summary>
        /// <param name="span"></param>
        /// <returns></returns>
         public static bool SetSpan(double span,int socketIndex)
         {
             if (InstrumentType == TywCommon.SPECT_TYPE_E444X)
             {
                 //E4440 N9030
                 return Set4440Span(span, socketIndex);
             }
             else if (InstrumentType == TywCommon.SPECT_TYPE_FSW50)
             {
                 //FSW
                 return SetFSW50Span(span, socketIndex);
             }
             else if (InstrumentType == TywCommon.SPECT_TYPE_AV4051X)
             {
                 //4051
                 return Set4440Span(span, socketIndex);
             }
             else if (InstrumentType == TywCommon.SPECT_TYPE_FSMR50)
             {
                 //FMSR
                 return SetFSW50Span(span, socketIndex);
             }
             return false;
         }
         /// <summary>
         /// 设置4440 span
         /// </summary>
         /// <param name="span"></param>
         /// <returns></returns>
         public static bool Set4440Span(double span,int socketIndex)
         {
             string cmdStr = "";
             cmdStr = string.Format("FREQ:SPAN {0}Hz", span.ToString());
             if (!OnSendMsg(cmdStr, socketIndex))
                 return false;
             Thread.Sleep(WAIT_WRITE);
             return true;
         }

        /// <summary>
        /// 设置FSW50 span
        /// </summary>
        /// <param name="span"></param>
        /// <returns></returns>
        public static bool SetFSW50Span (double span,int socketIndex)
        {
            string cmdStr = "";
            cmdStr = string.Format("FREQ:SPAN {0}Hz", span.ToString());
            if (!OnSendMsg(cmdStr, socketIndex))
                return false;
            Thread.Sleep(WAIT_WRITE);
            return true;
        }

         /// <summary>
         /// 读Peak值
         /// </summary>
         /// <param name="d"></param>
         /// <returns></returns>
         public static bool ReadPeak(out double d,int socketIndex)
         {
             d = 0;
             if (InstrumentType == TywCommon.SPECT_TYPE_E444X)
             {
                 //E4440 N9030
                 return Read4440Peak(out d, socketIndex);
             }
             else if (InstrumentType == TywCommon.SPECT_TYPE_FSW50)
             {
                 //FSW
                 return ReadFSW50Peak(out d, socketIndex);
             }
             else if (InstrumentType == TywCommon.SPECT_TYPE_AV4051X)
             {
                 //4051
                 return Read4440Peak(out d, socketIndex);
             }
             else if (InstrumentType == TywCommon.SPECT_TYPE_FSMR50)
             {
                 //FSMR50
                 return ReadFSW50Peak(out d, socketIndex);
             }
             return false;
         }

        /// <summary>
        /// E4440读PEAK值
        /// </summary>
        /// <param name="d"></param>
        /// <returns></returns>
         public static bool Read4440Peak(out double  d,int socketIndex)
         {
             int len = 0;
             string cmdStr = "";
             d = 0;       
             if (!OnSendMsg ("CALC:MARK1:MAX", socketIndex))   return false;   Thread.Sleep( WAIT_WRITE);
             if (!OnSendMsg("CALC:MARK1:Y?", socketIndex))     return false;   Thread.Sleep(WAIT_READ);
              //读PEAK值
             if (ReceiveServerMsg(out cmdStr, out len, socketIndex))
             {
                d= ConvertFunc.GetValueFromString(cmdStr);
                return true;         
             }
             return false;
         }



        /// <summary>
        /// FSW读Peak值
        /// </summary>
        /// <param name="d"></param>
        /// <returns></returns>
         public static bool ReadFSW50Peak(out double d,int socketIndex)
         {
             int len = 0;
             string cmdStr = "";
             d = 0;
             if (!OnSendMsg("CALC:MARK1:MAX", socketIndex)) return false; 
                 Thread.Sleep(50);
             if (!OnSendMsg("CALC:MARK1:Y?", socketIndex)) return false; 
                 Thread.Sleep(50);
             //读PEAK值
             if (ReceiveServerMsg(out cmdStr, out len,socketIndex))
             {
                 d = ConvertFunc.GetValueFromString(cmdStr);
                 return true;
             }
             return false;
         }


        /// <summary>
        /// 读
        /// </summary>
        /// <param name="pow"></param>
        /// <param name="refMax"></param>
        /// <param name="refMix"></param>
        /// <returns></returns>
       public static bool GetValidPeakFrom(out double pow, double refMax, double refMix,int  socketIndex)
       {
           pow = 0;
           if (InstrumentType == TywCommon.SPECT_TYPE_E444X)
           {
               //E4440 N9030
               return GetValidPeakFrom4440(out pow, refMax, refMix, socketIndex);
           }
           else if (InstrumentType == TywCommon.SPECT_TYPE_FSW50)
           {
               //FSW
               return GetValidPeakFromFSW50(out pow, refMax, refMix, socketIndex);
           }
           else if (InstrumentType == TywCommon.SPECT_TYPE_AV4051X)
           {
               //4051
               return GetValidPeakFrom4440(out pow, refMax, refMix, socketIndex);
           }
           else if (InstrumentType == TywCommon.SPECT_TYPE_FSMR50)
           {
               //FSMR50
               return GetValidPeakFromFSW50(out pow, refMax, refMix, socketIndex);
           }
           return false;
       }

        /// <summary>
        /// E4440读
        /// </summary>
        /// <param name="pow"></param>
        /// <param name="refMax"></param>
        /// <param name="refMix"></param>
        /// <returns></returns>
        public static bool GetValidPeakFrom4440 (out double pow, double refMax, double refMix,int socketIndex)
        {
            int i = 0;
            bool flag=false;
            pow=0;
            while(true)
            {
                flag=Read4440Peak(out pow, socketIndex);
                i++;
                if(i>3)  //最多读取4次
                {
                    return  false;
                }
                if(!flag || pow>refMax || pow<refMix)
                {
                    Thread.Sleep(1200);
                    continue;
                }
                else
                {
                    return true;
                }         
            }
        }
        /// <summary>
        /// FSW50读
        /// </summary>
        /// <param name="pow"></param>
        /// <param name="refMax"></param>
        /// <param name="refMix"></param>
        /// <returns></returns>
        public static bool GetValidPeakFromFSW50(out double pow, double refMax, double refMix,int socketIndex)
        {
            int i = 0;
            bool flag = false;
            pow = 0;
            while (true)
            {
                flag = ReadFSW50Peak(out pow, socketIndex);
                i++;
                if (i > 3)  //最多读取4次
                {
                    return false;
                }
                if (!flag || pow > refMax || pow < refMix)
                {
                    Thread.Sleep(1200);
                    continue;
                }
                else
                {
                    return true;
                }
            }
        }

        /// <summary>
        /// 设置起始频率
        /// </summary>
        /// <param name="start"></param>
        /// <returns></returns>
        public static bool SetStartFreq(double start,int socketIndex)
        {
            if (InstrumentType == TywCommon.SPECT_TYPE_E444X)
            {
                //E4440
                return Set4440StartFreq(start, socketIndex);
            }
            else if (InstrumentType == TywCommon.SPECT_TYPE_AV4051X)
            {
                //4051
                return Set4440StartFreq(start, socketIndex);
            }
            else if (InstrumentType == TywCommon.SPECT_TYPE_FSW50)
            {
                //FSW50
                return SetFSW50StartFreq(start, socketIndex);
            }
            else if (InstrumentType == TywCommon.SPECT_TYPE_FSMR50)
            {
                //FSMR50
                return SetFSW50StartFreq(start, socketIndex);
            }
            return false;
        }

        /// <summary>
        /// E4440起始频率
        /// </summary>
        /// <param name="start"></param>
        /// <returns></returns>
         public static bool Set4440StartFreq(double start,int socketIndex)
        {
            string cmdStr = "";
            cmdStr = string.Format("FREQ:START {0}Hz", start.ToString("0.000000"));
            if (!OnSendMsg(cmdStr, socketIndex)) 
                return false; 
             Thread.Sleep(WAIT_WRITE);
             return true;
        }

         /// <summary>
         /// FSW50起始频率
         /// </summary>
         /// <param name="start"></param>
         /// <returns></returns>
         public static bool SetFSW50StartFreq(double start, int socketIndex)
         {
             string cmdStr = "";
             cmdStr = string.Format("FREQ:START {0}Hz", start.ToString("0.000000"));
             if (!OnSendMsg(cmdStr, socketIndex))
                 return false;
             Thread.Sleep(WAIT_WRITE);
             return true;
         }

        /// <summary>
        /// 设置终止频率
        /// </summary>
        /// <param name="stop"></param>
        /// <returns></returns>
        public static bool SetStopFreq(double stop,int socketIndex)
        {
            if (InstrumentType == TywCommon.SPECT_TYPE_E444X)
            {
                //E4440 N9030
                return Set4440StopFreq(stop, socketIndex);
            }
            else if (InstrumentType == TywCommon.SPECT_TYPE_AV4051X)
            {
                //4051
                return Set4440StopFreq(stop, socketIndex);
            }
            else if (InstrumentType == TywCommon.SPECT_TYPE_FSW50)
            {
                //FSW50
                return SetFSW50StopFreq(stop, socketIndex);
            }

            else if (InstrumentType == TywCommon.SPECT_TYPE_FSMR50)
            {
                //FSMR50
                return SetFSW50StopFreq(stop, socketIndex);
            }
            return false;
        }



         /// <summary>
         /// E4440终止频率
         /// </summary>
         /// <param name="start"></param>
         /// <returns></returns>
         public static bool Set4440StopFreq(double stop,int socketIndex)
         {
             string cmdStr = "";
             cmdStr = string.Format("FREQ:STOP {0}Hz", stop.ToString("0.000000"));
             if (!OnSendMsg(cmdStr, socketIndex))
                 return false;
             Thread.Sleep(WAIT_WRITE);
             return true;
         }

         /// <summary>
         /// FSW50终止频率
         /// </summary>
         /// <param name="start"></param>
         /// <returns></returns>
         public static bool SetFSW50StopFreq(double stop,int socketIndex)
         {
             string cmdStr = "";
             cmdStr = string.Format("FREQ:STOP {0}Hz", stop.ToString("0.000000"));
             if (!OnSendMsg(cmdStr, socketIndex))
                 return false;
             Thread.Sleep(WAIT_WRITE);
             return true;
         }

         /// <summary>
         /// 设置TRACE为最大保持
         /// </summary>
         /// <returns></returns>
         public static bool SetTraceMaxHoldMode(int socketIndex)
         {
             if (InstrumentType == TywCommon.SPECT_TYPE_E444X)
             {
                 //E4440 N9030
                 return Set4440TraceMaxHoldMode(socketIndex);
             }
             else if (InstrumentType == TywCommon.SPECT_TYPE_AV4051X)
             {
                 //4051
                 return Set4440TraceMaxHoldMode(socketIndex);
             }
             else if (InstrumentType == TywCommon.SPECT_TYPE_FSW50)
             {
                 //FSW50
                 return SetFSW50TraceMaxHoldMode(socketIndex);
             }
             else if (InstrumentType == TywCommon.SPECT_TYPE_FSMR50)
             {
                 //FSMR50
                 return SetFSW50TraceMaxHoldMode(socketIndex);
             }
             return false;
         }

        /// <summary>
        /// E4440最大保持
        /// </summary>
        /// <returns></returns>
        public static bool Set4440TraceMaxHoldMode(int socketIndex)
        {
             if (!OnSendMsg("TRACE:MODE MAXH", socketIndex))
                 return false;
             Thread.Sleep(WAIT_WRITE);
             return true;
         }

        /// <summary>
        /// FSW50最大保持
        /// </summary>
        /// <returns></returns>
        public static bool SetFSW50TraceMaxHoldMode(int socketIndex)
        {
            if (!OnSendMsg("DISP:WIND:TRACE:MODE MAXH", socketIndex))
                return false;
            Thread.Sleep(WAIT_WRITE);
            return true;
        }

        /// <summary>
        /// 清除TRACE最大最小保持
        /// </summary>
        /// <returns></returns>
        public static bool SetTraceClear(int socketIndex)
        {
            if (InstrumentType == TywCommon.SPECT_TYPE_E444X)
            {
                //E4440 N9030
                return Set4440TraceClear(socketIndex);
            }
            else if (InstrumentType == TywCommon.SPECT_TYPE_AV4051X)
            {
                //4051
                return Set4440TraceClear(socketIndex);
            }
            else if (InstrumentType == TywCommon.SPECT_TYPE_FSW50)
            {
                //FSW50
                return SetFSW50TraceClear(socketIndex);
            }
            else if (InstrumentType == TywCommon.SPECT_TYPE_FSMR50)
            {
                //FSMR50
                return SetFSW50TraceClear(socketIndex);
            }
            return false;
        }

        /// <summary>
        /// E4440 清除TRACE最大最小保持
        /// </summary>
        /// <returns></returns>
        public static bool Set4440TraceClear(int socketIndex)
        {
            if (!OnSendMsg("TRACE:MODE WRIT", socketIndex))
                return false;
            Thread.Sleep(WAIT_WRITE);
            return true;
        }

        /// <summary>
        /// FSW50 清除TRACE最大最小保持
        /// </summary>
        /// <returns></returns>
        public static bool SetFSW50TraceClear(int socketIndex)
        {
            if (!OnSendMsg("DISP:WIND:TRACE:MODE WRIT", socketIndex))
                return false;
            Thread.Sleep(WAIT_WRITE);
            return true;
        }

        /// <summary>
        /// 设置接收机Amplle Range
        /// </summary>
        /// <param name="range"></param>
        /// <returns></returns>
        public static bool SetRange(double range,int  socketIndex)
        {
            if (InstrumentType == TywCommon.SPECT_TYPE_E444X)
            {
                //E4440 N9030
                return Set4440Range(range, socketIndex);
            }
            else if (InstrumentType == TywCommon.SPECT_TYPE_AV4051X)
            {
                //4051
                return Set4440Range(range, socketIndex);
            }
            else if (InstrumentType == TywCommon.SPECT_TYPE_FSW50)
            {
                //FSW
                return SetFSW50Range(range, socketIndex);
            }
            else if (InstrumentType == TywCommon.SPECT_TYPE_FSMR50)
            {
                //FSMR50
                return SetFSW50Range(range, socketIndex);
            }
            return false;
        }
        /// <summary>
        /// 设置4440接收机Amplle Range
        /// </summary>
        /// <param name="range"></param>
        /// <returns></returns>
         public static bool Set4440Range(double range,int socketIndex)
        {
            string cmdStr = "";
            cmdStr = string.Format("DISP:TRAC:Y {0}", range.ToString());
            if (!OnSendMsg(cmdStr, socketIndex))
                return false;
            Thread.Sleep(WAIT_WRITE);
            return true;
        }
        /// <summary>
         /// 设置FSW50接收机Amplle Range
        /// </summary>
        /// <param name="range"></param>
        /// <returns></returns>
        public static bool SetFSW50Range(double range,int socketIndex)
        {
            string cmdStr = "";
            cmdStr = string.Format("DISP:TRAC:Y {0}", range.ToString());
            if (!OnSendMsg(cmdStr, socketIndex))
                return false;
            Thread.Sleep(WAIT_WRITE);
            return true;
        }
      
        /// <summary>
        /// 设置慢屏点数
        /// </summary>
        /// <param name="points"></param>
        /// <returns></returns>
        public static bool SetSwePoint(int points,int socketIndex)
        {
            if (InstrumentType == TywCommon.SPECT_TYPE_E444X)
            {
                //E4440 N9030
                return Set4440SwePoint(points, socketIndex);
            }
            else if (InstrumentType == TywCommon.SPECT_TYPE_AV4051X)
            {
                //4051
                return Set4440SwePoint(points, socketIndex);
            }
            else if (InstrumentType == TywCommon.SPECT_TYPE_FSW50)
            {
                //FSW50
                return SetFSW50SwePoint(points, socketIndex);
            }
            else if (InstrumentType == TywCommon.SPECT_TYPE_FSMR50)
            {
                //FSMR50
                return SetFSW50SwePoint(points, socketIndex);
            }
            return false;
        }



        /// <summary>
        /// E4440设置满屏点数
        /// </summary>
        /// <param name="points"></param>
        /// <returns></returns>
        public static bool Set4440SwePoint(int points,int socketIndex)
        {
            string cmdStr = "";
            cmdStr = string.Format(":SWE:POIN {0}", points.ToString());
            if (!OnSendMsg(cmdStr, socketIndex))
                return false;
            Thread.Sleep(WAIT_WRITE);
            return true;
        }

        /// <summary>
        /// FSW50设置慢屏点数
        /// </summary>
        /// <param name="points"></param>
        /// <returns></returns>
        public static bool SetFSW50SwePoint(int points,int socketIndex)
        {
            string cmdStr = "";
            cmdStr = string.Format(":SWE:POIN {0}", points.ToString());
            if (!OnSendMsg(cmdStr, socketIndex))
                return false;
            Thread.Sleep(WAIT_WRITE);
            return true;
        }

        /// <summary>
        /// 读Trace迹线数据
        /// </summary>
        /// <param name="pTraceData"></param>
        /// <returns></returns>
        public static bool GetTraceData(out string pTraceData,int socketIndex)
        {
            pTraceData = "";
            if (InstrumentType == TywCommon.SPECT_TYPE_E444X)
            {
                //E4440 N9030
                return Get4440TraceData(out pTraceData, socketIndex);
            }
            else if (InstrumentType == TywCommon.SPECT_TYPE_AV4051X)
            {
                //4051
                return Get4440TraceData(out pTraceData, socketIndex);
            }
            else if (InstrumentType == TywCommon.SPECT_TYPE_FSW50)
            {
                //FSW
                return GetFSW50TraceData(out pTraceData, socketIndex);
            }
            else if (InstrumentType == TywCommon.SPECT_TYPE_FSMR50)
            {
                //FSMR50
                return GetFSMR50TraceData(out pTraceData, socketIndex);
            }
            return false;
        }

         /// <summary>
         /// E4440读Trace迹线数据
         /// </summary>
         /// <param name="pTraceData"></param>
         /// <returns></returns>
         public static bool Get4440TraceData(out string pTraceData,int socketIndex)
         {
             int len=0;
             pTraceData="";
             //step1: 准备获取TRACE1数据
             if (!OnSendMsg(":TRAC:DATA? TRACE1", socketIndex))
                return false;
            Thread.Sleep(WAIT_WRITE);
             //step2: 准备
             if (!OnSendMsg(":TRA?", socketIndex))
                return false;
            Thread.Sleep(WAIT_WRITE);
            //step3: 读
            if (!ReceiveServerMsg(out pTraceData, out len, socketIndex)) return false;
            if(len<1)
            {
                return  false;
            }
             return true;
         }

         /// <summary>
         /// FSW50 读Trace迹线数据
         /// </summary>
         /// <param name="pTraceData"></param>
         /// <returns></returns>
         public static bool GetFSW50TraceData(out string pTraceData,int socketIndex)
         {
              int len=0;
              pTraceData="";
               //step1: 准备获取TRACE1数据
              if (!OnSendMsg(":TRAC:DATA? TRACE1", socketIndex))
                return false;
              Thread.Sleep(WAIT_WRITE);
               //step3: 读
              if (!ReceiveServerMsg(out pTraceData, out len, socketIndex)) return false;
              if(len<1)
              {
                return  false;
              }
              return true;
          }

        /// <summary>
         /// FSMR50 读Trace迹线数据
        /// </summary>
        /// <param name="pTraceData"></param>
        /// <returns></returns>
         public static bool GetFSMR50TraceData(out string pTraceData,int socketIndex)
         {
             int len = 0;
             pTraceData = "";
             //step1: 准备获取TRACE1数据
             if (!OnSendMsg("TRAC? TRACE1", socketIndex))
                 return false;
             Thread.Sleep(WAIT_WRITE);
             //step3: 读
             if (!ReceiveServerMsg(out pTraceData, out len, socketIndex)) return false;
             if (len < 1)
             {
                 return false;
             }
             return true;
         }
       
        /// <summary>
        /// Mark到频率为xFreq的点
        /// </summary>
        /// <param name="xFreq"></param>
        /// <returns></returns>
        public static bool SetMarkXfreq(double xFreq,int socketIndex)
        {
            if (InstrumentType == TywCommon.SPECT_TYPE_E444X)
            {
                //E4440 N9030
                return Set4440MarkXfreq(xFreq, socketIndex);
            }
            else if (InstrumentType == TywCommon.SPECT_TYPE_FSMR50)
            {
                //FSMR50
                return SetFSW50MarkXfreq(xFreq, socketIndex);
            }
            else if (InstrumentType == TywCommon.SPECT_TYPE_FSW50)
            {
                //FSW
                return SetFSW50MarkXfreq(xFreq, socketIndex);
            }
            else if (InstrumentType == TywCommon.SPECT_TYPE_AV4051X)
            {
                //4051
                return Set4440MarkXfreq(xFreq, socketIndex);
            }
            return false;
        }


        /// <summary>
        /// E4440 Mark到频率为xFreq的点
        /// </summary>
        /// <param name="xFreq"></param>
        /// <returns></returns>
        public static bool Set4440MarkXfreq(double xFreq,int socketIndex)
        {
            string cmdStr = "";
            cmdStr = string.Format("CALC:MARK:X {0}Hz", xFreq);
            if (!OnSendMsg(cmdStr, socketIndex))
                return false;
            Thread.Sleep(WAIT_WRITE);
            return true;
        }


        /// <summary>
        /// FSW50 Mark到频率为xFreq的点
        /// </summary>
        /// <param name="xFreq"></param>
        /// <returns></returns>
        public static bool SetFSW50MarkXfreq(double xFreq,int socketIndex)
        {
            string cmdStr = "";
            cmdStr = string.Format("CALC:MARK:X {0}Hz", xFreq);
            if (!OnSendMsg(cmdStr, socketIndex))
                return false;
            Thread.Sleep(WAIT_WRITE);
            return true;
        }
         
        /// <summary>
        /// 获取Mark 读数(Y方向)
        /// </summary>
        /// <param name="powerDBm"></param>
        /// <returns></returns>
         public static bool GetMarkYdBm(out double powerDBm,int socketIndex)
        {
            powerDBm = 0;
            if (InstrumentType == TywCommon.SPECT_TYPE_E444X)
            {
                //E4440 N9030
                return Get4440MarkYdBm(out powerDBm, socketIndex);
            }
            else if (InstrumentType == TywCommon.SPECT_TYPE_FSW50)
            {
                //FSW
                return GetFSW50MarkYdBm(out powerDBm, socketIndex);
            }
            else if (InstrumentType == TywCommon.SPECT_TYPE_AV4051X)
            {
                //4051
                return Get4440MarkYdBm(out powerDBm, socketIndex);
            }
            else if (InstrumentType == TywCommon.SPECT_TYPE_FSMR50)
            {
                //FSMR50
                return GetFSW50MarkYdBm(out powerDBm, socketIndex);
            }
            return false;
        }
        /// <summary>
        /// E4440取Mark位置值
        /// </summary>
        /// <param name="powerDBm"></param>
        /// <returns></returns>
        public static bool Get4440MarkYdBm(out double powerDBm,int socketIndex)
        {
            int len=0;
            powerDBm=0;
            string cmdStr = "";
            cmdStr = string.Format("CALC:MARK:Y?");
            if (!OnSendMsg(cmdStr, socketIndex))  return false;  Thread.Sleep(WAIT_WRITE);
            if (!ReceiveServerMsg(out cmdStr, out len,socketIndex)) return false;
            if(len<1)
            {
               return  false;
            }
            powerDBm = ConvertFunc.GetValueFromString(cmdStr);
           return true;           
        }


        /// <summary>
        /// FSW50取Mark位置值
        /// </summary>
        /// <param name="powerDBm"></param>
        /// <returns></returns>
        public static bool GetFSW50MarkYdBm(out double powerDBm,int socketIndex)
        {
            int len = 0;
            powerDBm = 0;
            string cmdStr = "";
            cmdStr = string.Format("CALC:MARK:Y?");
            if (!OnSendMsg(cmdStr, socketIndex)) return false; Thread.Sleep(WAIT_WRITE);
            if (!ReceiveServerMsg(out cmdStr, out len,socketIndex)) return false;
            if (len < 1)
            {
                return false;
            }
            powerDBm = ConvertFunc.GetValueFromString(cmdStr);
            return true;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="fCarry"></param>
        /// <param name="fBB"></param>
        /// <param name="pPeakF"></param>
        /// <param name="pImageF"></param>
        /// <returns></returns>
        public static bool GetFSW50TwoFreq(double fCarry, double fBB, out double pPeakF, out double pImageF,int socketIndex)
        {
            double peakF=0;
            int len=0;
            string cmdStr="";
            pPeakF = 0;
            pImageF = 0;
            if (!OnSendMsg("CALC:MARK1:MAX", socketIndex)) return false; Thread.Sleep(WAIT_WRITE);
            if (!OnSendMsg("CALC:MARK1:X?", socketIndex))  return false; Thread.Sleep(WAIT_WRITE);
            if (!ReceiveServerMsg(out cmdStr, out len,socketIndex)) return false;
            pPeakF=ConvertFunc.GetValueFromString(cmdStr);
            if (Math.Abs(Math.Abs(peakF-fCarry) - fBB) > 100000000.0)
		    return false;
            pPeakF  = peakF;
            pImageF = 2.0 * fCarry - peakF;        
            return true;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="span"></param>
        /// <param name="f"></param>
        /// <param name="pMrkY"></param>
        /// <returns></returns>
        public static bool ReadFSW50MrkWith_threedot (double span, double f, out double pMrkY,int socketIndex)
       {
            int len=0;
            string cmdStr="";
	        double   xstep = span / 1001.0;
            pMrkY = 0;

            cmdStr=string.Format("CALC:MARK1:X {0}Hz",f.ToString("0.000000"));
            if (!OnSendMsg("CALC:MARK1:Y?", socketIndex)) return false; Thread.Sleep(WAIT_WRITE);
            if (!ReceiveServerMsg(out cmdStr, out len,socketIndex)) return false;
            pMrkY = ConvertFunc.GetValueFromString(cmdStr);
            return true;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="span"></param>
        /// <param name="fPeak"></param>
        /// <param name="fImage"></param>
        /// <param name="pPeak"></param>
        /// <param name="pImage"></param>
        /// <param name="avgCount"></param>
        /// <returns></returns>
        public static bool GetFSW50TwoAmpl(double span, double fPeak, double fImage, out double pPeak,out double pImage, int avgCount,int socketIndex)
        {         
            int i;
            double d1 = 0, d2 = 0, d;
            pPeak = 0;
            pImage = 0;
            Thread.Sleep(200);    //WAIT_WRITE_S  

            for (i = 0; i < avgCount; i++)
            {
                Thread.Sleep(100);
                if (!ReadFSW50MrkWith_threedot(span, fPeak, out d, socketIndex)) return false;
                d1 += d;
                Thread.Sleep(100);
                if (!ReadFSW50MrkWith_threedot(span, fImage, out d, socketIndex)) return false;
                d2 += d;
            }
            pPeak = d1 / (double)avgCount;
            pImage = d2 / (double)avgCount;
            return true;
        }



        /// <summary>
        /// Mark->Peak
        /// </summary>
        /// <param name="xFreq"></param>
        /// <returns></returns>
        public static bool SetFSW13MarkToPeak (int socketIndex)
        {	
             if (!OnSendMsg("CALC:MARK:X PEAK", socketIndex)) 
             return false; 
             Thread.Sleep(WAIT_WRITE);	    
	         return true;
        }


        /// <summary>
        /// Mark->Peak
        /// </summary>
        /// <param name="xFreq"></param>
        /// <returns></returns>
        public static bool SetFSMR50MarkToPeak(int socketIndex)
        {
            if (!OnSendMsg("CALC:MARK:X PEAK", socketIndex))
                return false;
            Thread.Sleep(WAIT_WRITE);
            return true;
        }




        //***************************************** MDO4104B示波器 ******************************************
        //***************************************** MDO4104B示波器 ******************************************
        //***************************************** MDO4104B示波器 ******************************************
        //***************************************** MDO4104B示波器 ******************************************
        /// <summary>
        /// 复位
        /// </summary>
        /// <returns></returns>
        public static bool  ResetMDO4104(int socketIndex)
        {
            if (!OnSendMsg("*RST", socketIndex))
            {
                return false;
            }
            else
            {
                Thread.Sleep(3000);
                return true;
            }	
        }

        /// <summary>
        /// 自动
        /// </summary>
        /// <returns></returns>
        public static bool  AutoSetMDO4104(int socketIndex)
        {
            if (!OnSendMsg("AUTOSET EXECute", socketIndex))
            {
                return false;
            }
            else
            {
                Thread.Sleep(6000);
                return true;
            }
        }

       /// <summary>
        /// 设置通道chx的阻抗位50欧姆或者1M欧姆
       /// </summary>
       /// <param name="chx">1/2/3/4</param>
       /// <param name="ImpMode"></param>
       /// <returns></returns>
        public static bool SetMDO4104InputMode(string ImpMode = "50", int Channel = 1,int socketIndex=0)
        {
            string strImp = ImpMode.ToUpper();
            string cmdStr = "";
            if (Channel < 1 || Channel > 4)
            {
                return false;
            }
            if (strImp == "1M")
            {
                cmdStr = string.Format("CH{0}:TERMINATION MEG", Channel.ToString()); 
            }
            else if(strImp=="50")
            {
                cmdStr = string.Format("CH{0}:TERMINATION FIF", Channel.ToString()); 
            }
            if (!OnSendMsg(cmdStr, socketIndex))
                return false;
            Thread.Sleep(WAIT_WRITE);	            
            return true;
        }

        /// <summary>
        /// 设置纵坐标刻度
        /// </summary>
        /// <param name="yScale"></param>
        /// <param name="Channel"></param>
        /// <returns></returns>
        public static bool SetMDO4104YScale(double yScale, int Channel = 1,int socketIndex=0)
        {
            string cmdStr = "";
            cmdStr = string.Format("CH{0}:SCALE {1}", Channel.ToString(),yScale.ToString("0.000000")); 
            if (!OnSendMsg(cmdStr, socketIndex))
                return false;
            Thread.Sleep(WAIT_WRITE);	            
            return true;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="pScale"></param>
        /// <param name="Channel"></param>
        /// <returns></returns>
        public static bool GetMDO4104YScale(out double pScale, int Channel = 1,int socketIndex=0)
        {
            int len=0;
            string cmdStr = "";
            pScale=0;
            cmdStr = string.Format("CH{0}:SCALE?", Channel.ToString()); 
            if (!OnSendMsg(cmdStr, socketIndex))
                return false;
            Thread.Sleep(WAIT_READ);	
            if (!ReceiveServerMsg(out cmdStr, out len,socketIndex)) 
            {
                return false;
            }
            if(len<1)  return false;
            pScale=ConvertFunc.GetValueFromString(cmdStr);
            return true;        
        }

        /// <summary>
        /// 设置横坐标刻度
        /// </summary>
        /// <param name="xScale"></param>
        /// <returns></returns>
        public static bool SetMDO4104XScale(double xScale,int socketIndex)
        {
            string cmdStr = "";
            cmdStr = string.Format("HORIZONTAL:SCALE {0}", xScale.ToString("0.000000"));
            if (!OnSendMsg(cmdStr, socketIndex))
                return false;
            Thread.Sleep(WAIT_WRITE);	       
	        return true;
        }

        /// <summary>
        /// 获取横坐标刻度
        /// </summary>
        /// <param name="pScale"></param>
        /// <returns></returns>
        public static bool GetMDO4104XScale(out double pScale,int socketIndex)
        {
            int len = 0;
            string cmdStr = "";
            pScale = 0;
            if (!OnSendMsg("HORIZONTAL:SCALE?", socketIndex))
                return false;
            Thread.Sleep(WAIT_READ);
            if (!ReceiveServerMsg(out cmdStr, out len,socketIndex))
            {
                return false;
            }
            if (len < 1) return false;
            pScale = ConvertFunc.GetValueFromString(cmdStr);
            return true;
        }

        /// <summary>
        /// 设置通道x水平旋钮位置
        /// </summary>
        /// <param name="ch1Pos"></param>
        /// <param name="index"></param>
        /// <returns></returns>
        public static bool SetMDO4104CHxPos(double channelPos = 0, int channel = 1,int socketIndex=0)
        {
            string cmdStr = "";
            cmdStr = string.Format("CH{0}:POS {1}" ,channel.ToString(), channelPos.ToString("0.000000"));
            if (!OnSendMsg(cmdStr, socketIndex))
		    return false;  
	        Thread.Sleep(2000);//延时2s
            return true;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="offset"></param>
        /// <returns></returns>
        public static bool SetMDO4104CH1Offset(double offset = 0,int socketIndex=0)
        {
             string cmdStr = "";
             cmdStr = string.Format("CH1:OFFS {0}" ,offset.ToString("0.000000"));
             if (!OnSendMsg(cmdStr, socketIndex))
		    return false;  
	        Thread.Sleep(2000);//延时2s
            return true;
        }

        /// <summary>
        /// 返回MDO4104第index个测量量的平均值,index<=8
        /// </summary>
        /// <param name="index"></param>
        /// <returns></returns>
        public static bool GetMDO4104MEAS(out double d, int index = 1,int socketIndex=0)
        {
            int len = 0;
            string cmdStr = "";      
            d = 0;
            if (index < 0 || index > 8)
            {
                return false;
            }
            cmdStr = string.Format("MEASU:MEAS{0}:VAL?", index.ToString());
            if (!OnSendMsg(cmdStr, socketIndex))
                return false;
            Thread.Sleep(WAIT_READ);

            if (!ReceiveServerMsg(out cmdStr, out len,socketIndex))
            {
                return false;
            }
            if (len < 1) return false;
            d = ConvertFunc.GetValueFromString(cmdStr);
            return true;
        }

        /// <summary>
        /// type为测量类型,index表示第index个测量量
        /// </summary>
        /// <param name="type"></param>
        /// <param name="index"></param>
        /// <returns></returns>
        public static bool SetMDO4104MEASTYPE(int type, int index = 1,int socketIndex=0)
        {
            //type==1:设置测量类型为FREQUENCY 频率
            //type==2:设置测量类型为PK2PK     峰峰值
            //type==3:设置测量类型为MEAN      平均
            string cmdStr = "";      
            if (type < 1 || type > 3)    return false;
            if (index < 0 || index > 8)  return false;
            switch (type)
            {
                case 1: cmdStr = string.Format(":MEASU:MEAS{0}:TYP FREQUENCY", index.ToString());break;
                case 2: cmdStr = string.Format(":MEASU:MEAS{0}:TYP PK2PK", index.ToString()); break;
                case 3: cmdStr = string.Format(":MEASU:MEAS{0}:TYP MEAN", index.ToString()); break;
                default: break;
            }
            if (!OnSendMsg(cmdStr, socketIndex))
                return false;
            Thread.Sleep(WAIT_WRITE);
            return true;
        }

        /// <summary>
        /// 获取index测量类型
        /// </summary>
        /// <param name="type"></param>
        /// <param name="index"></param>
        /// <returns></returns>
        public static bool GetMDO4104MEASTYPE(out int type, int index = 1,int socketIndex=0)
        {
            //type==1:设置测量类型为FREQUENCY 频率
            //type==2:设置测量类型为PK2PK     峰峰值
            //type==3:设置测量类型为MEAN      平均
            int len=0;
            string cmdStr = ""; 
            type=0;
            cmdStr=string.Format("MEASU:MEAS{0}:TYP?", index.ToString());
            if (!OnSendMsg(cmdStr, socketIndex))
                return false;
            Thread.Sleep(WAIT_READ);
            if (!ReceiveServerMsg(out cmdStr, out len,socketIndex))
            {
                return false;
            }
            if (len < 1) return false;
            cmdStr=cmdStr.ToUpper();
            if(cmdStr=="FREQUENCY" || cmdStr=="FREQ")
            {
                type=1;
            }
            if(cmdStr=="PK2PK")
            {
                type=2;
            }
            if(cmdStr=="MEAN")
            {
                type=3;
            }
            return true;
        }

        /// <summary>
        /// 设置第index个测量量测量源为channel个通道
        /// </summary>
        /// <param name="index"></param>
        /// <param name="channel"></param>
        /// <returns></returns>
        public static bool SetMEASxTOChannel(int index = 1, int channel = 1,int socketIndex=0)
        {
            string cmdStr = ""; 
            cmdStr=string.Format("MEASU:MEAS{0}:SOU CH{1}", index.ToString(),channel.ToString());
             if (!OnSendMsg(cmdStr, socketIndex))
                return false;
            Thread.Sleep(WAIT_WRITE);           
		    return true;
        }




        //***************************************** Keysight 34470A ******************************************
        //***************************************** Keysight 34470A ******************************************
        //***************************************** Keysight 34470A ******************************************
        //***************************************** Keysight 34470A ******************************************
        /// <summary>
        /// 复位
        /// </summary>
        /// <returns></returns>
        public static bool ReSetKeysight34470A(int socketIndex)
        {
            if (!OnSendMsg("*RST", socketIndex))
            {
                return false;
            }
            else
            {
                Thread.Sleep(2000);
                return true;
            }
        }
        /// <summary>
        /// 34470A初始状态
        /// </summary>
        /// <returns></returns>
        public static bool InitSetKeysight34470A(int socketIndex)
        {
            if (!OnSendMsg("CONF:VOLT:DC", socketIndex))
                return false;
            Thread.Sleep(600);
            if (!OnSendMsg("VOLT:ZERO:AUTO OFF", socketIndex))
                return false;
            Thread.Sleep(600);
            return true;
        }

        
        /// <summary>
        /// 
        /// </summary>
        /// <param name="d"></param>
        /// <returns></returns>
        public static bool GetOFTValue(out double d,int socketIndex)
        {
            int len = 0;
            string cmdStr = "";
            d = 0;
            //第1次读
            if (!OnSendMsg("READ?", socketIndex))
                return false;
            Thread.Sleep(1000);
            if (!ReceiveServerMsg(out cmdStr, out len,socketIndex))
            {
                return false;
            }
            if (len < 1)
                return false;
            d = ConvertFunc.GetValueFromString(cmdStr);
            return true;
        }



        //***************************************** MSO254A示波器 ******************************************
        //***************************************** MSO254A示波器 ******************************************
        //***************************************** MSO254A示波器 ******************************************
        //***************************************** MSO254A示波器 ******************************************
        /// <summary>
        /// 复位MSO254A
        /// </summary>
        /// <returns></returns>
        public static bool RstMSO254A(int socketIndex)
        {
            if (!OnSendMsg("*RST", socketIndex))
            {
                return false;
            }
            Thread.Sleep(3000);
            return true;
        }
        /// <summary>
        /// MSO254A自动1次
        /// </summary>
        /// <returns></returns>
        public static bool AutoSetMSO254A(int socketIndex)
        {
            if (!OnSendMsg("AUT", socketIndex))
            {
                return false;
            }
            else
            {
                Thread.Sleep(5000);
                return true;
            }
        }
        /// <summary>
        /// MSO254A 1通道幅度校准初始化
        /// </summary>
        /// <returns></returns>
        public static bool InitMSO254AOneChannel(int osc_channel,int socketIndex)
        {
            string cmdStr = "";          
            if (osc_channel < 1 || osc_channel > 4)
                return false;
            cmdStr = string.Format(":CHAN{0}:DISP ON", osc_channel.ToString());
            if (!OnSendMsg(cmdStr, socketIndex))
                return false;
            cmdStr = string.Format(":CHAN{0}:INP DC50", osc_channel.ToString());
            if (!OnSendMsg(cmdStr,socketIndex))
                return false;
            cmdStr = string.Format(":CHAN{0}:SCAL 0.2V", osc_channel.ToString());
            if (!OnSendMsg(cmdStr,socketIndex))
                return false;
            if (!OnSendMsg(":TIM:SCAL 0.00000005s", socketIndex))
                return false;
            cmdStr = string.Format(":TRIG:EDGE:SOUR  CHAN{0}", osc_channel.ToString());
            if (!OnSendMsg(cmdStr,socketIndex))
                return false;
            cmdStr = string.Format(":MEAS:SOUR CHAN{0}", osc_channel.ToString());
            if (!OnSendMsg(cmdStr,socketIndex))
                return false;
            if (!OnSendMsg(":MEAS:FREQ", socketIndex))
                return false;
            if (!OnSendMsg(":MEAS:Vaverage", socketIndex))
                return false;
            if (!OnSendMsg(":MEAS:VPP", socketIndex))
                return false;
            if (!OnSendMsg(":MEAS:V AVG", socketIndex))
                return false; 
            cmdStr = string.Format(":CHAN%d:OFFS 0.00V", osc_channel.ToString());
            if (!OnSendMsg(cmdStr,socketIndex))
                return false;
            return true;
        }

        /// <summary>
        /// MSO254A 通道1-2通道幅度校准初始化
        /// </summary>
        /// <returns></returns>
        public static bool InitMSO254ATwoChannel(int socketIndex)
        {
            int i = 1;
            string cmdStr = "";
            for (i = 1; i <= 2; i++)
            {
                cmdStr = string.Format(":CHAN{0}:DISP ON", socketIndex, i.ToString());
                if (!OnSendMsg(cmdStr,socketIndex))
                    return false;              
                cmdStr = string.Format(":CHAN{0}:INP DC50", socketIndex, i.ToString());
                if (!OnSendMsg(cmdStr,socketIndex))
                    return false;              
                cmdStr = string.Format(":CHAN{0}:SCAL 0.2V", socketIndex, i.ToString());
                if (!OnSendMsg(cmdStr,socketIndex))
                    return false;           
            }

            if (!OnSendMsg(":TIM:SCAL 0.00000005s", socketIndex))
                return false;           
            for (i = 1; i <= 2; i++)
            {
                cmdStr = string.Format(":MEAS:SOUR CHAN{0}", i.ToString());
                if (!OnSendMsg(cmdStr,socketIndex))
                    return false;            
                if (!OnSendMsg(":MEAS:Vaverage", socketIndex))
                    return false;              
                if (!OnSendMsg(":MEAS:VPP", socketIndex))
                    return false;
            }

            if (!OnSendMsg(":MEAS:V AVG", socketIndex))
                return false;        
            for (i = 1; i <= 2; i++)
            {
                cmdStr = string.Format(":CHAN%d:OFFS 0.00V", i.ToString());
                if (!OnSendMsg(cmdStr,socketIndex))
                    return false;           
            }
            return true;
        }


        /// <summary>
        /// MSO254A 4通道幅度校准初始化
        /// </summary>
        /// <returns></returns>
        public static bool InitMSO254AFourChannel(int socketIndex)
        {
            int i = 1;
            string cmdStr = "";
            for (i = 1; i <= 2; i++)
            {
                cmdStr = string.Format(":CHAN{0}:DISP ON", i.ToString());
                if (!OnSendMsg(cmdStr,socketIndex))
                    return false;
                Thread.Sleep(50);
                cmdStr = string.Format(":CHAN{0}:INP DC50", i.ToString());
                if (!OnSendMsg(cmdStr,socketIndex))
                    return false;
                Thread.Sleep(50);
                cmdStr = string.Format(":CHAN{0}:SCAL 0.2V", i.ToString());
                if (!OnSendMsg(cmdStr,socketIndex))
                    return false;
                Thread.Sleep(50);
            }

            if (!OnSendMsg(":TIM:SCAL 0.00000005s", socketIndex))
                return false;
            Thread.Sleep(50);

            for (i = 1; i <= 2; i++)
            {
                cmdStr = string.Format(":MEAS:SOUR CHAN{0}", i.ToString());
                if (!OnSendMsg(cmdStr,socketIndex))
                    return false;
                Thread.Sleep(50);

                if (!OnSendMsg(":MEAS:Vaverage", socketIndex))
                    return false;
                Thread.Sleep(50);

                if (!OnSendMsg(":MEAS:VPP", socketIndex))
                    return false;
            }

            if (!OnSendMsg(":MEAS:V AVG", socketIndex))
                return false;
            Thread.Sleep(100);

            for (i = 1; i <= 2; i++)
            {
                cmdStr = string.Format(":CHAN%d:OFFS 0.00V", i.ToString());
                if (!OnSendMsg(cmdStr,socketIndex))
                    return false;
                Thread.Sleep(600);
            }
            return true;
        }

        /// <summary>
        /// 设置MSO254A时基
        /// </summary>
        /// <param name="time"></param>
        /// <returns></returns>
        public static bool SetMSO254ATime(double time, int socketIndex)
        {
            string cmdStr = "";
            cmdStr = string.Format("TIM:SCAL {0}", time.ToString("0.00000000000"));	//精确到1ps
            if (!OnSendMsg(cmdStr,socketIndex))
                return false;        
            return true;
        }

        /// <summary>
        /// 读MSO254A时基
        /// </summary>
        /// <param name="pTime"></param>
        /// <returns></returns>
        public static bool GetMSO254ATime(out double pTime,int socketIndex)
        {
            int i, len = 0;
            string strRecv = "";
            pTime = 0;
            //发读指令
            OnSendMsg(":TIM:SCAL?", socketIndex);
            if (!ReceiveServerMsg(out strRecv, out len,socketIndex))
            {
                for (i = 0; i < 5; i++)
                {
                    OnSendMsg(":TIM:SCAL?", socketIndex);
                    if (ReceiveServerMsg(out strRecv, out len,socketIndex))
                    {
                        break;
                    }
                }
                //5次尝试都未能读出
                if (i == 5)
                    return false;
            }
            if (len < 1)
                return false;
            pTime = ConvertFunc.GetValueFromString(strRecv);
            return true;
        }

        /// <summary>
        /// MSO254A 通道阻抗设置 channel取值范围1/2/3/4
        /// </summary>
        /// <param name="imp"></param>
        /// <param name="channel"></param>
        /// <returns></returns>
        public static bool SetMSO254AOhm(int imp, int channel,int socketIndex)
        {
            string cmdStr = "";
            if (imp == TywCommon.OHM_1M)
            {
                //输入阻抗为1M欧姆
                cmdStr = string.Format(":CHAN{0}:INP DC", channel.ToString());
            }
            else
            {
                //输入阻抗为50欧姆
                cmdStr = string.Format(":CHAN{0}:INP DC50", channel.ToString());
            }
            if (!OnSendMsg(cmdStr,socketIndex))
                return false;
            return true;
        }


        /// <summary>
        /// MSO254A设置垂直偏移  CHAN4:OFFS 0.1\r\n
        /// </summary>
        /// <param name="pos"></param>
        /// <param name="channel"></param>
        /// <returns></returns>
        public static bool SetMSO254AYPOS(double pos, int channel,int socketIndex)
        {
            string cmdStr = "";
            cmdStr = string.Format("CHAN{0}:OFFS {1}", channel.ToString(), pos.ToString("0.000"));  //精确到1mv 
            if (!OnSendMsg(cmdStr,socketIndex))
                return false;
            return true;
        }

        /// <summary>
        /// 读通道垂直位置
        /// </summary>
        /// <param name="pos"></param>
        /// <param name="channel"></param>
        /// <returns></returns>
        public static bool GetMSO254AYPOS(out double pos, int channel,int socketIndex)
        {
            int i, len = 0;
            string cmdStr = "";
            string strRecv = "";
            pos = 0;
            cmdStr = string.Format("CHAN{0}:OFFS?", channel);  //精确到1mv 
            //发读指令
            OnSendMsg(cmdStr,socketIndex);
            if (!ReceiveServerMsg(out strRecv, out len,socketIndex))
            {
                for (i = 0; i < 5; i++)
                {
                    OnSendMsg(cmdStr,socketIndex);
                    if (ReceiveServerMsg(out strRecv, out len,socketIndex))
                    {
                        break;
                    }
                }
                //5次尝试都未能读出
                if (i == 5)
                    return false;
            }
            if (len < 1)
                return false;
            pos = ConvertFunc.GetValueFromString(strRecv);
            return true;
        }

        /// <summary>
        /// 设置通道垂直刻度
        /// </summary>
        /// <param name="yScale"></param>
        /// <param name="channel"></param>
        /// <returns></returns>
        public static bool SetMSO254AYScale(double yScale, int channel,int socketIndex)
        {
            string cmdStr = "";
            cmdStr = string.Format("CHAN{0}:SCALE {1}", channel.ToString(), yScale.ToString("0.000"));  //精确到1mv 
            if (!OnSendMsg(cmdStr,socketIndex))
                return false;
            return true;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="pScale"></param>
        /// <param name="channel"></param>
        /// <returns></returns>
        public static bool GetMSO254AYScale(out double pScale, int channel,int socketIndex)
        {
            int i, len = 0;
            string cmdStr = "";
            string strRecv = "";
            pScale = 0;
            cmdStr = string.Format("CHAN{0}:SCALE?", channel);  //精确到1mv 
            //发读指令
            OnSendMsg(cmdStr,socketIndex);
            if (!ReceiveServerMsg(out strRecv, out len,socketIndex))
            {
                for (i = 0; i < 5; i++)
                {
                    OnSendMsg(cmdStr,socketIndex);
                    if (ReceiveServerMsg(out strRecv, out len,socketIndex))
                    {
                        break;
                    }
                }
                //5次尝试都未能读出
                if (i == 5)
                    return false;
            }
            if (len < 1)
                return false;
            pScale = ConvertFunc.GetValueFromString(strRecv);
            return true;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="xScale"></param>
        /// <returns></returns>
        public static bool SetMSO254AXScale(double xScale,int socketIndex)
        {
            string cmdStr = "";
            cmdStr = string.Format("HORIZONTAL:SCALE {0}", xScale.ToString("0.000000000"));
            if (!OnSendMsg(cmdStr,socketIndex))
                return false;
            return true;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="pScale"></param>
        /// <returns></returns>
        public static bool GetMSO254AXScale(out double pScale,int socketIndex)
        {
            int i, len = 0;
            string strRecv = "";
            pScale = 0;
            //发读指令
            OnSendMsg("HORIZONTAL:SCALE?", socketIndex);
            if (!ReceiveServerMsg(out strRecv, out len,socketIndex))
            {
                for (i = 0; i < 5; i++)
                {
                    OnSendMsg("HORIZONTAL:SCALE?", socketIndex);
                    //Thread.Sleep(WAIT_READ);
                    if (ReceiveServerMsg(out strRecv, out len,socketIndex))
                    {
                        break;
                    }
                }
                //5次尝试都未能读出
                if (i == 5)
                    return false;
            }
            if (len < 1)
                return false;
            pScale = ConvertFunc.GetValueFromString(strRecv);
            return true;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="pVpp"></param>
        /// <param name="channel"></param>
        /// <returns></returns>
        public static bool GetMSO254AVpp(out double pVpp, int channel, int socketIndex )
        {
            int i, len = 0;
            string cmdStr = "";
            string strRecv = "";
            pVpp = 0;
            cmdStr = string.Format(":MEAS:SOUR CHAN{0}", channel);
            //发读指令
            OnSendMsg(cmdStr,socketIndex);
            OnSendMsg(":MEAS:VPP?", socketIndex);
            if (!ReceiveServerMsg(out strRecv, out len,socketIndex))
            {
                for (i = 0; i < 5; i++)
                {
                    OnSendMsg(cmdStr,socketIndex);
                    OnSendMsg(":MEAS:VPP?", socketIndex);
                    Thread.Sleep(WAIT_READ);
                    if (ReceiveServerMsg(out strRecv, out len,socketIndex))
                    {
                        break;
                    }
                }
                //5次尝试都未能读出
                if (i == 5)
                    return false;
            }
            if (len < 1)
                return false;
            pVpp = ConvertFunc.GetValueFromString(strRecv);
            return true;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="pVoft"></param>
        /// <param name="channel"></param>
        /// <returns></returns>
        public static bool GetMSO254ADC(out double pVoft, int channel,int socketIndex)
        {
            int i, len = 0;
            string cmdStr = "";
            string strRecv = "";
            pVoft = 0;
            cmdStr = string.Format(":MEAS:SOUR CHAN{0}", channel);
            //发读指令
            OnSendMsg(cmdStr,socketIndex);
            Thread.Sleep(WAIT_WRITE_S);
            OnSendMsg(":MEAS:Vaverage?", socketIndex);
            Thread.Sleep(WAIT_READ);
            if (!ReceiveServerMsg(out strRecv, out len,socketIndex))
            {
                for (i = 0; i < 5; i++)
                {
                    OnSendMsg(cmdStr,socketIndex);
                    Thread.Sleep(WAIT_WRITE_S);
                    OnSendMsg(":MEAS:Vaverage?", socketIndex);
                    Thread.Sleep(WAIT_READ);
                    if (ReceiveServerMsg(out strRecv, out len,socketIndex))
                    {
                        break;
                    }
                }
                //5次尝试都未能读出
                if (i == 5)
                    return false;
            }
            if (len < 1)
                return false;
            pVoft = ConvertFunc.GetValueFromString(strRecv);
            return true;
        }


        //***************************************** MXR604A示波器 ******************************************
        //***************************************** MXR604A示波器 ******************************************
        //***************************************** MXR604A示波器 ******************************************
        //***************************************** MXR604A示波器 ******************************************
        /// <summary>
        /// 复位MXR604A
        /// </summary>
        /// <returns></returns>
        public static bool RstMXR604A(int socketIndex)
        {
            if (!OnSendMsg("*RST", socketIndex))
            {
                return false;
            }
            Thread.Sleep(3000);
            return true;
        }
        /// <summary>
        /// MXR604A自动1次
        /// </summary>
        /// <returns></returns>
        public static bool AutoSetMXR604A(int socketIndex)
        {
            if (!OnSendMsg("AUT", socketIndex))
            {
                return false;
            }
            else
            {
                Thread.Sleep(5000);
                return true;
            }
        }
        /// <summary>
        /// MXR604A 通道幅度校准初始化
        /// </summary>
        /// <param name="osc_channel">示波器通道1-4</param>
        /// <returns></returns>
        public static bool InitMXR604AOneChannel(int osc_channel,int socketIndex)
        {
            string cmdStr = "";
            if (osc_channel < 1 || osc_channel > 4)
                return false;
            cmdStr = string.Format(":CHAN{0}:DISP ON", osc_channel.ToString());
            if (!OnSendMsg(cmdStr,socketIndex))
                    return false;
            cmdStr = string.Format(":CHAN{0}:INP DC50", osc_channel.ToString());
            if (!OnSendMsg(cmdStr,socketIndex))
                    return false;
            cmdStr = string.Format(":CHAN{0}:SCAL 0.2V",osc_channel.ToString());
            if (!OnSendMsg(cmdStr,socketIndex))
                    return false;          
            if (!OnSendMsg(":TIM:SCAL 0.00000005s", socketIndex))
                return false;
            cmdStr = string.Format(":TRIG:EDGE:SOUR  CHAN{0}", osc_channel.ToString());
            if (!OnSendMsg(cmdStr,socketIndex))
                return false;     
            cmdStr = string.Format(":MEAS:SOUR CHAN{0}", osc_channel.ToString());
            if (!OnSendMsg(cmdStr,socketIndex))
                    return false;
            if (!OnSendMsg(":MEAS:FREQ", socketIndex))
                return false;
            if (!OnSendMsg(":MEAS:Vaverage", socketIndex))
                    return false;
            if (!OnSendMsg(":MEAS:VPP", socketIndex))
                    return false;
            cmdStr = string.Format(":CHAN%d:OFFS 0.00V", osc_channel.ToString());
            if (!OnSendMsg(cmdStr,socketIndex))
                    return false;      
            return true;
        }

        /// <summary>
        /// MXR604A 通道1-2通道幅度校准初始化
        /// </summary>
        /// <returns></returns>
        public static bool InitMXR604ATwoChannel(int socketIndex)
        {
            int i = 1;
            string cmdStr = "";
            for (i = 1; i <= 2; i++)
            {
                cmdStr = string.Format(":CHAN{0}:DISP ON", i.ToString());
                if (!OnSendMsg(cmdStr,socketIndex))
                    return false;
                cmdStr = string.Format(":CHAN{0}:INP DC50", i.ToString());
                if (!OnSendMsg(cmdStr,socketIndex))
                    return false;
                cmdStr = string.Format(":CHAN{0}:SCAL 0.2V", i.ToString());
                if (!OnSendMsg(cmdStr,socketIndex))
                    return false;
            }

            if (!OnSendMsg(":TIM:SCAL 0.00000005s", socketIndex))
                return false;
            for (i = 1; i <= 2; i++)
            {
                cmdStr = string.Format(":MEAS:SOUR CHAN{0}", i.ToString());
                if (!OnSendMsg(cmdStr,socketIndex))
                    return false;
                if (!OnSendMsg(":MEAS:Vaverage", socketIndex))
                    return false;
                if (!OnSendMsg(":MEAS:VPP", socketIndex))
                    return false;
            }

            
            for (i = 1; i <= 2; i++)
            {
                cmdStr = string.Format(":CHAN%d:OFFS 0.00V", i.ToString());
                if (!OnSendMsg(cmdStr,socketIndex))
                    return false;
            }
            return true;
        }


        /// <summary>
        /// MXR604A 4通道幅度校准初始化
        /// </summary>
        /// <returns></returns>
        public static bool InitMXR604AFourChannel(int socketIndex)
        {
            int i = 1;
            string cmdStr = "";
            for (i = 1; i <= 4; i++)
            {
                cmdStr = string.Format(":CHAN{0}:DISP ON", i.ToString());
                if (!OnSendMsg(cmdStr,socketIndex))
                    return false;
                Thread.Sleep(50);
                cmdStr = string.Format(":CHAN{0}:INP DC50", i.ToString());
                if (!OnSendMsg(cmdStr,socketIndex))
                    return false;
                Thread.Sleep(50);
                cmdStr = string.Format(":CHAN{0}:SCAL 0.2V", i.ToString());
                if (!OnSendMsg(cmdStr,socketIndex))
                    return false;
                Thread.Sleep(50);
            }

            if (!OnSendMsg(":TIM:SCAL 0.00000005s", socketIndex))
                return false;
            Thread.Sleep(50);

            for (i = 1; i <= 4; i++)
            {
                cmdStr = string.Format(":MEAS:SOUR CHAN{0}", i.ToString());
                if (!OnSendMsg(cmdStr,socketIndex))
                    return false;
                Thread.Sleep(50);

                if (!OnSendMsg(":MEAS:Vaverage", socketIndex))
                    return false;
                Thread.Sleep(50);

                if (!OnSendMsg(":MEAS:VPP", socketIndex))
                    return false;
            }


            for (i = 1; i <= 2; i++)
            {
                cmdStr = string.Format(":CHAN%d:OFFS 0.00V", i.ToString());
                if (!OnSendMsg(cmdStr,socketIndex))
                    return false;
                Thread.Sleep(600);
            }
            return true;
        }

        /// <summary>
        /// 设置MXR604A时基
        /// </summary>
        /// <param name="time"></param>
        /// <returns></returns>
        public static bool SetMXR604ATime(double time,int socketIndex)
        {
            string cmdStr = "";
            cmdStr = string.Format("TIM:SCAL {0}", time.ToString("0.00000000000"));	//精确到10ps
            if (!OnSendMsg(cmdStr,socketIndex))
                return false;
            return true;
        }

        /// <summary>
        /// 读MXR604A时基
        /// </summary>
        /// <param name="pTime"></param>
        /// <returns></returns>
        public static bool GetMXR604ATime(out double pTime,int socketIndex)
        {
            int i, len = 0;
            string strRecv = "";
            pTime = 0;
            //发读指令
            OnSendMsg(":TIM:SCAL?", socketIndex);
            if (!ReceiveServerMsg(out strRecv, out len,socketIndex))
            {
                for (i = 0; i < 5; i++)
                {
                    OnSendMsg(":TIM:SCAL?", socketIndex);
                    if (ReceiveServerMsg(out strRecv, out len,socketIndex))
                    {
                        break;
                    }
                }
                //5次尝试都未能读出
                if (i == 5)
                    return false;
            }
            if (len < 1)
                return false;
            pTime = ConvertFunc.GetValueFromString(strRecv);
            return true;
        }

        /// <summary>
        /// MXR604A 通道阻抗设置 channel取值范围1/2/3/4
        /// </summary>
        /// <param name="imp"></param>
        /// <param name="channel"></param>
        /// <returns></returns>
        public static bool SetMXR604AOhm(int imp, int channel,int socketIndex)
        {
            string cmdStr = "";
            if (imp == TywCommon.OHM_1M)
            {
                //输入阻抗为1M欧姆
                cmdStr = string.Format(":CHAN{0}:INP DC", channel.ToString());
            }
            else
            {
                //输入阻抗为50欧姆
                cmdStr = string.Format(":CHAN{0}:INP DC50", channel.ToString());
            }
            if (!OnSendMsg(cmdStr,socketIndex))
                return false;
            return true;
        }


        /// <summary>
        /// MXR604A设置垂直偏移  CHAN4:OFFS 0.1\r\n
        /// </summary>
        /// <param name="pos"></param>
        /// <param name="channel"></param>
        /// <returns></returns>
        public static bool SetMXR604AYPOS(double pos, int channel,int socketIndex)
        {
            string cmdStr = "";
            cmdStr = string.Format("CHAN{0}:OFFS {1}", channel.ToString(), pos.ToString("0.000"));  //精确到1mv 
            if (!OnSendMsg(cmdStr,socketIndex))
                return false;
            return true;
        }

        /// <summary>
        /// 读MXR604A通道垂直位置
        /// </summary>
        /// <param name="pos"></param>
        /// <param name="channel"></param>
        /// <returns></returns>
        public static bool GetMXR604AYPOS(out double pos, int channel,int socketIndex)
        {
            int i, len = 0;
            string cmdStr = "";
            string strRecv = "";
            pos = 0;
            cmdStr = string.Format("CHAN{0}:OFFS?", channel);  //精确到1mv 
            //发读指令
            OnSendMsg(cmdStr,socketIndex);
            if (!ReceiveServerMsg(out strRecv, out len,socketIndex))
            {
                for (i = 0; i < 5; i++)
                {
                    OnSendMsg(cmdStr,socketIndex);
                    if (ReceiveServerMsg(out strRecv, out len,socketIndex))
                    {
                        break;
                    }
                }
                //5次尝试都未能读出
                if (i == 5)
                    return false;
            }
            if (len < 1)
                return false;
            pos = ConvertFunc.GetValueFromString(strRecv);
            return true;
        }

        /// <summary>
        /// 设置MXR604A通道垂直刻度
        /// </summary>
        /// <param name="yScale"></param>
        /// <param name="channel"></param>
        /// <returns></returns>
        public static bool SetMXR604AYScale(double yScale, int channel,int socketIndex)
        {
            string cmdStr = "";
            cmdStr = string.Format("CHAN{0}:SCALE {1}", channel.ToString(), yScale.ToString("0.000"));  //精确到1mv 
            if (!OnSendMsg(cmdStr,socketIndex))
                return false;
            return true;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="pScale"></param>
        /// <param name="channel"></param>
        /// <returns></returns>
        public static bool GetMXR604AYScale(out double pScale, int channel,int socketIndex)
        {
            int i, len = 0;
            string cmdStr = "";
            string strRecv = "";
            pScale = 0;
            cmdStr = string.Format("CHAN{0}:SCALE?", channel);  //精确到1mv 
            //发读指令
            OnSendMsg(cmdStr,socketIndex);
            if (!ReceiveServerMsg(out strRecv, out len,socketIndex))
            {
                for (i = 0; i < 5; i++)
                {
                    OnSendMsg(cmdStr,socketIndex);
                    if (ReceiveServerMsg(out strRecv, out len,socketIndex))
                    {
                        break;
                    }
                }
                //5次尝试都未能读出
                if (i == 5)
                    return false;
            }
            if (len < 1)
                return false;
            pScale = ConvertFunc.GetValueFromString(strRecv);
            return true;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="xScale"></param>
        /// <returns></returns>
        public static bool SetMXR604AXScale(double xScale,int socketIndex)
        {
            string cmdStr = "";
            cmdStr = string.Format("HORIZONTAL:SCALE {0}", xScale.ToString("0.000000000"));
            if (!OnSendMsg(cmdStr,socketIndex))
                return false;
            return true;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="pScale"></param>
        /// <returns></returns>
        public static bool GetMXR604AXScale(out double pScale,int socketIndex)
        {
            int i, len = 0;
            string strRecv = "";
            pScale = 0;
            //发读指令
            OnSendMsg("HORIZONTAL:SCALE?", socketIndex);
            if (!ReceiveServerMsg(out strRecv, out len,socketIndex))
            {
                for (i = 0; i < 5; i++)
                {
                    OnSendMsg("HORIZONTAL:SCALE?", socketIndex);
                    //Thread.Sleep(WAIT_READ);
                    if (ReceiveServerMsg(out strRecv, out len,socketIndex))
                    {
                        break;
                    }
                }
                //5次尝试都未能读出
                if (i == 5)
                    return false;
            }
            if (len < 1)
                return false;
            pScale = ConvertFunc.GetValueFromString(strRecv);
            return true;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="pVpp"></param>
        /// <param name="channel"></param>
        /// <returns></returns>
        public static bool GetMXR604AVpp(out double pVpp, int channel,int socketIndex)
        {
            int i, len = 0;
            string cmdStr = "";
            string strRecv = "";
            pVpp = 0;
            cmdStr = string.Format(":MEAS:SOUR CHAN{0}", channel);
            //发读指令
            OnSendMsg(cmdStr,socketIndex);
            OnSendMsg(":MEAS:VPP?", socketIndex);
            if (!ReceiveServerMsg(out strRecv, out len,socketIndex))
            {
                for (i = 0; i < 5; i++)
                {
                    OnSendMsg(cmdStr,socketIndex);
                    OnSendMsg(":MEAS:VPP?", socketIndex);
                    Thread.Sleep(WAIT_READ);
                    if (ReceiveServerMsg(out strRecv, out len,socketIndex))
                    {
                        break;
                    }
                }
                //5次尝试都未能读出
                if (i == 5)
                    return false;
            }
            if (len < 1)
                return false;
            pVpp = ConvertFunc.GetValueFromString(strRecv);
            return true;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="pVoft"></param>
        /// <param name="channel"></param>
        /// <returns></returns>
        public static bool GetMXR604ADC(out double pVoft, int channel,int socketIndex)
        {
            int i, len = 0;
            string cmdStr = "";
            string strRecv = "";
            pVoft = 0;
            cmdStr = string.Format(":MEAS:SOUR CHAN{0}", channel);
            //发读指令
            OnSendMsg(cmdStr,socketIndex);
            Thread.Sleep(WAIT_WRITE_S);
            OnSendMsg(":MEAS:Vaverage?", socketIndex);
            Thread.Sleep(WAIT_READ);
            if (!ReceiveServerMsg(out strRecv, out len,socketIndex))
            {
                for (i = 0; i < 5; i++)
                {
                    OnSendMsg(cmdStr,socketIndex);
                    Thread.Sleep(WAIT_WRITE_S);
                    OnSendMsg(":MEAS:Vaverage?", socketIndex);
                    Thread.Sleep(WAIT_READ);
                    if (ReceiveServerMsg(out strRecv, out len,socketIndex))
                    {
                        break;
                    }
                }
                //5次尝试都未能读出
                if (i == 5)
                    return false;
            }
            if (len < 1)
                return false;
            pVoft = ConvertFunc.GetValueFromString(strRecv);
            return true;
        }


        //-------------------------------------------------NRX功率计：
        //-------------------------------------------------NRX功率计：
        //-------------------------------------------------NRX功率计：
        //-------------------------------------------------NRX功率计：
        /// <summary>
        /// 复位功功率计
        /// </summary>
        /// <returns></returns>
        public static bool RstNRX(int socketIndex)
        {
            if(!OnSendMsg("*RST", socketIndex)) 
	        {	
		        return false;
	        }
	        Thread.Sleep(2000);
	        return true;	
        }

        /// <summary>
        /// 设置功率计频率
        /// </summary>
        /// <param name="carrierFreq"></param>
        /// <returns></returns>
        public static bool SetNRXCarrierFreq(double carrierFreq,int socketIndex)
        {
            string cmdStr = "";
            carrierFreq =TywCommon.MaxMin(carrierFreq,50000000.0,100000000000.0);//50MHz-100GHz
            cmdStr = string.Format("SENS5:FREQ {0}", carrierFreq.ToString("0.000"));  
            if (!OnSendMsg(cmdStr,socketIndex))
                return false;
            Thread.Sleep(200);
            return true;

        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="_carrierFreq"></param>
        /// <returns></returns>
        public static bool GetNRXCarrierFreq(out double _carrierFreq,int socketIndex)
        {
            int i, len = 0;
            string strRecv = "";
            _carrierFreq = 0;
            //发读指令
            OnSendMsg("SENS5:FREQ?", socketIndex);
            Thread.Sleep(WAIT_READ);
            if (!ReceiveServerMsg(out strRecv, out len,socketIndex))
            {
                for (i = 0; i < 5; i++)
                {
                    Thread.Sleep(100);
                    OnSendMsg("SENS5:FREQ?", socketIndex);
                    Thread.Sleep(WAIT_READ);
                    if (ReceiveServerMsg(out strRecv, out len,socketIndex))
                    {
                        break;
                    }
                }
                //5次尝试都未能读出
                if (i == 5)
                    return false;
            }
            if (len < 1)
                return false;
            _carrierFreq = ConvertFunc.GetValueFromString(strRecv);
            return true;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="pow"></param>
        /// <param name="sleep_ms"></param>
        /// <returns></returns>
        public static bool GetNRXPower1(out double pow,int sleep_ms,int socketIndex)
        {
            int i, len = 0;
            string strRecv = "";
            pow = 0;
            //发读指令
            OnSendMsg("MEAS1?", socketIndex);
            Thread.Sleep(sleep_ms);
            if (!ReceiveServerMsg(out strRecv, out len,socketIndex))
            {
                for (i = 0; i < 5; i++)
                {
                    Thread.Sleep(100);
                    OnSendMsg("MEAS1?", socketIndex);
                    Thread.Sleep(sleep_ms);
                    if (ReceiveServerMsg(out strRecv, out len,socketIndex))
                    {
                        break;
                    }
                }
                //5次尝试都未能读出
                if (i == 5)
                    return false;
            }
            if (len < 1)
                return false;
            pow = ConvertFunc.GetValueFromString(strRecv);
            return true;
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="pow"></param>
        /// <param name="sleep_ms"></param>
        /// <returns></returns>
        public static bool GetNRXPower2(out double pow,int sleep_ms,int socketIndex)
        {
            int i, len = 0;
            string strRecv = "";
            pow = 0;
            //发读指令
            OnSendMsg("FETC?", socketIndex);
            Thread.Sleep(sleep_ms);
            if (!ReceiveServerMsg(out strRecv, out len,socketIndex))
            {
                for (i = 0; i < 5; i++)
                {
                    Thread.Sleep(100);
                    OnSendMsg("FETC?", socketIndex);
                    Thread.Sleep(sleep_ms);
                    if (ReceiveServerMsg(out strRecv, out len,socketIndex))
                    {
                        break;
                    }
                }
                //5次尝试都未能读出
                if (i == 5)
                    return false;
            }
            if (len < 1)
                return false;
            pow = ConvertFunc.GetValueFromString(strRecv);
            return true;
        }




    }
}
