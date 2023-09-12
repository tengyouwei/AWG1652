using Microsoft.Win32;
using System;
using System.ComponentModel;
using System.IO;
using System.Windows;

namespace AWGTestCerter.ParseData
{
    /// <summary>
    /// .awf文件头
    /// </summary>
    public struct AWFileInfo
    {
        public UInt16 _identfyLow; //帧识别码L
        public UInt16 _identfyHigh;//帧识别码H
        public UInt16 _framwareVersion;//固件版本
        public UInt16 _dataWidth;   //数据位宽
        public UInt64 _lengthBytes; //波形数据长度,字节
        public UInt64 _sampleRate;  //采样率,Hz
        public UInt16 _instrument;  //仪器识别码
        public UInt16 _fileFormat;    //文件存储格式识别码，I路、Q路、IQ交叉、RF
        public UInt16 _genSource;     //文件产生来源，标识哪个插件生成的
        public UInt16 _playflag;      //文件播放时处理标识，1：Real模式直接 2：Duc模式硬件上变频

        public UInt16 _defaultInfo3;//预留3
        public UInt16 _defaultInfo4;//预留4
        public UInt16 _defaultInfo5;//预留5
        public UInt16 _defaultInfo6;//预留6
        public UInt16 _defaultInfo7;//预留7
        public UInt16 _defaultInfo8;//预留8
        public UInt16 _defaultInfo9;//预留9
        public UInt16 _defaultInfo10;//预留10
    }

  
    public static partial class ParseData
    {

        public const string EditWaveDir = "./EditWave/";
        public const string DefineDir = "./DEFINE/";
        public const string SeqDir = "./SEQ/";
        public const string AWGDir = "./AWG/";
        public const string HsspDir = "./HSSP/";
        public const string MutliWave = "./MultiWaveSequency/";
        public const string IFRFDir = ".\\SignalSimulator\\IFRFSignal\\";
        public const string IQDir = ".\\SignalSimulator\\IQSignal\\";
        public const string OFDMDir = ".\\SignalSimulator\\OFDMSignal\\";
        public const string RadarDir = ".\\SignalSimulator\\RadarSignal\\";

        public const string MapBpsk = "\\SignalSimulator\\MAPS\\DefaultMap\\bpsk_star.conf";
        public const string MapQpsk = "\\SignalSimulator\\MAPS\\DefaultMap\\qpsk_star.conf";
        public const string MapQam16 = "\\SignalSimulator\\MAPS\\DefaultMap\\qam16_star.conf";
        public const string MapQam32 = "\\SignalSimulator\\MAPS\\DefaultMap\\qam32_star.conf";
        public const string MapQam64 = "\\SignalSimulator\\MAPS\\DefaultMap\\qam64_star.conf";
        public const string MapQam128 = "\\SignalSimulator\\MAPS\\DefaultMap\\qam128_star.conf";
        public const string MapFile = "\\SignalSimulator\\MAPS\\star.conf";

        public const int ChannelStatusHeight = 50;
        public const int OtherHeight = 460;
        public const int ChannelChartHeight = 210;

        public const int INIT_CHAR_VIEW_LEN = 40000;//显示控件chart一次性能显示的最大字节数
        public const int MOVE_MIN_LEN = 200;
        public const long PLUG_MAX_NUM = 4000000;//信号模拟插件生成最大数据量

        public static int FrameHeaderSize = 48;     // awf 数据帧头字节数,48个字节
        public static int SEQCyCSize = 1024;        //各子序列循环次数信息512个*2=1024字节
        public static int SEQSubLengthSize = 4096;  //各子序列数据量信息512个*8=4096字节
        public static int SEQFrameHeaderSize = 5168; // seq 数据帧头字节数（包含1、普通头信息48字节，同awf文件；2、各子序列循环次数信息512个*2=1024字节；3、各子序列数据量信息512个*8=4096字节，（2）+（3）=5120字节）
        public static UInt64 SEQSubMaxByteSize = 10000000;//多波序列中单个子序列字节数10000000，即5MSa
        public static AWFileInfo GetAwfileInfoInit( )
        {
            AWFileInfo fileHead = new AWFileInfo();
            fileHead._identfyHigh = 0x3412;
            fileHead._identfyLow = 0xBBAA;
            fileHead._sampleRate = 5000000000;//采样率：         5GSa/s
            fileHead._lengthBytes = 0;        //裸数据长度：     0字节
            fileHead._fileFormat = 1;         //文件存储格式：   0x0001表示单路I数据；0x0002表示单路Q数据；0x0003表示IQ数据；0x0004表示RF数据；
            fileHead._framwareVersion = 1;    //固件版本：       V1.0
            fileHead._instrument = 0x0002;    //仪器型号：       1652B
            fileHead._dataWidth = 2;          //数据位宽：       2Bytes，16bit

            fileHead._genSource = 0;    //文件数据来源 0：Com 通用数据文件
            fileHead._playflag  = 1;    //播放是否需要硬件上变频 1：Real模式不需要  2：DUC模式需要（对应IQ插件下IQ存储格式的数据文件）
            fileHead._defaultInfo3  = 0;
            fileHead._defaultInfo4  = 0;
            fileHead._defaultInfo5  = 0;
            fileHead._defaultInfo6  = 0;
            fileHead._defaultInfo7  = 0;//循环播放次数 0-65535
            fileHead._defaultInfo8  = 0;
            fileHead._defaultInfo9  = 0;
            fileHead._defaultInfo10 = 0;

            return fileHead;
        }
        /// <summary>
        /// 从offset位置开始读最大长度为dataArray.Length的数据文件
        /// </summary>
        /// <param name="origStream"></param>
        /// <param name="dataArray"></param>
        /// <param name="offset">偏移（单位：字节）</param>
        /// <returns></returns>
        public static uint ReadUInt16FromStream(FileStream origStream, UInt16[] dataArray, long offset = 0)
        {
            int byteSize = dataArray.Length * sizeof(UInt16);
            byte[] tempArray = null;
            int readLength = 0;
            if (offset < 0)
            {
                offset = 0;
            }
            try
            {
                tempArray = new byte[byteSize];
                origStream.Seek(offset, SeekOrigin.Begin);
                readLength = origStream.Read(tempArray, 0, tempArray.Length);
                readLength = readLength / sizeof(UInt16);
                for (int i = 0; i < readLength; i++)
                {
                    dataArray[i] = BitConverter.ToUInt16(tempArray, i * sizeof(UInt16));
                }
            }
            catch
            {
                readLength = 0;
            }         
            return (uint)readLength;
        }
        /// <summary>
        /// 从offset位置开始，读取全部文件
        /// </summary>
        /// <param name="origStream"></param>
        /// <param name="dataArray"></param>
        /// <param name="offset">偏移（单位：字节）</param>
        /// <returns></returns>
        public static uint ReadUInt32FromStream(FileStream origStream, UInt32[] dataArray, int offset)
        {
            int byteSize = dataArray.Length * sizeof(UInt32);
            byte[] tempArray = new byte[byteSize];
            if (offset < 0) offset = 0;
            origStream.Seek(offset, SeekOrigin.Begin);
            int readLength = origStream.Read(tempArray, 0, tempArray.Length);
            readLength = readLength / sizeof(UInt32);
            for (int i = 0; i < readLength; i++)
            {
                dataArray[i] = BitConverter.ToUInt32(tempArray, i * sizeof(UInt32));
            }
            return (uint)readLength;
        }
        /// <summary>
        /// 从offset位置开始，读取全部文件
        /// </summary>
        /// <param name="origStream"></param>
        /// <param name="dataArray"></param>
        /// <param name="offset">偏移（单位：字节）</param>
        /// <returns></returns>
        public static uint ReadUInt64FromStream(FileStream origStream, UInt64[] dataArray, int offset)
        {
            int byteSize = dataArray.Length * sizeof(UInt64);
            byte[] tempArray = new byte[byteSize];
            if (offset < 0) offset = 0;
            origStream.Seek(offset, SeekOrigin.Begin);
            int readLength = origStream.Read(tempArray, 0, tempArray.Length);
            readLength = readLength / sizeof(UInt64);
            for (int i = 0; i < readLength; i++)
            {
                dataArray[i] = BitConverter.ToUInt32(tempArray, i * sizeof(UInt64));
            }
            return (uint)readLength;
        }
        /// <summary>
        /// 获取取文件大小（包含文件头信息）字节数
        /// </summary>
        public static long GetFileLength(string filePath)
        {
            FileStream origStream = null;
            if (File.Exists(filePath))
            {
                try
                {
                    using (origStream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                    {
                        long len = origStream.Length;
                        origStream.Close();
                        return len;
                    }
                }catch
                {
                    MessageBox.Show("查看文件是否已打开!");
                    return 0;
                }         
            }
            return 0;
        }


        /// <summary>
        /// 获取取.awf文件裸数据字节数
        /// </summary>
        public static long GetFileNoHeadLength(string filePath)
        {
            FileStream origStream = null;
            if (File.Exists(filePath))
            {
                using (origStream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    long len = origStream.Length - FrameHeaderSize;
                    origStream.Close();
                    return len;
                }
            }
            return 0;
        }

        /// <summary>
        /// 获取取串行数据文件裸数据字节数
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public static long GetHsspFileLength(string filePath)
        {
            FileStream origStream = null;
            if (File.Exists(filePath))
            {
                using (origStream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    long len = origStream.Length;
                    origStream.Close();
                    return len;
                }
            }
            return 0;
        }
        

        /// <summary>
        /// 取文件数据部分，认为Mark1和Mark2也是数据的一部分
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="start">start开始(start==0时从第1个裸数据读起,单位为字节)</param>
        /// <param name="offset">offset长度的数据（单位为字节），offset不可过大，超出可动态分配内存长度</param>
        /// <param name="data"></param>
        /// <returns></returns>
        public static long ParseDat(string filePath, long start, long offset, out int[] data)
        {
            //读取 .awf  
            uint PointNumber = 0;
            FileStream origStream = null;
            start = start + FrameHeaderSize;
            if (File.Exists(filePath))
            {
                using (origStream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    uint allDataLength = 0;//有效数据点数（每个点占2个字节,16bit）
                    if (start + offset < origStream.Length)
                    {
                        allDataLength = (uint)(offset / sizeof(UInt16));//字节数转换为点数
                    }
                    else
                    {
                        allDataLength = (uint)((origStream.Length - start) / sizeof(UInt16));//字节数转换为点数
                    }
                    //分配16位无符号数内存
                    UInt16[] tempArray = new UInt16[allDataLength];
                    data = new int[allDataLength];
                    uint readLength = 0;
                    if (origStream.Position < origStream.Length)
                    {
                        readLength = ParseData.ReadUInt16FromStream(origStream, tempArray, start);
                    }
                    for (int x = 0; x < readLength; x++)
                    {
                        data[x] = (int)(tempArray[x]);
                    }
                    PointNumber = allDataLength;
                }
            }
            else
            {
                data = null;
                return 0;
            }
            return PointNumber;
        }
      
        /// <summary>
        /// 取文件数据部分，认为Mark1和Mark2也是数据的一部分，复用Mark1和Mark2数据
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="start">start开始(start==0时从第1个裸数据读起,单位为字节)</param>
        /// <param name="offset">offset长度的数据（单位为字节），offset不可过大，超出可动态分配内存长度</param>
        /// <param name="data"></param>
        /// <param name="mark1"></param>
        /// <param name="mark2"></param>
        /// <returns></returns>
        public static long ParseDat(string filePath, long start, long offset, out int[] data, out int[] mark1, out int[] mark2)
        {
            FileStream origStream = null;         
            UInt16[] tempArray = null;
            data = null;
            mark1 = null;
            mark2 = null;
            long ReadDataLength = 0;
            string _exName = Path.GetExtension(filePath);
            if (File.Exists(filePath)==false)
            {
                 return 0;
            }
            if (_exName == ".awf")
            {
                start = start + FrameHeaderSize;//跳过文件头 
            }
            else if (_exName == ".seq")
            {
                start = start + SEQFrameHeaderSize;//跳过文件头 
            }
            else if (_exName == ".usbx" || _exName == ".pciex" || _exName == ".fcx")
            {
                start = start + 0;
            }
            else
            {
                return 0;
            }
            uint readLength = 0;
            //打开文件取数据      
            using (origStream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                    if (start + offset < origStream.Length)//有效数据点数（每个点占2个字节,16bit）
                    {
                        ReadDataLength = (uint)(offset / sizeof(UInt16));
                    }
                    else
                    {
                        ReadDataLength = (uint)((origStream.Length - start) / sizeof(UInt16));
                    }
                    try
                    {
                        tempArray = new UInt16[ReadDataLength];
                        data = new int[ReadDataLength]; //数据
                        mark1 = new int[ReadDataLength]; //Mark1
                        mark2 = new int[ReadDataLength]; //Mark2
                    }
                    catch
                    {
                        tempArray = null;
                        data = null;  //数据
                        mark1 = null; //Mark1
                        mark2 = null; //Mark2
                        return 0;
                    }

                    if (origStream.Position < origStream.Length)
                    {
                        readLength = ParseData.ReadUInt16FromStream(origStream, tempArray, start);
                    }

                    for (int x = 0; x < ReadDataLength; x++)
                    {
                        data[x] = (int)(tempArray[x]);   //数据是16bit有效,Mark复用最低2bit数据而已               
                        mark1[x] = (int)((tempArray[x] & 0x0001) == 0 ? 0 : 1);
                        mark2[x] = (int)((tempArray[x] & 0x0002) == 0 ? 0 : 1);
                    }
             }     
            return ReadDataLength;
        }


        /// <summary>
        /// 一次性读取.awf文件，包含Mark。数据量过大出现异常
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="data">out 数据部分</param>
        /// <param name="mark1">out mark1部分</param>
        /// <param name="mark2">out mark2部分</param>
        /// <param name="MaxReadLength">最大读取长度，0时读取全部文件数据，非0时读取Min(MaxReadLength,文件长度)数据</param>
        /// <returns></returns>
        public static bool ParseDatFile(string filePath, out int[] data, out int[] mark1, out int[] mark2, long MaxReadLength=0)
        {
            AWFileInfo _fileHead = new AWFileInfo();           
            UInt16[] tempArrayData = null;
            data = null;
            mark1 = null;
            mark2 = null;
            uint readLength = 0;
            uint awfileHeadlength = (uint)((UInt64)FrameHeaderSize / sizeof(UInt16));
            UInt16[] tempArrayRawData = null;
           
            GetawfFileHeadInfo(filePath, out _fileHead);//文件头信息

            if (File.Exists(filePath) && Path.GetExtension(filePath)==".awf")
            {
                using (FileStream origStreamReadRawData = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite, 8192))
                {
                    if ((_fileHead._identfyLow != 0xBBAA) || (_fileHead._identfyHigh != 0x3412))
                    {
                        MessageBox.Show("数据帧头校验错误!");
                        origStreamReadRawData.Close();
                        return false;
                    }
                    if (_fileHead._lengthBytes != ((UInt64)(origStreamReadRawData.Length - FrameHeaderSize)))
                    {
                        MessageBox.Show("数据长度校验错误!");
                        origStreamReadRawData.Close();
                        return false;
                    }
                    try
                    {
                        if (MaxReadLength > 0 && MaxReadLength < origStreamReadRawData.Length / sizeof(UInt16))
                        {
                            tempArrayRawData = new UInt16[MaxReadLength + awfileHeadlength];//除了数据还有文件头部分
                        }
                        else
                        {
                            tempArrayRawData = new UInt16[origStreamReadRawData.Length / sizeof(UInt16)];//裸数据长度
                        }
                                            
                        
                        readLength = ParseData.ReadUInt16FromStream(origStreamReadRawData, tempArrayRawData, 0);
                        if (readLength == 0)
                        {
                            //一次性读取文件出错，多源于内存分配失败
                            origStreamReadRawData.Close();
                            return false;
                        }
                    }
                    catch
                    {
                        origStreamReadRawData.Close();
                        return false;
                    }
                  
                    //allDataLength = _fileHead._lengthBytes / _fileHead._dataWidth;//文件头中裸数据长度Bytes                
                    try
                    {
                        tempArrayData = new UInt16[readLength - awfileHeadlength];
                        data = new int[readLength - awfileHeadlength];
                        mark1 = new int[readLength - awfileHeadlength];
                        mark2 = new int[readLength - awfileHeadlength];
                    }
                    catch
                    {
                        origStreamReadRawData.Close();          
                        return false;
                    }

                    for (int i = 0; i < tempArrayData.Length; i++)
                    {
                        tempArrayData[i] = tempArrayRawData[i + awfileHeadlength];
                    }

                    for (int j = 0; j < tempArrayData.Length; j++)
                    {
                        data[j]  = (int)(tempArrayData[j]);
                        mark1[j] = (int)((tempArrayData[j] & 0x0001) == 0 ? 0 : 1);
                        mark2[j] = (int)((tempArrayData[j] & 0x0002) == 0 ? 0 : 1);
                    }
                }
                return true;
            }
            else
            {
                data = null;
                mark1 = null;
                mark2 = null;
                return false;
            }
        }
        /// <summary>
        /// 将.awf文件中的数据转换为[-1,+1]之间的绘制数据，保留Mark标记
        /// </summary>
        /// <param name="data"></param>
        /// <param name="hY"></param>
        /// <param name="mark1"></param>
        /// <param name="mark2"></param>
        /// <param name="hYwidth"></param>
        public static void ConvertDataToSeriesPointsWithMark(UInt16 data,out double hY,out int mark1,out int mark2)
        {                 
            mark1 = (int)((data & 0x0001) == 0x0001 ? 1 : 0);
            mark2 = (int)((data & 0x0002) == 0x0002 ? 1 : 0);
            data &= 0xFFFC;//Mark标记不参与量化预算，这里仅仅作为标志位
            hY = ParseData.Quantify(data, 0.0, 65532.0, -1.0, 1.0);
            return;
        }
        /// <summary>
        /// 将.awf文件中的数据转换为[-1,+1]之间的绘制数据，mark位最为数据部分参与量化运算
        /// </summary>
        /// <param name="data"></param>
        /// <param name="hY"></param>
        public static void ConvertDataToSeriesPointsNoMark(UInt16 data, out double hY)
        {         
            hY = ParseData.Quantify(data, 0.0, 65535.0, -1.0, 1.0);
            return;
        }
       
        /// <summary>
        /// 修改文件头信息，如果该文件不存在则返回
        /// </summary>
        /// <param name="fileHead"></param>
        /// <param name="filename"></param>
        /// <returns></returns>
        public static bool ModifyAWFileHeadInfo(AWFileInfo fileHead, string filename)
        {
            if (!File.Exists(filename))
            {
                return false;
            }
            byte[] B = null;
            try
            {
                using (FileStream fw = new FileStream(filename, FileMode.Open))
                {
                    B = BitConverter.GetBytes(fileHead._identfyLow);
                    fw.Write(B, 0, sizeof(UInt16));
                    B = BitConverter.GetBytes(fileHead._identfyHigh);
                    fw.Write(B, 0, sizeof(UInt16));
                    B = BitConverter.GetBytes(fileHead._framwareVersion);
                    fw.Write(B, 0, sizeof(UInt16));
                    B = BitConverter.GetBytes(fileHead._dataWidth);
                    fw.Write(B, 0, sizeof(UInt16));
                    B = BitConverter.GetBytes(fileHead._lengthBytes);
                    fw.Write(B, 0, sizeof(UInt64));
                    B = BitConverter.GetBytes(fileHead._sampleRate);
                    fw.Write(B, 0, sizeof(UInt64));


                    B = BitConverter.GetBytes(fileHead._instrument);
                    fw.Write(B, 0, sizeof(UInt16));
                    B = BitConverter.GetBytes(fileHead._genSource);
                    fw.Write(B, 0, sizeof(UInt16));
                    B = BitConverter.GetBytes(fileHead._fileFormat);
                    fw.Write(B, 0, sizeof(UInt16));
                    B = BitConverter.GetBytes(fileHead._defaultInfo3);

                    fw.Write(B, 0, sizeof(UInt16));
                    B = BitConverter.GetBytes(fileHead._playflag);
                    fw.Write(B, 0, sizeof(UInt16));
                    B = BitConverter.GetBytes(fileHead._defaultInfo4);
                    fw.Write(B, 0, sizeof(UInt16));
                    B = BitConverter.GetBytes(fileHead._defaultInfo5);
                    fw.Write(B, 0, sizeof(UInt16));
                    B = BitConverter.GetBytes(fileHead._defaultInfo6);
                    fw.Write(B, 0, sizeof(UInt16));

                    B = BitConverter.GetBytes(fileHead._defaultInfo7);
                    fw.Write(B, 0, sizeof(UInt16));
                    B = BitConverter.GetBytes(fileHead._defaultInfo8);
                    fw.Write(B, 0, sizeof(UInt16));
                    B = BitConverter.GetBytes(fileHead._defaultInfo9);
                    fw.Write(B, 0, sizeof(UInt16));
                    B = BitConverter.GetBytes(fileHead._defaultInfo10);
                    fw.Write(B, 0, sizeof(UInt16));
                }
            }
            catch
            {
                return false;
            }
           
            return true;
        }

         /// <summary>
        /// 导出文件，将.awf文件导出为txt/csv格式
        /// </summary>
        /// <param name="srcFile"></param>
        /// <param name="destFile"></param>
        /// <param name="_backgroundWorker"></param>
        /// <returns></returns>
        public static bool ExportFile(string srcFile, string destFile, BackgroundWorker _backgroundWorker = null)
        {
            //如果源文件不是.awf文件，返回
            if (!srcFile.Contains(".awf"))
            {
                return false;
            }
            long i = 0;
            int mark1 = 0;
            int mark2 = 0;
            UInt16 data = 0;//16bit无符号整形数据
            double serPoint = 0;//[-1,+1]

            AWFileInfo fileHead = new AWFileInfo();
            GetawfFileHeadInfo(srcFile, out fileHead);//获取文件头信息
            ulong _sampleRate = fileHead._sampleRate;
            ulong _size = fileHead._lengthBytes/sizeof(UInt16);//波形点数


            if (destFile.Contains(".csv") || destFile.Contains(".txt"))
            {
                //读写流进行数据传输         
                using (FileStream fr = new FileStream(srcFile, FileMode.Open))
                {
                    using (FileStream fw = new FileStream(destFile, FileMode.Create))
                    {
                        BinaryReader sr = new BinaryReader(fr);//读.pat文件
                        StreamWriter sw = new StreamWriter(fw);//写.csv或.txt文件

                        double persent = 0;
                        string tBuffer = string.Empty;
                        //写采样率
                        tBuffer = "#CLOCK=" + _sampleRate.ToString() + "\r\n";
                        sw.Write(tBuffer);
                        //写数据量
                        tBuffer = "#SIZE=" + _size.ToString() + "\r\n";
                        sw.Write(tBuffer);
                        //初始化百分比显示标记
                        flag = false;
                        compareValue = 1;
                        //跳过文件头
                        sr.BaseStream.Seek(ParseData.FrameHeaderSize, SeekOrigin.Begin);
                        while (sr.BaseStream.Position < sr.BaseStream.Length)
                        {
                            //读取2个字节
                            data = BitConverter.ToUInt16(sr.ReadBytes(2), 0);
                            //分离出数据和Mark位，Mark和数据最低2位复用
                            ParseData.ConvertDataToSeriesPointsWithMark(data, out serPoint,out mark1,out mark2);                       
                            //将读取的数据写入.csv或.txt文件                     
                            tBuffer = serPoint + "," + mark1 + "," + mark2 + "\r\n";
                            sw.Write(tBuffer);
                            persent = (double)sr.BaseStream.Position / (double)sr.BaseStream.Length;
                            ParseData.ReportPercent(persent, _backgroundWorker);
                            i++;
                        }
                        sw.Close();//不要遗漏，否则会导致最后几个数据没有写入文件
                        sr.Close();
                        fw.Close();
                    }
                    fr.Close();
                }
            }
            return true;
        }
        


        


        /// <summary>
        /// 根据新旧采样率以及数据量最小粒度计算新数据量（单位：Sa）
        /// </summary>
        /// <param name="sampleRate1">原采样率</param>
        /// <param name="sampleRate2">修改后采样率</param>
        /// <param name="step"></param>
        /// <returns></returns>
        public static int GetCountBySampleRateAndStep(ulong sampleRateOld,ulong count,int step,out ulong sampleRateNew,out ulong countNew)
        {
            int status = 0 - 1;
            ulong[] sampleRateList = new ulong[4];
            ulong sampleRateMin = 10000;//最小采样率
            countNew = 0;
            sampleRateNew = 0;           
            sampleRateList[0] =  625000000;
            sampleRateList[1] = 1250000000;
            sampleRateList[2] = 2500000000;
            sampleRateList[3] = 5000000000;

            if (sampleRateOld<10000)
            {
                //情形1：采样率小于最小采样率
                countNew = (sampleRateMin / sampleRateOld + 1) * (ulong)count;
                countNew = (ulong)(step * ParseData.ReturnNeastD((double)countNew / step));//新数据量
                sampleRateNew = (ulong)((double)countNew / (double)count * sampleRateOld);//新采样率
                status = 1;
            }
            else if (sampleRateOld <= 312500000)
            {
                //情形2：采样率处于10kSa/s-312.5MSa/s
                countNew = (ulong)(step * ParseData.ReturnNeastD((double)count / step));
                sampleRateNew = (ulong)(((double)countNew / count) * sampleRateOld);
                if(sampleRateNew>312500000)
                {
                    //修改后的采样率>312.5MSa/s，则通过抽值的方式使得新采样率位于区间[10000,312500000]
                    countNew = (ulong)step * (countNew / (ulong)step);
                    sampleRateNew = (ulong)(((double)countNew / count) * sampleRateOld);
                }
                else if(sampleRateNew<10000)
                {
                    //修改后的采样率<10kSa/s，则通过插值的方式使得新采样率位于区间[10000,312500000]
                    countNew = (ulong)step * (countNew / (ulong)step + 1);
                    sampleRateNew = (ulong)(((double)countNew / count) * sampleRateOld);
                }
                status = 2;
            }
            else if (sampleRateOld<=5000000000)
            {
                //情形3：采样率312.5MSa/s-5GSa/s
                for (int i = 0; i < sampleRateList.Length; i++)
                {
                    if (sampleRateOld <= sampleRateList[i])
                    {
                        sampleRateNew = sampleRateList[i];//新采样率
                        break;
                    }
                }
                //根据采样率调整数据量
                countNew = (ulong)((double)sampleRateNew / (double)sampleRateOld * count);
                //根据数据量"粒度"再次调整数据量
                countNew = (ulong)step * (ulong)ParseData.ReturnNeastD((double)countNew / step);
                status = 3;
            }
            else
            {
                //情形4：采样率>5GSa/s
                countNew = (ulong)(step * ParseData.ReturnNeastD((double)count / step));//新数据量
                sampleRateNew = 5000000000;//新采样率
                status = 4;
            }
            return status;
        }
        
        /// <summary>
        /// 消顶消谷
        /// </summary>
        /// <param name="dData"></param>
        /// <param name="minV"></param>
        /// <param name="maxV"></param>
        /// <returns></returns>
        public static bool CutDoubleArray(double [] dData,double minV,double maxV)
        {
            if (dData == null)
                return false;
            for (int i = 0; i < dData.Length;i++)
            {
                if(dData[i] < minV)
                    dData[i] = minV;
                if (dData[i] > maxV)
                    dData[i] = maxV;
            }
            return true;
        }



        /// <summary>
        /// 一次性读取Pat数据文件
        /// </summary>
        /// <param name="patData"></param>
        /// <returns></returns>
        public static bool ReadPatFile(string patFilePath,out double [] patData)
        {
            int index_first = 0;//第1个'#'位置
            int Mlen = 0;
            int fileCount = 0;
            byte[] ReadBuf = new byte[2];
            UInt16 Data = 0;
            fileCount = (int)ParseData.GetPatFileDataBytes(patFilePath, out index_first, out Mlen)/2;
            patData = new double[fileCount];
           
            using (FileStream origStreamRead = new FileStream(patFilePath, FileMode.Open, FileAccess.Read))
            {              
                 using (BinaryReader sr = new BinaryReader(origStreamRead))
                 {                    
                       int i = 0;
                       //移动到Pat波形数据首位置
                        sr.BaseStream.Seek(index_first + Mlen + 2, SeekOrigin.Begin);                         
                       //开始读波形数据
                        while (sr.BaseStream.Position < sr.BaseStream.Length && i < fileCount)
                       {
                           //读取2个字节（即1个波形数据点）
                           ReadBuf = sr.ReadBytes(2);
                           Data = BitConverter.ToUInt16(ReadBuf, 0);          //适合pat文件中的数据
                           Data = ParseData.ConvertPatUInt16toAwfUInt16(Data);//适合awf文件中的数据 0x0000-0xFFFF
                           patData[i] = Data;//存储pat文件中的数据
                           i++;
                       }                       
                 sr.Close();                 
                }
                origStreamRead.Close();
            }
            return true;
        }

        /// <summary>
        /// 将dataIn中数据通过插值或抽值的方式获取n个数据
        /// </summary>
        /// <param name="dataIn">输入数组</param>
        /// <param name="dataOut">输出</param>
        /// <returns></returns>
        public static double [] LinearInterpFunction(double[] dataIn,int n)
        {
            double []dataOut=new double[n];
            int len1 = dataIn.Length;
            int len2 = dataOut.Length;
            int i = 0;
            int index = 0;
            double step = 0.0;
            if (dataIn == null || dataOut == null)
            {
                dataOut = null;
                return null;
            }
            if (len1 == len2)
            {
                for (i = 0; i < len2; i++)
                {
                    dataOut[i] = dataIn[i];
                }
            }
            else
            {
                if (len1 < len2)
                {
                    step = (double)(len1 - 1.0) / (double)(len2 - 1.0);//step>1.0时抽值，step<1.0时插值
                }
                else
                {
                    step = (double)(len1 - 0.0) / (double)(len2 - 0.0);//step>1.0时抽值，step<1.0时插值
                }
                for (i = 0; i < len2; i++)
                {
                    index = ParseData.MaxMin(ParseData.ReturnNeastD(i * step), 0, len1 - 2);
                    dataOut[i] = ParseData.GetLinValue((double)index, dataIn[index], (double)(index + 1.0), dataIn[index + 1], (double)(i * step));
                }
            }
            return dataOut;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="x"></param>
        /// <param name="y"></param>
        /// <returns></returns>
        public static long Min(long x, long y)
        {
            if (x > y)
            {
                return y;
            }
            else
            {
                return x;
            }       
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="x"></param>
        /// <param name="y"></param>
        /// <returns></returns>
        public static long Max(long x, long y)
        {
            if (x > y)
            {
                return x;
            }
            else
            {
                return y;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="I"></param>
        /// <param name="x"></param>
        /// <param name="y"></param>
        /// <returns></returns>
        public static int MaxMin(int I, int x, int y)
        {
            if (x > y)
            {
                int t = x;
                x = y;
                y = t;
            }
            if (I < x) I = x;
            if (I > y) I = y;
            return I;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="D"></param>
        /// <param name="x"></param>
        /// <param name="y"></param>
        /// <returns></returns>
        public static double MaxMin(double D, double x, double y)
        {
            if (x > y)
            {
                double t = x;
                x = y;
                y = t;
            }
            if (D < x) D = x;
            if (D > y) D = y;
            return D;
        }
        /// <summary>
        /// 返回最近整数
        /// </summary>
        /// <param name="d"></param>
        /// <returns></returns>
        public static int ReturnNeastD(double d)
        {
            int I = 0;
            if (Math.Abs(d) < 0.00000001) return I;
            if (d > 0.0)
            {
                if (d - (int)d > 0.50000) I = (int)d + 1;
                else I = (int)d;
            }
            else
            {
                if ((int)d - d < 0.50000) I = (int)d;
                else I = (int)d - 1;
            }
            return I;
        }

        /// <summary>
        /// 求f1和f2之间的线性值
        /// </summary>
        /// <param name="f1"></param>
        /// <param name="am1"></param>
        /// <param name="f2"></param>
        /// <param name="am2"></param>
        /// <param name="f"></param>
        /// <returns></returns>
        public static double GetLinValue(double f1, double am1, double f2, double am2, double f)
        {
            //f1和f2不要相等
            double am;
            if (Math.Abs(f2 - f1) < 0.000001)
                return am1;
            am = (f - f1) * (am2 - am1) / (f2 - f1) + am1;
            return am;
        }

        /// <summary>
        /// 将Pat文件中的数据转换为Awf文件中的数据，保留Mark信息
        /// 因泰克Pat文件有效数据位宽为10位，所以需要重新量化
        /// </summary>
        /// <param name="D">[14]为Mark2信息，[13]为Mark1信息，[9:0]为有效数据，0对应最小值,0x3FF对应最大值</param>
        /// <returns></returns>
        public static UInt16 ConvertPatUInt16toAwfUInt16(UInt16 D)
        {
            //Pat文件中，数值范围 0-1023
            //awf文件中，数值范围 （0-16383)*4，最低两位为Mark2和Mark1,复用
            UInt16 D16 = 0;
            D &= 0x3FF;//取裸数据
            double d = ParseData.Quantify(D,0.0,1022.0,0.0,16383.0);//计算量化数据
            D16 = (UInt16)(d*4) ;//左移2位，给Mark2和Mark1“腾空间”
            D16 &= 0xFFFC;//初始化Mark2、Mark1为0
            if((D & 0x4000)==0x4000)
            {
                //Mark2为1
                D16 |= 0x2;
            }
            if ((D & 0x2000) == 0x2000)
            {
                //Mark1为1
                D16 |= 0x1;
            }       
            return D16;
        }

        /// <summary>
        /// 等效量化
        /// </summary>
        /// <param name="inputV">输入</param>
        /// <param name="intputVmin"></param>
        /// <param name="inputVmax"></param>
        /// <param name="outputVmin"></param>
        /// <param name="outPutVmax"></param>
        /// <returns></returns>
        public static double Quantify(double inputV,double intputVmin,double inputVmax,double outputVmin,double outPutVmax)
        {
            double d = 0;
            d=outputVmin +(outPutVmax - outputVmin)*(inputV-intputVmin)/(inputVmax -intputVmin);
            if (d < outputVmin) d = outputVmin;
            if (d > outPutVmax) d = outPutVmax;
            if (Math.Abs(d )< 1e-4)
            {
                d = 0;
            }
            return d;
        }


       /// <summary>
        /// 获取Pat文件裸数据长度，单位为字节
       /// </summary>
       /// <param name="filename"></param>
       /// <param name="firstpPos"></param>
       /// <param name="dataWidth"></param>
       /// <returns></returns>
        public static ulong GetPatFileDataBytes(string filename,out int firstpPos,out int dataWidth)
        {
            firstpPos=0;
            dataWidth=0;
            int index_first = 0;//第1个'#'位置
            int numBytes = 0; //波形数据部分长度（字节数）           
            int Mlen = 0;
            byte[] header = new byte[32];//pat文件头信息，不会超过32个字节  
            if (!File.Exists(filename))
            {
                return 0;
            }                       
            using (FileStream origStream = new FileStream(filename, FileMode.Open, FileAccess.Read))
            {
                using (BinaryReader sr = new BinaryReader(origStream))
                {
                    header = sr.ReadBytes(header.Length);
                }
                origStream.Close();
            }
            string strMsg = System.Text.Encoding.ASCII.GetString(header);//pat文件头信息
            if (strMsg.Contains("MAGIC"))
            {
                //找到第1个‘#’的位置
                index_first = strMsg.IndexOf("#");
                if (-1 == index_first)
                {
                    //文件头没有找到'#'，直接返回（任务文件已被破坏）
                    return 0;
                }
                firstpPos = index_first;
                //检查有效个数位宽是否是合法的数字ASCII，位宽不超过9位，即最大999999999
                if (strMsg[index_first + 1] < '0' || strMsg[index_first + 1] > '9')
                {
                    return 0;
                }
                else
                {
                    Mlen = strMsg[index_first + 1] - 48;//裸数据有效字节数位宽（1个位宽为4bits）,字符'0'对应值48
                    dataWidth = Mlen;
                }
                //获取pat文件有效字节数                                
                if (!int.TryParse(strMsg.Substring(index_first + 2, Mlen), out numBytes))
                {
                    return 0;
                }            
            }
            return (ulong)numBytes;
        }

        /// <summary>
        /// 获取.Pat格式文件的采样率信息
        /// </summary>
        /// <param name="filename"></param>
        /// <returns></returns>
        public static ulong GetPatFileSampleRate(string filename)
        {
            ulong sampleRate = 0;
            if (!File.Exists(filename))
            {
                return 0;
            }
            int index_first = 0;//第1个'CLOCK'位置       
            byte[] ender = new byte[32];//pat文件头信息，不会超过30个字节                 
            using (FileStream origStream = new FileStream(filename, FileMode.Open, FileAccess.Read))
            {
                origStream.Seek(origStream.Length-32,SeekOrigin.Begin);
                using (BinaryReader sr = new BinaryReader(origStream))
                {
                    ender = sr.ReadBytes(ender.Length);
                }
                origStream.Close();
            }
            string strMsg = System.Text.Encoding.ASCII.GetString(ender);//pat文件头信息       
            if (strMsg.Contains("CLOCK"))
            {
                //找到第1个‘CLOCK’的位置
                index_first = strMsg.IndexOf("CLOCK");
                if (-1 == index_first)
                {                 
                    return 0;
                }else
                {
                    strMsg=strMsg.Substring(index_first + 5);
                    double d=0;
                    if (!double.TryParse(strMsg, out d))
                    {
                        sampleRate=0;
                    }else
                    {
                        sampleRate=(ulong)d;
                    }                  
                }
            }
            return sampleRate;
        }



        /// <summary>
        /// 导出文件时弹出对话框获取用户保存路径
        /// </summary>
        public static string ExportDialog()
        {
            SaveFileDialog sf = new SaveFileDialog();
            sf.Title = "导出文件";
            sf.DefaultExt = "awf";
            sf.Filter = "awf files(*.awf)|*.awf|csv files(*.csv)|*.csv|txt files(*.txt)|*.txt";
            //sf.InitialDirectory = System.AppDomain.CurrentDomain.SetupInformation.ApplicationBase;
            sf.InitialDirectory = "c:\\";
            sf.FilterIndex = 1;
            sf.RestoreDirectory = true;

            if ((bool)sf.ShowDialog())
            {
                return sf.FileName;
            }
            else
            {
                return "";
            }
        }

        /// <summary>
        /// 导入文件时弹出对话框
        /// </summary>
        public static string ImportDialog()
        { 
            OpenFileDialog sf = new OpenFileDialog();
            sf.Title = "Open  Files";
            sf.DefaultExt = "awf";
            sf.Filter = "awf files(*.awf)|*.awf|pat files(*.pat)|*.pat|csv files(*.csv)|*.csv|txt files(*.txt)|*.txt";
            //sf.InitialDirectory = System.AppDomain.CurrentDomain.SetupInformation.ApplicationBase;
            sf.InitialDirectory = "c:\\";
            sf.FilterIndex = 1;
            sf.RestoreDirectory = true;

            if ((bool)sf.ShowDialog())
            {
                return sf.FileName;
            }
            else
            {
                return "";
            }       
        }

        /// <summary>
        /// 读入 .awf 文件
        /// </summary>
        public static string ImportDialogOnlyData(string subDir = "")
        {
            // 读取 .awf   
            OpenFileDialog sf = new OpenFileDialog();
            sf.Title = "打开 awf 文件";
            sf.DefaultExt = "awf";
            sf.Filter = "files(*.awf)|*.awf";
            sf.InitialDirectory = System.AppDomain.CurrentDomain.SetupInformation.ApplicationBase + subDir;
            sf.FilterIndex = 1;
            sf.RestoreDirectory = true;

            if ((bool)sf.ShowDialog())
            {

            }
            return sf.FileName;
        }
     
        /// <summary>
        /// 添加数据到awf文件
        /// </summary>
        public static void SaveDataDialogAppend(string filename, int[] data)
        {
            using (FileStream fs = new FileStream(filename, FileMode.Append))
            {
                using (BinaryWriter sw = new BinaryWriter(fs))
                {
                    for (int i = 0; i < data.Length; i++)
                    {
                        byte[] byteData = BitConverter.GetBytes(data[i]);
                        UInt16 tBuffer = (UInt16)(data[i] & 0xFFFF);
                        sw.Write(tBuffer);//写入1个点数据
                    }
                    sw.Close();
                }
                fs.Close();
            }
        }

      
    

        /// <summary>
        /// 纯二进制文件拷贝
        /// </summary>
        /// <param name="srcPath"></param>
        /// <param name="desPath"></param>
        /// <returns></returns>
        public static bool CopyFile(string srcPath, string desPath)
        {
            try
            {
                System.IO.File.Copy(srcPath, desPath, true);
                return true;
            }
            catch (Exception e)
            {
                return false;
            }
        }

        /// <summary>
        /// 从绝对路径中获取末尾的文件名
        /// </summary>
        /// <param name="srcPath"></param>
        /// <returns></returns>
        public static string GetNameFromPath(string srcPath)
        {
            string filename = "";

            if (srcPath == null || srcPath.Length <= 0)
            {
                return "";
            }

            int last_index = srcPath.LastIndexOf("/");

            if (last_index < 0)
            {
                last_index = srcPath.LastIndexOf("\\");
            }

            filename = srcPath.Substring(last_index + 1, srcPath.Length - last_index - 1);

            if (filename.Contains("."))
            {
                return filename;
            }

            return "";
        }

        public static bool IsExistsDir(string srcPath)
        {
            DirectoryInfo mypath = new DirectoryInfo(srcPath);
            if (mypath.Exists)
            {
                return true;
            }
            return false;
        }

        public static void ReNameFile(string srcPath, string desPath)
        {
            try
            {
                System.IO.File.Copy(srcPath, desPath, true);
                System.IO.File.Delete(srcPath);
            }
            catch (Exception e)
            {

            }
        }

       

        /// <summary>
        /// 创建新数据文件
        /// </summary>
        public static void CreateFile(string filename)
        {
            using (FileStream fs = new FileStream(filename, FileMode.Create))
            {
                using (BinaryWriter sw = new BinaryWriter(fs))
                {
                    fs.Close();
                }
            }
        }
        /// <summary>
        /// 创建新文件并写入数据。如果filename文件已存在，则删除已有文件并创建新文件
        /// </summary>
        /// <param name="filename">文件存储格式</param>
        /// <param name="data"></param>
        /// <param name="_fileFormat"></param>
        /// <param name="_sampleRateHz"></param>
        /// <param name="bkWorker"></param>
        public static void SaveToNewAwfile(string filename, int[] data, UInt16 _fileFormat = 0, UInt16 _genSource=0, UInt16 _playflag=0, UInt64 _sampleRateHz = 5000000000, BackgroundWorker bkWorker = null)
        {
            if (System.IO.File.Exists(filename))
            {
                ParseData.DeleteFile(filename);
            }
            //创建新文件
            ParseData.CreateFile(filename);
            AWFileInfo fileHead = ParseData.GetAwfileInfoInit();//文件头信息结构初始化
            fileHead._lengthBytes = (ulong)data.Length * fileHead._dataWidth;//裸数据长度
            fileHead._fileFormat = _fileFormat;//文件格式识别码
            fileHead._genSource=_genSource;//文件来源
            fileHead._playflag=_playflag;//播放方式
            fileHead._sampleRate = _sampleRateHz;//采样率
            //写文件头信息
            ModifyAWFileHeadInfo(fileHead, filename);
            UInt16 tBuffer = 0;
            //写数据
            using (FileStream origStreamWriter = new FileStream(filename, FileMode.Open, FileAccess.Write))
            {
                using (BinaryWriter sw = new BinaryWriter(origStreamWriter))
                {
                    //Seek到裸位置起始位置
                    sw.Seek(FrameHeaderSize, SeekOrigin.Begin);
                    for (int i = 0; i < data.Length; i++)
                    {
                        tBuffer = (UInt16)(data[i] & 0xFFFF);
                        sw.Write(tBuffer);//写入1个点数据
                    }
                    sw.Close();
                }
                origStreamWriter.Close();
            }
            return;
        }
        /// <summary>
        /// 修改awf文件，若文件不存在，则创建新文件
        /// 若添加的数据（后部分）超过filename裸数据区域，则将超出文件区域的数据写入末尾（此时修改了文件长度，所以需要更新文件头信息）
        /// 若startPosition即要添加的数据起始位置就超出filename裸数据区域（甚至中间留有空白），则采用将新数据紧贴filename末尾写入
        /// </summary>
        /// <param name="filename">文件名绝对路径</param>
        /// <param name="data">修改后的数据</param>
        /// <param name="startPosition">修改起始位置</param>
        /// <param name="bkWorker"></param>
        public static void SaveToAwfileEx(string filename, int[] data, int startPosition, BackgroundWorker bkWorker = null)
        {
            bool _isNewFile = false;
            if(startPosition<0)
            {
                return;
            }
            //文件头信息结构初始化
            AWFileInfo fileHead = ParseData.GetAwfileInfoInit();     
            if (!System.IO.File.Exists(filename) )
            {
                //创建新文件，文件格式为com普通文件
                using (FileStream origStreamWriter = new FileStream(filename, FileMode.Create, FileAccess.Write))
                {
                    using (BinaryWriter sw = new BinaryWriter(origStreamWriter))
                    {                     
                        sw.Close();
                    }
                    origStreamWriter.Close();
                }
                //写文件头信息
                ModifyAWFileHeadInfo(fileHead, filename);
                _isNewFile = true;
            }
            //检查文件是否是标准awf格式的
            if(!CheckawfFileFormat(filename))
            {
                return;
            }            
            //获取文件头信息
            GetawfFileHeadInfo(filename, out fileHead);
            UInt16 tBuffer = 0;
            ulong firstIndex = (ulong)(startPosition * fileHead._dataWidth + FrameHeaderSize);
            if (firstIndex >= fileHead._lengthBytes)
            {
                //全部超出裸数据区域
                SaveDataDialogAppend(filename, data);
                fileHead._lengthBytes += (ulong)data.Length * fileHead._dataWidth;
                //修改文件头信息
                ModifyAWFileHeadInfo(fileHead, filename);
            }
            else 
            {          
                using (FileStream origStreamWriter = new FileStream(filename, FileMode.Open, FileAccess.Write))
                {
                    using (BinaryWriter sw = new BinaryWriter(origStreamWriter))
                    {
                        //Seek到startPosition位置
                        sw.Seek(startPosition * fileHead._dataWidth + FrameHeaderSize, SeekOrigin.Begin);
                        for (int i = 0; i < data.Length; i++)
                        {
                            tBuffer = (UInt16)(data[i] & 0xFFFF);
                            sw.Write(tBuffer);//写入1个点数据
                        }
                        sw.Close();
                    }
                    origStreamWriter.Close();
                }
                if (firstIndex + (ulong)(data.Length * fileHead._dataWidth) > ((ulong)FrameHeaderSize + fileHead._lengthBytes))
                {
                    //部分新增数据超出裸数据区域，需要修改文件头信息
                    ulong addBytes = (ulong)data.Length * fileHead._dataWidth - ((ulong)FrameHeaderSize + fileHead._lengthBytes- firstIndex);//文件增加的字节数
                    fileHead._lengthBytes += addBytes;
                    //修改文件头信息
                    ModifyAWFileHeadInfo(fileHead, filename);
                }
            }  
            if( _isNewFile ==true)
            {          
                ModifyAWFileHeadInfo(fileHead, filename); //修改文件头信息
            }

            return;
        }
       
        /// <summary>
        /// 删除文件
        /// </summary>
        /// <param name="filename">被删除文件的绝对路径</param>
        public static void DeleteFile(string filename)
        {
            
            if (System.IO.File.Exists(filename))
            {
                try
                {
                     System.IO.File.Delete(filename);
                }catch
                {

                }               
            }
        }
        /// <summary>
        /// 删除旧文件，创建新文件
        /// </summary>
        /// <param name="filename"></param>
        public static void CreateNewFile(string filename)
        {
            ParseData.DeleteFile(filename);
            FileStream fs = new FileStream(filename, FileMode.Create);
            fs.Close();
            return;
        }

        /// <summary>
        /// 删除文件夹下的全部.awf文件
        /// </summary>
        /// <param name="folderFullName"></param>
        public static void DeleteAllFile(string folderFullName)
        {
            if (folderFullName == "" || folderFullName==null)
                return;
            DirectoryInfo TheFolder = new DirectoryInfo(folderFullName);

            if (!TheFolder.Exists)
            {
                return;
            }

            //遍历文件
            foreach (FileInfo NextFile in TheFolder.GetFiles())
            {
                if (NextFile.Name.Contains(".awf"))
                {
                    ParseData.DeleteFile(folderFullName + NextFile.Name);
                }
            }

        }

        /// <summary>
        /// 删除文件夹下的全部.xmlf文件
        /// </summary>
        /// <param name="folderFullName"></param>
        public static void DeleteAllXMLFile(string folderFullName)
        {
            if (folderFullName == "" || folderFullName == null)
                return;
            DirectoryInfo TheFolder = new DirectoryInfo(folderFullName);

            if (!TheFolder.Exists)
            {
                return;
            }

            //遍历文件
            foreach (FileInfo NextFile in TheFolder.GetFiles())
            {
                if (NextFile.Name.Contains(".xml"))
                {
                    ParseData.DeleteFile(folderFullName + NextFile.Name);
                }
            }

        }

        /// <summary>
        /// 创建destfile文件并将filePath1和filePath2中数据拷贝
        /// 新文件头信息同filePath1，数据量为两者之和
        /// </summary>
        /// <param name="destfile"></param>
        /// <param name="filePath1"></param>
        /// <param name="filePath2"></param>
        public static void CombineToAwfile(string destfile,string filePath1,string filePath2)
        {  
            ParseData.CreateNewFile(destfile);
            AWFileInfo _fileInfo1;
            AWFileInfo _fileInfo2;
            ParseData.GetawfFileHeadInfo(filePath1, out  _fileInfo1);
            ParseData.GetawfFileHeadInfo(filePath2, out  _fileInfo2);
            _fileInfo1._lengthBytes += _fileInfo2._lengthBytes;
            if (ParseData.CopyFileBytes(filePath1,destfile,null))
            {
                //修改头文件信息，这里只修改数据量信息
                ParseData.ModifyAWFileHeadInfo(_fileInfo1, destfile);

                byte[] ReadBuf = new byte[2];
                UInt16 Data = 0;
                using (FileStream origStreamRead = new FileStream(filePath2, FileMode.Open, FileAccess.Read))
                {
                    using (FileStream origStreamWriter = new FileStream(destfile, FileMode.Append))
                    {
                        using (BinaryReader sr = new BinaryReader(origStreamRead))
                        {
                            using (BinaryWriter sw = new BinaryWriter(origStreamWriter))
                            {                            
                                sr.BaseStream.Seek(FrameHeaderSize, SeekOrigin.Begin);
                                while (sr.BaseStream.Position < sr.BaseStream.Length)
                                {
                                    ReadBuf = sr.ReadBytes(2);
                                    Data = BitConverter.ToUInt16(ReadBuf, 0);
                                    sw.Write(BitConverter.GetBytes(Data));
                                }                                                               
                            }
                        }
                    }
                }
            }
            return;
        }



    }
}
