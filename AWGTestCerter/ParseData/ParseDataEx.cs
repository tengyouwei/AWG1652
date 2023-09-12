using System;
using System.ComponentModel;
using System.IO;
using System.Windows;


namespace AWGTestCerter.ParseData
{
    /**************************************************************************************
     * 设计者：滕友伟
     * 功  能：文件类型的相互转换，绘制数据和保存数据的相互转换扩展
     * 时  间：2020年6月11日
     * 修  改：
     * 
     * 
     * ************************************************************************************/
    public static partial class ParseData
    {
        /// <summary>
        /// 获取偏移位置子序列头信息
        /// </summary>
        /// <param name="origStream">.seq文件</param>
        /// <param name="offset">偏移</param>
        /// <param name="_fileInfo">头信息获取</param>
        public static void GetSubseqHeadInfo(FileStream origStream, int offset, out AWFileInfo _fileInfo)
        {
            int HeadlengthBytes = FrameHeaderSize / sizeof(UInt16);
            UInt16[] tempArrayFrameHeader = new UInt16[HeadlengthBytes];
            //移动到子序列头信息位置           
            ParseData.ReadUInt16FromStream(origStream, tempArrayFrameHeader, offset);
            if (tempArrayFrameHeader[1] == 0x3412 &&  tempArrayFrameHeader[0] == 0xBBAA)
            {
                //文件信息获取
                _fileInfo._identfyLow = tempArrayFrameHeader[0];
                _fileInfo._identfyHigh = tempArrayFrameHeader[1];
                _fileInfo._framwareVersion = tempArrayFrameHeader[2];
                _fileInfo._dataWidth = tempArrayFrameHeader[3];

                _fileInfo._lengthBytes = (UInt64)(tempArrayFrameHeader[4]
                                                + tempArrayFrameHeader[5] * Math.Pow(2, 16)
                                                + tempArrayFrameHeader[6] * Math.Pow(2, 32)
                                                + tempArrayFrameHeader[7] * Math.Pow(2, 48));

                _fileInfo._sampleRate = (UInt64)(tempArrayFrameHeader[8]
                                               + tempArrayFrameHeader[9] * Math.Pow(2, 16)
                                               + tempArrayFrameHeader[10] * Math.Pow(2, 32)
                                               + tempArrayFrameHeader[11] * Math.Pow(2, 48));

                _fileInfo._instrument    = tempArrayFrameHeader[12];
                _fileInfo._genSource     = tempArrayFrameHeader[13];//文件数据来源
                _fileInfo._fileFormat    = tempArrayFrameHeader[14];//文件数据存储标记  0x0001表示单路I数据；0x0002表示单路Q数据；0x0003表示IQ数据；0x0004表示RF数据；               
                _fileInfo._defaultInfo3  = tempArrayFrameHeader[15];

                _fileInfo._playflag      = tempArrayFrameHeader[16];//文件播放时处理标记 Real或DUC              
                _fileInfo._defaultInfo4  = tempArrayFrameHeader[17];
                _fileInfo._defaultInfo5  = tempArrayFrameHeader[18];
                _fileInfo._defaultInfo6  = tempArrayFrameHeader[19];

                _fileInfo._defaultInfo7  = tempArrayFrameHeader[20];//本子序列循环播放次数 0-65535
                _fileInfo._defaultInfo8  = tempArrayFrameHeader[21];
                _fileInfo._defaultInfo9  = tempArrayFrameHeader[22];
                _fileInfo._defaultInfo10 = tempArrayFrameHeader[23];
            }
            else
            {
                _fileInfo._identfyHigh =0;
                _fileInfo._identfyLow = 0;
                _fileInfo._sampleRate = 0;
                _fileInfo._lengthBytes = 0;        
                _fileInfo._fileFormat = 0;         
                _fileInfo._framwareVersion = 0;    
                _fileInfo._instrument = 0;    
                _fileInfo._dataWidth = 0;          
                _fileInfo._genSource = 0;
                _fileInfo._playflag = 0;
                _fileInfo._defaultInfo3 = 0;
                _fileInfo._defaultInfo4 = 0;
                _fileInfo._defaultInfo5 = 0;
                _fileInfo._defaultInfo6 = 0;
                _fileInfo._defaultInfo7 = 0;
                _fileInfo._defaultInfo8 = 0;
                _fileInfo._defaultInfo9 = 0;
                _fileInfo._defaultInfo10 = 0;
            }
             return;
        }



        /// <summary>
        /// 取awf文件头信息
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="_fileInfo"></param>
         public static void GetawfFileHeadInfo(string filePath, out AWFileInfo _fileInfo)
         {
             //初始化头信息
             _fileInfo._identfyLow = 0;
             _fileInfo._identfyHigh = 0;
             _fileInfo._framwareVersion = 0;
             _fileInfo._dataWidth = 0;
             _fileInfo._lengthBytes = 0;
             _fileInfo._sampleRate = 0;
             _fileInfo._instrument = 0;
             _fileInfo._fileFormat = 0;   //文件存储格式 ：I路、Q路、IQ交叉、RF

             _fileInfo._genSource = 0;    //文件数据来源 0：Com 通用数据文件
             _fileInfo._playflag = 1;     //播放是否需要硬件上变频 1：Real模式不需要  2：DUC模式需要（对应IQ插件下IQ存储格式的数据文件）
             _fileInfo._defaultInfo3  = 0;
             _fileInfo._defaultInfo4  = 0;
             _fileInfo._defaultInfo5  = 0;
             _fileInfo._defaultInfo6  = 0;
             _fileInfo._defaultInfo7  = 0;//循环播放次数 0-65535
             _fileInfo._defaultInfo8  = 0;
             _fileInfo._defaultInfo9  = 0;
             _fileInfo._defaultInfo10 = 0;

            if (File.Exists(filePath) )
            {
                if (Path.GetExtension(filePath) == ".awf" || Path.GetExtension(filePath) == ".seq")
                {
                    //读取文件
                    using (FileStream origStreamReadRawData = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite, 8192))
                    {
                        int HeadlengthBytes = FrameHeaderSize / sizeof(UInt16);
                        UInt16[] tempArrayFrameHeader = new UInt16[HeadlengthBytes];
                        ParseData.ReadUInt16FromStream(origStreamReadRawData, tempArrayFrameHeader, 0);
                        //文件信息获取
                        _fileInfo._identfyLow = tempArrayFrameHeader[0];
                        _fileInfo._identfyHigh = tempArrayFrameHeader[1];
                        _fileInfo._framwareVersion = tempArrayFrameHeader[2];
                        _fileInfo._dataWidth = tempArrayFrameHeader[3];
                        _fileInfo._lengthBytes = (UInt64)(tempArrayFrameHeader[4]
                                                        + tempArrayFrameHeader[5] * Math.Pow(2, 16)
                                                        + tempArrayFrameHeader[6] * Math.Pow(2, 32)
                                                        + tempArrayFrameHeader[7] * Math.Pow(2, 48));
                        _fileInfo._sampleRate = (UInt64)(tempArrayFrameHeader[8]
                                                       + tempArrayFrameHeader[9] * Math.Pow(2, 16)
                                                       + tempArrayFrameHeader[10] * Math.Pow(2, 32)
                                                       + tempArrayFrameHeader[11] * Math.Pow(2, 48));

                        _fileInfo._instrument   = tempArrayFrameHeader[12];
                        _fileInfo._genSource    = tempArrayFrameHeader[13];//数据文件来源（各功能插件）
                        _fileInfo._fileFormat   = tempArrayFrameHeader[14];//文件数据存储标记  0x0001表示单路I数据；0x0002表示单路Q数据；0x0003表示IQ数据；0x0004表示RF数据；
                        _fileInfo._defaultInfo3 = tempArrayFrameHeader[15];
                        _fileInfo._playflag     = tempArrayFrameHeader[16];//文件处理标记      0x0001：Real（实时播放）；0x0002：Duc（即上变频IQ数据）
                        _fileInfo._defaultInfo4 = tempArrayFrameHeader[17];
                        _fileInfo._defaultInfo5 = tempArrayFrameHeader[18];
                        _fileInfo._defaultInfo6 = tempArrayFrameHeader[19];
                        _fileInfo._defaultInfo7 = tempArrayFrameHeader[20];//本子序列循环播放次数 0-65535
                        _fileInfo._defaultInfo8 = tempArrayFrameHeader[21];
                        _fileInfo._defaultInfo9 = tempArrayFrameHeader[22];
                        _fileInfo._defaultInfo10 = tempArrayFrameHeader[23];
                    }
                }                          
             }
         }


        /// <summary>
        /// 检查.awf文件是否是标准格式
        /// 标准1：采样率不超过本机支持范围
        /// 标准2：波形数据位宽不小于1byte
        /// 标准3：数据部分不为空
        /// 标准4：帧识别码必须为0x3412和0xBBAA
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
         public static bool CheckawfFileFormat(string filePath)
         {
             AWFileInfo _fileInfo=ParseData.GetAwfileInfoInit();//初始化awf文件头指针
             GetawfFileHeadInfo(filePath, out  _fileInfo);
             if (_fileInfo._identfyHigh != 0x3412 || _fileInfo._identfyLow != 0xBBAA)
             {
                 //帧头识别码有误
                 return false;
             }
             if (_fileInfo._dataWidth <1)
             {
                 //数据位宽有误
                 return false;
             }
             if(_fileInfo._sampleRate<999999)
             {
                 //采样率小于10kHz
                 return false;
             }
             if (_fileInfo._sampleRate>312500000)
             {
                 //采样率大于312.5MHz时,采样率只能是如下几个值
                 if(_fileInfo._sampleRate!=625000000 && _fileInfo._sampleRate!=1250000000 && _fileInfo._sampleRate!=2500000000 && _fileInfo._sampleRate!=5000000000)
                 {
                     return false;
                 }
             }
             if (_fileInfo._lengthBytes <= (ulong)FrameHeaderSize)
             {
                 //数据量少于2bytes
                 return false;
             }
             return true;
         }
        
        /// <summary>
        /// 检查采样率是否匹配1652B 10M-312.5MHz，625MHz，1.25GHz，2.5GHz，5GHz
        /// </summary>
        /// <param name="_sampleRate"></param>
        /// <returns></returns>
        public static bool CheckIsEnableSampleRate(ulong _sampleRate)
         {
            bool flag = false;
            if(_sampleRate==625000000|| _sampleRate==1250000000 || _sampleRate==2500000000 || _sampleRate==5000000000)
            {
                flag = true;
            }
            if (_sampleRate >= 10000000 && _sampleRate<=312500000)
            {
                flag = true;
            }
             return flag;
         }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="_fileInfo"></param>
        /// <returns></returns>
         public static bool CheckawfFileFormat(AWFileInfo _fileInfo)
         {          
             if (_fileInfo._identfyHigh != 0x3412 || _fileInfo._identfyLow != 0xBBAA)
             {
                 //帧头识别码有误
                 return false;
             }
             if (_fileInfo._dataWidth < 1)
             {
                 //数据位宽有误
                 return false;
             }
             if (_fileInfo._sampleRate < 10000)
             {
                 //采样率小于10kHz
                 return false;
             }
             if (_fileInfo._sampleRate > 312500000)
             {
                 //采样率大于312.5MHz时,采样率只能是如下几个值
                 if (_fileInfo._sampleRate != 625000000 && _fileInfo._sampleRate != 1250000000 && _fileInfo._sampleRate != 2500000000 && _fileInfo._sampleRate != 5000000000)
                 {
                     return false;
                 }
             }
             if (_fileInfo._lengthBytes <= (ulong)FrameHeaderSize)
             {
                 //数据量少于2bytes
                 return false;
             }
             return true;
         }
        /// <summary>
        /// 读取序列seq文件头信息中的 a.数据量Bytes和 b.采样率信息
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
         public static void GetSeqFileBytesAndSampleHz(string filePath, out ulong totalBytes,out ulong sampleHz)
         {
             totalBytes = 0;
             sampleHz = 0;
             if (Path.GetExtension(filePath) != ".seq")
             {              
                 return;
             }
             //检查文件头，即序列中第1个子序列头信息
             AWFileInfo _firstSeq=new AWFileInfo();
             GetawfFileHeadInfo(filePath, out _firstSeq);
             totalBytes = _firstSeq._lengthBytes;
             sampleHz = _firstSeq._sampleRate;
             return;
         }

        /// <summary>
         /// 读数据文件，跳过文件头信息（仅用于读取文件并下载到DDR存储器中）
        /// </summary>
        /// <param name="origStream"></param>
        /// <param name="dataArray"></param>
        /// <param name="offset">相偏移</param>
        /// <returns></returns>
         public static uint tywReadUInt32ToDDR(FileStream origStream, UInt32[] dataArray, long offset = 0,int validReadBytes=0)
         {
             int byteSize = dataArray.Length * sizeof(UInt32);//读缓存大小,单位为Bytes
             byte[] tempArray = new byte[byteSize];
             int readLength = 0;
             origStream.Seek(offset, SeekOrigin.Begin);//首位置偏移   
             if (validReadBytes < 0 || validReadBytes > byteSize)
                 validReadBytes = 0;  //即读取长度取决于tempArray长度  
                       
             if (validReadBytes==0)
             {
                 readLength = origStream.Read(tempArray, 0, tempArray.Length);
             }else
             {
                 readLength = origStream.Read(tempArray, 0, validReadBytes);
             }
             
             //正常情况下readLength必定为偶数（因为1个波形点占2个字节）,readLength%4只能是0或者2
             if (readLength%4 ==2)
             {
                 //如果读取的数据量为奇数,需要添加最后1个数据（32位）的高位为0x8000（FPGA认为的0）
                 tempArray[readLength]    = 0x00;
                 tempArray[readLength +1] = 0x80;
                 readLength = readLength + 2;//这样就能确保（readLength%4==0，位宽对齐）
             }          
             readLength = readLength / sizeof(UInt32);
             for (int i = 0; i < readLength; i++)
             {
                 dataArray[i] = BitConverter.ToUInt32(tempArray, i * sizeof(UInt32));
             }
             return (uint)readLength;
         }



         private static bool flag = false;
         private static int compareValue = 1;

         /// <summary>
         /// 发100次执行百分比
         /// </summary>
         /// <param name="percent">[0,1]之间执行比例</param>
         /// <param name="_backgroundWorker"></param>
         public static void ReportPercent(double percent, BackgroundWorker _backgroundWorker)
         {
             percent = percent * 100;
             if (percent > compareValue)
             {
                 flag = true;
                 compareValue = compareValue + 1;
                 if (compareValue > 100)
                 {
                     compareValue = 1;
                 }
             }
             if (flag)
             {
                 _backgroundWorker.ReportProgress((int)(percent));
                 flag = false;
             }
             return;
         }

       

        

         /// <summary>
         /// 纯数据文件对拷，不做任何变换
         /// </summary>
         /// <param name="SourcefilePath"></param>
         /// <param name="destinationfilePath"></param>
         /// <param name="_backgroundWorker"></param>
         /// <returns></returns>
         public static bool CopyFileBytes(string SourcefilePath, string destinationfilePath, BackgroundWorker _backgroundWorker)
         {
             int num = 0;
             long offset = 0;
             byte[] arrBytes = null;
             FileStream _readStream = null;
             FileStream _writeStream = null;
             double percent = 0;         
             try
             {
                 _readStream = File.OpenRead(SourcefilePath);
                 if (File.Exists(destinationfilePath))
                 {
                     FileInfo fi = new FileInfo(destinationfilePath);
                     if (fi.Attributes.ToString().IndexOf("ReadOnly") != -1)
                         fi.Attributes = FileAttributes.Normal;
                     File.Delete(destinationfilePath);     //同名文件存在先删除
                 }
                 _writeStream = File.Create(destinationfilePath);
                 arrBytes = new byte[1024];//每次最多拷贝的字节数
                 //初始化百分比显示标记              
                 flag = false;
                 compareValue = 1;
                 percent = 0;//拷贝执行百分比              
                 //循环取数
                 while (offset < _readStream.Length)
                 {
                     num = _readStream.Read(arrBytes, 0, arrBytes.Length);//读数据
                     offset = offset + num;
                     _readStream.Seek(offset, SeekOrigin.Begin);                    
                     _writeStream.Write(arrBytes, 0, num);//写数据
                     if (_backgroundWorker!=null)
                     {
                         percent = (double)offset / (double)_readStream.Length;//计算百分比
                         ParseData.ReportPercent(percent, _backgroundWorker);//显示拷贝进度
                     }              
                 }
                 offset = _readStream.Length;
                 _readStream.Close();
                 _writeStream.Close();
             }
             catch
             {
                 return false;
             }
             _readStream.Close();
             _writeStream.Close();
             return true;
         }



         /// <summary>
         /// 将浮点值写入文件，文件不存在先创建
         /// </summary>
         /// <param name="d"></param>
         /// <param name="fileName"></param>
         public static void WriteDoubleArrayToFile(double[] d, string filename)
         {
             using (FileStream origStreamWriter = new FileStream(filename, FileMode.Append, FileAccess.Write))
             {
                 //新文件则起始位置无效               
                 using (BinaryWriter sw = new BinaryWriter(origStreamWriter))
                 {
                     for (int i = 0; i < d.Length; i++)
                     {
                         sw.Write(BitConverter.GetBytes(d[i]));
                     }
                     sw.Close();
                 }
                 origStreamWriter.Close();
             }
             return;
         }

         /// <summary>
         /// 将Bool值写入文件，文件不存在先创建
         /// </summary>
         /// <param name="d"></param>
         /// <param name="fileName"></param>
         public static void WriteBoolArrayToFile(bool[] d, string filename)
         {
             using (FileStream origStreamWriter = new FileStream(filename, FileMode.Append, FileAccess.Write))
             {
                 //新文件则起始位置无效               
                 using (BinaryWriter sw = new BinaryWriter(origStreamWriter))
                 {
                     for (int i = 0; i < d.Length; i++)
                     {
                         sw.Write(BitConverter.GetBytes(d[i]));
                     }
                     sw.Close();
                 }
                 origStreamWriter.Close();
             }
             return;
         }

         /// <summary>
         /// 读取文件中的Bool数据
         /// </summary>
         /// <param name="d"></param>
         /// <param name="fileName"></param>
         public static void ReadDoubleArrayToFile(out double[] d, string filename)
         {
             d = null;
             using (FileStream origStreamReader = new FileStream(filename, FileMode.Open, FileAccess.Read))
             {
                 //新文件则起始位置无效               
                 using (BinaryReader sr = new BinaryReader(origStreamReader))
                 {
                     d = new double[sr.BaseStream.Length / 8];
                     for (int i = 0; i < d.Length; i++)
                     {
                         d[i] = BitConverter.ToDouble(sr.ReadBytes(8), 0);
                     }
                     sr.Close();
                 }
                 origStreamReader.Close();
             }
             return;
         }




         /// <summary>
         /// 读取Matlab返回的数据
         /// </summary>
         /// <param name="filePath"></param>
         /// <param name="dataDouble"></param>
         static public bool ReadMatlabReturnValue(string filePath, out double[] dataDouble)
         {
             dataDouble = null;
             int byteSize = 0;
             byte[] tempArray = null;
             UInt64[] dataArray = null;
             int readLength = 0;
             int count = 0;
             try
             {
                 using (FileStream origStream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                 {
                     count = (int)(origStream.Length / sizeof(UInt64));//文件中UInt64数据个数
                     byteSize = sizeof(UInt64);
                     tempArray = new byte[count * byteSize];
                     dataArray = new UInt64[count];
                     dataDouble = new double[count];
                     readLength = origStream.Read(tempArray, 0, tempArray.Length);
                     for (int i = 0; i < count; i++)
                     {
                         dataArray[i] = BitConverter.ToUInt64(tempArray, i * sizeof(UInt64));
                     }
                     origStream.Close();
                     for (int j = 0; j < count; j++)
                     {
                         dataDouble[j] = (double)dataArray[j];
                     }
                 }
             }
             catch(Exception e)
             {
                 MessageBox.Show(e.ToString());
                 return false;
             }
             
             return true;
         }

         /// <summary>
         /// 读Matlab执行错误码
         /// </summary>
         /// <param name="filePath"></param>
         /// <param name="dataUint64"></param>
         /// <returns></returns>
         static public bool ReadMatlabReturnErrorCode(string filePath, out UInt64 dataUint64)
         {
             dataUint64 = 0xFFFFFFFF;
             int byteSize = 0;
             byte[] tempArray = null;
             UInt64[] dataArray = null;
             int readLength = 0;
             int count = 0;
             try
             {
                 using (FileStream origStream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                 {
                     count = (int)(origStream.Length / sizeof(UInt64));//文件中UInt64数据个数
                     byteSize = sizeof(UInt64);
                     tempArray = new byte[count * byteSize];
                     dataArray = new UInt64[count];
                     readLength = origStream.Read(tempArray, 0, tempArray.Length);
                     for (int i = 0; i < count; i++)
                     {
                         dataArray[i] = BitConverter.ToUInt64(tempArray, i * sizeof(UInt64));
                     }
                     origStream.Close();
                     if (dataArray.Length>=1)
                     {
                         dataUint64 = dataArray[0];
                     }                  
                 }
             }
             catch (Exception e)
             {
                 MessageBox.Show(e.ToString());
                 return false;
             }
             return true;
         }


         /// <summary>
         /// 读取Map调制星座图数据
         /// </summary>
         /// <param name="filePath"></param>
         /// <param name="dataDouble"></param>
         static public bool ReadStarScatterDoubleValue(string filePath, out double[] I, out double[] Q, out double[] Mag, out double[]Phase)
         {
             int i=0;
             int byteSize = 0;
             byte[] tempArray = null;
             int readLength = 0;
             int count = 0;
             I = null;
             Q = null;
             Mag = null;
             Phase = null;
             try
             {
                 using (FileStream origStream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                 {
                     count = (int)(origStream.Length / sizeof(double));//文件中double数据个数
                     int N = count / 4;
                     if (count % 4 != 0)
                     {
                         MessageBox.Show("星座图中数据量不能整除3!");
                         return false;
                     }
                     byteSize = sizeof(double);
                     tempArray = new byte[count * byteSize];//将全部数据读出                
                     readLength = origStream.Read(tempArray, 0, tempArray.Length);
                     origStream.Close();

                     I = new double[N];
                     Q = new double[N];
                     Mag = new double[N];
                     Phase = new double[N];

                     for (i = 0; i < count; i++)
                     {
                         if (i < N)
                         {
                             I[i] = BitConverter.ToDouble(tempArray, i * sizeof(double));
                         }
                         else if (i < 2 * N)
                         {
                             Q[i - N] = BitConverter.ToDouble(tempArray, i * sizeof(double));
                         }
                         else if (i < 3 * N)
                         {
                             Mag[i - 2 * N] = BitConverter.ToDouble(tempArray, i * sizeof(double));
                         }
                         else
                         {
                             Phase[i - 3 * N] = BitConverter.ToDouble(tempArray, i * sizeof(double));
                         }
                     }
                 }                   
             }
             catch (Exception e)
             {
                 MessageBox.Show(e.ToString());
                 return false;
             }
             return true;
         }



         /// <summary>
         /// 写Map调制星座图数据，供matlab调用
         /// </summary>
         /// <param name="filePath"></param>
         /// <param name="dataDouble"></param>
         static public bool WriteStarPlotDoubleValue(string filePath,  double[] I, double[] Q, double[] Mag,double []Phase)
         {         
             int len = I.Length;   
             //已有同名文件存在先删除
             if (File.Exists(filePath))
             {
                 FileInfo fi = new FileInfo(filePath);
                 if (fi.Attributes.ToString().IndexOf("ReadOnly") != -1)
                     fi.Attributes = FileAttributes.Normal;
                 File.Delete(filePath);    
             }

             //检查长度是否一致
             if (I.Length != Q.Length || I.Length != Mag.Length)
             {
                 return false;
             }
             double[] dataDouble = new double[len*4];
             Array.Copy(I,     0, dataDouble,          0, len);
             Array.Copy(Q,     0, dataDouble,        len, len);
             Array.Copy(Mag,   0, dataDouble,    2 * len, len);
             Array.Copy(Phase, 0, dataDouble,    3 * len, len);
             WriteDoubleArrayToFile(dataDouble, filePath);
             return true;
         }











    }
}
