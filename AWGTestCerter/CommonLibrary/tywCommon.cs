using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AWGTestCerter.Common
{

    /// <summary>
    /// 
    /// </summary>
    public static class TywCommon
    {
        /// <summary>
        /// 幅度校准信息
        /// </summary>
        public static string CALI_CH1_FILE_PATH = "C:\\Cali\\Cali_CH1_Vpp.bin";
        public static string CALI_CH2_FILE_PATH = "C:\\Cali\\Cali_CH2_Vpp.bin";
        public static string CALI_CH3_FILE_PATH = "C:\\Cali\\Cali_CH3_Vpp.bin";
        public static string CALI_CH4_FILE_PATH = "C:\\Cali\\Cali_CH4_Vpp.bin";

        public static string CALI_CH1_HIGH_PATH = "C:\\Cali\\Cali_CH1_HIGH_Vpp.bin";
        public static string CALI_CH2_HIGH_PATH = "C:\\Cali\\Cali_CH2_HIGH_Vpp.bin";
        public static string CALI_CH3_HIGH_PATH = "C:\\Cali\\Cali_CH3_HIGH_Vpp.bin";
        public static string CALI_CH4_HIGH_PATH = "C:\\Cali\\Cali_CH4_HIGH_Vpp.bin";

        public static string CALI_INSERTLOSS_FILE_PATH = "C:\\Cali\\Cali_Atten_DB.bin";

        public static int INSERTLOSS_COLUM = 201;
        public static double CALI_INSERTLOSS_START = 10000000.0;// （起始）频率
        public static double CALI_INSERTLOSS_STOP = 4000000000.0;// （终止）频率

        public static double DAC_FULL_SCALE = 1023.0;
        public static double DAC_ELEC_MIN = 8.0;//DAC输出电流范围8mA~40mA
        public static double DAC_ELEC_MAX = 40.0;//Ioutfs=32mA * (ANA_FULL_SCALE[9:0]/1023) + 8mA DAC输出电流计算公式

        public static int CURVE_ROW_DC = 14;
        public static int CURVE_ROW_DAC = 14;

        public static int CURVE_COLUM_DC_2G = 101; //DC-2GHz校准频点数目
        public static int CURVE_COLUM_DC_4G = 101; //DC-4GHz校准频点数目(DUC,NCO)

        public static int CURVE_ROW_AC = 4; //AC模式校准频响曲线数目
        public static int AC_ATT_0DB = 0; //AC模式校准的衰减器档位
        public static int AC_ATT_10DB = 10;
        public static int AC_ATT_20DB = 20;
        public static int AC_ATT_30DB = 30;

        public static int CURVE_ROW_AMP = 6; //AMPL模式校准频响曲线数目
        public static int AMP_ATT_0DB = 0; //AMPL模式校准的衰减器档位
        public static int AMP_ATT_10DB = 10;
        public static int AMP_ATT_20DB = 20;
        public static int AMP_ATT_30DB = 30;
        public static int AMP_ATT_40DB = 40;
        public static int AMP_ATT_50DB = 50;


        public static double CALI_AC_SWP_START = 15000000.0;// Normal AC模式下校准起始频率,15MHz
        public static double CALI_AC_SWP_STOP = 2000000000.0;// Normal AC模式下校准终止频率,2GHz

        public static double CALI_DC_SWP_START = 10000000.0;// Normal DC模式下校准起始频率
        public static double CALI_DC_SWP_STOP = 2000000000.0;// Normal DC模式下校准终止频率

        public static double CALI_HIGH_AC_SWP_START = 15000000.0;// AC/AMPL模式下校准（起始）频率
        public static double CALI_HIGH_AC_SWP_STOP = 4000000000.0;// AC/AMPL模式下校准（终止）频率

        public static double CALI_HIGH_DC_SWP_START = 10000000.0;// DC/DAC模式下校准（起始）频率
        public static double CALI_HIGH_DC_SWP_STOP = 4000000000.0;// DC/DAC模式下校准（终止）频率



        public static int OHM_50 = 0; //50欧姆阻抗
        public static int OHM_1M = 1; //1M欧姆阻抗

        public static int SPECT_TYPE_E444X = 0;//幅度校准仪器
        public static int SPECT_TYPE_AV4051X = 1;//幅度校准仪器
        public static int SPECT_TYPE_FSW50 = 2;//幅度校准仪器
        public static int SPECT_TYPE_FSMR50 = 3;//幅度校准仪器
        public static int SPECT_TYPE_MSO254A = 4;//幅度校准仪器
        public static int SPECT_TYPE_MXR604A = 5;//幅度校准仪器

        public static int VOFT_TYPE_MDO4000 = 6;//偏置校准仪器
        public static int VOFT_TYPE_34470A = 7;//偏置校准仪器
        public static int SPECT_TYPE_NRX = 8;//幅度校准仪器（功率计）
        public static int SPECT_TYPE_2438D = 9;//幅度校准仪器（功率计）

        public static int COUPLETYPE_50 = 0;// 50欧姆接地
        public static int COUPLETYPE_DC = 1;// 0.1Vpp-0.3Vpp-1.0Vpp  DC耦合
        public static int COUPLETYPE_AC = 2;// 0.2Vpp-0.6Vpp-2.0Vpp  AC耦合
        public static int COUPLETYPE_DAC = 3;
        public static int COUPLETYPE_ACAMP = 4;

        public static int WAVE_TYPE_SIN = 0;
        public static int WAVE_TYPE_EXP = 1;
        public static int WAVE_TYPE_TRA = 2;
        public static int WAVE_TYPE_TRU = 3;
        public static int WAVE_TYPE_DC = 4;
        public static int WAVE_TYPE_NOS = 5;
        public static int WAVE_TYPE_SWP = 6;
        public static int WAVE_TYPE_PULSE = 7;
        public static int WAVE_TYPE_COMPULSE = 8;

        //---------------------------对外开放(FPGA工作模式)
        public static int FPGA_DUC_MODE_FUNCTION_SIN_HIGH = 0;//DDFS 正弦波2GHz以上
        public static int FPGA_DUC_MODE_FUNCTION_NORMAL = 1;//DDFS 正弦波2GHz一下以及其它函数波形

        public static int FPGA_DUC_MODE_DDR_REAL_5000MS = 2;//DDWS Real 5GHz采样率
        public static int FPGA_DUC_MODE_DDR_REAL_2500MS = 3;//DDWS Real 2.5GHz采样率
        public static int FPGA_DUC_MODE_DDR_REAL_1250MS = 4;//DDWS Real 1.25GHz采样率
        public static int FPGA_DUC_MODE_DDR_REAL_625MS = 5;//DDWS Real 0.625GHz采样率
        public static int FPGA_DUC_MODE_DDR_REAL_BELOW312MS = 10;//DDWS Real 0.3125GHz以下采样率

        public static int FPGA_DUC_MODE_DDR_IQ_2500MS = 6;//DDWS DUC 2.5GHz采样率
        public static int FPGA_DUC_MODE_DDR_IQ_1250MS = 7;//DDWS DUC 1.25GHz采样率
        public static int FPGA_DUC_MODE_DDR_IQ_625MS = 8;//DDWS DUC 0.625GHz采样率
        public static int FPGA_DUC_MODE_DDR_IQ_312MS = 9;//DDWS DUC 0.3125GHz采样率


        static double[] ElECVALUE_DC = new double[14]
        {
           1023,884,760,650,551,463,385,315,253,198,149,105,65,31
        };

        static double[] ElECVALUE_DAC = new double[14]
        {
           1023,884,760,650,551,463,385,315,253,198,149,105,65,31
        };

        public static double ElECVALUE_AC = 760;
        /// <summary>
        /// 偏置校准信息
        /// </summary>
        /// 
        public static string VOFT_CALI_CH1_FILE_PATH = "C:\\Cali\\CaliOft_CH1.bin";
        public static string VOFT_CALI_CH2_FILE_PATH = "C:\\Cali\\CaliOft_CH2.bin";
        public static string VOFT_CALI_CH3_FILE_PATH = "C:\\Cali\\CaliOft_CH3.bin";
        public static string VOFT_CALI_CH4_FILE_PATH = "C:\\Cali\\CaliOft_CH4.bin";

        public static int DCVOFT_COUNT = 49;
        public static int DACVOFT_COUNT = 25;
        public static int ACVOFT_COUNT = 99;


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
        /// 求数组中的最大值
        /// </summary>
        /// <param name="pValue"></param>
        /// <returns></returns>
        public static bool GetMaxValue(double[] pValue, out double maxValue)
        {
            int i;
            maxValue = 0;
            if (pValue == null || pValue.Length < 1)
            {
                return false;
            }
            else if (pValue.Length == 1)
            {
                maxValue = pValue[0];
            }
            else
            {
                maxValue = pValue[0];
                for (i = 1; i < pValue.Length; i++)
                {
                    if (pValue[i] > maxValue)
                    {
                        maxValue = pValue[i];
                    }
                }
            }
            return true;
        }


        /// <summary>
        /// 求数组中的最小值
        /// </summary>
        /// <param name="pValue"></param>
        /// <returns></returns>
        public static bool GetMinValue(double[] pValue, out double minValue)
        {
            int i;
            minValue = 0;
            if (pValue == null || pValue.Length < 1)
            {
                return false;
            }
            else if (pValue.Length == 1)
            {
                minValue = pValue[0];
            }
            else
            {
                minValue = pValue[0];
                for (i = 1; i < pValue.Length; i++)
                {
                    if (pValue[i] < minValue)
                    {
                        minValue = pValue[i];
                    }
                }
            }
            return true;
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
        /// 差值或抽值
        /// </summary>
        /// <param name="pValue1"></param>
        /// <param name="len1"></param>
        /// <param name="pValue2"></param>
        /// <param name="len2"></param>
        public static void GetInterruptOrSampleValue(double[] pValue1, double[] pValue2)
        {
            //pValue1,len 数据源
            //pValue2抽样或者插值得到的数据
            int len1 = pValue1.Length;
            int len2 = pValue2.Length;
            int i = 0;
            int index = 0;
            double step = 0.0;
            if (pValue1 == null || pValue2 == null)
            {
                pValue2 = null;
                return;
            }
            if (len1 == len2)
            {
                for (i = 0; i < len2; i++)
                {
                    pValue2[i] = pValue1[i];
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
                    index = MaxMin(ReturnNeastD(i * step), 0, len1 - 2);
                    pValue2[i] = GetLinValue((double)index, pValue1[index], (double)(index + 1.0), pValue1[index + 1], (double)(i * step));
                }
            }
            return;
        }

        /// <summary>
        /// 将回读的频响曲线字符串数据转换为double型数组,最多检出5000个double数据
        /// </summary>
        /// <param name="str"></param>
        /// <param name="rValue"></param>
        /// <returns>true检出正常数据</returns>
        public static bool LongCharsToValues(string str, out double[] rValue)
        {
            int count = 0;
            int maxCount = 5000;//str中最多检索5000个以","为分割的数据
            List<double> realValue = new List<double>();
            int index = 0;
            double d = 0;
            string strValue = "";
            while (count < maxCount)
            {
                index = str.IndexOf(",");
                if (-1 == index)
                {
                    if (double.TryParse(str, out d))
                    {
                        realValue.Add(d);
                        break;
                    }
                    else
                    {
                        break;
                    }
                }
                if (index < str.Length)
                {
                    strValue = str.Substring(0, index);
                    str = str.Substring(index + 1);
                    if (double.TryParse(strValue, out d))
                    {
                        realValue.Add(d);
                    }
                }
                count++;
            }
            if (realValue.Count < 1)
            {
                rValue = null;
                return false;
            }
            else
            {
                rValue = new double[realValue.Count];
                for (int i = 0; i < realValue.Count; i++)
                {
                    rValue[i] = realValue[i];
                }
            }
            return true;
        }

        /// <summary>
        /// 求中间线性值
        /// </summary>
        /// <param name="slot">取值范围[first,last]</param>
        /// <param name="first"></param>
        /// <param name="last"></param>
        /// <param name="pD"></param>
        /// <returns></returns>
        public static bool GetLineSlotValue(double slot, double first, double last, double[] pD, out double d)
        {
            //first->*pD,last->*(pD+len-1)
            int len = pD.Length;
            int i, index;
            double step;
            double[] pX = new double[pD.Length];
            d = 0.0;
            if (pD == null)
            {
                return false;
            }

            //确保升序
            if (first > last)
            {
                d = first;
                first = last;
                last = d;
                for (i = 0; i < len / 2; i++)
                {
                    d = pD[i];
                    pD[i] = pD[len - 1 - i];
                    pD[len - 1 - i] = d;
                }
            }
            //等间隔插入搜索点
            step = (last - first) / (len - 1);
            for (i = 0; i < len; i++)
            {
                pX[i] = first + i * step;
            }
            if (pX[len - 1] > last)
            {
                pX[len - 1] = last;
            }
            //查找slot位置
            if (slot < first + 0.0000001)
            {
                d = GetLinValue(pX[0], pD[0], pX[1], pD[1], slot);
            }
            else if (slot > last - 0.1)//近似于把slot==last归结于该分支	
            {
                d = GetLinValue(pX[len - 2], pD[len - 2], pX[len - 1], pD[len - 1], slot);
            }
            else
            {
                index = 0;
                for (i = 1; i < len; i++)
                {
                    if (slot < pX[i])
                    {
                        index = i;
                        break;
                    }
                }
                if (i == len && index == 0)
                {
                    //slot 超出first-last范围,用最后1个区间计算
                    index = len - 1;
                }
                index = MaxMin(index, 1, len - 1);//限定范围
                d = GetLinValue(pX[index - 1], pD[index - 1], pX[index], pD[index], slot);
            }
            return true;
        }

    }



    /// <summary>
    /// 转换函数类
    /// </summary>
    public static class ConvertFunc
    {

        /// <summary>
        /// .awf文件格式是由哪个插件生成的
        /// </summary>
        /// <param name="_formatCode"></param>
        /// <returns></returns>
        public static string ConvertAwfilePlugToString(long _formatCode)
        {
            string _formatS = String.Empty;
            if (_formatCode == 0x0001)
            {
                _formatS = "DDWS";//普通数据文件，即DDWS模式下生成的所有数据
            }
            else if (_formatCode == 0x0002)
            {
                _formatS = "IQ_Plug";//I路
            }
            else if (_formatCode == 0x0003)
            {
                _formatS = "IF/RF_Plug";//IF/RF
            }
            else if (_formatCode == 0x0004)
            {
                _formatS = "OFDM_Plug";//OFDM
            }
            else if (_formatCode == 0x0005)
            {
                _formatS = "Radar_Plug";//Radar
            }
            else if (_formatCode == 0x0006)
            {
                _formatS = "GSM_Plug";//GSM
            }
            else if (_formatCode == 0x0007)
            {
                _formatS = "WIFI_Plug";//WIFI
            }
            else if (_formatCode == 0x0008)
            {
                _formatS = "WCDMA_Plug";//WCDMA
            }
            else if (_formatCode == 0x0009)
            {
                _formatS = "UWB_Plug";//UWB
            }
            else if (_formatCode == 0x000A)
            {
                _formatS = "Interf_Plug";//ComplexInterf
            }
            else if (_formatCode == 0x000B)
            {
                _formatS = "EME_Plug";//EME
            }
            else if (_formatCode == 0x000C)
            {
                _formatS = "PulseRadar_Plug";//PulseRadar
            }
            else if (_formatCode == 0x1000)
            {
                _formatS = "Drawing";//任意波形编辑
            }
            else
            {
                _formatS = "unknown";//未知
            }
            return _formatS;
        }



        /// <summary>
        /// .awf文件的存储形式，I路数据存储，Q路数据存储，IQ数据存储（仅限IQ插件下的一种特殊存储形式，需要硬件做上变频调制），RF数据存储（软件处理时即加载波调制）
        /// </summary>
        /// <param name="_formatCode"></param>
        /// <returns></returns>
        public static string ConvertAwfileSaveFormatToString(long _formatCode)
        {
            string _formatS = "unknown";
            if (_formatCode == 0)
            {
                _formatS = "Com";//常规数据，非IQ调制；如DDWS生成的信号和波形编辑产生的任意波形信号
            }
            if (_formatCode == 0x0001)
            {
                _formatS = "I";//I路数据
            }
            else if (_formatCode == 0x0002)
            {
                _formatS = "Q";//Q路数据
            }
            else if (_formatCode == 0x0003)
            {
                _formatS = "IQ";//1个I和1个Q路数据交叉存储
            }
            else if (_formatCode == 0x0004)
            {
                _formatS = "RF";//IQ数据合并，并加载波处理
            }
            return _formatS;
        }

        /// <summary>
        /// 阻抗50欧姆，单位转换 dBm->Vpp
        /// </summary>
        /// <param name="_dbm"></param>
        /// <returns></returns>
        public static double DbmToVpp(double _dbm)
        {
            double d = 0;
            if (_dbm < -200.0)
            {
                d = 0;
            }
            else
            {
                d = 20.0 * Math.Pow(10.0, (_dbm / 20.0 - 1.50));
            }
            return d;
        }

        /// <summary>
        /// 阻抗50欧姆，单位转换 dBm->Vpp
        /// </summary>
        /// <param name="_dbm"></param>
        /// <returns></returns>
        public static double VppToDbm(double _vpp)
        {
            //50欧姆阻抗V->dBm转换
            double dbm = 0;
            if (_vpp < 0.0000001)
            {
                dbm = 0 - 200.0;
            }
            else
            {
                dbm = 20 * Math.Log10(_vpp / 20.0) + 30;
            }
            return dbm;
        }

        /// <summary>
        /// 峰峰值转有效值
        /// </summary>
        /// <param name="_vpp"></param>
        /// <returns></returns>
        public static double VppToVrms(double _vpp)
        {
            double _vrms = 0;
            _vrms = _vpp / 2.8284;
            return _vrms;
        }

        /// <summary>
        ///有效值转峰峰值
        /// </summary>
        /// <param name="_vpp"></param>
        /// <returns></returns>
        public static double VrmsToVpp(double _vrms)
        {
            double _vpp = 0;
            _vpp = _vrms * 2.8284;
            return _vpp;
        }

        /// <summary>
        ///有效值转dBm
        /// </summary>
        /// <param name="_vpp"></param>
        /// <returns></returns>
        public static double VrmsToDbm(double _vrms)
        {
            return VppToDbm(_vrms * 2.8284);
        }

        /// <summary>
        ///dBm转有效值
        /// </summary>
        /// <param name="_vpp"></param>
        /// <returns></returns>
        public static double DbmToVrms(double _dbm)
        {
            return DbmToVpp(_dbm) / 2.8284;
        }

        /// <summary>
        /// .awf文件采样率
        /// </summary>
        /// <param name="sampleRate"></param>
        /// <returns></returns>
        public static string ConvertSampleRateToString(double sampleRate)
        {
            string strValue = String.Empty;
            string strUnit = String.Empty;
            if (sampleRate < 312500000.1)
            {
                //10kHz - 312.5MHz连续可变
                if (sampleRate < 1000000.0)
                {
                    strUnit = "kSa/s";
                    strValue = (sampleRate / 1000.0).ToString("0.0");
                }
                else
                {
                    strUnit = "MSa/s";
                    strValue = (sampleRate / 1000000.0).ToString("0.00");
                }
            }
            else if (sampleRate < 1000000000.0)
            {
                strUnit = "MSa/s";
                strValue = (sampleRate / 1000000.0).ToString("0.00");
            }
            else
            {
                strUnit = "GSa/s";
                strValue = (sampleRate / 1000000000.0).ToString("0.00");
            }
            return strValue + strUnit;
        }




        /// <summary>
        /// 
        /// </summary>
        /// <param name="counts"></param>
        /// <returns></returns>
        public static string ConvertCountsToString(double counts)
        {
            string strValue = String.Empty;
            string strUnit = String.Empty;
            if (counts < 1000.0)
            {
                strUnit = "";
                strValue = counts.ToString("0");
            }
            else if (counts < 1000000.0)
            {
                strUnit = "k";
                strValue = (counts / 1000.0).ToString("0.000");
            }
            else if (counts < 1000000000.0)
            {
                strUnit = "M";
                strValue = (counts / 1000000.0).ToString("0.000");
            }
            else if (counts < 1000000000000.0)
            {
                strUnit = "G";
                strValue = (counts / 1000000000.0).ToString("0.0000");
            }
            return strValue + strUnit;
        }





        /// <summary>
        /// 将时间转换为字符串
        /// </summary>
        /// <param name="timeS"></param>
        /// <returns></returns>
        public static string ConvertTimeToString(double timeS)
        {
            string strUnit = string.Empty;
            string strValue = String.Empty;

            if (Math.Abs(timeS - 0) < double.Epsilon)
            {
                strUnit = "us";
                strValue = "0.00";
            }
            else if (timeS >= 1)
            {
                strUnit = "s";
                strValue = timeS.ToString("0.000");
            }
            else if (timeS * 1000 >= 1)
            {
                strUnit = "ms";
                strValue = (timeS * 1000).ToString("0.000");
            }
            else if (timeS * 1000 * 1000 >= 1)
            {
                strUnit = "μs";
                strValue = (timeS * 1000 * 1000).ToString("0.000");
            }
            else if (timeS * 1000 * 1000 * 1000 >= 1)
            {
                strUnit = "ns";
                strValue = (timeS * 1000 * 1000 * 1000).ToString("0.000");
            }

            return strValue + " " + strUnit;
        }

        /// <summary>
        /// 将字符串转换为时间
        /// </summary>
        /// <param name="strTime"></param>
        /// <returns></returns>
        public static double ConvertStringToTime(string strTime)
        {
            double multi = 1;

            strTime = strTime.ToLower();
            if (strTime.Contains("ns")) { multi = 1000 * 1000 * 1000; }
            else if (strTime.Contains("us")) { multi = 1000 * 1000; }
            else if (strTime.Contains("ms")) { multi = 1000; }
            else if (strTime.Contains("s")) { multi = 1; }
            else { multi = 1000 * 1000; }

            double value = GetValueFromString(strTime);
            return value / multi;
        }



        #region 频率转换函数

        /// <summary>
        /// 
        /// </summary>
        /// <param name="frequencyHz"></param>
        /// <returns></returns>
        public static string ConvertFrequencyToString(double frequencyHz)
        {
            string strValue = String.Empty;
            string strUnit = String.Empty;

            if (Math.Abs(frequencyHz - 0) < double.Epsilon)
            {
                strUnit = "Hz";
                strValue = "0.00";
            }
            else if (frequencyHz >= 1000 * 1000 * 1000)
            {
                strUnit = "GHz";
                strValue = (frequencyHz / 1000 / 1000 / 1000).ToString("0.000000000");
            }
            else if (frequencyHz >= 1000 * 1000)
            {
                strUnit = "MHz";
                strValue = (frequencyHz / 1000 / 1000).ToString("0.000000");
            }
            else if (frequencyHz >= 1000)
            {
                strUnit = "KHz";
                strValue = (frequencyHz / 1000).ToString("0.000");
            }
            else
            {
                strUnit = "Hz";
                strValue = frequencyHz.ToString("0");
            }

            return strValue + " " + strUnit;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="strFrequencyHz"></param>
        /// <returns></returns>
        public static double ConvertStringToFrequency(string strFrequencyHz)
        {
            double multi = 1;
            strFrequencyHz = strFrequencyHz.ToUpper();
            if (strFrequencyHz.Contains("GHz"))
            {
                multi = 1000 * 1000 * 1000;
            }
            else if (strFrequencyHz.Contains("MHZ"))
            {
                multi = 1000 * 1000;
            }
            else if (strFrequencyHz.Contains("KHZ"))
            {
                multi = 1000;
            }
            else if (strFrequencyHz.Contains("HZ"))
            {
                multi = 1;
            }
            double value = GetValueFromString(strFrequencyHz);
            return value * multi;
        }

        #endregion

        #region 将角度转换为字符串

        /// <summary>
        /// 将角度转换为字符串
        /// </summary>
        /// <param name="degreeValue"></param>
        /// <returns></returns>
        public static string ConvertDegreeToString(double degreeValue)
        {
            string strValue = String.Empty;
            string strUnit = String.Empty;

            strValue = degreeValue.ToString("0.00");
            strUnit = "°";

            return strValue + " " + strUnit;
        }

        /// <summary>
        /// 将字符串转换为度
        /// </summary>
        /// <param name="strDegreeValue"></param>
        /// <returns></returns>
        public static double ConvertStringToDegree(string strDegreeValue)
        {
            double value = GetValueFromString(strDegreeValue);
            if (value < 0 || value > 360)
            {
                value = 0;
            }
            return value;
        }

        #endregion

        #region 将Dbm转换为字符串

        /// <summary>
        /// 将角度转换为字符串
        /// </summary>
        /// <param name="degreeValue"></param>
        /// <returns></returns>
        public static string ConvertDbmToString(double degreeValue)
        {
            string strValue = String.Empty;
            string strUnit = String.Empty;

            strValue = degreeValue.ToString("0.00");
            strUnit = "dBm";

            return strValue + " " + strUnit;
        }

        /// <summary>
        /// 将字符串转换为角度
        /// </summary>
        /// <param name="strDegreeValue"></param>
        /// <returns></returns>
        public static double ConvertStringToDbm(string strDegreeValue)
        {
            double value = GetValueFromString(strDegreeValue);
            if (value < 0 || value > 360)
            {
                value = 0;
            }
            return value;
        }

        #endregion

        #region 将扫描速度转换为字符串

        /// <summary>
        /// 将角度转换为字符串
        /// </summary>
        /// <param name="degreeValue"></param>
        /// <returns></returns>
        public static string ConvertScanRateToString(double degreeValue)
        {
            string strValue = String.Empty;
            string strUnit = String.Empty;

            strValue = degreeValue.ToString("0.00");
            strUnit = " ° / s";

            return strValue + " " + strUnit;
        }

        /// <summary>
        /// 将字符串转换为角度
        /// </summary>
        /// <param name="strDegreeValue"></param>
        /// <returns></returns>
        public static double ConvertStringToScanRate(string strDegreeValue)
        {
            double value = GetValueFromString(strDegreeValue);
            return value;
        }

        #endregion

        /// <summary>
        /// 获取给定字符串最开头的数字值
        /// </summary>
        /// <param name="strValue"></param>
        /// <returns></returns>
        public static double GetValueFromString(string strValue)
        {
            int index = -1;
            double result;
            if (double.TryParse(strValue, out result)) { return result; }

            for (int i = 1; i <= strValue.Length; i++)
            {
                //此处在末尾加0，为的是防止在前面有“+”或者“-”的情况下，会发生问题
                if (!double.TryParse(strValue.Substring(0, i) + "0", out result))
                {
                    index = i;
                    break;
                }
            }

            if (index == -1 || index == 1)
            {
                return 0;
            }

            return double.Parse(strValue.Substring(0, index - 1));
        }








    }

}