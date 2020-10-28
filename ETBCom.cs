using System;
using System.Diagnostics;
using System.IO.Ports;
using System.Text;
using System.Collections.Generic;

namespace ETB2020
{
    /// <summary>
    /// 概要:EME TDDB シリアルコム通信 基本クラス
    /// 作成日:2020/09/30 作成者:murata
    /// </summary>
    class ETBCom
    {
        #region "プロパティ"
        /// <summary>
        /// プロパティ
        /// 作成日:2020/09/30 作成者:murata
        /// </summary> 
        private SerialPort mPort = new SerialPort();        // シリアル通信オブジェクト
        private int mPortNo;                                // ポート番号
        private int mBaudRate;                              // 通信速度
        private int mDataBit;                               // データ長
        private Parity mParity;                             // パリティ
        private StopBits mStopBit;                          // ストップビット
        private int mTimeOut;                               // タイムアウト時間
        private List<Byte> mSend = new List<Byte>();        // 送信データ
        private List<Byte> mRecv = new List<Byte>();        // 受信データ
        private const byte COMMA = 0x2c;                    // コンマ(ASCIIコード）
        private const byte DELIMITER = 0xa;                 // デリミタ
        private const long SEND_NUM = 18;                   // 送信バイト数
        private const long RECV_NUM = 20;                   // 受信バイト数
        private const int TIMEOUT_PC = 3000;               // パソコン側タイムアウト(ms)
        #endregion

        #region "アクセサ"
        /// <summary>
        /// アクセサ
        /// 作成日:2020/09/30 作成者:murata
        /// </summary>         
        //ポート状態
        public bool IsOpen
        {
            get { return mPort.IsOpen; }
        }
        #endregion 

        #region "コンストラクタ"
        /// <summary>
        /// コンストラクタ
        /// 作成日:2020/09/30 作成者:murata
        /// </summary>         
        public ETBCom(int port, int rate, int bit, Parity parity, StopBits stop)
        {
            mPortNo = port;
            mBaudRate = rate;
            mDataBit = bit;
            mParity = parity;
            mStopBit = stop;
            mTimeOut = TIMEOUT_PC;

        }
        #endregion

        #region "公開メソッド"
        /// <summary>
        /// ポートオープン
        /// 作成日:2020/09/30 作成者:murata
        /// </summary>
        public bool Open()
        {
            bool ret = false;

            try
            {
                mPort.Close();
                mPort.PortName = string.Format("COM{0}", mPortNo);
                mPort.BaudRate = mBaudRate;
                mPort.DataBits = mDataBit;
                mPort.Parity = mParity;
                mPort.StopBits = mStopBit;
                mPort.Handshake = Handshake.None;
                mPort.RtsEnable = false;
                mPort.ReadTimeout = TIMEOUT_PC;
                mPort.WriteTimeout = TIMEOUT_PC;
                mPort.Open();
                ret = mPort.IsOpen;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }

            return ret;
        }

        /// <summary>
        /// ポートクローズ
        /// 作成日:2020/09/30 作成者:murata
        /// </summary>
        public void Close()
        {
            try
            {
                mPort.Close();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        /// <summary>
        /// バージョンチェックコマンド
        /// 作成日:2020/09/30 作成者:murata
        /// </summary>
        public bool VersionCheck(string boardno, out string version)
        {
            bool bRet;
            string cmd = "";
            string bno = "";
            string data = "";
            string apply = "";

            //データ初期化
            version = "";

            //送信データ作成
            bRet = MakeSendData(boardno, "VER", "12345678");
            if (bRet == false)
            {
                Console.WriteLine("送信データ作成エラー");
                return (false);
            }

            //送信処理
            bRet = SendData();
            if (bRet == false)
            {
                Console.WriteLine("送信エラー");
                return (false);
            }

            //受信処理
            bRet = RecvData();
            if (bRet == false)
            {
                Console.WriteLine("受信エラー:タイムアウト");
                return (false);
            }

            //受信データ解析
            bRet = RecvDataAnalyze(out bno, out cmd, out data, out apply);
            if (bRet == false)
            {
                Console.WriteLine("受信エラー:データ異常");
                return (false);
            }

            //ボード番号チェック
            if (bno != boardno)
            {
                Console.WriteLine("受信エラー:ボードNo");
                return (false);
            }

            //応答コードチェック
            if (apply != "0")
            {

                Console.WriteLine("受信エラー:応答コードNG コード:" + apply);
                return (false);
            }

            //コマンドチェック
            if (cmd != "VER")
            {
                Console.WriteLine("受信エラー:コマンドID");
                return (false);
            }

            //バージョン番号取得
            version = data;


            return (true);
        }
        /// <summary>
        /// 通信チェックコマンド
        /// </summary>
        public bool ConnectCheck(string boardno)
        {
            bool bRet;
            string cmd = "";
            string bno = "";
            string data = "";
            string apply = "";

            //送信データ作成
            bRet = MakeSendData(boardno, "CCK", "12345678");
            if (bRet == false)
            {
                Console.WriteLine("送信データ作成エラー");
                return (false);
            }

            //送信処理
            bRet = SendData();
            if (bRet == false)
            {
                Console.WriteLine("送信エラー");
                return (false);
            }

            //受信処理
            bRet = RecvData();
            if (bRet == false)
            {
                Console.WriteLine("受信エラー:タイムアウト");
                return (false);
            }

            //受信データ解析
            bRet = RecvDataAnalyze(out bno, out cmd, out data, out apply);
            if (bRet == false)
            {
                Console.WriteLine("受信エラー:データ異常");
                return (false);
            }

            //ボード番号チェック
            if (bno != boardno)
            {
                Console.WriteLine("受信エラー:ボードNo");
                return (false);
            }

            //応答コードチェック
            if (apply != "0")
            {

                Console.WriteLine("受信エラー:応答コードNG コード:" + apply);
                return (false);
            }

            //コマンドチェック
            if (cmd != "CCK")
            {
                Console.WriteLine("受信エラー:コマンドID");
                return (false);
            }

            return (true);
        }

        /// <summary>
        /// 電流設定コマンド
        /// 作成日:2020/09/30 作成者:murata
        /// </summary>
        public bool VoltSet(string boardno, int value)
        {
            bool bRet;
            string cmd = "";
            string bno = "";
            string data = "";
            string apply = "";
            string datasend = value.ToString("00000000");

            //送信データ作成
            bRet = MakeSendData(boardno, "SCR", datasend);
            if (bRet == false)
            {
                Console.WriteLine("送信データ作成エラー");
                return (false);
            }

            //送信処理
            bRet = SendData();
            if (bRet == false)
            {
                Console.WriteLine("送信エラー");
                return (false);
            }

            //受信処理
            bRet = RecvData();
            if (bRet == false)
            {
                Console.WriteLine("受信エラー:タイムアウト");
                return (false);
            }

            //受信データ解析
            bRet = RecvDataAnalyze(out bno, out cmd, out data, out apply);
            if (bRet == false)
            {
                Console.WriteLine("受信エラー:データ異常");
                return (false);
            }

            //ボード番号チェック
            if (bno != boardno)
            {
                Console.WriteLine("受信エラー:ボードNo");
                return (false);
            }

            //応答コードチェック
            if (apply != "0")
            {

                Console.WriteLine("受信エラー:応答コードNG コード:" + apply);
                return (false);
            }

            //コマンドチェック
            if (cmd != "SCR")
            {
                Console.WriteLine("受信エラー:コマンドID");
                return (false);
            }

            return (true);
        }

        /// <summary>
        /// 電流設定確認コマンド
        /// 作成日:2020/10/09 作成者:huy
        /// </summary>
        public bool VoltCheck(string boardno, out int volt)
        {
            bool bRet;
            string cmd = "";
            string bno = "";
            string data = "";
            string apply = "";

            //データ初期化
            volt = 0;

            //送信データ作成
            bRet = MakeSendData(boardno, "GCR", "00000000");
            if (bRet == false)
            {
                Console.WriteLine("送信データ作成エラー");
                return (false);
            }

            //送信処理
            bRet = SendData();
            if (bRet == false)
            {
                Console.WriteLine("送信エラー");
                return (false);
            }

            //受信処理
            bRet = RecvData();
            if (bRet == false)
            {
                Console.WriteLine("受信エラー:タイムアウト");
                return (false);
            }

            //受信データ解析
            bRet = RecvDataAnalyze(out bno, out cmd, out data, out apply);
            if (bRet == false)
            {
                Console.WriteLine("受信エラー:データ異常");
                return (false);
            }

            //ボード番号チェック
            if (bno != boardno)
            {
                Console.WriteLine("受信エラー:ボードNo");
                return (false);
            }

            //応答コードチェック
            if (apply != "0")
            {

                Console.WriteLine("受信エラー:応答コードNG コード:" + apply);
                return (false);
            }

            //コマンドチェック
            if (cmd != "GCR")
            {
                Console.WriteLine("受信エラー:コマンドID");
                return (false);
            }

            //電源値取得
            if (int.TryParse(data, out volt) == false)
            {
                Console.WriteLine("受信エラー:データ変換エラー");
                return (false);
            }

            return (true);
        }

        /// <summary>
        /// 試験実行時間クリアコマンド
        /// 作成日:2020/10/09 作成者:huy
        /// </summary>
        public bool ClearTestTime(string boardno)
        {
            bool bRet;
            string cmd = "";
            string bno = "";
            string data = "";
            string apply = "";

            //送信データ作成
            bRet = MakeSendData(boardno, "CET", "12345678");
            if (bRet == false)
            {
                Console.WriteLine("送信データ作成エラー");
                return (false);
            }

            //送信処理
            bRet = SendData();
            if (bRet == false)
            {
                Console.WriteLine("送信エラー");
                return (false);
            }

            //受信処理
            bRet = RecvData();
            if (bRet == false)
            {
                Console.WriteLine("受信エラー:タイムアウト");
                return (false);
            }

            //受信データ解析
            bRet = RecvDataAnalyze(out bno, out cmd, out data, out apply);
            if (bRet == false)
            {
                Console.WriteLine("受信エラー:データ異常");
                return (false);
            }

            //ボード番号チェック
            if (bno != boardno)
            {
                Console.WriteLine("受信エラー:ボードNo");
                return (false);
            }

            //応答コードチェック
            if (apply != "0")
            {

                Console.WriteLine("受信エラー:応答コードNG コード:" + apply);
                return (false);
            }

            //コマンドチェック
            if (cmd != "CET")
            {
                Console.WriteLine("受信エラー:コマンドID");
                return (false);
            }

            return (true);
        }

        /// <summary>
        /// 試験実行時間取得コマンド
        /// 作成日:2020/10/09 作成者:huy
        /// </summary>
        public bool GetTestTime(string boardno, int chno, out int time)
        {

            bool bRet;
            string cmd = "";
            string bno = "";
            string data = "";
            string apply = "";
            string datasend = chno.ToString("00000000");

            //データ初期化
            time = 0;

            //送信データ作成
            bRet = MakeSendData(boardno, "GET", datasend);
            if (bRet == false)
            {
                Console.WriteLine("送信データ作成エラー");
                return (false);
            }

            //送信処理
            bRet = SendData();
            if (bRet == false)
            {
                Console.WriteLine("送信エラー");
                return (false);
            }

            //受信処理
            bRet = RecvData();
            if (bRet == false)
            {
                Console.WriteLine("受信エラー:タイムアウト");
                return (false);
            }

            //受信データ解析
            bRet = RecvDataAnalyze(out bno, out cmd, out data, out apply);
            if (bRet == false)
            {
                Console.WriteLine("受信エラー:データ異常");
                return (false);
            }

            //ボード番号チェック
            if (bno != boardno)
            {
                Console.WriteLine("受信エラー:ボードNo");
                return (false);
            }

            //応答コードチェック
            if (apply != "0")
            {

                Console.WriteLine("受信エラー:応答コードNG コード:" + apply);
                return (false);
            }


            //コマンドチェック
            if (cmd != "GET")
            {
                Console.WriteLine("受信エラー:コマンドID");
                return (false);
            }

            //試験実行時間取得
            if (int.TryParse(data, out time) == false)
            {
                Console.WriteLine("受信エラー:データ変換エラー");
                return (false);
            }

            return (true);
        }


        /// <summary>
        /// ストレス印可開始コマンド
        /// 作成日:2020/10/09 作成者:huy
        /// </summary>
        public bool StartApplyStress(string boardno, int chno)
        {
            bool bRet;
            string cmd = "";
            string bno = "";
            string data = "";
            string apply = "";
            string datasend = chno.ToString("00000000");

            //送信データ作成
            bRet = MakeSendData(boardno, "SSS", datasend);
            if (bRet == false)
            {
                Console.WriteLine("送信データ作成エラー");
                return (false);
            }

            //送信処理
            bRet = SendData();
            if (bRet == false)
            {
                Console.WriteLine("送信エラー");
                return (false);
            }

            //受信処理
            bRet = RecvData();
            if (bRet == false)
            {
                Console.WriteLine("受信エラー:タイムアウト");
                return (false);
            }

            //受信データ解析
            bRet = RecvDataAnalyze(out bno, out cmd, out data, out apply);
            if (bRet == false)
            {
                Console.WriteLine("受信エラー:データ異常");
                return (false);
            }

            //ボード番号チェック
            if (bno != boardno)
            {
                Console.WriteLine("受信エラー:ボードNo");
                return (false);
            }

            //応答コードチェック
            if (apply != "0")
            {
                Console.WriteLine("受信エラー:応答コードNG コード:" + apply);
                return (false);
            }

            //応答チャンネル　チェック
            if (datasend != data)
            {
                Console.WriteLine("受信エラー:チャンネル番号異常");
                return (false);
            }

            //コマンドチェック
            if (cmd != "SSS")
            {
                Console.WriteLine("受信エラー:コマンドID");
                return (false);
            }

            return (true);
        }

        /// <summary>
        /// ストレス印可終了コマンド
        /// 作成日:2020/10/09 作成者:huy
        /// </summary>
        public bool StopApplyStress(string boardno, int chno)
        {
            bool bRet;
            string cmd = "";
            string bno = "";
            string data = "";
            string apply = "";
            string datasend = chno.ToString("00000000");

            //送信データ作成
            bRet = MakeSendData(boardno, "SSE", datasend);
            if (bRet == false)
            {
                Console.WriteLine("送信データ作成エラー");
                return (false);
            }

            //送信処理
            bRet = SendData();
            if (bRet == false)
            {
                Console.WriteLine("送信エラー");
                return (false);
            }

            //受信処理
            bRet = RecvData();
            if (bRet == false)
            {
                Console.WriteLine("受信エラー:タイムアウト");
                return (false);
            }

            //受信データ解析
            bRet = RecvDataAnalyze(out bno, out cmd, out data, out apply);
            if (bRet == false)
            {
                Console.WriteLine("受信エラー:データ異常");
                return (false);
            }

            //ボード番号チェック
            if (bno != boardno)
            {
                Console.WriteLine("受信エラー:ボードNo");
                return (false);
            }

            //応答コードチェック
            if (apply != "0")
            {

                Console.WriteLine("受信エラー:応答コードNG コード:" + apply);
                return (false);
            }

            //応答チャンネル　チェック
            if (datasend != data)
            {
                Console.WriteLine("受信エラー:チャンネル番号異常");
                return (false);
            }


            //コマンドチェック
            if (cmd != "SSE")
            {
                Console.WriteLine("受信エラー:コマンドID");
                return (false);
            }

            return (true);
        }

        /// <summary>
        /// 測定開始コマンド
        /// 作成日:2020/10/09 作成者:huy
        /// </summary>
        public bool StartTest(string boardno, int chno)
        {
            bool bRet;
            string cmd = "";
            string bno = "";
            string data = "";
            string apply = "";
            string datasend = chno.ToString("00000000");

            //送信データ作成
            bRet = MakeSendData(boardno, "STS", datasend);
            if (bRet == false)
            {
                Console.WriteLine("送信データ作成エラー");
                return (false);
            }

            //送信処理
            bRet = SendData();
            if (bRet == false)
            {
                Console.WriteLine("送信エラー");
                return (false);
            }

            //受信処理
            bRet = RecvData();
            if (bRet == false)
            {
                Console.WriteLine("受信エラー:タイムアウト");
                return (false);
            }

            //受信データ解析
            bRet = RecvDataAnalyze(out bno, out cmd, out data, out apply);
            if (bRet == false)
            {
                Console.WriteLine("受信エラー:データ異常");
                return (false);
            }

            //ボード番号チェック
            if (bno != boardno)
            {
                Console.WriteLine("受信エラー:ボードNo");
                return (false);
            }

            //応答コードチェック
            if (apply != "0")
            {

                Console.WriteLine("受信エラー:応答コードNG コード:" + apply);
                return (false);
            }

            //応答チャンネル　チェック
            if (datasend != data)
            {
                Console.WriteLine("受信エラー:チャンネル番号異常");
                return (false);
            }

            //コマンドチェック
            if (cmd != "STS")
            {
                Console.WriteLine("受信エラー:コマンドID");
                return (false);
            }

            return (true);
        }

        /// <summary>
        /// 測定終了コマンド
        /// 作成日:2020/10/09 作成者:huy
        /// </summary>
        public bool StopTest(string boardno, int chno)
        {
            bool bRet;
            string cmd = "";
            string bno = "";
            string data = "";
            string apply = "";
            string datasend = chno.ToString("00000000");

            //送信データ作成
            bRet = MakeSendData(boardno, "STE", datasend);
            if (bRet == false)
            {
                Console.WriteLine("送信データ作成エラー");
                return (false);
            }

            //送信処理
            bRet = SendData();
            if (bRet == false)
            {
                Console.WriteLine("送信エラー");
                return (false);
            }

            //受信処理
            bRet = RecvData();
            if (bRet == false)
            {
                Console.WriteLine("受信エラー:タイムアウト");
                return (false);
            }

            //受信データ解析
            bRet = RecvDataAnalyze(out bno, out cmd, out data, out apply);
            if (bRet == false)
            {
                Console.WriteLine("受信エラー:データ異常");
                return (false);
            }

            //ボード番号チェック
            if (bno != boardno)
            {
                Console.WriteLine("受信エラー:ボードNo");
                return (false);
            }

            //応答コードチェック
            if (apply != "0")
            {

                Console.WriteLine("受信エラー:応答コードNG コード:" + apply);
                return (false);
            }

            //応答チャンネル　チェック
            if (datasend != data)
            {
                Console.WriteLine("受信エラー:チャンネル番号異常");
                return (false);
            }

            //コマンドチェック
            if (cmd != "STE")
            {
                Console.WriteLine("受信エラー:コマンドID");
                return (false);
            }

            return (true);
        }

        /// <summary>
        /// 動作状態取得コマンド
        /// 作成日:2020/10/09 作成者:huy
        /// </summary>
        public bool GetStatus(string boardno, int chno, out int status)
        {
            bool bRet;
            string cmd = "";
            string bno = "";
            string data = "";
            string apply = "";
            string datasend = chno.ToString("00000000");

            //データ初期化
            status = 0;

            //送信データ作成
            bRet = MakeSendData(boardno, "GST", datasend);
            if (bRet == false)
            {
                Console.WriteLine("送信データ作成エラー");
                return (false);
            }

            //送信処理
            bRet = SendData();
            if (bRet == false)
            {
                Console.WriteLine("送信エラー");
                return (false);
            }

            //受信処理
            bRet = RecvData();
            if (bRet == false)
            {
                Console.WriteLine("受信エラー:タイムアウト");
                return (false);
            }

            //受信データ解析
            bRet = RecvDataAnalyze(out bno, out cmd, out data, out apply);
            if (bRet == false)
            {
                Console.WriteLine("受信エラー:データ異常");
                return (false);
            }

            //ボード番号チェック
            if (bno != boardno)
            {
                Console.WriteLine("受信エラー:ボードNo");
                return (false);
            }

            //応答コードチェック
            if (apply != "0")
            {
                Console.WriteLine("受信エラー:応答コードNG コード:" + apply);
                return (false);
            }

            //コマンドチェック
            if (cmd != "GST")
            {
                Console.WriteLine("受信エラー:コマンドID");
                return (false);
            }

            //応答チャンネル　チェック
            if (chno.ToString("0000") != data.Substring(4, 4))
            {
                Console.WriteLine("受信エラー:チャンネル番号異常");
                return (false);
            }

            //ステータス セット
            status |= (data.Substring(0, 1) == "1" ? (1 << 3) : 0);
            status |= (data.Substring(1, 1) == "1" ? (1 << 2) : 0);
            status |= (data.Substring(2, 1) == "1" ? (1 << 1) : 0);
            status |= (data.Substring(3, 1) == "1" ? (1 << 0) : 0);

            return (true);
        }

        /// <summary>
        /// ストレス電流確認コマンド
        /// 作成日:2020/10/09 作成者:huy
        /// </summary>
        public bool GetStressCurrent(string boardno, int value, out int current)
        {
            bool bRet;
            string cmd = "";
            string bno = "";
            string data = "";
            string apply = "";
            string datasend = value.ToString("00000000");

            //データ初期化
            current = 0;

            //送信データ作成
            bRet = MakeSendData(boardno, "GSC", datasend);
            if (bRet == false)
            {
                Console.WriteLine("送信データ作成エラー");
                return (false);
            }

            //送信処理
            bRet = SendData();
            if (bRet == false)
            {
                Console.WriteLine("送信エラー");
                return (false);
            }

            //受信処理
            bRet = RecvData();
            if (bRet == false)
            {
                Console.WriteLine("受信エラー:タイムアウト");
                return (false);
            }

            //受信データ解析
            bRet = RecvDataAnalyze(out bno, out cmd, out data, out apply);
            if (bRet == false)
            {
                Console.WriteLine("受信エラー:データ異常");
                return (false);
            }

            //ボード番号チェック
            if (bno != boardno)
            {
                Console.WriteLine("受信エラー:ボードNo");
                return (false);
            }

            //応答コードチェック
            if (apply != "0")
            {

                Console.WriteLine("受信エラー:応答コードNG コード:" + apply);
                return (false);
            }

            //コマンドチェック
            if (cmd != "GSC")
            {
                Console.WriteLine("受信エラー:コマンドID");
                return (false);
            }

            //電流値取得
            if (int.TryParse(data, out current) == false)
            {
                Console.WriteLine("受信エラー:データ変換エラー");
                return (false);
            }

            return (true);
        }

        /// <summary>
        /// パルス周波数設定コマンド
        /// 作成日:2020/10/09 作成者:huy
        /// </summary>
        public bool PulseCycleSet(string boardno, int value)
        {
            bool bRet;
            string cmd = "";
            string bno = "";
            string data = "";
            string apply = "";
            string datasend = value.ToString("00000000");

            //送信データ作成
            bRet = MakeSendData(boardno, "SPF", datasend);
            if (bRet == false)
            {
                Console.WriteLine("送信データ作成エラー");
                return (false);
            }

            //送信処理
            bRet = SendData();
            if (bRet == false)
            {
                Console.WriteLine("送信エラー");
                return (false);
            }

            //受信処理
            bRet = RecvData();
            if (bRet == false)
            {
                Console.WriteLine("受信エラー:タイムアウト");
                return (false);
            }

            //受信データ解析
            bRet = RecvDataAnalyze(out bno, out cmd, out data, out apply);
            if (bRet == false)
            {
                Console.WriteLine("受信エラー:データ異常");
                return (false);
            }

            //ボード番号チェック
            if (bno != boardno)
            {
                Console.WriteLine("受信エラー:ボードNo");
                return (false);
            }

            //応答コードチェック
            if (apply != "0")
            {

                Console.WriteLine("受信エラー:応答コードNG コード:" + apply);
                return (false);
            }

            //コマンドチェック
            if (cmd != "SPF")
            {
                Console.WriteLine("受信エラー:コマンドID");
                return (false);
            }

            return (true);
        }

        /// <summary>
        /// パルス周波数取得コマンド
        /// 作成日:2020/10/09 作成者:huy
        /// </summary>
        public bool PulseCycleGet(string boardno, out int cycle)
        {
            bool bRet;
            string cmd = "";
            string bno = "";
            string data = "";
            string apply = "";

            //データ初期化
            cycle = 0;

            //送信データ作成
            bRet = MakeSendData(boardno, "GPF", "00000000");
            if (bRet == false)
            {
                Console.WriteLine("送信データ作成エラー");
                return (false);
            }

            //送信処理
            bRet = SendData();
            if (bRet == false)
            {
                Console.WriteLine("送信エラー");
                return (false);
            }

            //受信処理
            bRet = RecvData();
            if (bRet == false)
            {
                Console.WriteLine("受信エラー:タイムアウト");
                return (false);
            }

            //受信データ解析
            bRet = RecvDataAnalyze(out bno, out cmd, out data, out apply);
            if (bRet == false)
            {
                Console.WriteLine("受信エラー:データ異常");
                return (false);
            }

            //ボード番号チェック
            if (bno != boardno)
            {
                Console.WriteLine("受信エラー:ボードNo");
                return (false);
            }

            //応答コードチェック
            if (apply != "0")
            {

                Console.WriteLine("受信エラー:応答コードNG コード:" + apply);
                return (false);
            }

            //コマンドチェック
            if (cmd != "GPF")
            {
                Console.WriteLine("受信エラー:コマンドID");
                return (false);
            }


            //周波数取得
            if (int.TryParse(data, out cycle) == false)
            {
                Console.WriteLine("受信エラー:データ変換エラー");
                return (false);
            }

            return (true);
        }

        /// <summary>
        /// パルスデューティ設定コマンド
        /// 作成日:2020/10/09 作成者:huy
        /// </summary>
        public bool PulseDutySet(string boardno, int value)
        {
            bool bRet;
            string cmd = "";
            string bno = "";
            string data = "";
            string apply = "";
            string datasend = value.ToString("00000000");

            //送信データ作成
            bRet = MakeSendData(boardno, "SPD", datasend);
            if (bRet == false)
            {
                Console.WriteLine("送信データ作成エラー");
                return (false);
            }

            //送信処理
            bRet = SendData();
            if (bRet == false)
            {
                Console.WriteLine("送信エラー");
                return (false);
            }

            //受信処理
            bRet = RecvData();
            if (bRet == false)
            {
                Console.WriteLine("受信エラー:タイムアウト");
                return (false);
            }

            //受信データ解析
            bRet = RecvDataAnalyze(out bno, out cmd, out data, out apply);
            if (bRet == false)
            {
                Console.WriteLine("受信エラー:データ異常");
                return (false);
            }

            //ボード番号チェック
            if (bno != boardno)
            {
                Console.WriteLine("受信エラー:ボードNo");
                return (false);
            }

            //応答コードチェック
            if (apply != "0")
            {

                Console.WriteLine("受信エラー:応答コードNG コード:" + apply);
                return (false);
            }

            //コマンドチェック
            if (cmd != "SPD")
            {
                Console.WriteLine("受信エラー:コマンドID");
                return (false);
            }

            return (true);
        }

        /// <summary>
        /// パルスデューティ取得コマンド
        /// 作成日:2020/10/09 作成者:huy
        /// </summary>
        public bool PulseDutyGet(string boardno, out int duty)
        {
            bool bRet;
            string cmd = "";
            string bno = "";
            string data = "";
            string apply = "";

            //データ初期化
            duty = 0;

            //送信データ作成
            bRet = MakeSendData(boardno, "GPD", "00000000");
            if (bRet == false)
            {
                Console.WriteLine("送信データ作成エラー");
                return (false);
            }

            //送信処理
            bRet = SendData();
            if (bRet == false)
            {
                Console.WriteLine("送信エラー");
                return (false);
            }

            //受信処理
            bRet = RecvData();
            if (bRet == false)
            {
                Console.WriteLine("受信エラー:タイムアウト");
                return (false);
            }

            //受信データ解析
            bRet = RecvDataAnalyze(out bno, out cmd, out data, out apply);
            if (bRet == false)
            {
                Console.WriteLine("受信エラー:データ異常");
                return (false);
            }

            //ボード番号チェック
            if (bno != boardno)
            {
                Console.WriteLine("受信エラー:ボードNo");
                return (false);
            }

            //応答コードチェック
            if (apply != "0")
            {

                Console.WriteLine("受信エラー:応答コードNG コード:" + apply);
                return (false);
            }

            //コマンドチェック
            if (cmd != "GPD")
            {
                Console.WriteLine("受信エラー:コマンドID");
                return (false);
            }

            //デューティ取得
            if (int.TryParse(data, out duty) == false)
            {
                Console.WriteLine("受信エラー:データ変換エラー");
                return (false);
            }

            return (true);
        }

        /// <summary>
        /// パルスON/OFF設定コマンド
        /// 作成日:2020/10/09 作成者:huy
        /// </summary>
        public bool PulseOnOffSet(string boardno, int value)
        {
            bool bRet;
            string cmd = "";
            string bno = "";
            string data = "";
            string apply = "";
            string datasend = value.ToString("00000000");

            //送信データ作成
            bRet = MakeSendData(boardno, "SPE", datasend);
            if (bRet == false)
            {
                Console.WriteLine("送信データ作成エラー");
                return (false);
            }

            //送信処理
            bRet = SendData();
            if (bRet == false)
            {
                Console.WriteLine("送信エラー");
                return (false);
            }

            //受信処理
            bRet = RecvData();
            if (bRet == false)
            {
                Console.WriteLine("受信エラー:タイムアウト");
                return (false);
            }

            //受信データ解析
            bRet = RecvDataAnalyze(out bno, out cmd, out data, out apply);
            if (bRet == false)
            {
                Console.WriteLine("受信エラー:データ異常");
                return (false);
            }

            //ボード番号チェック
            if (bno != boardno)
            {
                Console.WriteLine("受信エラー:ボードNo");
                return (false);
            }

            //応答コードチェック
            if (apply != "0")
            {

                Console.WriteLine("受信エラー:応答コードNG コード:" + apply);
                return (false);
            }

            //コマンドチェック
            if (cmd != "SPE")
            {
                Console.WriteLine("受信エラー:コマンドID");
                return (false);
            }

            return (true);
        }

        /// <summary>
        /// パルスON/OFF確認コマンド
        /// 作成日:2020/10/09 作成者:huy
        /// </summary>
        public bool PulseOnOffGet(string boardno, out int status)
        {
            bool bRet;
            string cmd = "";
            string bno = "";
            string data = "";
            string apply = "";

            //データ初期化
            status = 0;

            //送信データ作成
            bRet = MakeSendData(boardno, "GPE", "00000000");
            if (bRet == false)
            {
                Console.WriteLine("送信データ作成エラー");
                return (false);
            }

            //送信処理
            bRet = SendData();
            if (bRet == false)
            {
                Console.WriteLine("送信エラー");
                return (false);
            }

            //受信処理
            bRet = RecvData();
            if (bRet == false)
            {
                Console.WriteLine("受信エラー:タイムアウト");
                return (false);
            }

            //受信データ解析
            bRet = RecvDataAnalyze(out bno, out cmd, out data, out apply);
            if (bRet == false)
            {
                Console.WriteLine("受信エラー:データ異常");
                return (false);
            }

            //ボード番号チェック
            if (bno != boardno)
            {
                Console.WriteLine("受信エラー:ボードNo");
                return (false);
            }

            //応答コードチェック
            if (apply != "0")
            {

                Console.WriteLine("受信エラー:応答コードNG コード:" + apply);
                return (false);
            }

            //コマンドチェック
            if (cmd != "GPE")
            {
                Console.WriteLine("受信エラー:コマンドID");
                return (false);
            }

            //デューティ取得
            if (int.TryParse(data, out status) == false)
            {
                Console.WriteLine("受信エラー:データ変換エラー");
                return (false);
            }

            return (true);
        }

        /// <summary>
        /// 温度確認コマンド
        /// 作成日:2020/10/09 作成者:huy
        /// </summary>
        public bool GetTemp(string boardno, out int temp, out int status)
        {
            bool bRet;
            string cmd = "";
            string bno = "";
            string data = "";
            string apply = "";

            //データ初期化
            temp = 0;
            status = 0;

            //送信データ作成
            bRet = MakeSendData(boardno, "GTP", "00000000");
            if (bRet == false)
            {
                Console.WriteLine("送信データ作成エラー");
                return (false);
            }

            //送信処理
            bRet = SendData();
            if (bRet == false)
            {
                Console.WriteLine("送信エラー");
                return (false);
            }

            //受信処理
            bRet = RecvData();
            if (bRet == false)
            {
                Console.WriteLine("受信エラー:タイムアウト");
                return (false);
            }

            //受信データ解析
            bRet = RecvDataAnalyze(out bno, out cmd, out data, out apply);
            if (bRet == false)
            {
                Console.WriteLine("受信エラー:データ異常");
                return (false);
            }

            //ボード番号チェック
            if (bno != boardno)
            {
                Console.WriteLine("受信エラー:ボードNo");
                return (false);
            }

            //応答コードチェック
            if (apply != "0")
            {

                Console.WriteLine("受信エラー:応答コードNG コード:" + apply);
                return (false);
            }

            //コマンドチェック
            if (cmd != "GTP")
            {
                Console.WriteLine("受信エラー:コマンドID");
                return (false);
            }

            //温度取得
            if (int.TryParse(data.Substring(data.Length - 3), out temp) == false)
            {
                Console.WriteLine("受信エラー:データ変換エラー");
                return (false);
            }
            if (int.TryParse(data.Substring(0, 1), out status) == false)
            {
                Console.WriteLine("受信エラー:データ変換エラー");
                return (false);
            }

            return (true);

        }

        /// <summary>
        /// 電圧極性設定コマンド
        /// 作成日:2020/10/09 作成者:huy
        /// </summary>
        public bool SetVoltPolarity(string boardno, int value)
        {
            bool bRet;
            string cmd = "";
            string bno = "";
            string data = "";
            string apply = "";
            string datasend = value.ToString("00000000");

            //送信データ作成
            bRet = MakeSendData(boardno, "SVP", datasend);
            if (bRet == false)
            {
                Console.WriteLine("送信データ作成エラー");
                return (false);
            }

            //送信処理
            bRet = SendData();
            if (bRet == false)
            {
                Console.WriteLine("送信エラー");
                return (false);
            }

            //受信処理
            bRet = RecvData();
            if (bRet == false)
            {
                Console.WriteLine("受信エラー:タイムアウト");
                return (false);
            }

            //受信データ解析
            bRet = RecvDataAnalyze(out bno, out cmd, out data, out apply);
            if (bRet == false)
            {
                Console.WriteLine("受信エラー:データ異常");
                return (false);
            }

            //ボード番号チェック
            if (bno != boardno)
            {
                Console.WriteLine("受信エラー:ボードNo");
                return (false);
            }

            //応答コードチェック
            if (apply != "0")
            {

                Console.WriteLine("受信エラー:応答コードNG コード:" + apply);
                return (false);
            }

            //コマンドチェック
            if (cmd != "SVP")
            {
                Console.WriteLine("受信エラー:コマンドID");
                return (false);
            }

            return (true);
        }

        /// <summary>
        /// 電圧極性取得コマンド
        /// 作成日:2020/10/09 作成者:huy
        /// </summary>
        public bool GetVoltPolarity(string boardno, out int polarity)
        {
            bool bRet;
            string cmd = "";
            string bno = "";
            string data = "";
            string apply = "";

            //データ初期化
            polarity = 0;

            //送信データ作成
            bRet = MakeSendData(boardno, "GVP", "00000000");
            if (bRet == false)
            {
                Console.WriteLine("送信データ作成エラー");
                return (false);
            }

            //送信処理
            bRet = SendData();
            if (bRet == false)
            {
                Console.WriteLine("送信エラー");
                return (false);
            }

            //受信処理
            bRet = RecvData();
            if (bRet == false)
            {
                Console.WriteLine("受信エラー:タイムアウト");
                return (false);
            }

            //受信データ解析
            bRet = RecvDataAnalyze(out bno, out cmd, out data, out apply);
            if (bRet == false)
            {
                Console.WriteLine("受信エラー:データ異常");
                return (false);
            }

            //ボード番号チェック
            if (bno != boardno)
            {
                Console.WriteLine("受信エラー:ボードNo");
                return (false);
            }

            //応答コードチェック
            if (apply != "0")
            {

                Console.WriteLine("受信エラー:応答コードNG コード:" + apply);
                return (false);
            }

            //コマンドチェック
            if (cmd != "GVP")
            {
                Console.WriteLine("受信エラー:コマンドID");
                return (false);
            }

            //電圧極性取得
            if (int.TryParse(data, out polarity) == false)
            {
                Console.WriteLine("受信エラー:データ変換エラー");
                return (false);
            }

            return (true);
        }
        #endregion

        #region "非公開メソッド"
        /// <summary>
        /// Ascii 判定関数
        /// 作成日:2020/09/30 作成者:murata
        /// </summary>
        private bool AsciiCheck(Byte character)
        {
            if (character == 0)
            {
                return true;
            }
            if ((character < 0x20) || (character > 0x7E))
            {
                return false;
            }
            return true;
        }

        /// <summary>
        /// Ascii 数値("0"～"9")判定関数
        /// 作成日:2020/09/30 作成者:murata
        /// </summary>
        private bool AsciiNumericCheck(Byte character)
        {
            if (character == 0)
            {
                return true;
            }
            if ((character < 0x30) || (character > 0x39))
            {
                return false;
            }
            return true;
        }

        /// <summary>
        /// 送信データ作成
        /// 作成日:2020/09/30 作成者:murata
        /// </summary>
        private bool MakeSendData(string addr, string cmd, string data)
        {
            byte[] bytes;
            long sum = 0;
            long ckData;

            //送信データ クリア
            mSend.Clear();

            //アドレス設定
            bytes = Encoding.ASCII.GetBytes(addr);
            if (bytes.Length < 1) return (false);
            if (AsciiCheck(bytes[0]) == false) return (false); mSend.Add(bytes[0]); sum += bytes[0];

            //コンマ
            mSend.Add(COMMA); sum += COMMA;

            //コマンド設定
            bytes = Encoding.ASCII.GetBytes(cmd);
            if (bytes.Length < 3) return (false);
            if (AsciiCheck(bytes[0]) == false) return (false); mSend.Add(bytes[0]); sum += bytes[0];
            if (AsciiCheck(bytes[1]) == false) return (false); mSend.Add(bytes[1]); sum += bytes[1];
            if (AsciiCheck(bytes[2]) == false) return (false); mSend.Add(bytes[2]); sum += bytes[2];

            //コンマ
            mSend.Add(COMMA); sum += COMMA;

            //データ設定
            bytes = Encoding.ASCII.GetBytes(data);
            if (bytes.Length < 8) return (false);
            if (AsciiCheck(bytes[0]) == false) return (false); mSend.Add(bytes[0]); sum += bytes[0];
            if (AsciiCheck(bytes[1]) == false) return (false); mSend.Add(bytes[1]); sum += bytes[1];
            if (AsciiCheck(bytes[2]) == false) return (false); mSend.Add(bytes[2]); sum += bytes[2];
            if (AsciiCheck(bytes[3]) == false) return (false); mSend.Add(bytes[3]); sum += bytes[3];
            if (AsciiCheck(bytes[4]) == false) return (false); mSend.Add(bytes[4]); sum += bytes[4];
            if (AsciiCheck(bytes[5]) == false) return (false); mSend.Add(bytes[5]); sum += bytes[5];
            if (AsciiCheck(bytes[6]) == false) return (false); mSend.Add(bytes[6]); sum += bytes[6];
            if (AsciiCheck(bytes[7]) == false) return (false); mSend.Add(bytes[7]); sum += bytes[7];

            //コンマ
            mSend.Add(COMMA); sum += COMMA;

            //チェックサム 計算と設定
            ckData = (0xff - (((sum >> 8) & 0xff) + (sum & 0xff)));
            bytes = Encoding.ASCII.GetBytes(ckData.ToString("X2"));
            mSend.Add(bytes[0]);
            mSend.Add(bytes[1]);

            //デリミタ
            mSend.Add(DELIMITER);

            //送信データバイト数チェック
            if (mSend.Count != SEND_NUM)
            {
                return (false);
            }

            return (true);
        }

        /// <summary>
        /// 受信データ解析
        /// 作成日:2020/09/30 作成者:murata
        /// </summary>
        private bool RecvDataAnalyze(out string boardno,
                                     out string cmd,
                                     out string data,
                                     out string res)
        {
            int i;
            long sum = 0;
            long ckData;
            byte[] bytes;

            boardno = "";
            cmd = "";
            data = "";
            res = "";

            //受信カウント数チェック
            if (mRecv.Count != RECV_NUM)
            {
                Console.WriteLine("受信データエラー：データ個数");
                return (false);
            }

            //チェックサム チェック
            sum = 0;
            for (i = 0; i < 17; i++)
            {
                sum += mRecv[i];
            }
            ckData = (0xff - (((sum >> 8) & 0xff) + (sum & 0xff)));

            bytes = Encoding.ASCII.GetBytes(ckData.ToString("X2"));

            if (mRecv[17] != bytes[0] || mRecv[18] != bytes[1])
            {
                Console.WriteLine("受信データエラー：チェックサム");
                return (false);
            }

            //応答コード取得
            res = ((char)(mRecv[15])).ToString();

            //データ取得
            data = "";
            data += ((char)(mRecv[6])).ToString();         //データ　1文字目
            data += ((char)(mRecv[7])).ToString();         //データ　2文字目
            data += ((char)(mRecv[8])).ToString();         //データ　3文字目
            data += ((char)(mRecv[9])).ToString();         //データ　4文字目
            data += ((char)(mRecv[10])).ToString();        //データ　5文字目
            data += ((char)(mRecv[11])).ToString();        //データ　6文字目
            data += ((char)(mRecv[12])).ToString();        //データ　7文字目
            data += ((char)(mRecv[13])).ToString();        //データ　8文字目

            //制御コマンド取得
            cmd = "";
            cmd += ((char)(mRecv[2])).ToString();         //データ  1文字目
            cmd += ((char)(mRecv[3])).ToString();         //データ　2文字目
            cmd += ((char)(mRecv[4])).ToString();         //データ　3文字目

            //基板番号
            boardno = ((char)(mRecv[0])).ToString();      //データ  1文字目

            //タイムアウトチェック
            if (cmd == "TMO")
            {
                Console.WriteLine("タイムアウト受信");
            }

            return (true);
        }

        /// <summary>
        /// 送信処理
        /// 作成日:2020/09/30 作成者:murata
        /// </summary>
        private bool SendData()
        {
            bool ret = false;

            try
            {
                mPort.Write(mSend.ToArray(), 0, mSend.Count);
                ret = true;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                ret = false;
            }
            return ret;
        }

        /// <summary>
        /// 受信処理
        /// 作成日:2020/09/30 作成者:murata
        /// </summary>
        private bool RecvData()
        {
            byte bData;
            bool ret = false;
            Stopwatch sw = new Stopwatch();

            mRecv.Clear();

            sw.Start();

            try
            {
                while (true)
                {
                    //タイムアウトエラーチェック
                    if (sw.ElapsedMilliseconds > mTimeOut)
                    {
                        ret = false;
                        break;
                    }

                    bData = (byte)mPort.ReadByte();
                    mRecv.Add(bData);

                    if (bData == DELIMITER)
                    {
                        ret = true;
                        break;
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                ret = false;
            }
            return ret;
        }
        #endregion
    }
}
