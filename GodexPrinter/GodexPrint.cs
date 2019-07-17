using System;
using System.Collections;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using EzioDLL;


namespace AddIn
{
    public enum PaperMode
    {
        GapLabel = 0,
        PlainPaperLabel = 1,
        BlackMarkLabel = 2
    }

    public enum PortType
    {
        LPT1 = 0,
        LPT2 = 1,
        LPT3 = 2,
        COM1 = 3,
        COM2 = 4,
        COM3 = 5,        
        COM4 = 6,
        COM5 = 7,
        COM6 = 8,
        COM7 = 9,
        COM8 = 10,
        USB = 11,
        NET = 12
    }

    [ProgId("AddIn.GodexPrinter")]
    [ComVisible(true), Guid("ACFC56D7-321A-49AE-A4B0-FD4DB9253F11")]
    public class GodexPrinter : ComponentBase
    {
        static Hashtable PrinterPort = new Hashtable();
        private static void GetPrinterPort()
        {
            if (PrinterPort.Count > 0) return;
            PrinterPort[PortType.LPT1] = 0;
            PrinterPort[PortType.LPT2] = 5;
            PrinterPort[PortType.COM1] = 1;
            PrinterPort[PortType.COM2] = 2;
            PrinterPort[PortType.COM3] = 3;
            PrinterPort[PortType.COM4] = 4;
            PrinterPort[PortType.USB] = 6;
        }

        private PortType numberPort;
        [Alias("НомерПорта")]
        public PortType NumberPort
        {
            get { return numberPort; }
            set
            {
                try
                {
                    numberPort = (PortType)value;
                }
                catch
                {
                    numberPort = 0;
                }                
            }
        }
        private int speed;
        [Alias("Скорость")]
        public int Speed
        {
            get { return speed; }
            set { speed = value; }
        }
        private int darkness;
        [Alias("Яркость")]
        public int Darkness
        {
            get { return darkness; }
            set
            {
                if (value < 20 && value >=0)
                    darkness = value;
                else
                    darkness = 10;
            }
        }
        private int stripper;
        [Alias("Отделитель")]
        public int Stripper
        {
            get { return stripper; }
            set
            {
                if (value < 2 && value >=0)
                    stripper = value;
                else
                    stripper = 0;
            }
        }
        private int cutter;
        [Alias("Резак")]
        public int Cutter
        {
            get { return cutter; }
            set
            {
                if (value < 32768 && value >= 0)
                    cutter = value;
                else
                    cutter = 0;
            }
        }
        private int labelHeight;
        [Alias("ВысотаЭтикетки")]
        public int LabelHeight
        {
            get { return labelHeight; }
            set
            {
                if (value < 10000 && value >= 0)
                    labelHeight = value;
                else
                    labelHeight = 25;
            }
        }
        private int labelWidth;
        [Alias("ШиринаЭтикетки")]
        public int LabelWidth
        {
            get { return labelWidth; }
            set
            {
                if (value < 10000 && value >= 0)
                    labelWidth = value;
                else
                    labelWidth = 43;
            }
        }
        private int labelLeftMargin;
        [Alias("ОтступСлеваЭтикетки")]
        public int LabelLeftMargin
        {
            get { return labelLeftMargin; }
            set
            {
                if (value < 400 && value >= 0)
                    labelLeftMargin = value;
                else
                    labelLeftMargin = 0;
            }
        }
        private int labelTopMargin;
        [Alias("ОтступСверхуЭтикетки")]
        public int LabelTopMargin
        {
            get { return labelTopMargin; }
            set
            {
                if (value > -101 && value < 101)
                    labelTopMargin = value;
                else
                    value = 0;
            }
        }
        // Это толщина тонкой линии ШК, а не то, что вы подумали.
        private int barcodeWidth;
        [Alias("ШиринаШтрихкода")]
        public int BarcodeWidth
        {
            get { return barcodeWidth; }
            set
            {
                if (value < 11 && value >= 0) barcodeWidth = value;
                else barcodeWidth = 2;
                setWide();
            }
        }
        private int barcodeReadable;
        [Alias("ЧитаемыйШтрихкод")]
        public int BarcodeReadable
        {
            get { return barcodeReadable; }
            set { barcodeReadable = value; }
        }
        // это отношение толстой линии к тонкой
        private int barcodeNarrow;
        private int wide;
        [Alias("ПлотностьШтрихкода")]
        public int BarcodeNarrow
        {
            get { return barcodeNarrow; }
            set
            {
                if (value < 2) barcodeNarrow = value;
                else barcodeNarrow = 0;
                setWide();
            }
        }
        // собственно тут считаем толщину линии, необходимую для комманды форматирования
        private void setWide()
        { 
            if (barcodeNarrow == 2)
               wide = barcodeWidth* 5 / 2;
            else if (barcodeNarrow == 1)
               wide = barcodeWidth* 3;
            else
               wide = barcodeWidth* 2;
        }
        private PaperMode labelType;
        [Alias("ТипЭтикетки")]
        public int LabelType
        {
            get { return (int)labelType; }
            set
            {
                if (value >= 0 && value < 3)
                    labelType = (PaperMode)value;
                else
                    value = 0;
            }
        }
        private int gapLength;
        [Alias("ДлинаЗазора")]
        public int GapLength
        {
            get { return gapLength; }
            set
            {
                if (value >= 0 && value < 14)
                    gapLength = value;
                else
                    gapLength = 3;
            }
        }
        private int feedLength;
        [Alias("ДлинаПрогона")]
        public int FeedLength
        {
            get { return feedLength; }
            set
            {
                if (value >= 0 && value < 10000)
                    feedLength = value;
                else
                    feedLength = 3;
            }
        }
        private int blackMarkLength;
        [Alias("ДлинаТемнойМетки")]
        public int BlackMarkLength
        {
            get { return blackMarkLength; }
            set
            {
                if (value >= 0 && value < 10000)
                    blackMarkLength = value;
                else
                    blackMarkLength = 3;
            }
        }
        private int blackMarkPosition;
        [Alias("СмещениеОтТемнойМетки")]
        public int BlackMarkPosition
        {
            get { return blackMarkLength; }
            set
            {
                if (value >= 0 && value < 10000)
                    blackMarkLength = value;
                else
                    blackMarkLength = 3;
            }
        }
        private int stopPosition;
        [Alias("ПозицияОстановки")]
        public int StopPosition
        {
            get { return stopPosition; }
            set
            {
                if (value >= 0 && value < 41)
                    stopPosition = value;
                else
                    stopPosition = 0;
            }
        }
        private int labelRotate;
        [Alias("ПеревернутьЭтикетку")]
        public int LabelRotate
        {
            get { return labelRotate; }
            set
            {
                if (value >= 0 && value <= 1)
                    labelRotate = value;
                else
                    labelRotate = 0;
            }
        }
        private int termotransfer;
        [Alias("Термотрансферный")]
        public int Termotransfer
        {
            get { return termotransfer; }
            set
            {
                if (value >= 0 && value <= 1)
                    termotransfer = value;
                else
                    termotransfer = 0;
            }
        }
        private string ip;
        [Alias("СетевойАдрес")]
        public string IP
        {
            get { return ip; }
            set { ip = value; }
        }
        private int port;
        [Alias("СетевойПорт")]
        public int Port
        {
            get { return port; }
            set { port = value; }
        }
        
        // методы
        [Alias("НоваяЭтикетка")]
        public void StartLabel()
        {

            AddFormatLabel("^L");
        }
        [Alias("ВывестиТекст")]
        public void PutText(int x, int y, int stretch_x, int stretch_y, int gap, int rotate, string font, string str)
        {
            string param = String.Join(",", new string[] { x.ToString(), y.ToString(), stretch_x.ToString(), stretch_y.ToString(), gap.ToString(), rotate.ToString(), str });
            AddFormatLabel("A" + font + "," + param);
        }
        [Alias("ВывестиЗагруженныйТекст")]
        public void PutloadedText(int x, int y, int stretch_x, int stretch_y, int gap, int rotate, string font, string str)
        {
            string param = String.Join(",", new string[] { x.ToString(), y.ToString(), stretch_x.ToString(), stretch_y.ToString(), gap.ToString(), rotate.ToString(), str });
            AddFormatLabel("V" + font + "," + param);
        }
        [Alias("СформироватьШтрихкод")]
        public void PutBarcode(string type, int x, int y, int height, int rotate, string str)
        {
            
            string param = string.Join(",", new string[] { type, x.ToString(), y.ToString(), barcodeWidth.ToString(), wide.ToString(), height.ToString(), rotate.ToString(), barcodeReadable.ToString(), str });
            AddFormatLabel("B" + param);
        }
        [Alias("ВывестиПрямоугольник")]
        public void PutRectangle(int x, int y, int x1, int y1, int width_y, int width_x)
        {

        }
        [Alias("ИнвертироватьПрямоугольник")]
        public void InverseRectangle(int x, int y, int x1, int y1)
        {

        }
        [Alias("Печатать")]
        public int Print(int count)
        {
            switch (numberPort)
            {
                case PortType.LPT1:
                case PortType.LPT2:
                case PortType.LPT3:
                    Open(numberPort);
                    break;
                case PortType.COM1:
                case PortType.COM2:
                case PortType.COM3:
                case PortType.COM4:
                case PortType.COM5:
                case PortType.COM6:
                case PortType.COM7:
                case PortType.COM8:
                    Open(numberPort);
                    break;
                case PortType.USB:
                    GetPrinterPort();
                    Open(PrinterPort[PortType.USB].ToString());
                    break;
                case PortType.NET:
                    if (!EZioApi.OpenNet(ip, port.ToString()))
                    {
                        showMеssages = true;
                        errorMessage = "Не удалось соединиться с принтером.";
                        throw new COMException(errorMessage, (int)HRESULT.S_FALSE);
                        return 1;
                    }                    
                    break;
            }
            int err = 0;
            if (count == 0)
            {
                string str = String.Join("\r\n", printBuffer.ToArray());
                err = EZioApi.sendcommand(str);                
            }
            else
            {
                SetupExec();
                AddSetup("^P" + count.ToString());
                AddFormatLabel("E");
                err = EZioApi.sendcommand(setupData + "\r\n" + labelFormat);               
            }
            printBuffer.Clear();
            ClearSetup();
            ClearFormatLabel();
            EZioApi.closeport();
            return err;
        }
        
        private List<String> printBuffer = new List<string>();
        [Alias("ПечататьВБуфер")]
        public int PrintToBuff(int count)
        {
            SetupExec();
            AddSetup("^P" + count.ToString());
            AddFormatLabel("E");
            printBuffer.Add(setupData + "\r\n" + labelFormat);
            ClearSetup();
            ClearFormatLabel();
            return 0;
        }
        [Alias("ВывестиКартинку")]
        public int PutImage(int x, int y, string imageName)
        {
            return 0;
        }

        private string setupData;
        private void AddSetup(string data)
        {
            if (setupData == null) { setupData = data; return; }
            setupData = String.Join("\r\n", new string[] { setupData, data });
        }
        private string labelFormat;
        private void AddFormatLabel(string data)
        {
            if (labelFormat == null) { labelFormat = data; return; }
            labelFormat = string.Join("\r\n", new string[] { labelFormat, data });
        }

        private void SetupExec()
        {
            ClearSetup();
            if (labelType == PaperMode.GapLabel)
                AddSetup("^Q" + labelHeight.ToString() + "," + gapLength.ToString());
            else if (labelType == PaperMode.PlainPaperLabel)
                AddSetup("^Q" + labelHeight.ToString() + ",0," + gapLength.ToString());
            else
                AddSetup("^Q" + labelHeight.ToString() + blackMarkLength.ToString() + blackMarkPosition.ToString() + "0+");
            AddSetup("^W" + labelWidth.ToString());
            AddSetup("^H" + darkness.ToString());
            AddSetup("^S" + speed.ToString());
            if (termotransfer == 0)
                AddSetup("^AD");
            else
                AddSetup("^AT");
            AddSetup("^R" + labelLeftMargin.ToString());
            if (labelTopMargin >= 0)
                AddSetup("~Q+" + labelTopMargin.ToString());
            else
                AddSetup("~Q" + labelTopMargin.ToString());
            AddSetup("^O" + stripper.ToString());
            AddSetup("^D" + cutter.ToString());
            AddSetup("^E" + stopPosition.ToString());
        }
        private void ClearSetup()
        {
            setupData = "";
        }
        private void ClearFormatLabel()
        {
            labelFormat = "";
        }
        [Alias("ОписаниеОшибки")]
        public string GetErrorMessage()
        {
            return errorMessage;
        }

        public GodexPrinter()
        {
            numberPort = 0;
            speed = 2;
            darkness = 10;
            stripper = 0;
            cutter = 0;
            labelHeight = 25;
            labelWidth = 43;
            labelLeftMargin = 0;
            labelTopMargin = 0;
            BarcodeWidth = 2;
            barcodeReadable = 1;
            BarcodeNarrow = 0;
            labelType = 0;
            gapLength = 3;
            feedLength = 3;
            blackMarkLength = 0;
            blackMarkPosition = 0;
            stopPosition = 0;
            labelRotate = 0;
            Termotransfer = 1;
        }

        //[Alias("Печать")]
        //public void Print()
        //{
        //    byte[] version = new byte[100];
        //    Array.Clear(version, 0, version.Length);
        //    EZioApi.GetDllVersion(version);
        //    MessageBox.Show(System.Text.Encoding.ASCII.GetString(version));
        //}
        [Alias("ОткрытьСетевойПорт")]
        public void OpenNet(string IP, int port)
        {
            EZioApi.OpenNet(IP, port.ToString());
        }
        [Alias("ЗакрытьПорт")]
        public void Close()
        {
            EZioApi.closeport();
        }
        [Alias("ОтправитьКоманду")]
        public int Send(string Cmd)
        {
            return EZioApi.sendcommand(Cmd);
        }
        //[Alias("Скорость")]
        //public void Speed(int nSpeed)
        //{
        //    EZioApi.sendcommand("^S" + nSpeed.ToString());
        //}
        //[Alias("Яркость")]
        //public void Darkness(int nDark)
        //{
        //    EZioApi.sendcommand("^H" + nDark.ToString());
        //}
        [Alias("ОткрытьПорт")]
        public void Open(PortType mPortType)
        {
            EZioApi.openport(((int)mPortType).ToString());
        }

        public void Open(string PortName)
        {
              EZioApi.openport(PortName);
        }
        public void Open(string strIP, int nPort)
        {
            EZioApi.OpenNet(strIP, nPort.ToString());
        }

        public int SetBaudrate(int nBaud)
        {
            return EZioApi.setbaudrate(nBaud);
        }
        [Alias("ПоказыватьСообщения")]
        public bool ShowMessages
        {
            get { return showMеssages; }
            set { showMеssages = value; }                
        }       
    }
}
