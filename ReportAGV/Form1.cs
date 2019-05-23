using GemBox.Spreadsheet;
using GemBox.Spreadsheet.WinFormsUtilities;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Forms;


namespace ReportAGV
{
    public partial class Form1 : Form
    {
        public enum StatusOrderResponseCode
        {
            NODETERMINED = 0,
            ORDER_STATUS_RESPONSE_SUCCESS = 200,
            ORDER_STATUS_RESPONSE_ERROR_DATA = 201,
            ORDER_STATUS_RESPONSE_NOACCEPTED = 202,
            ORDER_STATUS_DOOR_BUSY = 203,
            PENDING = 300,
            DELIVERING = 301,
            FINISHED = 302,
            ROBOT_ERROR = 303,
            NO_BUFFER_DATA = 304,
            CHANGED_FORKLIFT = 305,
            DESTROYED = 306,


        }


        public enum TyeRequest
        {
            TYPEREQUEST_FORLIFT_TO_BUFFER = 1,
            TYPEREQUEST_BUFFER_TO_MACHINE = 2,
            TYPEREQUEST_BUFFER_TO_RETURN = 3,
            TYPEREQUEST_MACHINE_TO_RETURN = 4,
            TYPEREQUEST_RETURN_TO_GATE = 5,
            TYPEREQUEST_CLEAR = 6,
            TYPEREQUEST_OPEN_FRONTDOOR_DELIVERY_PALLET = 7,
            TYPEREQUEST_CLOSE_FRONTDOOR_DELIVERY_PALLET = 8,
            TYPEREQUEST_OPEN_FRONTDOOR_RETURN_PALLET = 9,
            TYPEREQUEST_CLOSE_FRONTDOOR_RETURN_PALLET = 10,
            TYPEREQUEST_CLEAR_FORLIFT_TO_BUFFER = 11,
            TYPEREQUEST_FORLIFT_TO_MACHINE = 12, // santao jujeng cap bottle
            TYPEREQUEST_WMS_RETURNPALLET_BUFFER = 13, // santao jujeng cap bottle
            TYPEREQUEST_CHARGE = 14, // santao jujeng cap bottle
            TYPEREQUEST_GOTO_READY = 15, // santao jujeng cap bottle
        }

        public class DataPallet
        {
            public int row;
            public int bay;
            public int directMain;
            public int directSub;
            public int directOut;
            public int line_ord;
            public PalletCtrl palletCtrl;
            public Pose linePos;
        }
        public class Pose
        {
            public Pose(Point p, double Angle) // Angle gốc
            {
                this.Position = p;
                this.AngleW = Angle * Math.PI / 180.0;
                this.Angle = Angle;
            }
            public Pose(int X, int Y, double Angle) // Angle gốc
            {
                this.Position = new Point(X, Y);
                this.AngleW = Angle * Math.PI / 180.0;
                this.Angle = Angle;
            }
            public Pose() { }
            public void Destroy() // hủy vị trí robot để robot khác có thể làm việc trong quá trình detect
            {
                //this.Position = new Point (-1000, -1000);
                //this.AngleW = 0;
            }
            public Point Position { get; set; }
            public double VFbx { get; set; }
            public double VFby { get; set; }
            public double VFbw { get; set; }

            public double VCtrlx { get; set; }
            public double VCtrly { get; set; }
            public double VCtrlw { get; set; }

            public double AngleW { get; set; } // radian
            public double Angle { get; set; } // radian
        }
        public enum PalletCtrl
        {
            Pallet_CTRL_DOWN = 0,
            Pallet_CTRL_UP = 1

        }
        public class OrderItem
        {
            public OrderItem() { }
            public String userName { get; set; }
            public String robot { get; set; }
            public StatusOrderResponseCode status { get; set; }
            private String OrderId { get; set; }
            public int planId { get; set; }
            public int deviceId;
            public String productDetailName { get; set; }
            public int productId { get; set; }
            public int productDetailId { get; set; }


            public TyeRequest typeReq { get; set; } // FL: ForkLift// BM: BUFFER MACHINE // PR: Pallet return
            public String activeDate;
            public String dateTime { get; set; }

            public int timeWorkId;
            public String palletStatus;
            public int palletId;
            public int updUsrId;
            public int lengthPallet;
            public String dataRequest;
            // public bool status = false; // chua hoan thanh
            public DataPallet palletAtMachine;

            public int bufferId;
            public int palletAmount;
            public DateTime startTimeProcedure = new DateTime();
            public DateTime endTimeProcedure = new DateTime();
            public double totalTimeProcedure { get; set; }
            public bool onAssiged = false;

        }


        public Form1()
        {
            InitializeComponent();
            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.Filter = "All Files (*.*)|*.*";
            if (fileDialog.ShowDialog() == DialogResult.OK)
            {
                string path = fileDialog.FileName;
                String[] lines = File.ReadAllLines(path);
                List<OrderItem> orderlist = new List<OrderItem>();
                foreach (String strj in lines)
                {
                    OrderItem order = JsonConvert.DeserializeObject<OrderItem>(strj);
                    orderlist.Add(order);
                }
                dataGridViewReport.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
                dataGridViewReport.DataSource = orderlist;

            }
           
        }

        private void Button2_Click(object sender, EventArgs e)
        {
          
            var saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "XLS files (*.xls)|*.xls|XLT files (*.xlt)|*.xlt|XLSX files (*.xlsx)|*.xlsx|XLSM files (*.xlsm)|*.xlsm|XLTX (*.xltx)|*.xltx|XLTM (*.xltm)|*.xltm|ODS (*.ods)|*.ods|OTS (*.ots)|*.ots|CSV (*.csv)|*.csv|TSV (*.tsv)|*.tsv|HTML (*.html)|*.html|MHTML (.mhtml)|*.mhtml|PDF (*.pdf)|*.pdf|XPS (*.xps)|*.xps|BMP (*.bmp)|*.bmp|GIF (*.gif)|*.gif|JPEG (*.jpg)|*.jpg|PNG (*.png)|*.png|TIFF (*.tif)|*.tif|WMP (*.wdp)|*.wdp";
            saveFileDialog.FilterIndex = 3;

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                var workbook = new ExcelFile();
                var worksheet = workbook.Worksheets.Add("Sheet1");

                // From DataGridView to ExcelFile.
                DataGridViewConverter.ImportFromDataGridView(worksheet, this.dataGridViewReport, new ImportFromDataGridViewOptions() { ColumnHeaders = true });

                workbook.Save(saveFileDialog.FileName);
            }
        }

        private void DataGridViewReport_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
           
        }
    }
    
}
