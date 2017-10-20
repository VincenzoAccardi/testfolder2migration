using System;
using System.Configuration;
using System.Data.OleDb;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace WinConsultingOrder
{
    public partial class ConsultingOrder : Form
    {
        private IntPtr ptrHook;
        private LowLevelKeyboardProc objKeyboardProcess;

        private System.Threading.Timer changeHour;
        private System.Threading.Timer requestOrder;
        private System.Threading.Timer responseOrder;
        private string numOrder = string.Empty;

        public ConsultingOrder()
        {
            ProcessModule objCurrentModule = Process.GetCurrentProcess().MainModule; //Get Current Module
            objKeyboardProcess = new LowLevelKeyboardProc(captureKey); //Assign callback function each time keyboard process
            ptrHook = SetWindowsHookEx(13, objKeyboardProcess, GetModuleHandle(objCurrentModule.ModuleName), 0); //Setting Hook of Keyboard Process for current module
            InitializeComponent();

            var result = AddFontResource(@"../../Fonts/ufonxts.com_cocon-bold-opentype.ttf");
            var error = Marshal.GetLastWin32Error();
        }
        private void ConsultingOrder_Load(object sender, EventArgs e)
        {
            var size = this.Size;

            try
            {
                var LabelSite = ConfigurationManager.AppSettings["LabelSite"].ToString();
                var sizeLabel1 = Convert.ToInt32(ConfigurationManager.AppSettings["sizeLabel1"]);
                var sizeLabel2 = Convert.ToInt32(ConfigurationManager.AppSettings["sizeLabel2"]);
                var sizeLabel3 = Convert.ToInt32(ConfigurationManager.AppSettings["sizeLabel3"]);

                this.label1.Location = new Point(0, (size.Height / 2));
                this.label1.Size = new Size(this.Size.Width, 100);
                this.label1.Font = new Font("Cocon", sizeLabel1, FontStyle.Bold);
                this.label1.ForeColor = Color.FromArgb(0, 85, 48);

                this.label2.Location = new Point(0, this.label1.Location.Y + this.label1.Size.Height);
                this.label2.Size = new Size(this.Size.Width, 100);
                this.label2.Font = new Font("Cocon", sizeLabel2, FontStyle.Bold);
                this.label2.ForeColor = Color.FromArgb(0, 85, 48);

                this.label4.Location = new Point(0, this.label2.Location.Y + this.label2.Size.Height + 30);
                this.label4.Size = new Size(this.Size.Width, 100);
                this.label4.Font = new Font("Cocon", sizeLabel3, FontStyle.Bold);
                this.label4.ForeColor = Color.FromArgb(0, 85, 48);
                this.label4.Text = LabelSite;

            }
            catch (Exception ettt)
            {

            }
        }
        private void backgroundWorker1_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
        {
            changeHour = new System.Threading.Timer(ChangeHour, null, 0, 1000);
            requestOrder = new System.Threading.Timer(GetRequest, null, 0, 10000);
            responseOrder = new System.Threading.Timer(GetResponse, null, 0, 10000);
        }
        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);

            backgroundWorker1.RunWorkerAsync();
        }

        private void GetRequest(object state)
        {
            try
            {
                var connString = System.Configuration.ConfigurationManager.ConnectionStrings["localDB"].ConnectionString;
                var existRequest = false;

                using (var conn = new OleDbConnection(connString))
                {
                    conn.Open();
                    var qry = new StringBuilder();

                    qry.AppendLine("SELECT [TPPosDB].[dbo].[ecommerce_request].[transaction_id],");
                    qry.AppendLine("    count([TPPosDB].[dbo].[ecommerce_request].[transaction_id]) count  ");
                    qry.AppendLine(" FROM [TPPosDB].[dbo].[ecommerce_request] WITH(NOLOCK)");
                    qry.AppendLine(" WHERE  ");
                    qry.AppendLine("         NOT EXISTS ");
                    qry.AppendLine("(SELECT * ");
                    qry.AppendLine("    FROM [TPPosDB].[dbo].[ecommerce_response] WITH(NOLOCK)");
                    qry.AppendLine("    WHERE [TPPosDB].[dbo].[ecommerce_request].[transaction_id] = [TPPosDB].[dbo].[ecommerce_response].[transaction_id]");
                    qry.AppendLine("    AND [TPPosDB].[dbo].[ecommerce_response].transaction_id = ?)");
                    qry.AppendLine(" GROUP BY");
                    qry.AppendLine("    [TPPosDB].[dbo].[ecommerce_request].[transaction_id]");

                    using (var cmd = new OleDbCommand(qry.ToString(), conn))
                    {
                        cmd.Parameters.Add(new OleDbParameter("@transaction_id", numOrder));
                        using (var reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                existRequest = true;
                                numOrder = reader["transaction_id"].ToString();
                            }
                        }
                    }
                }

                if (existRequest)
                {
                    var lblRunning = ConfigurationManager.AppSettings["LabelOrderRunning"].ToString();

                    label4.Invoke(new Action(() => label4.Visible = false));
                    label1.Invoke(new Action(() => label1.Text = "ORDINE N. " + numOrder));
                    label2.Invoke(new Action(() => label2.Text = lblRunning));
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void GetResponse(object state)
        {
            try
            {
                if (!string.IsNullOrEmpty(numOrder))
                {
                    var existResponse = false;
                    var tranError = false;
                    var connString = System.Configuration.ConfigurationManager.ConnectionStrings["localDB"].ConnectionString;

                    using (var conn = new OleDbConnection(connString))
                    {
                        conn.Open();
                        var qry = new StringBuilder();

                        qry.AppendLine("SELECT [TPPosDB].[dbo].[ecommerce_response].[transaction_id],");
                        qry.AppendLine("    row_type  ");
                        qry.AppendLine(" FROM [TPPosDB].[dbo].[ecommerce_response] WITH(NOLOCK)");
                        qry.AppendLine(" WHERE  ");
                        qry.AppendLine("    [TPPosDB].[dbo].[ecommerce_response].transaction_id = ? AND");
                        qry.AppendLine("         NOT EXISTS ");
                        qry.AppendLine("(SELECT * ");
                        qry.AppendLine("    FROM [TPPosDB].[dbo].[ecommerce_request] WITH(NOLOCK)");
                        qry.AppendLine("    WHERE [TPPosDB].[dbo].[ecommerce_request].[transaction_id] = [TPPosDB].[dbo].[ecommerce_response].[transaction_id]");
                        qry.AppendLine("    AND [TPPosDB].[dbo].[ecommerce_request].transaction_id = ?)");
                        qry.AppendLine(" GROUP BY");
                        qry.AppendLine("    [TPPosDB].[dbo].[ecommerce_response].[transaction_id],");
                        qry.AppendLine("    row_type");


                        using (var cmd = new OleDbCommand(qry.ToString(), conn))
                        {
                            cmd.Parameters.Add(new OleDbParameter("@transaction_id_response", numOrder));
                            cmd.Parameters.Add(new OleDbParameter("@transaction_id", numOrder));

                            using (var reader = cmd.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    if (reader["row_type"].ToString() == "E")
                                        tranError = true;
                                    else
                                        existResponse = true;

                                }
                            }
                        }
                    }

                    var timeoutError = Convert.ToInt32(ConfigurationManager.AppSettings["timeOutError"].ToString());
                    var timeoutRegards = Convert.ToInt32(ConfigurationManager.AppSettings["timeOutRegards"].ToString());
                    var lblEndOrder = ConfigurationManager.AppSettings["LabelOrderRunning"].ToString();
                    var lblOrderError = ConfigurationManager.AppSettings["LabelError"].ToString();

                    if (tranError)
                    {
                        label4.Invoke(new Action(() => label4.Visible = false));
                        numOrder = string.Empty;
                        label2.Invoke(new Action(() => label2.ForeColor = Color.Red));
                        label2.Invoke(new Action(() => label2.Text = lblOrderError));
                        Thread.Sleep(timeoutError);
                        label2.Invoke(new Action(() => label2.ForeColor = Color.FromArgb(0, 85, 48)));
                        label1.Invoke(new Action(() => label1.Text = "SCEGLI IL MODO PIU'"));
                        label2.Invoke(new Action(() => label2.Text = "COMODO DI FARE LA SPESA"));
                        label4.Invoke(new Action(() => label4.Visible = true));
                    }
                    else if (existResponse)
                    {
                        label4.Invoke(new Action(() => label4.Visible = false));
                        numOrder = string.Empty;
                        label2.Invoke(new Action(() => label2.Text = lblEndOrder));
                        Thread.Sleep(timeoutError);
                        label2.Invoke(new Action(() => label2.Text = "GRAZIE E ARRIVEDERCI"));
                        Thread.Sleep(timeoutRegards);
                        label1.Invoke(new Action(() => label1.Text = "SCEGLI IL MODO PIU'"));
                        label2.Invoke(new Action(() => label2.Text = "COMODO DI FARE LA SPESA"));
                        label4.Invoke(new Action(() => label4.Visible = true));
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void ChangeHour(object state)
        {
            label3.Invoke((MethodInvoker)(() =>
            {
                label3.Text = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");
            }));

        }


        #region "BlockWinKey"

        private IntPtr captureKey(int nCode, IntPtr wp, IntPtr lp)
        {
            if (nCode >= 0)
            {
                KBDLLHOOKSTRUCT objKeyInfo = (KBDLLHOOKSTRUCT)Marshal.PtrToStructure(lp, typeof(KBDLLHOOKSTRUCT));

                if (objKeyInfo.key == Keys.RWin || objKeyInfo.key == Keys.LWin) // Disabling Windows keys
                {
                    return (IntPtr)1;
                }
            }
            return CallNextHookEx(ptrHook, nCode, wp, lp);
        }

        [DllImport("gdi32.dll", EntryPoint = "AddFontResourceW", SetLastError = true)]
        public static extern int AddFontResource([In][MarshalAs(UnmanagedType.LPWStr)]
                                             string lpFileName);

        [StructLayout(LayoutKind.Sequential)]
        private struct KBDLLHOOKSTRUCT
        {
            public Keys key;
            public int scanCode;
            public int flags;
            public int time;
            public IntPtr extra;

            public int vkCode;
            int dwExtraInfo;
        }

        //System level functions to be used for hook and unhook keyboard input
        private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);
        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int id, LowLevelKeyboardProc callback, IntPtr hMod, uint dwThreadId);
        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern bool UnhookWindowsHookEx(IntPtr hook);
        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr CallNextHookEx(IntPtr hook, int nCode, IntPtr wp, IntPtr lp);
        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr GetModuleHandle(string name);
        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern short GetAsyncKeyState(Keys key);
        #endregion

    }

}
