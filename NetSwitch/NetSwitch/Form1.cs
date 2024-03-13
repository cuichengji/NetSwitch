using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using System.Management;
using System.IO;
using System.Drawing.Imaging;

namespace NetSwitch
{
    public partial class Form1 : Form
    {

        private bool printedInterfacesStartup = false;
        private bool toogleConfigActive = false;
        private bool allowshowdisplay = false;
        // private const string deviceDisabled = "22";
        private const string menuExitText = "Exit";
        private const string menuToggleText = "Toggle";
        private const string toggleFileConfig = "toggle.conf";
        private List<string> toggleToDeactivate = new List<string>();
        private string toggleToActivate = "";
        private List<ManagementObject> devList = new List<ManagementObject>();
        private ToolStripMenuItem menuExit = new ToolStripMenuItem(menuExitText);

        private const string DEVICE_DISABLED = "22";
        private const int ICON_W = 64;
        private const int ICON_H = 64;
        private const string ICON_GIF_FILENAME = @"setting{0}";
        //private const string menuShowText = "Show";
        //private ToolStripMenuItem menuShow = new ToolStripMenuItem(menuShowText);

        //タスクトレイにアニメで表示するアイコン
        private Icon[] tasktrayIcons = new Icon[0] { };
        //アニメで現在表示しているアイコンのインデックス
        private int currentTasktrayIconIndex = 0;


        public Form1()
        {
            InitializeComponent();

            // parse toggle options, if configured:
            // loads the content of the file. Each line must contain the exact name of the interface. Only one of the interfaces is activated each toggle.
            if (File.Exists(toggleFileConfig))
            {
                List<string> interfaces = new List<string>();
                string[] lines = File.ReadAllLines(toggleFileConfig);
                foreach(string line in lines)
                {
                    if (line.Count() > 0)
                        interfaces.Add(line);
                }
                if(interfaces.Count() > 1)
                {
                    Console.WriteLine("toggle configuration found");
                    toogleConfigActive = true;
                    toggleToActivate = interfaces.ElementAt(0);
                    interfaces.Remove(toggleToActivate);
                    toggleToDeactivate = interfaces;
                }
            }
            //this.Shown += new EventHandler(this.ShownStatus);
            //this.Closing += new CancelEventHandler(this.CloseForm);
            //this.MinimumSizeChanged += new EventHandler(this.MinimumForm);
            //this.VisibleChanged += new EventHandler(this.ShownStatus);

            InitAnimate();
            this.Visible = false;
            Hide(); // will be only in notification area
            UpdateItems();
        }

        protected override void SetVisibleCore(bool value)
        {
            base.SetVisibleCore(allowshowdisplay ? value : allowshowdisplay);
        }

        private string Uid(ManagementObject mo)
        {
            // create the interface name to be displayed 
            //string sName = obj["Index"].ToString() + ": " + obj["Name"].ToString();
            int ID = CommonFunctions.GetInt(mo["DeviceID"]);

            string CreationClassName = CommonFunctions.GetString(mo["CreationClassName"]);

            string NetConnectionID = CommonFunctions.GetString(mo["NetConnectionID"]);

            string Name = CommonFunctions.GetString(mo["Name"]);

            string MACAddress = CommonFunctions.GetString(mo["MACAddress"]);

            //Console.WriteLine(ID.ToString() + "/" + NetConnectionID + "/" + CreationClassName + "/" + Name + "/" + MACAddress);
            string sName = ID.ToString() + "/" + NetConnectionID + "/" + CreationClassName + "/" + Name + "/" + MACAddress;


            return sName;
        }

        private void UpdateItems()
        {
            UpdateInterfaces();
            UpdateMenu();
        }

        private void InitAnimate()
        {
            //タイマーを無効にしておく（初めはアニメしない）
            this.timer1.Enabled = false;
            //アニメ時は、1秒毎にアイコンを変更する
            this.timer1.Interval = 500;

            //タスクトレイにアニメで表示するアイコンを指定する
            this.makeIconAnime();

        }

        private void UpdateInterfaces()
        {
            string NamespacePath = "\\\\.\\ROOT\\cimv2";
            string ClassName = "Win32_NetworkAdapter";
            
            // this message is useful to copy interface names and configuring toggle 
            if (printedInterfacesStartup == false)
            {
                Console.WriteLine("add two or more of the following names to configure " + toggleFileConfig + " and then activate the Toggle function:");
            }

            devList.Clear();
            ManagementClass oClass = new ManagementClass(NamespacePath + ":" + ClassName);
            foreach (ManagementObject oObject in oClass.GetInstances())
            {
                if (IsNetworkDevice(oObject))
                {
                    devList.Add(oObject);
                    // this message is useful to copy interface names and configuring toggle.
                    if (printedInterfacesStartup == false)
                    {
                        Console.WriteLine(Uid(oObject));
                    }
                }
            }
            // display only once the interface names
            printedInterfacesStartup = true;
        }

        private Boolean IsNetworkDevice(ManagementObject obj)
        {
            return (obj["NetEnabled"] != null);
        }

        private Boolean IsDisabled(ManagementObject obj)
        {
            if (obj["ConfigManagerErrorCode"] != null)
            {
                return (obj["ConfigManagerErrorCode"].ToString() == DEVICE_DISABLED);
            }

            return false;
        }

        private void UpdateMenu()
        {
            // erase all entries in menu
            contextMenuStrip1.Items.Clear();
            ToolStripMenuItem nics = new ToolStripMenuItem("NICs");
            contextMenuStrip1.Items.Add(nics);

            foreach (ManagementObject oObject in devList)
            {
                // set as checked the enabled devices 
                ToolStripMenuItem m = new ToolStripMenuItem(Uid(oObject))
                {
                    Checked = IsDisabled(oObject) ? false : true,
                    CheckOnClick = true
                };
                m.Click += NicItemMenu_ItemClicked;
                // add to menu
                nics.DropDownItems.Add(m);
            }

            // add the toggle option to toggle interfaces defined by user in configuration
            
            ToolStripMenuItem toggleOption = new ToolStripMenuItem(menuToggleText)
            {
                Enabled = toogleConfigActive
            };

            nics.DropDownItems.Add(toggleOption);

            //// Form表示メニュー
            //this.contextMenuStrip1.Items.Add(menuShow);
            // 処理終了メニュー
            this.contextMenuStrip1.Items.Add(menuExit);
        }

        private void ContextMenuStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            // if the item is the exit, bye bye
            if (menuExitText == e.ClickedItem.Text)
            {
                Dispose();
                return;
            }
            //else if (menuShowText == e.ClickedItem.Text)
            //{
            //    Show();
            //}

        }

        private void NicItemMenu_ItemClicked(object sender, EventArgs e)
        {
            // if the item is the exit, bye bye
            if (menuToggleText == sender.ToString())
            {
                ToggleUserInterfaces();
            }

            // タスクトレイアニメ開始
            this.currentTasktrayIconIndex = 0;
            //タイマーが動いている時は止め、止まっているときは動かす
            this.timer1.Enabled = true;
            //アニメを開始したときは、初めのアイコンを表示する
            this.ChangeAnimatedTasktrayIcon();

            // find the interface in or device list 
            foreach (ManagementObject oObject in devList)
            {
                if (sender.ToString() == Uid(oObject))
                {
                    //command in powershell: get-netadapter -ifIndex 7 | Disable-NetAdapter -Confirm:$false
                    string method = IsDisabled(oObject) ? "Enable" : "Disable";
                    oObject.InvokeMethod(method, null);
                }
            }

            // refresh all 
            UpdateItems();

            this.timer1.Enabled = false;
            //アニメが終了した時は、アイコンを元に戻す
            this.notifyIcon1.Icon = this.Icon;

            // 完了メッセージ
            ShowBalloonTip();
        }

        private void ToggleUserInterfaces()
        {
            string aux = toggleToActivate;

            // if there is no element to deactivate, do not toggle
            if (toggleToDeactivate.Count() < 1)
                return;

            toggleToActivate = toggleToDeactivate.ElementAt(0);
            toggleToDeactivate.Remove(toggleToActivate);
            toggleToDeactivate.Add(aux); // add in the end

            Console.WriteLine("activating " + toggleToActivate);
            foreach(string str in toggleToDeactivate)
            {
                Console.WriteLine("deactive " + str);
            }

            foreach (ManagementObject oObject in devList)
            {
                // disable all interfaces and enable only one of the user set (toggle mode)
                foreach (string iface in toggleToDeactivate)
                {
                    if (iface == Uid(oObject))
                    {
                        string method = "Disable";
                        oObject.InvokeMethod(method, null);
                    }
                }

                if (toggleToActivate == Uid(oObject))
                {
                    string method = "Enable";
                    oObject.InvokeMethod(method, null);
                }

                UpdateMenu();
            }
        }

        private void ShowBalloonTip()
        {
            // バルーンヒントを表示する
            //this.notifyIcon1.BalloonTipIcon = ToolTipIcon.Info;
            this.notifyIcon1.BalloonTipTitle = "お知らせタイトル";
            this.notifyIcon1.BalloonTipText = "お知らせメッセージ";
            this.notifyIcon1.ShowBalloonTip(5000);
        }

        private void makeIconAnime()
        {
            this.currentTasktrayIconIndex = 0;
            // サポートするイメージ種類取得
            List<string> imageFiles = new List<string>();
            ImageCodecInfo[] decoders = ImageCodecInfo.GetImageDecoders();
            foreach (ImageCodecInfo ici in decoders)
            {
                if (string.IsNullOrEmpty(ici.FilenameExtension))
                {
                    continue;
                }
                else
                {
                    string[] exts = ici.FilenameExtension.Split(';');
                    foreach(string ext in exts)
                    {
                        string fileName = string.Format(ICON_GIF_FILENAME, ext.Replace("*",""));
                        if (File.Exists(fileName))
                        {
                            imageFiles.Add(fileName);
                        }
                    }
                }
            }

            // gif読込
            if (imageFiles.Count == 0)
            {
                return;
            }
            string iconFile = imageFiles[0];
            // icon サイズ
            Size size = new Size(ICON_W, ICON_H);
            Image imgGif = Image.FromFile(iconFile, true);


            //Create a new FrameDimension object from this image
            var ImgFrmDim = new FrameDimension(imgGif.FrameDimensionsList[0]);

            //Determine the number of frames in the image
            //Note that all images contain at least 1 frame,
            //but an animated GIF will contain more than 1 frame.
            int n = imgGif.GetFrameCount(ImgFrmDim);

            // Save every frame into icon format
            List<Icon> icons = new List<Icon>();
            for (int i = 0; i < n; i++)
            {
                imgGif.SelectActiveFrame(ImgFrmDim, i);
                //imgGif.Save(string.Format(@"Frame{0}.png", i), ImageFormat.Icon);
                Bitmap iconBm = new Bitmap(imgGif, size);
                Icon icon = Icon.FromHandle(iconBm.GetHicon());
                icons.Add(icon);
            }
            this.tasktrayIcons = icons.ToArray();

        }

        //アニメ表示時にタスクトレイアイコンを変更する
        private void ChangeAnimatedTasktrayIcon()
        {
            if (this.tasktrayIcons.Length == 0)
            {
                return;
            }
            //タスクトレイアイコンを変更する
            this.notifyIcon1.Icon = this.tasktrayIcons[this.currentTasktrayIconIndex];

            //次に表示するアイコンを決める
            this.currentTasktrayIconIndex++;
            if (this.currentTasktrayIconIndex >= this.tasktrayIcons.Length)
                this.currentTasktrayIconIndex = 0;
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            this.ChangeAnimatedTasktrayIcon();
        }
        private void CloseForm(object sender, CancelEventArgs e)
        {
            this.ShowInTaskbar = false;
            this.Hide();
            e.Cancel = true;
        }
        private void MinimumForm(object sender, EventArgs e)
        {
            this.ShowInTaskbar = false;
            this.Hide();
        }


        private void ShownStatus(object sender, EventArgs e)
        {
            if (this.Visible)
            {
                this.ShowInTaskbar = true;
                this.rtStatuses.Clear();
                this.rtStatuses.Text = NetworkFunctions.GetNetworkStatus();
            }
            else
            {
                this.rtStatuses.Clear();
                this.ShowInTaskbar = false;
            }
        }

    }
}
