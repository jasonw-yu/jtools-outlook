using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace jtools_outlook
{
    public partial class ThisAddIn
    {
        private const string AppVersion = "v1.1.1";

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // 备注: Outlook不会再触发这个事件。如果具有
            //    在 Outlook 关闭时必须运行，详请参阅 https://go.microsoft.com/fwlink/?LinkId=506785
        }

        #region VSTO 生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }

    public class ProgressForm : Form
    {
        private ProgressBar progressBar;
        private Label lblStatus;
        private Button btnCancel;

        public bool IsCancelled { get; private set; }

        public ProgressForm()
        {
            IsCancelled = false;
            this.Text = "保存进度";
            this.Width = 550;
            this.Height = 180;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            var tableLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 3,
                Padding = new Padding(15)
            };

            // 状态标签
            lblStatus = new Label
            {
                Text = "准备保存...",
                Dock = DockStyle.Fill,
                TextAlign = System.Drawing.ContentAlignment.MiddleLeft,
                Font = new System.Drawing.Font("Microsoft Sans Serif", 10),
                Height = 25
            };

            // 进度条
            progressBar = new ProgressBar
            {
                Dock = DockStyle.Fill,
                Minimum = 0,
                Maximum = 100,
                Height = 25
            };

            // 取消按钮
            btnCancel = new Button
            {
                Text = "停止保存",
                Width = 100,
                Height = 30,
                Anchor = AnchorStyles.None
            };
            btnCancel.Click += (s, e) =>
            {
                IsCancelled = true;
                lblStatus.Text = "正在停止，请稍候...";
                btnCancel.Enabled = false;
            };

            var buttonPanel = new Panel { Height = 50 };
            buttonPanel.Controls.Add(btnCancel);
            btnCancel.Left = (buttonPanel.Width - btnCancel.Width) / 2;
            btnCancel.Top = 10;

            tableLayout.Controls.Add(lblStatus, 0, 0);
            tableLayout.Controls.Add(progressBar, 0, 1);
            tableLayout.Controls.Add(buttonPanel, 0, 2);

            tableLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));
            tableLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 35));
            tableLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 50));

            this.Controls.Add(tableLayout);
        }

        public void SetProgress(int current, int total)
        {
            if (total > 0)
            {
                progressBar.Maximum = total;
                progressBar.Value = Math.Min(current, total);
            }
            else
            {
                progressBar.Maximum = 1;
                progressBar.Value = 0;
            }
        }

        public void IncrementProgress()
        {
            if (progressBar.Value < progressBar.Maximum)
            {
                progressBar.Value++;
            }
        }

        public void UpdateStatus(string message)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new System.Action<string>(UpdateStatus), message);
                return;
            }
            lblStatus.Text = message;
        }
    }

    public class DateRangePickerForm : Form
    {
        private DateTimePicker startDatePicker;
        private DateTimePicker endDatePicker;
        private TextBox pathTextBox;
        private Button browseButton;
        private Button okButton;
        private Button cancelButton;
        private CheckBox chkInbox;
        private CheckBox chkSentItems;

        public DateTime StartDate { get; private set; }
        public DateTime EndDate { get; private set; }
        public string SavePath { get; private set; }
        public bool SaveInbox { get; private set; }
        public bool SaveSentItems { get; private set; }

        public DateRangePickerForm()
        {
            this.Text = "保存邮件附件";
            this.Width = 480;
            this.Height = 480;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.StartPosition = FormStartPosition.CenterScreen;

            // 主容器
            var tableLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 11,
                Padding = new Padding(20)
            };



            // 保存路径标签
            var pathLabel = new Label
            {
                Text = "保存路径：",
                Dock = DockStyle.Fill,
                TextAlign = System.Drawing.ContentAlignment.BottomLeft,
                Font = new System.Drawing.Font("Microsoft Sans Serif", 10, System.Drawing.FontStyle.Bold),
                Height = 25
            };

            // 路径选择面板
            var pathPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Height = 30
            };

            pathTextBox = new TextBox
            {
                Left = 0,
                Top = 2,
                Width = 330,
                Height = 25,
                ReadOnly = true
            };

            browseButton = new Button
            {
                Text = "浏览...",
                Left = 340,
                Top = 0,
                Width = 70,
                Height = 28
            };
            browseButton.Click += BrowseButton_Click;

            pathPanel.Controls.Add(pathTextBox);
            pathPanel.Controls.Add(browseButton);

            // 起始日期标签
            var startLabel = new Label
            {
                Text = "起始日期：",
                Dock = DockStyle.Fill,
                TextAlign = System.Drawing.ContentAlignment.BottomLeft,
                Font = new System.Drawing.Font("Microsoft Sans Serif", 10, System.Drawing.FontStyle.Bold),
                Height = 25
            };

            // 起始日期选择器
            startDatePicker = new DateTimePicker
            {
                Dock = DockStyle.Fill,
                Format = DateTimePickerFormat.Short,
                Height = 25
            };

            // 分隔
            var spacer = new Label
            {
                Text = "",
                Dock = DockStyle.Fill,
                Height = 5
            };

            // 结束日期标签
            var endLabel = new Label
            {
                Text = "结束日期：",
                Dock = DockStyle.Fill,
                TextAlign = System.Drawing.ContentAlignment.BottomLeft,
                Font = new System.Drawing.Font("Microsoft Sans Serif", 10, System.Drawing.FontStyle.Bold),
                Height = 25
            };

            // 结束日期选择器
            endDatePicker = new DateTimePicker
            {
                Dock = DockStyle.Fill,
                Format = DateTimePickerFormat.Short,
                Height = 25
            };

            // 分隔
            var spacer2 = new Label
            {
                Text = "",
                Dock = DockStyle.Fill,
                Height = 5
            };

            // 文件夹选择标签
            var folderLabel = new Label
            {
                Text = "选择要保存附件的文件夹：",
                Dock = DockStyle.Fill,
                TextAlign = System.Drawing.ContentAlignment.BottomLeft,
                Font = new System.Drawing.Font("Microsoft Sans Serif", 10, System.Drawing.FontStyle.Bold),
                Height = 25
            };

            // 收件箱复选框
            chkInbox = new CheckBox
            {
                Text = "收件箱",
                Dock = DockStyle.Fill,
                Height = 25,
                Checked = true
            };

            // 已发送邮件复选框
            chkSentItems = new CheckBox
            {
                Text = "已发送邮件",
                Dock = DockStyle.Fill,
                Height = 25,
                Checked = false
            };

            // 按钮面板
            var buttonPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Height = 40
            };

            okButton = new Button
            {
                Text = "确定",
                Width = 80,
                Height = 30,
                Left = 120,
                Top = 5
            };

            cancelButton = new Button
            {
                Text = "取消",
                Width = 80,
                Height = 30,
                Left = 220,
                Top = 5
            };

            okButton.Click += (sender, e) =>
            {
                if (string.IsNullOrWhiteSpace(pathTextBox.Text))
                {
                    MessageBox.Show("请先选择保存路径！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (!chkInbox.Checked && !chkSentItems.Checked)
                {
                    MessageBox.Show("请至少选择一个文件夹！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                SavePath = pathTextBox.Text;
                StartDate = startDatePicker.Value;
                EndDate = endDatePicker.Value;
                SaveInbox = chkInbox.Checked;
                SaveSentItems = chkSentItems.Checked;
                this.DialogResult = DialogResult.OK;
                this.Close();
            };

            cancelButton.Click += (sender, e) =>
            {
                this.DialogResult = DialogResult.Cancel;
                this.Close();
            };

            buttonPanel.Controls.Add(okButton);
            buttonPanel.Controls.Add(cancelButton);

            // 添加到布局
            tableLayout.Controls.Add(pathLabel, 0, 0);
            tableLayout.Controls.Add(pathPanel, 0, 1);
            tableLayout.Controls.Add(startLabel, 0, 2);
            tableLayout.Controls.Add(startDatePicker, 0, 3);
            tableLayout.Controls.Add(spacer, 0, 4);
            tableLayout.Controls.Add(endLabel, 0, 5);
            tableLayout.Controls.Add(endDatePicker, 0, 6);
            tableLayout.Controls.Add(spacer2, 0, 7);
            tableLayout.Controls.Add(folderLabel, 0, 8);
            tableLayout.Controls.Add(chkInbox, 0, 9);
            tableLayout.Controls.Add(chkSentItems, 0, 10);
            tableLayout.Controls.Add(buttonPanel, 0, 11);

            // 设置行高 - 增加间距避免重叠
            tableLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 35));  // 保存路径标签
            tableLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 40));  // 路径选择面板
            tableLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 35));  // 起始日期标签
            tableLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 40));  // 起始日期选择器
            tableLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 15));  // 分隔
            tableLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 35));  // 结束日期标签
            tableLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 40));  // 结束日期选择器
            tableLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20));  // 分隔
            tableLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 35));  // 文件夹选择标签
            tableLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30));  // 收件箱复选框
            tableLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30));  // 已发送邮件复选框
            tableLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 55));  // 按钮面板

            this.Controls.Add(tableLayout);
        }

        private void BrowseButton_Click(object sender, EventArgs e)
        {
            using (var folderDialog = new FolderBrowserDialog())
            {
                folderDialog.Description = "请选择附件保存的根文件夹";
                if (folderDialog.ShowDialog() == DialogResult.OK)
                {
                    pathTextBox.Text = folderDialog.SelectedPath;
                }
            }
        }
    }

    public class SaveResultForm : Form
    {
        public SaveResultForm(string resultText)
        {
            this.Text = "保存结果详情";
            this.Width = 700;
            this.Height = 500;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.Sizable;
            this.MinimizeBox = false;

            var tableLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 2,
                Padding = new Padding(10)
            };

            // 标题
            var titleLabel = new Label
            {
                Text = "保存结果详情",
                Dock = DockStyle.Fill,
                TextAlign = System.Drawing.ContentAlignment.MiddleCenter,
                Font = new System.Drawing.Font("Microsoft Sans Serif", 12, System.Drawing.FontStyle.Bold),
                ForeColor = System.Drawing.Color.SteelBlue,
                Height = 30
            };

            // 文本框显示结果
            var textBox = new TextBox
            {
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Both,
                Dock = DockStyle.Fill,
                Text = resultText,
                Font = new System.Drawing.Font("Consolas", 9),
                BackColor = System.Drawing.Color.White
            };

            // 按钮面板
            var buttonPanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 40
            };

            var okButton = new Button
            {
                Text = "确定",
                Width = 80,
                Height = 30,
                Left = 300,
                Top = 5,
                DialogResult = DialogResult.OK
            };
            buttonPanel.Controls.Add(okButton);

            // 复制按钮
            var copyButton = new Button
            {
                Text = "复制到剪贴板",
                Width = 100,
                Height = 30,
                Left = 180,
                Top = 5
            };
            copyButton.Click += (s, e) =>
            {
                Clipboard.SetText(resultText);
                MessageBox.Show("已复制到剪贴板", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            };
            buttonPanel.Controls.Add(copyButton);

            tableLayout.Controls.Add(titleLabel, 0, 0);
            tableLayout.Controls.Add(textBox, 0, 1);
            tableLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 35));
            tableLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

            this.Controls.Add(tableLayout);
            this.Controls.Add(buttonPanel);

            this.AcceptButton = okButton;
        }
    }

    /// <summary>
    /// 数据文件信息
    /// </summary>
    public class StoreInfo
    {
        public Outlook.Store Store { get; set; }
        public string DisplayName { get; set; }
        public bool IsArchive { get; set; }

        public override string ToString()
        {
            return DisplayName ?? "未知数据文件";
        }
    }

    /// <summary>
    /// 关于对话框
    /// </summary>
    public class AboutForm : Form
    {
        public AboutForm()
        {
            this.Text = "关于 JTools-outlook";
            this.Width = 450;
            this.Height = 380;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            var tableLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 6,
                Padding = new Padding(25)
            };

            // 应用名称
            var lblTitle = new Label
            {
                Text = "JTools-outlook",
                Dock = DockStyle.Fill,
                TextAlign = System.Drawing.ContentAlignment.MiddleCenter,
                Font = new System.Drawing.Font("Microsoft Sans Serif", 20, System.Drawing.FontStyle.Bold),
                ForeColor = System.Drawing.Color.SteelBlue,
                Height = 45
            };

            // 版本号
            var lblVersion = new Label
            {
                Text = "版本 v1.1.1",
                Dock = DockStyle.Fill,
                TextAlign = System.Drawing.ContentAlignment.MiddleCenter,
                Font = new System.Drawing.Font("Microsoft Sans Serif", 11),
                Height = 25
            };

            // 分隔线
            var separator = new Label
            {
                Text = "",
                Dock = DockStyle.Fill,
                Height = 2,
                BorderStyle = BorderStyle.Fixed3D
            };

            // 描述
            var lblDescription = new Label
            {
                Text = "Outlook功能增强工具",
                Dock = DockStyle.Fill,
                TextAlign = System.Drawing.ContentAlignment.MiddleCenter,
                Font = new System.Drawing.Font("Microsoft Sans Serif", 9),
                Height = 50
            };

            // 版权信息
            var lblCopyright = new Label
            {
                Text = "Copyright © 2025 Jason\n基于 MIT 协议开源",
                Dock = DockStyle.Fill,
                TextAlign = System.Drawing.ContentAlignment.MiddleCenter,
                Font = new System.Drawing.Font("Microsoft Sans Serif", 8),
                ForeColor = System.Drawing.Color.Gray,
                Height = 40
            };

            // 按钮面板
            var buttonPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Height = 45
            };

            var btnOK = new Button
            {
                Text = "确定",
                Width = 80,
                Height = 30,
                Left = 160,
                Top = 8,
                DialogResult = DialogResult.OK
            };
            buttonPanel.Controls.Add(btnOK);

            tableLayout.Controls.Add(lblTitle, 0, 0);
            tableLayout.Controls.Add(lblVersion, 0, 1);
            tableLayout.Controls.Add(separator, 0, 2);
            tableLayout.Controls.Add(lblDescription, 0, 3);
            tableLayout.Controls.Add(lblCopyright, 0, 4);
            tableLayout.Controls.Add(buttonPanel, 0, 5);

            tableLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 50));
            tableLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));
            tableLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 15));
            tableLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 55));
            tableLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 45));
            tableLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

            this.Controls.Add(tableLayout);
            this.AcceptButton = btnOK;
        }
    }

    #region 下载联机功能

    /// <summary>
    /// 年份统计信息
    /// </summary>
    public class YearStats
    {
        public int Year { get; set; }
        public int InboxCount { get; set; }
        public int SentCount { get; set; }
        public int TotalCount { get { return InboxCount + SentCount; } }

        public override string ToString()
        {
            return $"{Year} 年 (收件箱: {InboxCount}, 已发送: {SentCount}, 共: {TotalCount})";
        }
    }

    /// <summary>
    /// 下载联机窗体
    /// </summary>
    public class DownloadOnlineForm : Form
    {
        private Outlook.Application _application;
        private CancellationTokenSource _cancellationTokenSource;
        private bool _isRunning = false;
        private Dictionary<int, YearStats> _yearStats = new Dictionary<int, YearStats>();
        private HashSet<string> _downloadedEntryIds = new HashSet<string>();

        // UI 控件
        private ComboBox cmbSourceStore;
        private Button btnAnalyze;
        private CheckedListBox chkYears;
        private TextBox txtTargetFolder;
        private Button btnBrowseFolder;
        private ProgressBar progressBar;
        private Label lblProgress;
        private TextBox txtLog;
        private Button btnStart;
        private Button btnCancel;
        private Button btnClose;
        private Panel selectPanel;
        private Panel progressPanel;
        private GroupBox grpYears;

        public DownloadOnlineForm(Outlook.Application application)
        {
            _application = application;
            InitializeComponent();
            LoadStores();
        }

        private void InitializeComponent()
        {
            this.Text = "下载联机存档";
            this.Width = 750;
            this.Height = 650;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.Sizable;
            this.MaximizeBox = false;
            this.MinimumSize = new Size(650, 550);

            // 主面板
            var mainPanel = new Panel { Dock = DockStyle.Fill, Padding = new Padding(15) };

            // 选择面板
            selectPanel = new Panel { Dock = DockStyle.Top, Height = 320 };
            var lblTitle = new Label
            {
                Text = "下载联机存档到本地文件夹",
                Dock = DockStyle.Top,
                Height = 30,
                Font = new Font("Microsoft YaHei", 12, FontStyle.Bold),
                ForeColor = Color.SteelBlue
            };

            // 源数据文件选择
            var lblSource = new Label
            {
                Text = "源数据文件（联机存档）:",
                Dock = DockStyle.Top,
                Height = 25,
                Margin = new Padding(0, 10, 0, 0)
            };

            var sourcePanel = new Panel { Dock = DockStyle.Top, Height = 32 };
            cmbSourceStore = new ComboBox
            {
                Dock = DockStyle.Fill,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Height = 28
            };
            btnAnalyze = new Button
            {
                Text = "分析",
                Width = 80,
                Height = 28,
                Dock = DockStyle.Right
            };
            btnAnalyze.Click += BtnAnalyze_Click;
            sourcePanel.Controls.Add(cmbSourceStore);
            sourcePanel.Controls.Add(btnAnalyze);

            // 年份选择区域
            grpYears = new GroupBox
            {
                Text = "选择要下载的年份（分析后显示）",
                Dock = DockStyle.Top,
                Height = 150,
                Margin = new Padding(0, 10, 0, 0)
            };

            chkYears = new CheckedListBox
            {
                Dock = DockStyle.Fill,
                CheckOnClick = true,
                Font = new Font("Microsoft YaHei", 9),
                Margin = new Padding(5)
            };
            grpYears.Controls.Add(chkYears);

            // 目标文件夹选择
            var lblTarget = new Label
            {
                Text = "目标文件夹:",
                Dock = DockStyle.Top,
                Height = 25,
                Margin = new Padding(0, 10, 0, 0)
            };

            var folderPanel = new Panel { Dock = DockStyle.Top, Height = 32 };
            txtTargetFolder = new TextBox
            {
                Dock = DockStyle.Fill,
                Height = 28
            };
            btnBrowseFolder = new Button
            {
                Text = "浏览...",
                Width = 80,
                Height = 28,
                Dock = DockStyle.Right
            };
            btnBrowseFolder.Click += BtnBrowseFolder_Click;
            folderPanel.Controls.Add(txtTargetFolder);
            folderPanel.Controls.Add(btnBrowseFolder);

            selectPanel.Controls.AddRange(new Control[] {
                folderPanel, lblTarget,
                grpYears,
                sourcePanel, lblSource,
                lblTitle
            });

            // 进度面板
            progressPanel = new Panel { Dock = DockStyle.Top, Height = 55, Margin = new Padding(0, 10, 0, 0) };

            progressBar = new ProgressBar
            {
                Dock = DockStyle.Top,
                Height = 25,
                Minimum = 0,
                Maximum = 100
            };

            lblProgress = new Label
            {
                Text = "0 / 0 (0%)",
                Dock = DockStyle.Top,
                Height = 25,
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("Microsoft YaHei", 9, FontStyle.Bold)
            };

            progressPanel.Controls.AddRange(new Control[] { lblProgress, progressBar });

            // 日志面板
            var logPanel = new Panel { Dock = DockStyle.Fill, Margin = new Padding(0, 10, 0, 0) };

            var lblLog = new Label
            {
                Text = "操作日志",
                Dock = DockStyle.Top,
                Height = 25,
                Font = new Font("Microsoft YaHei", 9, FontStyle.Bold)
            };

            txtLog = new TextBox
            {
                Dock = DockStyle.Fill,
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                ReadOnly = true,
                Font = new Font("Consolas", 9),
                BackColor = Color.WhiteSmoke
            };

            logPanel.Controls.AddRange(new Control[] { txtLog, lblLog });

            // 按钮面板
            var buttonPanel = new Panel { Dock = DockStyle.Bottom, Height = 50 };

            btnStart = new Button
            {
                Text = "开始下载",
                Width = 100,
                Height = 32,
                Left = 15,
                Top = 9,
                Enabled = false  // 初始禁用，分析完成后启用
            };
            btnStart.Click += BtnStart_Click;

            btnCancel = new Button
            {
                Text = "取消",
                Width = 80,
                Height = 32,
                Left = 125,
                Top = 9,
                Enabled = false
            };
            btnCancel.Click += BtnCancel_Click;

            btnClose = new Button
            {
                Text = "关闭",
                Width = 80,
                Height = 32,
                Left = 630,
                Top = 9,
                Anchor = AnchorStyles.Top | AnchorStyles.Right
            };
            btnClose.Click += (s, e) => this.Close();

            buttonPanel.Controls.AddRange(new Control[] { btnStart, btnCancel, btnClose });

            mainPanel.Controls.AddRange(new Control[] { logPanel, progressPanel, selectPanel });
            this.Controls.AddRange(new Control[] { mainPanel, buttonPanel });
        }

        private void BtnBrowseFolder_Click(object sender, EventArgs e)
        {
            using (var dialog = new FolderBrowserDialog())
            {
                dialog.Description = "选择保存邮件的目标文件夹";
                dialog.ShowNewFolderButton = true;
                if (!string.IsNullOrEmpty(txtTargetFolder.Text) && Directory.Exists(txtTargetFolder.Text))
                {
                    dialog.SelectedPath = txtTargetFolder.Text;
                }
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    txtTargetFolder.Text = dialog.SelectedPath;
                }
            }
        }

        private void LoadStores()
        {
            try
            {
                cmbSourceStore.Items.Clear();
                int archiveCount = 0;

                foreach (Outlook.Store store in _application.Session.Stores)
                {
                    try
                    {
                        bool isArchive = store.ExchangeStoreType != Outlook.OlExchangeStoreType.olNotExchange &&
                                        store.ExchangeStoreType != Outlook.OlExchangeStoreType.olPrimaryExchangeMailbox;

                        if (isArchive)
                        {
                            var info = new StoreInfo
                            {
                                Store = store,
                                DisplayName = store.DisplayName,
                                IsArchive = true
                            };
                            cmbSourceStore.Items.Add(info);
                            archiveCount++;
                        }
                    }
                    catch { }
                }

                if (cmbSourceStore.Items.Count > 0)
                {
                    cmbSourceStore.SelectedIndex = 0;
                }

                AddLog($"已加载 {archiveCount} 个联机存档");
                AddLog("请选择源数据文件后点击\"分析\"按钮");
            }
            catch (System.Exception ex)
            {
                AddLog($"加载数据文件失败: {ex.Message}");
            }
        }

        private void AddLog(string message)
        {
            if (this.InvokeRequired)
            {
                this.BeginInvoke(new System.Action(() => AddLog(message)));
                return;
            }

            string timestamp = DateTime.Now.ToString("HH:mm:ss");
            txtLog.AppendText($"[{timestamp}] {message}\r\n");
            txtLog.SelectionStart = txtLog.Text.Length;
            txtLog.ScrollToCaret();
        }

        private async void BtnAnalyze_Click(object sender, EventArgs e)
        {
            var sourceInfo = cmbSourceStore.SelectedItem as StoreInfo;
            if (sourceInfo?.Store == null)
            {
                MessageBox.Show("请选择源数据文件", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            btnAnalyze.Enabled = false;
            chkYears.Items.Clear();
            _yearStats.Clear();

            // 先更新 UI 显示状态
            grpYears.Text = "选择要下载的年份（正在分析...）";
            AddLog($"正在分析 {sourceInfo.DisplayName}...");

            // 强制刷新 UI，确保用户看到状态变化
            this.Refresh();
            System.Windows.Forms.Application.DoEvents();

            try
            {
                // 获取 StoreId（在主线程获取，避免跨线程 COM 问题）
                string storeId = sourceInfo.Store.StoreID;
                string storeName = sourceInfo.DisplayName;

                // 在后台线程执行分析
                await Task.Run(() => AnalyzeStoreInBackground(storeId, storeName));

                // 显示年份统计
                grpYears.Text = $"选择要下载的年份（共 {_yearStats.Count} 个年份）";
                foreach (var stats in _yearStats.Values.OrderByDescending(y => y.Year))
                {
                    chkYears.Items.Add(stats, true);
                }

                AddLog($"分析完成，共发现 {_yearStats.Count} 个年份");
                
                // 分析完成后启用开始下载按钮
                if (_yearStats.Count > 0)
                {
                    btnStart.Enabled = true;
                }
            }
            catch (System.Exception ex)
            {
                AddLog($"分析失败: {ex.Message}");
                MessageBox.Show($"分析失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                btnAnalyze.Enabled = true;
            }
        }

        /// <summary>
        /// 在后台线程分析邮件（使用 STAThread 处理 COM）
        /// </summary>
        private void AnalyzeStoreInBackground(string storeId, string storeName)
        {
            System.Exception bgException = null;

            // 在 STAThread 后台线程中执行分析
            var thread = new Thread(() =>
            {
                try
                {
                    LogToUi("[后台] 开始创建 Outlook 实例...");

                    // 创建新的 Outlook Application 实例
                    var app = new Outlook.Application();
                    var ns = app.GetNamespace("MAPI");

                    LogToUi("[后台] Outlook 实例创建成功");

                    try
                    {
                        // 重新获取 Store
                        LogToUi("[后台] 正在查找数据文件...");
                        Outlook.Store bgStore = null;
                        int storeCount = 0;
                        foreach (Outlook.Store s in ns.Stores)
                        {
                            storeCount++;
                            try
                            {
                                if (s.StoreID == storeId)
                                {
                                    bgStore = s;
                                    LogToUi($"[后台] 找到目标数据文件 (共扫描 {storeCount} 个)");
                                    break;
                                }
                            }
                            catch { }
                        }

                        if (bgStore == null)
                        {
                            LogToUi("[后台] 未找到目标数据文件！");
                        }

                        if (bgStore != null)
                        {
                            // 获取文件夹 EntryID
                            LogToUi("[后台] 正在查找收件箱...");
                            string inboxEntryId = null;
                            string sentEntryId = null;

                            try
                            {
                                var inbox = FindFolder(bgStore, "收件箱", "Inbox");
                                if (inbox != null)
                                {
                                    inboxEntryId = inbox.EntryID;
                                    LogToUi($"[后台] 找到收件箱: {inbox.Name}");
                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(inbox);
                                }
                                else
                                {
                                    LogToUi("[后台] 未找到收件箱");
                                }
                            }
                            catch (System.Exception ex)
                            {
                                LogToUi($"[后台] 查找收件箱失败: {ex.Message}");
                            }

                            LogToUi("[后台] 正在查找已发送邮件...");
                            try
                            {
                                var sent = FindFolder(bgStore, "已发送邮件", "Sent Items", "已发送");
                                if (sent != null)
                                {
                                    sentEntryId = sent.EntryID;
                                    LogToUi($"[后台] 找到已发送: {sent.Name}");
                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(sent);
                                }
                                else
                                {
                                    LogToUi("[后台] 未找到已发送邮件文件夹");
                                }
                            }
                            catch (System.Exception ex)
                            {
                                LogToUi($"[后台] 查找已发送失败: {ex.Message}");
                            }

                            // 分析收件箱
                            if (!string.IsNullOrEmpty(inboxEntryId))
                            {
                                LogToUi("[后台] 开始分析收件箱...");
                                try
                                {
                                    var inbox = ns.GetFolderFromID(inboxEntryId);
                                    if (inbox != null)
                                    {
                                        AnalyzeFolderByYearUsingTable(inbox, "Inbox");
                                        System.Runtime.InteropServices.Marshal.ReleaseComObject(inbox);
                                        LogToUi("[后台] 收件箱分析完成");
                                    }
                                }
                                catch (System.Exception ex)
                                {
                                    LogToUi($"[后台] 分析收件箱失败: {ex.Message}");
                                }
                            }

                            // 分析已发送
                            if (!string.IsNullOrEmpty(sentEntryId))
                            {
                                LogToUi("[后台] 开始分析已发送邮件...");
                                try
                                {
                                    var sent = ns.GetFolderFromID(sentEntryId);
                                    if (sent != null)
                                    {
                                        AnalyzeFolderByYearUsingTable(sent, "Sent");
                                        System.Runtime.InteropServices.Marshal.ReleaseComObject(sent);
                                        LogToUi("[后台] 已发送邮件分析完成");
                                    }
                                }
                                catch (System.Exception ex)
                                {
                                    LogToUi($"[后台] 分析已发送失败: {ex.Message}");
                                }
                            }

                            System.Runtime.InteropServices.Marshal.ReleaseComObject(bgStore);
                        }
                    }
                    finally
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(ns);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
                        LogToUi("[后台] Outlook 资源已释放");
                    }
                }
                catch (System.Exception ex)
                {
                    bgException = ex;
                    LogToUi($"[后台] 分析失败: {ex.Message}");
                    System.Diagnostics.Debug.WriteLine($"后台分析失败: {ex.Message}");
                }
            });

            thread.SetApartmentState(ApartmentState.STA);
            LogToUi("[主线程] 启动后台分析线程...");
            thread.Start();
            thread.Join();  // 等待线程完成
            LogToUi("[主线程] 后台分析线程已完成");

            if (bgException != null)
            {
                throw bgException;
            }
        }

        /// <summary>
        /// 线程安全的日志输出
        /// </summary>
        private void LogToUi(string message)
        {
            if (this.InvokeRequired)
            {
                this.BeginInvoke(new System.Action(() => LogToUi(message)));
                return;
            }
            AddLog(message);
        }

        /// <summary>
        /// 使用 GetTable() 高效分析文件夹（比遍历 Items 快得多）
        /// </summary>
        private void AnalyzeFolderByYearUsingTable(Outlook.MAPIFolder folder, string folderType)
        {
            try
            {
                LogToUi($"[分析] 获取 {folderType} 邮件总数...");

                // 先获取总数量
                var items = folder.Items;
                int totalCount = items.Count;
                LogToUi($"[分析] {folderType} 共有 {totalCount} 封邮件");

                if (totalCount == 0)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(items);
                    return;
                }

                // 获取年份范围
                LogToUi($"[分析] 获取 {folderType} 年份范围...");
                items.Sort("[ReceivedTime]", true);

                DateTime? minDate = null;
                DateTime? maxDate = null;

                // 获取最早和最晚的邮件日期
                try
                {
                    var firstItem = items[1];
                    if (firstItem is Outlook.MailItem firstMail)
                    {
                        minDate = firstMail.ReceivedTime;
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(firstMail);
                    }
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(firstItem);
                }
                catch { }

                try
                {
                    var lastItem = items[totalCount];
                    if (lastItem is Outlook.MailItem lastMail)
                    {
                        maxDate = lastMail.ReceivedTime;
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(lastMail);
                    }
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(lastItem);
                }
                catch { }

                System.Runtime.InteropServices.Marshal.ReleaseComObject(items);

                if (minDate.HasValue && maxDate.HasValue)
                {
                    int minYear = minDate.Value.Year;
                    int maxYear = maxDate.Value.Year;

                    // 确保年份范围正确（minYear 可能大于 maxYear，因为排序是降序）
                    int startYear = Math.Min(minYear, maxYear);
                    int endYear = Math.Max(minYear, maxYear);

                    LogToUi($"[分析] {folderType} 年份范围: {startYear} - {endYear}");

                    // 按年份统计数量（使用 Restrict 过滤）
                    for (int year = startYear; year <= endYear; year++)
                    {
                        try
                        {
                            // 使用 Restrict 按年份过滤，然后取 Count
                            string filter = $"[ReceivedTime] >= '{year}/1/1' AND [ReceivedTime] < '{year + 1}/1/1'";
                            var yearItems = folder.Items.Restrict(filter);
                            int yearCount = yearItems.Count;

                            if (yearCount > 0)
                            {
                                lock (_yearStats)
                                {
                                    if (!_yearStats.ContainsKey(year))
                                    {
                                        _yearStats[year] = new YearStats { Year = year };
                                    }
                                    if (folderType == "Inbox")
                                        _yearStats[year].InboxCount = yearCount;
                                    else
                                        _yearStats[year].SentCount = yearCount;
                                }
                                LogToUi($"[分析] {folderType} {year}年: {yearCount} 封");
                            }

                            System.Runtime.InteropServices.Marshal.ReleaseComObject(yearItems);
                        }
                        catch (System.Exception ex)
                        {
                            LogToUi($"[分析] 统计 {year} 年失败: {ex.Message}");
                        }
                    }
                }
                else
                {
                    LogToUi($"[分析] 无法获取 {folderType} 年份范围，使用遍历方式...");
                    AnalyzeFolderByYearFallback(folder, folderType);
                }
            }
            catch (System.Exception ex)
            {
                LogToUi($"[分析] 分析失败: {ex.Message}，尝试回退方法...");
                System.Diagnostics.Debug.WriteLine($"分析失败: {ex.Message}");
                AnalyzeFolderByYearFallback(folder, folderType);
            }
        }

        /// <summary>
        /// 回退方法：当 GetTable 不可用时使用（分批处理避免阻塞）
        /// </summary>
        private void AnalyzeFolderByYearFallback(Outlook.MAPIFolder folder, string folderType)
        {
            try
            {
                var items = folder.Items;
                int count = items.Count;
                int batchSize = 100;  // 每批处理100封

                for (int batch = 0; batch < (count + batchSize - 1) / batchSize; batch++)
                {
                    int start = batch * batchSize + 1;
                    int end = Math.Min(start + batchSize - 1, count);

                    for (int i = start; i <= end; i++)
                    {
                        try
                        {
                            var item = items[i];
                            if (item is Outlook.MailItem mail)
                            {
                                int year = mail.ReceivedTime.Year;

                                lock (_yearStats)
                                {
                                    if (!_yearStats.ContainsKey(year))
                                    {
                                        _yearStats[year] = new YearStats { Year = year };
                                    }
                                    if (folderType == "Inbox")
                                        _yearStats[year].InboxCount++;
                                    else
                                        _yearStats[year].SentCount++;
                                }

                                System.Runtime.InteropServices.Marshal.ReleaseComObject(mail);
                            }
                            if (item != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(item);
                        }
                        catch { }
                    }

                    // 每批后让出线程
                    Thread.Sleep(1);
                }

                System.Runtime.InteropServices.Marshal.ReleaseComObject(items);
            }
            catch { }
        }

        private Outlook.MAPIFolder FindFolder(Outlook.Store store, params string[] possibleNames)
        {
            try
            {
                var root = store.GetRootFolder();
                foreach (Outlook.MAPIFolder folder in root.Folders)
                {
                    try
                    {
                        foreach (var name in possibleNames)
                        {
                            if (folder.Name.Equals(name, StringComparison.OrdinalIgnoreCase))
                            {
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(root);
                                return folder;
                            }
                        }
                    }
                    catch { }
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(folder);
                }
                System.Runtime.InteropServices.Marshal.ReleaseComObject(root);
            }
            catch { }
            return null;
        }

        private async void BtnStart_Click(object sender, EventArgs e)
        {
            if (_isRunning) return;

            var sourceInfo = cmbSourceStore.SelectedItem as StoreInfo;
            if (sourceInfo?.Store == null)
            {
                MessageBox.Show("请选择源数据文件", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (chkYears.CheckedItems.Count == 0)
            {
                MessageBox.Show("请选择要下载的年份", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (string.IsNullOrWhiteSpace(txtTargetFolder.Text))
            {
                MessageBox.Show("请选择目标文件夹", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 确保目标目录存在
            string targetFolder = txtTargetFolder.Text;
            if (!Directory.Exists(targetFolder))
            {
                try
                {
                    Directory.CreateDirectory(targetFolder);
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show($"无法创建目标目录: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }

            // 获取选中的年份
            var selectedYears = chkYears.CheckedItems.Cast<YearStats>().Select(s => s.Year).OrderByDescending(y => y).ToList();

            _isRunning = true;
            btnStart.Enabled = false;
            btnCancel.Enabled = true;
            btnAnalyze.Enabled = false;
            cmbSourceStore.Enabled = false;
            btnBrowseFolder.Enabled = false;

            _cancellationTokenSource = new CancellationTokenSource();

            try
            {
                await DownloadEmailsAsync(sourceInfo.Store, targetFolder, selectedYears, _cancellationTokenSource.Token);
            }
            catch (OperationCanceledException)
            {
                AddLog("下载已取消");
            }
            catch (System.Exception ex)
            {
                AddLog($"下载失败: {ex.Message}");
                MessageBox.Show($"下载失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                _isRunning = false;
                btnStart.Enabled = true;
                btnCancel.Enabled = false;
                btnAnalyze.Enabled = true;
                cmbSourceStore.Enabled = true;
                btnBrowseFolder.Enabled = true;
                _cancellationTokenSource?.Dispose();
                _cancellationTokenSource = null;

                // 显示完成提示 - 不再调用 UpdateProgress，保持最后的状态
                MessageBox.Show("下载完成！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            if (_cancellationTokenSource != null && !_cancellationTokenSource.IsCancellationRequested)
            {
                _cancellationTokenSource.Cancel();
                AddLog("正在取消下载...");
                btnCancel.Enabled = false;
            }
        }

        private async Task DownloadEmailsAsync(Outlook.Store sourceStore, string targetFolder, List<int> selectedYears, CancellationToken cancellationToken)
        {
            AddLog($"开始下载邮件...");
            AddLog($"源: {sourceStore.DisplayName}");
            AddLog($"目标: {targetFolder}");
            AddLog($"选中年份: {string.Join(", ", selectedYears)}");

            int totalDownloaded = 0;
            int totalSkipped = 0;

            // 在主线程获取文件夹 EntryID（不传递 COM 对象）
            var folderEntryIds = new List<(string EntryId, string StoreId, string FolderName)>();
            string storeId = sourceStore.StoreID;

            try
            {
                // 获取源文件夹 EntryID
                var sourceInbox = FindFolder(sourceStore, "收件箱", "Inbox");
                if (sourceInbox != null)
                {
                    folderEntryIds.Add((sourceInbox.EntryID, storeId, "收件箱"));
                    AddLog($"找到收件箱: {sourceInbox.Name}");
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(sourceInbox);
                }

                var sourceSent = FindFolder(sourceStore, "已发送邮件", "Sent Items", "已发送");
                if (sourceSent != null)
                {
                    folderEntryIds.Add((sourceSent.EntryID, storeId, "已发送邮件"));
                    AddLog($"找到已发送: {sourceSent.Name}");
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(sourceSent);
                }

                if (folderEntryIds.Count == 0)
                {
                    AddLog("未找到可下载的文件夹");
                    return;
                }

                // 按年份下载
                foreach (var year in selectedYears)
                {
                    if (cancellationToken.IsCancellationRequested)
                        break;

                    AddLog($"");
                    AddLog($"--- 下载 {year} 年邮件 ---");

                    // 创建年份目录
                    string yearFolder = Path.Combine(targetFolder, year.ToString());
                    if (!Directory.Exists(yearFolder))
                    {
                        Directory.CreateDirectory(yearFolder);
                    }

                    // 处理每个源文件夹
                    foreach (var (folderEntryId, folderStoreId, folderName) in folderEntryIds)
                    {
                        if (cancellationToken.IsCancellationRequested)
                            break;

                        // 创建子文件夹
                        string subFolder = Path.Combine(yearFolder, folderName);
                        if (!Directory.Exists(subFolder))
                        {
                            Directory.CreateDirectory(subFolder);
                        }

                        AddLog($"下载 {folderName} 到 {subFolder}...");

                        int yearDownloaded = 0;
                        int yearSkipped = 0;

                        // 传递 EntryID 到后台线程
                        await Task.Run(() =>
                        {
                            DownloadFolderToFilesByEntryId(folderEntryId, folderStoreId, year, subFolder, cancellationToken,
                                ref yearDownloaded, ref yearSkipped);
                        }, cancellationToken);

                        totalDownloaded += yearDownloaded;
                        totalSkipped += yearSkipped;

                        AddLog($"  {folderName}: 下载 {yearDownloaded}，跳过 {yearSkipped}");

                        if (cancellationToken.IsCancellationRequested)
                            break;
                    }

                    if (cancellationToken.IsCancellationRequested)
                        break;
                }
            }
            finally
            {
                AddLog($"");
                AddLog($"========== 下载完成 ==========");
                AddLog($"总计: 下载 {totalDownloaded} 封，跳过 {totalSkipped} 封");
                
                // 更新状态显示为最终结果
                UpdateProgress(0, 0, totalDownloaded, totalSkipped, 0, "完成");
            }
        }

        /// <summary>
        /// 通过 EntryID 在后台线程获取文件夹并下载邮件
        /// </summary>
        private void DownloadFolderToFilesByEntryId(string folderEntryId, string storeId, int year, string targetPath,
            CancellationToken cancellationToken, ref int downloadedCount, ref int skippedCount)
        {
            Outlook.NameSpace ns = null;
            Outlook.MAPIFolder folder = null;

            try
            {
                AddLog("[后台] 开始获取文件夹...");

                // 使用现有的 Application 实例（通过 GetNamespace 获取）
                ns = Globals.ThisAddIn.Application.GetNamespace("MAPI");

                AddLog("[后台] 正在通过 EntryID 获取文件夹...");

                // 通过 EntryID 获取文件夹
                folder = ns.GetFolderFromID(folderEntryId, storeId);
                if (folder == null)
                {
                    AddLog($"无法获取文件夹: {folderEntryId}");
                    return;
                }

                AddLog($"[后台] 成功获取文件夹: {folder.Name}");

                // 使用 Restrict 过滤该年份的邮件
                string filter = $"[ReceivedTime] >= '{year}/1/1' AND [ReceivedTime] < '{year + 1}/1/1'";
                var filteredItems = folder.Items.Restrict(filter);
                int total = filteredItems.Count;

                AddLog($"[后台] {folder.Name} {year}年共 {total} 封邮件");

                if (total == 0)
                {
                    AddLog($"[后台] 没有需要下载的邮件");
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(filteredItems);
                    return;
                }

                int processed = 0;
                int batchSize = 20;

                // 直接遍历过滤后的邮件
                for (int i = 1; i <= total; i++)
                {
                    if (cancellationToken.IsCancellationRequested)
                        break;

                    object item = null;
                    Outlook.MailItem mailItem = null;

                    try
                    {
                        item = filteredItems[i];
                        processed++;

                        if (item is Outlook.MailItem mail)
                        {
                            mailItem = mail;

                            // 获取邮件信息
                            string subject = mailItem.Subject;
                            DateTime receivedTime = mailItem.ReceivedTime;

                            // 使用精确时间戳生成文件名
                            string safeSubject = GetSafeFileNameWithTimestamp(subject, receivedTime);
                            string filePath = Path.Combine(targetPath, safeSubject + ".msg");

                            if (File.Exists(filePath))
                            {
                                skippedCount++;
                            }
                            else
                            {
                                // 保存邮件
                                mailItem.SaveAs(filePath);
                                downloadedCount++;
                            }
                        }

                        // 更新进度
                        if (processed % 10 == 0)
                        {
                            UpdateProgress(processed, total, downloadedCount, skippedCount, year, "下载中");
                        }

                        // 每处理50封输出一次日志
                        if (processed % 50 == 0)
                        {
                            AddLog($"[下载] 已处理 {processed}/{total}，下载 {downloadedCount}，跳过 {skippedCount}");
                        }

                        // 批次间让出线程
                        if (processed % batchSize == 0)
                        {
                            Thread.Sleep(10);
                        }
                    }
                    catch (System.Exception ex)
                    {
                        AddLog($"处理邮件失败: {ex.Message}");
                    }
                    finally
                    {
                        if (mailItem != null)
                        {
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(mailItem);
                            mailItem = null;
                        }
                        if (item != null && item != mailItem)
                        {
                            try { System.Runtime.InteropServices.Marshal.ReleaseComObject(item); } catch { }
                        }
                    }
                }

                System.Runtime.InteropServices.Marshal.ReleaseComObject(filteredItems);
            }
            catch (System.Exception ex)
            {
                AddLog($"下载文件夹失败: {ex.Message}");
            }
            finally
            {
                // 释放 COM 对象
                if (folder != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(folder);
                    folder = null;
                }
                if (ns != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ns);
                    ns = null;
                }
            }
        }

        private string GetSafeFileNameWithTimestamp(string subject, DateTime receivedTime)
        {
            // 移除非法字符
            string safeName = subject ?? "无主题";
            foreach (char c in Path.GetInvalidFileNameChars())
            {
                safeName = safeName.Replace(c, '_');
            }
            if (safeName.Length > 100)
            {
                safeName = safeName.Substring(0, 100);
            }
            // 使用精确时间戳（精确到秒）作为唯一标识
            string timestamp = receivedTime.ToString("yyyyMMdd_HHmmss");
            return $"{safeName}_{timestamp}";
        }

        private string GetSafeFileName(string subject, string entryId)
        {
            // 移除非法字符
            string safeName = subject ?? "无主题";
            foreach (char c in Path.GetInvalidFileNameChars())
            {
                safeName = safeName.Replace(c, '_');
            }
            // 限制长度
            if (safeName.Length > 100)
            {
                safeName = safeName.Substring(0, 100);
            }
            // EntryID 是 Outlook MAPI 的核心标识符，正常邮件一定有
            // 使用前8位作为后缀确保唯一性
            if (!string.IsNullOrEmpty(entryId) && entryId.Length >= 8)
            {
                return $"{safeName}_{entryId.Substring(0, 8)}";
            }
            else
            {
                // EntryID 获取失败时，使用 GUID 确保唯一性
                return $"{safeName}_{Guid.NewGuid().ToString("N").Substring(0, 8)}";
            }
        }

        private void UpdateProgress(int processed, int total, int downloaded, int skipped, int year, string stage = "")
        {
            if (this.InvokeRequired)
            {
                this.BeginInvoke(new System.Action(() => UpdateProgress(processed, total, downloaded, skipped, year, stage)));
                return;
            }

            int percent = total > 0 ? (int)((double)processed / total * 100) : 0;
            progressBar.Value = Math.Min(percent, 100);
            lblProgress.Text = $"{processed} / {total} ({percent}%)";
        }
    }

    /// <summary>
    /// 阻止域对话框
    /// </summary>
    public class BlockDomainDialog : Form
    {
        private string _domain;
        private Label lblMessage;
        private Button btnOK;
        private Button btnCancel;

        public BlockDomainDialog(string domain)
        {
            _domain = domain;
            this.Text = "阻止域";
            this.Width = 400;
            this.Height = 180;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            var tableLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 3,
                Padding = new Padding(20)
            };

            lblMessage = new Label
            {
                Text = $"确定要阻止来自 @{_domain} 的所有邮件吗？",
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("Microsoft YaHei", 10),
                Height = 50
            };

            var buttonPanel = new Panel { Height = 50 };
            btnOK = new Button
            {
                Text = "确定",
                Width = 80,
                Height = 30,
                Left = 100,
                Top = 10,
                DialogResult = DialogResult.OK
            };
            btnCancel = new Button
            {
                Text = "取消",
                Width = 80,
                Height = 30,
                Left = 200,
                Top = 10,
                DialogResult = DialogResult.Cancel
            };

            buttonPanel.Controls.Add(btnOK);
            buttonPanel.Controls.Add(btnCancel);

            tableLayout.Controls.Add(lblMessage, 0, 0);
            tableLayout.Controls.Add(buttonPanel, 0, 2);

            this.Controls.Add(tableLayout);
            this.AcceptButton = btnOK;
            this.CancelButton = btnCancel;
        }
    }

    #endregion

    #region 导入邮件功能

    /// <summary>
    /// 导入邮件窗体
    /// </summary>
    public class ImportEmailsForm : Form
    {
        private Outlook.Application _application;
        private CancellationTokenSource _cancellationTokenSource;
        private bool _isRunning = false;

        // UI 控件
        private TextBox txtSourceFolder;
        private Button btnBrowseSource;
        private ComboBox cmbTargetStore;
        private ComboBox cmbTargetFolder;
        private ProgressBar progressBar;
        private Label lblStatus;
        private Label lblProgress;
        private TextBox txtLog;
        private Button btnStart;
        private Button btnCancel;
        private Button btnClose;

        public ImportEmailsForm(Outlook.Application application)
        {
            _application = application;
            InitializeComponent();
            LoadStores();
        }

        private void InitializeComponent()
        {
            this.Text = "导入邮件";
            this.Width = 600;
            this.Height = 500;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.Sizable;
            this.MaximizeBox = false;
            this.MinimumSize = new Size(500, 400);

            var mainPanel = new Panel { Dock = DockStyle.Fill, Padding = new Padding(15) };

            // 源文件夹选择
            var lblSource = new Label
            {
                Text = "源文件夹（包含.msg文件）:",
                Dock = DockStyle.Top,
                Height = 25
            };

            var sourcePanel = new Panel { Dock = DockStyle.Top, Height = 32 };
            txtSourceFolder = new TextBox { Dock = DockStyle.Fill, Height = 28 };
            btnBrowseSource = new Button { Text = "浏览...", Width = 80, Height = 28, Dock = DockStyle.Right };
            btnBrowseSource.Click += BtnBrowseSource_Click;
            sourcePanel.Controls.Add(txtSourceFolder);
            sourcePanel.Controls.Add(btnBrowseSource);

            // 目标PST选择
            var lblTargetStore = new Label
            {
                Text = "目标PST文件:",
                Dock = DockStyle.Top,
                Height = 25,
                Margin = new Padding(0, 10, 0, 0)
            };

            cmbTargetStore = new ComboBox
            {
                Dock = DockStyle.Top,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Height = 28
            };
            cmbTargetStore.SelectedIndexChanged += CmbTargetStore_SelectedIndexChanged;

            // 目标文件夹选择
            var lblTargetFolder = new Label
            {
                Text = "目标文件夹:",
                Dock = DockStyle.Top,
                Height = 25,
                Margin = new Padding(0, 10, 0, 0)
            };

            cmbTargetFolder = new ComboBox
            {
                Dock = DockStyle.Top,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Height = 28
            };

            // 进度面板
            var progressPanel = new Panel { Dock = DockStyle.Top, Height = 80, Margin = new Padding(0, 10, 0, 0) };

            lblStatus = new Label { Text = "就绪", Dock = DockStyle.Top, Height = 25 };
            progressBar = new ProgressBar { Dock = DockStyle.Top, Height = 25, Minimum = 0, Maximum = 100 };
            lblProgress = new Label { Text = "0 / 0 (0%)", Dock = DockStyle.Top, Height = 25, TextAlign = ContentAlignment.MiddleCenter };

            progressPanel.Controls.AddRange(new Control[] { lblProgress, progressBar, lblStatus });

            // 日志面板
            var logPanel = new Panel { Dock = DockStyle.Fill, Margin = new Padding(0, 10, 0, 0) };
            var lblLog = new Label { Text = "操作日志", Dock = DockStyle.Top, Height = 25 };
            txtLog = new TextBox
            {
                Dock = DockStyle.Fill,
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                ReadOnly = true,
                BackColor = Color.WhiteSmoke
            };
            logPanel.Controls.AddRange(new Control[] { txtLog, lblLog });

            // 按钮面板
            var buttonPanel = new Panel { Dock = DockStyle.Bottom, Height = 50 };
            btnStart = new Button { Text = "开始导入", Width = 100, Height = 32, Left = 15, Top = 9 };
            btnStart.Click += BtnStart_Click;
            btnCancel = new Button { Text = "取消", Width = 80, Height = 32, Left = 125, Top = 9, Enabled = false };
            btnCancel.Click += BtnCancel_Click;
            btnClose = new Button { Text = "关闭", Width = 80, Height = 32, Left = 480, Top = 9, Anchor = AnchorStyles.Top | AnchorStyles.Right };
            btnClose.Click += (s, e) => this.Close();
            buttonPanel.Controls.AddRange(new Control[] { btnStart, btnCancel, btnClose });

            mainPanel.Controls.AddRange(new Control[] { logPanel, progressPanel });
            mainPanel.Controls.AddRange(new Control[] { cmbTargetFolder, lblTargetFolder, cmbTargetStore, lblTargetStore, sourcePanel, lblSource });
            this.Controls.AddRange(new Control[] { mainPanel, buttonPanel });
        }

        private void BtnBrowseSource_Click(object sender, EventArgs e)
        {
            using (var dialog = new FolderBrowserDialog())
            {
                dialog.Description = "选择包含.msg文件的源文件夹";
                if (!string.IsNullOrEmpty(txtSourceFolder.Text) && Directory.Exists(txtSourceFolder.Text))
                {
                    dialog.SelectedPath = txtSourceFolder.Text;
                }
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    txtSourceFolder.Text = dialog.SelectedPath;
                }
            }
        }

        private void LoadStores()
        {
            try
            {
                cmbTargetStore.Items.Clear();
                foreach (Outlook.Store store in _application.Session.Stores)
                {
                    try
                    {
                        // 只显示本地PST文件
                        if (store.ExchangeStoreType == Outlook.OlExchangeStoreType.olNotExchange ||
                            store.ExchangeStoreType == Outlook.OlExchangeStoreType.olPrimaryExchangeMailbox)
                        {
                            var info = new StoreInfo
                            {
                                Store = store,
                                DisplayName = store.DisplayName,
                                IsArchive = false
                            };
                            cmbTargetStore.Items.Add(info);
                        }
                    }
                    catch { }
                }

                if (cmbTargetStore.Items.Count > 0)
                {
                    cmbTargetStore.SelectedIndex = 0;
                }

                AddLog($"已加载 {cmbTargetStore.Items.Count} 个本地PST文件");
            }
            catch (System.Exception ex)
            {
                AddLog($"加载PST文件失败: {ex.Message}");
            }
        }

        private void CmbTargetStore_SelectedIndexChanged(object sender, EventArgs e)
        {
            LoadFolders();
        }

        private void LoadFolders()
        {
            try
            {
                cmbTargetFolder.Items.Clear();
                var storeInfo = cmbTargetStore.SelectedItem as StoreInfo;
                if (storeInfo?.Store == null) return;

                var rootFolder = storeInfo.Store.GetRootFolder();
                AddFoldersToList(rootFolder, "");
                System.Runtime.InteropServices.Marshal.ReleaseComObject(rootFolder);

                if (cmbTargetFolder.Items.Count > 0)
                {
                    cmbTargetFolder.SelectedIndex = 0;
                }
            }
            catch (System.Exception ex)
            {
                AddLog($"加载文件夹失败: {ex.Message}");
            }
        }

        private void AddFoldersToList(Outlook.MAPIFolder folder, string path)
        {
            try
            {
                var folderInfo = new FolderInfo { Folder = folder, DisplayPath = string.IsNullOrEmpty(path) ? folder.Name : $"{path}\\{folder.Name}" };
                cmbTargetFolder.Items.Add(folderInfo);

                foreach (Outlook.MAPIFolder subFolder in folder.Folders)
                {
                    AddFoldersToList(subFolder, folderInfo.DisplayPath);
                }
            }
            catch { }
        }

        private void AddLog(string message)
        {
            if (this.InvokeRequired)
            {
                this.BeginInvoke(new System.Action(() => AddLog(message)));
                return;
            }

            string timestamp = DateTime.Now.ToString("HH:mm:ss");
            txtLog.AppendText($"[{timestamp}] {message}\r\n");
            txtLog.SelectionStart = txtLog.Text.Length;
            txtLog.ScrollToCaret();
        }

        private async void BtnStart_Click(object sender, EventArgs e)
        {
            if (_isRunning) return;

            if (string.IsNullOrWhiteSpace(txtSourceFolder.Text) || !Directory.Exists(txtSourceFolder.Text))
            {
                MessageBox.Show("请选择有效的源文件夹", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var storeInfo = cmbTargetStore.SelectedItem as StoreInfo;
            if (storeInfo?.Store == null)
            {
                MessageBox.Show("请选择目标PST文件", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var folderInfo = cmbTargetFolder.SelectedItem as FolderInfo;
            if (folderInfo?.Folder == null)
            {
                MessageBox.Show("请选择目标文件夹", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            _isRunning = true;
            btnStart.Enabled = false;
            btnCancel.Enabled = true;
            cmbTargetStore.Enabled = false;
            cmbTargetFolder.Enabled = false;
            btnBrowseSource.Enabled = false;

            _cancellationTokenSource = new CancellationTokenSource();

            try
            {
                string storeId = storeInfo.Store.StoreID;
                string folderId = folderInfo.Folder.EntryID;
                string sourcePath = txtSourceFolder.Text;

                await Task.Run(() => ImportEmailsAsync(storeId, folderId, sourcePath, _cancellationTokenSource.Token));
            }
            catch (OperationCanceledException)
            {
                AddLog("导入已取消");
            }
            catch (System.Exception ex)
            {
                AddLog($"导入失败: {ex.Message}");
            }
            finally
            {
                _isRunning = false;
                btnStart.Enabled = true;
                btnCancel.Enabled = false;
                cmbTargetStore.Enabled = true;
                cmbTargetFolder.Enabled = true;
                btnBrowseSource.Enabled = true;
                _cancellationTokenSource?.Dispose();
                _cancellationTokenSource = null;
            }
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            if (_cancellationTokenSource != null && !_cancellationTokenSource.IsCancellationRequested)
            {
                _cancellationTokenSource.Cancel();
                AddLog("正在取消导入...");
                btnCancel.Enabled = false;
            }
        }

        private void ImportEmailsAsync(string storeId, string folderId, string sourcePath, CancellationToken cancellationToken)
        {
            int imported = 0;
            int skipped = 0;
            int failed = 0;

            try
            {
                var msgFiles = Directory.GetFiles(sourcePath, "*.msg", SearchOption.TopDirectoryOnly);
                int total = msgFiles.Length;

                AddLog($"找到 {total} 个.msg文件");

                if (total == 0)
                {
                    AddLog("没有找到可导入的邮件文件");
                    return;
                }

                Outlook.NameSpace ns = null;
                Outlook.MAPIFolder targetFolder = null;

                try
                {
                    ns = _application.GetNamespace("MAPI");
                    targetFolder = ns.GetFolderFromID(folderId, storeId);

                    if (targetFolder == null)
                    {
                        AddLog("无法获取目标文件夹");
                        return;
                    }

                    // 初始化：扫描目标文件夹，建立已存在邮件的 Message-ID 集合
                    AddLog("正在扫描目标文件夹中的已有邮件...");
                    var existingMessageIds = new HashSet<string>();
                    var items = targetFolder.Items;
                    int scannedCount = 0;
                    
                    foreach (object item in items)
                    {
                        if (item is Outlook.MailItem existingMail)
                        {
                            try
                            {
                                string existingMessageId = GetMessageId(existingMail);
                                if (!string.IsNullOrEmpty(existingMessageId))
                                {
                                    existingMessageIds.Add(existingMessageId);
                                }
                            }
                            catch { }
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(existingMail);
                            scannedCount++;
                        }
                    }
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(items);
                    
                    AddLog($"目标文件夹已有 {existingMessageIds.Count} 封邮件（共扫描 {scannedCount} 封）");

                    // 遍历待导入文件
                    for (int i = 0; i < msgFiles.Length; i++)
                    {
                        if (cancellationToken.IsCancellationRequested)
                            break;

                        string file = msgFiles[i];
                        try
                        {
                            var mailItem = _application.Session.OpenSharedItem(file) as Outlook.MailItem;
                            if (mailItem != null)
                            {
                                // 获取 Message-ID
                                string messageId = GetMessageId(mailItem);

                                bool exists = false;
                                if (!string.IsNullOrEmpty(messageId) && existingMessageIds.Contains(messageId))
                                {
                                    exists = true;
                                }

                                if (exists)
                                {
                                    skipped++;
                                }
                                else
                                {
                                    mailItem.Move(targetFolder);
                                    imported++;
                                    
                                    // 将新导入的 Message-ID 添加到集合中
                                    if (!string.IsNullOrEmpty(messageId))
                                    {
                                        existingMessageIds.Add(messageId);
                                    }
                                }
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(mailItem);
                            }
                        }
                        catch (System.Exception ex)
                        {
                            failed++;
                            if (failed <= 5)
                            {
                                AddLog($"导入失败 ({Path.GetFileName(file)}): {ex.Message}");
                            }
                        }

                        // 更新进度（每处理一封或每10封更新一次）
                        if ((i + 1) % 10 == 0 || i + 1 == total)
                        {
                            UpdateProgress(i + 1, total, imported, skipped, failed);
                        }

                        Thread.Sleep(10);
                    }

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(targetFolder);
                }
                finally
                {
                    if (ns != null)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(ns);
                }

                AddLog($"导入完成: 成功 {imported}，跳过 {skipped}，失败 {failed}");
            }
            catch (System.Exception ex)
            {
                AddLog($"导入过程出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 获取邮件的 Message-ID，并进行标准化处理
        /// </summary>
        private string GetMessageId(Outlook.MailItem mail)
        {
            try
            {
                string messageId = mail.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x1035001F")?.ToString() ?? "";
                
                if (string.IsNullOrEmpty(messageId))
                {
                    // 如果无法获取 Message-ID，使用主题+发件人+时间作为备选
                    messageId = $"{mail.Subject}_{mail.SenderName}_{mail.ReceivedTime:yyyyMMddHHmmss}";
                }
                
                // 标准化处理：去除尖括号和空白
                messageId = messageId.Trim().TrimStart('<').TrimEnd('>');
                
                return messageId;
            }
            catch
            {
                // 回退方案：使用主题+发件人+时间
                try
                {
                    return $"{mail.Subject}_{mail.SenderName}_{mail.ReceivedTime:yyyyMMddHHmmss}".Trim();
                }
                catch
                {
                    return "";
                }
            }
        }

        private void UpdateProgress(int processed, int total, int imported, int skipped, int failed)
        {
            if (this.InvokeRequired)
            {
                this.BeginInvoke(new System.Action(() => UpdateProgress(processed, total, imported, skipped, failed)));
                return;
            }

            int percent = total > 0 ? (int)((double)processed / total * 100) : 0;
            progressBar.Value = Math.Min(percent, 100);
            lblProgress.Text = $"{processed} / {total} ({percent}%)";
            lblStatus.Text = $"导入中 - 成功: {imported}，跳过: {skipped}，失败: {failed}";
        }
    }

    /// <summary>
    /// 文件夹信息
    /// </summary>
    public class FolderInfo
    {
        public Outlook.MAPIFolder Folder { get; set; }
        public string DisplayPath { get; set; }

        public override string ToString()
        {
            return DisplayPath ?? "未知文件夹";
        }
    }

    #endregion
}
