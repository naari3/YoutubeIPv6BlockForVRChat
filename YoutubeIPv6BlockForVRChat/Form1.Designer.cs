namespace YoutubeIPv6BlockForVRChat
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            notifyIcon1 = new NotifyIcon(components);
            contextMenuStrip1 = new ContextMenuStrip(components);
            toolStripMenuItem1 = new ToolStripMenuItem();
            toolStripSeparator1 = new ToolStripSeparator();
            ShowMenu = new ToolStripMenuItem();
            IPv6Block = new ToolStripMenuItem();
            Exit = new ToolStripMenuItem();
            checkBoxAutoStart = new CheckBox();
            buttonBlock = new Button();
            buttonUnblock = new Button();
            groupBox1 = new GroupBox();
            timerCheckVRCRunning = new System.Windows.Forms.Timer(components);
            label1 = new Label();
            timerCheckVRCInitializing = new System.Windows.Forms.Timer(components);
            contextMenuStrip1.SuspendLayout();
            groupBox1.SuspendLayout();
            SuspendLayout();
            // 
            // notifyIcon1
            // 
            notifyIcon1.ContextMenuStrip = contextMenuStrip1;
            notifyIcon1.Icon = (Icon)resources.GetObject("notifyIcon1.Icon");
            notifyIcon1.Text = "YoutubeIPv6BlockForVRChat";
            notifyIcon1.Visible = true;
            // 
            // contextMenuStrip1
            // 
            contextMenuStrip1.Items.AddRange(new ToolStripItem[] { toolStripMenuItem1, toolStripSeparator1, ShowMenu, IPv6Block, Exit });
            contextMenuStrip1.Name = "contextMenuStrip1";
            contextMenuStrip1.Size = new Size(225, 98);
            // 
            // toolStripMenuItem1
            // 
            toolStripMenuItem1.Enabled = false;
            toolStripMenuItem1.Name = "toolStripMenuItem1";
            toolStripMenuItem1.Size = new Size(224, 22);
            toolStripMenuItem1.Text = "YoutubeIPv6BlockForVRChat";
            // 
            // toolStripSeparator1
            // 
            toolStripSeparator1.Name = "toolStripSeparator1";
            toolStripSeparator1.Size = new Size(221, 6);
            // 
            // ShowMenu
            // 
            ShowMenu.Name = "ShowMenu";
            ShowMenu.Size = new Size(224, 22);
            ShowMenu.Text = "メニュー画面表示";
            ShowMenu.Click += ShowMenu_Click;
            // 
            // IPv6Block
            // 
            IPv6Block.Name = "IPv6Block";
            IPv6Block.Size = new Size(224, 22);
            IPv6Block.Text = "IPv6ブロック";
            IPv6Block.Click += IPv6Block_Click;
            // 
            // Exit
            // 
            Exit.Name = "Exit";
            Exit.Size = new Size(224, 22);
            Exit.Text = "アプリ終了";
            Exit.Click += Exit_Click;
            // 
            // checkBoxAutoStart
            // 
            checkBoxAutoStart.AutoSize = true;
            checkBoxAutoStart.Location = new Point(12, 73);
            checkBoxAutoStart.Name = "checkBoxAutoStart";
            checkBoxAutoStart.Size = new Size(194, 19);
            checkBoxAutoStart.TabIndex = 1;
            checkBoxAutoStart.Text = "Windows起動時にアプリ自動起動";
            checkBoxAutoStart.UseVisualStyleBackColor = true;
            checkBoxAutoStart.CheckedChanged += checkBoxAutoStart_CheckedChanged;
            // 
            // buttonBlock
            // 
            buttonBlock.Location = new Point(6, 22);
            buttonBlock.Name = "buttonBlock";
            buttonBlock.Size = new Size(75, 23);
            buttonBlock.TabIndex = 2;
            buttonBlock.Text = "ブロック";
            buttonBlock.UseVisualStyleBackColor = true;
            buttonBlock.Click += buttonBlock_Click;
            // 
            // buttonUnblock
            // 
            buttonUnblock.Enabled = false;
            buttonUnblock.Location = new Point(87, 22);
            buttonUnblock.Name = "buttonUnblock";
            buttonUnblock.Size = new Size(75, 23);
            buttonUnblock.TabIndex = 2;
            buttonUnblock.Text = "ブロック解除";
            buttonUnblock.UseVisualStyleBackColor = true;
            buttonUnblock.Click += buttonUnblock_Click;
            // 
            // groupBox1
            // 
            groupBox1.Controls.Add(buttonBlock);
            groupBox1.Controls.Add(buttonUnblock);
            groupBox1.Location = new Point(12, 12);
            groupBox1.Name = "groupBox1";
            groupBox1.Size = new Size(179, 55);
            groupBox1.TabIndex = 4;
            groupBox1.TabStop = false;
            groupBox1.Text = "Youtube IPv6ブロック手動制御";
            // 
            // timerCheckVRCRunning
            // 
            timerCheckVRCRunning.Interval = 1000;
            timerCheckVRCRunning.Tick += timerCheckVRCRunning_Tick;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(294, 74);
            label1.Name = "label1";
            label1.Size = new Size(34, 15);
            label1.TabIndex = 5;
            label1.Text = "v1.02";
            // 
            // timerCheckVRCInitializing
            // 
            timerCheckVRCInitializing.Interval = 1000;
            timerCheckVRCInitializing.Tick += timerCheckVRCInitializing_Tick;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(340, 98);
            Controls.Add(label1);
            Controls.Add(groupBox1);
            Controls.Add(checkBoxAutoStart);
            FormBorderStyle = FormBorderStyle.Fixed3D;
            Icon = (Icon)resources.GetObject("$this.Icon");
            MaximizeBox = false;
            Name = "Form1";
            Text = "YoutubeIPv6BlockForVRChat";
            FormClosing += Form1_FormClosing;
            Shown += Form1_Shown;
            SizeChanged += Form1_SizeChanged;
            contextMenuStrip1.ResumeLayout(false);
            groupBox1.ResumeLayout(false);
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion
        private NotifyIcon notifyIcon1;
        private ContextMenuStrip contextMenuStrip1;
        private ToolStripMenuItem ShowMenu;
        private ToolStripMenuItem IPv6Block;
        private ToolStripMenuItem Exit;
        private CheckBox checkBoxAutoStart;
        private Button buttonBlock;
        private Button buttonUnblock;
        private GroupBox groupBox1;
        private System.Windows.Forms.Timer timerCheckVRCRunning;
        private ToolStripSeparator toolStripSeparator1;
        private ToolStripMenuItem toolStripMenuItem1;
        private Label label1;
        private System.Windows.Forms.Timer timerCheckVRCInitializing;
    }
}
