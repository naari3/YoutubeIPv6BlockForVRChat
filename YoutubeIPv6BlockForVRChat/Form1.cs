using Microsoft.Win32.TaskScheduler;
using System.Diagnostics;


namespace YoutubeIPv6BlockForVRChat
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private bool isInitializing = true;
        private void Form1_Shown(object sender, EventArgs e)
        {
            Debug.WriteLine("App Init Start");
            WindowState = FormWindowState.Minimized;

            //残存ファイアウォールルールのクリーンアップ
            DeleteFirewallRule();

            //タスク有無を確認し、タスクが登録済みの場合はスタートアップのチェックボックスON
            if (IsTaskExist())
            {
                checkBoxAutoStart.Checked = true;
            }

            //VRC起動監視 開始
            timerCheckVRCInitializing.Start();

            isInitializing = false;
            Debug.WriteLine("App Init Finish");
        }

        private void buttonBlock_Click(object sender, EventArgs e)
        {
            if (IsVRCRunning())
            {
                MessageBox.Show("VRChat起動中は操作できません", Properties.Resources.AppName);
                return;
            }
            ExecuteBlock();
        }

        private void buttonUnblock_Click(object sender, EventArgs e)
        {
            if (IsVRCRunning())
            {
                MessageBox.Show("VRChat起動中は操作できません", Properties.Resources.AppName);
                return;
            }
            ExecuteUnblock();
        }

        private void checkBoxAutoStart_CheckedChanged(object sender, EventArgs e)
        {
            if (isInitializing) return;

            var Checkbox = (CheckBox)sender;
            if (Checkbox.Checked)
            {
                CreateTask();
            }
            else
            {
                DeleteTask();
            }
        }

        private void timerCheckVRCInitializing_Tick(object sender, EventArgs e)
        {
            if (!IsVRChatInitializing())
            {
                return;
            }
            Debug.WriteLine("VRC Detect");
            timerCheckVRCInitializing.Stop();
            ExecuteBlock();
            timerCheckVRCRunning.Start();
        }

        private void timerCheckVRCRunning_Tick(object sender, EventArgs e)
        {
            if (IsVRChatInitializing())
            {
                return;
            }
            if (IsVRCRunning())
            {
                return;
            }
            Debug.WriteLine("VRC Shutdown");
            timerCheckVRCRunning.Stop();
            ExecuteUnblock();
            timerCheckVRCInitializing.Start();
        }

        private void CreateTask()
        {
            Debug.WriteLine("Create Task");
            var TaskName = Properties.Resources.AppName;
            var TaskExeFile = Environment.ProcessPath;
            using (TaskService TaskData = new TaskService())
            {
                TaskDefinition TaskDefine = TaskData.NewTask();
                TaskDefine.Principal.RunLevel = TaskRunLevel.Highest;
                TaskDefine.Principal.LogonType = TaskLogonType.InteractiveToken;
                TaskDefine.Actions.Add(new ExecAction(TaskExeFile));
                TaskDefine.Triggers.Add(new LogonTrigger());
                TaskDefine.RegistrationInfo.Author = "NyaHo";
                TaskDefine.Settings.DisallowStartIfOnBatteries = false;
                TaskData.RootFolder.RegisterTaskDefinition(TaskName, TaskDefine, TaskCreation.CreateOrUpdate, null, null, TaskLogonType.InteractiveToken, null);
            }
            return;
        }

        private void DeleteTask()
        {
            Debug.WriteLine("Delete Task");
            var TaskName = Properties.Resources.AppName;
            var TaskService = new TaskService();
            TaskService.RootFolder.DeleteTask(TaskName);
        }

        private bool IsTaskExist()
        {
            using (var TaskService = new TaskService())
            {
                var Task = TaskService.FindTask(Properties.Resources.AppName);
                return Task != null;
            }
        }

        private void ExecuteBlock()
        {
            Debug.WriteLine("Execute Block");
            CreateFirewallRule();
            buttonBlock.Enabled = false;
            buttonUnblock.Enabled = true;
            IPv6Block.Checked = true;
        }

        private void ExecuteUnblock()
        {
            Debug.WriteLine("Execute Unblock");
            DeleteFirewallRule();
            buttonBlock.Enabled = true;
            buttonUnblock.Enabled = false;
            IPv6Block.Checked = false;
        }

        private void CreateFirewallRule()
        {
            Debug.WriteLine("Create FirewallRule");
            ExecutePowerShellCommand($"New-NetFirewallRule -DisplayName \"{Properties.Resources.AppName}\" -Direction Outbound -Action Block -RemoteAddress \"2404:6800::/32\" -Profile Any -Protocol Any -Enabled True");
        }

        private void DeleteFirewallRule()
        {
            Debug.WriteLine("Delete FirewallRule");
            ExecutePowerShellCommand($"Remove-NetFirewallRule -DisplayName \"{Properties.Resources.AppName}\"");
        }

        private void IPv6Block_Click(object sender, EventArgs e)
        {
            if (IsVRCRunning())
            {
                MessageBox.Show("VRChat起動中は操作できません", Properties.Resources.AppName);
                return;
            }

            if (IPv6Block.Checked)
            {
                buttonUnblock_Click(sender, e);
            }
            else
            {
                buttonBlock_Click(sender, e);
            }
        }

        private void ShowMenu_Click(object sender, EventArgs e)
        {
            WindowState = FormWindowState.Normal;
        }

        private void Exit_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void Form1_SizeChanged(object sender, EventArgs e)
        {
            if (WindowState == FormWindowState.Minimized)
            {
                ShowInTaskbar = false;
            }
            else if (WindowState == FormWindowState.Normal)
            {
                ShowInTaskbar = true;
            }
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            //残存ファイアウォールルールのクリーンアップ
            DeleteFirewallRule();
        }

        private bool IsVRCRunning()
        {
            return Process.GetProcessesByName("VRChat").Any();
        }

        private bool IsVRChatInitializing()
        {
            var targetProcessName = "start_protected_game";
            try
            {
                var procs = Process.GetProcessesByName(targetProcessName);
                var proc = procs.FirstOrDefault();
                if (proc != null && proc.MainModule != null)
                {
                    var processFilePath = proc.MainModule.FileName;
                    var parentDirName = Directory.GetParent(processFilePath)?.Name;
                    if (parentDirName == "VRChat")
                    {
                        return true;
                    }
                    else
                    {
                        Debug.WriteLine("No VRChat EAC");
                        return false;
                    }
                }
                else
                {
                    return false;
                }
            }
            catch (Exception)
            {
                return false;
            }
        }

        private string ExecutePowerShellCommand(string psCommandWithArgs)
        {
            var psInfo = new ProcessStartInfo();

            psInfo.FileName = @"PowerShell.exe";
            psInfo.CreateNoWindow = true;
            psInfo.WindowStyle = ProcessWindowStyle.Hidden;
            psInfo.UseShellExecute = false;
            psInfo.Arguments = psCommandWithArgs;
            psInfo.RedirectStandardOutput = true; // 標準出力をリダイレクト
            psInfo.RedirectStandardError = true;  // 標準エラー出力をリダイレクト

            var p = Process.Start(psInfo);
            if (p == null)
            {
                throw new InvalidOperationException("PowerShell プロセスの開始に失敗しました。");
            }

            var Result = p.StandardOutput.ReadToEnd();   // 標準出力の読み取り
            return Result;
        }
    }
}
