using Microsoft.Win32.TaskScheduler;
using System.Diagnostics;
using System.Net;
using System.Text.Json;


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
            ExecuteBlock();
        }

        private void buttonUnblock_Click(object sender, EventArgs e)
        {
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
            if (CreateFirewallRule())
            {
                buttonBlock.Enabled = false;
                buttonUnblock.Enabled = true;
                IPv6Block.Checked = true;
            }
        }

        private void ExecuteUnblock()
        {
            Debug.WriteLine("Execute Unblock");
            DeleteFirewallRule();
            buttonBlock.Enabled = true;
            buttonUnblock.Enabled = false;
            IPv6Block.Checked = false;
        }

        private bool CreateFirewallRule()
        {
            Debug.WriteLine("Create FirewallRule");
            
            var config = LoadNetblocksConfig();
            if (config == null || config.GoogleIPv6Netblocks == null || config.GoogleIPv6Netblocks.Count == 0)
            {
                MessageBox.Show("Google DNSの設定が見つかりません。先に「Google DNS取得」ボタンをクリックしてIPv6アドレス範囲を取得してください。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            
            var addresses = string.Join("\", \"", config.GoogleIPv6Netblocks);
            ExecutePowerShellCommand($"New-NetFirewallRule -DisplayName \"{Properties.Resources.AppName}\" -Direction Outbound -Action Block -RemoteAddress \"{addresses}\" -Profile Any -Protocol Any -Enabled True");
            
            Debug.WriteLine($"Blocked addresses: {addresses}");
            return true;
        }

        private void DeleteFirewallRule()
        {
            Debug.WriteLine("Delete FirewallRule");
            ExecutePowerShellCommand($"Remove-NetFirewallRule -DisplayName \"{Properties.Resources.AppName}\"");
        }

        private void IPv6Block_Click(object sender, EventArgs e)
        {
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

        private async void buttonFetchDNS_Click(object sender, EventArgs e)
        {
            try
            {
                buttonFetchDNS.Enabled = false;
                buttonFetchDNS.Text = "取得中...";
                
                var netblocks = await FetchGoogleNetblocks();
                
                if (netblocks != null && netblocks.Count > 0)
                {
                    SaveNetblocksConfig(netblocks);
                    MessageBox.Show($"Google DNSから{netblocks.Count}個のIPv6アドレス範囲を取得し、設定ファイルに保存しました。", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("Google DNSからIPv6アドレス範囲を取得できませんでした。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"エラーが発生しました: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                buttonFetchDNS.Enabled = true;
                buttonFetchDNS.Text = "Google DNS取得";
            }
        }

        private async System.Threading.Tasks.Task<List<string>> FetchGoogleNetblocks()
        {
            var netblocks = new List<string>();

            var spfCommand = "nslookup -type=TXT _spf.google.com 8.8.8.8";
            var spfResult = await System.Threading.Tasks.Task.Run(() => ExecutePowerShellCommand(spfCommand));

            var netblockDomains = ExtractNetblockDomains(spfResult);

            if (netblockDomains.Count == 0)
            {
                Debug.WriteLine("No netblock domains found in SPF record");
                return netblocks;
            }

            foreach (var domain in netblockDomains)
            {
                var nslookupCommand = $"nslookup -type=TXT {domain} 8.8.8.8";
                var result = await System.Threading.Tasks.Task.Run(() => ExecutePowerShellCommand(nslookupCommand));

                var lines = result.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                foreach (var line in lines)
                {
                    if (line.Contains("ip6:"))
                    {
                        var matches = System.Text.RegularExpressions.Regex.Matches(line, @"ip6:([0-9a-fA-F:]+\/\d+)");
                        foreach (System.Text.RegularExpressions.Match match in matches)
                        {
                            if (match.Success)
                            {
                                netblocks.Add(match.Groups[1].Value);
                                Debug.WriteLine($"Found IPv6 block from {domain}: {match.Groups[1].Value}");
                            }
                        }
                    }
                }
            }

            return netblocks.Distinct().ToList();
        }

        private List<string> ExtractNetblockDomains(string spfResult)
        {
            var domains = new List<string>();
            var lines = spfResult.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);

            foreach (var line in lines)
            {
                if (line.Contains("include:"))
                {
                    var matches = System.Text.RegularExpressions.Regex.Matches(line, @"include:(_netblocks\d*\.google\.com)");
                    foreach (System.Text.RegularExpressions.Match match in matches)
                    {
                        if (match.Success)
                        {
                            domains.Add(match.Groups[1].Value);
                            Debug.WriteLine($"Found netblock domain: {match.Groups[1].Value}");
                        }
                    }
                }
            }

            return domains;
        }

        private void SaveNetblocksConfig(List<string> netblocks)
        {
            var appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            var configDir = Path.Combine(appDataPath, "YoutubeIPv6BlockForVRChat");
            
            if (!Directory.Exists(configDir))
            {
                Directory.CreateDirectory(configDir);
            }
            
            var configPath = Path.Combine(configDir, "config.json");
            
            var config = new
            {
                LastUpdated = DateTime.Now,
                GoogleIPv6Netblocks = netblocks
            };
            
            var jsonOptions = new JsonSerializerOptions
            {
                WriteIndented = true
            };
            
            var json = JsonSerializer.Serialize(config, jsonOptions);
            File.WriteAllText(configPath, json);
            
            Debug.WriteLine($"Config saved to: {configPath}");
        }

        private NetblocksConfig? LoadNetblocksConfig()
        {
            var appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            var configPath = Path.Combine(appDataPath, "YoutubeIPv6BlockForVRChat", "config.json");
            
            if (!File.Exists(configPath))
            {
                return null;
            }
            
            try
            {
                var json = File.ReadAllText(configPath);
                return JsonSerializer.Deserialize<NetblocksConfig>(json);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Failed to load config: {ex.Message}");
                return null;
            }
        }
    }
}
