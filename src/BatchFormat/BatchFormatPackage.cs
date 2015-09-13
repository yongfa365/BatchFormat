using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;

namespace YongFa365.BatchFormat
{
    [PackageRegistration(UseManagedResourcesOnly = true)]
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)]
    [ProvideOptionPage(typeof(OptionPage), "BatchFormat", "General", 101, 106, true)]
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [Guid(GuidList.GuidBatchFormatPkgString)]
    public sealed class BatchFormatPackage : Package
    {
        private DTE2 _dte;
        private PkgCmdIdList _selectedMenu = PkgCmdIdList.Null;
        private List<string> _lstAlreadyOpenFiles = new List<string>();
        private List<string> _lstExcludePath;
        private OutputWindowPane _myOutPane;
        private int _count;

        protected override void Initialize()
        {
            base.Initialize();
            _dte = (DTE2)GetService(typeof(DTE));
            _myOutPane = _dte.ToolWindows.OutputWindow.OutputWindowPanes.Add("BatchFormat");

            var mcs = GetService(typeof(IMenuCommandService)) as MenuCommandService;

            if (null == mcs) return;

            foreach (int item in Enum.GetValues(typeof(PkgCmdIdList)))
            {
                var menuCommandId = new CommandID(GuidList.GuidBatchFormatCmdSet, item);
                var menuItem = new MenuCommand(Excute, menuCommandId);
                mcs.AddCommand(menuItem);
            }
        }

        private void Excute(object sender, EventArgs e)
        {
            _myOutPane.Activate();
            _lstExcludePath = _dte.GetExcludePathList();
            _count = 0;

            _selectedMenu = (PkgCmdIdList)((MenuCommand)sender).CommandID.ID;

            var table = new RunningDocumentTable(this);
            _lstAlreadyOpenFiles = table.Select(p => p.Moniker).ToList();

            var selectedItem = _dte.SelectedItems.Item(1);

            WriteLog($"{Environment.NewLine}====================================================================================");
            WriteLog($"Start: {DateTime.Now}");

           
            var sp = new Stopwatch();
            sp.Start();

            //以下判断只会进一次
            if (selectedItem.Project == null && selectedItem.ProjectItem == null) //从解决方案一级进来的
            {
                ProcessSolution();
            }
            else if (selectedItem.Project != null) //从项目一级进来的
            {
                ProcessProject(selectedItem.Project);
            }
            else if (selectedItem.ProjectItem != null) //从文件夹一级或者文件一级进来的
            {
                if (selectedItem.ProjectItem.ProjectItems.Count > 0) //此项下页还有文件，如：文件夹，T4模板
                {
                    ProcessProjectItem(selectedItem.ProjectItem);
                    ProcessProjectItems(selectedItem.ProjectItem.ProjectItems, selectedItem.ProjectItem.Name);
                }
                else //此项下什么都没有了，直接处理
                {
                    ProcessProjectItem(selectedItem.ProjectItem);
                }
            }
            sp.Stop();

            WriteLog($"Finish: {DateTime.Now}  Files: {_count - 2}  Elapsed: {sp.ElapsedMilliseconds / 1000}s");
            _dte.ExecuteCommand("View.Output");
            _myOutPane.Activate();

        }

        private void ProcessSolution()
        {
            if (_dte.Solution != null)
            {
                var projects = (from prj in new ProjectIterator(_dte.Solution)
                                where prj.Kind == GuidList.GuidCsharpProjectString //只处理C#项目
                                select prj);

                projects.ForEach(ProcessProject);
            }
        }

        private void ProcessProject(Project project)
        {
            if (project != null)
            {
                ProcessProjectItems(project.ProjectItems);
            }
        }

        private void ProcessProjectItems(ProjectItems projectItems, string parentFileName = null)
        {
            if (projectItems != null)
            {
                new ProjectItemIterator(projectItems).ForEach(item => ProcessProjectItem(item, parentFileName));
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="projectItem"></param>
        /// <param name="parentFileName"> //TODO:because can't get projectItem.Parent,so give it</param>
        private void ProcessProjectItem(ProjectItem projectItem, string parentFileName = null)
        {
            _myOutPane.Activate();
            string fileName;
            if (projectItem != null)
            {
                fileName = projectItem.FileNames[1];
                if (IsNotNeedProcess(projectItem, parentFileName))
                {
                    return;
                }

                WriteLog($"Doing: {fileName}");

                var window = _dte.OpenFile("{7651A703-06E5-11D1-8EBD-00A0C90F26EA}", fileName);
                window.Activate();
                #region 执行命令
                try
                {
                    switch (_selectedMenu)
                    {
                        case PkgCmdIdList.CmdidRemoveUnusedUsings:
                            _dte.ExecuteCommand("Edit.RemoveUnusedUsings");
                            break;
                        case PkgCmdIdList.CmdidSortUsings:
                            _dte.ExecuteCommand("Edit.SortUsings");
                            break;
                        case PkgCmdIdList.CmdidRemoveAndSortUsings:
                            _dte.ExecuteCommand("Edit.RemoveAndSort");
                            break;
                        case PkgCmdIdList.CmdidFormatDocument:
                            _dte.ExecuteCommand("Edit.FormatDocument");
                            break;
                        case PkgCmdIdList.CmdidSortUsingsAndFormatDocument:
                            _dte.ExecuteCommand("Edit.SortUsings");
                            _dte.ExecuteCommand("Edit.FormatDocument");
                            break;
                        case PkgCmdIdList.CmdidRemoveAndSortUsingsAndFormatDocument:
                            _dte.ExecuteCommand("Edit.RemoveAndSort");
                            _dte.ExecuteCommand("Edit.FormatDocument");
                            break;
                        case PkgCmdIdList.Null:
                            break;
                        default:
                            throw new ArgumentOutOfRangeException();
                    }
                }
                catch (COMException)
                {
                }

                #endregion

                if (_lstAlreadyOpenFiles.Exists(file => file.Equals(fileName, StringComparison.OrdinalIgnoreCase)))
                {
                    _dte.ActiveDocument.Save(fileName);
                }
                else
                {
                    window.Close(vsSaveChanges.vsSaveChangesYes);
                }
            }
        }


        private bool IsNotNeedProcess(ProjectItem projectItem, string parentFileName)
        {
            if (projectItem.FileCodeModel == null)
            {
                return true;
            }
            if (parentFileName != null)
            {
                if (_dte.IsIgnoreT4Child() && parentFileName.ToLower().EndsWith(".tt"))
                {
                    return true;
                }
            }

            var input = projectItem.FileNames[1];

            return _lstExcludePath.Any(p => input.ToLower().Contains(p.ToLower()));
        }

        private void WriteLog(string log)
        {
            _dte.StatusBar.Text = $"BatchFormat {log}";
            _myOutPane.OutputString($"{log}{Environment.NewLine}");
            _count++;
        }
    }
}
