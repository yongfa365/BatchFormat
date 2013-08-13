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
    [ProvideOptionPageAttribute(typeof(OptionPage), "BatchFormat", "General", 101, 106, true)]
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [Guid(GuidList.GuidBatchFormatPkgString)]
    public sealed class BatchFormatPackage : Package
    {
        private DTE2 dte = null;
        private PkgCmdIDList selectedMenu = PkgCmdIDList.Null;
        private List<string> lstAlreadyOpenFiles = new List<string>();
        private List<string> lstExcludeEndsWith = null;
        private OutputWindowPane myOutPane = null;
        private int count;

        protected override void Initialize()
        {
            base.Initialize();
            dte = (DTE2)base.GetService(typeof(DTE));
            myOutPane = dte.ToolWindows.OutputWindow.OutputWindowPanes.Add("BatchFormat");

            var mcs = GetService(typeof(IMenuCommandService)) as MenuCommandService;
            if (null != mcs)
            {
                foreach (int item in Enum.GetValues(typeof(PkgCmdIDList)))
                {
                    var menuCommandID = new CommandID(GuidList.GuidBatchFormatCmdSet, item);
                    var menuItem = new MenuCommand(Excute, menuCommandID);
                    mcs.AddCommand(menuItem);
                }
            }
        }


        private void Excute(object sender, EventArgs e)
        {
            lstExcludeEndsWith = dte.GetValue("ExcludeEndsWith");
            count = 0;

            selectedMenu = (PkgCmdIDList)((MenuCommand)sender).CommandID.ID;

            var table = new RunningDocumentTable(this);
            lstAlreadyOpenFiles = (from info in table select info.Moniker).ToList<string>();

            var selectedItem = dte.SelectedItems.Item(1);
            WriteLog("\r\n====================================================================================");
            WriteLog("Start ：" + DateTime.Now.ToString());

            Stopwatch sp = new Stopwatch();
            sp.Start();

            if (selectedItem.Project == null && selectedItem.ProjectItem == null)
            {
                ProcessSolution();
            }
            else if (selectedItem.Project != null)
            {
                ProcessProject(selectedItem.Project);
            }
            else if (selectedItem.ProjectItem != null)
            {
                if (selectedItem.ProjectItem.ProjectItems.Count > 0)
                {
                    ProcessProjectItem(selectedItem.ProjectItem);
                    ProcessProjectItems(selectedItem.ProjectItem.ProjectItems);
                }
                else
                {
                    ProcessProjectItem(selectedItem.ProjectItem);
                }
            }
            sp.Stop();

            WriteLog(string.Format("Finish：{0}  Times：{1}s  Files：{2}", DateTime.Now.ToString(), sp.ElapsedMilliseconds / 1000, count - 2));
            dte.ExecuteCommand("View.Output");
            myOutPane.Activate();

        }


        private void ProcessSolution()
        {
            if (dte.Solution != null)
            {
                var projects = (from prj in new ProjectIterator(dte.Solution)
                                where prj.Kind == GuidList.GuidCsharpProjectString //只处理C#项目
                                select prj);

                projects.ForEach(prj => ProcessProject(prj));
            }
        }

        private void ProcessProject(Project project)
        {
            if (project != null)
            {
                ProcessProjectItems(project.ProjectItems);
            }
        }

        private void ProcessProjectItems(ProjectItems projectItems)
        {
            if (projectItems != null)
            {
                new ProjectItemIterator(projectItems).ForEach(item => ProcessProjectItem(item));
            }
        }

        private void ProcessProjectItem(ProjectItem projectItem)
        {
            string fileName;
            if (projectItem != null)
            {
                fileName = projectItem.FileNames[1];
                if (IsNotNeedProcess(projectItem))
                {
                    return;
                }

                WriteLog("Doing ：" + fileName);

                Window window = dte.OpenFile("{7651A703-06E5-11D1-8EBD-00A0C90F26EA}", fileName);
                window.Activate();
                #region 执行命令
                try
                {
                    switch (selectedMenu)
                    {
                        case PkgCmdIDList.cmdidRemoveUnusedUsings:
                            dte.ExecuteCommand("Edit.RemoveUnusedUsings");
                            break;
                        case PkgCmdIDList.cmdidSortUsings:
                            dte.ExecuteCommand("Edit.SortUsings");
                            break;
                        case PkgCmdIDList.cmdidRemoveAndSortUsings:
                            dte.ExecuteCommand("Edit.RemoveAndSort");
                            break;
                        case PkgCmdIDList.cmdidFormatDocument:
                            dte.ExecuteCommand("Edit.FormatDocument");
                            break;
                        case PkgCmdIDList.cmdidSortUsingsAndFormatDocument:
                            dte.ExecuteCommand("Edit.SortUsings");
                            dte.ExecuteCommand("Edit.FormatDocument");
                            break;
                        case PkgCmdIDList.cmdidRemoveAndSortUsingsAndFormatDocument:
                            dte.ExecuteCommand("Edit.RemoveAndSort");
                            dte.ExecuteCommand("Edit.FormatDocument");
                            break;
                        default:
                            break;
                    }

                }
                catch (COMException)
                {
                }
                #endregion

                if (lstAlreadyOpenFiles.Exists(file => file.Equals(fileName, StringComparison.OrdinalIgnoreCase)))
                {
                    dte.ActiveDocument.Save(fileName);
                }
                else
                {
                    window.Close(vsSaveChanges.vsSaveChangesYes);
                }
            }
        }

        private bool IsNotNeedProcess(ProjectItem projectItem)
        {
            if (projectItem.FileCodeModel == null)
            {
                return true;
            }

            var input = projectItem.FileNames[0];

            foreach (var item in lstExcludeEndsWith)
            {
                if (input.EndsWith(item, true, null))
                {
                    return true;
                }
            }
            return false;
        }


        private void WriteLog(string log)
        {
            dte.StatusBar.Text = "BatchFormat " + log;
            myOutPane.OutputString(log + "\r\n");
            count++;
        }

    }
}
