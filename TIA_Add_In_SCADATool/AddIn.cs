using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Siemens.Engineering;
using Siemens.Engineering.AddIn.Menu;
using Siemens.Engineering.HW;
using Siemens.Engineering.HW.Features;
using Siemens.Engineering.Online;
using Siemens.Engineering.SW;
using Siemens.Engineering.SW.Blocks;

namespace TIA_Add_In_SCADATool
{
    public class AddIn : ContextMenuAddIn
    {
        #region Definition

        /// <summary>
        /// 博图实例1
        /// </summary>
        private readonly TiaPortal _tiaPortal;

        /// <summary>
        /// Base class for projects
        /// can be use in multi-user environment
        /// </summary>
        private ProjectBase _projectBase;

        /// <summary>
        /// Path of the project file
        /// </summary>
        private string _projectPath;

        /// <summary>
        /// Path of the project directory
        /// </summary>
        private string _projectDir;

        /// <summary>
        /// 导出文件夹信息
        /// </summary>
        private DirectoryInfo _exportDirInfo;

        /// <summary>
        /// 获取PLC目标
        /// </summary>
        private PlcSoftware _plcSoftware;

        /// <summary>
        /// 获取实例的DB名称
        /// </summary>
        private string _instanceName;

        /// <summary>
        /// 获取实例的FB名称
        /// </summary>
        private string _instanceOfName;

        /// <summary>
        /// 获取DB的编号
        /// </summary>
        private string _number;

        /// <summary>
        /// 获取程序语言类型
        /// </summary>
        private string _programmingLanguage;

        /// <summary>
        /// 获取InstanceDB Xml文件路径
        /// </summary>
        private string _iDBXmlFilePath;

        /// <summary>
        /// 获取FB Xml文件路径
        /// </summary>
        private string _fBXmlFilePath;

        /// <summary>
        /// 写入数据流
        /// </summary>
        public StreamWriter _streamWriter;

        #endregion
        public AddIn(TiaPortal tiaPortal) : base("SCADA工具")
        {
            _tiaPortal = tiaPortal;
        }

        protected override void BuildContextMenuItems(ContextMenuAddInRoot addInRootSubmenu)
        {
            addInRootSubmenu.Items.AddActionItem<InstanceDB>("生成数据", Generate_OnClick);
            addInRootSubmenu.Items.AddActionItem<IEngineeringObject>("如需使用，请选中实例数据块", menuSelectionProvider => { }, GenerateStatus);
        }

        private void Generate_OnClick(MenuSelectionProvider<InstanceDB> menuSelectionProvider)
        {
            //TODO 确定PLC的在线状态
            if (!IsOffline())
            {
                throw new Exception(string.Format(CultureInfo.InvariantCulture,
                            "PLC 在线状态！"));
            }
            //TODO 获取项目数据
            GetProjectData();
            try
            {
                //TODO 打开窗口获取.csv文件保存位置
                FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();
                if (folderBrowserDialog.ShowDialog(new Form()
                { TopMost = true, WindowState = FormWindowState.Maximized }) == DialogResult.OK)
                {
                    //TODO 独占窗口
                    using (ExclusiveAccess exclusiveAccess = _tiaPortal.ExclusiveAccess("导出中……"))
                    {
                        //定义csv数据流
                        _streamWriter = new StreamWriter(Path.Combine(_projectDir, "SCADA.csv"));
                        //写入标题行
                        _streamWriter.WriteLine("\"地址\",\"数据类别0状态1错误2警告\",\"报警文本en-US\",\"报警文本zh-CN\"");

                        foreach (InstanceDB instanceDB in menuSelectionProvider.GetSelection())
                        {
                            if (exclusiveAccess.IsCancellationRequested)
                            {
                                break;
                            }

                            //TODO 获取InstanceName
                            _instanceName = instanceDB.Name;
                            //TODO 获取InstanceOfName
                            _instanceOfName = instanceDB.InstanceOfName;
                            //TODO 获取Number
                            _number = instanceDB.Number.ToString();
                            //TODO 获取ProgrammingLanguage
                            _programmingLanguage = instanceDB.ProgrammingLanguage.ToString();

                            //创建导出文件夹
                            string exportFileDir = Path.Combine(_projectDir, "SCADA");
                            _exportDirInfo = Directory.CreateDirectory(exportFileDir);
                            if (Directory.Exists(exportFileDir))
                            {
                                //TODO 导出InstanceDB
                                _iDBXmlFilePath = Path.Combine(exportFileDir, StringHandle(_instanceName), ".xml");
                                exclusiveAccess.Text = "导出中-> " + Export(instanceDB, _iDBXmlFilePath);

                                //判断文件夹中是否已包含
                                _fBXmlFilePath = Path.Combine(exportFileDir, StringHandle(_instanceOfName), ".xml");
                                string[] fileNames = Directory.GetFiles(exportFileDir, Path.GetFileName(_fBXmlFilePath));
                                if (fileNames.Length < 0)
                                {
                                    //TODO 导出FB
                                    exclusiveAccess.Text = "导出中-> " + Export(GetFB(), _fBXmlFilePath);
                                }
                            }

                            //TODO 处理Xml
                            XmlEditor xmlEditor = new XmlEditor
                            {
                                _fBXmlFilePath = _fBXmlFilePath,
                                _iDBXmlFilePath = _iDBXmlFilePath,
                                _instanceName = _instanceName,
                                _number = _number,
                                _programmingLanguage = _programmingLanguage,
                                _instanceOfName = _instanceOfName,
                                _streamWriter = _streamWriter
                            };
                        }
                        //TODO 写入csv数据流
                        _streamWriter.Close();
                        //TODO 导出完成确认是否打开导出文件夹
                        DialogResult dialogResult = MessageBox.Show("是否打开导出文件夹？", "导出完成", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (dialogResult == DialogResult.Yes)
                        {
                            if (Directory.Exists(_projectDir))
                            {
                                Process.Start("explorer.exe",_projectDir);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "异常", MessageBoxButtons.OK, MessageBoxIcon.Error);
                throw;
            }
            finally
            {
                //TODO 删除导出的文件夹
                _exportDirInfo.Delete();
            }
        }

        /// <summary>
        /// PLC是否为离线模式
        /// </summary>
        /// <param name="item"></param>
        /// <returns></returns>
        private bool IsOffline()
        {
            bool ret = false;

            foreach (Device device in _projectBase.Devices)
            {
                DeviceItem deviceItem = device.DeviceItems[1];
                if (deviceItem.GetAttribute("Classification") is DeviceItemClassifications.CPU)
                {
                    OnlineProvider onlineProvider = deviceItem.GetService<OnlineProvider>();
                    ret = (onlineProvider.State == OnlineState.Offline);
                }
            }

            return ret;
        }

        private FB GetFB()
        {
            PlcBlockGroup plcBlockGroups = _plcSoftware.BlockGroup;
            foreach (FB fB in plcBlockGroups.Blocks)
            {
                if (fB.Name == _instanceOfName)
                {
                    return fB;
                }
            }
            foreach (PlcBlockGroup plcBlockGroup in plcBlockGroups.Groups)
            {
                return EnumerateAllBlocks(plcBlockGroup);
            }

            return null;
        }

        private FB EnumerateAllBlocks(PlcBlockGroup blockGroup)
        {
            foreach (PlcBlockUserGroup subBlockGroup in blockGroup.Groups)
            {
                foreach (FB fB in subBlockGroup.Blocks)
                {
                    if (fB.Name == _instanceOfName)
                    {
                        return fB;
                    }
                }
                EnumerateAllBlocks(subBlockGroup);
            }

            return null;
        }

        // Returns PlcSoftware
        private void GetPlcSoftware()
        {
            foreach (Device device in _projectBase.Devices)
            {
                if (device.DeviceItems[1].GetAttribute("Classification") is DeviceItemClassifications.CPU)
                {
                    DeviceItemComposition deviceItemComposition = device.DeviceItems;
                    foreach (DeviceItem deviceItem in deviceItemComposition)
                    {
                        SoftwareContainer softwareContainer = deviceItem.GetService<SoftwareContainer>();
                        if (softwareContainer != null)
                        {
                            Software softwareBase = softwareContainer.Software;
                            _plcSoftware = softwareBase as PlcSoftware;
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 字符串处理将"/"替换为"_"
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        private string StringHandle(string input)
        {
            //查询名称是否包含“/”，如果包含替换更“_”
            string ret = input;
            while (ret.Contains("/"))
            {
                ret = ret.Replace("/", "_");
            }
            return ret;
        }

        /// <summary>
        /// 导出PLC数据
        /// </summary>
        /// <param name="exportItem"></param>
        /// <param name="exportPath"></param>
        /// <returns></returns>
        private string Export(IEngineeringObject exportItem, string exportPath)
        {
            const ExportOptions exportOption = ExportOptions.WithDefaults & ExportOptions.WithReadOnly;

            switch (exportItem)
            {
                case PlcBlock item:
                    {
                        if (item.ProgrammingLanguage == ProgrammingLanguage.ProDiag ||
                            item.ProgrammingLanguage == ProgrammingLanguage.ProDiag_OB)
                            return null;
                        if (item.IsConsistent)
                        {
                            // filePath = Path.Combine(filePath, AdjustNames.AdjustFileName(GetObjectName(item)) + ".xml");
                            if (File.Exists(exportPath))
                            {
                                File.Delete(exportPath);
                            }

                            item.Export(new FileInfo(exportPath), exportOption);

                            return exportPath;
                        }

                        throw new EngineeringException(string.Format(CultureInfo.InvariantCulture,
                            "程序块: {0} 是不一致的! 请编译程序块! 导出将终止!", item.Name));
                    }
            }

            return null;
        }

        /// <summary>
        /// 获取ProjectBase：支持多用户项目
        /// </summary>
        private void GetProjectData()
        {
            try
            {
                // Multi-user support
                // If TIA Portal is in multi user environment (connected to project server)
                if (_tiaPortal.LocalSessions.Any())
                {
                    _projectBase = _tiaPortal.LocalSessions
                        .FirstOrDefault(s => s.Project != null && s.Project.IsPrimary)?.Project;
                }
                else
                {
                    // Get local project
                    _projectBase = _tiaPortal.Projects.FirstOrDefault(p => p.IsPrimary);
                }

                _projectPath = _projectBase?.Path.FullName;
                _projectDir = _projectBase?.Path.Directory?.FullName;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "异常", MessageBoxButtons.OK, MessageBoxIcon.Error);
                throw;
            }
        }

        /// <summary>
        /// 获取选中项类型关闭显示项目树按钮
        /// </summary>
        /// <param name="menuSelectionProvider"></param>
        /// <returns></returns>
        private static MenuStatus GenerateStatus(MenuSelectionProvider<IEngineeringObject> menuSelectionProvider)
        {
            var show = false;

            foreach (IEngineeringObject engineeringObject in menuSelectionProvider.GetSelection())
            {
                if (!(engineeringObject is InstanceDB))
                {
                    show = true;
                    break;
                }
            }
            return show ? MenuStatus.Disabled : MenuStatus.Hidden;
        }
    }
}