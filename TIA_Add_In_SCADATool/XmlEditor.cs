using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;

namespace TIA_Add_In_SCADATool
{
    public class XmlEditor
    {
        /// <summary>
        /// 获取Xml的命名空间
        /// </summary>
        private XmlNamespaceManager _xmlns;

        /// <summary>
        /// 获取实例的DB名称
        /// </summary>
        public string _instanceName;

        /// <summary>
        /// 获取实例的FB名称
        /// </summary>
        public string _instanceOfName;

        /// <summary>
        /// 获取DB的编号
        /// </summary>
        public string _number;

        /// <summary>
        /// 获取程序语言类型
        /// </summary>
        public string _programmingLanguage;

        /// <summary>
        /// 获取InstanceDB Xml文件路径
        /// </summary>
        public string _iDBXmlFilePath;

        /// <summary>
        /// 获取FB Xml文件路径
        /// </summary>
        public string _fBXmlFilePath;

        /// <summary>
        /// 写入数据流
        /// </summary>
        public StreamWriter _streamWriter;

        /// <summary>
        /// 获取监控列表
        /// </summary>
        private readonly List<string[]> _supervisions = new List<string[]>();

        public void Run()
        {
            //读取Xml文件
            XmlDocument xmlDocument = new XmlDocument();
            xmlDocument.Load(_iDBXmlFilePath);
            //获取"Sections"节点
            XmlNode sections = xmlDocument.GetElementsByTagName("Sections")[0];
            //获取"BlockInstSupervisionGroups"节点
            XmlNode bISupervisionGroups = xmlDocument.GetElementsByTagName("BlockInstSupervisionGroups")[0];
            //从Sections和BlockInstSupervisionGroups节点获取NameSpace
            _xmlns = GetXmlns(xmlDocument, sections.NamespaceURI, bISupervisionGroups.NamespaceURI, "");

            //获取FB的监控属性
            xmlDocument.Load(_fBXmlFilePath);
            //获取"BlockInstSupervisionGroups"节点
            XmlNode bTSupervisions = xmlDocument.GetElementsByTagName("BlockTypeSupervisions")[0];
            //从Sections、BlockInstSupervisionGroups、BlockTypeSupervisions节点获取NameSpace
            _xmlns = GetXmlns(xmlDocument, sections.NamespaceURI, bISupervisionGroups.NamespaceURI, bTSupervisions.NamespaceURI);
            //获取诊断变量和报警文本
            foreach (XmlNode supervisionGroup in bISupervisionGroups)
            {
                foreach (XmlNode supervision in supervisionGroup)
                {
                    string[] supervisioninfo = new string[4];
                    supervisioninfo[0] = GetName(GetSupervision(supervision));//DB监控定义
                    supervisioninfo[1] = GetSupervisionType(supervision);//DB监控类型
                    foreach (XmlNode bTSupervision in bTSupervisions)
                    {

                        if (GetName(GetSupervision(supervision)).Contains(GetName(GetSupervisedOperand(bTSupervision))))
                        {
                            supervisioninfo[2] = GetMultiLanguageText(bTSupervision, "en-US");//获取英文报警文本
                            supervisioninfo[3] = GetMultiLanguageText(bTSupervision, "zh-CN");//获取中文报警文本
                        }
                    }
                    _supervisions.Add(supervisioninfo);
                }
            }

            //获取变量和偏移量
            foreach (XmlNode members in sections)
            {
                foreach (XmlNode member in members)
                {
                    switch (GetName(member))
                    {
                        //获取运行状态变量
                        case "Q_M01_Forward": //正转运行状态
                            //string a = CalOffset(GetOffset(member));
                            SetSacdaTag(CalOffset(GetOffset(member)), "Q_M01_Forward");
                            break;
                        case "Q_M01_Reverse": //反转运行状态
                            //string b = CalOffset(GetOffset(member));
                            SetSacdaTag(CalOffset(GetOffset(member)), "Q_M01_Reverse");
                            break;
                        //提升机、顶升机
                        case "Q_M02_Forward":
                            SetSacdaTag(CalOffset(GetOffset(member)), "Q_M02_Forward");
                            break;
                        case "Q_M02_Reverse":
                            SetSacdaTag(CalOffset(GetOffset(member)), "Q_M02_Reverse");
                            break;
                        //拆盘机
                        case "Q_M01_Raise":
                            SetSacdaTag(CalOffset(GetOffset(member)), "Q_M01_Raise");
                            break;
                        case "Q_M01_Lower":
                            SetSacdaTag(CalOffset(GetOffset(member)), "Q_M01_Lower");
                            break;
                        case "Q_M02_ExtendLeftFork":
                            SetSacdaTag(CalOffset(GetOffset(member)), "Q_M02_ExtendLeftFork");
                            break;
                        case "Q_M02_RetractLeftFork":
                            SetSacdaTag(CalOffset(GetOffset(member)), "Q_M02_RetractLeftFork");
                            break;
                        case "Q_M03_ExtendRightFork":
                            SetSacdaTag(CalOffset(GetOffset(member)), "Q_M03_ExtendRightFork");
                            break;
                        case "Q_M03_RetractRightFork":
                            SetSacdaTag(CalOffset(GetOffset(member)), "Q_M03_RetractRightFork");
                            break;
                        //TODO 叠盘机托盘数量
                        case "R_DV_PalletStackCount":
                            SetSacdaTag(CalOffset(GetOffset(member)), "R_DV_PalletStackCount",true);
                            break;
                        case "Z_Data"://托盘数据
                            foreach (XmlNode memberOfZ_Data in GetSection(member))
                            {
                                switch (GetName(memberOfZ_Data))
                                {
                                    case "Position01":
                                        //int d = GetOffset(memberOfZ_Data);
                                        foreach (XmlNode memberOfPosition01 in GetSection(memberOfZ_Data))
                                        {
                                            switch (GetName(memberOfPosition01))
                                            {
                                                //获取原地址变量
                                                case "Origin": 
                                                    SetSacdaTag(CalOffset(GetOffset(member) + GetOffset(memberOfZ_Data) + GetOffset(memberOfPosition01)), "Z_Data_Position01_Origin");
                                                    break;
                                                //获取目的地变量
                                                case "Destination":
                                                    SetSacdaTag(CalOffset(GetOffset(member) + GetOffset(memberOfZ_Data) + GetOffset(memberOfPosition01)), "Z_Data_Position01_Destination");
                                                    break;
                                                case "PalletID":
                                                    //获取条码的字符长度
                                                    foreach (XmlNode item in GetSection(memberOfPosition01))
                                                    {
                                                        //Array[1..20] of Char
                                                        string barcodeType = item.Attributes.GetNamedItem("Datatype").Value;
                                                        SetSacdaTag(CalOffset(GetOffset(member) + GetOffset(memberOfZ_Data) + GetOffset(memberOfPosition01)), GetBarcodeLenth(barcodeType), "Z_Data_Position01_PalletID");
                                                    }
                                                    break;
                                                //托盘外检拒绝信息 
                                                case "RejectCode":
                                                    foreach (XmlNode item in GetSection(memberOfPosition01))
                                                    {
                                                        SetSacdaTag(CalOffset(GetOffset(member) + GetOffset(memberOfZ_Data) + GetOffset(memberOfPosition01) + GetOffset(item)), "Z_Data_Position01_" + GetName(item));
                                                    }
                                                    break;
                                            }
                                        }
                                        break;
                                    case "Position02":
                                        //int f = GetOffset(memberOfZ_Data);
                                        // 判断InstanceOfName实例决定是否需要占位2数据
                                        if (!_instanceOfName.Contains("_2P"))
                                            break;
                                        foreach (XmlNode memberOfPosition02 in GetSection(memberOfZ_Data))
                                        {
                                            switch (GetName(memberOfPosition02))
                                            {
                                                //获取原地址变量
                                                case "Origin":
                                                    SetSacdaTag(CalOffset(GetOffset(member) + GetOffset(memberOfZ_Data) + GetOffset(memberOfPosition02)), "Z_Data_Position02_Origin");
                                                    break;
                                                //获取目的地变量
                                                case "Destination":
                                                    SetSacdaTag(CalOffset(GetOffset(member) + GetOffset(memberOfZ_Data) + GetOffset(memberOfPosition02)), "Z_Data_Position02_Destination");
                                                    break;
                                                case "PalletID":
                                                    //f += GetOffset(memberOfPosition02);
                                                    //string g = CalOffset(c + f);
                                                    //获取条码的字符长度
                                                    foreach (XmlNode item in GetSection(memberOfPosition02))
                                                    {
                                                        //Array[1..20] of Char
                                                        string barcodeType = item.Attributes.GetNamedItem("Datatype").Value;
                                                        //string barcodeLenth = GetBarcodeLenth(barcodeType);
                                                        SetSacdaTag(CalOffset(GetOffset(member) + GetOffset(memberOfZ_Data) + GetOffset(memberOfPosition02)), GetBarcodeLenth(barcodeType), "Z_Data_Position02_PalletID");
                                                    }
                                                    break;
                                                //托盘外检拒绝信息 
                                                case "RejectCode":
                                                    foreach (XmlNode item in GetSection(memberOfPosition02))
                                                    {
                                                        SetSacdaTag(CalOffset(GetOffset(member) + GetOffset(memberOfZ_Data) + GetOffset(memberOfPosition02) + GetOffset(item)), "Z_Data_Position02_" + GetName(item));
                                                    }
                                                    break;
                                            }
                                        }
                                        break;
                                }
                            }
                            break;
                        case "Z_Status"://输送机状态
                            foreach (XmlNode memberOfZ_Status in GetSection(member))
                            {
                                switch (GetName(memberOfZ_Status))
                                {
                                    case "Fault":
                                        SetSacdaTag(CalOffset(GetOffset(member) + GetOffset(memberOfZ_Status)), "Fault");
                                        break;
                                    case "Occupied01":
                                        SetSacdaTag(CalOffset(GetOffset(member) + GetOffset(memberOfZ_Status)), "Occupied01");
                                        break;
                                    case "Occupied02":
                                        // 判断InstanceOfName实例决定是否需要占位2数据
                                        if (!_instanceOfName.Contains("_2P") & !_instanceOfName.Contains("_PStackUni_"))
                                            break;
                                        SetSacdaTag(CalOffset(GetOffset(member) + GetOffset(memberOfZ_Status)), "Occupied02");
                                        break;
                                    //叠盘机自动模式
                                    case "AutoModeEnabled":
                                        SetSacdaTag(CalOffset(GetOffset(member) + GetOffset(memberOfZ_Status)), "AutoModeEnabled");
                                        break;
                                    //叠盘机手动模式
                                    case "JogSelected":
                                        SetSacdaTag(CalOffset(GetOffset(member) + GetOffset(memberOfZ_Status)), "JogSelected");
                                        break;
                                }
                            }
                            break;
                    }
                    //获取诊断变量和报警文本
                    foreach (string[] supervision in _supervisions)
                    {
                        if (supervision[0].Contains(GetName(member)))
                        {
                            SetSacdaTag(CalOffset(GetOffset(member)), GetName(member), supervision[1], supervision[2], supervision[3]);
                        }
                    }
                }
            }
            //写入csv数据流
            //_streamWriter.Close();

            // 删除导出的Xml文件
        }

        /// <summary>
        /// 获取命名空间
        /// </summary>
        /// <param name="xmlDocument"></param>
        /// <param name="namespaceURI1">Sections命名空间</param>
        /// <param name="namespaceURI2">BlockInstSupervisionGroups命名空间</param>
        /// <returns>xmlns="http://www.siemens.com/automation/Openness/SW/Interface/v5"
        /// xmlns="http://www.siemens.com/automation/Openness/SW/BlockInstanceSupervisions/v3"</returns>
        private XmlNamespaceManager GetXmlns(XmlDocument xmlDocument, string namespaceURI1, string namespaceURI2, string namespaceURI3)
        {
            XmlNamespaceManager xmlns = new XmlNamespaceManager(xmlDocument.NameTable);
            xmlns.AddNamespace("x", namespaceURI1);
            xmlns.AddNamespace("y", namespaceURI2);
            xmlns.AddNamespace("z", namespaceURI3);
            return xmlns;
        }

        /// <summary>
        /// 获取偏移量节点
        /// </summary>
        /// <param name="xmlNode"></param>
        /// <returns>321</returns>
        private int GetOffset(XmlNode xmlNode)
        {
            XmlNode integerAttribute = xmlNode.SelectSingleNode("./x:AttributeList", _xmlns).FirstChild.FirstChild;
            //获取偏移量
            int offset = Convert.ToInt16(integerAttribute.Value);
            return offset;
        }

        /// <summary>
        /// 获取监控节点
        /// </summary>
        /// <param name="xmlNode"></param>
        /// <returns></returns>
        private XmlNode GetSupervision(XmlNode xmlNode)
        {
            XmlNode stateStruct = xmlNode.SelectSingleNode("./y:StateStruct", _xmlns);
            return stateStruct;
        }

        /// <summary>
        /// 获取监控类别节点
        /// </summary>
        /// <param name="xmlNode"></param>
        /// <returns></returns>
        private string GetSupervisionType(XmlNode xmlNode)
        {
            string str = xmlNode.SelectSingleNode("./y:BlockTypeSupervisionNumber", _xmlns).FirstChild.Value;
            return str;
        }

        /// <summary>
        /// 获取FB的监控定义
        /// </summary>
        /// <param name="xmlNode"></param>
        /// <returns></returns>
        private XmlNode GetSupervisedOperand(XmlNode xmlNode)
        {
            XmlNode supervisedOperand = xmlNode.SelectSingleNode("./z:SupervisedOperand", _xmlns);
            return supervisedOperand;
        }

        /// <summary>
        /// 获取FB的监控的文本定义
        /// </summary>
        /// <param name="xmlNode"></param>
        /// <param name="lang"></param>
        /// <returns></returns>
        private string GetMultiLanguageText(XmlNode xmlNode, string lang)
        {
            XmlNode specificField = xmlNode.SelectSingleNode("./z:SpecificField", _xmlns);
            XmlNode specificFieldText = specificField.SelectSingleNode("./z:SpecificFieldText", _xmlns);
            XmlNode multiLanguageText = specificFieldText.SelectSingleNode(string.Format("./z:MultiLanguageText[@Lang=\"{0}\"]", lang), _xmlns);
            //如果报警文本没有配置中文/英文，修复报错。
            string str = "";

            if (multiLanguageText == null)
            {
                return str;
            }
            else
            {
                str = multiLanguageText.FirstChild.Value;
            }
            return str;
        }

        /// <summary>
        /// 计算偏移量
        /// </summary>
        /// <param name="getOffset"></param>
        /// <returns>40.1</returns>
        private string CalOffset(int getOffset)
        {
            //计算偏移量
            string consult = (getOffset / 8).ToString(); //商，"40"
            string rem = (getOffset % 8).ToString(); //余数，"1"
            string offset = consult + "." + rem; //偏移量结果：40.1
            // Console.WriteLine(offset);

            return offset;
        }

        /// <summary>
        /// 获取Section
        /// </summary>
        /// <param name="xmlNode"></param>
        /// <returns></returns>
        private XmlNode GetSection(XmlNode xmlNode)
        {
            XmlNode section = xmlNode.SelectSingleNode("./x:Sections", _xmlns).SelectSingleNode("./x:Section", _xmlns);
            return section;
        }

        /// <summary>
        /// 获取节点的元素为Name的值
        /// </summary>
        /// <param name="xmlNode"></param>
        /// <returns>Name="Q_M01_Forward"</returns>
        private string GetName(XmlNode xmlNode)
        {
            string name = xmlNode.Attributes.GetNamedItem("Name").Value;
            return name;
        }

        /// <summary>
        /// 获取条码长度
        /// </summary>
        /// <param name="input"></param>
        /// <returns>20</returns>
        private string GetBarcodeLenth(string input)
        {

            //string input = "Array[1..20] of Char";
            // 查找 ".." 的位置
            int rangeIndex = input.IndexOf("..");

            if (rangeIndex != -1)
            {
                // 向前搜索数字的边界
                int startIndex = rangeIndex - 1;
                while (startIndex >= 0 && char.IsDigit(input[startIndex]))
                {
                    startIndex--;
                }
                startIndex++;

                // 向后搜索数字的边界
                int endIndex = rangeIndex + 2;
                while (endIndex < input.Length && char.IsDigit(input[endIndex]))
                {
                    endIndex++;
                }
                endIndex--;

                // 提取数字
                string beforeRange = input.Substring(startIndex, rangeIndex - startIndex);
                string afterRange = input.Substring(rangeIndex + 2, endIndex - rangeIndex - 1);

                int beforeNumber = int.Parse(beforeRange);
                int afterNumber = int.Parse(afterRange);

                // 计算差值
                string lenth = (afterNumber - beforeNumber + 1).ToString();
                //Console.WriteLine("前数字：" + beforeRange);
                //Console.WriteLine("后数字：" + afterRange);

                return lenth;
            }
            else
            {
                //Console.WriteLine("未找到匹配的范围。");
                return string.Empty;
            }
        }

        /// <summary>
        /// 设置SCADA标签值-Bool
        /// </summary>
        /// <param name="offset"></param>
        /// <param name="input"></param>
        /// <returns>"DB304,X40.0;P1005_Q_M01_Forward","0"</returns>
        private string SetSacdaTag(string offset, string input)
        {
            // BOOL：DB编号,X偏移量;标签
            string str = string.Format("\"{0}{1},X{2};{3}_{4}\",\"0\"", _programmingLanguage, _number, offset, _instanceName, input);
            _streamWriter.WriteLine(str);
            Console.WriteLine(str);
            return str;
        }

        /// <summary>
        /// 设置SCADA标签值-Int
        /// </summary>
        /// <param name="offset"></param>
        /// <param name="input"></param>
        /// <returns>"DB304,INT40;P1005_R_DV_PalletStackCountd","0"</returns>
        private string SetSacdaTag(string offset, string input, bool isInt)
        {
            string str = "";

            if (!isInt)
            {
                return str;
            }
            offset = offset.Replace(".0","");
            // Int：DB编号,INT偏移量;标签
            str = string.Format("\"{0}{1},INT{2};{3}_{4}\",\"0\"", _programmingLanguage, _number, offset, _instanceName, input);
            _streamWriter.WriteLine(str);
            Console.WriteLine(str);
            return str;
        }

        /// <summary>
        /// 设置SCADA标签值-字符串
        /// </summary>
        /// <param name="offset"></param>
        /// <param name="input"></param>
        /// <returns>"DB304,STRING178.20;P1005_Z_Data_Position01_PalletID","0"</returns>
        private string SetSacdaTag(string offset, string barcodeLenth, string input)
        {
            // String：DB编号,SATRING偏移量.字符长度;标签
            string str = string.Format("\"{0}{1},STRING{2}.{3};{4}_{5}\",\"0\"", _programmingLanguage, _number, offset.Replace(".0", ""), barcodeLenth, _instanceName, input);
            _streamWriter.WriteLine(str);
            Console.WriteLine(str);
            return str;
        }

        /// <summary>
        /// 设置SCADA标签值-监控报警
        /// </summary>
        /// <param name="offset"></param>
        /// <param name="input"></param>
        /// <returns>"DB304,X122.1;P1005_Y_DV_JogSelected","1","zh-CN","en-US"</returns>
        private string SetSacdaTag(string offset, string input, string msgType, string textUs, string textZh)
        {
            // BOOL：DB编号,X偏移量;标签
            string str = string.Format("\"{0}{1},X{2};{3}_{4}\",\"{5}\",\"{6}\",\"{7}\"", _programmingLanguage, _number, offset, _instanceName, input, msgType, textUs, textZh);
            _streamWriter.WriteLine(str);
            Console.WriteLine(str);
            return str;
        }
    }
}