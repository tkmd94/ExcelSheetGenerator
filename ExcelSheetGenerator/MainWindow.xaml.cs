using ClosedXML.Excel;
using EsapiEssentials.Plugin;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Windows;
using System.Xml.Linq;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;

namespace ExcelSheetGenerator
{
    /// <summary>
    /// MainWindow
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        // Define variables
        public string PreferencePath;
        public IEnumerable<XElement> PreferenceList;
        public PlanSetup plan;
        public string user;

        /// <summary>
        /// Initialize MainWindow and load configuration.
        /// </summary>
        /// <param name="context"></param>
        public MainWindow(PluginScriptContext context)
        {
            InitializeComponent();

            // change your cofiguration
            PreferencePath = @"D:\Dshare\ExcelSheetGenerator\ExcelSheetGeneratorPreference.xml";

            if (!File.Exists(PreferencePath))
            {
                MessageBox.Show(string.Format("File is not found. {0}", PreferencePath));
                return;
            }

            plan = context.PlanSetup;
            user = context.CurrentUser.Name;

            //load configuration
            XDocument xml = XDocument.Load(PreferencePath);
            XElement xmlPreference = xml.Element("preference");
            XElement xmlDefault = xmlPreference.Element("default");

            XElement xmlLists = xmlPreference.Element("lists");
            PreferenceList = xmlLists.Elements("list");
            IEnumerable<string> ListNames = PreferenceList.Select(p => p.Element("templatename").Value);
            foreach (var name in ListNames)
            {
                PreferenceComboBox.Items.Add(name);
            }
            PreferenceComboBox.SelectedIndex = 0;
        }

        /// <summary>
        /// Action when GenerateButton is clicked.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void GenerateButton_Click(object sender, RoutedEventArgs e)
        {

            string patientName = plan.Course.Patient.LastName + " " + plan.Course.Patient.FirstName;
            string patientId = plan.Course.Patient.Id;
            string courseId = plan.Course.Id; ;
            string planId = plan.Id;
            string exportFilePath = "NA";
            string dosePerFraction = plan.DosePerFraction.ToString();
            string totalDose = plan.TotalDose.ToString();
            string noOfFraction = plan.NumberOfFractions.ToString();
            string planApprovalDate = "NA";
            string creationDate = DateTime.Now.ToString("yyyyMMddHHmmss");
            string planApprovalStatus = plan.ApprovalStatus.ToString();
            DateTime PlanApprovalDateTime;

            if (DateTime.TryParse(plan.PlanningApprovalDate, out PlanApprovalDateTime))
            {
                planApprovalDate = PlanApprovalDateTime.ToString("yyyy/MM/dd HH:mm:ss");
            }
            string planApprover = plan.PlanningApproverDisplayName;
            string planApproverWithDateTime = plan.PlanningApproverDisplayName + " ( " + planApprovalDate + " ) ";

            string printDate = DateTime.Now.ToString();
            string printDateWithUser = "printed on " + printDate + " by " + user;


            int currentSelectedIndex = PreferenceComboBox.SelectedIndex;
            XElement xmlList = PreferenceList.ElementAt(currentSelectedIndex);

            string templatePath = xmlList.Element("templatePath").Value;           
            string exportRootPath = xmlList.Element("exportPath").Value;
            exportRootPath = exportRootPath.Replace("$PatientID", patientId);
            exportRootPath = exportRootPath.Replace("$PatientName", patientName);
            exportRootPath = exportRootPath.Replace("$CourseID", courseId);
            exportRootPath = exportRootPath.Replace("$PlanID", planId);
            exportRootPath = exportRootPath.Replace("$CreationDate", creationDate);

            //check template file
            if (!File.Exists(templatePath))
            {
                MessageBox.Show(string.Format("Template excel file is not found. {0}", templatePath));
                return;
            }

            //check export directory
            if (!Directory.Exists(exportRootPath))
            {
                Directory.CreateDirectory(exportRootPath);
            }

            string fileName = xmlList.Element("filename").Value;
            fileName = fileName.Replace("$PatientID", patientId);
            fileName = fileName.Replace("$PatientName", patientName);
            fileName = fileName.Replace("$CourseID", courseId);
            fileName = fileName.Replace("$PlanID", planId);
            fileName = fileName.Replace("$CreationDate", creationDate);
            string fileInfo = Path.GetExtension(templatePath);
            exportFilePath = exportRootPath + @"\" + fileName + fileInfo;

            //check export file name 
            if (File.Exists(exportFilePath))
            {
                MessageBox.Show("Check sheet already exists.");
                System.Diagnostics.Process.Start(exportRootPath);
                return;
            }
            IXLWorkbook wb = new XLWorkbook(templatePath);
            var ws = wb.Worksheet("Sheet1");

            XElement xmlCells = xmlList.Element("cells");
            IEnumerable<XElement> xmlCell = xmlCells.Elements("cell");

            foreach (XElement cell in xmlCell)
            {
                XElement elementType = cell.Element("type");
                XElement elementData = cell.Element("data");
                XElement elementAddress = cell.Element("address");
                if (elementType != null & elementData != null && elementAddress != null)
                {
                    if (elementType.Value == "info")
                    {
                        //Patient Id                          
                        if (elementData.Value == "PatientID")
                        {
                            ws.Cell(elementAddress.Value).SetValue(patientId);
                            ws.Cell(elementAddress.Value).SetDataType(XLDataType.Text);
                        }

                        //Patient Name
                        if (elementData.Value == "PatientName")
                        {
                            ws.Cell(elementAddress.Value).SetValue(patientName);
                            ws.Cell(elementAddress.Value).SetDataType(XLDataType.Text);
                        }

                        //Course ID
                        if (elementData.Value == "CourseID")
                        {
                            ws.Cell(elementAddress.Value).SetValue(courseId);
                            ws.Cell(elementAddress.Value).SetDataType(XLDataType.Text);
                        }

                        //Plan ID                           
                        if (elementData.Value == "PlanID")
                        {
                            ws.Cell(elementAddress.Value).SetValue(planId);
                            ws.Cell(elementAddress.Value).SetDataType(XLDataType.Text);
                        }

                        //DosePerFraction
                        if (elementData.Value == "DosePerFraction")
                        {
                            ws.Cell(elementAddress.Value).SetValue(dosePerFraction);
                            ws.Cell(elementAddress.Value).SetDataType(XLDataType.Text);
                        }
                        
                        //Total dose
                        if (elementData.Value == "TotalDose")
                        {
                            ws.Cell(elementAddress.Value).SetValue(totalDose);
                            ws.Cell(elementAddress.Value).SetDataType(XLDataType.Text);
                        }

                        //Number of fraction
                        if (elementData.Value == "NoOfFraction")
                        {
                            ws.Cell(elementAddress.Value).SetValue(noOfFraction);
                            ws.Cell(elementAddress.Value).SetDataType(XLDataType.Number);
                        }

                        //PrintDateWithUser
                        if (elementData.Value == "PrintDateWithUser")
                        {
                            ws.Cell(elementAddress.Value).SetValue(printDateWithUser);
                            ws.Cell(elementAddress.Value).SetDataType(XLDataType.Text);
                        }

                        //PlanApproverWithDateTime
                        if (elementData.Value == "PlanApproverWithDateTime")
                        {
                            ws.Cell(elementAddress.Value).SetValue(planApproverWithDateTime);
                            ws.Cell(elementAddress.Value).SetDataType(XLDataType.Text);
                        }

                        //PlanApprovalStatus
                        if (elementData.Value == "PlanApprovalStatus")
                        {                            
                            ws.Cell(elementAddress.Value).SetValue(planApprovalStatus);
                            ws.Cell(elementAddress.Value).SetDataType(XLDataType.Text);
                        }

                        // Isocenter.position
                        if (elementData.Value == "Isocenter.position.X" ||
                            elementData.Value == "Isocenter.position.Y" ||
                            elementData.Value == "Isocenter.position.Z")
                        {
                            var isoPosition = IsoToText(plan, elementData.Value);
                            if (isoPosition != null)
                            {
                                ws.Cell(elementAddress.Value).SetValue(isoPosition);
                                ws.Cell(elementAddress.Value).SetDataType(XLDataType.Number);
                            }
                            else
                            {
                                return;
                            }
                        }
                    }
                    else if (elementType.Value == "DVH")
                    {
                        XElement structure = cell.Element("structure");
                        if (structure != null)
                        {
                            //structure volume
                            if (elementData.Value == "Volume")
                            {
                                Structure targetStructure = null;
                                string structureVolume = "NA";
                                if (structure != null)
                                {
                                    try
                                    {
                                        targetStructure = plan.StructureSet.Structures.Where(s => s.Id == structure.Value).FirstOrDefault();
                                    }
                                    catch
                                    { }
                                    if (targetStructure != null && targetStructure.IsEmpty != true)
                                    {
                                        structureVolume = Math.Round(targetStructure.Volume).ToString();
                                    }
                                }                                
                                ws.Cell(elementAddress.Value).SetValue(structureVolume);
                                ws.Cell(elementAddress.Value).SetDataType(XLDataType.Number);
                            }
                        }
                    }
                }
            }
            // save as exportFilePath
            wb.SaveAs(exportFilePath);

            // open directory
            System.Diagnostics.Process.Start(exportRootPath);
        }

        /// <summary>
        /// Function to calculate isocenter coordinates and return them in text format.
        /// </summary>
        /// <param name="plan"></param>
        /// <param name="direction"></param>
        /// <returns></returns>
        private Nullable<double> IsoToText(PlanSetup plan, string direction)
        {
            //Single isocenter check
            var image = plan.StructureSet.Image;
            var isoPositionTable = new List<string>();
            VVector isoCNT = plan.Beams.ElementAt(0).IsocenterPosition;
            var iso_X = isoCNT.x * 0.1 - image.UserOrigin.x * 0.1;
            var iso_Y = isoCNT.y * 0.1 - image.UserOrigin.y * 0.1;
            var iso_Z = isoCNT.z * 0.1 - image.UserOrigin.z * 0.1;

            isoPositionTable.Add(string.Format("[ {0:f2} / {1:f2} / {2:f2} ]", (isoCNT.x * 0.1), (isoCNT.y * 0.1), (isoCNT.z * 0.1)));

            foreach (var beam in plan.Beams)
            {
                isoCNT = beam.IsocenterPosition;
                string isoPosition = string.Format("[ {0:f2} / {1:f2} / {2:f2} ]", (isoCNT.x * 0.1), (isoCNT.y * 0.1), (isoCNT.z * 0.1));
                bool isoDupFlg = false;
                foreach (var iso in isoPositionTable)
                {
                    if (isoPosition == iso)
                        isoDupFlg = true;
                }
                if (isoDupFlg == false)
                    isoPositionTable.Add(isoPosition);
            }

            if (isoPositionTable.Count() > 1)
            {
                MessageBox.Show("This plan is not single isocenter.");
                return null;
            }
            if (direction == "Isocenter.position.X")
            {
                return iso_X;
            }
            else if (direction == "Isocenter.position.Y")
            {
                return iso_Y;
            }
            else if (direction == "Isocenter.position.Z")
            {
                return iso_Z;
            }
            else
            {
                return null;
            }
        }
    }
}
