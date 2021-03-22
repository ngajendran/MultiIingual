using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Windows.Forms;
using System.Xml;
using System.Configuration;
using System.Collections.Specialized;
using System.Linq;

namespace MultilingualResourceFileGenerator
{
    public partial class MainForm : Form
    {
        internal static DataTable dtDefaultSheet;
        public MainForm()
        {
            InitializeComponent();
        }


        public DataSet ReadValuesFromExcel(string path)

        {

            Microsoft.Office.Interop.Excel.Application objXL = null;
            Microsoft.Office.Interop.Excel.Workbook objWB = null;

            DataSet dsWorkBook = new DataSet();
            try
            {

                objXL = new Microsoft.Office.Interop.Excel.Application();
                objWB = objXL.Workbooks.Open(path);
                foreach (Microsoft.Office.Interop.Excel.Worksheet objSHT in objWB.Worksheets)
                {
                    int irows = objSHT.UsedRange.Rows.Count;
                    int icols = objSHT.UsedRange.Columns.Count;
                    DataTable dtWorkSheet = new DataTable();
                    int noofrow = 1;
                    dtWorkSheet.TableName = objSHT.Name;

                    for (int iCurItem = 1; iCurItem <= icols; iCurItem++)
                    {
                        string strcolname = objSHT.Cells[1, iCurItem].Text;
                        dtWorkSheet.Columns.Add(strcolname);
                        noofrow = 2;
                    }

                    for (int iCurRow = noofrow; iCurRow <= irows; iCurRow++)
                    {
                        DataRow dr = dtWorkSheet.NewRow();
                        for (int iCurCol = 1; iCurCol <= icols; iCurCol++)
                        {
                            string currvalue = objSHT.Cells[iCurRow, iCurCol].Text;
                            dr[iCurCol - 1] = currvalue.Trim().Replace("#VALUE!", "");
                        }

                        dtWorkSheet.Rows.Add(dr);
                    }
                    dsWorkBook.Tables.Add(dtWorkSheet);
                }

                objWB.Close();
                objXL.Quit();
            }

            catch (Exception ex)
            {
                objWB.Saved = true;
                objWB.Close();
                objXL.Quit();
                MessageBox.Show("Unable to Read the Excel File due to exception : " + ex.Message.ToString());
            }
            return dsWorkBook;

        }

        private void openFileforGridView(string strExcelpath)
        {
            try
            {
                if (File.Exists(strExcelpath))
                {
                    DataSet dsExcel = ReadValuesFromExcel(strExcelpath);
                    if (dsExcel.Tables.Count > 0 && dsExcel.Tables != null)
                    {
                        dtDefaultSheet = dsExcel.Tables[0]; //Getting the First WorkSheet!
                        dataGridView1.DataSource = dtDefaultSheet;

                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Please load the excel sheet to process : " + ex.Message.ToString());
            }

        }



       

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void CreateResourceFile()
        {
            string strResourceFormatFile = Path.Combine(Environment.CurrentDirectory, "Resource_FormatFile.aspx.resx");

            try
            {
                if (File.Exists(strResourceFormatFile))
                {
                    List<string> lstColumnHeaders = new List<string>();
                    bool bContainskeyNameColumn = false;

                    foreach (DataColumn dc in dtDefaultSheet.Columns)
                    {
                        if (!string.Equals(dc.ColumnName.ToString(), "name", StringComparison.OrdinalIgnoreCase))
                            lstColumnHeaders.Add(dc.ColumnName);
                        else
                            bContainskeyNameColumn = true;
                    }

                   if(!bContainskeyNameColumn)
                    {
                        MessageBox.Show("The Loaded worksheet doesn't have the Name Column to generate the resource file ");
                        return;
                    }
                    
                    if (lstColumnHeaders.Count > 0)
                    {
                        foreach(string currColHeader in lstColumnHeaders)
                        {
                            bool bContainsCulture = ConfigurationManager.AppSettings.AllKeys
                                               .Where(key => key.Equals(currColHeader))
                                               .Any();

                            if(!bContainsCulture)
                            {
                                MessageBox.Show($"No Language match is found for  : {currColHeader} - in Appsettings to Read the culture Info", "File Creation Failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                return;
                            }

                        }


                        foreach (string currColHeader in lstColumnHeaders)
                        {
                            XmlDataDocument xmldoc = new XmlDataDocument();
                            FileStream fs = new FileStream(strResourceFormatFile, FileMode.Open, FileAccess.Read);
                            xmldoc.Load(fs);

                            foreach (DataRow dr in dtDefaultSheet.Rows)
                            {
                                string NodeName = dr["Name"].ToString();

                                if (string.IsNullOrEmpty(NodeName) || string.IsNullOrWhiteSpace(NodeName))
                                    continue;

                                if (xmldoc.SelectSingleNode("//root/data[@name='" + NodeName + "']/value") != null)
                                    continue;

                                XmlElement data = xmldoc.CreateElement("data");
                                data.SetAttribute("name", dr["Name"].ToString());
                                data.SetAttribute("xml:space", "preserve");
                                XmlElement value = xmldoc.CreateElement("value");
                                value.InnerText = dr[currColHeader].ToString().Trim().Replace("#VALUE!","");
                                data.AppendChild(value);
                                xmldoc.DocumentElement.AppendChild(data);
                            }
                            string strOutputFilename = dtDefaultSheet.TableName;

                            if (!strOutputFilename.EndsWith(".aspx.resx"))
                                strOutputFilename = string.Concat(dtDefaultSheet.TableName, ".aspx.resx");

                          
                            string strCultureName = ConfigurationManager.AppSettings.AllKeys
                                                  .Where(key => key.Equals(currColHeader))
                                                  .Select(key => ConfigurationManager.AppSettings[key])
                                                  .FirstOrDefault();

                            if (!string.IsNullOrEmpty(strCultureName) && !string.IsNullOrWhiteSpace(strCultureName))
                            {
                                strOutputFilename = string.Concat(strOutputFilename.Remove(strOutputFilename.LastIndexOf("resx")), strCultureName.Trim(), ".resx");
                            }
                           

                            string strAbsFileName = Path.Combine(Environment.CurrentDirectory, strOutputFilename);
                            xmldoc.Save(strAbsFileName);

                            if (!File.Exists(strAbsFileName))
                            {
                                MessageBox.Show($"Unable to generate resource file for {currColHeader}", "FileCreation", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                return;
                            }

                        }

                        MessageBox.Show($"The Excel Sheet is processsed and all resource files are generated", "File Creation Completed",MessageBoxButtons.OK, MessageBoxIcon.Information);

                    }

                }
                else
                {
                    MessageBox.Show("Resource format file is missing in  Application directory : ", "Format File missing", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception Occured, Unable to generate the file due to : " + ex.Message, "File Creation Failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            lblFilePathDisplay.Text = string.Empty;
        }

        private void btnCreateResource_Click(object sender, EventArgs e)
        {
            if (dtDefaultSheet != null && dtDefaultSheet.Rows.Count > 0)
            {
                CreateResourceFile();
            }
            else
            {
                MessageBox.Show("Unable to identify the dataset in the grid, Please select a valid excel sheet");
            }
        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            OpenFileDialog oFD = new OpenFileDialog();
            oFD.Filter = "Excel Files|*.xls;*.xlsx";
            oFD.CheckFileExists = true;

            if (oFD.ShowDialog() == DialogResult.OK)
            {
                if (string.IsNullOrEmpty(oFD.FileName) || string.IsNullOrWhiteSpace(oFD.FileName))
                {
                    MessageBox.Show("Please Select the Excel sheet", "File Selection", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    lblFilePathDisplay.Text = oFD.FileName;
                    lblFilePathDisplay.BorderStyle = BorderStyle.FixedSingle;
                    openFileforGridView(oFD.FileName);
                }
            }
        }
    }
}

