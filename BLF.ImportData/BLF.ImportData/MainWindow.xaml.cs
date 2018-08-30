using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;
using System.Xml.Linq;
using System.Xml;
using System.IO;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Collections;

namespace BLF.ImportData
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {

        XSSFWorkbook xssfworkbook;

        public MainWindow()
        {
            InitializeComponent();
            this.filePath.Text = "";
            this.Title = "Access BLF Data";
            
        }

        void InitializeWorkbook(string path)
        {
            using (FileStream file = new FileStream(path, FileMode.Open, FileAccess.Read))
            {
                xssfworkbook = new XSSFWorkbook(file);
            }
        }

        System.Data.DataTable InitializeForeignOrderDataTable()
        {

            System.Data.DataTable dtForeignOrder = new System.Data.DataTable();
            //dtAccount.TableName = Constants.AccountTable;
            dtForeignOrder.Columns.Add("CountryRegion");
            dtForeignOrder.Columns.Add("CalendarDay");
            dtForeignOrder.Columns.Add("AccountNoForCustomer");
            dtForeignOrder.Columns.Add("CustomerNameLocalLanguage");
            dtForeignOrder.Columns.Add("CountryRegionSoldTo");
            dtForeignOrder.Columns.Add("SalesOffice");
            dtForeignOrder.Columns.Add("NameOfSalesEmployee");
            dtForeignOrder.Columns.Add("IC0");
            dtForeignOrder.Columns.Add("IC1");
            dtForeignOrder.Columns.Add("IC5");
            dtForeignOrder.Columns.Add("PostalCode");
            dtForeignOrder.Columns.Add("CustomerNameInEnglish");
            dtForeignOrder.Columns.Add("CustomerNameInEnglish2");
            dtForeignOrder.Columns.Add("CityInEnglish");
            dtForeignOrder.Columns.Add("OrderNumber");
            dtForeignOrder.Columns.Add("InvoiceNo");
            dtForeignOrder.Columns.Add("DocumentType");
            dtForeignOrder.Columns.Add("MaterialDescription");
            dtForeignOrder.Columns.Add("OrderCode");
            dtForeignOrder.Columns.Add("ProductGroup");
            dtForeignOrder.Columns.Add("ProductArea");
            dtForeignOrder.Columns.Add("IOGross_SC ");
            dtForeignOrder.Columns.Add("IOQty_BaseUnit");
            dtForeignOrder.Columns.Add("SLGross_SC");
            dtForeignOrder.Columns.Add("SLQty_BaseUnit");

            return dtForeignOrder;
        }

        System.Data.DataTable InitializeAccountDataTable()
        {

            System.Data.DataTable dtAccount = new System.Data.DataTable();
            //dtAccount.TableName = Constants.AccountTable;
            dtAccount.Columns.Add("SAPNo");
            dtAccount.Columns.Add("CustomerName");
            dtAccount.Columns.Add("CustClass");
            dtAccount.Columns.Add("Paymentterms");
            dtAccount.Columns.Add("Created");
            dtAccount.Columns.Add("Region");
            dtAccount.Columns.Add("Office");
            dtAccount.Columns.Add("SalesEngineer");
            dtAccount.Columns.Add("SalesGroup");
            dtAccount.Columns.Add("Province");
            dtAccount.Columns.Add("City");
            dtAccount.Columns.Add("IndustryCode");
            dtAccount.Columns.Add("IndustryCodeDescription");
            dtAccount.Columns.Add("IndustryCode1");
            dtAccount.Columns.Add("IndustryCode1Description");
            dtAccount.Columns.Add("CustomerType");
            dtAccount.Columns.Add("IndustryCode3");
            dtAccount.Columns.Add("IndustryCode4");
            dtAccount.Columns.Add("IndustryCode5");
            dtAccount.Columns.Add("DeleteFlag");
            dtAccount.Columns.Add("PostalCode");
            dtAccount.Columns.Add("CustomerEnglishName");
            dtAccount.Columns.Add("CustomerEnglishName2");
            dtAccount.Columns.Add("CityEn");
            dtAccount.Columns.Add("CompaniesToVerify");
            dtAccount.Columns.Add("FirstOrderDate");
            dtAccount.Columns.Add("OperatingOrganization");
            dtAccount.Columns.Add("VerticalDevision");
            dtAccount.Columns.Add("Status");  
            return dtAccount;
        }

        System.Data.DataTable InitializeOrderDataTable()
        {

            System.Data.DataTable dtOrder = new System.Data.DataTable();
            //dtOrder.TableName = Constants.OrderTable;
            dtOrder.Columns.Add("IOCreatedYear");
            dtOrder.Columns.Add("CalendarWeek");
            dtOrder.Columns.Add("Year");
            dtOrder.Columns.Add("CalendarDay");
            dtOrder.Columns.Add("SAPNo");
            dtOrder.Columns.Add("CustomerName");
            dtOrder.Columns.Add("CustomerClassification");
            dtOrder.Columns.Add("PaymentTerms");
            dtOrder.Columns.Add("CreatedAt");
            dtOrder.Columns.Add("Region");
            dtOrder.Columns.Add("SaleOffice");
            dtOrder.Columns.Add("SaleEngineer");
            dtOrder.Columns.Add("City");
            dtOrder.Columns.Add("IndustryCode");
            dtOrder.Columns.Add("IndustryCodeDescription");
            dtOrder.Columns.Add("IndustryCode1");
            dtOrder.Columns.Add("IndustryCode1Description");
            dtOrder.Columns.Add("CustomerType");
            dtOrder.Columns.Add("DeleteFlag");
            dtOrder.Columns.Add("PostalCode");
            dtOrder.Columns.Add("CustomerEnglishName");
            dtOrder.Columns.Add("Name2");
            dtOrder.Columns.Add("OrderNo");
            dtOrder.Columns.Add("OrderItemNo");
            dtOrder.Columns.Add("OrderNoSoldtoParty");
            dtOrder.Columns.Add("ConditionType");
            dtOrder.Columns.Add("OrderType");
            dtOrder.Columns.Add("MaterialID");
            dtOrder.Columns.Add("Material");
            dtOrder.Columns.Add("OrderCode");
            dtOrder.Columns.Add("PlanGrp1MDID");
            dtOrder.Columns.Add("PlanGrp1MD");
            dtOrder.Columns.Add("PlanGrp2MDID");
            dtOrder.Columns.Add("PlanGrp2MD");
            dtOrder.Columns.Add("ProdhierLevel2");
            dtOrder.Columns.Add("ProductAreaMD");
            dtOrder.Columns.Add("ProductArea");
            dtOrder.Columns.Add("ProductGroupMD");
            dtOrder.Columns.Add("ProductGroup");
            dtOrder.Columns.Add("IOnetprcSD_SC");
            dtOrder.Columns.Add("IOqty");
            dtOrder.Columns.Add("SLnetprcSD_SC");
            dtOrder.Columns.Add("SLqty");
            dtOrder.Columns.Add("UnitPrice");
            dtOrder.Columns.Add("GlobalLine");
            dtOrder.Columns.Add("KeyIndustry");
            dtOrder.Columns.Add("OED");
            dtOrder.Columns.Add("MarkforSQL");
            dtOrder.Columns.Add("CustGroup");
            return dtOrder;
        }

        System.Data.DataTable InitializeMaterialPriceDataTable()
        {

            System.Data.DataTable dtMaterialPrice = new System.Data.DataTable();
            //dtMaterialPrice.TableName = Constants.MtlsPriceTable;
            dtMaterialPrice.Columns.Add("Material");
            dtMaterialPrice.Columns.Add("CLPinCNY");
            dtMaterialPrice.Columns.Add("GLPinEuro");
            dtMaterialPrice.Columns.Add("ProductArea");
            dtMaterialPrice.Columns.Add("ProductGroup");
            dtMaterialPrice.Columns.Add("NewProduct");
            dtMaterialPrice.Columns.Add("OrderCode");
           
            return dtMaterialPrice;
        }

        System.Data.DataTable InitializeOpenOrderDataTable()
        {
            System.Data.DataTable dtOpenOrder = new System.Data.DataTable();
            dtOpenOrder.Columns.Add("CalendarDay");
            dtOpenOrder.Columns.Add("OrderNumber");
            dtOpenOrder.Columns.Add("OrderItemNumber");
            dtOpenOrder.Columns.Add("OpenValue");
            dtOpenOrder.Columns.Add("Period");
            return dtOpenOrder;
        }

        System.Data.DataTable InitializeExchangeRatesDataTable()
        {
            System.Data.DataTable dtExchangeRates = new System.Data.DataTable();
            dtExchangeRates.Columns.Add("CountryRegion");
            dtExchangeRates.Columns.Add("rate");
            return dtExchangeRates;
        }

        private void tables_Loaded(object sender, RoutedEventArgs e)
        {
            // List<string> tables = new List<string>();
            SqlCommand cmd = new SqlCommand();
            SqlConnection sqlCon = new SqlConnection();
            try
            {
               
                sqlCon.ConnectionString = ConfigurationSettings.AppSettings["conStr"];
                cmd.CommandText = string.Format(Constants.GetTables, Constants.DBName);
                cmd.Connection = sqlCon;
                SqlDataAdapter sda = new SqlDataAdapter(cmd);
                System.Data.DataTable dt = new System.Data.DataTable("tables");
                sda.Fill(dt);
                var comboBox = sender as ComboBox;
                comboBox.ItemsSource = dt.DefaultView;
                comboBox.DisplayMemberPath = dt.Columns[0].ToString();
                comboBox.SelectedValuePath = dt.Columns[0].ToString();
            }
            catch (Exception ex)
            {
                this.Error.Text = ex.Message;
            }
            finally
            {
                cmd.Dispose();
                sqlCon.Close();
            }
        }

        private void ExecuteInsertSQL(SqlCommand cmd)
        {
           
            SqlConnection sqlCon = new SqlConnection();
            try
            {

                sqlCon.ConnectionString = ConfigurationSettings.AppSettings["conStr"];
                cmd.Connection = sqlCon;
                sqlCon.Open();
                cmd.ExecuteNonQuery();
                
               
            }
            catch (Exception ex)
            {
                if (cmd.Parameters.Contains("@CustomerName"))
                {

                    this.Error.Text = this.Error.Text + "\n" + "Current Item: " + cmd.Parameters["@CustomerName"].Value + " Error: " + ex.Message;
                }
                else if (cmd.Parameters.Contains("@SaleEngineer"))
                {
                    this.Error.Text = this.Error.Text + "\n" + "Current Item: " + cmd.Parameters["@SaleEngineer"].Value + " Error: " + ex.Message;
                }
                else
                {
                    this.Error.Text = this.Error.Text + "\n" + " Error: " + ex.Message;
                }
            }
            finally
            {
                cmd.Dispose();
                sqlCon.Close();
            }
            
        }
        private void ButtonImport_Click(object sender, RoutedEventArgs e)
        {

            string filePath = this.filePath.Text.Trim();
            if (string.IsNullOrEmpty(filePath))
            {
                this.Error.Text = "Please input XML file path...";
            }
            else
            {
                string tableName = this.tables.SelectedValue.ToString();
                this.Error.Text = "";
                if (tableName == Constants.AccountTable || tableName == Constants.OrderTable || tableName == Constants.MtlsPriceTable || tableName == Constants.OpenOrderTable || tableName == Constants.ForeignOrderTable || tableName == Constants.ExchangeRatesTable)
                {
                    //excel upload edit by justin li 
                    //create xml file by excel file
                    ArrayList xmlPathList= CreateXml(tableName, filePath);
                    //upload
                    foreach (string xmlPath in xmlPathList)
                    {
                        InsertDataToTable(tableName, xmlPath);
                    }
                    MessageBox.Show("upload completed..", "Upload Data", MessageBoxButton.OK);
                   
                }
                else if ( tableName == Constants.BudgetTable  || tableName == Constants.SalesTable || tableName == Constants.AccountMappingTable || tableName == Constants.MtlsTable)
                {
                    //xml upload
                    InsertDataToTable(tableName, filePath);
                }
                else
                {
                    this.Error.Text = "Can not upload " + tableName + "...";
                }
            }
        }

        
        #region unused functions
        
        private void ButtonUpdate_Click(object sender, RoutedEventArgs e)
        {

            string filePath = this.filePath.Text.Trim();
            if (string.IsNullOrEmpty(filePath))
            {
                this.Error.Text = "Please input XML file path...";
            }
            else
            {
                string tableName = this.tables.SelectedValue.ToString();
                this.Error.Text = "";
                if (tableName == Constants.OrderTable)
                {
                    UpdateDataToTable(tableName, filePath);
                }
                else
                {
                    this.Error.Text = "Can not upload " + tableName + "...";
                }
            }
        }

        private void ButtonExport_Click(object sender, RoutedEventArgs e)
        {
            string tableName = this.tables.SelectedValue.ToString();
            if (string.IsNullOrEmpty(tableName))
            {
                this.Error.Text = "Please select one table...";
            }
            else
            {
                
            }
        }

        #endregion
        //new by justin li 
        private ArrayList CreateXml(string tableName, string path)
        {
            string xmlPath = "";
            ArrayList xmlPathlist = new ArrayList();
            System.Data.DataTable dt = null;
            string sheetName = "";
            switch (tableName)
            {
                case Constants.AccountTable:
                    dt = InitializeAccountDataTable();
                    sheetName = "Customer Master";
                    break;
                case Constants.OrderTable:
                    dt = InitializeOrderDataTable();
                    sheetName = "Weekly Data";
                    break;
                case Constants.MtlsPriceTable:
                    dt = InitializeMaterialPriceDataTable();
                    sheetName = "Material Price";
                    break;
                case Constants.OpenOrderTable:
                    dt = InitializeOpenOrderDataTable();
                    sheetName = "Open Order";
                    break;
                case Constants.ForeignOrderTable:
                    dt = InitializeForeignOrderDataTable();
                    sheetName = "Data APAC";
                    break;
                case Constants.ExchangeRatesTable:
                    dt = InitializeExchangeRatesDataTable();
                    sheetName = "Exchange Rates";
                    break;
                default:
                    dt = null;
                    break;
            }
            if(dt != null)
            {
                //'Item' is tagName in xml
                dt.TableName = "Item";
            }
            
            InitializeWorkbook(path);
            ISheet sheet = xssfworkbook.GetSheet(sheetName);
            System.Collections.IEnumerator rows = sheet.GetRowEnumerator();

            if(dt == null)
            {
                return null;
            }
            //start at line 2
            rows.MoveNext();
            int count = 0;
            int index = 0;              
            while (rows.MoveNext())
            {
                if (count >= Constants.RowLimit)
                {
                    count = 0;
                    xmlPath = CreateXml(tableName, index, dt);
                    //xmlPath = @"E:\Upload BLF_Data\" + tableName + DateTime.Now.ToString("yyyyMMddhhmmss") +"_"+ index + ".xml";
                    //create xml file to local
                    //dt.WriteXml(xmlPath);
                    xmlPathlist.Add(xmlPath);
                    dt.Clear();
                    index++;
                }

                 
                IRow row = (XSSFRow)rows.Current;
                DataRow dr = dt.NewRow();

                for (int i = 0; i < row.LastCellNum; i++)
                {
                    ICell cell = row.GetCell(i);
                    if (cell == null)
                    {
                        dr[i] = null;
                    }
                    else
                    {
                        //var evaluator = xssfworkbook.GetCreationHelper().CreateFormulaEvaluator();
                        switch (cell.CellType) {
                            case CellType.Numeric:
                                dr[i] = cell.NumericCellValue.ToString();
                                break;
                            case CellType.Boolean:
                                dr[i] = cell.BooleanCellValue.ToString();
                                break;
                            case CellType.String:
                                dr[i] = cell.StringCellValue.ToString();
                                break;
                            case CellType.Error:
                                dr[i] = cell.ErrorCellValue.ToString();
                                break;
                            case CellType.Blank:
                                dr[i] = cell.ToString();
                                break;
                            case CellType.Formula:
                                // dr[i] = evaluator.EvaluateFormulaCell(cell);
                                throw new Exception(string.Format("formula cell at cell {0}", i+1));
                                
                            default:
                                    dr[i] = cell.ToString();
                                break;
                        }
                        
                    }
                }
                if(dr[0].ToString() != "")
                {
                    dt.Rows.Add(dr);
                }
                count++;
               
                
            }
            xmlPath = CreateXml(tableName, index, dt);
            //xmlPath = @"E:\Upload BLF_Data\" + tableName + DateTime.Now.ToString("yyyyMMddhhmmss") +"_"+ index + ".xml";
            //create xml file to local
            //dt.WriteXml(xmlPath);
            xmlPathlist.Add(xmlPath);
            dt.Dispose();
            
            
            return xmlPathlist;

        }

        private string CreateXml(string tableName,int index, System.Data.DataTable dt)
        {
            string xmlPath = @"E:\Upload BLF_Data\" + tableName + DateTime.Now.ToString("yyyyMMddhhmmss") + "_" + index + ".xml";
            //create xml file to local
            dt.WriteXml(xmlPath);
            return xmlPath;
        }

        private void InsertDataToTable(string tableName,string filePath)
        {
            XmlDataDocument xmlDoc = new XmlDataDocument();
            SqlCommand cmd = new SqlCommand();

            try
            {
                using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    if (fs != null)
                    {
                        xmlDoc.Load(fs);
                        XmlNodeList NodeList = xmlDoc.GetElementsByTagName("Item");
                        for (int i = 0; i < NodeList.Count; i++)
                        {
                            //string ID = NodeList[i];
                            cmd = SetParameters(tableName, NodeList[i]);
                            ExecuteInsertSQL(cmd);
                            if (!string.IsNullOrEmpty(this.Error.Text.ToString()))
                            {
                                continue;
                            }
                        }
                        
                    }
                }

            }
            catch (Exception ex)
            {
                this.Error.Text = ex.Message;
            }
            finally
            {
                cmd.Dispose();
            }
        }


        
        //private void InsertDataToTable(string tableName, string filePath, string tagName)
        //{
        //    XmlDataDocument xmlDoc = new XmlDataDocument();
        //    SqlCommand cmd = new SqlCommand();

        //    try
        //    {
        //        using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
        //        {
        //            if (fs != null)
        //            {
        //                xmlDoc.Load(fs);
        //                XmlNodeList NodeList = xmlDoc.GetElementsByTagName(tagName);
        //                for (int i = 0; i < NodeList.Count; i++)
        //                {
        //                    //string ID = NodeList[i];
        //                    cmd = SetParameters(tableName, NodeList[i]);
        //                    ExecuteInsertSQL(cmd);
        //                    if (!string.IsNullOrEmpty(this.Error.Text.ToString()))
        //                    {
        //                        continue;
        //                    }
        //                }
        //                if (string.IsNullOrEmpty(this.Error.Text.ToString()))
        //                {
        //                    MessageBox.Show("upload successfully..", "Upload Data", MessageBoxButton.OK);
        //                }
        //            }
        //        }

        //    }
        //    catch (Exception ex)
        //    {
        //        this.Error.Text = ex.Message;
        //    }
        //    finally
        //    {
        //        cmd.Dispose();
        //    }
        //}

        private void UpdateDataToTable(string tableName, string filePath)
        {
            XmlDataDocument xmlDoc = new XmlDataDocument();
            SqlCommand cmd = new SqlCommand();

            try
            {
                using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    if (fs != null)
                    {
                        xmlDoc.Load(fs);
                        XmlNodeList NodeList = xmlDoc.GetElementsByTagName("Item");
                        for (int i = 0; i < NodeList.Count; i++)
                        {
                            //string ID = NodeList[i];
                            cmd = SetParametersForUpdate(tableName, NodeList[i]);
                            ExecuteInsertSQL(cmd);
                            if (!string.IsNullOrEmpty(this.Error.Text.ToString()))
                            {
                                continue;
                            }
                        }
                        if (string.IsNullOrEmpty(this.Error.Text.ToString()))
                        {
                            MessageBox.Show("update successfully..", "Update Data", MessageBoxButton.OK);
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                this.Error.Text = ex.Message;
            }
            finally
            {
                cmd.Dispose();
            }
        }
        
        private SqlCommand SetParameters(string tableName, XmlNode node)
        {
            SqlCommand cmd = new SqlCommand();
            switch (tableName)
            {
                case Constants.AccountTable:
                        cmd.CommandText = Constants.strInsertAccount;
                        cmd.Parameters.AddWithValue("@SAPNo", node["SAPNo"] != null ? node["SAPNo"].InnerText : "");
                        cmd.Parameters.AddWithValue("@CustomerName", node["CustomerName"] != null ? node["CustomerName"].InnerText : "");
                        cmd.Parameters.AddWithValue("@CustClass", node["CustClass"] != null ? node["CustClass"].InnerText : "");
                        cmd.Parameters.AddWithValue("@Paymentterms", node["Paymentterms"] != null ? node["Paymentterms"].InnerText : "");
                        cmd.Parameters.AddWithValue("@Created", node["Created"] != null ? Convert.ToInt32(node["Created"].InnerText) : 0);
                        cmd.Parameters.AddWithValue("@Region", node["Region"] != null ? node["Region"].InnerText : "");
                        cmd.Parameters.AddWithValue("@Office", node["Office"] != null ? node["Office"].InnerText : "");
                        cmd.Parameters.AddWithValue("@SalesEngineer", node["SalesEngineer"] != null ? node["SalesEngineer"].InnerText : "");
                        cmd.Parameters.AddWithValue("@SalesGroup", node["SalesGroup"] != null ? node["SalesGroup"].InnerText : "");
                        cmd.Parameters.AddWithValue("@Province", node["Province"] != null ? node["Province"].InnerText : "");
                        cmd.Parameters.AddWithValue("@City", node["City"] != null ? node["City"].InnerText : "");
                        cmd.Parameters.AddWithValue("@IndustryCode", node["IndustryCode"] != null ? node["IndustryCode"].InnerText : "");
                        cmd.Parameters.AddWithValue("@IndustryCodeDescription", node["IndustryCodeDescription"] != null ? node["IndustryCodeDescription"].InnerText : "");
                        cmd.Parameters.AddWithValue("@IndustryCode1", node["IndustryCode1"] != null ? node["IndustryCode1"].InnerText : "");
                        cmd.Parameters.AddWithValue("@IndustryCode1Description", node["IndustryCode1Description"] != null ? node["IndustryCode1Description"].InnerText : "");
                        cmd.Parameters.AddWithValue("@CustomerType", node["CustomerType"] != null ? node["CustomerType"].InnerText : "");
                        cmd.Parameters.AddWithValue("@IndustryCode3", node["IndustryCode3"] != null ? node["IndustryCode3"].InnerText : "");
                        cmd.Parameters.AddWithValue("@IndustryCode4", node["IndustryCode4"] != null ? node["IndustryCode4"].InnerText : "");
                        cmd.Parameters.AddWithValue("@IndustryCode5", node["IndustryCode5"] != null ? node["IndustryCode5"].InnerText : "");
                        cmd.Parameters.AddWithValue("@DeleteFlag", node["DeleteFlag"] != null ? node["DeleteFlag"].InnerText : "");
                        cmd.Parameters.AddWithValue("@PostalCode", node["PostalCode"] != null ? node["PostalCode"].InnerText : "");
                        cmd.Parameters.AddWithValue("@CustomerEnglishName", node["CustomerEnglishName"] != null ? node["CustomerEnglishName"].InnerText : "");
                        cmd.Parameters.AddWithValue("@CustomerEnglishName2", node["CustomerEnglishName2"] != null ? node["CustomerEnglishName2"].InnerText : "");
                        cmd.Parameters.AddWithValue("@CityEn", node["CityEn"] != null ? node["CityEn"].InnerText : "");
                        cmd.Parameters.AddWithValue("@CompaniesToVerify", node["CompaniesToVerify"] != null ? node["CompaniesToVerify"].InnerText : "");
                        cmd.Parameters.AddWithValue("@OperatingOrganization", node["OperatingOrganization"] != null && node["OperatingOrganization"].InnerText != "0" ? node["OperatingOrganization"].InnerText : "");
                        cmd.Parameters.AddWithValue("@VerticalDevision", node["VerticalDevision"] != null && node["VerticalDevision"].InnerText != "" ? node["VerticalDevision"].InnerText : "");
                        cmd.Parameters.AddWithValue("@FirstOrderDate", node["FirstOrderDate"] != null && node["FirstOrderDate"].InnerText != "" ? Convert.ToInt32(node["FirstOrderDate"].InnerText) : 0);
                        cmd.Parameters.AddWithValue("@Status", node["Status"] != null ? node["Status"].InnerText : "");
                        if (node["IndustryCode"].InnerText.Trim().StartsWith("M") || node["IndustryCode"].InnerText.Trim().StartsWith("D"))
                        {
                            cmd.Parameters.AddWithValue("@OED", "Distribution");
                        }
                        else if (node["IndustryCode"].InnerText.Trim().StartsWith("E"))
                        { 
                            cmd.Parameters.AddWithValue("@OED", "Enduser");
                        }
                        else if (node["IndustryCode"].InnerText.Trim().StartsWith("B"))
                        {
                            cmd.Parameters.AddWithValue("@OED", "BAL");
                        }
                        else if (node["IndustryCode"].InnerText.Trim().StartsWith("O") || node["IndustryCode"].InnerText.Trim().StartsWith("S"))
                        {
                            cmd.Parameters.AddWithValue("@OED", "OEM");
                        }
                        else
                        {
                            cmd.Parameters.AddWithValue("@OED", "");
                        }
                    break;
                case Constants.BudgetTable:
                        cmd.CommandText = Constants.strInsertBudget;
                        cmd.Parameters.AddWithValue("@CustomerName", node["CustomerName"] != null ? node["CustomerName"].InnerText : "");
                        cmd.Parameters.AddWithValue("@CustomerEnglishName", node["CustomerEnglishName"] != null ? node["CustomerEnglishName"].InnerText : "");
                        cmd.Parameters.AddWithValue("@Budget", node["Budget"] != null ? node["Budget"].InnerText : "");
                        cmd.Parameters.AddWithValue("@SAPNo", node["SAPNo"] != null ? node["SAPNo"].InnerText : "");
                        cmd.Parameters.AddWithValue("@Year", node["Year"] != null ? node["Year"].InnerText : "");
                    break;
                case Constants.AccountMappingTable:
                        cmd.CommandText = Constants.strInsertAccountMapping;
                        cmd.Parameters.AddWithValue("@SAPNO", node["SAPNO"] != null ? node["SAPNO"].InnerText : "");
                        cmd.Parameters.AddWithValue("@CustomerName", node["CustomerName"] != null ? node["CustomerName"].InnerText : "");
                        cmd.Parameters.AddWithValue("@CustomerEnglishName", node["CustomerEnglishName"] != null ? node["CustomerEnglishName"].InnerText : "");
                        cmd.Parameters.AddWithValue("@IndustryCode", node["IndustryCode"] != null ? node["IndustryCode"].InnerText : "");
                        cmd.Parameters.AddWithValue("@IndustryCode1", node["IndustryCode1"] != null ? node["IndustryCode1"].InnerText : "");
                        cmd.Parameters.AddWithValue("@CustomerType", node["CustomerType"] != null ? node["CustomerType"].InnerText : "");
                        cmd.Parameters.AddWithValue("@SalesEngineer", node["SalesEngineer"] != null ? node["SalesEngineer"].InnerText : "");
                        cmd.Parameters.AddWithValue("@AreaSalesOffice", node["AreaSalesOffice"] != null ? node["AreaSalesOffice"].InnerText : "");
                        cmd.Parameters.AddWithValue("@CustomerClassification", node["CustomerClassification"] != null ? node["CustomerClassification"].InnerText : "");
                    break;
                case Constants.OrderTable:
                        cmd.CommandText = Constants.strInsertOrder;
                        cmd.Parameters.AddWithValue("@CustomerName", node["CustomerName"] != null ? node["CustomerName"].InnerText : "");
                        cmd.Parameters.AddWithValue("@CustomerEnglishName", node["CustomerEnglishName"] != null ? node["CustomerEnglishName"].InnerText : "");//
                        cmd.Parameters.AddWithValue("@SaleEngineer", node["SaleEngineer"] != null ? node["SaleEngineer"].InnerText : "");
                        cmd.Parameters.AddWithValue("@SaleOffice", node["SaleOffice"] != null ? node["SaleOffice"].InnerText : "");//
                        cmd.Parameters.AddWithValue("@IOCreatedYear", node["IOCreatedYear"] != null && node["IOCreatedYear"].InnerText != "" ? System.Convert.ToInt32(node["IOCreatedYear"].InnerText) : 0);
                        cmd.Parameters.AddWithValue("@CalendarWeek", node["CalendarWeek"] != null && node["CalendarWeek"].InnerText != "" ? System.Convert.ToInt32(node["CalendarWeek"].InnerText) : 0);
                        cmd.Parameters.AddWithValue("@Year", node["Year"] != null && node["Year"].InnerText != "" ? System.Convert.ToInt32(node["Year"].InnerText) : 0);
                        cmd.Parameters.AddWithValue("@CalendarDay", node["CalendarDay"] != null && node["CalendarDay"].InnerText != "" ? System.Convert.ToInt32(node["CalendarDay"].InnerText) : 0);
                        cmd.Parameters.AddWithValue("@MaterialID", node["MaterialID"] != null && node["MaterialID"].InnerText != "" ? System.Convert.ToInt32(node["MaterialID"].InnerText) : 0);
                        cmd.Parameters.AddWithValue("@Material", node["Material"] != null ? node["Material"].InnerText : "");
                        cmd.Parameters.AddWithValue("@OrderCode", node["OrderCode"] != null ? node["OrderCode"].InnerText : "");
                        cmd.Parameters.AddWithValue("@PlanGrp1MDID", node["PlanGrp1MDID"] != null ? node["PlanGrp1MDID"].InnerText : "");
                        cmd.Parameters.AddWithValue("@PlanGrp1MD", node["PlanGrp1MD"] != null ? node["PlanGrp1MD"].InnerText : "");
                        cmd.Parameters.AddWithValue("@PlanGrp2MDID", node["PlanGrp2MDID"] != null ? node["PlanGrp2MDID"].InnerText : "");
                        cmd.Parameters.AddWithValue("@PlanGrp2MD", node["PlanGrp2MD"] != null ? node["PlanGrp2MD"].InnerText : "");
                        cmd.Parameters.AddWithValue("@OrderItemNo", node["OrderItemNo"] != null ? node["OrderItemNo"].InnerText : "");
                        cmd.Parameters.AddWithValue("@IOnetprcSD_SC", node["IOnetprcSD_SC"] != null && node["IOnetprcSD_SC"].InnerText != "" ? System.Convert.ToDecimal(node["IOnetprcSD_SC"].InnerText) : 0);
                        cmd.Parameters.AddWithValue("@IOqty", node["IOqty"] != null && node["IOqty"].InnerText != "" ? System.Convert.ToInt32(node["IOqty"].InnerText) : 0);
                        cmd.Parameters.AddWithValue("@OrderNo", node["OrderNo"] != null ? node["OrderNo"].InnerText : "");
                        cmd.Parameters.AddWithValue("@SAPNo", node["SAPNo"] != null && node["SAPNo"].InnerText != "" ? System.Convert.ToInt32(node["SAPNo"].InnerText) : 0);
                        cmd.Parameters.AddWithValue("@Region", node["Region"] != null ? node["Region"].InnerText : "");
                        cmd.Parameters.AddWithValue("@CustomerClassification", node["CustomerClassification"] != null ? node["CustomerClassification"].InnerText : "");
                        cmd.Parameters.AddWithValue("@OED", node["OED"] != null ? node["OED"].InnerText : "");
                        cmd.Parameters.AddWithValue("@IndustryCode", node["IndustryCode"] != null ? node["IndustryCode"].InnerText : "");
                        cmd.Parameters.AddWithValue("@IndustryCode1", node["IndustryCode1"] != null ? node["IndustryCode1"].InnerText : "");
                        cmd.Parameters.AddWithValue("@SLnetprcSD_SC", node["SLnetprcSD_SC"] != null && node["SLnetprcSD_SC"].InnerText != "" ? System.Convert.ToDecimal(node["SLnetprcSD_SC"].InnerText) : 0);
                        cmd.Parameters.AddWithValue("@SLqty", node["SLqty"] != null ? node["SLqty"].InnerText : "");
                        cmd.Parameters.AddWithValue("@PaymentTerms", node["PaymentTerms"] != null ? node["PaymentTerms"].InnerText : "");
                        cmd.Parameters.AddWithValue("@CreatedAt", node["CreatedAt"] != null && node["CreatedAt"].InnerText != "" ? System.Convert.ToInt32(node["CreatedAt"].InnerText) : 0);
                        cmd.Parameters.AddWithValue("@City", node["City"] != null ? node["City"].InnerText : "");
                        cmd.Parameters.AddWithValue("@CustomerType", node["CustomerType"] != null ? node["CustomerType"].InnerText : "");
                        // Figure Booking Fields
                        cmd.Parameters.AddWithValue("@FigureBooking", node["FigureBooking"] != null ? (object)node["FigureBooking"].InnerText : DBNull.Value);
                        cmd.Parameters.AddWithValue("@CRMProjectID", node["CRMProjectID"] != null ? (object)System.Convert.ToInt32(node["CRMProjectID"].InnerText) : DBNull.Value);
                        cmd.Parameters.AddWithValue("@ProjectName", node["ProjectName"] != null ? (object)node["ProjectName"].InnerText : DBNull.Value);
                        cmd.Parameters.AddWithValue("@ApplicantSalesOffice", node["ApplicantSalesOffice"] != null ? (object)node["ApplicantSalesOffice"].InnerText : DBNull.Value);
                        cmd.Parameters.AddWithValue("@ApplicantSalesEngineer", node["ApplicantSalesEngineer"] != null ? (object)node["ApplicantSalesEngineer"].InnerText : DBNull.Value);
                        cmd.Parameters.AddWithValue("@EndUserName", node["EndUserName"] != null ? (object)node["EndUserName"].InnerText : DBNull.Value);
                        cmd.Parameters.AddWithValue("@SalesEngineerByEndUer", node["SalesEngineerByEndUer"] != null ? (object)node["SalesEngineerByEndUer"].InnerText : DBNull.Value);
                        cmd.Parameters.AddWithValue("@SalesOfficeByEndUser", node["SalesOfficeByEndUser"] != null ? (object)node["SalesOfficeByEndUser"].InnerText : DBNull.Value);
                        cmd.Parameters.AddWithValue("@SalesRegionByEndUser", node["SalesRegionByEndUser"] != null ? (object)node["SalesRegionByEndUser"].InnerText : DBNull.Value);
                        cmd.Parameters.AddWithValue("@ProductArea", node["ProductArea"] != null ? (object)node["ProductArea"].InnerText : DBNull.Value);
                        cmd.Parameters.AddWithValue("@ProductGroup", node["ProductGroup"] != null ? (object)node["ProductGroup"].InnerText : DBNull.Value);
                        cmd.Parameters.AddWithValue("@OrderType", node["OrderType"] != null ? (object)node["OrderType"].InnerText : DBNull.Value);
                        cmd.Parameters.AddWithValue("@ConditionType", node["ConditionType"] != null ? (object)node["ConditionType"].InnerText : DBNull.Value);

                    break;
                case Constants.SalesTable:
                        cmd.CommandText = Constants.strInsertSales;
                        cmd.Parameters.AddWithValue("@SaleEngineer", node["SaleEngineer"] != null ? node["SaleEngineer"].InnerText : "");
                        cmd.Parameters.AddWithValue("@SaleOffice", node["SaleOffice"] != null ? node["SaleOffice"].InnerText : "");
                        cmd.Parameters.AddWithValue("@Status", node["Status"] != null ? node["Status"].InnerText : "");
                        cmd.Parameters.AddWithValue("@OnBoardTime", node["OnBoardTime"] != null ? node["OnBoardTime"].InnerText : "");
                        cmd.Parameters.AddWithValue("@ExitTime", node["ExitTime"] != null ? node["ExitTime"].InnerText : "");
                    break;
                case Constants.MtlsTable:
                    cmd.CommandText = Constants.strInsertMtl;
                    cmd.Parameters.AddWithValue("@Mtl", node["Mtl"] != null ? node["Mtl"].InnerText.Trim() : "");
                    cmd.Parameters.AddWithValue("@C28", node["Cam"].InnerText.ToString().Trim() == "C28" ? 1 : 0);
                    cmd.Parameters.AddWithValue("@C26", node["Cam"].InnerText.ToString().Trim() == "C26" ? 1 : 0);
                    cmd.Parameters.AddWithValue("@C15", node["Cam"].InnerText.ToString().Trim() == "C15" ? 1 : 0);
                    cmd.Parameters.AddWithValue("@C4", node["Cam"].InnerText.ToString().Trim() == "C15" ? 1 : 0);
                    break;
                case Constants.MtlsPriceTable:
                    cmd.CommandText = Constants.strInsertMtlPrice;
                    cmd.Parameters.AddWithValue("@Material", node["Material"] != null ? node["Material"].InnerText.Trim() : "");
                    cmd.Parameters.AddWithValue("@CLPinCNY", node["CLPinCNY"] != null ? System.Convert.ToDecimal(node["CLPinCNY"].InnerText) : 0);
                    cmd.Parameters.AddWithValue("@GLPinEuro", node["GLPinEuro"] != null ? System.Convert.ToDecimal(node["GLPinEuro"].InnerText) : 0);
                    cmd.Parameters.AddWithValue("@ProductArea", node["ProductArea"] != null ? node["ProductArea"].InnerText.Trim() : "");
                    cmd.Parameters.AddWithValue("@ProductGroup", node["ProductGroup"] != null ? node["ProductGroup"].InnerText.Trim() : "");
                    cmd.Parameters.AddWithValue("@NewProduct", node["NewProduct"] != null ? node["NewProduct"].InnerText.Trim() : "");
                    cmd.Parameters.AddWithValue("@OrderCode", node["OrderCode"] != null ? node["OrderCode"].InnerText.Trim() : "");
                    break;
                case Constants.OpenOrderTable:
                    cmd.CommandText = Constants.strInsertOpenOrder;
                    cmd.Parameters.AddWithValue("@OrderNumber", node["OrderNumber"] != null ? node["OrderNumber"].InnerText.Trim() : "");
                    cmd.Parameters.AddWithValue("@OrderItemNumber", node["OrderItemNumber"] != null ? node["OrderItemNumber"].InnerText.Trim() : "");
                    cmd.Parameters.AddWithValue("@OpenValue", node["OpenValue"] != null ? System.Convert.ToDecimal(node["OpenValue"].InnerText) : 0);
                    cmd.Parameters.AddWithValue("@Period", node["Period"] != null ? node["Period"].InnerText.Trim() : "");
                    cmd.Parameters.AddWithValue("@CalendarDay", node["CalendarDay"] != null && node["CalendarDay"].InnerText != "" ? System.Convert.ToInt32(node["CalendarDay"].InnerText) : 0);
                    break;
                case Constants.ForeignOrderTable:
                    cmd.CommandText = Constants.strInsertForeignOrder;
                    cmd.Parameters.AddWithValue("@CountryRegion", node["CountryRegion"] != null ? node["CountryRegion"].InnerText : "");
                    cmd.Parameters.AddWithValue("@CalendarDay", node["CalendarDay"] != null && node["CalendarDay"].InnerText != "" ? System.Convert.ToInt32(node["CalendarDay"].InnerText) : 0);
                    cmd.Parameters.AddWithValue("@AccountNoForCustomer", node["AccountNoForCustomer"] != null ? node["AccountNoForCustomer"].InnerText : "");//
                    cmd.Parameters.AddWithValue("@CustomerNameLocalLanguage", node["CustomerNameLocalLanguage"] != null ? node["CustomerNameLocalLanguage"].InnerText : "");
                    cmd.Parameters.AddWithValue("@CountryRegionSoldTo", node["CountryRegionSoldTo"] != null ? node["CountryRegionSoldTo"].InnerText : "");
                    cmd.Parameters.AddWithValue("@SalesOffice", node["SalesOffice"] != null ? node["SalesOffice"].InnerText : "");//
                    cmd.Parameters.AddWithValue("@NameOfSalesEmployee", node["NameOfSalesEmployee"] != null ? node["NameOfSalesEmployee"].InnerText : "");
                    cmd.Parameters.AddWithValue("@IC0", node["IC0"] != null ? node["IC0"].InnerText : "");
                    cmd.Parameters.AddWithValue("@IC1", node["IC1"] != null ? node["IC1"].InnerText : "");
                    cmd.Parameters.AddWithValue("@IC5", node["IC5"] != null ? node["IC5"].InnerText : "");
                    cmd.Parameters.AddWithValue("@PostalCode", node["PostalCode"] != null ? node["PostalCode"].InnerText : "");
                    cmd.Parameters.AddWithValue("@CustomerNameInEnglish", node["CustomerNameInEnglish"] != null ? node["CustomerNameInEnglish"].InnerText : "");
                    cmd.Parameters.AddWithValue("@CustomerNameInEnglish2", node["CustomerNameInEnglish2"] != null ? node["CustomerNameInEnglish2"].InnerText : "");
                    cmd.Parameters.AddWithValue("@CityInEnglish", node["CityInEnglish"] != null ? node["CityInEnglish"].InnerText : "");
                    cmd.Parameters.AddWithValue("@OrderNumber", node["OrderNumber"] != null ? node["OrderNumber"].InnerText : "");
                    cmd.Parameters.AddWithValue("@InvoiceNo", node["InvoiceNo"] != null ? node["InvoiceNo"].InnerText : "");
                    cmd.Parameters.AddWithValue("@DocumentType", node["DocumentType"] != null ? node["DocumentType"].InnerText : "");
                    cmd.Parameters.AddWithValue("@MaterialDescription", node["MaterialDescription"] != null ? node["MaterialDescription"].InnerText : "");
                    cmd.Parameters.AddWithValue("@OrderCode", node["OrderCode"] != null ? node["OrderCode"].InnerText : "");
                    cmd.Parameters.AddWithValue("@ProductGroup", node["ProductGroup"] != null ? node["ProductGroup"].InnerText : "");
                    cmd.Parameters.AddWithValue("@ProductArea", node["ProductArea"] != null ? node["ProductArea"].InnerText : "");
                    cmd.Parameters.AddWithValue("@IOGross_SC", node["IOGross_SC"] != null && node["IOGross_SC"].InnerText != "" && node["IOGross_SC"].InnerText.Trim() != "-" ? System.Convert.ToDecimal(node["IOGross_SC"].InnerText) : 0);
                    cmd.Parameters.AddWithValue("@IOQty_BaseUnit", node["IOQty_BaseUnit"] != null && node["IOQty_BaseUnit"].InnerText != "" && node["IOQty_BaseUnit"].InnerText.Trim() != "-" ? System.Convert.ToDecimal(node["IOQty_BaseUnit"].InnerText) : 0);
                    cmd.Parameters.AddWithValue("@SLGross_SC", node["SLGross_SC"] != null && node["SLGross_SC"].InnerText != "" && node["SLGross_SC"].InnerText.Trim() != "-" ? System.Convert.ToDecimal(node["SLGross_SC"].InnerText) : 0);
                    cmd.Parameters.AddWithValue("@SLQty_BaseUnit", node["SLQty_BaseUnit"] != null && node["SLQty_BaseUnit"].InnerText != "" && node["SLQty_BaseUnit"].InnerText.Trim() != "-" ? System.Convert.ToDecimal(node["SLQty_BaseUnit"].InnerText) : 0);
                    break;
                case Constants.ExchangeRatesTable:
                    cmd.CommandText = Constants.strInsertExchangeRates;
                    cmd.Parameters.AddWithValue("@CountryRegion", node["CountryRegion"] != null ? node["CountryRegion"].InnerText.Trim() : "");
                    cmd.Parameters.AddWithValue("@Rate", node["Rate"] != null ? System.Convert.ToDecimal(node["Rate"].InnerText) : 0);
                    break;
                default:
                    break;
            }
            return cmd;
        }

        private SqlCommand SetParametersForUpdate(string tableName, XmlNode node)
        {
            SqlCommand cmd = new SqlCommand();
            switch (tableName)
            {
                case Constants.OrderTable:
                    cmd.CommandText = Constants.strInsertOrderUpdate;
                    cmd.Parameters.AddWithValue("@OrderNo", node["OrderNo"] != null ? node["OrderNo"].InnerText : "");
                    cmd.Parameters.AddWithValue("@OrderType", node["OrderType"] != null ? node["OrderType"].InnerText : "");
                    break;
                default:
                    break;
            }
            return cmd;
        }

      
        private void Browse_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.Filter = "excel files (*.xlsx)|*.xls;*.xlsx|All files (*.*)|*.*";
            Nullable<bool> result = dlg.ShowDialog();

            if(result == true)
            {
                this.filePath.Text = dlg.FileName;
            }
        }
    }
}
