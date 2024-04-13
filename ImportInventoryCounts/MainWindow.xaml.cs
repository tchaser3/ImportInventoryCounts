/* Title:           Import Inventory
 * Date:            5-13-20
 * Author:          Terry Holmes
 * 
 * Description:     This will allow a complete warehouse to be imported for select part numbers */

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
using DataValidationDLL;
using InventoryDLL;
using NewEmployeeDLL;
using NewEventLogDLL;
using NewPartNumbersDLL;
using Microsoft.Win32;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;

namespace ImportInventoryCounts
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        //setting up the classes
        WPFMessagesClass TheMessagesClass = new WPFMessagesClass();
        DataValidationClass TheDataValidationClass = new DataValidationClass();
        InventoryClass TheInventoryClass = new InventoryClass();
        EmployeeClass TheEmployeeClass = new EmployeeClass();
        EventLogClass TheEventLogClass = new EventLogClass();
        PartNumberClass ThePartNumberClass = new PartNumberClass();

        //setting up the data
        FindPartsWarehousesDataSet TheFindPartsWarehousesDateSet = new FindPartsWarehousesDataSet();
        ImportedPartsDataSet TheImportedPartsDataSet = new ImportedPartsDataSet();
        FindPartByPartNumberDataSet TheFindPartByPartNumberDataSet = new FindPartByPartNumberDataSet();
        FindPartByJDEPartNumberDataSet TheFindPartByJDEPartNumberDataSet = new FindPartByJDEPartNumberDataSet();
        FindWarehouseInventoryPartDataSet TheFindWarehouseInventoryPartDataSet = new FindWarehouseInventoryPartDataSet();

        //setting up global variables
        int gintWarehouseID;

        public MainWindow()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            int intCounter;
            int intNumberOfRecords;

            TheFindPartsWarehousesDateSet = TheEmployeeClass.FindPartsWarehouses();

            intNumberOfRecords = TheFindPartsWarehousesDateSet.FindPartsWarehouses.Rows.Count - 1;
            cboSelectWarehouse.Items.Add("Select Warehouse");

            for(intCounter = 0; intCounter <= intNumberOfRecords; intCounter++)
            {
                cboSelectWarehouse.Items.Add(TheFindPartsWarehousesDateSet.FindPartsWarehouses[intCounter].FirstName);
            }

            cboSelectWarehouse.SelectedIndex = 0;
            expImportExcel.IsEnabled = false;
            expProcessImport.IsEnabled = false;
        }

        private void expCloseProgram_Expanded(object sender, RoutedEventArgs e)
        {
            expCloseProgram.IsExpanded = false;
            TheMessagesClass.CloseTheProgram();
        }

        private void Grid_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            DragMove();
        }

        private void cboSelectWarehouse_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            int intSelectedIndex;

            intSelectedIndex = cboSelectWarehouse.SelectedIndex - 1;

            if (intSelectedIndex > -1)
            {
                gintWarehouseID = TheFindPartsWarehousesDateSet.FindPartsWarehouses[intSelectedIndex].EmployeeID;

                expImportExcel.IsEnabled = true;
            }
            
        }

        private void expImportExcel_Expanded(object sender, RoutedEventArgs e)
        {
            Excel.Application xlDropOrder;
            Excel.Workbook xlDropBook;
            Excel.Worksheet xlDropSheet;
            Excel.Range range;

            int intColumnRange = 0;
            int intCounter;
            int intNumberOfRecords;
            string strJDEPartNumber;
            string strPartNumber;
            int intCurrentCount;
            int intPartID = 0;
            string strPartDescription;
            int intOldCount = 0;
            string strCurrentCount;
            int intRecordsReturned;
            int intTransactionID = 0;

            try
            {
                TheImportedPartsDataSet.importedparts.Rows.Clear();

                Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
                dlg.FileName = "Document"; // Default file name
                dlg.DefaultExt = ".xlsx"; // Default file extension
                dlg.Filter = "Excel (.xlsx)|*.xlsx"; // Filter files by extension

                // Show open file dialog box
                Nullable<bool> result = dlg.ShowDialog();

                // Process open file dialog box results
                if (result == true)
                {
                    // Open document
                    string filename = dlg.FileName;
                }

                PleaseWait PleaseWait = new PleaseWait();
                PleaseWait.Show();

                xlDropOrder = new Excel.Application();
                xlDropBook = xlDropOrder.Workbooks.Open(dlg.FileName, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                xlDropSheet = (Excel.Worksheet)xlDropOrder.Worksheets.get_Item(1);

                range = xlDropSheet.UsedRange;
                intNumberOfRecords = range.Rows.Count;
                intColumnRange = range.Columns.Count;

                for (intCounter = 5; intCounter <= intNumberOfRecords; intCounter++)
                {
                    strJDEPartNumber = Convert.ToString((range.Cells[intCounter, 1] as Excel.Range).Value2).ToUpper();
                    strPartNumber = Convert.ToString((range.Cells[intCounter, 3] as Excel.Range).Value2).ToUpper();
                    strCurrentCount = Convert.ToString((range.Cells[intCounter, 6] as Excel.Range).Value2).ToUpper();
                    intCurrentCount = Convert.ToInt32(strCurrentCount);

                    TheFindPartByPartNumberDataSet = ThePartNumberClass.FindPartByPartNumber(strPartNumber);

                    intRecordsReturned = TheFindPartByPartNumberDataSet.FindPartByPartNumber.Rows.Count;

                    if(intRecordsReturned == 1)
                    {
                        intPartID = TheFindPartByPartNumberDataSet.FindPartByPartNumber[0].PartID;
                        strPartDescription = TheFindPartByPartNumberDataSet.FindPartByPartNumber[0].PartDescription;
                    }
                    else
                    {
                        TheFindPartByJDEPartNumberDataSet = ThePartNumberClass.FindPartByJDEPartNumber(strJDEPartNumber);

                        intRecordsReturned = TheFindPartByJDEPartNumberDataSet.FindPartByJDEPartNumber.Rows.Count;

                        if(intRecordsReturned == 1)
                        {
                            intPartID = TheFindPartByJDEPartNumberDataSet.FindPartByJDEPartNumber[0].PartID;
                            strPartDescription = TheFindPartByJDEPartNumberDataSet.FindPartByJDEPartNumber[0].PartDescription;
                            
                        }
                        else
                        {
                            TheMessagesClass.ErrorMessage("Something is messed up");
                            return;
                        }
                    }

                    TheFindWarehouseInventoryPartDataSet = TheInventoryClass.FindWarehouseInventoryPart(intPartID, gintWarehouseID);

                    intRecordsReturned = TheFindWarehouseInventoryPartDataSet.FindWarehouseInventoryPart.Rows.Count;

                    if(intRecordsReturned == 1)
                    {
                        intOldCount = TheFindWarehouseInventoryPartDataSet.FindWarehouseInventoryPart[0].Quantity;
                        intTransactionID = TheFindWarehouseInventoryPartDataSet.FindWarehouseInventoryPart[0].TransactionID;
                    }
                    else
                    {
                        intOldCount = 0;
                    }


                    ImportedPartsDataSet.importedpartsRow NewPartRow = TheImportedPartsDataSet.importedparts.NewimportedpartsRow();

                    NewPartRow.JDEPartNumber = strJDEPartNumber;
                    NewPartRow.NewCount = intCurrentCount;
                    NewPartRow.OldCount = intOldCount;
                    NewPartRow.PartDescription = strPartDescription;
                    NewPartRow.PartID = intPartID;
                    NewPartRow.PartNumber = strPartNumber;
                    NewPartRow.TransactionID = intTransactionID;
                    NewPartRow.WarehouseID = gintWarehouseID;

                    TheImportedPartsDataSet.importedparts.Rows.Add(NewPartRow);
                }

                dgrInventory.ItemsSource = TheImportedPartsDataSet.importedparts;

                PleaseWait.Close();

                expProcessImport.IsEnabled = true;
                expImportExcel.IsExpanded = false;

            }
            catch (Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Import Inventory Counts // Import Excel  " + Ex.Message);

                TheMessagesClass.ErrorMessage(Ex.ToString());
            }
        }

        private void expProcessImport_Expanded(object sender, RoutedEventArgs e)
        {
            int intCounter;
            int intNumberOfRecords;
            bool blnFatalError = false;
            int intNewCount;
            int intTransactionID;

            try
            {
                intNumberOfRecords = TheImportedPartsDataSet.importedparts.Rows.Count - 1;

                for(intCounter = 0; intCounter <= intNumberOfRecords; intCounter++)
                {
                    intNewCount = TheImportedPartsDataSet.importedparts[intCounter].NewCount;
                    intTransactionID = TheImportedPartsDataSet.importedparts[intCounter].TransactionID;

                    blnFatalError = TheInventoryClass.UpdateInventoryPart(intTransactionID, intNewCount);

                    if(blnFatalError == true)
                        throw new Exception();
                }

                TheMessagesClass.InformationMessage("Counts Have Been Updated");
            }
            catch (Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Import Inventory Counts // Process Import Expander " + Ex.Message);

                TheMessagesClass.ErrorMessage(Ex.ToString());
            }
        }
    }
}
