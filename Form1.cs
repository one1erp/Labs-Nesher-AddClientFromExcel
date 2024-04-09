using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using Common;
using DAL;
using Microsoft.Office.Interop.Excel;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;

namespace AddClientFromExcel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            //   AddSingleClient();





            //addressId = dal.GetNewID("SQ_ADDRESS");
            //clientId = dal.GetNewID("SQ_CLIENT");
            //Forms1();
            //    AddADress();


        }



        private long clientId = 1;
        private long addressId = 1;

        private Client AddSingleClient(string name, string code)
        {

            var client = new Client();
            client.ClientId = clientId++; //todo:take id from sequence
            client.Name = name;
            client.ClientCode = code;
            client.VERSION_STATUS = "A";
            client.VERSION = "1";
            return client;


        }

        private List<Client> _clients = new List<Client>();

        public void Forms1()
        {
            //Get excel
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();

            Microsoft.Office.Interop.Excel.Workbook wb = app.Workbooks.Open(@"C:\temp\Clients1.xlsx", Type.Missing,
                Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing,
                Type.Missing,
                Type.Missing, Type.Missing, Type.Missing,
                Type.Missing,
                Type.Missing, Type.Missing);

            //Get worksheet
            Worksheet sheet = (Worksheet)wb.Sheets["Clients"];

            Range excelRange = sheet.UsedRange;

            foreach (Microsoft.Office.Interop.Excel.Range row in excelRange.Rows)
            {
                int rowNumber = row.Row;


                if (rowNumber > 1)
                {



                    string[] A4D4 = GetRange("A" + rowNumber + ":M" + rowNumber + "", sheet);
                }

            }
            dal.SaveChanges();
            radGridView1.DataSource = clients;
            radGridView2.DataSource = addresses;

        }

        public
            string[] GetRange(string range, Worksheet excelWorksheet)
        {
            Microsoft.Office.Interop.Excel.Range workingRangeCells =
                excelWorksheet.get_Range(range, Type.Missing);
            //workingRangeCells.Select();

            System.Array array = (System.Array)workingRangeCells.Cells.Value2;
            string[] arrayS = this.ConvertToStringArray(array);

            return arrayS;

        }

        private List<Client> clients = new List<Client>();
        private List<Address> addresses = new List<Address>();
        private DataLayer dal;


        private string[] ConvertToStringArray(Array array)
        {

            var foo = new List<string>();
            var s = array.GetEnumerator();


            while (s.MoveNext())
            {
                if (s.Current != null) foo.Add(s.Current.ToString());
                else
                {
                    foo.Add("");
                }
            }

            //d  string[] foo = array.OfType<object>().Select(o => o.ToString()).ToArray();




            if (!string.IsNullOrEmpty(foo[0]) || !string.IsNullOrEmpty(foo[1]))
            {
                Client client;
                client = dal.GetClientByName(foo[1]);
                if (client != null)
                {

                    var BBaddress = dal.TestAddress(client.ClientId).FirstOrDefault(x => x.AddressType == "B");
                    var CCaddress = dal.TestAddress(client.ClientId).FirstOrDefault(x => x.AddressType == "C");
                    if (BBaddress != null)
                    {
                        UpdateBilingAddress(client, foo, BBaddress);
                    }
                    else
                    {
                        AddBilingAddress(client, foo);
                    }
                    if (CCaddress != null)
                    {
                        UpdateCompanyaddress(client, foo, CCaddress);
                    }
                    else
                    {
                        AddCompanyaddress(client, foo);
                    }
                }
                else //add new client
                {
                    client = AddSingleClient(foo[1], foo[0]);
                    clients.Add(client);
                    dal.AddClient(client);
                    AddNewAddress(client, foo);
                }





                //if (address != null)
                //{
                //    if (foo[11].Length >= 20)
                //    {
                //        address.Phone = foo[11].Substring(0, 20);
                //    }
                //}
                //else
                //{
                //    address.Phone = foo[11];
                //}






            }


            return null;
            //   dal.SaveChanges();
            //  }
            //client.ClientCode = foo[0];
            //client.Name = foo[1];
            //client.VERSION_STATUS = "A";
            //client.VERSION = "1";
            //clients.Add(client);
        }



        public void AddNewAddress(Client client, List<string> foo)
        {
            var ca = AddCompanyaddress(client, foo);

            var ba = AddBilingAddress(client, foo);
        }

        private Address AddBilingAddress(Client client, List<string> foo)
        {
            if (!string.IsNullOrEmpty(foo[10]))
            {
                var dAddress = new Address();
                dAddress.AddressId = addressId++; //todo:take id from sequence
                dAddress.AddressType = "B";
                dAddress.AddressTableName = "CLIENT";
                dAddress.AddressItemId = client.ClientId;


                dAddress.ContactMan = foo[10];
                if (foo[11].Length >= 20)
                    dAddress.Phone = foo[11].Substring(0, 20);
                else
                {
                    dAddress.Phone = foo[11];
                }

                dAddress.Email = foo[12];
                addresses.Add(dAddress);
                dal.AddAddress(dAddress);
                return dAddress;
            }
            else
            {
                return null;
            }
        }

        private Address AddCompanyaddress(Client client, List<string> foo)
        {
            var cAddress = new Address();
            cAddress.AddressId = addressId++; //todo:take id from sequence
            cAddress.AddressItemId = client.ClientId;
            cAddress.AddressTableName = "CLIENT";


            cAddress.AddressType = "C";
            if (foo[2].Length >= 50)
                cAddress.FullAddress = foo[2].Substring(0, 50);
            else
                cAddress.FullAddress = foo[2];


            cAddress.ContactMan = foo[4];
            if (foo[5].Length >= 20)
                cAddress.Phone = foo[5].Substring(0, 20);
            else
                cAddress.Phone = foo[5];

            if (foo[6].Length >= 20)
                cAddress.ContactMobile = foo[6].Substring(0, 20);
            else
                cAddress.ContactMobile = foo[6];



            if (foo[7].Length >= 20)
                cAddress.Fax = foo[7].Substring(0, 20);
            else
                cAddress.Fax = foo[7];

            cAddress.Email = foo[8];

            cAddress.H_P = foo[9];
            addresses.Add(cAddress);

            dal.AddAddress(cAddress);
            return cAddress;
        }




        private Address UpdateBilingAddress(Client client, List<string> foo, Address dAddress)
        {
            if (!string.IsNullOrEmpty(foo[10]))
            {




                dAddress.ContactMan = foo[10];
                if (foo[11].Length >= 20)
                    dAddress.Phone = foo[11].Substring(0, 20);
                else
                {
                    dAddress.Phone = foo[11];
                }

                dAddress.Email = foo[12];
                addresses.Add(dAddress);
                return dAddress;
            }
            else
            {
                return null;
            }
        }

        private void UpdateCompanyaddress(Client client, List<string> foo, Address cAddress)
        {







            if (foo[2].Length >= 50)
                cAddress.FullAddress = foo[2].Substring(0, 50);
            else
                cAddress.FullAddress = foo[2];


            cAddress.ContactMan = foo[4];
            if (foo[5].Length >= 20)
                cAddress.Phone = foo[5].Substring(0, 20);
            else
                cAddress.Phone = foo[5];

            if (foo[6].Length >= 20)
                cAddress.ContactMobile = foo[6].Substring(0, 20);
            else
                cAddress.ContactMobile = foo[6];


            if (foo[7].Length >= 20)
                cAddress.Fax = foo[7].Substring(0, 20);
            else
                cAddress.Fax = foo[7];

            cAddress.Email = foo[8];

            cAddress.H_P = foo[9];
            addresses.Add(cAddress);



        }

        private void button1_Click(object sender, EventArgs e)
        {
            //int id = 10;
            //try
            //{
            //    dal = new DataLayer();
            //    dal.Connect();


            //    var clients = dal.GetAll<XmlStorage>();
            //    foreach (var client in clients)
            //    {




            //        var cd = new U_CLIENT_DATA();
            //        cd.U_CLIENT_DATA_ID = id++;
            //        cd.NAME = client.Name + "," + "Cosmetics"; //cosmetic
            //        cd.U_CLIENT_ID = client.ClientId;
            //        cd.U_LAB_ID = 3;
            //        cd.VERSION = "1";
            //        cd.VERSION_STATUS = "A";
            //        cd.U_COA_COLUMNS = client.DefaultCOA_column;
            //        dal.AddClientData(cd);


            //        cd = new U_CLIENT_DATA();
            //        cd.U_CLIENT_DATA_ID = id++;

            //        cd.NAME = client.Name + "," + "Water"; //water
            //        cd.U_CLIENT_ID = client.ClientId;
            //        cd.U_LAB_ID = 2;
            //        cd.VERSION = "1";
            //        cd.VERSION_STATUS = "A";
            //        cd.U_COA_COLUMNS = client.DefaultCOA_column;
            //        dal.AddClientData(cd);

            //    }
            //    dal.SaveChanges();
            //}
            //catch (Exception ex)
            //{

            //    MessageBox.Show("Error " + ex.Message);
            //}
        }

        private void button2_Click(object sender, EventArgs e)
        {

            try
            {


                var dal = new DataLayer();
                dal.Connect();
                var xs = dal.GetAll<XmlStorage>().Where(x => x.TableName != "U_LABS_INFO");
                int id = 700;
                foreach (XmlStorage xmlStorage in xs)
                {
                    var s = new XmlStorage();
                    s.XmlStorageId = id++;
                    s.EntityId = xmlStorage.EntityId;
                    s.LAB_ID = 2;
                    s.XmlData = xmlStorage.XmlData;
                    s.TableName = xmlStorage.TableName;
                    dal.AddXmlStorage(s);
                }
                dal.SaveChanges();

            }

            catch (Exception ee)
            {
                Logger.WriteLogFile(ee);
                MessageBox.Show(ee.Message);
            }
            MessageBox.Show("END");




        }
    }
}

//public void ReadFromExcel()
//{
//    this.openFileDialog1.FileName = "*.xls";
//    if (this.openFileDialog1.ShowDialog() == DialogResult.OK)
//    {

//        Excel.Workbook theWorkbook = ExcelObj.Workbooks.Open(
//           openFileDialog1.FileName, 0, true, 5,
//            "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false,
//            0, true);
//        Excel.Sheets sheets = theWorkbook.Worksheets;
//        Excel.Worksheet worksheet = (Excel.Worksheet)sheets.get_Item(1);
//        for (int i = 1; i <= 10; i++)
//        {
//            Excel.Range range = worksheet.get_Range("A" + i.ToString(), "J" + i.ToString());
//            System.Array myvalues = (System.Array)range.Cells.Value;
//            string[] strArray = ConvertToStringArray(myvalues);
//        }
//    }
//}
//        public void ExcelDB()
//        {
//            try
//            {

//                excel.Open(Application.StartupPath + "\\MAILING.XLS");
//                excel.Range();
//                string file = Application.StartupPath + "\\MAILING.XLS";
//                sbConn.Append(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + file);
//                sbConn.Append(";Extended Properties=");
//                sbConn.Append(Convert.ToChar(34));
//                sbConn.Append("Excel 8.0;HDR=NO;IMEX=2");
//                sbConn.Append(Convert.ToChar(34));
//                var cnExcel = new OleDbConnection(sbConn.ToString());
//                var cmdExcel = new OleDbCommand("Select * From Clientlist", cnExcel);
//            }
//            catch
//            {
//                MessageBox.Show("Ernstige fout, klantenbestand niet gevonden!", "Error 03", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Stop);
//            }
//        }
//        public void FillList()
//        {
//            try
//            {
//                Cursor.Current = Cursors.WaitCursor;
//                cnExcel.Open();
//                drExcel = cmdExcel.ExecuteReader();
//                while (drExcel.Read())
//                {
//                    if (drExcel["Name"].ToString() != "")
//                    {
//                        LogicControl.form.clientList.Items.Add(drExcel["Name"].ToString());
//                        string[] listArray = {drExcel["Name"].ToString(),
//drExcel["Address"].ToString(),
//drExcel["City"].ToString(),
//drExcel["Phone"].ToString(),
//drExcel["Fax"].ToString(),
//drExcel["VAT"].ToString()
//};
//                        lvi = new ListViewItem(listArray);
//                        LogicControl.form.klantenList.Items.Add(lvi);
//                    }
//                }
//                drExcel.Close();
//                cnExcel.Close();
//                Cursor.Current = Cursors.Arrow;
//            }
//            catch
//            {
//                excel.Quit();
//                MessageBox.Show("Ernstige fout, foutief klantenbestand!", "Error 05", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Stop);
//            }
//        }


