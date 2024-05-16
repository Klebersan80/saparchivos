using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ItemAttachments
{
    class Program
    {
        static void Main(string[] args)
        {
            //#################################################################
            // SAP Business One 9.2 Item Master Attachment Console Application
            // Changes Before Running the Application
            //
            // 1. Change the Appliations File Path and File Name. See below
            // 2. Change the Database connection details to suit your "own" database setup. See the App.config file
            //
            // Most Importantly before running this on Production please test on a Test installation first.

            // The following Lines need to be changed to suite whatever location you have store the file
            string FilePath = @"\\10.61.1.15\share\Staff\staff\BrendanB\";
            string FileName = "DataFile2.csv";

            // Check Source File Exists
            if (CheckFileExists(FilePath, FileName))
            {

                SAPbobsCOM.Company oCompany;
                SAP_Bob SAPBOB = new SAP_Bob();
                oCompany = SAPBOB._getCompanyConnection();

                Console.WriteLine(string.Format("Connecting To Database {0}.", oCompany.CompanyDB));
                // Check Connection is Opened
                if (oCompany.Connected == true)
                {

                    Console.WriteLine(string.Format("Connection Successful"));


                    Console.WriteLine(string.Format("Getting Records from File {0}.", FileName));
                    // Get Data from File
                    var FileData = GetFileContent(FilePath, FileName);

                    if (FileData.Count > 0)
                    {

                        Console.WriteLine(string.Format("Found {0} rows.", FileData.Count));

                        // Need to Group by Item Codes as I only have one AbsEntry Record to Link into OITM Field
                        var ItemCodes = FileData.GroupBy(n => new { n.ItemCode })
                                                    .Select(g => new {
                                                        g.Key.ItemCode
                                                    }).ToList();

                        foreach (var Item in ItemCodes)
                        {
                            Console.WriteLine(string.Format("Processing Item Code {0}.", Item.ItemCode));

                            // Get the Files That use the Item Code
                            var Files = FileData.Where(s => s.ItemCode.Equals(Item.ItemCode)).ToList();

                            Console.WriteLine(string.Format("Attaching Files {0} to Item Code {1}.", Files.Count, Item.ItemCode));
                            
                            // Add the one or More Attachments to the Item Code
                            int intABSEntry = AddAttachments(oCompany, Files);

                            // If added, the and ABS entry is returned, this is used to "link" the files to the Item Code
                            if (intABSEntry > 0)
                            {
                                Console.WriteLine(string.Format("Files Added, Attaching to Item Code {0}.", Item.ItemCode));

                                UpdateItem(oCompany, Item.ItemCode, intABSEntry);
                            }
                        }
                    }
                    else
                    {
                        Console.WriteLine(string.Format("No Records found in File {0}.", FileName));
                    }
                    
                }
                else
                {
                    Console.WriteLine(string.Format("Connection Failed"));
                }                
            }
            else
            {
                Console.WriteLine(string.Format("The File Name: {0} Could not be found at the Location: {1}.", FileName, FilePath));
            }
        }

        static bool CheckFileExists(string FileLocation,string FileName)
        {
            if (System.IO.File.Exists(string.Concat(FileLocation, FileName)))
            {
                return true;
            }
            return false;
        }

        static IList<Models.Attachments> GetFileContent(string FilePath, string FileName)
        {
            List<Models.Attachments> Att = new List<Models.Attachments>();

            using (var data = new System.IO.StreamReader(string.Concat(FilePath, FileName)))
            {
                while (!data.EndOfStream)
                {
                    var LineItem = data.ReadLine();
                    var Fields = LineItem.Split(',');

                    Att.Add(new Models.Attachments { ItemCode = Fields[0].ToString(), FilePath = Fields[1].ToString(), FileName = Fields[2].ToString(), FileExtension = Fields[3].ToString() });
                }
            }
            return Att;
        }

        static void UpdateItem(SAPbobsCOM.Company oCompany, string ItemCode, int absEntry)
        {
            if (oCompany.Connected == true)
            {
                SAPbobsCOM.Items oSAPItem = null;
                oSAPItem = (SAPbobsCOM.Items)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems);
                oSAPItem.GetByKey(ItemCode);

                oSAPItem.AttachmentEntry = absEntry;
                
                int StatusCode = oSAPItem.Update();

                if (StatusCode != 0)
                {
                    string Message = oCompany.GetLastErrorDescription();
                }
                else
                {
                    string ItemNumer;
                    oCompany.GetNewObjectCode(out ItemNumer);
                }
            }
        }                 

        static int AddAttachments(SAPbobsCOM.Company oCompany, List<Models.Attachments> Attachments)
        {
            int ABSEntry = -1;

            if (oCompany.Connected == true)
            {
                SAPbobsCOM.Attachments2 oSAPItemAttach = null;
                oSAPItemAttach = (SAPbobsCOM.Attachments2)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oAttachments2);

                int LineNo = 0;

                foreach (var Attachment in Attachments)
                {
                    oSAPItemAttach.Lines.Add();
                    oSAPItemAttach.Lines.SetCurrentLine(LineNo);
                    oSAPItemAttach.Lines.SourcePath = Attachment.FilePath;
                    oSAPItemAttach.Lines.FileName = Attachment.FileName;
                    oSAPItemAttach.Lines.FileExtension = Attachment.FileExtension;
                    oSAPItemAttach.Lines.Override = SAPbobsCOM.BoYesNoEnum.tYES;
                    
                    LineNo++;
                }

                ABSEntry = oSAPItemAttach.Add();

                if (ABSEntry != 0)
                {
                    string Message = oCompany.GetLastErrorDescription();
                }
                else
                {
                    string ItemNumer;
                    oCompany.GetNewObjectCode(out ItemNumer);
                    return int.Parse(ItemNumer);
                }
            }

            return ABSEntry;
        }
    }
}
