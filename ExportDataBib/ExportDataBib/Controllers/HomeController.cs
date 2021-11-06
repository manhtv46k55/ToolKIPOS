using ExportDataBib.Data;
using ExportDataBib.Models;
using MarcXmlParser;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace ExportDataBib.Controllers
{
    public class HomeController : Controller
    {
        ConnectDB cnn = new ConnectDB();
        public ActionResult Index()
        {
            //Lấy ra List Bộ Sưu tập

            List<KPS_DMD_COLLECTION> ListCollection = new List<KPS_DMD_COLLECTION>();
            ListCollection = cnn.KPS_DMD_COLLECTION.ToList();

            ViewBag.ListCL = ListCollection;


            return View();
        }

        public FileResult GetData(/*string sArrNameField, */int iCollectionID, DateTime? dtFromDate, DateTime? dtToDate)
        {
            List<KPS_DMD> ListBib = cnn.KPS_DMD.ToList();
            if (iCollectionID == 0)
            {
                ListBib = ListBib.ToList();
            }

            if (iCollectionID > 0)
            {
                ListBib = ListBib.Where(x => x.COLLECTION_ID.Equals(iCollectionID)).ToList();
            }
            if (dtFromDate != null)
            {
                ListBib = ListBib.Where(x => x.RECORD_CREATED_DATE >= dtFromDate).ToList();
            }
            if (dtToDate != null)
            {
                ListBib = ListBib.Where(x => x.RECORD_CREATED_DATE <= dtToDate).ToList();
            }
            List<FinalDataBib> ListData = new List<FinalDataBib>();

            FileInfo file = new FileInfo(Server.MapPath(@"~/Template/DataBib.xlsx"));
            ExcelPackage excelPack = new ExcelPackage(file);
            ExcelWorkbook excelWorkBook = excelPack.Workbook;
            ExcelWorksheet excelWorkSheet = excelWorkBook.Worksheets["Sheet1"];
            int numRow = 1;
            int stt = 0;

            foreach (var bib in ListBib)
            {
                CRecord myRec = new CRecord();
                CControlfield Cf = new CControlfield();
                CDatafield Df = new CDatafield();
                CSubfield Sf = new CSubfield();
                myRec.load_Xml(bib.XMLDATA);

                FinalDataBib DT = new FinalDataBib();

                //Control Field
                Cf = myRec.Controlfields.Controlfield("001");
                DT.Cf001 = Cf.Value;

                Cf.ReConstruct();
                Cf = myRec.Controlfields.Controlfield("002");
                DT.Cf002 = Cf.Value;

                Cf.ReConstruct();
                Cf = myRec.Controlfields.Controlfield("004");
                DT.Cf004 = Cf.Value;

                Cf.ReConstruct();
                Cf = myRec.Controlfields.Controlfield("008");
                DT.Cf008 = Cf.Value;

                //Data field
                Df.ReConstruct();
                Df = myRec.Datafields.Datafield("039");
                DT.Df039 = Df.InnerText;

                Df.ReConstruct();
                Df = myRec.Datafields.Datafield("040");
                DT.Df040 = Df.InnerText;

                Df.ReConstruct();
                Df = myRec.Datafields.Datafield("041");
                DT.Df041 = Df.InnerText;

                Df.ReConstruct();
                Df = myRec.Datafields.Datafield("044");
                DT.Df044 = Df.InnerText;

                Df.ReConstruct();
                Df = myRec.Datafields.Datafield("082");
                DT.Df082 = Df.InnerText;

                Df.ReConstruct();
                Df = myRec.Datafields.Datafield("084");
                DT.Df084 = Df.InnerText;

                Df.ReConstruct();
                Df = myRec.Datafields.Datafield("082");
                DT.Df082 = Df.InnerText;

                Df.ReConstruct();
                Df = myRec.Datafields.Datafield("100");
                DT.Df100 = Df.InnerText;

                Df.ReConstruct();
                Df = myRec.Datafields.Datafield("110");
                DT.Df110 = Df.InnerText;

                Df.ReConstruct();
                Df = myRec.Datafields.Datafield("245");
                DT.Df245 = Df.InnerText;

                Df.ReConstruct();
                Df = myRec.Datafields.Datafield("260");
                DT.Df260 = Df.InnerText;

                Df.ReConstruct();
                Df = myRec.Datafields.Datafield("300");
                DT.Df300 = Df.InnerText;

                Df.ReConstruct();
                Df = myRec.Datafields.Datafield("500");
                DT.Df500 = Df.InnerText;

                Df.ReConstruct();
                Df = myRec.Datafields.Datafield("520");
                DT.Df520 = Df.InnerText;

                //650 - trường lặp
                Df.ReConstruct();
                string sValue650 = "";
                for (int j = 0; j < myRec.Datafields.Count; j++)
                {
                    Df = myRec.Datafields.Datafield(j);
                    if (Df.Tag == "650")
                    {
                        for (int k = 0; k < Df.Subfields.Count; k++)
                        {
                            sValue650 += Df.Subfields.Subfield(k).Value + "; ";
                        }
                    }

                }
                if (sValue650 != null && sValue650.Length > 0)
                {
                    sValue650 = sValue650.Remove(sValue650.Length - 2).ToString();
                    DT.Df650 = sValue650;
                }
                else
                {
                    DT.Df650 = "";
                }


                //653 - trường lặp
                Df.ReConstruct();
                string sValue653 = "";
                for (int j = 0; j < myRec.Datafields.Count; j++)
                {
                    Df = myRec.Datafields.Datafield(j);
                    if (Df.Tag == "653")
                    {
                        for (int k = 0; k < Df.Subfields.Count; k++)
                        {
                            sValue653 += Df.Subfields.Subfield(k).Value + "; ";
                        }
                    }

                }
                if (sValue653 != null && sValue653.Length > 0)
                {
                    sValue653 = sValue653.Remove(sValue653.Length - 2).ToString();
                    DT.Df653 = sValue653;
                }
                else
                {
                    DT.Df653 = "";
                }

                //700 - trường lặp
                Df.ReConstruct();
                string sValue700 = "";
                for (int j = 0; j < myRec.Datafields.Count; j++)
                {
                    Df = myRec.Datafields.Datafield(j);
                    if (Df.Tag == "700")
                    {
                        for (int k = 0; k < Df.Subfields.Count; k++)
                        {
                            sValue700 += Df.Subfields.Subfield(k).Value + "; ";
                        }
                    }

                }
                if (sValue700 != null && sValue700.Length > 0)
                {
                    sValue700 = sValue700.Remove(sValue700.Length - 2).ToString();
                    DT.Df700 = sValue700;
                }
                else
                {
                    DT.Df700 = "";
                }

                //710 - trường lặp
                Df.ReConstruct();
                string sValue710 = "";
                for (int j = 0; j < myRec.Datafields.Count; j++)
                {
                    Df = myRec.Datafields.Datafield(j);
                    if (Df.Tag == "710")
                    {
                        for (int k = 0; k < Df.Subfields.Count; k++)
                        {
                            sValue710 += Df.Subfields.Subfield(k).Value + "; ";
                        }
                    }

                }
                if (sValue710 != null && sValue710.Length > 0)
                {
                    sValue710 = sValue710.Remove(sValue710.Length - 2).ToString();
                    DT.Df710 = sValue710;
                }
                else
                {
                    DT.Df710 = "";
                }

                //773
                Df.ReConstruct();
                Df = myRec.Datafields.Datafield("773");
                DT.Df773 = Df.InnerText;

                //911
                Df.ReConstruct();
                Df = myRec.Datafields.Datafield("911");
                DT.Df911 = Df.InnerText;

                //912
                Df.ReConstruct();
                Df = myRec.Datafields.Datafield("912");
                DT.Df912 = Df.InnerText;

                //913
                Df.ReConstruct();
                Df = myRec.Datafields.Datafield("913");
                DT.Df913 = Df.InnerText;

                //925
                Df.ReConstruct();
                Df = myRec.Datafields.Datafield("925");
                DT.Df925 = Df.InnerText;

                //926
                Df.ReConstruct();
                Df = myRec.Datafields.Datafield("926");
                DT.Df926 = Df.InnerText;

                //927
                Df.ReConstruct();
                Df = myRec.Datafields.Datafield("927");
                DT.Df927 = Df.InnerText;

                ListData.Add(DT);
            }

            foreach (var b in ListData)
            {
                stt++;
                numRow++;
                excelWorkSheet.Cells["A" + numRow].Value = stt;
                excelWorkSheet.Cells["B" + numRow].Value = b.Cf001;
                excelWorkSheet.Cells["C" + numRow].Value = b.Cf002;
                excelWorkSheet.Cells["D" + numRow].Value = b.Cf004;
                excelWorkSheet.Cells["E" + numRow].Value = b.Cf005;
                excelWorkSheet.Cells["F" + numRow].Value = b.Cf008;
                excelWorkSheet.Cells["G" + numRow].Value = b.Cf009;
                excelWorkSheet.Cells["H" + numRow].Value = b.Df039;
                excelWorkSheet.Cells["I" + numRow].Value = b.Df040;
                excelWorkSheet.Cells["J" + numRow].Value = b.Df041;
                excelWorkSheet.Cells["K" + numRow].Value = b.Df044;
                excelWorkSheet.Cells["L" + numRow].Value = b.Df082;
                excelWorkSheet.Cells["M" + numRow].Value = b.Df084;
                excelWorkSheet.Cells["N" + numRow].Value = b.Df100;
                excelWorkSheet.Cells["O" + numRow].Value = b.Df110;
                excelWorkSheet.Cells["P" + numRow].Value = b.Df245;
                excelWorkSheet.Cells["Q" + numRow].Value = b.Df260;
                excelWorkSheet.Cells["R" + numRow].Value = b.Df300;
                excelWorkSheet.Cells["S" + numRow].Value = b.Df500;
                excelWorkSheet.Cells["T" + numRow].Value = b.Df520;
                excelWorkSheet.Cells["U" + numRow].Value = b.Df650;
                excelWorkSheet.Cells["V" + numRow].Value = b.Df653;
                excelWorkSheet.Cells["W" + numRow].Value = b.Df700;
                excelWorkSheet.Cells["X" + numRow].Value = b.Df710;
                excelWorkSheet.Cells["Y" + numRow].Value = b.Df773;
                excelWorkSheet.Cells["Z" + numRow].Value = b.Df911;
                excelWorkSheet.Cells["AA" + numRow].Value = b.Df912;
                excelWorkSheet.Cells["AB" + numRow].Value = b.Df913;
                excelWorkSheet.Cells["AC" + numRow].Value = b.Df925;
                excelWorkSheet.Cells["AD" + numRow].Value = b.Df926;
                excelWorkSheet.Cells["AE" + numRow].Value = b.Df927;

            }
            return File(excelPack.GetAsByteArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Data.xlsx");
        }
        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
    }
}