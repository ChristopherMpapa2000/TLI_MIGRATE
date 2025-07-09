using MigrateTLI.Item;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using WolfApprove.API2.Controllers.Utils;
using WolfApprove.API2.Extension;

namespace MigrateTLI
{
    class Program
    {
        public static void Log(String iText)
        {
            string pathlog = ItemConfig.LogFile;
            String logFolderPath = System.IO.Path.Combine(pathlog, DateTime.Now.ToString("yyyyMMdd"));

            if (!System.IO.Directory.Exists(logFolderPath))
            {
                System.IO.Directory.CreateDirectory(logFolderPath);
            }
            String logFilePath = System.IO.Path.Combine(logFolderPath, DateTime.Now.ToString("yyyyMMdd") + ".txt");

            try
            {
                using (System.IO.StreamWriter outfile = new System.IO.StreamWriter(logFilePath, true))
                {
                    System.Text.StringBuilder sbLog = new System.Text.StringBuilder();

                    String[] listText = iText.Split('|').ToArray();

                    foreach (String s in listText)
                    {
                        sbLog.AppendLine($"[{DateTime.Now:HH:mm:ss}] {s}");
                    }

                    outfile.WriteLine(sbLog.ToString());
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error writing log file: {ex.Message}");
            }
        }
        public static void LogError(String iText)
        {

            string pathlog = ItemConfig.LogFile;
            String logFolderPath = System.IO.Path.Combine(pathlog, DateTime.Now.ToString("yyyyMMdd"));

            if (!System.IO.Directory.Exists(logFolderPath))
            {
                System.IO.Directory.CreateDirectory(logFolderPath);
            }
            String logFilePath = System.IO.Path.Combine(logFolderPath, DateTime.Now.ToString("yyyyMMdd") + "LogError.txt");

            try
            {
                using (System.IO.StreamWriter outfile = new System.IO.StreamWriter(logFilePath, true))
                {
                    System.Text.StringBuilder sbLog = new System.Text.StringBuilder();

                    String[] listText = iText.Split('|').ToArray();

                    foreach (String s in listText)
                    {
                        sbLog.AppendLine($"[{DateTime.Now:HH:mm:ss}] {s}");
                    }

                    outfile.WriteLine(sbLog.ToString());
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error writing log file: {ex.Message}");
            }
        }
        static void Main()
        {
            try
            {
                Log("====== Start Process ====== : " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));
                Log(string.Format("Run batch as :{0}", System.Security.Principal.WindowsIdentity.GetCurrent().Name));

                ContextDataContext db = new ContextDataContext(ItemConfig.dbConnectionString);
                if (db.Connection.State == ConnectionState.Open)
                {
                    db.Connection.Close();
                    db.Connection.Open();
                }
                db.Connection.Open();
                db.CommandTimeout = 0;
                List<MSTEmployee> ListViewEmployee = db.MSTEmployees.ToList();
                var excelData = GetExcel.ReadExcelToDataTables(ItemConfig.PathFileExcel);
                if (excelData != null)
                {
                    Console.WriteLine("====== Start Migrate_Data ======");
                    List<Classcustom> listmemo = new List<Classcustom>();
                    string sheetName = ItemConfig.SheetName_Migrate;
                    if (excelData.ContainsKey(sheetName))
                    {
                        DataTable sheetData = excelData[sheetName];

                        for (int index = 3; index < sheetData.Rows.Count; index++)
                        {
                            string MAbvanceForm_Value = ImportAbvanceForm(db, sheetData.Rows[2], sheetData.Rows[index]);
                            var memo = InsertTrnmemo(db, MAbvanceForm_Value, sheetData.Rows[2], sheetData.Rows[index], ListViewEmployee);
                            if (memo != null)
                            {
                                string running = InsertTRNControlRunning(db, memo, sheetData.Rows[2], sheetData.Rows[index]);
                                if (!string.IsNullOrEmpty(running))
                                {
                                    InsertTRNAttachFile(db, memo,running);
                                }

                                InsertTrnLineapprove(db, ListViewEmployee, memo);
                                InsertTRNReferenceDoc(sheetData.Rows[index], db, memo);
                            }
                            else
                            {
                                LogError("InsertTrnmemo error: " + index);
                            }
                        }
                    }
                    else
                    {
                        Console.WriteLine($"Sheet '{sheetName}' ไม่พบในไฟล์ Excel");
                        LogError($"Sheet '{sheetName}' ไม่พบในไฟล์ Excel");
                    }
                    Console.WriteLine("====== End Migrate_Data ======");
                    Console.WriteLine("COMPLETE Count : " + listmemo.Count());
                    Console.WriteLine("Exit COMPLETE");
                    Log("COMPLETE Count : " + listmemo.Count());
                    Log("Exit COMPLETE");
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR");
                Console.WriteLine("Exit ERROR");
                LogError("ERROR");
                LogError("message: " + ex.Message);
                LogError("Exit ERROR");
            }
            finally
            {
                Log("====== End Process Process ====== : " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));
            }
        }
        public static string ImportAbvanceForm(ContextDataContext db, DataRow Rowheader, DataRow RowIndex)
        {
            try
            {
                var DestinationTemplate = db.MSTTemplates.Where(a => a.TemplateId == ItemConfig.TemplateID).ToList();
                JObject jsonAdvanceForm = JsonUtils.createJsonObject(DestinationTemplate.First().AdvanceForm);
                JArray itemsArray = (JArray)jsonAdvanceForm["items"];
                for (int r = 0; r < Rowheader.ItemArray.Length; r++)
                {
                    var label = Rowheader.ItemArray[r];
                    string value = string.Empty;
                    foreach (JObject jItems in itemsArray)
                    {
                        JArray jLayoutArray = (JArray)jItems["layout"];
                        if (jLayoutArray.Count >= 1)
                        {
                            JObject jTemplateL = (JObject)jLayoutArray[0]["template"];
                            JObject jData = (JObject)jLayoutArray[0]["data"];
                            if ((String)jTemplateL["label"] == label.ToString())
                            {

                                if (jData != null)
                                {
                                    #region value
                                    if (jData["value"] != null && jData["value"].Type == JTokenType.Null)
                                    {
                                        //กรณีเป็น value 
                                        if (label.ToString().Contains("วันที่"))
                                        {
                                            string Date = (string)RowIndex.ItemArray[r];
                                            if (Date != "" && Date != "-")
                                            {
                                                try
                                                {
                                                    double date = double.Parse(Date);
                                                    Date = DateTime.FromOADate(date).ToString("dd MMM yyyy", new CultureInfo("en-US"));
                                                }
                                                catch
                                                {
                                                    try
                                                    {
                                                        DateTime oDate = DateTime.ParseExact(Date, "dd/MM/yyyy", new CultureInfo("en-US"));
                                                        Date = oDate.ToString("dd MMM yyyy", new CultureInfo("en-US"));
                                                    }
                                                    catch
                                                    {
                                                        try
                                                        {
                                                            Date = Convert.ToDateTime(Date).ToString("dd MMM yyyy", new CultureInfo("en-US"));
                                                        }
                                                        catch
                                                        {
                                                            DateTime oDate = DateTime.Parse(Date);
                                                            Date = oDate.ToString("dd MMM yyyy", new CultureInfo("en-US"));
                                                        }
                                                    }
                                                }
                                            }
                                            else
                                            {
                                                Date = null;
                                            }
                                            jData["value"] = Date;
                                            break;
                                        }
                                        else
                                        {
                                            string type = (String)jTemplateL["type"];
                                            if (type == "cb")
                                            {
                                                JArray itemArray = new JArray();
                                                JObject attribute = (JObject)jTemplateL["attribute"];
                                                JArray items = (JArray)attribute["items"];

                                                int indexitem = items.Count();
                                                string[] Boxitem = new string[indexitem];
                                                if (jData != null)
                                                {
                                                    string itemvalue = (string)RowIndex.ItemArray[r];
                                                    string[] dataListArray = itemvalue.Split('|');

                                                    for (int F = 0; F < indexitem; F++)
                                                    {
                                                        string itemc = items[F]["item"].ToString();
                                                        if (dataListArray.Contains(itemc))
                                                        {
                                                            Boxitem[F] = "Y";
                                                        }
                                                        else
                                                        {
                                                            Boxitem[F] = "N";
                                                        }
                                                    }
                                                    string Valuecd = (string)RowIndex.ItemArray[r];
                                                    jData["value"] = Valuecd.Replace('|', ',');
                                                    jData["item"] = new JArray(Boxitem);
                                                    break;
                                                }
                                            }
                                            else
                                            {
                                                jData["value"] = (string)RowIndex.ItemArray[r];
                                                break;
                                            }
                                        }
                                    }
                                    #endregion
                                }
                            }
                            if (jLayoutArray.Count > 1)
                            {
                                JObject jTemplateR = (JObject)jLayoutArray[1]["template"];
                                JObject jData2 = (JObject)jLayoutArray[1]["data"];
                                if ((String)jTemplateR["label"] == label.ToString())
                                {
                                    if (jData2 != null)
                                    {
                                        #region value
                                        if (jData2["value"] != null && jData2["value"].Type == JTokenType.Null)
                                        {
                                            if (label.ToString().Contains("วันที่"))
                                            {
                                                string Date = (string)RowIndex.ItemArray[r];
                                                if (Date != "" && Date != "-")
                                                {
                                                    try
                                                    {
                                                        double date = double.Parse(Date);
                                                        Date = DateTime.FromOADate(date).ToString("dd MMM yyyy", new CultureInfo("en-US"));
                                                    }
                                                    catch
                                                    {
                                                        try
                                                        {
                                                            DateTime oDate = DateTime.ParseExact(Date, "dd/MM/yyyy", new CultureInfo("en-US"));
                                                            Date = oDate.ToString("dd MMM yyyy", new CultureInfo("en-US"));
                                                        }
                                                        catch
                                                        {
                                                            try
                                                            {

                                                                Date = Convert.ToDateTime(Date).ToString("dd MMM yyyy", new CultureInfo("en-US"));
                                                            }
                                                            catch
                                                            {
                                                                DateTime oDate = DateTime.Parse(Date);
                                                                Date = oDate.ToString("dd MMM yyyy", new CultureInfo("en-US"));
                                                            }
                                                        }
                                                    }
                                                }
                                                else
                                                {
                                                    Date = null;
                                                }
                                                jData2["value"] = Date;
                                                break;
                                            }
                                            else
                                            {
                                                string type = (String)jTemplateR["type"];
                                                if (type == "cb")
                                                {
                                                    JArray itemArray = new JArray();
                                                    JObject attribute = (JObject)jTemplateR["attribute"];
                                                    JArray items = (JArray)attribute["items"];

                                                    int indexitem = items.Count();
                                                    string[] Boxitem = new string[indexitem];
                                                    if (jData2 != null)
                                                    {
                                                        string itemvalue = (string)RowIndex.ItemArray[r];
                                                        string[] dataListArray = itemvalue.Split('|');

                                                        for (int F = 0; F < indexitem; F++)
                                                        {
                                                            string itemc = items[F]["item"].ToString();
                                                            if (dataListArray.Contains(itemc))
                                                            {
                                                                Boxitem[F] = "Y";
                                                            }
                                                            else
                                                            {
                                                                Boxitem[F] = "N";
                                                            }
                                                        }
                                                        string Valuecd = (string)RowIndex.ItemArray[r];
                                                        jData2["value"] = Valuecd.Replace('|', ',');
                                                        jData2["item"] = new JArray(Boxitem);
                                                        break;
                                                    }
                                                }
                                                else
                                                {
                                                    jData2["value"] = (string)RowIndex.ItemArray[r];
                                                    break;
                                                }
                                            }
                                        }
                                        #endregion
                                    }
                                }
                            }
                            #region Row
                            if (jData["row"] != null && jData["row"].Type == JTokenType.Null)
                            {
                                if ((String)jTemplateL["label"] == "รายละเอียดภาษี")
                                {
                                    string Taxtype1 = (string)RowIndex.ItemArray[72];
                                    string Taxcode1 = (string)RowIndex.ItemArray[73];
                                    string WHTCode1 = (string)RowIndex.ItemArray[74];
                                    string TypeofRep1 = (string)RowIndex.ItemArray[75];
                                    string Subjecttax1 = (string)RowIndex.ItemArray[76];

                                    string Taxtype2 = (string)RowIndex.ItemArray[77];
                                    string Taxcode2 = (string)RowIndex.ItemArray[78];
                                    string WHTCode2 = (string)RowIndex.ItemArray[79];
                                    string TypeofRep2 = (string)RowIndex.ItemArray[80];
                                    string Subjecttax2 = (string)RowIndex.ItemArray[81];

                                    string Taxtype3 = (string)RowIndex.ItemArray[82];
                                    string Taxcode3 = (string)RowIndex.ItemArray[83];
                                    string WHTCode3 = (string)RowIndex.ItemArray[84];
                                    string TypeofRep3 = (string)RowIndex.ItemArray[85];
                                    string Subjecttax3 = (string)RowIndex.ItemArray[86];

                                    string Taxtype4 = (string)RowIndex.ItemArray[87];
                                    string Taxcode4 = (string)RowIndex.ItemArray[88];
                                    string WHTCode4 = (string)RowIndex.ItemArray[89];
                                    string TypeofRep4 = (string)RowIndex.ItemArray[90];
                                    string Subjecttax4 = (string)RowIndex.ItemArray[91];
                                    string valuerow = $"[{{\"value\": \"WHT Type 1\"}},{{\"value\": \"{Taxtype1}\"}},{{\"value\": \"{Taxcode1}\"}},{{\"value\": \"{WHTCode1}\"}},{{\"value\": \"{TypeofRep1}\"}},{{\"value\": \"{Subjecttax1}\"}}]" +
                                        $",[{{\"value\": \"WHT Type 2\"}},{{\"value\": \"{Taxtype2}\"}},{{\"value\": \"{Taxcode2}\"}},{{\"value\": \"{WHTCode2}\"}},{{\"value\": \"{TypeofRep2}\"}},{{\"value\": \"{Subjecttax2}\"}}]" +
                                        $",[{{\"value\": \"WHT Type 3\"}},{{\"value\": \"{Taxtype3}\"}},{{\"value\": \"{Taxcode3}\"}},{{\"value\": \"{WHTCode3}\"}},{{\"value\": \"{TypeofRep3}\"}},{{\"value\": \"{Subjecttax3}\"}}]" +
                                        $",[{{\"value\": \"WHT Type 4\"}},{{\"value\": \"{Taxtype4}\"}},{{\"value\": \"{Taxcode4}\"}},{{\"value\": \"{WHTCode4}\"}},{{\"value\": \"{TypeofRep4}\"}},{{\"value\": \"{Subjecttax4}\"}}]";
                                    value = $"[{valuerow}]";

                                    jData.Remove("row");
                                    jData.Add("row", JArray.Parse(value));
                                }
                            }
                            #endregion
                        }
                    }
                }
                return JsonConvert.SerializeObject(jsonAdvanceForm);
            }
            catch (Exception ex)
            {
                LogError("ImportAbvanceForm: " + ex);
                LogError("ImportAbvanceForm: " + ex.StackTrace);
                return string.Empty;
            }
        }
        public static TRNMemo InsertTrnmemo(ContextDataContext db, string MAbvanceForm, DataRow Rowheader, DataRow RowIndex, List<MSTEmployee> ListViewEmployee)
        {
            try
            {
                string guids = Guid.NewGuid().ToString().Replace("-", "");
                Log("Start InsertTrnmemo");
                int sDoccode = 0;
                int sSubject = 0;
                int sCreator = 0;
                int sRequestor = 0;
                for (int r = 0; r < Rowheader.ItemArray.Length; r++)
                {
                    var label = Rowheader.ItemArray[r];
                    if (label.ToString().Contains("หมายเลขเอกสาร") || label.ToString().Contains("เลขที่เอกสาร") || label.ToString().Contains("รหัสเอกสาร") || label.ToString().Contains("รหัสผู้ค้า") || label.ToString().Contains("รหัสผู้ค้า :"))
                    {
                        sDoccode = r;
                    }
                    if (label.ToString().Contains("ชื่อเอกสาร") || label.ToString().Contains("ชื่อผู้ค้า :"))
                    {
                        sSubject = r;
                    }
                    if (label.ToString() == "Creator")
                    {
                        sCreator = r;
                    }
                    if (label.ToString() == "Requestor")
                    {
                        sRequestor = r;
                    }
                }
                MSTEmployee EmpCreator = new MSTEmployee();
                string Creator = (string)RowIndex.ItemArray[sCreator];
                EmpCreator = ListViewEmployee.Where(x => x.Email.ToUpper().Contains(Creator.ToUpper())).FirstOrDefault();
                if (EmpCreator == null)
                {
                    EmpCreator = ListViewEmployee.Where(x => x.Email.Contains(ItemConfig.Admin)).FirstOrDefault();
                }

                MSTEmployee EmpRequester = new MSTEmployee();
                string Requester = (string)RowIndex.ItemArray[sRequestor];
                EmpRequester = ListViewEmployee.Where(x => x.Email.ToUpper().Contains(Requester.ToUpper())).FirstOrDefault();
                if (EmpRequester == null)
                {
                    EmpRequester = ListViewEmployee.Where(x => x.Email.Contains(ItemConfig.Admin)).FirstOrDefault();
                }

                MSTEmployee EmpWaitFor = new MSTEmployee();
                string WaitFor = ItemConfig.WaitForApprove;
                if (!string.IsNullOrEmpty(WaitFor))
                {
                    EmpWaitFor = ListViewEmployee.Where(x => x.Email.ToUpper().Contains(WaitFor.ToUpper())).FirstOrDefault();
                    if (EmpWaitFor == null)
                    {
                        EmpWaitFor = ListViewEmployee.Where(x => x.Email.Contains(ItemConfig.Admin)).FirstOrDefault();
                    }
                }

                var DestinationTemplate = db.MSTTemplates.Where(a => a.TemplateId == ItemConfig.TemplateID).ToList();
                var lstCompany = db.MSTCompanies.ToList();

                TRNMemo objMemo = new TRNMemo();
                objMemo.StatusName = ItemConfig.Status;
                if (objMemo.StatusName.ToLower() == "draft")
                {
                    objMemo.PersonWaitingId = EmpCreator.EmployeeId;
                    objMemo.PersonWaiting = EmpCreator.NameTh;
                }
                if (objMemo.StatusName.ToLower() == "wait for approve")
                {
                    objMemo.PersonWaitingId = EmpWaitFor.EmployeeId;
                    objMemo.PersonWaiting = EmpWaitFor.NameTh;
                    objMemo.CurrentApprovalLevel = 3;
                }
                else
                {
                    objMemo.PersonWaitingId = null;
                    objMemo.PersonWaiting = "";
                }
                objMemo.CreatedDate = DateTime.Now;
                objMemo.CreatedBy = EmpCreator.NameEn;
                objMemo.CreatorId = EmpCreator.EmployeeId;
                objMemo.CNameTh = EmpCreator.NameTh;
                objMemo.CNameEn = EmpCreator.NameEn;
                objMemo.CPositionId = EmpCreator.PositionId;
                if (objMemo.CPositionId != null)
                {
                    objMemo.CPositionTh = db.MSTPositions.Where(x => x.PositionId == EmpCreator.PositionId).Select(s => s.NameTh).FirstOrDefault();
                    objMemo.CPositionEn = db.MSTPositions.Where(x => x.PositionId == EmpCreator.PositionId).Select(s => s.NameEn).FirstOrDefault();
                }
                objMemo.CDepartmentId = EmpCreator.DepartmentId;
                if (objMemo.CDepartmentId != null)
                {
                    objMemo.CDepartmentTh = db.MSTDepartments.Where(x => x.DepartmentId == EmpCreator.DepartmentId).Select(s => s.NameTh).FirstOrDefault();
                    objMemo.CDepartmentEn = db.MSTDepartments.Where(x => x.DepartmentId == EmpCreator.DepartmentId).Select(s => s.NameEn).FirstOrDefault();
                }
                objMemo.RequesterId = EmpRequester.EmployeeId;
                objMemo.RNameTh = EmpRequester.NameTh;
                objMemo.RNameEn = EmpRequester.NameEn;
                objMemo.RPositionId = EmpRequester.PositionId;
                if (objMemo.RPositionId != null)
                {
                    objMemo.RPositionTh = db.MSTPositions.Where(x => x.PositionId == EmpRequester.PositionId).Select(s => s.NameTh).FirstOrDefault();
                    objMemo.RPositionEn = db.MSTPositions.Where(x => x.PositionId == EmpRequester.PositionId).Select(s => s.NameEn).FirstOrDefault();
                }
                objMemo.RDepartmentId = EmpRequester.DepartmentId;
                if (objMemo.RDepartmentId != null)
                {
                    objMemo.RDepartmentTh = db.MSTDepartments.Where(x => x.DepartmentId == EmpRequester.DepartmentId).Select(s => s.NameTh).FirstOrDefault();
                    objMemo.RDepartmentEn = db.MSTDepartments.Where(x => x.DepartmentId == EmpRequester.DepartmentId).Select(s => s.NameEn).FirstOrDefault();
                }

                objMemo.ModifiedDate = DateTime.Now;
                objMemo.ModifiedBy = objMemo.ModifiedBy;
                objMemo.TemplateId = DestinationTemplate.First().TemplateId;
                objMemo.TemplateName = DestinationTemplate.First().TemplateName;
                objMemo.GroupTemplateName = DestinationTemplate.First().GroupTemplateName;
                objMemo.RequestDate = DateTime.Now;
                var CurrentCom = lstCompany.Find(a => a.CompanyId == ItemConfig.company);
                objMemo.CompanyId = CurrentCom.CompanyId;
                objMemo.CompanyName = CurrentCom.NameTh;

                objMemo.MAdvancveForm = MAbvanceForm;
                objMemo.TAdvanceForm = MAbvanceForm;
                objMemo.MemoSubject = "ขอแก้ไขข้อมูลผู้ค้า" + (string)RowIndex.ItemArray[sDoccode] + " : " + (string)RowIndex.ItemArray[sSubject];
                objMemo.TemplateSubject = "ขอแก้ไขข้อมูลผู้ค้า" + (string)RowIndex.ItemArray[sDoccode] + " : " + (string)RowIndex.ItemArray[sSubject];
                objMemo.TemplateDetail = guids;
                objMemo.ProjectID = 0;
                objMemo.DocumentCode = GenControlRunning(EmpCreator, DestinationTemplate.First().DocumentCode, objMemo, db);
                objMemo.DocumentNo = objMemo.DocumentCode;
                db.TRNMemos.InsertOnSubmit(objMemo);
                db.SubmitChanges();
                Console.WriteLine("GenerateTrnMemo success : Memoid >> " + objMemo.MemoId);
                Log("GenerateTrnMemo success : Memoid >> " + objMemo.MemoId);
                Log("End InsertTRNMemo");

                return objMemo;
            }
            catch (Exception ex)
            {
                LogError("InsertTrnmemo: " + ex);
                return null;
            }
        }
        public static void InsertTrnLineapprove(ContextDataContext db, List<MSTEmployee> ListViewEmployee, TRNMemo memo)
        {
            try
            {
                #region InsertTrnLineapprove
                Log("Start InsertTrnLineapprove");
                int sequence = 1;
                List<ViewEmployee> lstemp = new List<ViewEmployee>();
                List<TRNLineApprove> lstapprove = new List<TRNLineApprove>();
                //var emp1 = db.ViewEmployees.Where(e => e.Email == "supportpurchase@thailife.com").FirstOrDefault();
                //lstemp.Add(emp1);
                var emp2 = db.ViewEmployees.Where(e => e.Email == "natee.ler@thailife.com").FirstOrDefault();
                lstemp.Add(emp2);
                var emp3 = db.ViewEmployees.Where(e => e.Email == "pensri@thailife.com").FirstOrDefault();
                lstemp.Add(emp3);
                var emp4 = db.ViewEmployees.Where(e => e.Email == "supportpurchase@thailife.com").FirstOrDefault();
                lstemp.Add(emp4);
                foreach (var item in lstemp)
                {
                    TRNLineApprove approve = new TRNLineApprove();
                    approve.LineApproveId = 0;
                    approve.MemoId = memo.MemoId;
                    approve.Seq = sequence;
                    approve.EmployeeId = item.EmployeeId;
                    approve.EmployeeCode = item.EmployeeCode;
                    approve.NameTh = item.NameTh;
                    approve.NameEn = item.NameEn;
                    approve.PositionTH = item.PositionNameTh;
                    approve.PositionEN = item.PositionNameEn;
                    //if (sequence == 1)
                    //{
                    //    approve.SignatureId = 2195;
                    //    approve.SignatureTh = "ดำเนินการ";
                    //    approve.SignatureEn = "ดำเนินการ";
                    //}
                    if (sequence == 1)
                    {
                        approve.SignatureId = 2019;
                        approve.SignatureTh = "อนุมัติ";
                        approve.SignatureEn = "อนุมัติ";
                    }
                    else if (sequence == 2)
                    {
                        approve.SignatureId = 2019;
                        approve.SignatureTh = "อนุมัติ";
                        approve.SignatureEn = "อนุมัติ";
                    }
                    else if (sequence == 3)
                    {
                        approve.SignatureId = 2195;
                        approve.SignatureTh = "ดำเนินการ";
                        approve.SignatureEn = "ดำเนินการ";
                    }
                    approve.IsActive = true;
                    lstapprove.Add(approve);
                    db.TRNLineApproves.InsertOnSubmit(approve);
                    db.SubmitChanges();
                    sequence++;
                }
                Log("End InsertTrnLineapprove");
                #endregion
                #region InsertTrnTRNActionHistory
                Log("Start InsertTrnTRNActionHistory");
                if (ItemConfig.Status.ToLower() == "completed")
                {
                    int sequence2 = 1;
                    foreach (var item in lstapprove)
                    {
                        var emp = db.ViewEmployees.Where(e => e.EmployeeId == item.EmployeeId).FirstOrDefault();
                        TRNActionHistory actionHistory = new TRNActionHistory();
                        actionHistory.MemoId = memo.MemoId;
                        actionHistory.ActorName = emp.NameEn;
                        actionHistory.StartDate = DateTime.Now;
                        actionHistory.ActionProcess = "approve";
                        actionHistory.ActionDate = DateTime.Now;
                        actionHistory.ActionStatus = "Wait for Approve";
                        actionHistory.SignatureId = item.SignatureId;
                        actionHistory.Platform = "web";
                        actionHistory.ActorNameTh = emp.NameTh;
                        actionHistory.ActorNameEn = emp.NameEn;
                        actionHistory.ActorPositionId = emp.PositionId;
                        actionHistory.ActorPositionNameTh = emp.PositionNameTh;
                        actionHistory.ActorPositionNameEn = emp.PositionNameEn;
                        actionHistory.ActorDepartmentId = emp.DepartmentId;
                        actionHistory.ActorDepartmentNameTh = emp.DepartmentNameTh;
                        actionHistory.ActorDepartmentNameEn = emp.DepartmentNameEn;
                        actionHistory.HAdvancveForm = memo.MAdvancveForm;
                        actionHistory.LineApproveSeq = sequence2;
                        db.TRNActionHistories.InsertOnSubmit(actionHistory);
                        db.SubmitChanges();
                        sequence2++;
                    }
                }
                else if (ItemConfig.Status.ToLower() == "wait for approve")
                {
                    int sequence2 = 1;
                    foreach (var item in lstapprove)
                    {
                        if (sequence2 <= 2)
                        {
                            var emp = db.ViewEmployees.Where(e => e.EmployeeId == item.EmployeeId).FirstOrDefault();
                            TRNActionHistory actionHistory = new TRNActionHistory();
                            actionHistory.MemoId = memo.MemoId;
                            actionHistory.ActorName = emp.NameEn;
                            actionHistory.StartDate = DateTime.Now;
                            actionHistory.ActionProcess = "approve";
                            actionHistory.ActionDate = DateTime.Now;
                            actionHistory.ActionStatus = "Wait for Approve";
                            actionHistory.SignatureId = item.SignatureId;
                            actionHistory.Platform = "web";
                            actionHistory.ActorNameTh = emp.NameTh;
                            actionHistory.ActorNameEn = emp.NameEn;
                            actionHistory.ActorPositionId = emp.PositionId;
                            actionHistory.ActorPositionNameTh = emp.PositionNameTh;
                            actionHistory.ActorPositionNameEn = emp.PositionNameEn;
                            actionHistory.ActorDepartmentId = emp.DepartmentId;
                            actionHistory.ActorDepartmentNameTh = emp.DepartmentNameTh;
                            actionHistory.ActorDepartmentNameEn = emp.DepartmentNameEn;
                            actionHistory.HAdvancveForm = memo.MAdvancveForm;
                            actionHistory.LineApproveSeq = sequence2;
                            db.TRNActionHistories.InsertOnSubmit(actionHistory);
                            db.SubmitChanges();
                            sequence2++;
                        }
                    }
                }
                Log("End InsertTrnTRNActionHistory");
                #endregion
            }
            catch (Exception ex)
            {
                LogError("InsertTrnLineapprove: " + ex);
                LogError("InsertTrnLineapprove: " + ex.StackTrace);
            }
        }
        public static void InsertTRNAttachFile(ContextDataContext db, TRNMemo memo, string running)
        {
            try
            {
                Log("Start InsertTRNAttachFile");
                string docCode = "000"+running; // ชื่อโฟลเดอร์ Doccode
                if (!string.IsNullOrEmpty(docCode))
                {
                    string rootFolder = ItemConfig.SourcePath; // โฟลเดอร์หลักที่มีโฟลเดอร์ Doccode
                    string targetBasePath = ConfigurationSettings.AppSettings["TargetPath"].ToString();

                    string sourceFolder = Path.Combine(rootFolder, docCode);

                    // ตรวจสอบว่าโฟลเดอร์ Doccode มีอยู่หรือไม่
                    if (!Directory.Exists(sourceFolder))
                    {
                        LogError($"Folder for Doccode {docCode} not found.");
                        return;
                    }
                    string targetFolder = Path.Combine(targetBasePath, memo.TemplateDetail);
                    if (!Directory.Exists(targetFolder))
                    {
                        Directory.CreateDirectory(targetFolder);
                    }

                    // ดึงไฟล์ทั้งหมดในโฟลเดอร์ Doccode
                    string[] files = Directory.GetFiles(sourceFolder);

                    // คัดลอกไฟล์ทั้งหมดไปยังโฟลเดอร์ปลายทาง
                    foreach (string file in files)
                    {
                        try
                        {
                            string destFileName = Path.GetFileName(file).Replace(" ", "-");
                            string destFilePath = Path.Combine(targetFolder, destFileName);

                            File.Copy(file, destFilePath, true);
                            Log($"Download success: {file} to {destFilePath} MemoId: {memo.MemoId}");
                            Console.WriteLine($"Download success: {file} MemoId: {memo.MemoId}");

                            TRNAttachFile attach = new TRNAttachFile
                            {
                                FileName = destFileName,
                                AttachedDate = DateTime.Now,
                                AttachFile = Path.GetFileName(file),
                                FilePath = $"/TempAttachment/{memo.TemplateDetail}/{destFileName}",
                                IsMergePDF = true,
                                ActorId = memo.CreatorId,
                                MemoId = memo.MemoId,
                                ActorName = memo.CNameEn
                            };

                            db.TRNAttachFiles.InsertOnSubmit(attach);
                            db.SubmitChanges();
                        }
                        catch (Exception ex)
                        {
                            LogError($"Error copying file {file}: {ex.Message}");
                        }
                    }
                }
                Log("End InsertTRNAttachFile");
            }
            catch (Exception ex)
            {
                LogError($"Error in InsertTRNAttachFile: {ex.StackTrace}");
            }
        }
        public static string InsertTRNControlRunning(ContextDataContext db, TRNMemo memo, DataRow Rowheader, DataRow RowIndex)
        {
            try
            {
                TRNControlRunning objControlRunning = new TRNControlRunning();
                Log("Start InsertTRNControlRunning");
                int ControlRunning = 0;
                for (int r = 0; r < Rowheader.ItemArray.Length; r++)
                {
                    var label = Rowheader.ItemArray[r];
                    if (label.ToString().Contains("หมายเลขเอกสาร") || label.ToString().Contains("เลขที่เอกสาร") || label.ToString().Contains("รหัสเอกสาร") || label.ToString().Contains("รหัสผู้ค้า"))
                    {
                        ControlRunning = r;
                    }
                }
                string Number = (string)RowIndex.ItemArray[ControlRunning];
                if (string.IsNullOrEmpty(Number))
                {
                    var lastNumber = db.TRNControlRunnings.OrderBy(o => o.Running).Last();
                    objControlRunning.TemplateId = ItemConfig.TemplateID;
                    objControlRunning.Prefix = "1";
                    objControlRunning.Digit = ItemConfig.DigitCon;
                    objControlRunning.CreateBy = memo.CreatorId.ToString();
                    objControlRunning.CreateDate = DateTime.Now;
                    objControlRunning.RunningNumber = lastNumber.RunningNumber + 1;
                    objControlRunning.Running = Convert.ToInt32(objControlRunning.RunningNumber.ToString().Substring(objControlRunning.RunningNumber.ToString().Length - 4));
                    db.TRNControlRunnings.InsertOnSubmit(objControlRunning);
                    db.SubmitChanges();
                }
                else
                {
                    objControlRunning.TemplateId = ItemConfig.TemplateID;
                    objControlRunning.Prefix = "1";
                    objControlRunning.Digit = ItemConfig.DigitCon;
                    objControlRunning.CreateBy = memo.CreatorId.ToString();
                    objControlRunning.CreateDate = DateTime.Now;
                    objControlRunning.RunningNumber = Number;
                    objControlRunning.Running = Convert.ToInt32(objControlRunning.RunningNumber.ToString().Substring(objControlRunning.RunningNumber.ToString().Length - 4));
                    db.TRNControlRunnings.InsertOnSubmit(objControlRunning);
                    db.SubmitChanges();
                }
                Log("End InsertTRNControlRunning");
                return Number;
            }
            catch (Exception ex)
            {
                return null;
                LogError("InsertTRNControlRunning: " + ex.StackTrace);
            }
        }
        public static string GenControlRunning(MSTEmployee Emp, string DocumentCode, TRNMemo objTRNMemo, ContextDataContext db)
        {
            string TempCode = DocumentCode;
            String sPrefixDocNo = $"{TempCode}-{DateTime.Now.Year.ToString()}-";
            int iRunning = 1;
            List<TRNMemo> temp = db.TRNMemos.Where(a => a.DocumentNo.ToUpper().Contains(sPrefixDocNo.ToUpper())).ToList();
            if (temp.Count > 0)
            {
                String sLastDocumentNo = temp.OrderBy(a => a.DocumentNo).Last().DocumentNo;
                if (!String.IsNullOrEmpty(sLastDocumentNo))
                {
                    List<String> list_LastDocumentNo = sLastDocumentNo.Split('-').ToList();

                    if (list_LastDocumentNo.Count >= 3)
                    {
                        iRunning = checkDataIntIsNull(list_LastDocumentNo[list_LastDocumentNo.Count - 1]) + 1;
                    }
                }
            }

            String sDocumentNo = $"{sPrefixDocNo}{iRunning.ToString().PadLeft(6, '0')}";


            try
            {

                var mstMasterDataList = db.MSTMasterDatas.Where(a => a.MasterType == "DocNo").ToList();

                if (mstMasterDataList != null)
                    if (mstMasterDataList.Count() > 0)
                    {
                        var getCompany = db.MSTCompanies.Where(a => a.CompanyId == objTRNMemo.CompanyId).ToList();
                        var getDepartment = db.MSTDepartments.Where(a => a.DepartmentId == Emp.DepartmentId).ToList();
                        var getDivision = db.MSTDivisions.Where(a => a.DivisionId == Emp.DivisionId).ToList();

                        string CompanyCode = "";
                        string DepartmentCode = "";
                        string DivisionCode = "";
                        if (getCompany != null)
                            if (!string.IsNullOrWhiteSpace(getCompany.First().CompanyCode)) CompanyCode = getCompany.First().CompanyCode;
                        if (DepartmentCode != null)
                            if (!string.IsNullOrWhiteSpace(getDepartment.First().DepartmentCode)) DepartmentCode = getDepartment.First().DepartmentCode;
                        if (DivisionCode != null)
                        {
                            if (getDivision.Count > 0)
                                if (!string.IsNullOrWhiteSpace(getDivision.First().DivisionCode)) DivisionCode = getDivision.First().DivisionCode;
                        }
                        foreach (var getMaster in mstMasterDataList)
                        {
                            if (!string.IsNullOrWhiteSpace(getMaster.Value2))
                            {
                                var Tid_array = getMaster.Value2.Split('|');
                                string FixDoc = getMaster.Value1;
                                if (Tid_array.Count() > 0)
                                {
                                    if (Tid_array.Contains(objTRNMemo.TemplateId.ToString()))
                                    {
                                        sDocumentNo = DocNoGenerate(FixDoc, TempCode, CompanyCode, DepartmentCode, DivisionCode, db);
                                    }
                                }
                            }
                            else
                            {
                                string FixDoc = getMaster.Value1;
                                sDocumentNo = DocNoGenerate(FixDoc, TempCode, CompanyCode, DepartmentCode, DivisionCode, db);
                            }
                        }

                    }
            }
            catch (Exception ex) { }
            return sDocumentNo;
        }
        public static int checkDataIntIsNull(object Input)
        {
            int Results = 0;
            if (Input != null)
                int.TryParse(Input.ToString().Replace(",", ""), out Results);

            return Results;
        }
        public static string DocNoGenerate(string FixDoc, string DocCode, string CCode, string DCode, string DSCode, ContextDataContext db)
        {
            string sDocumentNo = "";
            int iRunning;
            if (!string.IsNullOrWhiteSpace(FixDoc))
            {
                string y4 = DateTime.Now.ToString("yyyy");
                string y2 = DateTime.Now.ToString("yy");
                string CompanyCode = CCode;
                string DepartmentCode = DCode;
                string DivisionCode = DSCode;
                string FixCode = FixDoc;
                FixCode = FixCode.Replace("[CompanyCode]", CompanyCode);
                FixCode = FixCode.Replace("[DepartmentCode]", DepartmentCode);
                FixCode = FixCode.Replace("[DocumentCode]", DocCode);
                FixCode = FixCode.Replace("[DivisionCode]", DivisionCode);

                FixCode = FixCode.Replace("[YYYY]", y4);
                FixCode = FixCode.Replace("[YY]", y2);
                sDocumentNo = FixCode;
                List<TRNMemo> tempfixDoc = db.TRNMemos.Where(a => a.DocumentNo.ToUpper().Contains(sDocumentNo.ToUpper())).ToList();


                List<TRNMemo> tempfixDocByYear = db.TRNMemos.ToList();

                tempfixDocByYear = tempfixDocByYear.FindAll(a => a.DocumentNo != ("Auto Generate") & Convert.ToDateTime(a.RequestDate).Year.ToString().Equals(y4)).ToList();
                if (tempfixDocByYear.Count > 0)
                {
                    tempfixDocByYear = tempfixDocByYear.OrderByDescending(a => a.MemoId).ToList();

                    String sLastDocumentNofix = tempfixDocByYear.First().DocumentNo;
                    if (!String.IsNullOrEmpty(sLastDocumentNofix))
                    {
                        List<String> list_LastDocumentNofix = sLastDocumentNofix.Split('-').ToList();
                        //Arm Edit 2020-05-18 Bug If Prefix have '-' will no running because list_LastDocumentNo.Count > 3

                        if (list_LastDocumentNofix.Count >= 3)
                        {
                            iRunning = checkDataIntIsNull(list_LastDocumentNofix[list_LastDocumentNofix.Count - 1]) + 1;
                            sDocumentNo = $"{sDocumentNo}-{iRunning.ToString().PadLeft(6, '0')}";
                        }
                    }
                }
                else
                {
                    sDocumentNo = $"{sDocumentNo}-{1.ToString().PadLeft(6, '0')}";

                }
            }
            return sDocumentNo;

        }
        public static void InsertTRNReferenceDoc(DataRow RowIndex, ContextDataContext db, TRNMemo memo)
        {
            string memoid = (string)RowIndex.ItemArray[98];
            var memoref = db.TRNMemos.Where(x => x.MemoId == Convert.ToInt32(memoid)).FirstOrDefault();
            TRNReferenceDoc mmmref = new TRNReferenceDoc();
            mmmref.MemoID = memo.MemoId;
            mmmref.MemoRefDocID = memoref.MemoId;
            mmmref.DocumentNo = memo.DocumentNo;
            mmmref.TemplateId = memo.TemplateId;
            mmmref.TemplateName = memo.TemplateName;
            mmmref.MemoSubject = memo.MemoSubject;
            mmmref.CreatedBy = "Adminchris";
            mmmref.CreatedDate = DateTime.Now;
            db.TRNReferenceDocs.InsertOnSubmit(mmmref);
            db.SubmitChanges();
        }
    }
}
