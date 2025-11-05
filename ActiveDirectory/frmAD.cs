//using Microsoft.Exchange.WebServices;
//using Microsoft.Exchange.WebServices.Data;
//using Microsoft.Exchange.Data.Directory;
//using Microsoft.Exchange.Data.Directory.Management;
//using Microsoft.Exchange.Data.Directory.Recipient;
//using Microsoft.Exchange.Data.Directory.SystemConfiguration;

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.DirectoryServices.ActiveDirectory;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;

using System.Security.Principal;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Management;


namespace ActiveDirectory
{
    public partial class frmAD : Form
    {
      
        private SearchResultCollection rsc;
        private Dictionary<string, SearchResult> resdic = new Dictionary<string, SearchResult>();
        private BackgroundWorker m_BGW_ExcelExport = new BackgroundWorker();

        public frmAD()
        {
            InitializeComponent();

            m_BGW_ExcelExport.WorkerReportsProgress = true;
            m_BGW_ExcelExport.WorkerSupportsCancellation = true;
            m_BGW_ExcelExport.DoWork += m_BGW_ExcelExport_DoWork;
            m_BGW_ExcelExport.ProgressChanged += m_BGW_ExcelExport_ProgressChanged;
            m_BGW_ExcelExport.RunWorkerCompleted += m_BGW_ExcelExport_RunWorkerCompleted;
        }

        void m_BGW_ExcelExport_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Error != null)
            {
                MessageBox.Show(e.Error.Message);
            }
            else if (e.Cancelled)
            {
                MessageBox.Show("Отменено!");
            }
            else
            {
                MessageBox.Show("Экспорт завершен!");
            }
        }

        void m_BGW_ExcelExport_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            //throw new NotImplementedException();
        }

        void m_BGW_ExcelExport_DoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker BGW = sender as BackgroundWorker;

            ExportUsersToExcelArgument Arg = e.Argument as ExportUsersToExcelArgument;

            if (Arg == null)
                return;

            if (Arg.ADSR == null || string.IsNullOrWhiteSpace(Arg.sExcelWBFileName))
                return;

            ExportToExcel(Arg.sExcelWBFileName, Arg.ADSR);
        }



        private void frmAD_Load(object sender, EventArgs e)
        {
            txtUsername.Text = Environment.UserName;
            txtPassword.Focus();

            btnSearchUserName.Select();

            txtAddress.Text = GetSystemDomain();            
        }

        private string GetSystemDomain()
        {
            try
            {
                return Domain.GetComputerDomain().ToString().ToLower();
            }
            catch (Exception e)
            {
                e.Message.ToString();
                return string.Empty;
            }
        }
         
        private void GetUserInformation(string username, string passowrd, string domain)
        {
                     Cursor.Current = Cursors.WaitCursor;
            SearchResultCollection rs = null;
       
               if (rbUseremail.Checked)
               {
                   
                   if (txtSearchUser.Text.Trim().IndexOf("@") > 0)
                   {
                       rs = SearchUserByEmail(GetDirectorySearcher(username, passowrd, domain), txtSearchUser.Text.Trim() + "*");
                   }
                   else
                   {
                       rs = SearchUserByUserName(GetDirectorySearcher(username, passowrd, domain), txtSearchUser.Text.Trim() + "*");
                   }

               }
            else if (rbFirstname.Checked)
               {
               
                rs = SearchUserByFirstName(GetDirectorySearcher(username, passowrd, domain), txtSearchUser.Text.Trim() + "*");
               }
            
            else if (rbPhoneNo.Checked)
               
                   rs = SearchUserByPhoneNo(GetDirectorySearcher(username, passowrd, domain), txtSearchUser.Text.Trim() + "*");
              
           if (rs.Count != 0)
            {
                ShowUserInformation();
            }
            else
            {
                MessageBox.Show("Пользователь не найден!", "Информация по поиску", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void ClearUserDetales()
        {
    
                lblUsernameDisplay.Text = "Данные отсутствуют";
                lblFirstname.Text = "ФИО :";
                lblTitle.Text = "Должность :";
                lblCompany.Text = "Компания :";
                lbOfficeName.Text = "Помещение :";
                lbMobile.Text = "Мобильный телефон :";
                lbDepartment.Text = "Отдел :";
                lblCity.Text = "Город :";
                lblTelephone.Text = "Внутренний номер :";
                lblDescription.Text = "Описание :";
            
                lbCreated.Text = "Запись создана :";
                lbManager.Text = "Руководитель :";
                lblastLogon.Text = "Входа не было";
                lbpwdLastSet.Text = "Пароль не менялся";
                
            NoteTextBox1.Clear();
                tbcaconicalname.Clear();
                pbUserImg.Image = null;
                listView2.Items.Clear();
                listView3.Items.Clear();
                listView4.Items.Clear();
                listView5.Items.Clear();
                chbSmartRequired.Checked = false;
                chbPasswdDontExpire.Checked = false;
                chbPassCantchange.Checked = false;
                chbAccountDisabled.Checked = false;

        }

        private void ShowUserDetales(string _SID)
        {
            string dAccExp = "01.01.1970 0:00:00";

            if (!resdic.ContainsKey(_SID))
                return;
            SearchResult rs = resdic[_SID];
    
            if (rs.GetDirectoryEntry().Properties["samaccountname"].Value != null)
                lblUsernameDisplay.Text = "Имя пользователя : " + rs.GetDirectoryEntry().Properties["samaccountname"].Value.ToString();

            if (rs.GetDirectoryEntry().Properties["displayName"].Value != null)
                lblFirstname.Text = "ФИО : " + rs.GetDirectoryEntry().Properties["displayName"].Value.ToString();

            if (rs.GetDirectoryEntry().Properties["title"].Value != null)
                lblTitle.Text = "Должность : " + rs.GetDirectoryEntry().Properties["title"].Value.ToString();

            if (rs.GetDirectoryEntry().Properties["company"].Value != null)
                lblCompany.Text = "Компания : " + rs.GetDirectoryEntry().Properties["company"].Value.ToString();

            if (rs.GetDirectoryEntry().Properties["physicalDeliveryOfficeName"].Value != null)
                lbOfficeName.Text = "Помещение : " + rs.GetDirectoryEntry().Properties["physicalDeliveryOfficeName"].Value.ToString();
 
            if (rs.GetDirectoryEntry().Properties["mobile"].Value != null)
                lbMobile.Text = "Мобильный телефон : " + rs.GetDirectoryEntry().Properties["mobile"].Value.ToString();


            if (rs.GetDirectoryEntry().InvokeGet("PasswordExpirationDate").ToString() != dAccExp) 
            {
                DateTime dt = Convert.ToDateTime(rs.GetDirectoryEntry().InvokeGet("PasswordExpirationDate"));
                if (dt <= DateTime.Now)
                {
                    lbDateAccounExp.ForeColor = Color.Red;
                    lbDateAccounExp.Font = new Font(lbDateAccounExp.Font,FontStyle.Bold);
                    lbDateAccounExp.Text = "Учетная запись отключена: " +dt.ToString("dd.MM.yyyy");
                }
                else
                {
                    lbDateAccounExp.ForeColor = Color.Black;
                    lbDateAccounExp.Font = new Font(lbDateAccounExp.Font, FontStyle.Regular);
                    lbDateAccounExp.Text = "Срок действия пароля: " + dt.ToString("dd.MM.yyyy");
                }                  
            }

            else 
            {
                lbDateAccounExp.ForeColor = Color.Black;
                lbDateAccounExp.Font = new Font(lbDateAccounExp.Font, FontStyle.Regular);
                lbDateAccounExp.Text = "Срок действия пароля не ограничен";
            }
                
            if (rs.GetDirectoryEntry().Properties["department"].Value != null)
                lbDepartment.Text = "Отдел : " + rs.GetDirectoryEntry().Properties["department"].Value.ToString();
            
            if (rs.GetDirectoryEntry().Properties["l"].Value != null)
                lblCity.Text = "Город : " + rs.GetDirectoryEntry().Properties["l"].Value.ToString();
                        
            if (rs.GetDirectoryEntry().Properties["telephoneNumber"].Value != null)
                lblTelephone.Text = "Внутренний номер : " + rs.GetDirectoryEntry().Properties["telephoneNumber"].Value.ToString();

            if (rs.GetDirectoryEntry().Properties["description"].Value != null)
                lblDescription.Text = "Описание : " + rs.GetDirectoryEntry().Properties["description"].Value.ToString();
            
            if (rs.GetDirectoryEntry().Properties["info"].Value != null)
            NoteTextBox1.Text = rs.GetDirectoryEntry().Properties["info"].Value.ToString();

            if (rs.GetDirectoryEntry().Properties["whenCreated"].Value != null)
                lbCreated.Text = "Запись создана : " + rs.GetDirectoryEntry().Properties["whenCreated"].Value.ToString();

            if (rs.GetDirectoryEntry().Properties["thumbnailPhoto"].Value != null)
            {
                pbUserImg.Image = GetUserPicture(rs.GetDirectoryEntry().Properties["thumbnailPhoto"].Value);
            }

            if (rs.GetDirectoryEntry().Properties["distinguishedName"].Value != null)
            tbcaconicalname.Text = rs.GetDirectoryEntry().Properties["distinguishedName"].Value.ToString().Replace("CN=","").Replace("OU=","").Replace("DC=",".").Replace(",","\\");

            if (rs.GetDirectoryEntry().Properties["accountExpires"].Value != null)
            {
                object acexp = rs.GetDirectoryEntry().Properties["accountExpires"].Value;

                long expires = LongFromLargeInteger(acexp);

                if (expires == long.MaxValue || expires <= 0 || DateTime.MaxValue.ToFileTime() <= expires)
                {
                    lbAccuntExperi.ForeColor = Color.Black;
                    lbAccuntExperi.Font = new Font(lbDateAccounExp.Font, FontStyle.Regular);
                    lbAccuntExperi.Text = "Дата отключения записи не установлена";
                }
                else
                {
                    lbAccuntExperi.ForeColor = Color.Red;
                    lbAccuntExperi.Font = new Font(lbDateAccounExp.Font, FontStyle.Bold);
                    lbAccuntExperi.Text = "Дата отключения УЗ: " + DateTime.FromFileTime(expires).ToString("dd.MM.yyyy");
                }
                
            }

            if (rs.GetDirectoryEntry().Properties["lastLogonTimestamp"].Value != null)
            {
                object lastlogon = rs.GetDirectoryEntry().Properties["lastLogonTimestamp"].Value;

                long expires = LongFromLargeInteger(lastlogon);

                if (expires == long.MaxValue || expires <= 0 || DateTime.MaxValue.ToFileTime() <= expires)
                {
                    lblastLogon.Text = "Входа не было";
                }
                else
                {
                    lblastLogon.Text = "Дата последнего входа: " + DateTime.FromFileTime(expires).ToString("dd.MM.yyyy");
                }
                
            }
            if (rs.GetDirectoryEntry().Properties["lastLogon"].Value != null)
            {
                object lastlogon = rs.GetDirectoryEntry().Properties["lastLogon"].Value;

                long expires = LongFromLargeInteger(lastlogon);

                if (expires == long.MaxValue || expires <= 0 || DateTime.MaxValue.ToFileTime() <= expires)
                {
                    lblastLogon1.Text = "Авторизации не было";
                }
                else
                {
                    lblastLogon1.Text = "Дата последней авторизации: " + DateTime.FromFileTime(expires).ToString("dd.MM.yyyy");
                }
                
            }
            

            if (rs.GetDirectoryEntry().Properties["pwdLastSet"].Value != null)
            {
                object pwdLastSet = rs.GetDirectoryEntry().Properties["pwdLastSet"].Value;

                long expires = LongFromLargeInteger(pwdLastSet);

                if (expires == long.MaxValue || expires <= 0 || DateTime.MaxValue.ToFileTime() <= expires)
                {
                    lbpwdLastSet.Text = "Пароль не менялся";
                }
                else
                {
                    lbpwdLastSet.Text = "Дата последнего изменения пароля: " + DateTime.FromFileTime(expires).ToString("dd.MM.yyyy");
                }

            }
           
            if (rs.GetDirectoryEntry().Properties["manager"].Value != null)
            {
                string usrManager = null;
                usrManager = rs.GetDirectoryEntry().Properties["manager"].Value.ToString();

                int managerstart = usrManager.IndexOf("CN=");
                int managerend = usrManager.IndexOf(",", managerstart);
                usrManager = usrManager.Substring(managerstart + 3, managerend - managerstart - 3);

                lbManager.Text = "Руководитель : " + usrManager;
            }


            ShowX509Detalise(rs);
            ShowEmailAddress(rs);
            ShowMemberOf(rs);
            ShowDirectReport(rs);
            ShowAccountOptions(rs);

        }

        private long LongFromLargeInteger(object largeInteger)
        {
            System.Type type = largeInteger.GetType();
            int highPart = (int)type.InvokeMember("HighPart", BindingFlags.GetProperty, null, largeInteger, null);
            int lowPart = (int)type.InvokeMember("LowPart", BindingFlags.GetProperty, null, largeInteger, null);

            return (long)highPart << 32 | (uint)lowPart;
        }

        private void ExportToExcel(string _sFilePath, Dictionary<string, SearchResult> _SRDic)
        {
            Dictionary<string, Dictionary<string, List<UserExportInfo>>> OrgDepUserInfo = new Dictionary<string, Dictionary<string, List<UserExportInfo>>>();
            Excel.Application ex_app = null;
            Excel.Workbook ex_wb = null;

            try
            {
                GetExportInfo(OrgDepUserInfo, _SRDic,true);

                ex_app = new Excel.Application { DisplayAlerts = false };
                ex_app.Visible = false;
                ex_wb = ex_app.Workbooks.Add();

                UEIToExcel(ex_wb, OrgDepUserInfo);

                ex_wb.SaveAs(_sFilePath);
            }
            catch (Exception ex)
            {
                string str =  ex.ToString();
                MessageBox.Show(ex.ToString(),"Export to excel error.");
            }
            finally
            {
                if (ex_wb != null)
                    Marshal.ReleaseComObject(ex_wb);

                if (ex_app != null)
                {
                    ex_app.Quit();
                    Marshal.ReleaseComObject(ex_app);
                }
            }
        }

        /// <summary>
        /// Экспорт в эксель телефонного справочника пользователей.
        /// </summary>
        /// <param name="_ex_wb">Книга эксель для экспорта.</param>
        /// <param name="_SRDic">Справочник пользователей с группировкой по организации и отделу в формате [название_организации - [название_отдела-[список_пользователей]]]</param>
        private void UEIToExcel(Excel.Workbook _ex_wb, Dictionary<string, Dictionary<string, List<UserExportInfo>>> _SRDic)
        {
            Excel.Worksheet WS = _ex_wb.Sheets.Add() as Excel.Worksheet;
            WS.Name = "Справочник";
            WS.Range["B1"].Value = "Список телефонов";
            WS.Range["A2"].Value = "Фамилия, Имя, Отчество";
            WS.Range["B2"].Value = "Должность";
            WS.Range["C2"].Value = "Городской";
            WS.Range["D2"].Value = "Мобильный";
            WS.Range["E2"].Value = "Внут.";
            WS.Range["F2"].Value = "Работает с";

            Excel.Range HeaderRNG = WS.Range["A1:F2"];
            HeaderRNG.Interior.Pattern = Excel.XlPattern.xlPatternSolid;
            HeaderRNG.Interior.PatternColorIndex = Excel.XlPattern.xlPatternAutomatic;
            HeaderRNG.Interior.Color = 6974311;
            HeaderRNG.Interior.TintAndShade = 0;
            HeaderRNG.Interior.PatternTintAndShade = 0;
            HeaderRNG.Font.Bold = true;
            HeaderRNG.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            Marshal.ReleaseComObject(HeaderRNG);

            
            int iCurRow = 2;

            foreach (string sOrg in _SRDic.Keys)
            {
                // пробегаем по организациям
                iCurRow++;
                WS.Range["A" + iCurRow.ToString()].Value = sOrg;
                Excel.Range TitleRNG = WS.Range["A" + iCurRow.ToString() + ":" + "F" + iCurRow.ToString()];
                TitleRNG.Interior.Pattern = Excel.XlPattern.xlPatternSolid;
                TitleRNG.Interior.PatternColorIndex = Excel.XlPattern.xlPatternAutomatic;
                TitleRNG.Interior.Color = 8692350;
                TitleRNG.Interior.TintAndShade = 0;
                TitleRNG.Interior.PatternTintAndShade = 0;
                TitleRNG.Font.Bold = true;
                TitleRNG.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                Marshal.ReleaseComObject(TitleRNG);

                List<string> deplist = _SRDic[sOrg].Keys.ToList();
                deplist.Sort();

                foreach (string sDep in deplist)
                {
                    // пробегаем по отделам в организации
                    iCurRow++;
                    int iStart = iCurRow;
                    WS.Range["A" + iCurRow.ToString()].Value = sDep;
                    Excel.Range DepRNG = WS.Range["A" + iCurRow.ToString() + ":" + "F"  + iCurRow.ToString()];
                    DepRNG.Interior.Pattern = Excel.XlPattern.xlPatternSolid;
                    DepRNG.Interior.PatternColorIndex = Excel.XlPattern.xlPatternAutomatic;
                    DepRNG.Interior.Color = 12961221;
                    DepRNG.Interior.TintAndShade = 0;
                    DepRNG.Interior.PatternTintAndShade = 0;
                    DepRNG.Font.Bold = true;
                    DepRNG.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    Marshal.ReleaseComObject(DepRNG);

                    foreach (UserExportInfo uei in _SRDic[sOrg][sDep])
                    {
                        // выгружаем всех пользователей отдела
                        iCurRow++;
                        WS.Range["A" + iCurRow.ToString()].Value = uei.m_sName;
                        WS.Range["B" + iCurRow.ToString()].Value = uei.m_sTitle;
                        WS.Range["C" + iCurRow.ToString()].Value = uei.m_sPhoneNum;
                        WS.Range["D" + iCurRow.ToString()].Value = uei.m_sMobile;
                        WS.Range["E" + iCurRow.ToString()].Value = uei.m_sPhoneInternal;
                        WS.Range["F" + iCurRow.ToString()].Value = uei.m_sWCreated.ToString("dd.MM.yyyy");
                    }
                 

                    Excel.Range rng = WS.Rows[iStart.ToString() + ":" + iCurRow.ToString()] as Excel.Range;
                
                    Marshal.ReleaseComObject(rng);
                   iCurRow++;
                }
            }
            Excel.Range ResRNG = WS.Range["A1" +":F" + iCurRow.ToString()];
            // После выгрузки в эксель выравниваем столбцы
            WS.Range["A1"].EntireColumn.AutoFit();
            WS.Range["B1"].EntireColumn.AutoFit();
            WS.Range["C1"].EntireColumn.AutoFit();
            WS.Range["D1"].EntireColumn.AutoFit();
            WS.Range["E1"].EntireColumn.AutoFit();
            WS.Range["F1:Z1"].EntireColumn.Hidden=true;

            

            ResRNG.Borders[Excel.XlBordersIndex.xlDiagonalDown].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
            ResRNG.Borders[Excel.XlBordersIndex.xlDiagonalUp].LineStyle = Excel.XlLineStyle.xlLineStyleNone;

            ResRNG.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlContinuous;
            ResRNG.Borders[Excel.XlBordersIndex.xlInsideHorizontal].ColorIndex = 0;
            ResRNG.Borders[Excel.XlBordersIndex.xlInsideHorizontal].TintAndShade = 0;
            ResRNG.Borders[Excel.XlBordersIndex.xlInsideHorizontal].Weight = Excel.XlBorderWeight.xlThin;

            ResRNG.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
            ResRNG.Borders[Excel.XlBordersIndex.xlInsideVertical].ColorIndex = 0;
            ResRNG.Borders[Excel.XlBordersIndex.xlInsideVertical].TintAndShade = 0;
            ResRNG.Borders[Excel.XlBordersIndex.xlInsideVertical].Weight = Excel.XlBorderWeight.xlThin;

            ResRNG.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            ResRNG.Borders[Excel.XlBordersIndex.xlEdgeTop].ColorIndex = 0;
            ResRNG.Borders[Excel.XlBordersIndex.xlEdgeTop].TintAndShade = 0;
            ResRNG.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;

            ResRNG.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            ResRNG.Borders[Excel.XlBordersIndex.xlEdgeBottom].ColorIndex = 0;
            ResRNG.Borders[Excel.XlBordersIndex.xlEdgeBottom].TintAndShade = 0;
            ResRNG.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;

            ResRNG.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            ResRNG.Borders[Excel.XlBordersIndex.xlEdgeLeft].ColorIndex = 0;
            ResRNG.Borders[Excel.XlBordersIndex.xlEdgeLeft].TintAndShade = 0;
            ResRNG.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;

            ResRNG.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            ResRNG.Borders[Excel.XlBordersIndex.xlEdgeRight].ColorIndex = 0;
            ResRNG.Borders[Excel.XlBordersIndex.xlEdgeRight].TintAndShade = 0;
            ResRNG.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;
           
            Marshal.ReleaseComObject(ResRNG);

            // удаляем все листа в книге, кроме текущего
            List<string> WSNames = new List<string>();
            foreach (Excel._Worksheet CurWS in _ex_wb.Worksheets)
            {
                WSNames.Add(CurWS.Name);
            }

            foreach (string sName in WSNames)
            {
                if (!sName.Equals(WS.Name))
                    (_ex_wb.Worksheets[sName] as Excel._Worksheet).Delete();
            }

            Marshal.ReleaseComObject(WS);
        }

        /// <summary>
        /// Формируем справочник польователей с группировкой по организации и отделу из результата запроса в AD для экспорта в эсель.
        /// </summary>
        /// <param name="_ResInfoDic">Справочник пользователей с группировкой по организации и отделу в формате [название_организации - [название_отдела-[список_пользователей]]]</param>
        /// <param name="_SRDic">Список записей AD полученный по запросу в AD</param>
        /// <param name="_bNeedSort">Флаг задающий необходимость выполнить сортровку пользователей по алфавиту</param>
        private void GetExportInfo(Dictionary<string, Dictionary<string,List<UserExportInfo>>> _ResInfoDic, Dictionary<string, SearchResult> _SRDic, bool _bNeedSort)
        {
            foreach (SearchResult sr in _SRDic.Values)
            {
                // пропускаем отключенные записи
                if (IsAccountDisabled(sr))
                    continue;

                UserExportInfo uei = new UserExportInfo();
                object tmpobj = sr.GetDirectoryEntry().Properties["company"].Value;
                uei.m_sCompany = tmpobj != null ? tmpobj.ToString() : string.Empty;
                tmpobj = sr.GetDirectoryEntry().Properties["department"].Value;
                uei.m_sDepartment = tmpobj != null ? tmpobj.ToString() : string.Empty;
                tmpobj = sr.GetDirectoryEntry().Properties["mobile"].Value;
                uei.m_sMobile = tmpobj != null ? tmpobj.ToString() : string.Empty;
                tmpobj = sr.GetDirectoryEntry().Properties["displayName"].Value;
                uei.m_sName = tmpobj != null ? tmpobj.ToString() : string.Empty;
                tmpobj = sr.GetDirectoryEntry().Properties["homePhone"].Value;
                uei.m_sPhoneNum = tmpobj != null ? tmpobj.ToString() : string.Empty;
                tmpobj = sr.GetDirectoryEntry().Properties["l"].Value;
                uei.m_sSity = tmpobj != null ? tmpobj.ToString() : string.Empty;
                tmpobj = sr.GetDirectoryEntry().Properties["title"].Value;
                uei.m_sTitle = tmpobj != null ? tmpobj.ToString() : string.Empty;
                
                tmpobj = sr.GetDirectoryEntry().Properties["telephoneNumber"].Value;
                uei.m_sPhoneInternal = tmpobj != null ? tmpobj.ToString() : string.Empty;

                tmpobj = sr.GetDirectoryEntry().Properties["whenCreated"].Value;
                uei.m_sWCreated =  Convert.ToDateTime(tmpobj);
                

                //rs.GetDirectoryEntry().Properties["whenCreated"].Value.ToString();

                if (!_ResInfoDic.ContainsKey(uei.m_sCompany))
                {
                    _ResInfoDic.Add(uei.m_sCompany, new Dictionary<string, List<UserExportInfo>>());
                }

                if (!_ResInfoDic[uei.m_sCompany].ContainsKey(uei.m_sDepartment))
                {
                    _ResInfoDic[uei.m_sCompany].Add(uei.m_sDepartment, new List<UserExportInfo>());
                }

                _ResInfoDic[uei.m_sCompany][uei.m_sDepartment].Add(uei);
            }

            if (_bNeedSort)
            {
                foreach (string sOrgName in _ResInfoDic.Keys)
                {
                    foreach (string sDepName in _ResInfoDic[sOrgName].Keys)
                    {
                        _ResInfoDic[sOrgName][sDepName].Sort(CompareUserExportInfoByName);
                    }
                }
            }

        }

        /// <summary>
        /// компаратор для объектов типа UserExportInfo по имени.
        /// </summary>
        /// <param name="x"></param>
        /// <param name="y"></param>
        /// <returns></returns>
        private int CompareUserExportInfoByName(UserExportInfo x, UserExportInfo y)
        {
            if (x == null && y == null)
                return 0;
            else if (x == null)
                return -1;
            else if (y == null)
                return 1;
            else
                return x.m_sName.CompareTo(y.m_sName);
        }

        private void ShowUserInformation()
        {
            Cursor.Current = Cursors.Default;
            listView1.BeginUpdate();

            foreach (SearchResult rs in rsc)
            {
                ListViewItem listUser = new ListViewItem();


                //foreach (var res in rs.GetDirectoryEntry().Properties.PropertyNames)
                    
                //{
                //    Console.WriteLine(res.ToString() + " = " + rs.GetDirectoryEntry().Properties[res.ToString()].Value.ToString());
                //}


                string sid = new SecurityIdentifier((byte[])rs.GetDirectoryEntry().Properties["objectSid"].Value, 0).ToString();

                if (rs.GetDirectoryEntry().Properties["samaccountname"].Value != null)
                    listUser.SubItems[0].Text = rs.GetDirectoryEntry().Properties["samaccountname"].Value.ToString();
                else
                    listUser.SubItems[0].Text = "";

                if (rs.GetDirectoryEntry().Properties["displayName"].Value != null)
                    listUser.SubItems.Add(rs.GetDirectoryEntry().Properties["displayName"].Value.ToString());
                else
                    listUser.SubItems.Add("");

                if (rs.GetDirectoryEntry().Properties["mail"].Value != null)
                    listUser.SubItems.Add(rs.GetDirectoryEntry().Properties["mail"].Value.ToString());
                    
                else
                    listUser.SubItems.Add("");
                                
                    listUser.SubItems.Add(sid);

                resdic.Add(listUser.SubItems[3].Text, rs);
                listView1.Items.Add(listUser);
                
            }
            listView1.EndUpdate();
        }

        private bool IsAccountDisabled(SearchResult _rsuac)
        {
            int uAC = 0;
            if (_rsuac.GetDirectoryEntry().Properties["userAccountControl"].Value != null)
                uAC = Convert.ToInt32(_rsuac.GetDirectoryEntry().Properties["userAccountControl"].Value);
            else
                throw new ArgumentException(); 

            UserAccountControl userAccountControl = (UserAccountControl)uAC;

            return (userAccountControl & UserAccountControl.ACCOUNTDISABLE) == UserAccountControl.ACCOUNTDISABLE;
        }

        private void ShowAccountOptions(SearchResult _rsuac)
        {
            int uAC = 0;
            if (_rsuac.GetDirectoryEntry().Properties["userAccountControl"].Value != null)
                uAC = Convert.ToInt32(_rsuac.GetDirectoryEntry().Properties["userAccountControl"].Value);

            UserAccountControl userAccountControl = (UserAccountControl)uAC;

            // This gets a comma separated string of the flag names that apply.
            string userAccountControlFlagNames = userAccountControl.ToString();

            // This is how you test for an individual flag.
            // Password Expires, Account disable, Smartcard to logon, UnlockAccount
            bool isNormalAccount = (userAccountControl & UserAccountControl.NORMAL_ACCOUNT) == UserAccountControl.NORMAL_ACCOUNT;
          
            bool isAccountDisabled = (userAccountControl & UserAccountControl.ACCOUNTDISABLE) == UserAccountControl.ACCOUNTDISABLE;
               
            if (chbAccountDisabled.Checked = isAccountDisabled)
            {
                lbAccountDisabled.ForeColor = Color.Red;
                lbAccountDisabled.Font = new Font(lbAccountDisabled.Font, FontStyle.Bold);
                chbAccountDisabled.Checked = isAccountDisabled;
            }
            else
            {
                lbAccountDisabled.ForeColor = Color.Black;
                lbAccountDisabled.Font = new Font(lbAccountDisabled.Font, FontStyle.Regular);
            }
            
            
            
            bool isAccountPASSWD_CANT_CHANGE = (userAccountControl & UserAccountControl.PASSWD_CANT_CHANGE) == UserAccountControl.PASSWD_CANT_CHANGE;
            chbPassCantchange.Checked = isAccountPASSWD_CANT_CHANGE;
            
            bool isAccountDONT_EXPIRE_PASSWD = (userAccountControl & UserAccountControl.DONT_EXPIRE_PASSWD) == UserAccountControl.DONT_EXPIRE_PASSWD;
            if (chbPasswdDontExpire.Checked = isAccountDONT_EXPIRE_PASSWD)
            {
                lbPasswdDontExpire.ForeColor = Color.Green;
                chbPasswdDontExpire.Checked = isAccountDONT_EXPIRE_PASSWD;
            }
            else
            {
                lbPasswdDontExpire.ForeColor = Color.Black;
            }
            
            bool isAccountSMARTCARD_REQUIRED = (userAccountControl & UserAccountControl.SMARTCARD_REQUIRED) == UserAccountControl.SMARTCARD_REQUIRED;
            chbSmartRequired.Checked = isAccountSMARTCARD_REQUIRED;

            //bool isAccountLOCKOUT = (userAccountControl & UserAccountControl.LOCKOUT) == UserAccountControl.LOCKOUT;
            //cBLOCKOUT.Checked = isAccountLOCKOUT;
 
        }

        private void ShowEmailAddress(SearchResult _rsmail)
        {
           
           foreach (string eAddress in _rsmail.Properties["proxyAddresses"])
           {
              string[] mailarr = eAddress.Split(':');
               ListViewItem maillist = new ListViewItem();
               maillist.SubItems[0].Text = mailarr[0];
               if (mailarr[0].Equals("SMTP"))
               {
                   maillist.Font = new Font(maillist.Font, FontStyle.Bold);
               }
               maillist.SubItems.Add(mailarr[1]);
               listView3.Items.Add(maillist);
           }

        }

        private void ShowDirectReport(SearchResult _rsdirreport)
        {
            foreach (string userdirreport in _rsdirreport.Properties["directReports"])
            {
                string usrdirrep = userdirreport;
                int memofstart = userdirreport.IndexOf("CN=");
                int memofend = userdirreport.IndexOf(",", memofstart);
                usrdirrep = userdirreport.Substring(memofstart + 3, memofend - memofstart - 3);
                ListViewItem dirreplist = new ListViewItem();
                dirreplist.SubItems[0].Text = usrdirrep;
                listView5.Items.Add(dirreplist);
            }
        }

        private void ShowMemberOf(SearchResult _rsmemberof)
        {
            foreach (string usermemberof in _rsmemberof.Properties["memberOf"])
            {
                string usrmemof = usermemberof;
                int memofstart = usermemberof.IndexOf("CN=");
                int memofend = usermemberof.IndexOf(",", memofstart);
                usrmemof = usermemberof.Substring(memofstart + 3, memofend - memofstart - 3);
                ListViewItem memoflist = new ListViewItem();
                memoflist.SubItems[0].Text = usrmemof;
                listView4.Items.Add(memoflist);
            }
        }

        private void ShowX509Detalise(SearchResult _rs)
        {
            if (_rs.Properties.Contains("userCertificate"))
            {
                for (int i = 0; i < _rs.Properties["userCertificate"].Count; i++ )
                {
                    Byte[] b = (Byte[])_rs.Properties["userCertificate"][i];  //This is hard coded to the first element.  Some users may have multiples.  Use ADSI Edit to find out more.
                    X509Certificate2 cert1 = new X509Certificate2(b);
                    string x509Name = cert1.GetName();
                    int NamaStart = x509Name.IndexOf("CN=");
                    int NameEnd = x509Name.IndexOf(",", NamaStart);
                    if (NamaStart > 0 && NameEnd > 0)
                    {
                        x509Name = x509Name.Substring(NamaStart + 3, NameEnd - NamaStart - 3);
                        string x509DateFrom = cert1.GetEffectiveDateString();
                        string x509DataExp = cert1.GetExpirationDateString();
                        ListViewItem certlist = new ListViewItem();
                        certlist.SubItems[0].Text = x509Name;
                        certlist.SubItems.Add(x509DateFrom);
                        certlist.SubItems.Add(x509DataExp);
                        listView2.Items.Add(certlist);
                    }
                    else
                    {
                        ListViewItem certlist = new ListViewItem();
                        
                        certlist.SubItems[0].Text = cert1.GetName();
                        certlist.SubItems.Add(cert1.GetEffectiveDateString());
                        certlist.SubItems.Add(cert1.GetExpirationDateString());
                        listView2.Items.Add(certlist);
                    }
                    

                }
                
                
            }
        }

        public static Image GetUserPicture(object  _tf)
        {
            if (_tf == null)
                return null;

            var data = _tf as byte[];
            if (data != null)            
            {
                using (var s = new MemoryStream(data))
                { 
                    Bitmap OrigBmp = (Bitmap)Bitmap.FromStream(s);
                    double Scale = 1.0;
                    if (OrigBmp.Width > OrigBmp.Height)
                    {
                        Scale = 96.0 / OrigBmp.Width;
                    }
                    else
                    {
                        Scale = 96.0 / OrigBmp.Height;
                    }
                    int newW = (int) (OrigBmp.Width * Scale);
                    int newH = (int)(OrigBmp.Height * Scale);

                    Bitmap ResBmp = new Bitmap(OrigBmp, new Size(newW, newH));
                    return ResBmp;
                }
            }
            else
                return null;
        }

        private DirectorySearcher GetDirectorySearcher(string username, string passowrd, string domain)
         {
            if (chB_Authorize.Checked)
             {
                 username = txtUsername.Text;
                 passowrd = txtPassword.Text;
                 domain = txtAddress.Text;
              
                     DirectoryEntry ds = new DirectoryEntry("LDAP://" + domain, username, passowrd);
                     DirectorySearcher dirSearch = new DirectorySearcher(ds);
                     
                ////ExchangeService _service = new ExchangeService(ExchangeVersion.Exchange2010_SP1);
                ////_service.Credentials = new WebCredentials(username, passowrd);
                ////_service.AutodiscoverUrl(username + "@" + domain);

                ////     Server _exserv = new Server();
                ////     ExchangeServer _service = new ExchangeServer(_exserv);
                
                ////ADUser _exaduser = new ADUser();

                ////CASMailbox _casmail = new CASMailbox(_exaduser);
               
                ////string exus = _exaduser.DisplayName.ToString();
                ////string exserv = _service.Fqdn.ToString();
                ////string excas = _casmail.DisplayName.ToString();
                
               
                
                
                dirSearch.FindOne();
                     this.Font = new Font(this.Font, FontStyle.Italic);
                     this.Text = "Авторизован пользователь: " + username;    
                     return dirSearch;
             }
            else
             {
                 try
                 {
                     DirectoryEntry ds = new DirectoryEntry();
                     DirectorySearcher dirSearch = new DirectorySearcher(ds);
                     this.ForeColor = Color.Black;
                     this.Font = new Font(this.Font, FontStyle.Regular);
                     this.Text = "Querying Active Directory";    
                     return dirSearch;
                 }
                catch (DirectoryServicesCOMException e)
                {
                    MessageBox.Show("Ошибка соединения с Active Directory!! ", "ОШИБКА",MessageBoxButtons.OK,MessageBoxIcon.Error);
                    e.Message.ToString();
                    
                }
                
              
             }
              return null; 
     
        }

        private SearchResultCollection SearchUserByUserName(DirectorySearcher ds, string username)
        {

            
            ds.Filter = "(&(objectClass=user)(objectCategory=person)(ANR=" + username + "))";
            
            ds.SearchScope = SearchScope.Subtree;
            ds.ServerTimeLimit = TimeSpan.FromSeconds(90);
            
                rsc = ds.FindAll(); 
            
            
            if (rsc != null)
                return rsc;
            else
                return null;         
        }

        private SearchResultCollection SearchUserByEmail(DirectorySearcher ds, string email)
        {
            ds.Filter = "(&(objectClass=user)(objectCategory=person)(|(proxyAddresses=smtp:" + email + ")(proxyAddresses=SMTP:" + email + ")))";

            
            ds.SearchScope = SearchScope.Subtree;
            ds.ServerTimeLimit = TimeSpan.FromSeconds(90);

            rsc = ds.FindAll();
           

            if (rsc != null)
                return rsc;
            else
                return null;
        }

        private SearchResultCollection SearchUserByFirstName(DirectorySearcher ds, string FN)
        {
            ds.Filter = "(&(objectClass=user)(objectCategory=person)(|(givenName=" + FN + ")(sn=" + FN + ")))";
            ds.SearchScope = SearchScope.Subtree;
            ds.ServerTimeLimit = TimeSpan.FromSeconds(90);

            rsc = ds.FindAll();


            if (rsc != null)
                return rsc;
            else
                return null;
        }

        private SearchResultCollection SearchUserByPhoneNo(DirectorySearcher ds, string PhoneNo)
        {
            ds.Filter = "(&(objectClass=user)(objectCategory=person)(telephoneNumber=" + PhoneNo + "))";

            ds.SearchScope = SearchScope.Subtree;
            ds.ServerTimeLimit = TimeSpan.FromSeconds(90);

            rsc = ds.FindAll();


            if (rsc != null)
                return rsc;
            else
                return null;
        }

        private void btnSearchUserName_Click(object sender, EventArgs e)
            
        {
            resdic.Clear();
            ClearUserDetales();
            listView1.Items.Clear();
            listView2.Items.Clear();
         

            if (chB_Authorize.Checked)
            {
                if (string.IsNullOrEmpty(txtUsername.Text) || string.IsNullOrEmpty(txtPassword.Text) || string.IsNullOrEmpty(txtAddress.Text))
                {
                    MessageBox.Show("Не заполнены поля авторизации!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }

            try
            {
                GetUserInformation(txtUsername.Text.Trim(), txtPassword.Text.Trim(), txtAddress.Text.Trim());
            }
            catch (DirectoryServicesCOMException come)
            {
                MessageBox.Show(come.ToString(), "ОШИБКА", MessageBoxButtons.OK, MessageBoxIcon.Error);
                
                return;
            }
             catch (COMException COMex)
            {
                MessageBox.Show(COMex.ToString(), "ОШИБКА", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
                
           
        }

        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listView1.SelectedItems.Count>0)
            {
                ClearUserDetales();
                string sid = listView1.SelectedItems[0].SubItems[3].Text;

                ShowUserDetales(sid);
            }
        }

        private void MenuExit_Click(object sender, EventArgs e)
        {
            Application.Exit();
            
        }

        private void MenuAbout_Click(object sender, EventArgs e)
        {
            
            AboutBox1 AbBox = new AboutBox1();
            AbBox.ShowDialog();
    
        }

       private void exportToExcelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Dictionary<string, SearchResult> resexp = new Dictionary<string, SearchResult>();

            try
            {
                string sFileName = string.Empty;
                SaveFileDialog SFD = new SaveFileDialog();
                SFD.Filter = "excel wb (*.xlsx) | .xlsx";
                if (SFD.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    sFileName = SFD.FileName;
                }
                else
                {
                    return;
                }

                SearchResultCollection rs = SearchUserByPhoneNo(GetDirectorySearcher(txtUsername.Text.Trim(), txtPassword.Text.Trim(), txtAddress.Text.Trim()), "*");
                foreach (SearchResult rss in rsc)
                {
                    string sid = new SecurityIdentifier((byte[])rss.GetDirectoryEntry().Properties["objectSid"].Value, 0).ToString(); ;
                    resexp.Add(sid, rss);
                }

                if (m_BGW_ExcelExport.IsBusy)
                    return;

                //ExportToExcel(sFileName, resexp);
                ExportUsersToExcelArgument EUArg = new ExportUsersToExcelArgument();
                EUArg.ADSR = resexp;
                EUArg.sExcelWBFileName = sFileName;

                m_BGW_ExcelExport.RunWorkerAsync(EUArg);
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private Dictionary<string, SearchResult> PrepUserExp()
        {
            
            return null;
        }

        private void rbUseremail_CheckedChanged(object sender, EventArgs e)
        {
            exportToExcelToolStripMenuItem.Visible = false;
        }

        private void rbPhoneNo_CheckedChanged(object sender, EventArgs e)
        {
            exportToExcelToolStripMenuItem.Visible = true;
        }

        private void rbFirstname_CheckedChanged(object sender, EventArgs e)
        {
            exportToExcelToolStripMenuItem.Visible = false;
        }

        private void chB_Authorize_CheckedChanged(object sender, EventArgs e)
        {
            txtUsername.ReadOnly = !chB_Authorize.Checked;
            txtPassword.ReadOnly = !chB_Authorize.Checked;
            txtAddress.ReadOnly = !chB_Authorize.Checked;
            if (chB_Authorize.Checked)
            {
                lbAuthorise.ForeColor = Color.Red;
                lbAuthorise.Font = new Font(lbAccountDisabled.Font, FontStyle.Bold);
            }
            else
            {
                lbAuthorise.ForeColor = Color.Black;
                lbAuthorise.Font = new Font(lbAccountDisabled.Font, FontStyle.Regular);
            }
        }
        /// <summary>
        /// Поиск компьютеров сети отображение из конфигурации
        /// </summary>
        //////////////////private void btn_LocWMIQueryRun_Click_1(object sender, EventArgs e)
        //////////////////{
        //////////////////    ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT Name FROM Win32_Processor");
        //////////////////    ManagementObjectCollection searcherCollection = searcher.Get();
        //////////////////    foreach (ManagementObject mo in searcherCollection)
        //////////////////    {
        //////////////////        foreach (PropertyData prop in mo.Properties)
        //////////////////        {
        //////////////////            lbCPU.Text = prop.Value.ToString(); //вывожу название процессора в метку на форме
        //////////////////        }
        //////////////////    }
        //////////////////}

        private void btn_NetWMIQueryRun_Click_1(object sender, EventArgs e)
        {
            ManagementScope scope = new ManagementScope("\\\\" + tB_ComputerName.Text + "\\root\\cimv2"); //область поиска. tB_ComputerName.Text - ТекстБокс с которого я возьму имя компьютера
            ObjectQuery query = new ObjectQuery("SELECT Name FROM Win32_Processor"); // Запрос
            ManagementObjectSearcher searcher = new ManagementObjectSearcher(scope, query);
            ManagementObjectCollection searcherCollection = searcher.Get();
            foreach (ManagementObject mo in searcherCollection)
            {
                foreach (PropertyData prop in mo.Properties)
                {
                    lbCPU.Text = prop.Value.ToString(); //вывожу название процессора в метку на форме
                }
            }
        }

        private void btn_FindNetworkComputer_Click(object sender, EventArgs e)
        {
        
            try
            {
                
                DirectoryEntry enTry = new DirectoryEntry("LDAP://DC=feib,DC=local");
                DirectorySearcher mySearcher = new DirectorySearcher(enTry);
                int UF_ACCOUNTDISABLE = 0x0002; // Исключаем из поиска отключенный компьютеры
                String searchFilter = "(&(objectClass=computer)(!(userAccountControl:1.2.840.113556.1.4.803:=" + UF_ACCOUNTDISABLE.ToString() + ")))";
                mySearcher.Filter = (searchFilter);
                SearchResultCollection resEnt = mySearcher.FindAll();
                foreach (SearchResult srItem in resEnt)
                {
                    ListViewItem listPC = new ListViewItem();
                    

                    listPC.SubItems[0].Text = srItem.GetDirectoryEntry().Name.ToString().Substring(3).ToUpper();

                    listPC.SubItems.Add(srItem.GetDirectoryEntry().Properties["operatingsystem"].Value.ToString());
                    listView6.Items.Add(listPC);
                }
            }
            catch (Exception)
            {
            }
        }

        private void listView6_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listView6.SelectedItems.Count > 0)
            {
                try
                {
                    string PCSel = listView6.SelectedItems[0].SubItems[0].Text;

                    ManagementScope scope = new ManagementScope("\\\\" + PCSel + "\\root\\cimv2"); //область поиска. tB_ComputerName.Text - ТекстБокс с которого я возьму имя компьютера
                    ObjectQuery query = new ObjectQuery("SELECT Name FROM Win32_Processor"); // Запрос
                    ManagementObjectSearcher searcher = new ManagementObjectSearcher(scope, query);
                    ManagementObjectCollection searcherCollection = searcher.Get();
                    foreach (ManagementObject mo in searcherCollection)
                    {
                        foreach (PropertyData prop in mo.Properties)
                        {
                            lbCPU.Text = prop.Value.ToString(); //вывожу название процессора в метку на форме
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
             }

           
    }

    /// <summary>
    /// Flags that control the behavior of the user account.
    /// </summary>
    [Flags()]
    public enum UserAccountControl : int
    {
        /// <summary>
        /// The logon script is executed. 
        ///</summary>
        SCRIPT = 0x00000001,

        /// <summary>
        /// The user account is disabled. 
        ///</summary>
        ACCOUNTDISABLE = 0x00000002,

        /// <summary>
        /// The home directory is required. 
        ///</summary>
        HOMEDIR_REQUIRED = 0x00000008,

        /// <summary>
        /// The account is currently locked out. 
        ///</summary>
        LOCKOUT = 0x00000010,

        /// <summary>
        /// No password is required. 
        ///</summary>
        PASSWD_NOTREQD = 0x00000020,

        /// <summary>
        /// The user cannot change the password. 
        ///</summary>
        /// <remarks>
        /// Note:  You cannot assign the permission settings of PASSWD_CANT_CHANGE by directly modifying the UserAccountControl attribute. 
        /// For more information and a code example that shows how to prevent a user from changing the password, see User Cannot Change Password.
        // </remarks>
        PASSWD_CANT_CHANGE = 0x00000040,

        /// <summary>
        /// The user can send an encrypted password. 
        ///</summary>
        ENCRYPTED_TEXT_PASSWORD_ALLOWED = 0x00000080,

        /// <summary>
        /// This is an account for users whose primary account is in another domain. This account provides user access to this domain, but not 
        /// to any domain that trusts this domain. Also known as a local user account. 
        ///</summary>
        TEMP_DUPLICATE_ACCOUNT = 0x00000100,

        /// <summary>
        /// This is a default account type that represents a typical user. 
        ///</summary>
        NORMAL_ACCOUNT = 0x00000200,

        /// <summary>
        /// This is a permit to trust account for a system domain that trusts other domains. 
        ///</summary>
        INTERDOMAIN_TRUST_ACCOUNT = 0x00000800,

        /// <summary>
        /// This is a computer account for a computer that is a member of this domain. 
        ///</summary>
        WORKSTATION_TRUST_ACCOUNT = 0x00001000,

        /// <summary>
        /// This is a computer account for a system backup domain controller that is a member of this domain. 
        ///</summary>
        SERVER_TRUST_ACCOUNT = 0x00002000,

        /// <summary>
        /// Not used. 
        ///</summary>
        Unused1 = 0x00004000,

        /// <summary>
        /// Not used. 
        ///</summary>
        Unused2 = 0x00008000,

        /// <summary>
        /// The password for this account will never expire. 
        ///</summary>
        DONT_EXPIRE_PASSWD = 0x00010000,

        /// <summary>
        /// This is an MNS logon account. 
        ///</summary>
        MNS_LOGON_ACCOUNT = 0x00020000,

        /// <summary>
        /// The user must log on using a smart card. 
        ///</summary>
        SMARTCARD_REQUIRED = 0x00040000,

        /// <summary>
        /// The service account (user or computer account), under which a service runs, is trusted for Kerberos delegation. Any such service 
        /// can impersonate a client requesting the service. 
        ///</summary>
        TRUSTED_FOR_DELEGATION = 0x00080000,

        /// <summary>
        /// The security context of the user will not be delegated to a service even if the service account is set as trusted for Kerberos delegation. 
        ///</summary>
        NOT_DELEGATED = 0x00100000,

        /// <summary>
        /// Restrict this principal to use only Data Encryption Standard (DES) encryption types for keys. 
        ///</summary>
        USE_DES_KEY_ONLY = 0x00200000,

        /// <summary>
        /// This account does not require Kerberos pre-authentication for logon. 
        ///</summary>
        DONT_REQUIRE_PREAUTH = 0x00400000,

        /// <summary>
        /// The user password has expired. This flag is created by the system using data from the Pwd-Last-Set attribute and the domain policy. 
        ///</summary>
        PASSWORD_EXPIRED = 0x00800000,

        /// <summary>
        /// The account is enabled for delegation. This is a security-sensitive setting; accounts with this option enabled should be strictly 
        /// controlled. This setting enables a service running under the account to assume a client identity and authenticate as that user to 
        /// other remote servers on the network.
        ///</summary>
        TRUSTED_TO_AUTHENTICATE_FOR_DELEGATION = 0x01000000,

        /// <summary>
        /// 
        /// </summary>
        PARTIAL_SECRETS_ACCOUNT = 0x04000000,

        /// <summary>
        /// 
        /// </summary>
        USE_AES_KEYS = 0x08000000
    }

    public class UserExportInfo
    {
        public string m_sName;
        public string m_sDepartment;
        public string m_sSity;
        public string m_sCompany;
        public string m_sTitle;
        public string m_sPhoneNum;
        public string m_sPhoneInternal;
        public string m_sMobile;
        
        public DateTime m_sWCreated;

    }

    public class ExportUsersToExcelArgument
    {
        public Dictionary<string, SearchResult> ADSR = null;
        public string sExcelWBFileName = null;
    }

    

    

    

    
    
 }

}
