using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Diagnostics;
using System.Security.Permissions;
using System.DirectoryServices;
using System.DirectoryServices.ActiveDirectory;
using System.DirectoryServices.AccountManagement;
using System.Security.Cryptography.X509Certificates;


//[assembly: SecurityPermission(SecurityAction.RequestMinimum, Execution = true)]
//[assembly: DirectoryServicesPermission(SecurityAction.RequestMinimum)]

namespace ActiveDirectory
{
    public partial class frmAD : Form
    {
        public DirectorySearcher dirSearch = null;
        public SearchResultCollection rsc;
        public Dictionary<string, SearchResult> resdic= new Dictionary<string,SearchResult>();
        public frmAD()
        {
            InitializeComponent();
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
            if(txtSearchUser.Text.Trim().IndexOf("@") > 0)
                rs = SearchUserByEmail(GetDirectorySearcher(username, passowrd, domain), txtSearchUser.Text.Trim());
            else
                rs = SearchUserByUserName(GetDirectorySearcher(username, passowrd, domain), txtSearchUser.Text.Trim());

            if (rsc != null)
            {
                ShowUserInformation();
            }
            else
            {
                MessageBox.Show("User not found!!!", "Search Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void ClearUserDetales()
        {
            lblUsernameDisplay.Text = "Username : ";


                 lblFirstname.Text = "First Name : ";


                lblMiddleName.Text = "Middle Name : ";

            
                lblLastName.Text = "Last Name : " ;

            

                lblEmailId.Text = "Email ID : " ;

            
                lblTitle.Text = "Title : " ;

            
                lblCompany.Text = "Company : " ;

            
                lblCity.Text = "City : " ;

            
                lblState.Text = "State : " ;

            
                lblCountry.Text = "Country : " ;

            
                lblPostal.Text = "Postal Code : " ;


                lblTelephone.Text = "Telephone No. : ";

                pbUserImg.Image = null;
                listView2.Items.Clear();
            

        }
        private void ShowUserDetales(string _AccountName)
        {
            if (!resdic.ContainsKey(_AccountName))
                return;
            SearchResult rs = resdic[_AccountName];
    
            if (rs.GetDirectoryEntry().Properties["samaccountname"].Value != null)
            lblUsernameDisplay.Text = "Username : " + rs.GetDirectoryEntry().Properties["samaccountname"].Value.ToString();

            if (rs.GetDirectoryEntry().Properties["givenName"].Value != null)
            lblFirstname.Text = "First Name : " + rs.GetDirectoryEntry().Properties["givenName"].Value.ToString();

            if (rs.GetDirectoryEntry().Properties["initials"].Value != null)
                lblMiddleName.Text = "Middle Name : " + rs.GetDirectoryEntry().Properties["initials"].Value.ToString();

            if (rs.GetDirectoryEntry().Properties["sn"].Value != null)
                lblLastName.Text = "Last Name : " + rs.GetDirectoryEntry().Properties["sn"].Value.ToString();

            if (rs.GetDirectoryEntry().Properties["mail"].Value != null)

                lblEmailId.Text = "Email ID : " + rs.GetDirectoryEntry().Properties["mail"].Value.ToString();

            if (rs.GetDirectoryEntry().Properties["title"].Value != null)
                lblTitle.Text = "Title : " + rs.GetDirectoryEntry().Properties["title"].Value.ToString();

            if (rs.GetDirectoryEntry().Properties["company"].Value != null)
                lblCompany.Text = "Company : " + rs.GetDirectoryEntry().Properties["company"].Value.ToString();

            if (rs.GetDirectoryEntry().Properties["l"].Value != null)
                lblCity.Text = "City : " + rs.GetDirectoryEntry().Properties["l"].Value.ToString();

            if (rs.GetDirectoryEntry().Properties["st"].Value != null)
                lblState.Text = "State : " + rs.GetDirectoryEntry().Properties["st"].Value.ToString();

            if (rs.GetDirectoryEntry().Properties["co"].Value != null)
                lblCountry.Text = "Country : " + rs.GetDirectoryEntry().Properties["co"].Value.ToString();

            if (rs.GetDirectoryEntry().Properties["postalCode"].Value != null)
                lblPostal.Text = "Postal Code : " + rs.GetDirectoryEntry().Properties["postalCode"].Value.ToString();

            if (rs.GetDirectoryEntry().Properties["telephoneNumber"].Value != null)
                lblTelephone.Text = "Telephone No. : " + rs.GetDirectoryEntry().Properties["telephoneNumber"].Value.ToString();

            if (rs.GetDirectoryEntry().Properties["thumbnailPhoto"].Value != null)
            {
                string UserName = rs.GetDirectoryEntry().Properties["samaccountname"].Value.ToString();
                pbUserImg.Image = GetUserPicture(UserName);
            }

            ShowX509Detalise(rs);

        }
        private void ShowUserInformation()
        {
            Cursor.Current = Cursors.Default;
            listView1.BeginUpdate();

            foreach (SearchResult rs in rsc)
            {
                ListViewItem listUser = new ListViewItem();


                foreach (var res in rs.GetDirectoryEntry().Properties.PropertyNames)
                {
                    
                    Console.WriteLine(res.ToString() + " = " + rs.GetDirectoryEntry().Properties[res.ToString()].Value.ToString());
                }
                

                if (rs.GetDirectoryEntry().Properties["samaccountname"].Value != null)
                    listUser.SubItems[0].Text = rs.GetDirectoryEntry().Properties["samaccountname"].Value.ToString();
         
                if (rs.GetDirectoryEntry().Properties["givenName"].Value != null)
                    listUser.SubItems.Add(rs.GetDirectoryEntry().Properties["givenName"].Value.ToString());
         
                if (rs.GetDirectoryEntry().Properties["mail"].Value != null)
                    listUser.SubItems.Add(rs.GetDirectoryEntry().Properties["mail"].Value.ToString());
         

                resdic.Add(listUser.SubItems[0].Text, rs);
                listView1.Items.Add(listUser);
                
            }
            listView1.EndUpdate();
        }
        private void ShowX509Detalise(SearchResult _rs)
        {
            if (_rs.Properties.Contains("userCertificate"))
            {
                for (int i = 0; i < _rs.Properties["userCertificate"].Count; i++ )
                {
                    Byte[] b = (Byte[])_rs.Properties["userCertificate"][i];  //This is hard coded to the first element.  Some users may have multiples.  Use ADSI Edit to find out more.
                    X509Certificate cert1 = new X509Certificate(b);
                    string x509Name = cert1.GetName();
                    int NamaStart = x509Name.IndexOf("CN=");
                    int NameEnd = x509Name.IndexOf(",", NamaStart);
                    x509Name = x509Name.Substring(NamaStart + 3, NameEnd - NamaStart - 3);
                    
                    string x509DateFrom = cert1.GetEffectiveDateString();
                    string x509DataExp = cert1.GetExpirationDateString();
                    ListViewItem certlist = new ListViewItem();
                    certlist.SubItems[0].Text = x509Name;
                    certlist.SubItems.Add(x509DateFrom);
                    certlist.SubItems.Add(x509DataExp);
                    listView2.Items.Add(certlist);
                }
                
                
            }
        }
        public static Image GetUserPicture(string userName)
        {
            using (var dsSearcher = new DirectorySearcher())
            {
                var idx = userName.IndexOf('\\');
                if (idx > 0)
                    userName = userName.Substring(idx + 1);
                dsSearcher.Filter = string.Format("(&(objectClass=user)(objectCategory=person)(samaccountname={0}))", userName);
                SearchResult result = dsSearcher.FindOne();
                if (result != null)
                {
                    using (var user = new DirectoryEntry(result.Path))
                    {
                        var data = user.Properties["thumbnailPhoto"].Value as byte[];

                        if (data != null)
                        
                            using (var s = new MemoryStream(data))
                  //Image Convert         
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
                }
                return null;
            }
        }


        private DirectorySearcher GetDirectorySearcher(string username, string passowrd, string domain)
        {           
            if(dirSearch == null)
            {
                try
                {
                    dirSearch = new DirectorySearcher();
                        //new DirectoryEntry("LDAP://" + domain, username, passowrd));                    
                        
                }
                catch (DirectoryServicesCOMException e)
                {
                    MessageBox.Show("Connection Creditial is Wrong!!!, please Check.", "Erro Info",MessageBoxButtons.OK,MessageBoxIcon.Error);
                    e.Message.ToString();
                }
                return dirSearch;
            }
            else{
                return dirSearch;
            }
        }

        private SearchResultCollection SearchUserByUserName(DirectorySearcher ds, string username)
        {
            ds.Filter = "(&(objectClass=user)(objectCategory=person)(samaccountname=" + username + "))";

            ds.SearchScope = SearchScope.Subtree;
            ds.ServerTimeLimit = TimeSpan.FromSeconds(90);
            //ds.FindOne();
            rsc = ds.FindAll();
            //SearchResultCollection userObject = ds.FindAll();

            if (rsc != null)
                return rsc;
            else
                return null;         
        }

        private SearchResultCollection SearchUserByEmail(DirectorySearcher ds, string email)
        {
            ds.Filter = "(&(objectClass=user)(objectCategory=person)(proxyAddresses=" + email + "))";

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
            listView1.Items.Clear();
            listView2.Items.Clear();
                       
                GetUserInformation(txtUsername.Text.Trim(), txtPassword.Text.Trim(), txtAddress.Text.Trim());               
           
        }

       

        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listView1.SelectedItems.Count>0)
            {
                ClearUserDetales();
                string AccountName = listView1.SelectedItems[0].SubItems[0].Text;

                ShowUserDetales(AccountName);
            }
        }
    }
   
}
