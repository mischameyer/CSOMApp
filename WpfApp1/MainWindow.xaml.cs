using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Windows;
using SPClient = Microsoft.SharePoint.Client;
using Microsoft.Win32;
using System.IO;

/// <summary>
/// This WPF Examples uses the CSOM (Client Object Model) to demonstrate the basic operations with a remote SharePoint-Site:
/// https://docs.microsoft.com/en-us/sharepoint/dev/sp-add-ins/complete-basic-operations-using-sharepoint-client-library-code
/// </summary>
namespace WpfApp1
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private Connection.SpConnector ctx;
        public MainWindow()
        {
            InitializeComponent();

            this.tabControl.Visibility = Visibility.Hidden;
        }
        /// <summary>
        /// Trigger for connecting to Sharepoint
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnConnect_Click(object sender, RoutedEventArgs e)
        {

            try
            {
                ctx = new Connection.SpConnector(this.txbUsername.Text, this.passwordBox.SecurePassword);

                MessageBox.Show("Connected to: " + ctx.SPContext.Web.Title);                

                ctx.SPContext.Dispose();

                this.tabControl.Visibility = Visibility.Visible;

            } catch (Exception ex)
            {
                RaiseException(ex);
            }            

        }

        /// <summary>
        /// Gets all List items of site
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnGetLists_Click(object sender, RoutedEventArgs e)
        {

            Web web = ctx.SPContext.Web;
            ctx.SPContext.Load(web.Lists, lists => lists);

            IEnumerable<SPClient.List> result = ctx.SPContext.LoadQuery(web.Lists.Include( // For each list, retrieve Title and Id.
                                                                   list => list.Title,
                                                                   list => list.Id));

            ctx.SPContext.ExecuteQuery();

            foreach (var list in result)
            {
                txtOutputLists.Text = txtOutputLists.Text + "\n\r" + list.Title;
            }

        }

        /// <summary>
        /// Get all items from a specific list
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnGetListItems_Click(object sender, RoutedEventArgs e)
        {

            this.txtOutputItems.Text = "";

            Web web = ctx.SPContext.Web;

            SPClient.List announcementsList = ctx.SPContext.Web.Lists.GetByTitle("MyFirstList");

            CamlQuery query = CamlQuery.CreateAllItemsQuery(100);
            SPClient.ListItemCollection items = announcementsList.GetItems(query);

            ctx.SPContext.Load(items);
            ctx.SPContext.ExecuteQuery();

            foreach (ListItem listItem in items)
            {
                // We have all the list item data. For example, Title.
                txtOutputItems.Text = txtOutputItems.Text + "\n\r" + listItem["ID"] + " | " + listItem["Title"] + " | " + listItem["Description"];
            }

        }

        /// <summary>
        /// Insert a new list item to a specific list
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnInsertItem_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                SPClient.List announcementsList = ctx.SPContext.Web.Lists.GetByTitle("MyFirstList");

                ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                ListItem newItem = announcementsList.AddItem(itemCreateInfo);

                newItem["Title"] = this.txtBoxTitle.Text;
                newItem["Description"] = this.txtBoxDescription.Text;

                newItem.Update();

                ctx.SPContext.ExecuteQuery();

                MessageBox.Show("Item inserted");

                this.txtBoxTitle.Text = "";
                this.txtBoxDescription.Text = "";

            } catch(Exception ex)
            {
                RaiseException(ex);
            }
            

        }        

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnUpdateItem_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                SPClient.List announcementsList = ctx.SPContext.Web.Lists.GetByTitle("MyFirstList");

                ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                ListItem listItem = announcementsList.GetItemById(this.txtBoxId.Text);

                listItem["Description"] = this.txtBoxDescriptionUpdate.Text;
                listItem.Update();

                ctx.SPContext.ExecuteQuery();

                MessageBox.Show("Item inserted");

                this.txtBoxId.Text = "";
                this.txtBoxDescriptionUpdate.Text = "";

            }
            catch (Exception ex)
            {
                RaiseException(ex);
            }


        }

        private void btnOpenFile_Click(object sender, RoutedEventArgs e)
        {

            try
            {
                OpenFileDialog openFileDialog = new OpenFileDialog();

                if (openFileDialog.ShowDialog() == true)
                {
                    txtEditor.Text = System.IO.File.ReadAllText(openFileDialog.FileName);

                    var formLib = ctx.SPContext.Web.Lists.GetByTitle("MyFirstDocument");

                    ctx.SPContext.Load(formLib.RootFolder);
                    ctx.SPContext.ExecuteQuery();
                    string fileUrl = "";

                    using (var fs = new FileStream(openFileDialog.FileName, FileMode.Open))
                    {
                        var fi = new FileInfo(openFileDialog.FileName); //file Title  
                        fileUrl = String.Format("{0}/{1}", formLib.RootFolder.ServerRelativeUrl, fi.Name);
                        SPClient.File.SaveBinaryDirect(ctx.SPContext, fileUrl, fs, true);
                        ctx.SPContext.ExecuteQuery();
                    }

                    var libFields = formLib.Fields;
                    ctx.SPContext.Load(libFields);
                    ctx.SPContext.ExecuteQuery();
                    SPClient.File newFile = ctx.SPContext.Web.GetFileByServerRelativeUrl(fileUrl);
                    ListItem item = newFile.ListItemAllFields;
                    item.Update();
                    ctx.SPContext.ExecuteQuery();
                    MessageBox.Show("File saved to SharePoint");
                }

            }
            catch (Exception ex)
            {
                RaiseException(ex);
            }
                



        }

        /// <summary>
        /// Shows the exceptions within a messageBox
        /// </summary>
        /// <param name="ex"></param>
        private void RaiseException(Exception ex)
        {
            MessageBox.Show(ex.ToString());
        }

        
    }
}
