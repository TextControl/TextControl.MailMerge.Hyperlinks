using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using TXTextControl.DocumentServer.Fields;

namespace TextControl.MailMerge.Hyperlinks
{
    public partial class index : System.Web.UI.Page
    {
        private TXTextControl.DocumentServer.MailMerge mailMerge1;
        private TXTextControl.ServerTextControl serverTextControl1;
        private System.ComponentModel.IContainer components;
    
        protected void Page_Load(object sender, EventArgs e)
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.mailMerge1 = new TXTextControl.DocumentServer.MailMerge(this.components);
            this.serverTextControl1 = new TXTextControl.ServerTextControl();
            // 
            // mailMerge1
            // 
            this.mailMerge1.TextComponent = this.serverTextControl1;
            this.mailMerge1.FieldMerged += new TXTextControl.DocumentServer.MailMerge.FieldMergedHandler(this.mailMerge1_FieldMerged);
            // 
            // serverTextControl1
            // 
            this.serverTextControl1.SpellChecker = null;
        }

        protected void Button1_Click(object sender, EventArgs e)
        {
            // load the template
            mailMerge1.LoadTemplate(Server.MapPath("template.docx"), TXTextControl.DocumentServer.FileFormat.WordprocessingML);

            // create a DataSet as data source for the merge process
            DataSet ds = new DataSet();
            ds.ReadXml(Server.MapPath("report.xml"), XmlReadMode.Auto);

            // merge the data
            mailMerge1.Merge(ds.Tables[0]);
            
            string data;
            mailMerge1.SaveDocumentToMemory(out data, TXTextControl.StringStreamType.HTMLFormat, null);

            // return the document as HTML
            Response.Write(data);
        }

        private void mailMerge1_FieldMerged(object sender, TXTextControl.DocumentServer.MailMerge.FieldMergedEventArgs e)
        {
            byte[] data;

            if(e.MailMergeFieldAdapter.TypeName == "MERGEFIELD")
            {
                MergeField field = new MergeField(e.MailMergeFieldAdapter.ApplicationField);
             
                // if the merge field starts with "%LINK%", we want to convert it
                // to a hyperlink
                if (field.Text.StartsWith("%LINK%") == true)
                {
                    // create a temporary ServerTextControl
                    using (TXTextControl.ServerTextControl tx = new TXTextControl.ServerTextControl())
                    {
                        tx.Create();
                        
                        // load the merge field to clone the formatting
                        tx.Load(e.MergedField, TXTextControl.BinaryStreamType.InternalUnicodeFormat);
                        
                        // remove the content
                        tx.ApplicationFields.Clear(true);
                        tx.SelectAll();
                        tx.Selection.Text = "";

                        // add a new hyperlink with the text part of the hyperlink
                        tx.HypertextLinks.Add(new TXTextControl.HypertextLink(field.ApplicationField.Text.Split(',')[1], field.ApplicationField.Text.Split(',')[1]));
                        
                        // save the content
                        tx.Save(out data, TXTextControl.BinaryStreamType.InternalUnicodeFormat);
                    }

                    // return the hyperlink to the merge process
                    e.MergedField = data;
                }

            }
        }
    }
}