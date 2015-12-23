using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using GroupDocs.ReportGenerator.BusinessLayer;

namespace GroupDocs.ReportGenerator
{
    public partial class Default : System.Web.UI.Page
    {
        #region Properties

        static string uploadedTemplatePath;
        static string uploadedTemplateType;

        #endregion

        #region Page load and events

        // page load event
        protected void Page_Load(object sender, EventArgs e)
        {
            try
            {
                // verify page is not post back, so we can setup default page view.
                if (!Page.IsPostBack)
                {
                    // Load form view.
                    InitForm();
                }
            }
            catch (Exception exc)
            {
                lblMessage.Text = exc.Message;
            }
        }

        //Reload/Populate Tables & Views dropdownlist
        protected void TestConnectionAndLoad_Click(object sender, EventArgs e)
        {
            try
            {
                ddlSource.Attributes.Add("onchange", "javascript:sourceChanged();");
                // Creating business layer object with user provided connection string
                BusinessDB businessDBObj = new BusinessDB(txtConString.Text.Trim());

                // Verify connection string successfully connect database
                if (businessDBObj.IsValidConnection())
                {
                    // populate dataset with tables
                    DataSet dsDBObjects = businessDBObj.getDBObjectNames(DBObjects.Tables);
                    if (dsDBObjects != null)
                    {
                        ddlTables.DataSource = dsDBObjects;
                        ddlTables.DataTextField = "name";
                        ddlTables.DataValueField = "name";
                        ddlTables.DataBind();
                    }

                    // populate dataset with views
                    dsDBObjects = null;
                    dsDBObjects = businessDBObj.getDBObjectNames(DBObjects.Views);
                    if (dsDBObjects != null)
                    {
                        ddlViews.DataSource = dsDBObjects;
                        ddlViews.DataTextField = "name";
                        ddlViews.DataValueField = "name";
                        ddlViews.DataBind();
                    }
                    lblMessage.Text = "Successfully connected to database.";
                }
                else
                {
                    ddlTables.Items.Clear();
                    ddlViews.Items.Clear();
                    txtCustomQuery.Text = "";
                }
            }
            catch (Exception exc)
            {
                lblMessage.Text = exc.Message;
            }
        }

        //Verify and upoload template
        protected void UploadTemplate_Click(object sender, EventArgs e)
        {
            try
            {
                lblMessage.Text = "";
                ddlSource.Attributes.Add("onchange", "javascript:sourceChanged();");
                // Verify file and upload to server and save file location in constant variable/property.
                if (fuTemplate.HasFile)
                {
                    uploadedTemplatePath = "";
                    uploadedTemplateType = "";
                    uploadedTemplateType = fuTemplate.FileName.Split('.')[fuTemplate.FileName.Split('.').Length - 1];

                    if (!Directory.Exists(Server.MapPath("~/UploadTemplates")))
                    {
                        Directory.CreateDirectory(Server.MapPath("~/UploadTemplates"));
                    }
                    if (!Directory.Exists(Server.MapPath("~/GeneratedReports")))
                    {
                        Directory.CreateDirectory(Server.MapPath("~/GeneratedReports"));
                    }

                    // verify uploaded file (MS Word, MS Excel, MS Power Point)
                    if (uploadedTemplateType.ToLower().Equals("doc") || uploadedTemplateType.ToLower().Equals("docx") || uploadedTemplateType.ToLower().Equals("xls") || uploadedTemplateType.ToLower().Equals("xlsx") || uploadedTemplateType.ToLower().Equals("ppt") || uploadedTemplateType.ToLower().Equals("pptx"))
                    {
                        fuTemplate.SaveAs(Server.MapPath("~/UploadTemplates/" + fuTemplate.FileName));
                        uploadedTemplatePath = Server.MapPath("~/UploadTemplates/" + fuTemplate.FileName);
                    }
                    else
                    {
                        lblMessage.Text = "Please select valid template file (MS Word, MS Excel, MS Power Point)";
                        return;
                    }

                    lblMessage.Text = "Template uploaded successfully.";
                }
                else
                {
                    lblMessage.Text = "Please select template file.";
                }
            }
            catch (Exception exc)
            {
                lblMessage.Text = exc.Message;
            }
        }

        //Export data file as per selected export type
        protected void GenerateReport_Click(object sender, EventArgs e)
        {
            try
            {
                lblMessage.Text = "";
                ddlSource.Attributes.Add("onchange", "javascript:sourceChanged();");

                lblMessage.Text = "Data exported successfully.";
                if (uploadedTemplatePath != "")
                {
                    // Creating business layer object with user provided connection string
                    BusinessDB businessDBObj = new BusinessDB(txtConString.Text.Trim());

                    // Verify connection string successfully connect database
                    if (businessDBObj.IsValidConnection())
                    {
                        Response.Clear();
                        Response.Buffer = true;
                        Response.AddHeader("content-disposition", "attachment;filename=GroupDocs_GeneratedReport_" + DateTime.Now.Day.ToString() + "_" + DateTime.Now.Month.ToString() + "_" + DateTime.Now.Year.ToString() + "_" + DateTime.Now.Hour.ToString() + DateTime.Now.Minute.ToString() + DateTime.Now.Second.ToString() + "_" + DateTime.Now.Millisecond.ToString() + "." + uploadedTemplateType);
                        Response.Charset = "";
                        Response.ContentType = "Application/" + uploadedTemplateType;

                        // if data source type is Table
                        if (ddlSource.SelectedItem.Text == "Tables")
                        {
                            businessDBObj.GenerateReport_DBO(ddlTables.SelectedValue, uploadedTemplatePath.Replace("UploadTemplates", "GeneratedReports"), uploadedTemplatePath);
                            Response.WriteFile(businessDBObj.ReportDestinationPath);
                        }
                        // if data source type is Views
                        else if (ddlSource.SelectedItem.Text == "Views")
                        {
                            businessDBObj.GenerateReport_DBO(ddlViews.SelectedValue, uploadedTemplatePath.Replace("UploadTemplates", "GeneratedReports"), uploadedTemplatePath);
                            Response.WriteFile(businessDBObj.ReportDestinationPath);
                        }
                        // if data source type is custom query
                        else
                        {
                            businessDBObj.GenerateReport_Query(txtCustomQuery.Text.Trim(), uploadedTemplatePath.Replace("UploadTemplates", "GeneratedReports"), uploadedTemplatePath);
                            Response.WriteFile(businessDBObj.ReportDestinationPath);
                        }
                        Response.End();
                    }
                }
                else
                {
                    lblMessage.Text = "Please upload template file.";
                }
            }
            catch (Exception exc)
            {
                lblMessage.Text = exc.Message;
            }
        }

        //Clear form data
        protected void ClearForm_Click(object sender, EventArgs e)
        {
            try
            {
                lblMessage.Text = "";
                Response.Redirect("~/Default.aspx");
            }
            catch (Exception exc)
            {
                lblMessage.Text = exc.Message;
            }
        }

        #endregion

        #region Private methods

        //Form Initiating
        private void InitForm()
        {
            uploadedTemplatePath = "";
            ddlSource.Attributes.Add("onchange", "javascript:sourceChanged();");

            ddlTables.Style["display"] = "none";
            ddlViews.Style["display"] = "none";
            txtCustomQuery.Style["display"] = "none";

            switch (ddlSource.SelectedValue)
            {
                case "0":
                    ddlTables.Style["display"] = "block";
                    break;
                case "1":
                    ddlViews.Style["display"] = "block";
                    break;
                case "2":
                    txtCustomQuery.Style["display"] = "block";
                    break;
            }
        }

        #endregion
    }
}