<%@ Page Title="Free online document and report generator for DOCX, PPTX, XLSX, DOT and other formats" Language="C#" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="GroupDocs.Assembly.Live.Demos.UI.Assembly.Default" MasterPageFile="~/site.Master" %>

<asp:Content ContentPlaceHolderID="MainContent" runat="server">

    <script src="https://cms.dynabic.com/templates/aspose/js/bootstrap.js?v=331"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/angular.js/1.7.7/angular.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/angular.js/1.7.7/angular-route.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/angular.js/1.7.7/angular-animate.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/angular.js/1.7.7/angular-aria.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/angular.js/1.7.7/angular-resource.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/angular.js/1.7.7/angular-messages.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/angular-material/1.1.12/angular-material.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/danialfarid-angular-file-upload/12.2.13/ng-file-upload-all.min.js"></script>

    <script type="text/javascript">
        var REQUESTED_EXTENSION = <%= Page.RouteData.Values["extension"] != null ? "\"" + Page.RouteData.Values["extension"] + "\"" : "null" %>;
        var REQUESTED_FILEFORMAT = <%= "null" %>;
    </script>

    <script src="/AssemblyApp/const.js"></script>
    <script src="/AssemblyApp/app.js"></script>
    <script src="/AssemblyApp/app.config.js"></script>
    <script src="/AssemblyApp/app.run.js"></script>
    <script src="/AssemblyApp/app.controller.main.js"></script>
    <link rel="stylesheet" href="/AssemblyApp/style.css" />

    <div ng-app="GroupDocsAssemblyApp" ng-cloak ng-controller="Main">

        <div class="container-fluid GroupDocsApps pb5">
            <div class="container">
                <div class="row">

                    <div class="col-md-12 pt-5 pb-5">
                        <h1 id="hFeatureTitle" ng-hide="REQUESTED_EXTENSION">Free Online Document Assembly</h1>
                        <h1 id="hFeatureTitle" ng-show="REQUESTED_EXTENSION">Free Online {{REQUESTED_EXTENSION}} file Assembly</h1>
                        <h4 id="hPageTitle" ng-hide="REQUESTED_EXTENSION">Generate meaningful documents from template and raw data, supporting Word, Excel, Email and more than 50 types of documents</h4>
                        <h4 id="hPageTitle" ng-show="REQUESTED_EXTENSION">Generate meaningful documents from {{REQUESTED_EXTENSION}} template files, and many other supported file formats</h4>

                        <ng-form class="form" ng-show="stage === 1" onsubmit="return false">
                        <div class="uploadfile">

                            <div class="filedropdown">
                                <div class="filedrop">
                                    <label class="dz-message needsclick">
                                        Drop or upload your template file
                                    </label>
                                    <input type="file"
                                           ngf-select
                                           ng-model="templateFile.file"
                                           accept="*/*" ngf-max-size="100MB"
                                           ngf-model-invalid="templateFile.error"
                                           required/>
                                    <br/>

                                    <div class="fileupload" ng-show="templateFile.file">
                                        <div class="progress">
                                            <div class="progress-bar progress-bar-striped progress-bar-success progress-bar-animated"
                                                 style="width:{{templateFile.progress}}%"></div>
                                        </div>
                                        <span class="filename" ng-show="templateFile.file.name">
                                            <a ng-click="templateFile={}">
                                                <label class="custom-file-upload">{{templateFile.file.name}}</label>
                                                <i class="fa fa-times">&nbsp;</i>
                                            </a>
                                        </span>
                                    </div>
                                </div>

                                <p id="pMessage"
                                   style="width: 65%; position: relative; color: red; margin: 50px auto 30px;">
                                </p>

                                <div class="convertbtn">
                                    <span ng-click="uploadTemplateFile()" class="btn btn-success btn-lg"
                                            ng-disabled="!templateFile.file">
                                        <i class="fa fa-file-upload">&nbsp;</i>
                                        Upload Template
                                    </span>
                                </div>
                                <div class="convertbtn">
                                    <span class="btn btn-success btn-lg px-5"
                                            data-toggle="modal" data-target="#help-dialog-template">
                                        <i class="fa fa-info-circle">&nbsp;</i>
                                        Help
                                    </span>
                                </div>
                            </div>
                        </div>
                    </ng-form>

                        <ng-form class="form" ng-show="stage === 2" onsubmit="return false">
                        <div class="uploadfile">
                            <div class="filedropdown">
                                <div class="filedrop">
                                    <label class="dz-message needsclick">
                                        Drop or upload your data-source file
                                    </label>
                                    <input type="file"
                                           ngf-select
                                           ng-model="datasourceFile.file"
                                           accept="*/*" ngf-max-size="100MB"
                                           ngf-model-invalid="datasourceFile.error"
                                           required/>
                                    <br/>

                                    <div class="fileupload" ng-show="datasourceFile.file">
                                        <div class="progress">
                                            <div class="progress-bar progress-bar-striped progress-bar-success progress-bar-animated"
                                                 style="width:{{datasourceFile.progress}}%"></div>
                                        </div>
                                        <span class="filename" ng-show="datasourceFile.file.name">
                                             <a ng-click="datasourceFile={}">
                                                <label class="custom-file-upload">{{datasourceFile.file.name}}</label>
                                                <i class="fa fa-times">&nbsp;</i>
                                            </a>
                                        </span>
                                    </div>
                                </div>

                                <table class="saveas">
                                    <tr>
                                        <td align="right">
                                            <em>Table Index</em>
                                        </td>
                                        <td align="left">
                                            <div class="btn-group saveformat">
                                                <input class="btn" ng-model="datasourceTableIndex" type="number" min="0"
                                                       placeholder="Integer"/>
                                            </div>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td align="right">
                                            <em>Data Source Name</em>
                                        </td>
                                        <td align="left">
                                            <div class="btn-group saveformat">
                                                <input class="btn" ng-model="datasourceName" type="text"
                                                       placeholder="Name"/>
                                            </div>
                                        </td>
                                    </tr>
                                </table>
                                <br/>

                                <p id="pMessage"
                                   style="width: 65%; position: relative; color: red; margin: 50px auto 30px;">
                                </p>
                                <div class="convertbtn">
                                    <span ng-click="uploadDatasourceFile()" class="btn btn-success btn-lg"
                                            ng-disabled="!datasourceFile.file">
                                        <i class="fa fa-file-upload">&nbsp;</i>
                                        Upload Data Source
                                    </span>
                                </div>
                                <div class="convertbtn">
                                    <span ng-click="" class="btn btn-success btn-lg px-5"
                                            data-toggle="modal" data-target="#help-dialog-datasource">
                                        <i class="fa fa-info-circle">&nbsp;</i>
                                        Help
                                    </span>
                                </div>
                            </div>
                            <div class="clearfix">&nbsp;</div>
                            <a href="" ng-click="stage = 1" class="btn btn-link" style="color:white">
                                <i class="fa fa-arrow-left">&nbsp;</i>
                                Go Back
                            </a>
                        </div>
                    </ng-form>

                        <ng-form class="form" ng-show="stage === 3" onsubmit="return false">
                        <div class="uploadfile">
                            <div class="filedropdown">
                                <div class="filesuccess">
                                    <label class="dz-message needsclick">
                                        Your document is ready to be assembled
                                    </label>
                                    <span class="downloadbtn convertbtn">
                                        <a href="#" class="btn btn-success btn-lg" ng-click="assembleDocument()">
                                            <i class="fa fa-file-download">&nbsp;</i>
                                            ASSEMBLE NOW
                                        </a>
                                    </span>
                                </div>
                            </div>
                            <a href="" ng-click="stage = 2" class="btn btn-link" style="color:white">
                                <i class="fa fa-arrow-left">&nbsp;</i>
                                Go Back
                            </a>
                        </div>
                    </ng-form>

                    </div>
                </div>
            </div>
        </div>

        <div class="modal fade" id="help-dialog-template" tabindex="-1" role="dialog">
            <div class="modal-dialog" role="document">
                <div class="modal-content">
                    <div class="modal-header">
                        <h5 class="modal-title">Help &raquo; Upload Template</h5>
                    </div>
                    <div class="modal-body">
                        <p>
                            Please read the documentation at
                        <a href="https://docs.groupdocs.com/display/assemblynet">docs.groupdocs.com/display/assemblynet
                        </a>
                            for creating templates.
                        </p>
                        <p>
                            For quick testing, please create a new Word document
                        and paste the following.
                        </p>
                        <blockquote>
                            Product List:<br />
                            <br />
                            &lt;&lt;foreach [in products]&gt;&gt;<br />
                            Name: &lt;&lt;[ProductName]&gt;&gt;<br />
                            Price: &lt;&lt;[Price]&gt;&gt;<br />
                            Description: &lt;&lt;[Description]&gt;&gt;<br />
                            <br />
                            &lt;&lt;/foreach&gt;&gt;
                        </blockquote>
                        <p>
                            Save this file as <b>Template.docx</b> and upload it as template.
                        </p>
                    </div>
                    <div class="modal-footer">
                        <span class="btn btn-primary" data-dismiss="modal">Close</span>
                    </div>
                </div>
            </div>
        </div>

        <div class="modal fade" id="help-dialog-datasource" tabindex="-1" role="dialog">
            <div class="modal-dialog" role="document">
                <div class="modal-content">
                    <div class="modal-header">
                        <h5 class="modal-title">Help &raquo; Upload Data Source</h5>
                    </div>
                    <div class="modal-body">
                        <p>
                            Please read the documentation at
                        <a href="https://docs.groupdocs.com/display/assemblynet">docs.groupdocs.com/display/assemblynet
                        </a>
                            for creating data source files.
                        </p>
                        <p>
                            For quick testing, please create a new Excel worksheet
                        and add the data according to the following table.
                        </p>
                        <blockquote>
                            <table border="1">
                                <thead>
                                    <tr>
                                        <td>ProductName</td>
                                        <td>Price</td>
                                        <td>Description</td>
                                    </tr>
                                </thead>
                                <tbody>
                                    <tr>
                                        <td>Laptop</td>
                                        <td>1200.00</td>
                                        <td>MSI GP Series GP63 Leopard-428 15.6in</td>
                                    </tr>
                                    <tr>
                                        <td>Keyboard</td>
                                        <td>25.00</td>
                                        <td>Logitech K400 Plus Wireless Touch Keyboard</td>
                                    </tr>
                                    <tr>
                                        <td>Mouse</td>
                                        <td>15.00</td>
                                        <td>Logitech M325 RF Wireless Optical 1000</td>
                                    </tr>
                                    <tr>
                                        <td>LED Monitor</td>
                                        <td>300.00</td>
                                        <td>LG 34UB67-B Black 34in 21:9 Ultra wide</td>
                                    </tr>
                                </tbody>
                            </table>
                        </blockquote>
                        <p>
                            Save this file as <b>Data.xlsx</b> and upload it as data source.
                        <br />
                            Set Table Index: 0
                        <br />
                            Set Data Source Name: products
                        </p>
                    </div>
                    <div class="modal-footer">
                        <span class="btn btn-primary" data-dismiss="modal">Close</span>
                    </div>
                </div>
            </div>
        </div>

        <section class="col-md-12 pull-left d-flex d-wrap bg-gray appfeaturesectionlist" ng-show="REQUESTED_FILEFORMAT">
            <div class="col-md-6 cardbox tc col-md-offset-3 b6">
                <h3>
                    <em class="icofile"><b>{{REQUESTED_FILEFORMAT.Extension}}</b></em>
                    <span>{{REQUESTED_FILEFORMAT.Name}}</span>
                </h3>
                <p>
                    {{REQUESTED_FILEFORMAT.Description}}
                <br>
                    <br>
                    <a target="_blank" href="{{REQUESTED_FILEFORMAT.FileFormat_Com_URL}}" class="btn btn-success btn-sm">Read More</a>
                </p>
            </div>
        </section>

        <section class="col-md-12 pt-5 bg-gray tc" ng-show="REQUESTED_FILEFORMAT === null">
            <div class="container">
                <div class="col-md-12 pull-left">
                    <h2 class="h2title">GroupDocs.Assembly App</h2>
                    <p>GroupDocs.Assembly App Supported Document Formats</p>
                </div>
            </div>
            <div class="clearfix"></div>
            <div>
                <div class="diagram1 d2 d1-net">
                    <div class="d1-row">
                        <div class="d1-col d1-left">
                            <header>Microsoft Office Formats</header>
                            <ul>
                                <li><strong>Word:</strong> DOC, DOCX, DOT, DOTX, DOTM, DOCM, RTF, WordprocessingML (XML)</li>
                                <li><strong>Excel:</strong> XLS, XLSX, XLSM, XLSB, XLT, XLTM, XLTX, SpreadsheetML (XML)</li>
                                <li><strong>PowerPoint:</strong> PPT, PPTX, PPTM, PPS, PPSX, PPSM, POTX, POTM</li>
                                <li><strong>Outlook:</strong> EML, EMLX, MSG</li>
                            </ul>
                            <header>Supported Data Sources</header>
                            <ul>
                                <li>Spreadsheet as Table of Data</li>
                                <li>Word Processing Table as Table of Data</li>
                            </ul>
                        </div>
                        <div class="d1-col d1-right">
                            <header>Other Formats</header>
                            <ul>
                                <li><strong>OpenOffice Document Formats:</strong> ODT, OTT, ODS, ODP</li>
                                <li><strong>Email:</strong> MHT, MHTML</li>
                                <li><strong>Web:</strong> HTML</li>
                                <li><strong>Other :</strong> TXT</li>
                            </ul>
                            <header>Inter-Format Assembly Support</header>
                            <ul>
                                <li>Word Processing<strong> TO </strong>Word Processing, HTML, PDF, XPS, TIFF, MHTML, TXT, XAML, OpenXPS, EPUB, SVG, PS, PCL</li>
                                <li>Spreadsheet<strong> TO </strong>Spreadsheet, HTML, PDF, XPS, TIFF, MHTML</li>
                                <li>Presentation<strong> TO </strong>Presentation, HTML, PDF, XPS, TIFF</li>
                                <li>Email<strong> TO </strong>Word Processing, Email, HTML, PDF, XPS, TIFF, MHTML, TXT, XAML, OpenXPS, EPUB, SVG, PS, PCL</li>
                                <li>HTML &amp; TXT<strong> TO </strong>Word Processing, HTML, PDF, XPS, TIFF, MHTML, TXT, XAML, OpenXPS, EPUB, SVG, PS, PCL</li>
                            </ul>
                        </div>
                    </div>
                    <div class="d1-logo">
                        <img src="//www.groupdocs.cloud/templates/groupdocs/images/product-logos/90x90-noborder/groupdocs-assembly-net.png" alt="GroupDocs.Assembly App">
                        <header>GroupDocs.Assembly App</header>
                    </div>
                </div>
            </div>

        </section>

        <section class="col-md-12 tl bg-darkgray howtolist">
            <div class="container tl dflex">

                <div class="col-md-4 howtosectiongfx">
                    <img src="/img/howto.png">
                </div>
                <div class="howtosection col-md-8">
                    <div>
                        <h4><i class="fa fa-question-circle "></i>&nbsp;<b>How to assemble a {{REQUESTED_EXTENSION}} file using GroupDocs.Assembly App</b></h4>
                        <ul>
                            <li>Click inside the file drop area to upload a {{REQUESTED_EXTENSION}} template file or drag &amp; drop a {{REQUESTED_EXTENSION}} template file.</li>
                            <li>Click <b>Upload Template</b> to start uploading.</li>
                            <li>Click inside the file drop area to upload a data-source file or drag &amp; drop it.</li>
                            <li>Choose the appropriate <b>Table Index</b> and <b>Data Source Name</b> options.</li>
                            <ol style="list-style-type:none">
                                <li>Table Index is the zero-based serial number of the table that exist in the data-source file. For example, if you have a spreadsheet file as data-source file then its first sheet is at zero index.</li>
                                <li>Data Source Name is that name of data-source table which is used to reffer inside your {{REQUESTED_EXTENSION}} template file.</li>
                            </ol>
                            <li>In the next step, you can download your assembled {{REQUESTED_EXTENSION}} file.</li>
                        </ul>
                    </div>
                </div>
            </div>
        </section>

        <section class="col-md-12 pt-5 app-features-section">
            <div class="container tc pt-5">
                <div class="row m-0">
                    <div class="col-md-4">
                        <div class="imgcircle fasteasy">
                            <img src="//products.groupdocs.app/img/fast-easy.png">
                        </div>
                        <h4>Fast and Easy Assembly</h4>
                        <p>Upload your template and data-source files, choose the appropriate options and click on &ldquo;ASSEMBLE NOW&rdquo; button. You will get the output file as soon as the files are assembled.</p>
                    </div>
                    <div class="col-md-4">
                        <div class="imgcircle anywhere">
                            <img src="//products.groupdocs.app/img/anywhere.png">
                        </div>
                        <h4>Assemble from Anywhere</h4>
                        <p>It works from all platforms including Windows, Mac, Android and iOS. All files are processed on our servers. No plugin or software installation required for you.</p>
                    </div>
                    <div class="col-md-4">
                        <div class="imgcircle quality">
                            <img src="//products.groupdocs.app/img/quality.png">
                        </div>
                        <h4>Assembly Quality</h4>
                        <p>Powered by <a href="https://products.groupdocs.com/assembly/net" target="_blank">GroupDocs.Assembly</a>. All files are processed using GroupDocs APIs.</p>
                    </div>
                </div>
            </div>
        </section>

    </div>


</asp:Content>
