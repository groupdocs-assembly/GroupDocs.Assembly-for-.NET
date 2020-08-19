<%@ Page Language="C#" AutoEventWireup="true" ContentType="text/xml"%>

<%
    string[] extensions = "DOC DOCX DOT DOTX DOTM DOCM RTF XML XLS XLSX XLSM XLSB XLT XLTM XLTX PPT PPTX PPTM PPS PPSX PPSM POTX POTM EML EMLX MSG ODT OTT ODS ODP MHT MHTML HTML TXT".Split(' ');
%>

<!-- Generated on <%= System.DateTime.Now.ToString("MMMM d, yyyy") %> -->

<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">
    <url>
        <loc>https://products.groupdocs.app/assembly/total</loc>
        <priority>0.9</priority>
    </url>
    <%foreach (var e in extensions) { %><url>
        <loc>https://products.groupdocs.app/assembly/<%= e.ToLower() %></loc>
        <priority>0.9</priority>
    </url><% } %>
</urlset>
