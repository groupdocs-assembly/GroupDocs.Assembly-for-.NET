using System;
using System.IO;
using System.Threading.Tasks;
using System.Collections;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Web;
using System.Web.Http;
using GroupDocs.Assembly;
using GroupDocs.Assembly.Data;

namespace GroupDocs.Assembly.Live.Demos.UI.Controllers
{
    public class GroupDocsAssemblyController : ApiController
    {
        [HttpPost]
        public Hashtable CreateJob()
        {
            Hashtable result = new Hashtable();
            string jid = Guid.NewGuid().ToString();
            result.Add("jid", jid);
            Directory.CreateDirectory(Path.Combine(AppSettings.WorkingDirectory, jid));
            return result;
        }

        [HttpPost]
        public async Task<HttpResponseMessage> Upload(string jid)
        {
            if (!Request.Content.IsMimeMultipartContent())
            {
                throw new HttpResponseException(HttpStatusCode.UnsupportedMediaType);
            }

            var provider = new MultipartFormDataStreamProvider(AppSettings.WorkingDirectory);
            await Request.Content.ReadAsMultipartAsync(provider);

            Directory.CreateDirectory(Path.Combine(AppSettings.WorkingDirectory, jid));

            foreach (MultipartFileData file in provider.FileData)
            {
                string name = file.Headers.ContentDisposition.FileName.Trim('"');
                string path = Path.Combine(AppSettings.WorkingDirectory, jid, name);
                File.Copy(file.LocalFileName, path, true);
                File.Delete(file.LocalFileName);
            }

            return Request.CreateResponse(HttpStatusCode.NoContent);
        }

        [HttpPost]
        public HttpResponseMessage Template(string jid, string filename)
        {
            return Request.CreateResponse(HttpStatusCode.NoContent);
        }

        [HttpPost]
        public HttpResponseMessage DocumentTableDataSource(string jid, string name, string filename, int index)
        {
            return Request.CreateResponse(HttpStatusCode.NoContent);
        }

        [HttpPost]
        public async Task<Hashtable> Assemble(string jid, string templateFilename, string datasourceFilename, string assembledFilename, string datasourceName = null, int datasourceTableIndex = 0)
        {
            DocumentAssembler assembler = new DocumentAssembler();
            assembler.Options = DocumentAssemblyOptions.AllowMissingMembers;

            DataSourceInfo source = new DataSourceInfo();
            source.DataSource = new GroupDocs.Assembly.Data.DocumentTable(
                Path.Combine(AppSettings.WorkingDirectory, jid, datasourceFilename),
                datasourceTableIndex,
                new DocumentTableOptions()
                {
                    FirstRowContainsColumnNames = true
                }
            );
            source.Name = datasourceName;

            string sourcePath = Path.Combine(AppSettings.WorkingDirectory, jid, templateFilename);
            string targetPath = Path.Combine(AppSettings.WorkingDirectory, jid, assembledFilename);

            await Task.Run(() => {
                using (Stream sourceStream = new FileStream(sourcePath, FileMode.Open))
                {
                    using (Stream targetStream = new FileStream(targetPath, FileMode.Create))
                    {
                        assembler.AssembleDocument(sourceStream, targetStream, source);
                    }
                }
            });

            Hashtable result = new Hashtable();
            result.Add("filename", assembledFilename);
            return result;
        }

        [HttpGet]
        public HttpResponseMessage Download(string jid, string filename)
        {
            if (jid.IndexOfAny(Path.GetInvalidFileNameChars()) > 0)
            {
                return Request.CreateErrorResponse(HttpStatusCode.NotFound, "Invalid jid");
            }

            if (filename.IndexOfAny(Path.GetInvalidFileNameChars()) > 0)
            {
                return Request.CreateErrorResponse(HttpStatusCode.NotFound, "Invalid filename");
            }

            var path = Path.Combine(AppSettings.WorkingDirectory, jid, filename);

            Stream s;
            try
            {
                s = new FileStream(path, FileMode.Open);
            }
            catch (System.IO.FileNotFoundException x)
            {
                return Request.CreateErrorResponse(HttpStatusCode.NotFound, x);
            }

            var result = Request.CreateResponse(HttpStatusCode.OK);
            result.Content = new StreamContent(s);
            result.Content.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");
            result.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment")
            {
                FileName = filename
            };
            return result;
        }
    }
}
