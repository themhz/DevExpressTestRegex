using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
using System.Text.RegularExpressions;

namespace testReport2
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            test();
        }

        string targetFile = @"..\\..\\..\\Main.docx";
        string outputFile = @"..\\..\\..\\Output.docx";
        public void test()
        {
            using (RichEditDocumentServer parentWordProcessor = new RichEditDocumentServer())
            {
                parentWordProcessor.Document.BeginUpdate();
                parentWordProcessor.LoadDocument(targetFile);

                var regexSectionOpen = new Regex("^#.((.|\n)*?).*:");
                var regexSectionClose = new Regex("^#.*\\/");

                foreach (var paragraph in parentWordProcessor.Document.Paragraphs)
                {
                    var matchesOpen = parentWordProcessor.Document.FindAll(regexSectionOpen, paragraph.Range);
                    foreach (var range in matchesOpen)
                    {
                        // Replace matched text with "target1"
                        parentWordProcessor.Document.Replace(range, "target1");
                    }

                    var matchesClose = parentWordProcessor.Document.FindAll(regexSectionClose, paragraph.Range);
                    foreach (var range in matchesClose)
                    {
                        // Replace matched text with "target2"
                        parentWordProcessor.Document.Replace(range, "target2");
                    }
                }

                parentWordProcessor.Document.EndUpdate();
                parentWordProcessor.SaveDocument(outputFile, DocumentFormat.OpenXml);

                this.OpenDocFile();
            }
        
        }

        public DocumentRange getTextRange(string search, RichEditDocumentServer wp = null)
        {
            Regex myRegEx = new Regex(search);
            return wp.Document.FindAll(myRegEx).First();
        }

        public void OpenDocFile(){

            RichEditDocumentServer childWordPrecessor = new RichEditDocumentServer();
            childWordPrecessor.Document.BeginUpdate();
            childWordPrecessor.LoadDocument(targetFile);
            
        }
    }
}