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

        string targetFile = "d:\\ProjNet2022\\applications\\Building.Project\\Building.UI\\data.el\\reports\\Building.Accessible\\AccessibleReport\\Main.docx";
        public void test()
        {
            using (RichEditDocumentServer parentWordProcessor = new RichEditDocumentServer())
            {
                parentWordProcessor.Document.BeginUpdate();
                parentWordProcessor.LoadDocument(targetFile);

                var regexSectionOpen = new Regex("^#.*\\((.|\\n)*?\\):");                
                var regexSectionClose = new Regex("^#.*/");

                List<DocumentRange> openings = new List<DocumentRange>();
                List<DocumentRange> closings = new List<DocumentRange>();
              
                foreach (var paragraph in parentWordProcessor.Document.Paragraphs)
                {
                    var matches = parentWordProcessor.Document.FindAll(regexSectionOpen, paragraph.Range);


                    if (matches != null && matches.Length > 0)
                        openings.AddRange(matches);

                    matches = parentWordProcessor.Document.FindAll(regexSectionClose, paragraph.Range);
                    if (matches != null && matches.Length > 0)
                        closings.AddRange(matches);
                }

                parentWordProcessor.Document.EndUpdate();
                parentWordProcessor.SaveDocument(targetFile, DocumentFormat.OpenXml);

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