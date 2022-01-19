using QRCoder;
using Syncfusion.DocIO.DLS;
using System;
using System.Drawing;
using System.Threading.Tasks;

namespace test_app_src
{
    /// <summary>
    /// Helper class, to work with .docx file
    /// </summary>
    public class WordHelper
    {
        #region Private fields
        private WordDocument _wordDocument;
        private Aspose.Words.Document _wordDoc;
        private string _filePath;
        private string _userInfo;
        private string _tempPath;
        #endregion
        #region Constructor
        public WordHelper(string filePath, string userInfo)
        {
            _filePath = filePath;
            _userInfo = userInfo;
        }
        #endregion
        #region Methods
        /// <summary>
        /// Iterates over document sections, looking for section brakes
        /// </summary>
        /// <returns>Boolean (is file iterated successfully, or not)</returns>
        /// <exception cref="ArgumentException"></exception>
        public async Task<bool> AddQrCodes()
        {
            if (_filePath == string.Empty)
                throw new ArgumentException("Path to .docx file can't be empty");
            try
            {
                // Reading doc file
                _wordDoc = new Aspose.Words.Document(_filePath);
                _tempPath = $"tmp{Guid.NewGuid()}";
                int total_pages = _wordDoc.PageCount;
                // Splitting pages into individual files
                for(int i = 0; i < total_pages; i++)
                {
                    await Task.Run(() => {
                        var page = _wordDoc.ExtractPages(i, 1);
                        page.Save($"{_tempPath}/{i}.docx");
                    });
                    
                }
                // Initializing new document
                _wordDocument = new WordDocument();
                _wordDocument.MailMerge.RemoveEmptyParagraphs = true;
                QRCodeGenerator qrGenerator = new QRCodeGenerator();
                int current_page = 0;
                string ev_text = "Evaluation Only. Created with Aspose.Words. Copyright 2003-2022 Aspose Pty Ltd.";
                for (int i = 0; i < total_pages; i++)
                {
                    WordDocument tmp = new WordDocument($"{_tempPath}/{i}.docx");
                    // Replacing some bs
                    
                    string replaceWith = "";
                    tmp.MailMerge.RemoveEmptyParagraphs = true;
                    tmp.Replace(ev_text, replaceWith, false, true);
                    foreach (WSection section in tmp.Sections)
                    {
                        var clone = section.Clone();
                        foreach(WParagraph w in clone.Paragraphs)
                        {
                            w.ParagraphFormat.LineSpacing = 10f;
                            if (w.Text == string.Empty)
                            {
                                var index = clone.Paragraphs.IndexOf(w);
                                clone.Paragraphs.RemoveAt(index);
                            }
                        }
                        _wordDocument.Sections.Add(clone);
                    }
                }
                System.IO.Directory.Delete(_tempPath, true);
                foreach (WSection section in _wordDocument.Sections)
                {
                    section.HeadersFooters.LinkToPrevious = false;
                    // Add some gap between bottom and qr code
                    section.PageSetup.FooterDistance = 1;
                    section.PageSetup.DifferentFirstPage = false;
                    // Removing existing footers, if they existing
                    section.HeadersFooters.Footer.ChildEntities.Clear();
                    // Creating new picture for word
                    WPicture qr = new WPicture(_wordDocument);
                    // And generating qr code, using some word file data
                    QRCodeData qrCodeData = qrGenerator.CreateQrCode($"{_userInfo}, {++current_page} - {total_pages}", QRCodeGenerator.ECCLevel.Q);
                    QRCode qrCode = new QRCode(qrCodeData);
                    // Getting image
                    Bitmap qrCodeImage = qrCode.GetGraphic(100);
                    // Setting image to qr code
                    qr.LoadImage(qrCodeImage);
                    qr.Height = 70;
                    qr.Width = 70;
                    // Adding it to footer
                    section.HeadersFooters.Footer.AddParagraph().ChildEntities.Add(qr);
                }
                return true;
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show(ex.Message);
                return false;
            }
        }
        /// <summary>
        /// Saves the document to path, specified in parameter
        /// </summary>
        /// <param name="path">Path to file</param>
        /// <returns>Boolean value (is file saved successfully, or not)</returns>
        public bool Save(string path)
        {
            try
            {
                _wordDocument.Save(path);
                return true;
            }
            catch
            {
                return false;
            }
        }
        #endregion
    }
}
