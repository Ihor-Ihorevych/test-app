using QRCoder;
using Syncfusion.DocIO.DLS;
using System;
using System.Drawing;

namespace test_app_src
{
    public class WordHelper
    {
        #region Private fields
        private WordDocument _wordDocument;
        private string _filePath;
        #endregion
        #region Constructor
        public WordHelper(string filePath)
        {
            _filePath = filePath;
        }
        #endregion
        #region Methods
        public bool AddQrCodes()
        {
            if (_filePath == string.Empty)
                throw new ArgumentException("Path to .docx file can't be empty");
            try
            {
                _wordDocument = new WordDocument(_filePath);
                int current_page = 0;
                int pages_count = _wordDocument.BuiltinDocumentProperties.PageCount;
                QRCodeGenerator qrGenerator = new QRCodeGenerator();
                _wordDocument.LastSection.HeadersFooters.LinkToPrevious = false;
                string userName = Environment.UserName;
                foreach (WSection section in _wordDocument.Sections)
                {
                    WPicture qr = new WPicture(_wordDocument);
                    QRCodeData qrCodeData = qrGenerator.CreateQrCode($"{userName}, {++current_page} - {pages_count}", QRCodeGenerator.ECCLevel.Q);
                    QRCode qrCode = new QRCode(qrCodeData);
                    Bitmap qrCodeImage = qrCode.GetGraphic(100);
                    qr.LoadImage(qrCodeImage);
                    qr.Height = 120;
                    qr.Width = 120;
                    section.HeadersFooters.Footer.AddParagraph().ChildEntities.Add(qr);
                }
                return true;
            }
            catch
            {
                return false;
            }
        }
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
