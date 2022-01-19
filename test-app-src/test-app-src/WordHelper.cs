using QRCoder;
using Syncfusion.DocIO.DLS;
using System;
using System.Drawing;

namespace test_app_src
{
    /// <summary>
    /// Helper class, to work with .docx file
    /// </summary>
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
        /// <summary>
        /// Iterates over document sections, looking for section brakes
        /// </summary>
        /// <returns>Boolean (is file iterated successfully, or not)</returns>
        /// <exception cref="ArgumentException"></exception>
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
                    // Removing existing footers, if they existing
                    section.HeadersFooters.Footer.ChildEntities.Clear();
                    // Creating new picture for word
                    WPicture qr = new WPicture(_wordDocument);
                    // And generating qr code, using some word file data
                    QRCodeData qrCodeData = qrGenerator.CreateQrCode($"{userName}, {++current_page} - {pages_count}", QRCodeGenerator.ECCLevel.Q);
                    QRCode qrCode = new QRCode(qrCodeData);
                    // Getting image
                    Bitmap qrCodeImage = qrCode.GetGraphic(100);
                    // Setting image to qr code
                    qr.LoadImage(qrCodeImage);
                    qr.Height = 75;
                    qr.Width = 75;
                    // Adding it to footer
                    section.HeadersFooters.Footer.AddParagraph().ChildEntities.Add(qr);
                }
                return true;
            }
            catch
            {
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
