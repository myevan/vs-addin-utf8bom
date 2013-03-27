using System;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Extensibility;
using EnvDTE;
using EnvDTE80;
namespace UTF8BOM
{
    public class Connect : IDTExtensibility2
    {
        public Connect()
        {
        }

        public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
        {
            _applicationObject = (DTE2)application;
            _addInInstance = (AddIn)addInInst;
            _documentEvents = _applicationObject.Events.get_DocumentEvents(null);
            _documentEvents.DocumentSaved += new _dispDocumentEvents_DocumentSavedEventHandler(OnDocumentSaved);
        }

        public void OnDisconnection(ext_DisconnectMode disconnectMode, ref Array custom)
        {
            _documentEvents.DocumentSaved -= new _dispDocumentEvents_DocumentSavedEventHandler(OnDocumentSaved);
        }

        static byte[] UTF8BOM = { 0xef, 0xbb, 0xbf };

        bool IsDosTextBytes(byte[] srcBytes)
        {
            for (int i = 0; i != srcBytes.Length; ++i)
            {
                if (i > 0 && srcBytes[i] == (byte)'\n')
                {
                    if (srcBytes[i - 1] == (byte)'\r')
                    {
                        return true;
                    }
                }
            }
            return false;
        }

        void OnDocumentSaved(Document document)
        {
            var path = document.FullName;
            var ifs = new FileStream(path, FileMode.Open);
            var hasUTF8BOM = ifs.ReadByte() == 0xef && ifs.ReadByte() == 0xbb && ifs.ReadByte() == 0xbf;
            ifs.Close();

            var encoding = System.Text.Encoding.UTF8;
            if (hasUTF8BOM)
            {
                byte[] srcBytes = File.ReadAllBytes(path);
                if (IsDosTextBytes(srcBytes))
                {
                    string srcText = encoding.GetString(srcBytes);
                    string dstText = srcText.Replace("\r\n", "\n").Replace("\t", "    ");

                    var ofs = new FileStream(path, FileMode.Create);

                    byte[] dstBytes = System.Text.Encoding.UTF8.GetBytes(dstText);
                    ofs.Write(dstBytes, 0, dstBytes.Length);                    

                    ofs.Close();
                }
            }
            else
            {
                byte[] srcBytes = File.ReadAllBytes(path);
                if (IsDosTextBytes(srcBytes))
                    encoding = System.Text.Encoding.Default;

                string srcText = encoding.GetString(srcBytes);
                string dstText = srcText.Replace("\r\n", "\n").Replace("\t", "    ");

                var ofs = new FileStream(path, FileMode.Create);
                ofs.Write(UTF8BOM, 0, 3);

                byte[] dstBytes = System.Text.Encoding.UTF8.GetBytes(dstText);
                ofs.Write(dstBytes, 0, dstBytes.Length);                    

                ofs.Close();

                MessageBox.Show("UTF8_CONVERTED:" + document.FullName);
            }

        }

        public void OnAddInsUpdate(ref Array custom)
        {
        }

        public void OnStartupComplete(ref Array custom)
        {
        }

        public void OnBeginShutdown(ref Array custom)
        {
        }

        private DTE2 _applicationObject;
        private AddIn _addInInstance;
        private DocumentEvents _documentEvents;
    }
}
