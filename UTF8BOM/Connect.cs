using System;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Extensibility;
using EnvDTE;
using EnvDTE80;
namespace UTF8BOM
{
	/// <summary>추가 기능을 구현하는 개체입니다.</summary>
	/// <seealso class='IDTExtensibility2' />
	public class Connect : IDTExtensibility2
	{
		/// <summary>추가 기능 개체에 대한 생성자를 구현합니다. 이 메서드 안에 초기화 코드를 배치하십시오.</summary>
		public Connect()
		{
		}

		/// <summary>IDTExtensibility2 인터페이스의 OnConnection 메서드를 구현합니다. 추가 기능이 로드되고 있다는 알림 메시지를 받습니다.</summary>
		/// <param term='application'>호스트 응용 프로그램의 루트 개체입니다.</param>
		/// <param term='connectMode'>추가 기능이 로드되는 방법을 설명합니다.</param>
		/// <param term='addInInst'>이 추가 기능을 나타내는 개체입니다.</param>
		/// <seealso class='IDTExtensibility2' />
		public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
		{
			_applicationObject = (DTE2)application;
			_addInInstance = (AddIn)addInInst;
            _documentEvents = _applicationObject.Events.get_DocumentEvents(null);
            _documentEvents.DocumentSaved += new _dispDocumentEvents_DocumentSavedEventHandler(OnDocumentSaved);
		}

		/// <summary>IDTExtensibility2 인터페이스의 OnDisconnection 메서드를 구현합니다. 추가 기능이 언로드되고 있다는 알림 메시지를 받습니다.</summary>
		/// <param term='disconnectMode'>추가 기능이 언로드되는 방법을 설명합니다.</param>
		/// <param term='custom'>호스트 응용 프로그램과 관련된 매개 변수의 배열입니다.</param>
		/// <seealso class='IDTExtensibility2' />
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

		/// <summary>IDTExtensibility2 인터페이스의 OnAddInsUpdate 메서드를 구현합니다. 추가 기능의 컬렉션이 변경되면 알림 메시지를 받습니다.</summary>
		/// <param term='custom'>호스트 응용 프로그램과 관련된 매개 변수의 배열입니다.</param>
		/// <seealso class='IDTExtensibility2' />		
		public void OnAddInsUpdate(ref Array custom)
		{
		}

		/// <summary>IDTExtensibility2 인터페이스의 OnStartupComplete 메서드를 구현합니다. 호스트 응용 프로그램에서 로드가 완료되었다는 알림 메시지를 받습니다.</summary>
		/// <param term='custom'>호스트 응용 프로그램과 관련된 매개 변수의 배열입니다.</param>
		/// <seealso class='IDTExtensibility2' />
		public void OnStartupComplete(ref Array custom)
		{
		}

		/// <summary>IDTExtensibility2 인터페이스의 OnBeginShutdown 메서드를 구현합니다. 호스트 응용 프로그램이 언로드되고 있다는 알림 메시지를 받습니다.</summary>
		/// <param term='custom'>호스트 응용 프로그램과 관련된 매개 변수의 배열입니다.</param>
		/// <seealso class='IDTExtensibility2' />
		public void OnBeginShutdown(ref Array custom)
		{
		}
		
		private DTE2 _applicationObject;
		private AddIn _addInInstance;
        private DocumentEvents _documentEvents;
	}
}