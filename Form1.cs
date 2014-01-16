using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Runtime.InteropServices;
using System.Text;


/*
 * Majority of this program was implemented by Tom Archer
 * Source of original project can be found here: http://www.codeguru.com/columns/dotnettips/article.php/c7315/Monitoring-Clipboard-Activity-in-C.htm
*/

namespace ClipboardMonitor
{
	/// <summary>
	/// Summary description for Form1.
	/// </summary>
	public class Form1 : System.Windows.Forms.Form
	{
		[DllImport("User32.dll")]
		protected static extern int SetClipboardViewer(int hWndNewViewer);

		[DllImport("User32.dll", CharSet=CharSet.Auto)]
		public static extern bool ChangeClipboardChain(IntPtr hWndRemove, IntPtr hWndNewNext);

		[DllImport("user32.dll", CharSet=CharSet.Auto)]
		public static extern int SendMessage(IntPtr hwnd, int wMsg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        internal static extern bool OpenClipboard(IntPtr hWndNewOwner);

        [DllImport("user32.dll")]
        internal static extern bool CloseClipboard();

        [DllImport("user32.dll")]
        internal static extern bool SetClipboardData(uint uFormat, IntPtr data);

        [DllImport("USER32.DLL", CharSet=CharSet.Auto, SetLastError=true)]
        public static extern uint RegisterClipboardFormat(string format);

		private System.Windows.Forms.RichTextBox richTextBox1;

		IntPtr nextClipboardViewer;

		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public Form1()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();
			nextClipboardViewer = (IntPtr)SetClipboardViewer((int) this.Handle);

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			ChangeClipboardChain(this.Handle, nextClipboardViewer);
			if( disposing )
			{
				if (components != null) 
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			this.richTextBox1 = new System.Windows.Forms.RichTextBox();
			this.SuspendLayout();
			// 
			// richTextBox1
			// 
			this.richTextBox1.BorderStyle = System.Windows.Forms.BorderStyle.None;
			this.richTextBox1.Dock = System.Windows.Forms.DockStyle.Fill;
			this.richTextBox1.Location = new System.Drawing.Point(0, 0);
			this.richTextBox1.Name = "richTextBox1";
			this.richTextBox1.ReadOnly = true;
			this.richTextBox1.Size = new System.Drawing.Size(292, 273);
			this.richTextBox1.TabIndex = 0;
			this.richTextBox1.Text = "richTextBox1";
			this.richTextBox1.WordWrap = false;
			// 
			// Form1
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(292, 273);
			this.Controls.Add(this.richTextBox1);
			this.Name = "Form1";
			this.Text = "Clipboard Monitor Example";
			this.ResumeLayout(false);

		}
		#endregion

		/// <summary>
		/// The main entry point for the application.
		/// </summary>
		[STAThread]
		static void Main() 
		{
			Application.Run(new Form1());
		}

		protected override void WndProc(ref System.Windows.Forms.Message m)
		{
			// defined in winuser.h
			const int WM_DRAWCLIPBOARD = 0x308;
			const int WM_CHANGECBCHAIN = 0x030D;

			switch(m.Msg)
			{
				case WM_DRAWCLIPBOARD:
					DisplayClipboardData();
					SendMessage(nextClipboardViewer, m.Msg, m.WParam, m.LParam);
					break;

				case WM_CHANGECBCHAIN:
					if (m.WParam == nextClipboardViewer)
						nextClipboardViewer = m.LParam;
					else
						SendMessage(nextClipboardViewer, m.Msg, m.WParam, m.LParam);
					break;

				default:
					base.WndProc(ref m);
					break;
			}	
		}

		void DisplayClipboardData()		
		{
			try
			{
				IDataObject iData = new DataObject();  
				iData = Clipboard.GetDataObject();

                if (iData.GetDataPresent(DataFormats.Text))
                {
                    //string text = Clipboard.GetText(TextDataFormat.UnicodeText);
                    string data = ((string)iData.GetData(DataFormats.Text)).Replace("\n", "\r\n");
                    if (data.StartsWith("FMT"))
                    {
                        // change to byteArray, allocate global memory, copy byteArray to this memory
                        byte[] dataMemory = Encoding.UTF8.GetBytes(data);
                        OpenClipboard(IntPtr.Zero);
                        IntPtr ptr = Marshal.AllocHGlobal(dataMemory.Length);
                        Marshal.Copy(dataMemory, 0, ptr, dataMemory.Length);
                        uint formatId = RegisterClipboardFormat("ALEPH_DOC");
                        bool success = SetClipboardData(formatId, ptr);
                        CloseClipboard();

                        if (!success)
                        {
                            Marshal.FreeHGlobal(ptr);
                            richTextBox1.ForeColor = Color.Red;
                            richTextBox1.Text = "[Error occured during saving to clipboard]";
                        }
                        else
                        {
                            richTextBox1.ForeColor = Color.Green;
                            richTextBox1.Text = "[Clipboard text successfully converted]";
                        }
                    }
                    else
                    {
                        richTextBox1.ForeColor = Color.Black;
                        richTextBox1.Text = "[Clipboard text is not for ALEPH client]";
                    }
                }
                else
                {
                    richTextBox1.ForeColor = Color.Black;
                    richTextBox1.Text = "[Clipboard data is not RTF or ASCII Text]";
                }
			}
			catch(Exception e)
			{
				MessageBox.Show(e.ToString());
			}
		}
	}
}
