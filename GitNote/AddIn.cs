using Extensibility;
using Gitmanik.GitNote.Utilities;
using Microsoft.Office.Core;
using System;
using System.Diagnostics.CodeAnalysis;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Threading.Tasks;
using System.Windows.Forms;
using Application = Microsoft.Office.Interop.OneNote.Application;

namespace Gitmanik.GitNote
{
	[ComVisible(true)]
	[Guid("7562BB0E-F7F7-4B01-A3BA-9FD8C5711F14"), ProgId("Gitmanik.GitNote")]
	public class AddIn : IDTExtensibility2, IRibbonExtensibility
	{
		protected Application OneNoteApplication { get; set; }

		private MainForm mainForm;

		public AddIn()
		{
		}

		/// <summary>
		/// Returns the XML in Ribbon.xml so OneNote knows how to render our ribbon
		/// </summary>
		/// <param name="RibbonID"></param>
		/// <returns></returns>
		public string GetCustomUI(string RibbonID)
		{
			return Properties.Resources.ribbon;
		}

		public void OnAddInsUpdate(ref Array custom)
		{
		}

		/// <summary>
		/// Cleanup
		/// </summary>
		/// <param name="custom"></param>
		public void OnBeginShutdown(ref Array custom)
		{
		}

		/// <summary>
		/// Called upon startup.
		/// Keeps a reference to the current OneNote application object.
		/// </summary>
		/// <param name="application"></param>
		/// <param name="connectMode"></param>
		/// <param name="addInInst"></param>
		/// <param name="custom"></param>
		public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
		{
			OneNoteApplication = ((Application)Application);
		}

		[SuppressMessage("Microsoft.Reliability", "CA2001:AvoidCallingProblematicMethods", MessageId = "System.GC.Collect")]
		public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
		{
			OneNoteApplication = null;
			GC.Collect();
			GC.WaitForPendingFinalizers();
		}

		public void OnStartupComplete(ref Array custom)
		{
		}

		public async Task GitNoteButtonClicked(IRibbonControl control)
		{
			ShowForm();
			return;
		}

		public async Task GitNoteButtonClickedMath(IRibbonControl control)
		{
			micautLib.MathInputControl ctrl = new micautLib.MathInputControlClass();
			ctrl.EnableExtendedButtons(true);
			ctrl.SetCaptionText("ghitmenm");
			ctrl.Insert += Ctrl_Insert;
			ctrl.Close += () =>
			{
				ctrl.Hide();
			};
			ctrl.Show();
			return;
		}

		private void Ctrl_Insert(string RecoResult)
		{
			MessageBox.Show(RecoResult);
		}

		private void ShowForm()
		{
			mainForm = new MainForm(OneNoteApplication);
			System.Windows.Forms.Application.Run(mainForm);
			mainForm?.Invoke(new Action(() =>
			{
				mainForm?.Close();
				mainForm = null;
			}));
		}

		/// <summary>
		/// Specified in Ribbon.xml, this method returns the image to display on the ribbon button
		/// </summary>
		/// <param name="imageName"></param>
		/// <returns></returns>
		public IStream GetImage(string imageName)
		{
			MemoryStream imageStream = new MemoryStream();
			Properties.Resources.Logo.Save(imageStream, ImageFormat.Png);
			return new CCOMStreamWrapper(imageStream);
		}
	}
}