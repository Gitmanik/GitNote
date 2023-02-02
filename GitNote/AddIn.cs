using Extensibility;
using Gitmanik.GitNote.Utilities;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.OneNote;
using NLog;
using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Gitmanik.GitNote
{
	[ComVisible(true)]
	[Guid("7562BB0E-F7F7-4B01-A3BA-9FD8C5711F14"), ProgId("Gitmanik.GitNote")]
	public class AddIn : IDTExtensibility2, IRibbonExtensibility
	{
#pragma warning disable CS3001 // Argument type is not CLS-compliant
#pragma warning disable CS3003 // Type is not CLS-compliant

		private static NLog.Logger Logger = NLog.LogManager.GetCurrentClassLogger();

		protected IApplication OneNoteApplication
		{ get; set; }

		protected micautLib.MathInputControl mathInputControl;

		public AddIn()
		{
			AppDomain.CurrentDomain.UnhandledException += CatchUnhandledException;

			NLog.Config.LoggingConfiguration nlog_config = new NLog.Config.LoggingConfiguration();
			NLog.Targets.FileTarget nlog_logfile = new NLog.Targets.FileTarget("logfile")
			{
				Layout = "${longdate}\t${level:uppercase=true}\t${logger}\t${message:withexception=true}",
				FileName = "F:/Logs.txt",
				ArchiveEvery = NLog.Targets.FileArchivePeriod.Day,
				MaxArchiveDays = 30,
				ArchiveNumbering = NLog.Targets.ArchiveNumberingMode.Date,
				ArchiveFileName = "F:/Logs.{##}.txt",
			};

			nlog_config.AddRule(LogLevel.Debug, LogLevel.Fatal, nlog_logfile);
			NLog.LogManager.Configuration = nlog_config;
		}

		private void CatchUnhandledException(object sender, UnhandledExceptionEventArgs e)
		{
			Logger.Fatal($"Unhandled Exception:\n{(Exception)e.ExceptionObject}");
		}

		public string GetCustomUI(string RibbonID)
		{
			return Properties.Resources.ribbon;
		}

		public void OnAddInsUpdate(ref Array custom)
		{
			Logger.Info("OnAddInsUpdate");
		}

		public void OnBeginShutdown(ref Array custom)
		{
			Logger.Info("OnBeginShutdown");
		}

		public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
		{
			Logger.Info("OnConnection. ConnectMode: {ConnectMode}");

			//SetOneNoteApplication((Application)Application);

			//mathInputControl = new micautLib.MathInputControlClass();
			//mathInputControl.EnableExtendedButtons(true);
			//mathInputControl.SetCaptionText("ghitmenm");
			//mathInputControl.Insert += (string res) =>
			//{
			//	MessageBox.Show(res);
			//};
			//mathInputControl.Close += () =>
			//{
			//	mathInputControl.Hide();
			//};
		}

		//public void SetOneNoteApplication(Application application)
		//{
		//	OneNoteApplication = application;
		//}

		public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
		{
			Logger.Info("OnDisconnection. RemoveMode: {RemoveMode}");
			OneNoteApplication = null;
			mathInputControl = null;
			Logger.Info("Finalized. Bye.");
			NLog.LogManager.Shutdown();
			GC.Collect();
			GC.WaitForPendingFinalizers();
		}

		public void OnStartupComplete(ref Array custom)
		{
			Logger.Info("OnStartupComplete");
			//GetOneNoteHandle();
		}

		//private void GetOneNoteHandle()
		//{
		//	int retries = 0;

		//	try
		//	{
		//		while (retries < 3)
		//		{
		//			try
		//			{
		//				OneNoteApplication = new Microsoft.Office.Interop.OneNote.Application();

		//				Logger.Info($"Completed successfully after {retries} retries");

		//				retries = int.MaxValue;
		//			}
		//			catch (COMException exc)
		//			{
		//				retries++;
		//				var ms = 250 * retries;

		//				Logger.Info($"{exc} OneNote is busy, retyring in {ms}ms");
		//				System.Threading.Thread.Sleep(ms);
		//			}
		//		}
		//	}
		//	catch (Exception exc)
		//	{
		//		Logger.Info($"{exc} error instantiating OneNote IApplication after {retries} retries", exc);
		//	}
		//}

		public async Task GitNoteButtonClicked(IRibbonControl control)
		{
			Logger.Info("GitNoteButtonClicked. control: {control.Id} {control.Tag} {control.Context}");
			await Task.Run(() => ShowForm());
		}

		public async Task GitNoteButtonClickedMath(IRibbonControl control)
		{
			Logger.Info("GitNoteButtonClickedMath. control: {control.Id} {control.Tag} {control.Context}");
			await Task.Run(() => mathInputControl.Show());
		}

		public async Task GitNoteButtonClickedInfo(IRibbonControl control)
		{
			Logger.Info("GitNoteButtonClickedInfo. control: {control.Id} {control.Tag} {sscontrol.Context}");
			await Task.Run(() => MessageBox.Show($"GitNote by gitmanik.dev\nPersonal OneNote AddIn\nLoaded DLL: {Assembly.GetExecutingAssembly().CodeBase}", "GitNote"));
		}

		private void ShowForm()
		{
			//Logger.Info("ShowForm");
			//mainForm = new MainForm(OneNoteApplication);
			//System.Windows.Forms.Application.Run(mainForm);
			//mainForm?.Invoke(new Action(() =>
			//{
			//	mainForm?.Close();
			//}));
		}

		public IStream GetImage(string imageName)
		{
			Logger.Info($"GetImage. imageName: {imageName}");
			MemoryStream imageStream = new MemoryStream();
			((Bitmap)Properties.Resources.ResourceManager.GetObject(imageName)).Save(imageStream, ImageFormat.Png);
			return new CCOMStreamWrapper(imageStream);
		}
	}
}