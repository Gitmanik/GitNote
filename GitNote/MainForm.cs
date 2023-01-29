using Microsoft.Office.Interop.OneNote;
using System;
using System.Windows.Forms;
using Application = Microsoft.Office.Interop.OneNote.Application;

namespace Gitmanik.GitNote
{
	public partial class MainForm : Form
	{
		private Application onenoteApplication;

		public MainForm(Application application)
		{
			InitializeComponent();
			onenoteApplication = application;
			Load += MainForm_Load;
		}

		private void MainForm_Load(object sender, EventArgs e)
		{
			// Note: skipping error checking here, it's possible for there to not be a current page, for example.

			string pageId = onenoteApplication.Windows.CurrentWindow.CurrentPageId;
			string xml;
			onenoteApplication.GetPageContent(pageId, out xml);
			xmlTextBox.Text = xml;

			onenoteApplication.GetHierarchy(null, HierarchyScope.hsPages, out xml);
			hierarchyXmlTextBox.Text = xml;
		}
	}
}