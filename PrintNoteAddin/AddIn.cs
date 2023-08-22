/*
 *  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
 */

using System;
using System.Diagnostics.CodeAnalysis;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using Extensibility;
using Microsoft.Office.Core;
using PrintNoteAddin.Utilities;
using Application = Microsoft.Office.Interop.OneNote.Application;  // Conflicts with System.Windows.Forms
using System.Reflection;
using System.Drawing;
using Microsoft.Office.Interop.OneNote;
using System.Text;
using System.Linq;
using System.Threading;
using System.Web;
using System.Configuration;
using System.Globalization;

#pragma warning disable CS3003 // Type is not CLS-compliant

namespace PrintNoteAddin
{
	[ComVisible(true)]
	[Guid("6ED07FCB-07F5-4AC4-AEFB-286DC51F9C17"), ProgId("PrintNote.Addin")]

	public class AddIn : IDTExtensibility2, IRibbonExtensibility
	{
		protected Application OneNoteApplication
		{ get; set; }

        public XNamespace ns;

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
            return LoadRibbon();
        }

        private string LoadRibbon()
        {
            try
            {
                var workingDirectory = Path.Combine(Path.GetDirectoryName(Assembly.GetCallingAssembly().Location), "ribbon.xml");
                string file = File.ReadAllText(workingDirectory);
                return file;
            }
            catch (Exception e)
            {
                MessageBox.Show("Exception from Addin.LoadRibbon:" + e.Message);
                return "";
            }
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
			SetOneNoteApplication((Application)Application);
		}

		public void SetOneNoteApplication(Application application)
		{
			OneNoteApplication = application;
		}

		/// <summary>
		/// Cleanup
		/// </summary>
		/// <param name="RemoveMode"></param>
		/// <param name="custom"></param>
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

        //public async Task AddInButtonClicked(IRibbonControl control)
        public void AddInButtonClicked(IRibbonControl control)
        {
            MessageBox.Show("This is a demo button!");
        }

        /// <summary>
        /// Specified in Ribbon.xml, this method returns the image to display on the ribbon button
        /// </summary>
        /// <param name="imageName"></param>
        /// <returns></returns>
        public IStream GetImage(string imageName)
		{
			MemoryStream imageStream = new MemoryStream();
            //switch (imageName)
            //{
            //    case "CSharp.png":
            //        Properties.Resources.CSharp.Save(imageStream, ImageFormat.Png);
            //        break;
            //    default:
            //        Properties.Resources.Logo.Save(imageStream, ImageFormat.Png);
            //        break;
            //}

            BindingFlags flags = BindingFlags.Static | BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic;

            var b = typeof(Properties.Resources).GetProperty(imageName.Substring(0, imageName.IndexOf('.')), flags).GetValue(null, null) as Bitmap;
            b.Save(imageStream, ImageFormat.Png);

            return new CCOMStreamWrapper(imageStream);
		}
    }
}
