/*
 *  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
 */

using System;
using System.Diagnostics.CodeAnalysis;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Windows.Forms;
using System.Xml.Linq;
using Extensibility;
using Microsoft.Office.Core;
using PrintNoteAddin.Utilities;
using Application = Microsoft.Office.Interop.OneNote.Application;  // Conflicts with System.Windows.Forms
using System.Reflection;
using System.Drawing;
using Microsoft.Office.Interop.OneNote;
using System.Linq;
using static CreatePrintForm.CreatePrintForm;
using System.Collections.Generic;

#pragma warning disable CS3001 // Type is not CLS-compliant

namespace PrintNoteAddin
{
	[ComVisible(true)]
	[Guid("6ED07FCB-07F5-4AC4-AEFB-286DC51F9C17") /* {CLSID} */, ProgId("PrintNote.Addin")]

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
            string xmlHierarchy;
            OneNoteApplication.GetHierarchy(null, HierarchyScope.hsPages, out xmlHierarchy);
            var xdoc = XDocument.Parse(xmlHierarchy);
            ns = xdoc.Root.Name.Namespace;
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
            try
            {
                string path = "C:/Users/faisa/OneDrive/Desktop";

                // Get and parse page content 
                var pageId = OneNoteApplication.Windows.CurrentWindow.CurrentPageId;
                string xmlPage;
                OneNoteApplication.GetPageContent(pageId, out xmlPage);
                var page = XDocument.Parse(xmlPage);

                // Implementation of getting height & width by using attributes from <one:Position/> and <one:Size/>
                const float pxToMm = (float)0.2645833333; // Scalar constant which converts pixels to millimeters 
                float paperWidth = (float)215.9; // The width of the paper in millimeters
                List<float> heights = new List<float>();
                List<float> widths = new List<float>();
                foreach (var node in page.Descendants(ns + "Position"))
                {
                    var sizeNode = node.Ancestors().First().Descendants(ns + "Size").First();
                    (float x, float y) pos = (float.Parse(node.Attribute("x").Value), float.Parse(node.Attribute("y").Value));
                    (float x, float y) size = (float.Parse(sizeNode.Attribute("width").Value), float.Parse(sizeNode.Attribute("height").Value));
                    float nodeHeight = size.y + pos.y;
                    float nodeWidth = size.x + pos.x;
                    heights.Add(nodeHeight);
                    widths.Add(nodeWidth);
                }
                float mmHeight = heights.Max() * pxToMm;
                float mmWidth = widths.Max() * pxToMm;

                // Sets a minimum height of 11in
                if (mmHeight < 279.4 /* 11in - Letter height */)
                {
                    mmHeight = (float)279.4;
                }
                // OneNote scales content to paper width using a ratio, which can be used to obtain the new height
                if (mmWidth > 215.9 /* 8.5in - Letter width */)
                {
                    float ratio = (float)215.9 / (mmWidth - (float)12.7 /* .5in for safety */);
                    mmHeight = ratio * mmHeight;
                }
                // For content smaller than the paper width, a scalar constant is used to obtain the new height
                else
                {
                    float ratio = (float)1.4;
                    mmHeight = ratio * mmHeight;
                }

                // Needs administrative permisssions...
                // It may be possible to use the OneNote Publish feature with a IMsoDocExporter interface
                // The ideal solution would be a way to somehow parse the XML content
                AddCustomPaperSize("Microsoft Print to PDF", "PrintNote", (float)215.9, mmHeight);
                MessageBox.Show("Paper size set!");
            }
            catch (Exception e)
            {
                MessageBox.Show("Exception:\n" + e.Message);
            }
        }

        /// <summary>
        /// Specified in Ribbon.xml, this method returns the image to display on the ribbon button
        /// </summary>
        /// <param name="imageName"></param>
        /// <returns></returns>
        public IStream GetImage(string imageName)
		{
			MemoryStream imageStream = new MemoryStream();
            BindingFlags flags = BindingFlags.Static | BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic;

            var b = typeof(Properties.Resources).GetProperty(imageName.Substring(0, imageName.IndexOf('.')), flags).GetValue(null, null) as Bitmap;
            b.Save(imageStream, ImageFormat.Png);

            return new CCOMStreamWrapper(imageStream);
		}
    }
}
