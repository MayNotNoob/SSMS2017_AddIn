using EnvDTE;
using EnvDTE80;
using Microsoft.SqlServer.Management.UI.VSIntegration;
using System;
using System.Text;
using Microsoft.SqlServer.Management.UI.VSIntegration.Editors;

namespace SSMS2017_AddIn
{
    public class DTEApplicationController
    {
        /// <summary>
        /// Method for creating and writting to a debug window
        /// </summary>
        /// <param name="application"></param>
        /// <param name="message"></param>
        public void WriteToOutputWindow(DTE2 application, string message)
        {
            try
            {
                EnvDTE.Events events = application.Events;
                OutputWindow outputWindow = (OutputWindow)application.Windows.Item(Constants.vsWindowKindOutput).Object;

                // Find the "Test Pane" Output window pane; if it doesn't exist,  
                // create it.
                OutputWindowPane pane = null;
                try
                {
                    pane = outputWindow.OutputWindowPanes.Item("Format SQL");
                }
                catch
                {
                    pane = outputWindow.OutputWindowPanes.Add("Format SQL");
                }

                // Show the Output window and activate the new pane.
                outputWindow.Parent.AutoHides = false;
                outputWindow.Parent.Activate();
                pane.Activate();

                // Add a line of text to the new pane.
                pane.OutputString(message + "\r\n");
            }
            catch
            {
                //MessageBox.Show(message, "T-SQL Tidy Error");
            }
        }

        /// <summary>
        /// GetContentsOfCurrentQueryWindow
        /// </summary>
        /// <returns></returns>
        public string GetContentsOfCurrentQueryWindow()
        {
            EnvDTE.TextDocument doc = (EnvDTE.TextDocument)ServiceCache.ExtensibilityModel.Application.ActiveDocument.Object(null);

            if (doc != null)
            {
                return GetSQLString(doc);
            }
            else
            {
                return null;
            }

        }

        /// <summary>
        /// Get Selected Text or All Text in window
        /// </summary>
        /// <param name="doc"></param>
        /// <returns></returns>
        public string GetSQLString(EnvDTE.TextDocument doc)
        {
            // Get Selected Text on screen
            String selectedparttest = doc.Selection.Text;


            if (selectedparttest.Length == 0)
            {
                // get all text in window
                doc.Selection.SelectAll();
                selectedparttest = doc.Selection.Text;
            }

            return selectedparttest;
        }

        /// <summary>
        /// CreateNewScriptWindow
        /// </summary>
        /// <param name="buffer"></param>
        public void CreateNewScriptWindow(StringBuilder buffer)
        {
            ServiceCache.ScriptFactory.CreateNewBlankScript(ScriptType.Sql);
            // insert SQL definition to document
            var doc = (TextDocument)ServiceCache.ExtensibilityModel.Application.ActiveDocument.Object(null);
            doc.EndPoint.CreateEditPoint().Insert(buffer.ToString());
        }

    }
}