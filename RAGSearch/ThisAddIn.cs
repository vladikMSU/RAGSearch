using System;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

namespace RAGSearch
{
    public partial class ThisAddIn
    {
        private LocalServiceClient serviceClient;
        private OomGuardProbe oomGuardProbe;
        private NativeImportRunner nativeImportRunner;
        private NativeSearchPresenter nativeSearchPresenter;
        private SearchPaneControl searchPaneControl;
        private Microsoft.Office.Tools.CustomTaskPane searchTaskPane;
        private Outlook.Explorer taskPaneExplorer;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            StartupTrace.Step("BEGIN ThisAddIn_Startup");
            try
            {
                StartupTrace.Step("BEGIN LocalServiceClient constructor (loopback HTTP only)");
                serviceClient = new LocalServiceClient();
                StartupTrace.Step("END LocalServiceClient constructor");

                oomGuardProbe = new OomGuardProbe(Application);
                nativeImportRunner = new NativeImportRunner(serviceClient);
                StartupTrace.Step("constructed native runner and explicit OOM diagnostic");

                StartupTrace.Step("BEGIN Application.ActiveExplorer (window lookup; no mail/address getters)");
                taskPaneExplorer = Application.ActiveExplorer();
                StartupTrace.Step("END Application.ActiveExplorer; found=" + (taskPaneExplorer != null));
                if (taskPaneExplorer != null)
                {
                    nativeSearchPresenter = new NativeSearchPresenter(taskPaneExplorer);
                    StartupTrace.Step("constructed native view presenter bound to task-pane Explorer");

                    StartupTrace.Step("BEGIN SearchPaneControl constructor (WinForms only)");
                    searchPaneControl = new SearchPaneControl(
                        serviceClient,
                        oomGuardProbe,
                        nativeImportRunner,
                        nativeSearchPresenter.GetCurrentScope,
                        nativeSearchPresenter.Show,
                        nativeSearchPresenter.Clear);
                    StartupTrace.Step("END SearchPaneControl constructor");

                    StartupTrace.Step("BEGIN CustomTaskPanes.Add/show");
                    searchTaskPane = CustomTaskPanes.Add(
                        searchPaneControl,
                        "RAG Search",
                        taskPaneExplorer);
                    searchTaskPane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionTop;
                    searchTaskPane.Height = 126;
                    searchTaskPane.Visible = true;
                    StartupTrace.Step("END CustomTaskPanes.Add/show");
                }

                StartupTrace.Step("END ThisAddIn_Startup; production startup performed no mail/address OOM reads");
            }
            catch (Exception ex)
            {
                StartupTrace.Failure("ThisAddIn_Startup", ex);
                try
                {
                    if (searchPaneControl != null)
                    {
                        searchPaneControl.Dispose();
                    }
                }
                catch (Exception cleanupException)
                {
                    StartupTrace.Failure(
                        "ThisAddIn_Startup SearchPaneControl cleanup",
                        cleanupException);
                }
                finally
                {
                    searchPaneControl = null;
                }
                try
                {
                    if (nativeSearchPresenter != null)
                    {
                        nativeSearchPresenter.Dispose();
                    }
                }
                catch (Exception cleanupException)
                {
                    StartupTrace.Failure(
                        "ThisAddIn_Startup NativeSearchPresenter cleanup",
                        cleanupException);
                }
                finally
                {
                    nativeSearchPresenter = null;
                    ComRelease.Final(taskPaneExplorer);
                    taskPaneExplorer = null;
                }
                throw;
            }

            // Production ingestion is the explicit Extended MAPI child process.
            // The labelled diagnostic button performs one deliberate protected OOM
            // getter; there is no Outlook Object Model indexing path in the add-in.
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            if (searchTaskPane != null)
            {
                searchTaskPane.Visible = false;
                CustomTaskPanes.Remove(searchTaskPane);
                searchTaskPane = null;
            }

            if (searchPaneControl != null)
            {
                searchPaneControl.Dispose();
                searchPaneControl = null;
            }

            if (nativeSearchPresenter != null)
            {
                nativeSearchPresenter.Dispose();
                nativeSearchPresenter = null;
            }

            oomGuardProbe = null;

            if (nativeImportRunner != null)
            {
                nativeImportRunner.Dispose();
                nativeImportRunner = null;
            }

            if (serviceClient != null)
            {
                serviceClient.Dispose();
                serviceClient = null;
            }

            ComRelease.Final(taskPaneExplorer);
            taskPaneExplorer = null;

            // Outlook no longer raises this event during shutdown. See:
            // https://go.microsoft.com/fwlink/?LinkId=506785
        }

        #region VSTO generated startup hook

        private void InternalStartup()
        {
            StartupTrace.BeginSession();
            StartupTrace.Step("BEGIN InternalStartup event registration");
            Startup += new EventHandler(ThisAddIn_Startup);
            Shutdown += new EventHandler(ThisAddIn_Shutdown);
            StartupTrace.Step("END InternalStartup event registration");
        }

        #endregion
    }
}
