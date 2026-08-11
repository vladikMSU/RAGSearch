using System;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

namespace RAGSearch
{
    public partial class ThisAddIn
    {
        private LocalServiceClient serviceClient;
        private OutlookItemExtractor itemExtractor;
        private OutlookIndexer indexer;
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

                StartupTrace.Step("BEGIN OutlookItemExtractor constructor (filesystem spool only)");
                itemExtractor = new OutlookItemExtractor();
                StartupTrace.Step("END OutlookItemExtractor constructor");

                StartupTrace.Step("BEGIN OutlookIndexer constructor (OOM Session is deferred to Debug OOM click)");
                indexer = new OutlookIndexer(Application, serviceClient, itemExtractor);
                StartupTrace.Step("END OutlookIndexer constructor");

                oomGuardProbe = new OomGuardProbe(Application);
                nativeImportRunner = new NativeImportRunner(serviceClient);
                nativeSearchPresenter = new NativeSearchPresenter(Application);
                StartupTrace.Step("constructed native runner, lazy OOM helpers and search presenter");

                StartupTrace.Step("BEGIN SearchPaneControl constructor (WinForms only)");
                searchPaneControl = new SearchPaneControl(
                    serviceClient,
                    indexer,
                    oomGuardProbe,
                    nativeImportRunner,
                    nativeSearchPresenter.Show,
                    nativeSearchPresenter.Clear);
                StartupTrace.Step("END SearchPaneControl constructor");

                StartupTrace.Step("BEGIN Application.ActiveExplorer (window lookup; no mail/address getters)");
                taskPaneExplorer = Application.ActiveExplorer();
                StartupTrace.Step("END Application.ActiveExplorer; found=" + (taskPaneExplorer != null));
                if (taskPaneExplorer != null)
                {
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
                throw;
            }

            // Do not subscribe NewMailEx to the Outlook Object Model extractor.
            // Outlook can replay that event while the add-in is still starting,
            // which would trigger Object Model Guard before the user does anything.
            // Production ingestion is the explicit Extended MAPI child process;
            // the legacy OOM scanner is available only from the labelled debug button.
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

            if (indexer != null)
            {
                indexer.Dispose();
                indexer = null;
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
            nativeSearchPresenter = null;

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
