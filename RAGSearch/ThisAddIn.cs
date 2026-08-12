using System;
using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

namespace RAGSearch
{
    public partial class ThisAddIn
    {
        // CustomTaskPane dimensions are points, not WinForms device pixels.
        // 210 pt is approximately 280 px at 96 DPI.
        private const int DefaultExpandedPaneHeight = 210;
        private const int DefaultFloatingPaneWidth = 788;
        private const int DefaultFloatingPaneHeight = 240;

        private LocalServiceClient serviceClient;
        private OomGuardProbe oomGuardProbe;
        private NativeImportRunner nativeImportRunner;
        private SearchPaneControl searchPaneControl;
        private Microsoft.Office.Tools.CustomTaskPane searchTaskPane;
        private Outlook.Explorer taskPaneExplorer;
        private int lastExpandedPaneHeight = DefaultExpandedPaneHeight;

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
                    StartupTrace.Step("BEGIN SearchPaneControl constructor (WinForms only)");
                    searchPaneControl = new SearchPaneControl(
                        serviceClient,
                        oomGuardProbe,
                        nativeImportRunner,
                        OpenSearchResult,
                        SetSearchPaneCollapsed,
                        ToggleSearchPaneFloating);
                    StartupTrace.Step("END SearchPaneControl constructor");

                    StartupTrace.Step("BEGIN CustomTaskPanes.Add/show");
                    searchTaskPane = CustomTaskPanes.Add(
                        searchPaneControl,
                        "            RAG Search",
                        taskPaneExplorer);
                    searchTaskPane.DockPositionChanged += SearchTaskPaneOnDockPositionChanged;
                    searchTaskPane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionBottom;
                    searchTaskPane.Height = DefaultExpandedPaneHeight;
                    searchTaskPane.Visible = true;
                    searchPaneControl.SetPaneFloating(false);
                    StartupTrace.Step("END CustomTaskPanes.Add/show bottom result pane");
                }

                StartupTrace.Step("END ThisAddIn_Startup; production startup performed no mail/address OOM reads");
            }
            catch (Exception ex)
            {
                StartupTrace.Failure("ThisAddIn_Startup", ex);
                try
                {
                    if (searchTaskPane != null)
                    {
                        searchTaskPane.DockPositionChanged -= SearchTaskPaneOnDockPositionChanged;
                        CustomTaskPanes.Remove(searchTaskPane);
                    }
                }
                catch (Exception cleanupException)
                {
                    StartupTrace.Failure(
                        "ThisAddIn_Startup CustomTaskPane cleanup",
                        cleanupException);
                }
                finally
                {
                    searchTaskPane = null;
                }
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
                ComRelease.Final(taskPaneExplorer);
                taskPaneExplorer = null;
                throw;
            }

            // Production ingestion is the explicit Extended MAPI child process.
            // The labelled diagnostic button performs one deliberate protected OOM
            // getter; there is no Outlook Object Model indexing path in the add-in.
        }

        private void OpenSearchResult(SearchResultDto result)
        {
            if (result == null)
            {
                throw new ArgumentNullException("result");
            }
            if (string.IsNullOrWhiteSpace(result.entry_id) ||
                string.IsNullOrWhiteSpace(result.store_id))
            {
                throw new InvalidOperationException(
                    "В локальном индексе нет EntryID или StoreID этого письма.");
            }

            Outlook.NameSpace session = null;
            object outlookItem = null;
            try
            {
                // This callback runs directly from the WinForms double-click on
                // Outlook's UI/STA thread. Do not move this COM work to Task.Run.
                session = Application.Session;
                outlookItem = session.GetItemFromID(result.entry_id, result.store_id);
                var mailItem = outlookItem as Outlook.MailItem;
                if (mailItem == null)
                {
                    throw new InvalidOperationException(
                        "Найденный объект больше не является почтовым сообщением.");
                }
                mailItem.Display(false);
            }
            catch (COMException ex)
            {
                throw new InvalidOperationException(
                    "Outlook не нашёл письмо в подключённом хранилище.",
                    ex);
            }
            finally
            {
                ComRelease.Final(outlookItem);
                ComRelease.Final(session);
            }
        }

        private void SetSearchPaneCollapsed(bool collapsed)
        {
            if (searchTaskPane == null || searchPaneControl == null)
            {
                return;
            }
            if (searchTaskPane.DockPosition ==
                Office.MsoCTPDockPosition.msoCTPDockPositionFloating)
            {
                return;
            }
            if (searchTaskPane.DockPosition !=
                    Office.MsoCTPDockPosition.msoCTPDockPositionBottom &&
                searchTaskPane.DockPosition !=
                    Office.MsoCTPDockPosition.msoCTPDockPositionTop)
            {
                throw new InvalidOperationException(
                    "Панель можно свернуть только когда она закреплена сверху или снизу.");
            }

            var chromeHeight = Math.Max(
                0,
                searchTaskPane.Height - PixelsToPoints(searchPaneControl.ClientSize.Height));
            if (collapsed)
            {
                var currentHeight = searchTaskPane.Height;
                if (currentHeight <= 0)
                {
                    throw new InvalidOperationException(
                        "Outlook не вернул текущую высоту панели.");
                }
                lastExpandedPaneHeight = currentHeight;
                searchTaskPane.Height = chromeHeight +
                    PixelsToPoints(searchPaneControl.CollapsedClientHeight);
                return;
            }

            searchTaskPane.Height = Math.Max(
                lastExpandedPaneHeight,
                chromeHeight + PixelsToPoints(searchPaneControl.MinimumExpandedClientHeight));
        }

        private int PixelsToPoints(int pixels)
        {
            if (searchPaneControl == null)
            {
                return pixels;
            }
            var dpi = searchPaneControl.DeviceDpi <= 0
                ? 96
                : searchPaneControl.DeviceDpi;
            return Math.Max(1, (int)Math.Round(pixels * 72F / dpi));
        }

        private bool ToggleSearchPaneFloating()
        {
            if (searchTaskPane == null)
            {
                return false;
            }

            var previousDockPosition = searchTaskPane.DockPosition;
            var previousWidth = searchTaskPane.Width;
            var previousHeight = searchTaskPane.Height;
            if (previousDockPosition ==
                Office.MsoCTPDockPosition.msoCTPDockPositionFloating)
            {
                try
                {
                    searchTaskPane.DockPosition =
                        Office.MsoCTPDockPosition.msoCTPDockPositionBottom;
                    searchTaskPane.Height = lastExpandedPaneHeight;
                    return false;
                }
                catch
                {
                    RestorePanePosition(
                        previousDockPosition,
                        previousWidth,
                        previousHeight);
                    throw;
                }
            }

            if (previousHeight > 0)
            {
                lastExpandedPaneHeight = previousHeight;
            }
            try
            {
                searchTaskPane.DockPosition =
                    Office.MsoCTPDockPosition.msoCTPDockPositionFloating;
                searchTaskPane.Width = DefaultFloatingPaneWidth;
                searchTaskPane.Height = DefaultFloatingPaneHeight;
                return true;
            }
            catch
            {
                RestorePanePosition(
                    previousDockPosition,
                    previousWidth,
                    previousHeight);
                throw;
            }
        }

        private void RestorePanePosition(
            Office.MsoCTPDockPosition dockPosition,
            int width,
            int height)
        {
            try
            {
                searchTaskPane.DockPosition = dockPosition;
                if (dockPosition == Office.MsoCTPDockPosition.msoCTPDockPositionFloating)
                {
                    if (width > 0)
                    {
                        searchTaskPane.Width = width;
                    }
                    if (height > 0)
                    {
                        searchTaskPane.Height = height;
                    }
                }
                else if (dockPosition == Office.MsoCTPDockPosition.msoCTPDockPositionLeft ||
                         dockPosition == Office.MsoCTPDockPosition.msoCTPDockPositionRight)
                {
                    if (width > 0)
                    {
                        searchTaskPane.Width = width;
                    }
                }
                else if (height > 0)
                {
                    searchTaskPane.Height = height;
                }
            }
            catch (Exception rollbackException)
            {
                StartupTrace.Failure("RestorePanePosition", rollbackException);
            }
        }

        private void SearchTaskPaneOnDockPositionChanged(object sender, EventArgs eventArgs)
        {
            if (searchTaskPane == null || searchPaneControl == null)
            {
                return;
            }

            var floating = searchTaskPane.DockPosition ==
                           Office.MsoCTPDockPosition.msoCTPDockPositionFloating;
            searchPaneControl.SetPaneFloating(floating);
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            if (searchTaskPane != null)
            {
                searchTaskPane.DockPositionChanged -= SearchTaskPaneOnDockPositionChanged;
                searchTaskPane.Visible = false;
                CustomTaskPanes.Remove(searchTaskPane);
                searchTaskPane = null;
            }

            if (searchPaneControl != null)
            {
                searchPaneControl.Dispose();
                searchPaneControl = null;
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
