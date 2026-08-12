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
        private OutlookMapiImportRunner outlookMapiImportRunner;
        private SearchPaneControl searchPaneControl;
        private Microsoft.Office.Tools.CustomTaskPane searchTaskPane;
        private Outlook.Explorer taskPaneExplorer;
        private int lastExpandedPaneHeight = DefaultExpandedPaneHeight;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            try
            {
                serviceClient = new LocalServiceClient();
                outlookMapiImportRunner = new OutlookMapiImportRunner(serviceClient);

                taskPaneExplorer = Application.ActiveExplorer();
                if (taskPaneExplorer != null)
                {
                    searchPaneControl = new SearchPaneControl(
                        serviceClient,
                        outlookMapiImportRunner,
                        OpenSearchResult,
                        SetSearchPaneCollapsed,
                        ToggleSearchPaneFloating);

                    searchTaskPane = CustomTaskPanes.Add(
                        searchPaneControl,
                        "            RAG Search",
                        taskPaneExplorer);
                    searchTaskPane.DockPositionChanged += SearchTaskPaneOnDockPositionChanged;
                    searchTaskPane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionBottom;
                    searchTaskPane.Height = DefaultExpandedPaneHeight;
                    searchTaskPane.Visible = true;
                    searchPaneControl.SetPaneFloating(false);
                }
            }
            catch
            {
                try
                {
                    if (searchTaskPane != null)
                    {
                        searchTaskPane.DockPositionChanged -= SearchTaskPaneOnDockPositionChanged;
                        CustomTaskPanes.Remove(searchTaskPane);
                    }
                }
                catch
                {
                    // Preserve the original startup exception if task-pane cleanup fails.
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
                catch
                {
                    // Preserve the original startup exception if control cleanup fails.
                }
                finally
                {
                    searchPaneControl = null;
                }
                ComRelease.Final(taskPaneExplorer);
                taskPaneExplorer = null;
                throw;
            }

            // Production ingestion is the explicit Extended MAPI child process;
            // there is no Outlook Object Model indexing path in the add-in.
        }

        private void OpenSearchResult(SearchResultDto result)
        {
            if (result == null)
            {
                throw new ArgumentNullException("result");
            }
            if (!result.IsOutlookDocument)
            {
                throw new InvalidOperationException(
                    "Локальный индекс не вернул Outlook locator этого письма.");
            }

            Outlook.NameSpace session = null;
            object outlookItem = null;
            try
            {
                // This callback runs directly from the WinForms double-click on
                // Outlook's UI/STA thread. Do not move this COM work to Task.Run.
                session = Application.Session;
                outlookItem = session.GetItemFromID(
                    result.LocatorEntryId,
                    result.LocatorStoreId);
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
            catch
            {
                // Best-effort rollback must not mask the original docking failure.
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

            if (outlookMapiImportRunner != null)
            {
                outlookMapiImportRunner.Dispose();
                outlookMapiImportRunner = null;
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
            Startup += new EventHandler(ThisAddIn_Startup);
            Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
