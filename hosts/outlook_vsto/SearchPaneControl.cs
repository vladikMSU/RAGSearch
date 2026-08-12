using System;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Win32;

namespace RAGSearch
{
    internal sealed class SearchPaneControl : UserControl
    {
        private const int SearchLimit = 25;
        private const int EmSetCueBanner = 0x1501;

        private readonly LocalServiceClient serviceClient;
        private readonly OutlookMapiImportRunner outlookMapiImportRunner;
        private readonly Action<SearchResultDto> openSearchResult;
        private readonly Action<bool> setPaneCollapsed;
        private readonly Func<bool> togglePaneFloating;
        private readonly PanePalette palette;
        private readonly TableLayoutPanel rootLayout;
        private readonly Panel commandBar;
        private readonly Label modeLabel;
        private readonly TextBox queryBox;
        private readonly Button searchButton;
        private readonly Button clearButton;
        private readonly Button collapseButton;
        private readonly Button detachButton;
        private readonly Button settingsButton;
        private readonly OutlookSafeDataGridView resultsGrid;
        private readonly Label statusLabel;
        private readonly ContextMenuStrip settingsMenu;
        private readonly ToolStripMenuItem collapsePaneMenuButton;
        private readonly ToolStripMenuItem detachPaneMenuButton;
        private readonly ToolStripMenuItem indexButton;
        private readonly ToolStripMenuItem stopButton;
        private readonly ToolStripMenuItem resetIndexButton;
        private readonly List<Font> ownedFonts = new List<Font>();
        private readonly Font subjectFont;
        private CancellationTokenSource searchCancellation;
        private CancellationTokenSource resetCancellation;
        private int searchGeneration;
        private string emptyStateText;
        private bool adjustingColumnWidths;
        private bool columnWidthsCustomized;
        private bool paneFloating;
        private bool paneCollapsed;
        private bool resetRunning;

        public SearchPaneControl(
            LocalServiceClient serviceClient,
            OutlookMapiImportRunner outlookMapiImportRunner,
            Action<SearchResultDto> openSearchResult,
            Action<bool> setPaneCollapsed,
            Func<bool> togglePaneFloating)
        {
            this.serviceClient = serviceClient ?? throw new ArgumentNullException("serviceClient");
            this.outlookMapiImportRunner = outlookMapiImportRunner ?? throw new ArgumentNullException("outlookMapiImportRunner");
            this.openSearchResult = openSearchResult ?? throw new ArgumentNullException("openSearchResult");
            this.setPaneCollapsed = setPaneCollapsed ?? throw new ArgumentNullException("setPaneCollapsed");
            this.togglePaneFloating = togglePaneFloating ?? throw new ArgumentNullException("togglePaneFloating");

            AutoScaleMode = AutoScaleMode.Dpi;
            palette = PanePalette.Create();
            Dock = DockStyle.Fill;
            BackColor = palette.Background;
            Font = OwnFont(new Font("Segoe UI", 9F));
            MinimumSize = new Size(420, 40);
            subjectFont = OwnFont(new Font("Segoe UI Semibold", 8.5F, FontStyle.Bold));
            emptyStateText = "Введите запрос, чтобы найти письма по смыслу.";

            queryBox = new TextBox
            {
                BackColor = palette.InputBackground,
                AutoSize = false,
                BorderStyle = BorderStyle.FixedSingle,
                Dock = DockStyle.None,
                Font = OwnFont(new Font("Segoe UI", 9.5F)),
                ForeColor = palette.Text,
                Margin = new Padding(0),
                TabIndex = 0
            };
            queryBox.HandleCreated += (sender, args) => SetCueBanner(
                queryBox,
                "Опишите, что обсуждали в письмах...");
            queryBox.KeyDown += QueryBoxOnKeyDown;

            searchButton = CreatePrimaryButton("Найти");
            searchButton.TabIndex = 1;
            searchButton.Click += SearchButtonOnClick;

            clearButton = CreateCommandButton("×");
            clearButton.AccessibleName = "Очистить запрос и результаты";
            clearButton.Font = OwnFont(new Font("Segoe UI", 12F));
            clearButton.TabIndex = 2;
            clearButton.Click += ClearButtonOnClick;

            collapseButton = CreateCommandButton("Свернуть");
            collapseButton.TabIndex = 3;
            collapseButton.Click += CollapseButtonOnClick;

            detachButton = CreateCommandButton("Отделить");
            detachButton.TabIndex = 4;
            detachButton.Click += DetachButtonOnClick;

            settingsButton = CreateCommandButton("⚙");
            settingsButton.AccessibleName = "Настройки RAG Search";
            settingsButton.Font = OwnFont(new Font("Segoe UI Symbol", 10F));
            settingsButton.MinimumSize = new Size(32, 26);
            settingsButton.TabIndex = 5;

            settingsMenu = new ContextMenuStrip
            {
                BackColor = palette.Surface,
                Font = OwnFont(new Font("Segoe UI", 9F)),
                ForeColor = palette.Text,
                ShowImageMargin = false
            };
            if (palette.IsDark)
            {
                settingsMenu.Renderer = new ToolStripProfessionalRenderer(
                    new DarkMenuColorTable(palette));
            }
            collapsePaneMenuButton = new ToolStripMenuItem("Свернуть панель");
            collapsePaneMenuButton.Click += CollapseButtonOnClick;
            detachPaneMenuButton = new ToolStripMenuItem("Отделить панель");
            detachPaneMenuButton.Click += DetachButtonOnClick;
            indexButton = new ToolStripMenuItem("Индексировать PST + OST (MAPI)");
            indexButton.Click += IndexButtonOnClick;
            stopButton = new ToolStripMenuItem("Остановить индексацию")
            {
                Enabled = false
            };
            stopButton.Click += StopButtonOnClick;
            resetIndexButton = new ToolStripMenuItem("Очистить локальный индекс...");
            resetIndexButton.Click += ResetIndexButtonOnClick;
            settingsMenu.Items.Add(collapsePaneMenuButton);
            settingsMenu.Items.Add(detachPaneMenuButton);
            settingsMenu.Items.Add(new ToolStripSeparator());
            settingsMenu.Items.Add(indexButton);
            settingsMenu.Items.Add(stopButton);
            settingsMenu.Items.Add(new ToolStripSeparator());
            settingsMenu.Items.Add(resetIndexButton);
            foreach (ToolStripItem item in settingsMenu.Items)
            {
                item.BackColor = palette.Surface;
                item.ForeColor = palette.Text;
            }
            settingsButton.Click += SettingsButtonOnClick;

            resultsGrid = CreateResultsGrid();
            resultsGrid.CellDoubleClick += ResultsGridOnCellDoubleClick;
            resultsGrid.CellPainting += ResultsGridOnCellPainting;
            resultsGrid.KeyDown += ResultsGridOnKeyDown;
            resultsGrid.Paint += ResultsGridOnPaint;
            resultsGrid.ClientSizeChanged += ResultsGridOnClientSizeChanged;
            resultsGrid.ColumnWidthChanged += ResultsGridOnColumnWidthChanged;
            resultsGrid.ScrollRowsRequested += ResultsGridOnScrollRowsRequested;

            statusLabel = new Label
            {
                AutoEllipsis = true,
                BackColor = palette.Background,
                Dock = DockStyle.Fill,
                Font = OwnFont(new Font("Segoe UI", 8.5F)),
                ForeColor = palette.MutedText,
                Padding = new Padding(8, 2, 4, 0),
                Text = "Локальный сервис: проверка...",
                TextAlign = ContentAlignment.MiddleLeft
            };

            commandBar = CreateCommandBar();
            modeLabel = CreateModeLabel("Письма по смыслу");
            commandBar.Controls.Add(modeLabel);
            commandBar.Controls.Add(queryBox);
            commandBar.Controls.Add(searchButton);
            commandBar.Controls.Add(clearButton);
            commandBar.Controls.Add(collapseButton);
            commandBar.Controls.Add(detachButton);
            commandBar.Controls.Add(settingsButton);
            commandBar.Layout += CommandBarOnLayout;

            rootLayout = new TableLayoutPanel
            {
                BackColor = palette.Background,
                ColumnCount = 1,
                Dock = DockStyle.Fill,
                Margin = new Padding(0),
                Padding = new Padding(52, 0, 0, 0),
                RowCount = 3
            };
            rootLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            rootLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 44F));
            rootLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            rootLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 24F));
            rootLayout.Controls.Add(commandBar, 0, 0);
            rootLayout.Controls.Add(resultsGrid, 0, 1);
            rootLayout.Controls.Add(statusLabel, 0, 2);
            Controls.Add(rootLayout);

            ApplyScaledMetrics();
            HandleCreated += (sender, args) => ApplyScaledMetrics();
            DpiChangedAfterParent += (sender, args) => ApplyScaledMetrics();

            outlookMapiImportRunner.ProgressChanged += OutlookMapiImportRunnerOnProgressChanged;
            Load += async (sender, args) =>
            {
                await RefreshHealthAsync();
            };
        }

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern IntPtr SendMessage(
            IntPtr windowHandle,
            int message,
            IntPtr wordParameter,
            string longParameter);

        private static void SetCueBanner(TextBox textBox, string cueText)
        {
            if (textBox == null || textBox.IsDisposed || !textBox.IsHandleCreated)
            {
                return;
            }

            SendMessage(textBox.Handle, EmSetCueBanner, IntPtr.Zero, cueText);
        }

        internal int CollapsedClientHeight
        {
            get { return ScaleLogical(44); }
        }

        internal int MinimumExpandedClientHeight
        {
            get { return ScaleLogical(220); }
        }

        internal void SetPaneFloating(bool floating)
        {
            paneFloating = floating;
            rootLayout.Padding = new Padding(floating ? 0 : ScaleLogical(52), 0, 0, 0);
            detachButton.Text = floating ? "Прикрепить" : "Отделить";
            detachPaneMenuButton.Text = floating
                ? "Прикрепить панель снизу"
                : "Отделить панель";
            commandBar.PerformLayout();
        }

        private Panel CreateCommandBar()
        {
            return new Panel
            {
                BackColor = palette.Surface,
                Dock = DockStyle.Fill,
                Margin = new Padding(0),
                Padding = new Padding(0)
            };
        }

        private Label CreateModeLabel(string text)
        {
            return new Label
            {
                AutoEllipsis = true,
                BackColor = palette.ChipBackground,
                BorderStyle = BorderStyle.FixedSingle,
                Font = OwnFont(new Font("Segoe UI Semibold", 8.5F, FontStyle.Bold)),
                ForeColor = palette.ChipText,
                Margin = new Padding(0),
                Text = text,
                TextAlign = ContentAlignment.MiddleCenter
            };
        }

        private Button CreateCommandButton(string text)
        {
            var button = new Button
            {
                AutoSize = false,
                BackColor = palette.ButtonBackground,
                FlatStyle = FlatStyle.Flat,
                ForeColor = palette.Text,
                Margin = new Padding(0),
                Padding = new Padding(0),
                Text = text,
                UseVisualStyleBackColor = false
            };
            button.FlatAppearance.BorderColor = palette.Border;
            button.FlatAppearance.MouseDownBackColor = palette.ButtonPressed;
            button.FlatAppearance.MouseOverBackColor = palette.ButtonHover;
            return button;
        }

        private Button CreatePrimaryButton(string text)
        {
            var button = CreateCommandButton(text);
            button.BackColor = palette.Accent;
            button.FlatAppearance.BorderColor = palette.AccentBorder;
            button.FlatAppearance.MouseDownBackColor = palette.AccentPressed;
            button.FlatAppearance.MouseOverBackColor = palette.AccentHover;
            button.ForeColor = Color.White;
            return button;
        }

        private OutlookSafeDataGridView CreateResultsGrid()
        {
            var grid = new OutlookSafeDataGridView
            {
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AllowUserToOrderColumns = false,
                AllowUserToResizeRows = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None,
                AutoGenerateColumns = false,
                BackgroundColor = palette.Background,
                BorderStyle = BorderStyle.None,
                CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal,
                ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithoutHeaderText,
                ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single,
                ColumnHeadersHeight = 25,
                ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing,
                Dock = DockStyle.Fill,
                EditMode = DataGridViewEditMode.EditProgrammatically,
                EnableHeadersVisualStyles = false,
                GridColor = palette.Border,
                Margin = new Padding(0),
                MultiSelect = false,
                ReadOnly = true,
                RowHeadersVisible = false,
                RowTemplate = { Height = 25 },
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                ShowCellErrors = false,
                ShowCellToolTips = true,
                ShowEditingIcon = false,
                // The stock DataGridView wheel path updates its child Win32
                // scrollbar with SetScrollInfo.  Inside an Outlook custom task
                // pane that synchronous call can deadlock Outlook's UI thread.
                // Keep the native scrollbar disabled; OutlookSafeDataGridView
                // consumes WM_MOUSEWHEEL and scrolls rows without forwarding the
                // message into that path.
                ScrollBars = ScrollBars.None
            };
            grid.ColumnHeadersDefaultCellStyle.BackColor = palette.HeaderBackground;
            grid.ColumnHeadersDefaultCellStyle.ForeColor = palette.Text;
            grid.ColumnHeadersDefaultCellStyle.Font = OwnFont(new Font("Segoe UI", 8.5F));
            grid.ColumnHeadersDefaultCellStyle.Padding = new Padding(5, 0, 5, 0);
            grid.DefaultCellStyle.BackColor = palette.Background;
            grid.DefaultCellStyle.ForeColor = palette.Text;
            grid.DefaultCellStyle.Font = OwnFont(new Font("Segoe UI", 8.5F));
            grid.DefaultCellStyle.Padding = new Padding(4, 0, 4, 0);
            grid.DefaultCellStyle.SelectionBackColor = palette.Selection;
            grid.DefaultCellStyle.SelectionForeColor = palette.SelectionText;
            grid.AlternatingRowsDefaultCellStyle.BackColor = palette.AlternateBackground;

            grid.Columns.Add(CreateTextColumn("Rank", "#", 36));
            grid.Columns["Rank"].Resizable = DataGridViewTriState.False;
            grid.Columns.Add(new DataGridViewTextBoxColumn
            {
                AutoSizeMode = DataGridViewAutoSizeColumnMode.None,
                HeaderText = "Тема письма и фрагмент",
                MinimumWidth = 220,
                Name = "SubjectAndSnippet",
                ReadOnly = true,
                SortMode = DataGridViewColumnSortMode.NotSortable,
                Width = 720
            });
            grid.Columns.Add(CreateTextColumn("Sender", "Отправитель", 150));
            grid.Columns.Add(CreateTextColumn("Received", "Дата", 118));
            var folderColumn = CreateTextColumn("Folder", "Хранилище / папка", 190);
            folderColumn.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            folderColumn.FillWeight = 100F;
            folderColumn.MinimumWidth = 190;
            folderColumn.Resizable = DataGridViewTriState.False;
            grid.Columns.Add(folderColumn);
            return grid;
        }

        private static DataGridViewTextBoxColumn CreateTextColumn(
            string name,
            string header,
            int width)
        {
            return new DataGridViewTextBoxColumn
            {
                HeaderText = header,
                Name = name,
                ReadOnly = true,
                SortMode = DataGridViewColumnSortMode.NotSortable,
                Width = width
            };
        }

        private void CommandBarOnLayout(object sender, LayoutEventArgs eventArgs)
        {
            var width = commandBar.ClientSize.Width;
            var height = commandBar.ClientSize.Height;
            if (width <= 0 || height <= 0)
            {
                return;
            }

            var padding = ScaleLogical(10);
            var gap = ScaleLogical(6);
            var controlHeight = ScaleLogical(28);
            var top = Math.Max(0, (height - controlHeight) / 2);
            var right = width - padding;

            var gearWidth = ScaleLogical(34);
            settingsButton.SetBounds(right - gearWidth, top, gearWidth, controlHeight);
            right = settingsButton.Left - gap;

            collapseButton.Visible = !paneFloating;
            var showUtilityActions = width >= ScaleLogical(720);
            var showTextActions = width >= ScaleLogical(1040);
            detachButton.Visible = showUtilityActions;
            if (showUtilityActions)
            {
                var utilityWidth = showTextActions ? ScaleLogical(88) : ScaleLogical(34);
                detachButton.Text = showTextActions
                    ? (paneFloating ? "Прикрепить" : "Отделить")
                    : (paneFloating ? "↙" : "↗");
                detachButton.AccessibleName = paneFloating
                    ? "Прикрепить панель снизу"
                    : "Отделить панель";
                detachButton.SetBounds(
                    right - utilityWidth,
                    top,
                    utilityWidth,
                    controlHeight);
                right = detachButton.Left - gap;

                if (!paneFloating)
                {
                    collapseButton.Text = showTextActions
                        ? (paneCollapsed ? "Развернуть" : "Свернуть")
                        : (paneCollapsed ? "⌄" : "—");
                    collapseButton.AccessibleName = paneCollapsed
                        ? "Развернуть панель"
                        : "Свернуть панель";
                    collapseButton.SetBounds(
                        right - utilityWidth,
                        top,
                        utilityWidth,
                        controlHeight);
                    right = collapseButton.Left - ScaleLogical(12);
                }
            }

            var left = padding;
            var showMode = width >= ScaleLogical(940);
            modeLabel.Visible = showMode;
            if (showMode)
            {
                var modeWidth = ScaleLogical(142);
                modeLabel.SetBounds(left, top, modeWidth, controlHeight);
                left = modeLabel.Right + ScaleLogical(12);
            }

            var searchWidth = ScaleLogical(76);
            var clearWidth = ScaleLogical(30);
            var maximumQueryWidth = ScaleLogical(700);
            var availableForQuery = right - left - searchWidth - clearWidth - (gap * 2);
            if (availableForQuery < ScaleLogical(180) && showMode)
            {
                modeLabel.Visible = false;
                left = padding;
                availableForQuery = right - left - searchWidth - clearWidth - (gap * 2);
            }

            var queryWidth = Math.Max(
                ScaleLogical(90),
                Math.Min(maximumQueryWidth, availableForQuery));
            queryBox.SetBounds(left, top, queryWidth, controlHeight);
            searchButton.SetBounds(
                queryBox.Right + gap,
                top,
                searchWidth,
                controlHeight);
            clearButton.SetBounds(
                searchButton.Right + gap,
                top,
                clearWidth,
                controlHeight);
        }

        private void ApplyScaledMetrics()
        {
            rootLayout.Padding = new Padding(
                paneFloating ? 0 : ScaleLogical(52),
                0,
                0,
                0);
            rootLayout.RowStyles[0].Height = ScaleLogical(44);
            rootLayout.RowStyles[2].Height = paneCollapsed ? 0 : ScaleLogical(24);
            resultsGrid.ColumnHeadersHeight = ScaleLogical(27);
            var rowHeight = ScaleLogical(27);
            resultsGrid.RowTemplate.Height = rowHeight;
            foreach (DataGridViewRow row in resultsGrid.Rows)
            {
                row.Height = rowHeight;
            }
            adjustingColumnWidths = true;
            try
            {
                resultsGrid.Columns["Rank"].Width = ScaleLogical(36);
                resultsGrid.Columns["SubjectAndSnippet"].MinimumWidth = ScaleLogical(220);
                resultsGrid.Columns["Sender"].MinimumWidth = ScaleLogical(90);
                resultsGrid.Columns["Sender"].Width = ScaleLogical(150);
                resultsGrid.Columns["Received"].MinimumWidth = ScaleLogical(90);
                resultsGrid.Columns["Received"].Width = ScaleLogical(118);
                resultsGrid.Columns["Folder"].MinimumWidth = ScaleLogical(190);
            }
            finally
            {
                adjustingColumnWidths = false;
            }
            FitColumnsToViewport(null);
            commandBar.PerformLayout();
            resultsGrid.Invalidate();
        }

        private void ResultsGridOnClientSizeChanged(object sender, EventArgs eventArgs)
        {
            FitColumnsToViewport(null);
        }

        private void ResultsGridOnScrollRowsRequested(
            object sender,
            ScrollRowsRequestedEventArgs eventArgs)
        {
            if (eventArgs == null ||
                eventArgs.WheelDetents == 0 ||
                resultsGrid.IsDisposed ||
                resultsGrid.ClientSize.Height <= resultsGrid.ColumnHeadersHeight ||
                resultsGrid.Rows.Count == 0)
            {
                return;
            }

            var dataHeight = Math.Max(
                1,
                resultsGrid.ClientSize.Height - resultsGrid.ColumnHeadersHeight);
            var rowHeight = Math.Max(1, resultsGrid.Rows[0].Height);
            var visibleRowCount = Math.Max(1, dataHeight / rowHeight);
            var maximumFirstRow = Math.Max(
                0,
                resultsGrid.Rows.Count - visibleRowCount);
            var currentFirstRow = resultsGrid.FirstDisplayedScrollingRowIndex;
            if (currentFirstRow < 0)
            {
                currentFirstRow = 0;
            }

            var configuredLines = SystemInformation.MouseWheelScrollLines;
            if (configuredLines == 0)
            {
                return;
            }
            var rowsPerDetent = configuredLines < 0
                ? visibleRowCount
                : configuredLines;
            var requestedFirstRow = (long)currentFirstRow -
                                    ((long)eventArgs.WheelDetents * rowsPerDetent);
            var targetFirstRow = (int)Math.Max(
                0L,
                Math.Min((long)maximumFirstRow, requestedFirstRow));
            if (targetFirstRow == currentFirstRow)
            {
                return;
            }

            resultsGrid.FirstDisplayedScrollingRowIndex = targetFirstRow;
        }

        private void ResultsGridOnColumnWidthChanged(
            object sender,
            DataGridViewColumnEventArgs eventArgs)
        {
            if (adjustingColumnWidths ||
                eventArgs.Column == null ||
                eventArgs.Column.AutoSizeMode == DataGridViewAutoSizeColumnMode.Fill ||
                eventArgs.Column.Name == "Rank")
            {
                return;
            }

            columnWidthsCustomized = true;
            FitColumnsToViewport(eventArgs.Column);
        }

        private void FitColumnsToViewport(DataGridViewColumn preferredColumn)
        {
            if (resultsGrid == null ||
                resultsGrid.IsDisposed ||
                resultsGrid.ClientSize.Width <= 0 ||
                adjustingColumnWidths)
            {
                return;
            }

            var rankColumn = resultsGrid.Columns["Rank"];
            var subjectColumn = resultsGrid.Columns["SubjectAndSnippet"];
            var senderColumn = resultsGrid.Columns["Sender"];
            var receivedColumn = resultsGrid.Columns["Received"];
            var folderColumn = resultsGrid.Columns["Folder"];
            var maximumFixedWidth = resultsGrid.ClientSize.Width -
                                    folderColumn.MinimumWidth -
                                    SystemInformation.VerticalScrollBarWidth -
                                    ScaleLogical(2);
            if (maximumFixedWidth <= 0)
            {
                return;
            }

            adjustingColumnWidths = true;
            try
            {
                if (!columnWidthsCustomized)
                {
                    senderColumn.Width = ScaleLogical(150);
                    receivedColumn.Width = ScaleLogical(118);
                    var maximumSubjectWidth = maximumFixedWidth -
                                              rankColumn.Width -
                                              senderColumn.Width -
                                              receivedColumn.Width;
                    subjectColumn.Width = Math.Max(
                        subjectColumn.MinimumWidth,
                        Math.Min(ScaleLogical(720), maximumSubjectWidth));
                }

                var fixedWidth = rankColumn.Width +
                                 subjectColumn.Width +
                                 senderColumn.Width +
                                 receivedColumn.Width;
                var overflow = fixedWidth - maximumFixedWidth;
                if (overflow > 0 && preferredColumn != null)
                {
                    overflow -= ReduceColumnWidth(preferredColumn, overflow);
                }
                if (overflow > 0)
                {
                    overflow -= ReduceColumnWidth(subjectColumn, overflow);
                }
                if (overflow > 0)
                {
                    overflow -= ReduceColumnWidth(senderColumn, overflow);
                }
                if (overflow > 0)
                {
                    ReduceColumnWidth(receivedColumn, overflow);
                }
            }
            finally
            {
                adjustingColumnWidths = false;
            }
        }

        private static int ReduceColumnWidth(DataGridViewColumn column, int amount)
        {
            if (column == null || amount <= 0 || column.AutoSizeMode == DataGridViewAutoSizeColumnMode.Fill)
            {
                return 0;
            }

            var reducible = Math.Max(0, column.Width - column.MinimumWidth);
            var reduction = Math.Min(amount, reducible);
            if (reduction > 0)
            {
                column.Width -= reduction;
            }
            return reduction;
        }

        private int ScaleLogical(int value)
        {
            var dpi = DeviceDpi <= 0 ? 96 : DeviceDpi;
            return Math.Max(1, (int)Math.Round(value * dpi / 96F));
        }

        private void ResultsGridOnCellPainting(
            object sender,
            DataGridViewCellPaintingEventArgs eventArgs)
        {
            if (eventArgs.RowIndex < 0 ||
                eventArgs.ColumnIndex != resultsGrid.Columns["SubjectAndSnippet"].Index)
            {
                return;
            }

            var result = resultsGrid.Rows[eventArgs.RowIndex].Tag as SearchResultDto;
            if (result == null)
            {
                return;
            }

            eventArgs.Paint(
                eventArgs.ClipBounds,
                DataGridViewPaintParts.Background |
                DataGridViewPaintParts.SelectionBackground |
                DataGridViewPaintParts.Border |
                DataGridViewPaintParts.Focus);

            var selected = resultsGrid.Rows[eventArgs.RowIndex].Selected;
            var textColor = selected ? palette.SelectionText : palette.Text;
            var mutedColor = selected ? palette.SelectionMutedText : palette.MutedText;
            var bounds = Rectangle.Inflate(eventArgs.CellBounds, -ScaleLogical(7), 0);
            var subject = CleanInline(result.Subject);
            var snippet = BuildVisibleSnippet(result);
            if (subject.Length == 0)
            {
                subject = "(без темы)";
            }

            const TextFormatFlags flags = TextFormatFlags.EndEllipsis |
                                          TextFormatFlags.NoPadding |
                                          TextFormatFlags.NoPrefix |
                                          TextFormatFlags.SingleLine |
                                          TextFormatFlags.VerticalCenter;
            var measuredSubject = TextRenderer.MeasureText(
                eventArgs.Graphics,
                subject,
                subjectFont,
                new Size(int.MaxValue, bounds.Height),
                flags).Width;
            var subjectWidth = snippet.Length == 0
                ? bounds.Width
                : Math.Min(measuredSubject, Math.Max(1, bounds.Width / 2));
            var subjectBounds = new Rectangle(
                bounds.Left,
                bounds.Top,
                Math.Max(1, subjectWidth),
                bounds.Height);
            TextRenderer.DrawText(
                eventArgs.Graphics,
                subject,
                subjectFont,
                subjectBounds,
                textColor,
                flags);

            if (snippet.Length > 0 && subjectBounds.Right < bounds.Right)
            {
                var snippetText = "  —  " + snippet;
                var snippetBounds = new Rectangle(
                    subjectBounds.Right,
                    bounds.Top,
                    bounds.Right - subjectBounds.Right,
                    bounds.Height);
                TextRenderer.DrawText(
                    eventArgs.Graphics,
                    snippetText,
                    resultsGrid.DefaultCellStyle.Font,
                    snippetBounds,
                    mutedColor,
                    flags);
            }

            eventArgs.Handled = true;
        }

        private void ResultsGridOnPaint(object sender, PaintEventArgs eventArgs)
        {
            if (resultsGrid.Rows.Count != 0 || string.IsNullOrWhiteSpace(emptyStateText))
            {
                return;
            }

            var bounds = new Rectangle(
                ScaleLogical(12),
                resultsGrid.ColumnHeadersHeight + ScaleLogical(8),
                Math.Max(1, resultsGrid.ClientSize.Width - ScaleLogical(24)),
                Math.Max(
                    ScaleLogical(24),
                    resultsGrid.ClientSize.Height -
                    resultsGrid.ColumnHeadersHeight -
                    ScaleLogical(16)));
            TextRenderer.DrawText(
                eventArgs.Graphics,
                emptyStateText,
                Font,
                bounds,
                palette.MutedText,
                TextFormatFlags.HorizontalCenter |
                TextFormatFlags.VerticalCenter |
                TextFormatFlags.EndEllipsis |
                TextFormatFlags.NoPrefix);
        }

        private async void SearchButtonOnClick(object sender, EventArgs eventArgs)
        {
            await SearchAsync();
        }

        private async Task SearchAsync()
        {
            if (resetRunning)
            {
                return;
            }
            if (outlookMapiImportRunner.IsRunning)
            {
                statusLabel.Text = "Дождитесь завершения индексации или остановите её через ⚙.";
                return;
            }

            var query = queryBox.Text.Trim();
            if (query.Length == 0)
            {
                statusLabel.Text = "Введите запрос — можно описать смысл своими словами.";
                queryBox.Focus();
                return;
            }

            if (searchCancellation != null)
            {
                searchCancellation.Cancel();
            }

            var currentCancellation = new CancellationTokenSource();
            var currentGeneration = ++searchGeneration;
            searchCancellation = currentCancellation;
            queryBox.Enabled = false;
            searchButton.Enabled = false;
            resultsGrid.Cursor = Cursors.WaitCursor;
            UpdateResetIndexButtonEnabled();
            emptyStateText = "Ищу письма...";
            resultsGrid.Invalidate();
            statusLabel.Text = "Проверяю локальный сервис...";
            try
            {
                await outlookMapiImportRunner.EnsureServiceReadyAsync(currentCancellation.Token);
                currentCancellation.Token.ThrowIfCancellationRequested();
                statusLabel.Text = "Ищу по смыслу и точным совпадениям...";
                var response = await serviceClient.SearchAsync(
                    query,
                    SearchLimit,
                    currentCancellation.Token);
                currentCancellation.Token.ThrowIfCancellationRequested();
                if (currentGeneration != searchGeneration)
                {
                    throw new OperationCanceledException(currentCancellation.Token);
                }

                if (response == null || response.Results == null)
                {
                    throw new InvalidOperationException(
                        "Локальный сервис вернул ответ без обязательного списка results.");
                }
                var results = response.Results;
                PopulateResults(results);
                statusLabel.Text = FormatSearchStatus(response, results.Count);
            }
            catch (OperationCanceledException)
            {
                if (currentGeneration == searchGeneration)
                {
                    statusLabel.Text = "Поиск отменён.";
                    if (resultsGrid.Rows.Count == 0)
                    {
                        emptyStateText = "Поиск отменён.";
                        resultsGrid.Invalidate();
                    }
                }
            }
            catch (Exception ex)
            {
                if (currentGeneration == searchGeneration)
                {
                    statusLabel.Text = "Поиск не выполнен: " + ex.Message;
                    if (resultsGrid.Rows.Count == 0)
                    {
                        emptyStateText = "Не удалось выполнить поиск.";
                        resultsGrid.Invalidate();
                    }
                }
            }
            finally
            {
                if (ReferenceEquals(searchCancellation, currentCancellation))
                {
                    searchCancellation = null;
                    RunOnUi(() =>
                    {
                        queryBox.Enabled = !resetRunning;
                        searchButton.Enabled = !resetRunning;
                        resultsGrid.Cursor = Cursors.Default;
                        UpdateResetIndexButtonEnabled();
                    });
                }
                currentCancellation.Dispose();
            }
        }

        private void PopulateResults(IList<SearchResultDto> results)
        {
            resultsGrid.SuspendLayout();
            try
            {
                resultsGrid.Rows.Clear();
                for (var index = 0; index < results.Count; index++)
                {
                    var result = results[index];
                    if (result == null || result.Rank <= 0)
                    {
                        throw new InvalidOperationException(
                            "Локальный сервис вернул некорректную строку результата.");
                    }

                    var rowIndex = resultsGrid.Rows.Add(
                        result.Rank,
                        BuildSubjectAndSnippet(result),
                        BuildSender(result),
                        result.ReceivedDisplay,
                        BuildFolder(result));
                    var row = resultsGrid.Rows[rowIndex];
                    row.Tag = result;
                    row.Cells[1].ToolTipText = BuildResultToolTip(result);
                }

                if (resultsGrid.Rows.Count > 0)
                {
                    resultsGrid.ClearSelection();
                    resultsGrid.Rows[0].Selected = true;
                    emptyStateText = string.Empty;
                }
                else
                {
                    emptyStateText = "Релевантных писем не найдено.";
                }
                resultsGrid.Invalidate();
            }
            finally
            {
                resultsGrid.ResumeLayout();
            }
        }

        private static string BuildSubjectAndSnippet(SearchResultDto result)
        {
            var subject = CleanInline(result.Subject);
            var snippet = BuildVisibleSnippet(result);
            if (subject.Length == 0)
            {
                subject = "(без темы)";
            }
            return snippet.Length == 0 ? subject : subject + " — " + snippet;
        }

        private static string BuildSender(SearchResultDto result)
        {
            var sender = CleanInline(result.SenderName);
            return sender.Length == 0 ? CleanInline(result.SenderEmail) : sender;
        }

        private static string BuildVisibleSnippet(SearchResultDto result)
        {
            var snippet = CleanInline(result.Snippet);
            if (snippet.Length == 0 || result.MatchedSources == null)
            {
                return snippet;
            }

            // A subject/sender/folder hit is represented by the service's metadata
            // chunk. Showing that raw chunk leaks implementation labels such as
            // "Conversation" into the result row and merely repeats the subject.
            if (result.MatchedSources.Contains("message_metadata") &&
                snippet.StartsWith("Subject:", StringComparison.OrdinalIgnoreCase))
            {
                return "совпадение в теме или реквизитах письма";
            }

            return snippet;
        }

        private static string BuildFolder(SearchResultDto result)
        {
            var store = CleanInline(result.StoreName);
            var folder = CleanInline(result.FolderPath);
            if (store.Length == 0)
            {
                return folder;
            }
            if (folder.Length == 0)
            {
                return store;
            }
            return store + " · " + folder.TrimStart('\\', '/');
        }

        private static string BuildResultToolTip(SearchResultDto result)
        {
            return string.Format(
                "{0}\r\n{1}\r\n{2}\r\nДвойной щелчок открывает исходное письмо Outlook.",
                CleanInline(result.Subject),
                BuildVisibleSnippet(result),
                BuildFolder(result));
        }

        private static string CleanInline(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }

            return value
                .Replace('\r', ' ')
                .Replace('\n', ' ')
                .Replace('\t', ' ')
                .Trim();
        }

        private Font OwnFont(Font font)
        {
            ownedFonts.Add(font);
            return font;
        }

        private static string FormatSearchStatus(SearchResponse response, int resultCount)
        {
            if (resultCount == 0)
            {
                if (string.Equals(response.Mode, "single-token-no-literal", StringComparison.Ordinal))
                {
                    return "Точных совпадений нет. Для поиска только по смыслу введите фразу из двух или более слов.";
                }
                return "Релевантных писем не найдено. Outlook остался в текущей папке.";
            }

            return string.Format(
                "Найдено писем: {0}. Порядок — релевантность RAG Search; двойной щелчок открывает исходное письмо Outlook.",
                resultCount);
        }

        private void ResultsGridOnCellDoubleClick(
            object sender,
            DataGridViewCellEventArgs eventArgs)
        {
            if (eventArgs.RowIndex < 0)
            {
                return;
            }
            OpenResult(resultsGrid.Rows[eventArgs.RowIndex]);
        }

        private void ResultsGridOnKeyDown(object sender, KeyEventArgs eventArgs)
        {
            if (eventArgs.KeyCode != Keys.Enter || resultsGrid.CurrentRow == null)
            {
                return;
            }

            eventArgs.Handled = true;
            eventArgs.SuppressKeyPress = true;
            OpenResult(resultsGrid.CurrentRow);
        }

        private void OpenResult(DataGridViewRow row)
        {
            var result = row == null ? null : row.Tag as SearchResultDto;
            if (result == null)
            {
                return;
            }

            try
            {
                openSearchResult(result);
                statusLabel.Text = "Письмо открыто в отдельном окне Outlook.";
            }
            catch (Exception ex)
            {
                statusLabel.Text =
                    "Не удалось открыть письмо: " + ex.Message +
                    " Если его переместили или удалили, обновите индекс.";
            }
        }

        private void ClearButtonOnClick(object sender, EventArgs eventArgs)
        {
            searchGeneration++;
            if (searchCancellation != null)
            {
                searchCancellation.Cancel();
            }
            resultsGrid.Rows.Clear();
            emptyStateText = "Введите запрос, чтобы найти письма по смыслу.";
            resultsGrid.Invalidate();
            queryBox.Clear();
            queryBox.Focus();
            statusLabel.Text = "Результаты очищены. Текущий вид Outlook не изменён.";
        }

        private void CollapseButtonOnClick(object sender, EventArgs eventArgs)
        {
            if (paneFloating)
            {
                statusLabel.Text = "Сначала прикрепите панель снизу, затем её можно свернуть.";
                return;
            }
            var previousCollapsed = paneCollapsed;
            try
            {
                var nextCollapsed = !previousCollapsed;
                setPaneCollapsed(nextCollapsed);
                ApplyCollapsedState(nextCollapsed);
            }
            catch (Exception ex)
            {
                ApplyCollapsedState(previousCollapsed);
                statusLabel.Text = "Не удалось изменить размер панели: " + ex.Message;
            }
        }

        private void ApplyCollapsedState(bool collapsed)
        {
            paneCollapsed = collapsed;
            resultsGrid.Visible = !collapsed;
            statusLabel.Visible = !collapsed;
            rootLayout.RowStyles[2].Height = collapsed ? 0 : ScaleLogical(24);
            collapsePaneMenuButton.Text = collapsed
                ? "Развернуть панель"
                : "Свернуть панель";
            commandBar.PerformLayout();
        }

        private void DetachButtonOnClick(object sender, EventArgs eventArgs)
        {
            var previousCollapsed = paneCollapsed;
            var previousFloating = paneFloating;
            try
            {
                if (paneCollapsed)
                {
                    setPaneCollapsed(false);
                    ApplyCollapsedState(false);
                }
                var floating = togglePaneFloating();
                SetPaneFloating(floating);
            }
            catch (Exception ex)
            {
                ApplyCollapsedState(previousCollapsed);
                SetPaneFloating(previousFloating);
                statusLabel.Text = "Не удалось изменить положение панели: " + ex.Message;
            }
        }

        private void SettingsButtonOnClick(object sender, EventArgs eventArgs)
        {
            settingsMenu.Show(
                settingsButton,
                new Point(
                    settingsButton.Width - settingsMenu.PreferredSize.Width,
                    settingsButton.Height));
        }

        private async void IndexButtonOnClick(object sender, EventArgs eventArgs)
        {
            if (resetRunning)
            {
                statusLabel.Text = "Дождитесь завершения очистки локального индекса.";
                return;
            }
            if (outlookMapiImportRunner.IsRunning)
            {
                statusLabel.Text = "Индексация уже выполняется.";
                return;
            }
            if (searchCancellation != null)
            {
                statusLabel.Text = "Дождитесь завершения поиска или очистите запрос.";
                return;
            }

            indexButton.Enabled = false;
            UpdateResetIndexButtonEnabled();
            stopButton.Enabled = true;
            statusLabel.Text = "Готовлю read-only Extended MAPI индексацию...";
            try
            {
                await outlookMapiImportRunner.RunAsync();
            }
            catch (OperationCanceledException)
            {
                RunOnUi(() => statusLabel.Text = "Outlook MAPI-индексация остановлена пользователем.");
            }
            catch (Exception ex)
            {
                var message = "Outlook MAPI-индексация не выполнена: " + ex.Message;
                RunOnUi(() => statusLabel.Text = message);
            }
            finally
            {
                RunOnUi(() =>
                {
                    indexButton.Enabled = !resetRunning;
                    stopButton.Enabled = false;
                    UpdateResetIndexButtonEnabled();
                });
            }
        }

        private void StopButtonOnClick(object sender, EventArgs eventArgs)
        {
            if (outlookMapiImportRunner.IsRunning)
            {
                outlookMapiImportRunner.RequestStop();
                stopButton.Enabled = false;
            }
        }

        private async void ResetIndexButtonOnClick(object sender, EventArgs eventArgs)
        {
            if (resetRunning)
            {
                return;
            }
            if (outlookMapiImportRunner.IsRunning)
            {
                statusLabel.Text = "Сначала остановите текущую индексацию.";
                return;
            }
            if (searchCancellation != null)
            {
                statusLabel.Text = "Дождитесь завершения поиска.";
                return;
            }
            var confirmation = MessageBox.Show(
                this,
                "Удалить весь локальный поисковый индекс RAGSearch?\r\n\r\n" +
                "Будут удалены только данные локальной базы: проиндексированные сообщения, " +
                "вложения и векторные чанки. Письма и вложения в Outlook, PST и OST " +
                "не изменятся.\r\n\r\nДля поиска потребуется заново выполнить индексацию.",
                "RAGSearch: очистка локального индекса",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning,
                MessageBoxDefaultButton.Button2);
            if (confirmation != DialogResult.Yes)
            {
                return;
            }

            // A modal dialog pumps messages. Recheck all activity immediately
            // before issuing the destructive local-service request.
            if (outlookMapiImportRunner.IsRunning || searchCancellation != null)
            {
                statusLabel.Text = "Очистка отменена: другая операция уже выполняется.";
                UpdateResetIndexButtonEnabled();
                return;
            }

            var currentCancellation = new CancellationTokenSource();
            resetCancellation = currentCancellation;
            resetRunning = true;
            queryBox.Enabled = false;
            searchButton.Enabled = false;
            clearButton.Enabled = false;
            indexButton.Enabled = false;
            stopButton.Enabled = false;
            resetIndexButton.Enabled = false;
            statusLabel.Text = "Проверяю локальный сервис перед очисткой индекса...";
            var completionStatus = "Очистка локального индекса завершилась без результата.";
            try
            {
                await outlookMapiImportRunner.EnsureServiceReadyAsync(currentCancellation.Token);
                currentCancellation.Token.ThrowIfCancellationRequested();
                var response = await serviceClient.ResetIndexAsync(currentCancellation.Token);
                currentCancellation.Token.ThrowIfCancellationRequested();
                if (response == null)
                {
                    throw new InvalidOperationException(
                        "Локальный сервис вернул ответ без обязательных счётчиков очистки.");
                }

                completionStatus = string.Format(
                    "Локальный индекс очищен: сообщений {0}, вложений {1}, чанков {2}.",
                    response.DeletedMessages,
                    response.DeletedAttachments,
                    response.DeletedChunks);
                resultsGrid.Rows.Clear();
                emptyStateText = "Индекс очищен. Выполните индексацию через ⚙.";
                resultsGrid.Invalidate();
            }
            catch (OperationCanceledException)
            {
                completionStatus = "Очистка локального индекса отменена.";
            }
            catch (Exception ex)
            {
                completionStatus = "Не удалось очистить локальный индекс: " + ex.Message;
            }
            finally
            {
                if (ReferenceEquals(resetCancellation, currentCancellation))
                {
                    resetCancellation = null;
                }
                currentCancellation.Dispose();
                RunOnUi(() =>
                {
                    resetRunning = false;
                    statusLabel.Text = completionStatus;
                    queryBox.Enabled = true;
                    clearButton.Enabled = true;
                    searchButton.Enabled = searchCancellation == null;
                    var indexing = outlookMapiImportRunner.IsRunning;
                    indexButton.Enabled = !indexing;
                    stopButton.Enabled = indexing;
                    UpdateResetIndexButtonEnabled();
                });
            }
        }

        private void QueryBoxOnKeyDown(object sender, KeyEventArgs eventArgs)
        {
            if (eventArgs.KeyCode == Keys.Escape)
            {
                eventArgs.SuppressKeyPress = true;
                ClearButtonOnClick(sender, EventArgs.Empty);
                return;
            }
            if (eventArgs.KeyCode != Keys.Enter)
            {
                return;
            }

            eventArgs.SuppressKeyPress = true;
            if (searchButton.Enabled)
            {
                SearchButtonOnClick(sender, EventArgs.Empty);
            }
        }

        private void OutlookMapiImportRunnerOnProgressChanged(object sender, OutlookMapiImportProgress progress)
        {
            RunOnUi(() =>
            {
                var active = outlookMapiImportRunner.IsRunning;
                stopButton.Enabled = progress.IsRunning && active;
                indexButton.Enabled = !active && !resetRunning;
                UpdateResetIndexButtonEnabled();
                statusLabel.Text = progress.Status;
            });
        }

        private void UpdateResetIndexButtonEnabled()
        {
            resetIndexButton.Enabled = !resetRunning &&
                                       searchCancellation == null &&
                                       !outlookMapiImportRunner.IsRunning;
        }

        private async Task RefreshHealthAsync()
        {
            try
            {
                var health = await serviceClient.GetHealthAsync(CancellationToken.None);
                statusLabel.Text = health != null &&
                                   string.Equals(health.Status, "ok", StringComparison.Ordinal)
                    ? "Локальный сервис готов."
                    : "Локальный сервис ответил, но сообщил об ошибке.";
            }
            catch (Exception)
            {
                statusLabel.Text =
                    "Локальный сервис сейчас не запущен; он запустится автоматически при первом поиске.";
            }
        }

        private void RunOnUi(Action action)
        {
            if (IsDisposed)
            {
                return;
            }

            if (InvokeRequired)
            {
                try
                {
                    BeginInvoke(action);
                }
                catch (InvalidOperationException)
                {
                    // Outlook is closing or the task pane handle is gone.
                }
                return;
            }

            action();
        }

        private sealed class ScrollRowsRequestedEventArgs : EventArgs
        {
            public ScrollRowsRequestedEventArgs(int wheelDetents)
            {
                WheelDetents = wheelDetents;
            }

            public int WheelDetents { get; private set; }
        }

        private sealed class OutlookSafeDataGridView : DataGridView
        {
            private const int WmMouseWheel = 0x020A;
            private const int WmMouseHorizontalWheel = 0x020E;
            private const int WheelDelta = 120;

            private int wheelDeltaRemainder;

            public event EventHandler<ScrollRowsRequestedEventArgs> ScrollRowsRequested;

            protected override void OnMouseDown(MouseEventArgs eventArgs)
            {
                if (CanFocus && !Focused)
                {
                    Focus();
                }
                base.OnMouseDown(eventArgs);
            }

            protected override void WndProc(ref Message message)
            {
                if (message.Msg == WmMouseWheel)
                {
                    var delta = unchecked((short)(
                        (message.WParam.ToInt64() >> 16) & 0xffff));
                    if (delta != 0)
                    {
                        if (wheelDeltaRemainder != 0 &&
                            Math.Sign(wheelDeltaRemainder) != Math.Sign(delta))
                        {
                            wheelDeltaRemainder = 0;
                        }
                        wheelDeltaRemainder += delta;
                        var detents = wheelDeltaRemainder / WheelDelta;
                        wheelDeltaRemainder -= detents * WheelDelta;
                        if (detents != 0)
                        {
                            var handler = ScrollRowsRequested;
                            if (handler != null)
                            {
                                handler(this, new ScrollRowsRequestedEventArgs(detents));
                            }
                        }
                    }

                    // Never call DataGridView.OnMouseWheel here. Its native
                    // scrollbar update is the call that blocks Outlook's STA.
                    return;
                }

                if (message.Msg == WmMouseHorizontalWheel)
                {
                    // Columns are fitted to the pane, so horizontal wheel input
                    // has no useful target. Consume it inside the task pane too.
                    return;
                }

                base.WndProc(ref message);
            }
        }

        private sealed class PanePalette
        {
            private PanePalette(bool dark)
            {
                IsDark = dark;
                if (dark)
                {
                    Background = Color.FromArgb(31, 31, 31);
                    AlternateBackground = Color.FromArgb(35, 35, 35);
                    Surface = Color.FromArgb(38, 38, 38);
                    HeaderBackground = Color.FromArgb(43, 43, 43);
                    InputBackground = Color.FromArgb(48, 48, 48);
                    ButtonBackground = Color.FromArgb(45, 45, 45);
                    ButtonHover = Color.FromArgb(58, 58, 58);
                    ButtonPressed = Color.FromArgb(68, 68, 68);
                    Text = Color.FromArgb(243, 243, 243);
                    MutedText = Color.FromArgb(174, 174, 174);
                    Border = Color.FromArgb(73, 73, 73);
                    Selection = Color.FromArgb(14, 72, 112);
                    SelectionText = Color.White;
                    SelectionMutedText = Color.FromArgb(220, 235, 247);
                    ChipBackground = Color.FromArgb(24, 62, 88);
                    ChipText = Color.FromArgb(143, 211, 255);
                }
                else
                {
                    Background = Color.White;
                    AlternateBackground = Color.FromArgb(250, 250, 250);
                    Surface = Color.FromArgb(250, 250, 250);
                    HeaderBackground = Color.FromArgb(246, 246, 246);
                    InputBackground = Color.White;
                    ButtonBackground = Color.FromArgb(250, 250, 250);
                    ButtonHover = Color.FromArgb(235, 235, 235);
                    ButtonPressed = Color.FromArgb(220, 220, 220);
                    Text = Color.FromArgb(32, 32, 32);
                    MutedText = Color.FromArgb(96, 96, 96);
                    Border = Color.FromArgb(210, 210, 210);
                    Selection = Color.FromArgb(204, 228, 247);
                    SelectionText = Color.FromArgb(24, 24, 24);
                    SelectionMutedText = Color.FromArgb(72, 72, 72);
                    ChipBackground = Color.FromArgb(222, 238, 252);
                    ChipText = Color.FromArgb(0, 90, 158);
                }

                Accent = Color.FromArgb(15, 108, 189);
                AccentBorder = Color.FromArgb(0, 90, 158);
                AccentHover = Color.FromArgb(17, 119, 208);
                AccentPressed = Color.FromArgb(0, 90, 158);
            }

            public bool IsDark { get; private set; }
            public Color Background { get; private set; }
            public Color AlternateBackground { get; private set; }
            public Color Surface { get; private set; }
            public Color HeaderBackground { get; private set; }
            public Color InputBackground { get; private set; }
            public Color ButtonBackground { get; private set; }
            public Color ButtonHover { get; private set; }
            public Color ButtonPressed { get; private set; }
            public Color Text { get; private set; }
            public Color MutedText { get; private set; }
            public Color Border { get; private set; }
            public Color Selection { get; private set; }
            public Color SelectionText { get; private set; }
            public Color SelectionMutedText { get; private set; }
            public Color ChipBackground { get; private set; }
            public Color ChipText { get; private set; }
            public Color Accent { get; private set; }
            public Color AccentBorder { get; private set; }
            public Color AccentHover { get; private set; }
            public Color AccentPressed { get; private set; }

            public static PanePalette Create()
            {
                return new PanePalette(IsWindowsDarkMode());
            }

            private static bool IsWindowsDarkMode()
            {
                try
                {
                    using (var key = Registry.CurrentUser.OpenSubKey(
                        @"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"))
                    {
                        var value = key == null ? null : key.GetValue("AppsUseLightTheme");
                        return value is int && (int)value == 0;
                    }
                }
                catch (Exception)
                {
                    return false;
                }
            }
        }

        private sealed class DarkMenuColorTable : ProfessionalColorTable
        {
            private readonly PanePalette palette;

            public DarkMenuColorTable(PanePalette palette)
            {
                this.palette = palette;
                UseSystemColors = false;
            }

            public override Color ToolStripDropDownBackground
            {
                get { return palette.Surface; }
            }

            public override Color MenuBorder
            {
                get { return palette.Border; }
            }

            public override Color MenuItemBorder
            {
                get { return palette.Border; }
            }

            public override Color MenuItemSelected
            {
                get { return palette.ButtonHover; }
            }

            public override Color ImageMarginGradientBegin
            {
                get { return palette.Surface; }
            }

            public override Color ImageMarginGradientMiddle
            {
                get { return palette.Surface; }
            }

            public override Color ImageMarginGradientEnd
            {
                get { return palette.Surface; }
            }

            public override Color SeparatorDark
            {
                get { return palette.Border; }
            }

            public override Color SeparatorLight
            {
                get { return palette.Border; }
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                outlookMapiImportRunner.ProgressChanged -= OutlookMapiImportRunnerOnProgressChanged;
                if (outlookMapiImportRunner.IsRunning)
                {
                    outlookMapiImportRunner.RequestStop();
                }
                if (searchCancellation != null)
                {
                    var cancellation = searchCancellation;
                    searchCancellation = null;
                    cancellation.Cancel();
                }
                if (resetCancellation != null)
                {
                    var cancellation = resetCancellation;
                    resetCancellation = null;
                    cancellation.Cancel();
                }
                settingsMenu.Dispose();
                foreach (var font in ownedFonts)
                {
                    font.Dispose();
                }
                ownedFonts.Clear();
            }

            base.Dispose(disposing);
        }
    }
}
