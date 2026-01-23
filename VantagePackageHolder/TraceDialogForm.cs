using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace VantagePackageHolder
{
    internal enum TraceDialogMode
    {
        Precedents,
        Dependents
    }

    internal enum TraceItemKind
    {
        Root,
        Group,
        Cell,
        MultiCell,
        Name,
        Chart,
        ThreeDRange
    }

    internal sealed class TraceDialogForm : Form
    {
        private readonly Excel.Application _app;
        private readonly TraceDialogMode _mode;
        private readonly RichTextBox _formulaBox;
        private readonly TreeView _tree;
        private readonly Label _status;
        private readonly CheckBox _group;
        private readonly List<TraceItem> _items;
        private readonly List<TraceItem> _formulaItems;
        private readonly Color[] _palette;
        private readonly Color _highlightBack;
        private readonly Color _accentColor;
        private readonly Color _accentLight;
        private readonly Color _dividerColor;
        private readonly Color _outlineColor;
        private readonly Color _buttonColor;
        private Excel.Range _anchorCell;
        private string _formulaText;
        private Excel.Range _highlightedRange;
        private bool _isRefreshing;
        private string _highlightFormula;
        private TraceItem _lastNavigatedItem;
        private Stopwatch _refreshWatch;
        private bool _refreshTimedOut;
        private readonly List<Excel.Worksheet> _navUnhiddenSheets = new List<Excel.Worksheet>();
        private readonly ImageList _treeImages;
        private readonly bool _unhideEnabled = true;
        private readonly bool _highlightEnabled = true;
        private readonly bool _moveEnabled = true;
        private bool _hasVerticalScrollbar;
        private bool _pendingDeactivateClose;
        public bool HasContent { get; private set; }

        private const int MaxValueText = 48;
        private const int MaxTraceCells = 200;
        private const int MaxArrowLinks = 256;
        private const int MaxExpandCells = 200;
        private const int RefreshTimeoutMs = 10000;
        private const int ValueColumnWidth = 96;
        private static readonly bool UseArrowPrecedents = false;
        private const int ResizeBorder = 6;
        private const int HeaderDragHeight = 28;
        private const string HighlightFormula = "=LEN(\"VANTAGE_TRACE\")>0";
        private const int WM_NCHITTEST = 0x0084;
        private const int HTCLIENT = 1;
        private const int HTCAPTION = 2;
        private const int HTLEFT = 10;
        private const int HTRIGHT = 11;
        private const int HTTOP = 12;
        private const int HTTOPLEFT = 13;
        private const int HTTOPRIGHT = 14;
        private const int HTBOTTOM = 15;
        private const int HTBOTTOMLEFT = 16;
        private const int HTBOTTOMRIGHT = 17;
        private static readonly Color HighlightSelectionColor = Color.FromArgb(218, 240, 255);
        private static readonly Color HighlightCrosshairColor = Color.FromArgb(161, 190, 249);

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        private enum HighlightMode
        {
            Selection,
            Crosshairs
        }

        private HighlightMode _highlightMode = HighlightMode.Crosshairs;

        public TraceDialogForm(Excel.Application app, TraceDialogMode mode)
        {
            _app = app ?? throw new ArgumentNullException(nameof(app));
            _mode = mode;
            _items = new List<TraceItem>();
            _formulaItems = new List<TraceItem>();
            _palette = BuildPalette();
            _highlightBack = Color.FromArgb(212, 236, 252);
            _accentColor = Color.FromArgb(0, 150, 255);
            _accentLight = Color.FromArgb(161, 190, 249);
            _dividerColor = Color.FromArgb(220, 230, 240);
            _outlineColor = Color.FromArgb(0, 110, 210);
            _buttonColor = Color.FromArgb(0, 170, 255);
            _treeImages = new ImageList
            {
                ImageSize = new Size(10, 10),
                ColorDepth = ColorDepth.Depth32Bit,
                TransparentColor = Color.Transparent
            };
            _treeImages.Images.Add("expanded", CreateArrowBitmap(true));
            _treeImages.Images.Add("collapsed", CreateArrowBitmap(false));
            _treeImages.Images.Add("home", CreateHomeBitmap());

            Text = mode == TraceDialogMode.Precedents ? "Trace In" : "Trace Out";
            FormBorderStyle = FormBorderStyle.SizableToolWindow;
            ControlBox = true;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            TopMost = true;
            StartPosition = FormStartPosition.Manual;
            Size = new Size(520, 320);
            MinimumSize = new Size(420, 260);
            BackColor = SystemColors.Control;
            Font = SystemFonts.MessageBoxFont;
            RightToLeft = RightToLeft.No;
            RightToLeftLayout = false;
            KeyPreview = true;
            DoubleBuffered = true;
            SizeGripStyle = SizeGripStyle.Auto;
            Padding = new Padding(6);

            var contentPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = SystemColors.Control
            };

            var formulaPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 36,
                BackColor = SystemColors.Window,
                Padding = new Padding(2)
            };

            _formulaBox = new RichTextBox
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                BorderStyle = BorderStyle.FixedSingle,
                Font = new Font("Consolas", 9f),
                BackColor = SystemColors.Window,
                DetectUrls = false,
                HideSelection = false,
                ScrollBars = RichTextBoxScrollBars.None
            };
            formulaPanel.Controls.Add(_formulaBox);

            var optionsPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 0,
                BackColor = SystemColors.Control,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                Padding = new Padding(2, 2, 2, 2),
                AutoSize = false,
                Visible = false
            };
            _group = new CheckBox
            {
                Text = "Group dependents",
                AutoSize = true,
                Checked = false
            };

            var tableHeaderPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 20,
                BackColor = SystemColors.Control,
                Padding = new Padding(2, 0, 2, 0)
            };

            var headerTitle = new Label
            {
                Text = _mode == TraceDialogMode.Precedents ? "Precedents" : "Dependents",
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft
            };

            var headerValue = new Label
            {
                Text = "Value",
                Dock = DockStyle.Right,
                Width = ValueColumnWidth,
                TextAlign = ContentAlignment.MiddleRight
            };

            var headerDivider = new Panel
            {
                Dock = DockStyle.Right,
                Width = 1,
                BackColor = SystemColors.ControlDark
            };

            tableHeaderPanel.Controls.Add(headerTitle);
            tableHeaderPanel.Controls.Add(headerDivider);
            tableHeaderPanel.Controls.Add(headerValue);

            _tree = new TreeView
            {
                Dock = DockStyle.Fill,
                BorderStyle = BorderStyle.FixedSingle,
                HideSelection = false,
                FullRowSelect = true,
                RightToLeft = RightToLeft.No,
                ShowLines = false,
                ShowRootLines = false,
                ShowPlusMinus = false,
                DrawMode = TreeViewDrawMode.OwnerDrawText
            };
            _tree.ImageList = _treeImages;
            _tree.ImageIndex = -1;
            _tree.SelectedImageIndex = -1;

            _status = new Label
            {
                Dock = DockStyle.Bottom,
                Height = 18,
                TextAlign = ContentAlignment.MiddleLeft
            };

            contentPanel.Controls.Add(_tree);
            contentPanel.Controls.Add(tableHeaderPanel);
            contentPanel.Controls.Add(optionsPanel);
            contentPanel.Controls.Add(formulaPanel);
            contentPanel.Controls.Add(_status);
            Controls.Add(contentPanel);

            _tree.AfterSelect += (_, __) => OnTreeSelectionChanged();
            _tree.BeforeExpand += (_, e) => EnsureChildrenLoaded(e.Node);
            _tree.AfterExpand += (_, e) =>
            {
                UpdateNodeIconAfterToggle(e.Node);
                UpdateScrollbarState();
                _tree.Invalidate();
            };
            _tree.AfterCollapse += (_, e) =>
            {
                UpdateNodeIconAfterToggle(e.Node);
                UpdateScrollbarState();
                _tree.Invalidate();
            };
            _tree.NodeMouseDoubleClick += (_, __) => NavigateToSelected();
            _tree.Resize += (_, __) =>
            {
                UpdateScrollbarState();
                _tree.Invalidate();
            };
            _tree.DrawNode += (_, e) => DrawTreeNode(e);
            _tree.Paint += (_, e) => DrawTreeGrid(e.Graphics);

            Load += (_, __) => PositionNearExcelTopRight();
            Shown += (_, __) => FocusTree();
            Activated += (_, __) => FocusTree();
            Deactivate += (_, __) =>
            {
                if (!IsExcelForeground())
                {
                    _pendingDeactivateClose = true;
                    if (!_isRefreshing)
                    {
                        CloseFromDeactivate();
                    }
                }
            };
            // grouping disabled; keep checkbox hidden

            FormClosed += (_, __) =>
            {
                RestoreHiddenSheets();
                RemoveHighlight();
                ClearTraceArrows();
            };
        }

        public void RefreshFromSelection()
        {
            if (_isRefreshing)
            {
                return;
            }

            _isRefreshing = true;
            try
            {
                HasContent = false;
                _refreshTimedOut = false;
                _refreshWatch = Stopwatch.StartNew();

                _items.Clear();
                _formulaItems.Clear();
                _formulaText = string.Empty;
                _anchorCell = null;
                _lastNavigatedItem = null;
                _tree.Nodes.Clear();
                _status.Text = string.Empty;
                ClearTraceArrows();

                if (!RangeHelpers.TryGetRangeOrActiveCell(_app, out var range))
                {
                    SetStatus("No active cell.");
                    ApplyFormulaText(string.Empty, null);
                    return;
                }

                _anchorCell = range.Cells[1, 1];
                _formulaText = ReadFormulaText(_anchorCell);

                if (!HasFormula(_anchorCell) || IsRangeEmpty(_anchorCell))
                {
                    ShowNoDependenciesMessage();
                    return;
                }

                if (_mode == TraceDialogMode.Precedents)
                {
                    BuildPrecedents(_anchorCell);
                }
                else
                {
                    BuildDependents(_anchorCell);
                }

                if (_items.Count == 0)
                {
                    ShowNoDependenciesMessage();
                    return;
                }

                ApplyFormulaText(_formulaText, _items.FirstOrDefault());
                PopulateTree();
                UpdateScrollbarState();
                _tree.Invalidate();
                if (_tree.Nodes.Count > 0)
                {
                    _tree.SelectedNode = _tree.Nodes[0];
                }
                HasContent = true;
                FocusTree();
            }
            finally
            {
                ResetExcelCursor();
                _refreshWatch = null;
                _isRefreshing = false;
                if (_pendingDeactivateClose && !IsExcelForeground())
                {
                    CloseFromDeactivate();
                }
            }
        }

        private bool CheckTimeout()
        {
            if (_refreshWatch == null)
            {
                return false;
            }

            if (_refreshWatch.ElapsedMilliseconds <= RefreshTimeoutMs)
            {
                return false;
            }

            if (!_refreshTimedOut)
            {
                _refreshTimedOut = true;
                SetStatus($"Trace timed out after {RefreshTimeoutMs / 1000} seconds.");
            }

            return true;
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == Keys.Escape)
            {
                Close();
                return true;
            }

            if (keyData == Keys.Enter || keyData == Keys.Return)
            {
                NavigateToSelected();
                return true;
            }

            if (keyData == Keys.Up)
            {
                return MoveSelection(-1);
            }

            if (keyData == Keys.Down)
            {
                return MoveSelection(1);
            }

            if (keyData == Keys.Right)
            {
                var node = _tree.SelectedNode;
                if (node != null && node.Nodes.Count > 0 && !node.IsExpanded)
                {
                    node.Expand();
                    UpdateNodeIconAfterToggle(node);
                    _tree.Invalidate();
                    return true;
                }
                return MoveSelection(1);
            }

            if (keyData == Keys.Left)
            {
                var node = _tree.SelectedNode;
                if (node != null && node.IsExpanded)
                {
                    node.Collapse();
                    UpdateNodeIconAfterToggle(node);
                    _tree.Invalidate();
                    return true;
                }
                return MoveSelection(-1);
            }

            return base.ProcessCmdKey(ref msg, keyData);
        }

        protected override void WndProc(ref Message m)
        {
            if (FormBorderStyle != FormBorderStyle.None)
            {
                base.WndProc(ref m);
                return;
            }

            if (m.Msg == WM_NCHITTEST)
            {
                base.WndProc(ref m);
                if ((int)m.Result == HTCLIENT)
                {
                    int lParam = m.LParam.ToInt32();
                    int x = (short)(lParam & 0xFFFF);
                    int y = (short)((lParam >> 16) & 0xFFFF);
                    Point pos = PointToClient(new Point(x, y));

                    bool left = pos.X <= ResizeBorder;
                    bool right = pos.X >= ClientSize.Width - ResizeBorder;
                    bool top = pos.Y <= ResizeBorder;
                    bool bottom = pos.Y >= ClientSize.Height - ResizeBorder;

                    if (left && top)
                    {
                        m.Result = (IntPtr)HTTOPLEFT;
                        return;
                    }
                    if (right && top)
                    {
                        m.Result = (IntPtr)HTTOPRIGHT;
                        return;
                    }
                    if (left && bottom)
                    {
                        m.Result = (IntPtr)HTBOTTOMLEFT;
                        return;
                    }
                    if (right && bottom)
                    {
                        m.Result = (IntPtr)HTBOTTOMRIGHT;
                        return;
                    }
                    if (left)
                    {
                        m.Result = (IntPtr)HTLEFT;
                        return;
                    }
                    if (right)
                    {
                        m.Result = (IntPtr)HTRIGHT;
                        return;
                    }
                    if (top)
                    {
                        m.Result = (IntPtr)HTTOP;
                        return;
                    }
                    if (bottom)
                    {
                        m.Result = (IntPtr)HTBOTTOM;
                        return;
                    }

                    if (pos.Y <= HeaderDragHeight)
                    {
                        m.Result = (IntPtr)HTCAPTION;
                        return;
                    }
                }

                return;
            }

            base.WndProc(ref m);
        }

        private bool MoveSelection(int delta)
        {
            if (_tree.Nodes.Count == 0)
            {
                return false;
            }

            var node = _tree.SelectedNode ?? _tree.Nodes[0];
            var next = delta > 0 ? node.NextVisibleNode : node.PrevVisibleNode;
            if (next == null)
            {
                next = delta > 0 ? _tree.Nodes[0] : GetLastVisibleNode(_tree.Nodes[0]);
            }

            if (next == null)
            {
                return false;
            }

            _tree.SelectedNode = next;
            next.EnsureVisible();
            return true;
        }

        private static TreeNode GetLastVisibleNode(TreeNode start)
        {
            var node = start;
            var next = node?.NextVisibleNode;
            while (next != null)
            {
                node = next;
                next = node.NextVisibleNode;
            }

            return node;
        }

        private void BuildPrecedents(Excel.Range cell)
        {
            if (cell == null)
            {
                SetStatus("No active cell.");
                return;
            }

            if (CheckTimeout())
            {
                return;
            }

            var refs = string.IsNullOrWhiteSpace(_formulaText)
                ? new List<TraceItem>()
                : CollectReferences(_formulaText);

            var resolvedRefs = new List<TraceItem>();
            foreach (var item in refs)
            {
                ResolveReference(item);
                if (CheckTimeout())
                {
                    return;
                }

                if (ShouldIncludeReference(item))
                {
                    resolvedRefs.Add(item);
                }
            }

            _formulaItems.AddRange(resolvedRefs);

            if (CheckTimeout())
            {
                return;
            }

            var arrowRefs = CollectPrecedents(cell);
            var merged = MergePrecedents(resolvedRefs, arrowRefs);

            var root = BuildRootItem(cell, "Trace In");
            var children = BuildPrecedentChildren(merged);
            if (children.Count == 0)
            {
                if (!HasFormula(cell))
                {
                    SetStatus("Active cell has no formula.");
                }
                else
                {
                    SetStatus("No precedents found.");
                }
                return;
            }

            root.Children.AddRange(children);
            _items.Add(root);
            ApplyPalette(children);
        }

        private void BuildDependents(Excel.Range cell)
        {
            if (cell == null)
            {
                SetStatus("No active cell.");
                return;
            }

            if (CheckTimeout())
            {
                return;
            }

            if (!RangeHelpers.TryGetActiveRange(_app, out var selection))
            {
                selection = cell;
            }

            var dependents = CollectDependents(selection);
            if (CheckTimeout())
            {
                return;
            }

            var nameDependents = CollectNameDependents(selection);
            if (CheckTimeout())
            {
                return;
            }

            var chartDependents = CollectChartDependents(selection);

            if (dependents.Count == 0 && nameDependents.Count == 0 && chartDependents.Count == 0)
            {
                SetStatus("No dependents found.");
                return;
            }

            var root = BuildRootItem(selection, "Trace Out");
            var items = new List<TraceItem>();

            foreach (var dep in dependents)
            {
                items.Add(BuildRangeItem(dep, TraceItemKind.Cell));
            }

            items.AddRange(nameDependents);
            items.AddRange(chartDependents);

            root.Children.AddRange(items);
            _items.Add(root);
            ApplyPalette(items);
        }

        private void PopulateTree()
        {
            _tree.BeginUpdate();
            _tree.Nodes.Clear();
            foreach (var item in _items)
            {
                _tree.Nodes.Add(BuildTreeNode(item));
            }
            _tree.EndUpdate();
        }
        private void UpdateScrollbarState()
        {
            if (!_tree.IsHandleCreated)
            {
                _hasVerticalScrollbar = false;
                return;
            }

            int visibleNodes = 0;
            var node = _tree.Nodes.Count > 0 ? _tree.Nodes[0] : null;
            while (node != null)
            {
                visibleNodes++;
                node = node.NextVisibleNode;
            }

            if (visibleNodes == 0 || _tree.ItemHeight <= 0)
            {
                _hasVerticalScrollbar = false;
                return;
            }

            int capacity = Math.Max(1, _tree.ClientSize.Height / _tree.ItemHeight);
            _hasVerticalScrollbar = visibleNodes > capacity;
        }

        private void OnTreeSelectionChanged()
        {
            if (_tree.SelectedNode?.Tag is TraceItem item)
            {
                ApplyFormulaText(_formulaText, item.StartIndex >= 0 ? item : null);
                if (IsNavigable(item))
                {
                    NavigateToItem(item, false);
                }
                else
                {
                    RemoveHighlight();
                }
            }
            else
            {
                ApplyFormulaText(_formulaText, null);
                RemoveHighlight();
            }
        }

        private void NavigateToSelected()
        {
            if (_tree.SelectedNode?.Tag is TraceItem item)
            {
                if (IsNavigable(item))
                {
                    NavigateToItem(item, true);
                    return;
                }
            }

            var node = _tree.SelectedNode;
            if (node != null && node.Nodes.Count > 0)
            {
                node.Expand();
                _tree.SelectedNode = node.Nodes[0];
            }
        }

        private void NavigateToItem(TraceItem item, bool force)
        {
            if (item == null)
            {
                return;
            }

            ClearTraceArrows();

            if (item.Kind == TraceItemKind.Chart && item.Chart != null)
            {
                TryNavigateChart(item.Chart);
                _lastNavigatedItem = item;
                FocusTree();
                return;
            }

            if (item.Range != null && RangeHelpers.IsRangeValid(item.Range))
            {
                TryUnhide(item.Range);
                RangeHelpers.SafeActivateSheet(item.Range.Worksheet);
                RangeHelpers.SafeSelect(item.Range);
                if (_moveEnabled)
                {
                    MoveFormAsNeeded(item.Range);
                }
                if (_highlightEnabled)
                {
                    ApplyHighlight(item.Range);
                }
                _lastNavigatedItem = item;
                CenterActiveCell();
                FocusTree();
                return;
            }

            if (force && !string.IsNullOrWhiteSpace(item.Token))
            {
                TryGoto(item.Token);
                if (_highlightEnabled && _app.ActiveCell is Excel.Range active)
                {
                    ApplyHighlight(active);
                }
                CenterActiveCell();
                FocusTree();
            }
        }

        private void ApplyFormulaText(string formula, TraceItem selected)
        {
            _formulaBox.Text = formula ?? string.Empty;
            _formulaBox.SelectAll();
            _formulaBox.SelectionColor = Color.Black;
            _formulaBox.SelectionBackColor = Color.White;

            if (string.IsNullOrEmpty(formula))
            {
                return;
            }

            foreach (var item in _formulaItems)
            {
                if (item.Length <= 0 || item.StartIndex < 0 || item.StartIndex + item.Length > formula.Length)
                {
                    continue;
                }

                _formulaBox.Select(item.StartIndex, item.Length);
                _formulaBox.SelectionColor = item.Color;
                if (selected != null
                    && item.StartIndex == selected.StartIndex
                    && item.Length == selected.Length
                    && item.StartIndex >= 0)
                {
                    _formulaBox.SelectionBackColor = _highlightBack;
                }
                else
                {
                    _formulaBox.SelectionBackColor = Color.White;
                }
            }

            _formulaBox.Select(0, 0);
        }

        private static string BuildNodeText(TraceItem item)
        {
            if (item == null)
            {
                return string.Empty;
            }

            if (item.Kind == TraceItemKind.Group)
            {
                var label = item.Label ?? item.DisplayAddress ?? item.Token ?? string.Empty;
                if (item.Children.Count > 0)
                {
                    label = $"{label} ({item.Children.Count})";
                }
                return label;
            }

            if (ShouldDisplayNameLabel(item))
            {
                return item.Label ?? item.Token ?? item.DisplayAddress ?? string.Empty;
            }

            var text = item.DisplayAddress;
            if (string.IsNullOrWhiteSpace(text))
            {
                text = item.Label;
            }
            if (string.IsNullOrWhiteSpace(text))
            {
                text = item.Token;
            }

            return text ?? string.Empty;
        }

        private static bool ShouldDisplayNameLabel(TraceItem item)
        {
            if (item == null)
            {
                return false;
            }

            if (item.Kind == TraceItemKind.Name && !string.IsNullOrWhiteSpace(item.Label))
            {
                return true;
            }

            if (!string.IsNullOrWhiteSpace(item.Label) && IsBareNameToken(item.Token))
            {
                return true;
            }

            if (!string.IsNullOrWhiteSpace(item.Label)
                && !string.IsNullOrWhiteSpace(item.DisplayAddress)
                && !string.Equals(item.Label, item.DisplayAddress, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            return false;
        }

        private TraceItem BuildRootItem(Excel.Range range, string label)
        {
            var item = new TraceItem
            {
                Kind = TraceItemKind.Root
            };

            if (RangeHelpers.IsRangeValid(range))
            {
                item.Range = range;
                item.Label = TryGetAddress(range, false);
                item.DisplayAddress = GetDisplayAddress(range) ?? item.Label;
                item.ValueText = GetRangeValueText(range);
            }
            else
            {
                item.Label = label ?? string.Empty;
            }

            return item;
        }

        private TraceItem BuildRangeItem(ResolvedRange resolved, TraceItemKind kind)
        {
            if (resolved == null)
            {
                return new TraceItem { Kind = kind, StartIndex = -1 };
            }

            var item = new TraceItem
            {
                Kind = kind,
                Token = resolved.Token,
                Range = resolved.Range,
                DisplayAddress = resolved.DisplayAddress,
                ValueText = resolved.ValueText,
                StartIndex = -1
            };

            if (string.IsNullOrWhiteSpace(item.Label))
            {
                if (IsBareNameToken(item.Token))
                {
                    item.Label = item.Token;
                }
                else if (RangeHelpers.IsRangeValid(item.Range))
                {
                    item.Label = TryGetNameForRange(item.Range);
                }
            }

            if (RangeHelpers.IsRangeValid(item.Range))
            {
                item.DisplayAddress = GetDisplayAddress(item.Range) ?? item.DisplayAddress ?? TryGetAddress(item.Range, true);
            }

            if (string.IsNullOrWhiteSpace(item.ValueText) && RangeHelpers.IsRangeValid(item.Range))
            {
                item.ValueText = GetRangeValueText(item.Range);
            }

            if (kind == TraceItemKind.Cell && RangeHelpers.IsRangeValid(item.Range))
            {
                try
                {
                    if (Convert.ToInt64(item.Range.CountLarge) > 1)
                    {
                        var multiCell = new TraceItem
                        {
                            Kind = TraceItemKind.MultiCell,
                            Token = item.Token,
                            Range = item.Range,
                            DisplayAddress = item.DisplayAddress,
                            ValueText = GetRangeValueText(item.Range),
                            StartIndex = -1,
                            HasDeferredChildren = true
                        };
                        return multiCell;
                    }
                }
                catch
                {
                    // ignore
                }
            }

            MarkDeferredIfFormula(item);
            return item;
        }

        private TraceItem BuildRangeItem(Excel.Range range, TraceItemKind kind)
        {
            return BuildRangeItem(BuildResolved(range), kind);
        }

        private List<TraceItem> BuildPrecedentChildren(List<TraceItem> merged)
        {
            var children = new List<TraceItem>();
            if (merged == null || merged.Count == 0)
            {
                return children;
            }

            var ordered = merged
                .OrderBy(item => item.StartIndex < 0 ? int.MaxValue : item.StartIndex)
                .ThenBy(item => item.StartIndex < 0 ? (item.DisplayAddress ?? item.Token ?? string.Empty) : string.Empty)
                .ToList();

            foreach (var source in ordered)
            {
                if (source == null)
                {
                    continue;
                }

                if (TryBuildThreeDRange(source, out var threeDItem))
                {
                    children.Add(threeDItem);
                    continue;
                }

                var child = BuildPrecedentItem(source, source.Range);
                if (child != null)
                {
                    children.Add(child);
                }
            }

            return children;
        }

        private TraceItem BuildPrecedentItem(TraceItem source, Excel.Range range)
        {
            if (source == null)
            {
                return null;
            }

            if (!RangeHelpers.IsRangeValid(range))
            {
                return new TraceItem
                {
                    Kind = TraceItemKind.Cell,
                    Token = source.Token,
                    DisplayAddress = source.DisplayAddress ?? source.Token,
                    ValueText = source.ValueText,
                    Label = source.Label,
                    ArgumentLabel = source.ArgumentLabel,
                    Color = source.Color,
                    StartIndex = source.StartIndex,
                    Length = source.Length
                };
            }

            bool isMultiCell = false;
            try
            {
                isMultiCell = Convert.ToInt64(range.CountLarge) > 1;
            }
            catch
            {
                isMultiCell = false;
            }

            if (!isMultiCell)
            {
                var cellItem = new TraceItem
                {
                    Kind = TraceItemKind.Cell,
                    Range = range,
                    Token = source.Token,
                    DisplayAddress = GetDisplayAddress(range) ?? source.DisplayAddress ?? TryGetAddress(range, true),
                    ValueText = string.IsNullOrWhiteSpace(source.ValueText) ? GetRangeValueText(range) : source.ValueText,
                    Label = source.Label,
                    ArgumentLabel = source.ArgumentLabel,
                    Color = source.Color,
                    StartIndex = source.StartIndex,
                    Length = source.Length
                };
                if (string.IsNullOrWhiteSpace(cellItem.Label))
                {
                    cellItem.Label = TryGetNameForRange(range);
                }
                MarkDeferredIfFormula(cellItem);
                return cellItem;
            }

            var multiCell = new TraceItem
            {
                Kind = TraceItemKind.MultiCell,
                Range = range,
                Token = source.Token,
                DisplayAddress = GetDisplayAddress(range) ?? source.DisplayAddress ?? TryGetAddress(range, true),
                ValueText = GetRangeValueText(range),
                Label = source.Label,
                ArgumentLabel = source.ArgumentLabel,
                Color = source.Color,
                StartIndex = source.StartIndex,
                Length = source.Length,
                HasDeferredChildren = true
            };
            if (string.IsNullOrWhiteSpace(multiCell.Label))
            {
                multiCell.Label = TryGetNameForRange(range);
            }

            return multiCell;
        }

        private void MarkDeferredIfFormula(TraceItem item)
        {
            if (item == null || item.HasDeferredChildren)
            {
                return;
            }

            if (item.Kind == TraceItemKind.Root || item.Kind == TraceItemKind.Group || item.Kind == TraceItemKind.ThreeDRange)
            {
                return;
            }

            if (!RangeHelpers.IsRangeValid(item.Range))
            {
                return;
            }

            try
            {
                if (Convert.ToInt64(item.Range.CountLarge) > 1)
                {
                    return;
                }
            }
            catch
            {
                return;
            }

            if (!HasFormula(item.Range))
            {
                return;
            }

            if (!HasFormulaReferences(item.Range))
            {
                return;
            }

            item.HasDeferredChildren = true;
        }

        private bool HasFormulaReferences(Excel.Range range)
        {
            if (!RangeHelpers.IsRangeValid(range))
            {
                return false;
            }

            var formula = ReadFormulaText(range);
            if (string.IsNullOrWhiteSpace(formula))
            {
                return false;
            }

            var refs = CollectReferences(formula);
            if (refs.Count == 0)
            {
                return false;
            }

            foreach (var item in refs)
            {
                if (item == null || string.IsNullOrWhiteSpace(item.Token))
                {
                    continue;
                }

                if (IsBareNameToken(item.Token) && !IsDefinedName(item.Token) && !IsReservedNameToken(item.Token))
                {
                    continue;
                }

                return true;
            }

            return false;
        }

        private void AddMultiCellChildren(TraceItem parent, Excel.Range range, TraceItem source)
        {
            if (parent == null || !RangeHelpers.IsRangeValid(range))
            {
                return;
            }

            int count = 0;
            try
            {
                foreach (Excel.Range cell in range.Cells)
                {
                    if (CheckTimeout())
                    {
                        break;
                    }

                    if (cell == null)
                    {
                        continue;
                    }

                    if (++count > MaxExpandCells)
                    {
                        break;
                    }

                    var child = new TraceItem
                    {
                        Kind = TraceItemKind.Cell,
                        Range = cell,
                        DisplayAddress = GetDisplayAddress(cell) ?? TryGetAddress(cell, true),
                        ValueText = GetRangeValueText(cell),
                        Color = source?.Color ?? Color.Empty,
                        StartIndex = source?.StartIndex ?? -1,
                        Length = source?.Length ?? 0
                    };
                    parent.Children.Add(child);
                }
            }
            catch
            {
                // ignore enumeration failures
            }
        }

        private void EnsureChildrenLoaded(TreeNode node)
        {
            if (node == null)
            {
                return;
            }

            if (node.Tag is TraceItem item && item.HasDeferredChildren && !item.ChildrenLoaded)
            {
                item.ChildrenLoaded = true;
                item.Children.Clear();
                if (item.Kind == TraceItemKind.MultiCell)
                {
                    AddMultiCellChildren(item, item.Range, item);
                }
                else if (RangeHelpers.IsRangeValid(item.Range) && HasFormula(item.Range))
                {
                    var ancestorKeys = GetAncestorKeys(node);
                    var children = BuildNestedPrecedents(item.Range, ancestorKeys, item);
                    item.Children.AddRange(children);
                }

                node.Nodes.Clear();
                foreach (var child in item.Children)
                {
                    node.Nodes.Add(BuildTreeNode(child));
                }

                if (node.Nodes.Count == 0)
                {
                    item.HasDeferredChildren = false;
                }

                UpdateNodeIconAfterToggle(node);
            }
        }

        private void UpdateNodeIconAfterToggle(TreeNode node)
        {
            if (node == null)
            {
                return;
            }

            if (node.Tag is TraceItem item && item.Kind == TraceItemKind.Root)
            {
                SetNodeIcon(node, "home");
                return;
            }

            if (!ShouldShowExpandIcon(node))
            {
                ClearNodeIcon(node);
                return;
            }

            UpdateNodeIcon(node, node.IsExpanded);
        }

        private bool ShouldShowExpandIcon(TreeNode node)
        {
            if (node == null)
            {
                return false;
            }

            if (!(node.Tag is TraceItem item))
            {
                return node.Nodes.Count > 0;
            }

            if (item.Kind == TraceItemKind.Root)
            {
                return true;
            }

            if (node.Nodes.Count == 0)
            {
                return false;
            }

            if (item.Kind == TraceItemKind.MultiCell || item.Kind == TraceItemKind.Group || item.Kind == TraceItemKind.ThreeDRange)
            {
                return true;
            }

            if (item.Kind == TraceItemKind.Name)
            {
                return node.Nodes.Count > 0;
            }

            if (item.Kind == TraceItemKind.Cell)
            {
                return RangeHelpers.IsRangeValid(item.Range) && HasFormula(item.Range);
            }

            return node.Nodes.Count > 0;
        }

        private List<TraceItem> BuildNestedPrecedents(Excel.Range cell, HashSet<string> ancestorKeys, TraceItem parent)
        {
            var children = new List<TraceItem>();
            if (cell == null)
            {
                return children;
            }

            var formula = ReadFormulaText(cell);
            if (string.IsNullOrWhiteSpace(formula))
            {
                return children;
            }

            var refs = CollectReferences(formula);
            var resolvedRefs = new List<TraceItem>();
            foreach (var item in refs)
            {
                ResolveReference(item);
                if (CheckTimeout())
                {
                    return children;
                }

                if (!ShouldIncludeReference(item))
                {
                    continue;
                }

                item.StartIndex = -1;
                item.Length = 0;
                if (parent != null && !parent.Color.IsEmpty)
                {
                    item.Color = parent.Color;
                }
                resolvedRefs.Add(item);
            }

            var arrowRefs = CollectPrecedents(cell);
            var merged = MergePrecedents(resolvedRefs, arrowRefs);
            var built = BuildPrecedentChildren(merged);
            if (ancestorKeys == null || ancestorKeys.Count == 0)
            {
                return built;
            }

            foreach (var child in built)
            {
                if (!IsAncestorMatch(child, ancestorKeys))
                {
                    children.Add(child);
                }
            }

            return children;
        }

        private static HashSet<string> GetAncestorKeys(TreeNode node)
        {
            var keys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var current = node;
            while (current != null)
            {
                if (current.Tag is TraceItem item && RangeHelpers.IsRangeValid(item.Range))
                {
                    var key = RangeHelpers.BuildRangeKey(item.Range);
                    if (!string.IsNullOrEmpty(key))
                    {
                        keys.Add(key);
                    }
                }

                current = current.Parent;
            }

            return keys;
        }

        private static bool IsAncestorMatch(TraceItem item, HashSet<string> ancestorKeys)
        {
            if (item == null || ancestorKeys == null || ancestorKeys.Count == 0)
            {
                return false;
            }

            if (RangeHelpers.IsRangeValid(item.Range))
            {
                var key = RangeHelpers.BuildRangeKey(item.Range);
                if (!string.IsNullOrEmpty(key) && ancestorKeys.Contains(key))
                {
                    return true;
                }
            }

            return false;
        }

        private bool TryBuildThreeDRange(TraceItem source, out TraceItem item)
        {
            item = null;
            if (source == null || string.IsNullOrWhiteSpace(source.Token))
            {
                return false;
            }

            if (!TryParseThreeDReference(source.Token, out var workbook, out var startSheet, out var endSheet, out var address))
            {
                return false;
            }

            var rangeValue = GetRangeValueText(source.Range);
            var threeD = new TraceItem
            {
                Kind = TraceItemKind.ThreeDRange,
                Token = source.Token,
                Label = source.DisplayAddress ?? source.Token,
                ValueText = string.IsNullOrWhiteSpace(rangeValue) ? "<array>" : rangeValue,
                ArgumentLabel = source.ArgumentLabel,
                Color = source.Color,
                StartIndex = source.StartIndex,
                Length = source.Length
            };

            foreach (var range in ExpandThreeDRanges(workbook, startSheet, endSheet, address))
            {
                var child = BuildPrecedentItem(new TraceItem
                {
                    Color = source.Color,
                    StartIndex = source.StartIndex,
                    Length = source.Length
                }, range);
                if (child != null)
                {
                    threeD.Children.Add(child);
                }
            }

            item = threeD;
            return true;
        }

        private bool TryParseThreeDReference(string token, out Excel.Workbook workbook, out string startSheet, out string endSheet, out string address)
        {
            workbook = null;
            startSheet = null;
            endSheet = null;
            address = null;

            if (!TryParseQualifiedReference(token, out var workbookName, out var sheetName, out address))
            {
                return false;
            }

            if (string.IsNullOrWhiteSpace(sheetName) || !sheetName.Contains(":"))
            {
                return false;
            }

            var parts = sheetName.Split(new[] { ':' }, 2);
            if (parts.Length != 2)
            {
                return false;
            }

            startSheet = parts[0];
            endSheet = parts[1];
            if (string.IsNullOrWhiteSpace(startSheet) || string.IsNullOrWhiteSpace(endSheet) || string.IsNullOrWhiteSpace(address))
            {
                return false;
            }

            workbook = GetWorkbookByName(workbookName) ?? _app.ActiveWorkbook;
            if (workbook == null)
            {
                return false;
            }

            return true;
        }

        private IEnumerable<Excel.Range> ExpandThreeDRanges(Excel.Workbook workbook, string startSheet, string endSheet, string address)
        {
            var ranges = new List<Excel.Range>();
            if (workbook == null || string.IsNullOrWhiteSpace(startSheet) || string.IsNullOrWhiteSpace(endSheet))
            {
                return ranges;
            }

            Excel.Worksheet start = null;
            Excel.Worksheet end = null;
            try
            {
                start = workbook.Worksheets[startSheet] as Excel.Worksheet;
                end = workbook.Worksheets[endSheet] as Excel.Worksheet;
            }
            catch
            {
                return ranges;
            }

            if (start == null || end == null)
            {
                return ranges;
            }

            int startIndex = start.Index;
            int endIndex = end.Index;
            if (startIndex > endIndex)
            {
                int temp = startIndex;
                startIndex = endIndex;
                endIndex = temp;
            }

            int added = 0;
            for (int i = startIndex; i <= endIndex; i++)
            {
                if (added >= MaxExpandCells)
                {
                    break;
                }

                Excel.Worksheet sheet = null;
                try
                {
                    sheet = workbook.Worksheets[i] as Excel.Worksheet;
                }
                catch
                {
                    sheet = null;
                }

                if (sheet == null)
                {
                    continue;
                }

                try
                {
                    var rng = sheet.Range[address];
                    if (RangeHelpers.IsRangeValid(rng))
                    {
                        ranges.Add(rng);
                        added++;
                    }
                }
                catch
                {
                    // ignore
                }
            }

            return ranges;
        }

        private void ApplyPalette(IEnumerable<TraceItem> items)
        {
            if (_palette == null || _palette.Length == 0)
            {
                return;
            }

            var colorMap = new Dictionary<string, Color>(StringComparer.OrdinalIgnoreCase);
            int index = 0;
            foreach (var item in _formulaItems)
            {
                if (item == null || item.StartIndex < 0 || item.Length <= 0)
                {
                    continue;
                }

                var key = $"{item.StartIndex}:{item.Length}";
                if (!colorMap.TryGetValue(key, out var color))
                {
                    color = _palette[index % _palette.Length];
                    colorMap[key] = color;
                    index++;
                }

                item.Color = color;
            }

            if (items == null)
            {
                return;
            }

            foreach (var item in items)
            {
                ApplyPaletteToItem(item, colorMap);
            }
        }

        private void ApplyPaletteToItem(TraceItem item, Dictionary<string, Color> colorMap)
        {
            if (item == null)
            {
                return;
            }

            if (item.StartIndex >= 0 && item.Length > 0 && item.Color.IsEmpty)
            {
                var key = $"{item.StartIndex}:{item.Length}";
                if (colorMap.TryGetValue(key, out var color))
                {
                    item.Color = color;
                }
            }

            foreach (var child in item.Children)
            {
                ApplyPaletteToItem(child, colorMap);
            }
        }

        private static Bitmap CreateArrowBitmap(bool expanded)
        {
            var fileName = expanded ? "trace_expanded.png" : "trace_collapsed.png";
            return LoadTreeBitmap(fileName);
        }

        private static Bitmap CreateHomeBitmap()
        {
            return LoadTreeBitmap("trace_home.png");
        }

        private static Bitmap LoadTreeBitmap(string fileName)
        {
            try
            {
                var path = GetResourcePath(fileName);
                if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
                {
                    return new Bitmap(10, 10);
                }

                using (var source = new Bitmap(path))
                {
                    var result = new Bitmap(10, 10);
                    using (var g = Graphics.FromImage(result))
                    {
                        g.Clear(Color.Transparent);
                        g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                        g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                        g.DrawImage(source, new Rectangle(0, 0, 10, 10));
                    }
                    result.MakeTransparent(Color.Black);
                    return result;
                }
            }
            catch
            {
                return new Bitmap(10, 10);
            }
        }

        private static string GetResourcePath(string fileName)
        {
            try
            {
                var assemblyDir = Path.GetDirectoryName(typeof(TraceDialogForm).Assembly.Location);
                if (!string.IsNullOrWhiteSpace(assemblyDir))
                {
                    var path = Path.Combine(assemblyDir, "Resources", fileName);
                    if (File.Exists(path))
                    {
                        return path;
                    }
                }
            }
            catch
            {
                // ignore
            }

            try
            {
                var baseDir = AppDomain.CurrentDomain.BaseDirectory;
                if (!string.IsNullOrWhiteSpace(baseDir))
                {
                    var path = Path.Combine(baseDir, "Resources", fileName);
                    if (File.Exists(path))
                    {
                        return path;
                    }
                }
            }
            catch
            {
                // ignore
            }

            return null;
        }

        private void UpdateNodeIcon(TreeNode node, bool expanded)
        {
            if (node == null || _treeImages == null)
            {
                return;
            }

            if (node.Tag is TraceItem item && item.Kind == TraceItemKind.Root)
            {
                SetNodeIcon(node, "home");
                return;
            }

            if (!ShouldShowExpandIcon(node))
            {
                ClearNodeIcon(node);
                return;
            }

            SetNodeIcon(node, expanded ? "expanded" : "collapsed");
        }

        private void SetNodeIcon(TreeNode node, string key)
        {
            if (node == null || _treeImages == null)
            {
                return;
            }

            if (string.IsNullOrWhiteSpace(key))
            {
                ClearNodeIcon(node);
                return;
            }

            node.ImageKey = key;
            node.SelectedImageKey = key;
            node.ImageIndex = _treeImages.Images.IndexOfKey(key);
            node.SelectedImageIndex = node.ImageIndex;
        }

        private static void ClearNodeIcon(TreeNode node)
        {
            if (node == null)
            {
                return;
            }

            node.ImageKey = null;
            node.SelectedImageKey = null;
            node.ImageIndex = -1;
            node.SelectedImageIndex = -1;
        }


        private TreeNode BuildTreeNode(TraceItem item)
        {
            var node = new TreeNode(BuildNodeText(item))
            {
                Tag = item,
                ImageIndex = -1,
                SelectedImageIndex = -1
            };

            foreach (var child in item.Children)
            {
                node.Nodes.Add(BuildTreeNode(child));
            }

            if (item.HasDeferredChildren
                && item.Kind != TraceItemKind.MultiCell
                && RangeHelpers.IsRangeValid(item.Range)
                && !HasFormulaReferences(item.Range))
            {
                item.HasDeferredChildren = false;
            }

            if (item.HasDeferredChildren && node.Nodes.Count == 0)
            {
                node.Nodes.Add(new TreeNode());
            }

            if (item.Kind == TraceItemKind.Root && node.Nodes.Count > 0)
            {
                node.Expand();
            }

            UpdateNodeIconAfterToggle(node);

            return node;
        }

        private void DrawTreeNode(DrawTreeNodeEventArgs e)
        {
            if (e.Node == null)
            {
                return;
            }

            e.DrawDefault = true;

            if (!(e.Node.Tag is TraceItem item) || string.IsNullOrWhiteSpace(item.ValueText))
            {
                return;
            }

            int valueStart = GetValueColumnStart();
            if (valueStart <= 0)
            {
                return;
            }

            int scrollbarWidth = _hasVerticalScrollbar ? SystemInformation.VerticalScrollBarWidth : 0;
            int valueRight = Math.Max(0, _tree.ClientSize.Width - scrollbarWidth);
            int valueWidth = Math.Max(0, valueRight - valueStart);
            if (valueWidth <= 0)
            {
                return;
            }

            bool selected = (e.State & TreeNodeStates.Selected) == TreeNodeStates.Selected;
            var backColor = selected ? SystemColors.Highlight : _tree.BackColor;
            var foreColor = selected ? SystemColors.HighlightText : _tree.ForeColor;
            var valueRect = new Rectangle(valueStart + 2, e.Bounds.Y, Math.Max(0, valueWidth - 4), e.Bounds.Height);

            using (var backBrush = new SolidBrush(backColor))
            {
                e.Graphics.FillRectangle(backBrush, valueRect);
            }

            TextRenderer.DrawText(
                e.Graphics,
                item.ValueText,
                _tree.Font,
                valueRect,
                foreColor,
                TextFormatFlags.VerticalCenter | TextFormatFlags.Right | TextFormatFlags.EndEllipsis);
        }

        private void DrawTreeGrid(Graphics graphics)
        {
            if (graphics == null)
            {
                return;
            }

            int valueStart = GetValueColumnStart();
            if (valueStart <= 0)
            {
                return;
            }

            using (var pen = new Pen(_dividerColor))
            {
                graphics.DrawLine(pen, valueStart, 0, valueStart, _tree.ClientSize.Height);
            }
        }

        private int GetValueColumnStart()
        {
            int scrollbarWidth = _hasVerticalScrollbar ? SystemInformation.VerticalScrollBarWidth : 0;
            int start = _tree.ClientSize.Width - ValueColumnWidth - scrollbarWidth;
            return Math.Max(0, start);
        }

        private bool IsNavigable(TraceItem item)
        {
            if (item == null)
            {
                return false;
            }

            if (item.Kind == TraceItemKind.Chart)
            {
                return item.Chart != null;
            }

            if (RangeHelpers.IsRangeValid(item.Range))
            {
                return true;
            }

            return !string.IsNullOrWhiteSpace(item.Token);
        }

        private void UpdateHighlightState()
        {
            if (!_highlightEnabled)
            {
                RemoveHighlight();
                return;
            }

            if (_tree.SelectedNode?.Tag is TraceItem item && IsNavigable(item) && RangeHelpers.IsRangeValid(item.Range))
            {
                ApplyHighlight(item.Range);
            }
        }

        private string BuildHighlightFormula(Excel.Range range, HighlightMode mode)
        {
            if (mode != HighlightMode.Crosshairs || !RangeHelpers.IsRangeValid(range))
            {
                return HighlightFormula;
            }

            try
            {
                var cell = range.Cells[1, 1];
                int row = cell.Row;
                int col = cell.Column;
                var baseExpr = HighlightFormula.StartsWith("=")
                    ? HighlightFormula.Substring(1)
                    : HighlightFormula;
                return $"=AND({baseExpr},NOT(AND(ROW()={row},COLUMN()={col})))";
            }
            catch
            {
                return HighlightFormula;
            }
        }

        private void ApplyHighlight(Excel.Range range)
        {
            if (!RangeHelpers.IsRangeValid(range))
            {
                return;
            }

            RemoveHighlight();

            try
            {
                if (Convert.ToInt64(range.CountLarge) > 1)
                {
                    return;
                }
            }
            catch
            {
                // ignore count check failures
            }

            try
            {
                if (range.Worksheet != null)
                {
                    if (Convert.ToInt64(range.Rows.CountLarge) >= Convert.ToInt64(range.Worksheet.Rows.CountLarge)
                        || Convert.ToInt64(range.Columns.CountLarge) >= Convert.ToInt64(range.Worksheet.Columns.CountLarge))
                    {
                        return;
                    }
                }
            }
            catch
            {
                // ignore
            }

            var mode = GetHighlightModeForRange(range);
            Excel.Range target = range;
            if (mode == HighlightMode.Crosshairs)
            {
                try
                {
                    target = _app.Union(range.EntireRow, range.EntireColumn);
                }
                catch
                {
                    target = range;
                }
            }

            if (target == null)
            {
                return;
            }

            try
            {
                var conditions = target.FormatConditions;
                var formula = BuildHighlightFormula(range, mode);
                conditions.Add(Excel.XlFormatConditionType.xlExpression, Type.Missing, formula);
                var format = (Excel.FormatCondition)conditions.Item(conditions.Count);
                var color = mode == HighlightMode.Crosshairs ? HighlightCrosshairColor : HighlightSelectionColor;
                format.Interior.Color = ColorTranslator.ToOle(color);
                format.Interior.Pattern = Excel.XlPattern.xlPatternSolid;
                _highlightFormula = formula;
            }
            catch
            {
                // ignore
            }

            _highlightedRange = range;
        }

        private void RemoveHighlight()
        {
            if (_highlightedRange == null)
            {
                _highlightFormula = null;
                return;
            }

            Excel.Range range = _highlightedRange;
            HighlightMode mode = GetHighlightModeForRange(range);
            Excel.Range target = range;
            if (mode == HighlightMode.Crosshairs)
            {
                try
                {
                    target = _app.Union(range.EntireRow, range.EntireColumn);
                }
                catch
                {
                    target = range;
                }
            }

            try
            {
                var matchFormula = _highlightFormula ?? HighlightFormula;
                var conditions = target?.FormatConditions;
                if (conditions != null)
                {
                    for (int i = conditions.Count; i >= 1; i--)
                    {
                        try
                        {
                            var condition = (Excel.FormatCondition)conditions.Item(i);
                            if (condition != null
                                && condition.Type == (int)Excel.XlFormatConditionType.xlExpression
                                && (string.Equals(condition.Formula1, matchFormula, StringComparison.OrdinalIgnoreCase)
                                    || string.Equals(condition.Formula1, HighlightFormula, StringComparison.OrdinalIgnoreCase)))
                            {
                                condition.Delete();
                                break;
                            }
                        }
                        catch
                        {
                            // ignore
                        }
                    }
                }
            }
            catch
            {
                // ignore
            }

            _highlightedRange = null;
            _highlightFormula = null;
        }

        private HighlightMode GetHighlightModeForRange(Excel.Range range)
        {
            if (!RangeHelpers.IsRangeValid(range))
            {
                return _highlightMode;
            }

            try
            {
                if (Convert.ToInt64(range.Rows.CountLarge) > 1 || Convert.ToInt64(range.Columns.CountLarge) > 1)
                {
                    return HighlightMode.Selection;
                }
            }
            catch
            {
                // ignore
            }

            return _highlightMode;
        }

        private void MoveFormAsNeeded(Excel.Range range)
        {
            if (!RangeHelpers.IsRangeValid(range))
            {
                return;
            }

            try
            {
                if (_app.ActiveWindow != null && _app.ActiveWindow.Split)
                {
                    return;
                }
            }
            catch
            {
                // ignore
            }

            Rectangle formRect = RectangleToScreen(ClientRectangle);
            Rectangle rangeRect = GetRangeScreenRect(range);
            if (rangeRect == Rectangle.Empty || (!formRect.IntersectsWith(rangeRect) && !rangeRect.IntersectsWith(formRect)))
            {
                return;
            }

            const int padding = 5;
            int x = formRect.X;
            int y = formRect.Y;
            int width = formRect.Width;
            int height = formRect.Height;

            var workingArea = Screen.FromControl(this).WorkingArea;
            bool canRight = rangeRect.Right + padding + width <= workingArea.Right;
            bool canLeft = rangeRect.Left - padding - width >= workingArea.Left;
            bool canBottom = rangeRect.Bottom + padding + height <= workingArea.Bottom;
            bool canTop = rangeRect.Top - padding - height >= workingArea.Top;

            var candidates = new List<Tuple<int, string>>
            {
                Tuple.Create(Math.Abs(x - (rangeRect.Right + padding)), "right"),
                Tuple.Create(Math.Abs(x - (rangeRect.Left - padding - width)), "left"),
                Tuple.Create(Math.Abs(y - (rangeRect.Bottom + padding)), "bottom"),
                Tuple.Create(Math.Abs(y - (rangeRect.Top - padding - height)), "top")
            };
            candidates.Sort((a, b) => a.Item1.CompareTo(b.Item1));

            foreach (var candidate in candidates)
            {
                switch (candidate.Item2)
                {
                    case "right" when canRight:
                        x = rangeRect.Right + padding;
                        break;
                    case "left" when canLeft:
                        x = rangeRect.Left - padding - width;
                        break;
                    case "bottom" when canBottom:
                        y = rangeRect.Bottom + padding;
                        break;
                    case "top" when canTop:
                        y = rangeRect.Top - padding - height;
                        break;
                    default:
                        continue;
                }

                var moved = new Rectangle(x, y, width, height);
                if (!moved.IntersectsWith(rangeRect) && !rangeRect.IntersectsWith(moved))
                {
                    Location = new Point(x, y);
                    break;
                }
            }
        }

        private Rectangle GetRangeScreenRect(Excel.Range range)
        {
            try
            {
                var window = _app.ActiveWindow;
                if (window == null)
                {
                    return Rectangle.Empty;
                }

                double zoom = 1.0;
                try
                {
                    zoom = Convert.ToDouble(window.Zoom) / 100.0;
                }
                catch
                {
                    zoom = 1.0;
                }

                double leftPoints = Convert.ToDouble(range.Left) * zoom;
                double topPoints = Convert.ToDouble(range.Top) * zoom;
                double rightPoints = leftPoints + Convert.ToDouble(range.Width) * zoom;
                double bottomPoints = topPoints + Convert.ToDouble(range.Height) * zoom;

                int left = window.PointsToScreenPixelsX((int)Math.Round(leftPoints));
                int top = window.PointsToScreenPixelsY((int)Math.Round(topPoints));
                int right = window.PointsToScreenPixelsX((int)Math.Round(rightPoints));
                int bottom = window.PointsToScreenPixelsY((int)Math.Round(bottomPoints));

                if (right < left)
                {
                    (left, right) = (right, left);
                }
                if (bottom < top)
                {
                    (top, bottom) = (bottom, top);
                }

                return new Rectangle(left, top, Math.Max(1, right - left), Math.Max(1, bottom - top));
            }
            catch
            {
                return Rectangle.Empty;
            }
        }

        private void TryNavigateChart(Excel.Chart chart)
        {
            if (chart == null)
            {
                return;
            }

            try
            {
                var chartObject = chart.Parent as Excel.ChartObject;
                if (chartObject != null)
                {
                    RangeHelpers.SafeActivateSheet(chartObject.Parent as Excel.Worksheet);
                    chartObject.Activate();
                    return;
                }

                var shape = chart.Parent as Excel.Shape;
                if (shape != null)
                {
                    RangeHelpers.SafeActivateSheet(shape.Parent as Excel.Worksheet);
                    shape.Select(Type.Missing);
                    return;
                }

                chart.Activate();
            }
            catch
            {
                // ignore
            }
        }

        private List<TraceItem> CollectNameDependents(Excel.Range selection)
        {
            var items = new List<TraceItem>();
            if (!RangeHelpers.IsRangeValid(selection))
            {
                return items;
            }

            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var workbook = selection.Worksheet?.Parent as Excel.Workbook ?? _app.ActiveWorkbook;

            void AddName(Excel.Name name)
            {
                if (name == null)
                {
                    return;
                }

                var nameKey = name.Name;
                if (!string.IsNullOrWhiteSpace(nameKey) && !seen.Add("name:" + nameKey))
                {
                    return;
                }

                Excel.Range refersTo = null;
                try
                {
                    refersTo = name.RefersToRange;
                }
                catch
                {
                    refersTo = null;
                }

                if (!RangeHelpers.IsRangeValid(refersTo))
                {
                    return;
                }

                Excel.Range intersect = null;
                try
                {
                    intersect = _app.Intersect(selection, refersTo);
                }
                catch
                {
                    intersect = null;
                }

                if (intersect == null)
                {
                    return;
                }

                var item = BuildRangeItem(refersTo, TraceItemKind.Name);
                item.Label = nameKey;
                item.Token = nameKey;
                item.ValueText = GetRangeValueText(refersTo);
                items.Add(item);
            }

            try
            {
                if (workbook != null)
                {
                    foreach (Excel.Name name in workbook.Names)
                    {
                        AddName(name);
                    }
                }
            }
            catch
            {
                // ignore
            }

            try
            {
                var sheet = selection.Worksheet;
                if (sheet != null)
                {
                    foreach (Excel.Name name in sheet.Names)
                    {
                        AddName(name);
                    }
                }
            }
            catch
            {
                // ignore
            }

            return items;
        }

        private List<TraceItem> CollectChartDependents(Excel.Range selection)
        {
            var items = new List<TraceItem>();
            if (!RangeHelpers.IsRangeValid(selection))
            {
                return items;
            }

            var workbook = selection.Worksheet?.Parent as Excel.Workbook ?? _app.ActiveWorkbook;
            if (workbook == null)
            {
                return items;
            }

            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            void AddChart(Excel.Chart chart)
            {
                if (chart == null)
                {
                    return;
                }

                var name = chart.Name;
                var sheet = GetChartWorksheet(chart);
                var sheetName = sheet?.Name ?? string.Empty;
                var key = string.IsNullOrWhiteSpace(sheetName) ? name : $"{sheetName}|{name}";
                if (!string.IsNullOrWhiteSpace(key) && !seen.Add(key))
                {
                    return;
                }

                if (!ChartDependsOnSelection(chart, selection))
                {
                    return;
                }

                items.Add(new TraceItem
                {
                    Kind = TraceItemKind.Chart,
                    Label = name,
                    ValueText = "Chart",
                    Chart = chart,
                    StartIndex = -1
                });
            }

            try
            {
                foreach (Excel.Worksheet sheet in workbook.Worksheets)
                {
                    Excel.ChartObjects chartObjects = null;
                    try
                    {
                        chartObjects = sheet.ChartObjects() as Excel.ChartObjects;
                    }
                    catch
                    {
                        chartObjects = null;
                    }

                    if (chartObjects != null)
                    {
                        for (int i = 1; i <= chartObjects.Count; i++)
                        {
                            try
                            {
                                var chartObject = chartObjects.Item(i) as Excel.ChartObject;
                                AddChart(chartObject?.Chart);
                            }
                            catch
                            {
                                // ignore
                            }
                        }
                    }

                    try
                    {
                        var shapes = sheet.Shapes;
                        if (shapes != null)
                        {
                            for (int i = 1; i <= shapes.Count; i++)
                            {
                                Excel.Shape shape = null;
                                try
                                {
                                    shape = shapes.Item(i);
                                }
                                catch
                                {
                                    shape = null;
                                }

                                if (shape == null)
                                {
                                    continue;
                                }

                                bool hasChart = false;
                                try
                                {
                                    hasChart = Convert.ToBoolean(shape.HasChart);
                                }
                                catch
                                {
                                    hasChart = false;
                                }

                                if (hasChart)
                                {
                                    AddChart(shape.Chart);
                                }
                            }
                        }
                    }
                    catch
                    {
                        // ignore
                    }
                }
            }
            catch
            {
                // ignore
            }

            try
            {
                foreach (Excel.Chart chart in workbook.Charts)
                {
                    AddChart(chart);
                }
            }
            catch
            {
                // ignore
            }

            return items;
        }

        private bool ChartDependsOnSelection(Excel.Chart chart, Excel.Range selection)
        {
            if (chart == null || !RangeHelpers.IsRangeValid(selection))
            {
                return false;
            }

            try
            {
                var seriesCollection = chart.SeriesCollection();
                foreach (Excel.Series series in seriesCollection)
                {
                    if (SeriesDependsOnSelection(series, selection))
                    {
                        return true;
                    }
                }
            }
            catch
            {
                return false;
            }

            return false;
        }

        private bool SeriesDependsOnSelection(Excel.Series series, Excel.Range selection)
        {
            if (series == null || !RangeHelpers.IsRangeValid(selection))
            {
                return false;
            }

            foreach (var range in GetSeriesRanges(series))
            {
                if (RangeHelpers.IsRangeValid(range) && RangesIntersect(range, selection))
                {
                    return true;
                }
            }

            string formula = null;
            try
            {
                formula = series.Formula;
            }
            catch
            {
                formula = null;
            }

            if (!string.IsNullOrWhiteSpace(formula))
            {
                foreach (var reference in CollectReferences(formula))
                {
                    if (string.IsNullOrWhiteSpace(reference.Token))
                    {
                        continue;
                    }

                    var resolved = TryResolveToken(reference.Token);
                    if (RangeHelpers.IsRangeValid(resolved.Range) && RangesIntersect(resolved.Range, selection))
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        private IEnumerable<Excel.Range> GetSeriesRanges(Excel.Series series)
        {
            var ranges = new List<Excel.Range>();
            if (series == null)
            {
                return ranges;
            }

            try
            {
                var valuesRange = series.Values as Excel.Range;
                if (valuesRange != null)
                {
                    ranges.Add(valuesRange);
                }
            }
            catch
            {
                // ignore
            }

            try
            {
                var xRange = series.XValues as Excel.Range;
                if (xRange != null)
                {
                    ranges.Add(xRange);
                }
            }
            catch
            {
                // ignore
            }

            return ranges;
        }

        private bool RangesIntersect(Excel.Range left, Excel.Range right)
        {
            try
            {
                return _app.Intersect(left, right) != null;
            }
            catch
            {
                return false;
            }
        }

        private List<TraceItem> BuildDependentGroups(List<TraceItem> items)
        {
            if (items == null || items.Count == 0)
            {
                return new List<TraceItem>();
            }

            var anchorKey = GetAnchorGroupKey();
            var groups = new List<TraceItem>();
            foreach (var bucket in BuildGroupBuckets(items))
            {
                if (!string.IsNullOrEmpty(anchorKey)
                    && string.Equals(bucket.Key, anchorKey, StringComparison.OrdinalIgnoreCase))
                {
                    foreach (var item in OrderItems(bucket.Items))
                    {
                        groups.Add(item);
                    }
                    continue;
                }

                var group = new TraceItem
                {
                    Kind = TraceItemKind.Group,
                    Label = bucket.Label,
                    StartIndex = -1
                };

                foreach (var item in OrderItems(bucket.Items))
                {
                    group.Children.Add(item);
                }

                groups.Add(group);
            }

            return groups;
        }

        private string GetAnchorGroupKey()
        {
            if (!RangeHelpers.IsRangeValid(_anchorCell))
            {
                return null;
            }

            try
            {
                var ws = _anchorCell.Worksheet;
                var wb = ws?.Parent as Excel.Workbook;
                var wbName = wb?.Name ?? string.Empty;
                var sheetName = ws?.Name ?? string.Empty;
                if (!string.IsNullOrEmpty(sheetName))
                {
                    return $"{wbName}|{sheetName}";
                }
            }
            catch
            {
                // ignore
            }

            return null;
        }

        private IEnumerable<TraceItem> OrderItems(IEnumerable<TraceItem> items)
        {
            return items
                .OrderBy(item => item.Range?.Worksheet?.Index ?? int.MaxValue)
                .ThenBy(item => item.Range?.Row ?? int.MaxValue)
                .ThenBy(item => item.Range?.Column ?? int.MaxValue)
                .ThenBy(item => item.Label ?? item.DisplayAddress ?? item.Token ?? string.Empty)
                .ToList();
        }

        private void ResolveReference(TraceItem item)
        {
            if (item == null)
            {
                return;
            }

            if (!string.IsNullOrWhiteSpace(_formulaText) && item.StartIndex >= 0)
            {
                item.ArgumentLabel = GetArgumentLabel(_formulaText, item.StartIndex);
            }

            if (string.IsNullOrWhiteSpace(item.Token))
            {
                return;
            }

            var resolved = TryResolveToken(item.Token);
            item.Range = resolved.Range;
            item.DisplayAddress = resolved.DisplayAddress;
            item.ValueText = resolved.ValueText;
            if (RangeHelpers.IsRangeValid(item.Range))
            {
                item.DisplayAddress = GetDisplayAddress(item.Range) ?? item.DisplayAddress;
            }
            else if (TryParseQualifiedReference(item.Token, out var wbName, out var sheetName, out var address))
            {
                if (IsSameAnchorSheet(sheetName, wbName))
                {
                    item.DisplayAddress = address;
                }
            }

            if (IsBareNameToken(item.Token) && string.IsNullOrWhiteSpace(item.Label))
            {
                item.Label = item.Token;
            }

            if (RangeHelpers.IsRangeValid(item.Range) && string.IsNullOrWhiteSpace(item.Label))
            {
                item.Label = TryGetNameForRange(item.Range);
            }
        }

        private bool ShouldIncludeReference(TraceItem item)
        {
            if (item == null || string.IsNullOrWhiteSpace(item.Token))
            {
                return false;
            }

            if (!IsBareNameToken(item.Token))
            {
                return true;
            }

            if (IsReservedNameToken(item.Token))
            {
                return false;
            }

            if (RangeHelpers.IsRangeValid(item.Range))
            {
                return true;
            }

            return IsDefinedName(item.Token);
        }

        private bool IsDefinedName(string token)
        {
            if (string.IsNullOrWhiteSpace(token))
            {
                return false;
            }

            try
            {
                var wb = _app.ActiveWorkbook;
                if (wb != null)
                {
                    var name = wb.Names.Item(token, Type.Missing, Type.Missing);
                    if (name != null)
                    {
                        return true;
                    }
                }
            }
            catch
            {
                // ignore
            }

            try
            {
                var ws = _app.ActiveSheet as Excel.Worksheet;
                if (ws != null)
                {
                    var name = ws.Names.Item(token, Type.Missing, Type.Missing);
                    if (name != null)
                    {
                        return true;
                    }
                }
            }
            catch
            {
                // ignore
            }

            return false;
        }

        private string TryGetNameForRange(Excel.Range range)
        {
            if (!RangeHelpers.IsRangeValid(range))
            {
                return null;
            }

            try
            {
                var wb = range.Worksheet?.Parent as Excel.Workbook ?? _app.ActiveWorkbook;
                if (wb != null)
                {
                    foreach (Excel.Name name in wb.Names)
                    {
                        Excel.Range refersTo = null;
                        try
                        {
                            refersTo = name.RefersToRange;
                        }
                        catch
                        {
                            refersTo = null;
                        }

                        if (!RangeHelpers.IsRangeValid(refersTo))
                        {
                            continue;
                        }

                        try
                        {
                            if (_app.Intersect(refersTo, range) != null
                                && string.Equals(refersTo.Address[true, true, Excel.XlReferenceStyle.xlA1, true],
                                    range.Address[true, true, Excel.XlReferenceStyle.xlA1, true],
                                    StringComparison.OrdinalIgnoreCase))
                            {
                                return NormalizeDefinedName(name.Name);
                            }
                        }
                        catch
                        {
                            // ignore
                        }
                    }
                }
            }
            catch
            {
                // ignore
            }

            try
            {
                var ws = range.Worksheet;
                if (ws != null)
                {
                    foreach (Excel.Name name in ws.Names)
                    {
                        Excel.Range refersTo = null;
                        try
                        {
                            refersTo = name.RefersToRange;
                        }
                        catch
                        {
                            refersTo = null;
                        }

                        if (!RangeHelpers.IsRangeValid(refersTo))
                        {
                            continue;
                        }

                        try
                        {
                            if (_app.Intersect(refersTo, range) != null
                                && string.Equals(refersTo.Address[true, true, Excel.XlReferenceStyle.xlA1, true],
                                    range.Address[true, true, Excel.XlReferenceStyle.xlA1, true],
                                    StringComparison.OrdinalIgnoreCase))
                            {
                                return NormalizeDefinedName(name.Name);
                            }
                        }
                        catch
                        {
                            // ignore
                        }
                    }
                }
            }
            catch
            {
                // ignore
            }

            return null;
        }

        private static string NormalizeDefinedName(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                return name;
            }

            int bang = name.LastIndexOf('!');
            if (bang >= 0 && bang < name.Length - 1)
            {
                return name.Substring(bang + 1);
            }

            return name;
        }

        private ResolvedRange TryResolveToken(string token)
        {
            var resolved = new ResolvedRange { Token = token };
            if (string.IsNullOrWhiteSpace(token))
            {
                return resolved;
            }

            string workbookName = null;
            string sheetName = null;
            string address = null;

            if (TryParseQualifiedReference(token, out workbookName, out sheetName, out address))
            {
                if (sheetName != null && sheetName.Contains(":"))
                {
                    resolved.DisplayAddress = token;
                    return resolved;
                }

                var wb = GetWorkbookByName(workbookName) ?? _app.ActiveWorkbook;
                if (wb != null && !string.IsNullOrWhiteSpace(sheetName))
                {
                    try
                    {
                        var ws = wb.Worksheets[sheetName] as Excel.Worksheet;
                        if (ws != null)
                        {
                            resolved.Range = ws.Range[address];
                            resolved.DisplayAddress = resolved.Range.Address[false, false, Excel.XlReferenceStyle.xlA1, true];
                            resolved.ValueText = GetRangeValueText(resolved.Range);
                            return resolved;
                        }
                    }
                    catch
                    {
                        // fall back to goto
                    }
                }
            }

            try
            {
                var sheet = _app.ActiveSheet as Excel.Worksheet;
                if (sheet != null)
                {
                    resolved.Range = sheet.Range[token];
                    resolved.DisplayAddress = resolved.Range.Address[false, false, Excel.XlReferenceStyle.xlA1, true];
                    resolved.ValueText = GetRangeValueText(resolved.Range);
                    return resolved;
                }
            }
            catch
            {
                // ignore
            }

            var nameRange = TryResolveNameReference(token);
            if (nameRange != null)
            {
                resolved.Range = nameRange;
                resolved.DisplayAddress = nameRange.Address[false, false, Excel.XlReferenceStyle.xlA1, true];
                resolved.ValueText = GetRangeValueText(nameRange);
                return resolved;
            }

            resolved.DisplayAddress = token;
            return resolved;
        }

        private Excel.Workbook GetWorkbookByName(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                return null;
            }

            try
            {
                return _app.Workbooks[name];
            }
            catch
            {
                return null;
            }
        }

        private bool TryGoto(string token)
        {
            try
            {
                _app.Goto(Reference: token, Scroll: false);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private void TryUnhide(Excel.Range range)
        {
            if (!_unhideEnabled || range == null)
            {
                return;
            }

            try
            {
                var sheet = range.Worksheet;
                if (sheet != null && sheet.Visible != Excel.XlSheetVisibility.xlSheetVisible)
                {
                    sheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;
                    if (!_navUnhiddenSheets.Contains(sheet))
                    {
                        _navUnhiddenSheets.Add(sheet);
                    }
                }
            }
            catch { }

            try { range.EntireRow.Hidden = false; } catch { }
            try { range.EntireColumn.Hidden = false; } catch { }
        }

        private void CenterActiveCell()
        {
            try
            {
                var win = _app.ActiveWindow;
                var cell = _app.ActiveCell as Excel.Range;
                if (win == null || cell == null)
                {
                    return;
                }

                var visible = win.VisibleRange;
                int visibleRows = visible.Rows.Count;
                int visibleCols = visible.Columns.Count;
                int targetRow = Math.Max(1, cell.Row - visibleRows / 2);
                int targetCol = Math.Max(1, cell.Column - visibleCols / 2);
                win.ScrollRow = targetRow;
                win.ScrollColumn = targetCol;
            }
            catch
            {
                // ignore
            }
        }

        private void SetStatus(string text)
        {
            _status.Text = text ?? string.Empty;
        }

        private void CloseFromDeactivate()
        {
            if (IsDisposed || Disposing)
            {
                _pendingDeactivateClose = false;
                return;
            }

            if (IsExcelForeground())
            {
                _pendingDeactivateClose = false;
                return;
            }

            _pendingDeactivateClose = false;
            ResetExcelCursor();
            RemoveHighlight();
            Close();
        }

        private void ClearTraceArrows()
        {
            try
            {
                _app.CommandBars?.ExecuteMso("ClearArrows");
            }
            catch
            {
                // ignore
            }

            try
            {
                _app.CommandBars?.ExecuteMso("RemoveArrows");
            }
            catch
            {
                // ignore
            }

            try
            {
                _app.CommandBars?.ExecuteMso("RemoveAllArrows");
            }
            catch
            {
                // ignore
            }

            try
            {
                if (IsExcelForeground())
                {
                    _app.CommandBars?.ReleaseFocus();
                }
            }
            catch
            {
                // ignore
            }

            try
            {
                _app.StatusBar = false;
            }
            catch
            {
                // ignore
            }

            try
            {
                _app.CutCopyMode = (Excel.XlCutCopyMode)0;
            }
            catch
            {
                // ignore
            }
        }

        private void FocusTree()
        {
            if (IsDisposed)
            {
                return;
            }

            if (IsHandleCreated)
            {
                try
                {
                    Activate();
                    _tree.Focus();
                    return;
                }
                catch
                {
                    // ignore
                }
            }

            try
            {
                BeginInvoke((Action)(() =>
                {
                    try
                    {
                        Activate();
                        _tree.Focus();
                    }
                    catch
                    {
                        // ignore
                    }
                }));
            }
            catch
            {
                // ignore
            }
        }

        private void PositionNearExcelTopRight()
        {
            try
            {
                if (GetWindowRect((IntPtr)_app.Hwnd, out var rect))
                {
                    int x = rect.Right - Width - 16;
                    int y = rect.Top + 80;
                    if (x < rect.Left + 8) x = rect.Left + 8;
                    if (y < rect.Top + 8) y = rect.Top + 8;
                    Location = new Point(x, y);
                    return;
                }
            }
            catch
            {
                // fall back
            }

            var wa = Screen.PrimaryScreen.WorkingArea;
            Location = new Point(wa.Right - Width - 20, wa.Top + 60);
        }

        private static bool HasFormula(Excel.Range cell)
        {
            try
            {
                return Convert.ToBoolean(cell.HasFormula);
            }
            catch
            {
                return false;
            }
        }

        private static bool IsRangeEmpty(Excel.Range cell)
        {
            if (cell == null)
            {
                return true;
            }

            try
            {
                var value = cell.Value2;
                if (value == null)
                {
                    return true;
                }

                if (value is string str)
                {
                    return string.IsNullOrWhiteSpace(str);
                }

                return false;
            }
            catch
            {
                return false;
            }
        }

        private void ResetExcelCursor()
        {
            try
            {
                _app.Cursor = Excel.XlMousePointer.xlDefault;
            }
            catch
            {
                // ignore
            }

            try
            {
                Cursor = Cursors.Default;
            }
            catch
            {
                // ignore
            }
        }

        private void ShowNoDependenciesMessage()
        {
            ResetExcelCursor();
            try
            {
                MessageBox.Show("No dependencies found for the selected cell.", "Trace", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch
            {
                // ignore
            }

            Close();
        }

        private bool IsExcelForeground()
        {
            try
            {
                IntPtr fg = GetForegroundWindow();
                if (fg == IntPtr.Zero)
                {
                    return false;
                }

                uint fgPid;
                GetWindowThreadProcessId(fg, out fgPid);
                uint excelPid;
                GetWindowThreadProcessId((IntPtr)_app.Hwnd, out excelPid);
                return fgPid == excelPid;
            }
            catch
            {
                return false;
            }
        }

        private static string ReadFormulaText(Excel.Range cell)
        {
            if (cell == null)
            {
                return string.Empty;
            }

            try
            {
                var formula2 = cell.GetType().InvokeMember(
                    "Formula2",
                    System.Reflection.BindingFlags.GetProperty,
                    null,
                    cell,
                    null);
                return Convert.ToString(formula2);
            }
            catch
            {
                // fall back
            }

            try
            {
                return Convert.ToString(cell.Formula);
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string GetRangeValueText(Excel.Range range)
        {
            if (range == null)
            {
                return string.Empty;
            }

            try
            {
                long count = Convert.ToInt64(range.CountLarge);
                if (count > 1)
                {
                    return "<array>";
                }
            }
            catch
            {
                // ignore count
            }

            try
            {
                var formatted = Convert.ToString(range.Text);
                if (!string.IsNullOrWhiteSpace(formatted) && !IsHashFill(formatted))
                {
                    return NormalizeValueText(formatted);
                }

                var value = range.Value2;
                if (value == null)
                {
                    return string.Empty;
                }

                var text = Convert.ToString(value);
                return NormalizeValueText(text);
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string NormalizeValueText(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return string.Empty;
            }

            text = text.Replace("\r", " ").Replace("\n", " ").Trim();
            if (text.Length > MaxValueText)
            {
                text = text.Substring(0, MaxValueText - 3) + "...";
            }
            return text;
        }

        private static bool IsHashFill(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return false;
            }

            foreach (var ch in text)
            {
                if (ch != '#')
                {
                    return false;
                }
            }

            return true;
        }

        private static List<TraceItem> CollectReferences(string formulaText)
        {
            var items = new List<TraceItem>();
            if (string.IsNullOrWhiteSpace(formulaText))
            {
                return items;
            }

            var matches = ReferenceRegex.Matches(formulaText);
            if (matches.Count == 0)
            {
                return items;
            }

            foreach (Match match in matches)
            {
                if (!match.Success)
                {
                    continue;
                }

                if (IsIndexInStringLiteral(formulaText, match.Index))
                {
                    continue;
                }

                if (IsFunctionLikeToken(formulaText, match.Index, match.Length))
                {
                    continue;
                }

                if (IsReservedNameToken(match.Value))
                {
                    continue;
                }

                items.Add(new TraceItem
                {
                    Token = match.Value,
                    StartIndex = match.Index,
                    Length = match.Length
                });
            }

            return items;
        }

        private static bool IsReservedNameToken(string token)
        {
            if (string.IsNullOrWhiteSpace(token))
            {
                return false;
            }

            switch (token.Trim().ToUpperInvariant())
            {
                case "TRUE":
                case "FALSE":
                    return true;
                default:
                    return false;
            }
        }

        private static bool IsFunctionLikeToken(string formula, int start, int length)
        {
            string token = formula.Substring(start, length);
            if (token.Contains("!") || token.Contains(":") || token.Contains("["))
            {
                return false;
            }

            int idx = start + length;
            while (idx < formula.Length)
            {
                char ch = formula[idx];
                if (ch == ' ' || ch == '\t')
                {
                    idx++;
                    continue;
                }

                return ch == '(';
            }

            return false;
        }

        private static bool IsBareNameToken(string token)
        {
            if (string.IsNullOrWhiteSpace(token))
            {
                return false;
            }

            if (token.IndexOf('!') >= 0 || token.IndexOf(':') >= 0 || token.IndexOf('[') >= 0)
            {
                return false;
            }

            return !PlainAddressRegex.IsMatch(token);
        }

        private static bool IsIndexInStringLiteral(string text, int index)
        {
            if (index < 0)
            {
                return false;
            }

            bool inString = false;
            for (int i = 0; i <= index && i < text.Length; i++)
            {
                if (text[i] == '"')
                {
                    if (inString)
                    {
                        if (i + 1 < text.Length && text[i + 1] == '"')
                        {
                            i++;
                        }
                        else
                        {
                            inString = false;
                        }
                    }
                    else
                    {
                        inString = true;
                    }
                }
            }

            return inString;
        }

        private static string GetArgumentLabel(string formula, int index)
        {
            if (string.IsNullOrEmpty(formula) || index < 0 || index > formula.Length)
            {
                return string.Empty;
            }

            bool inString = false;
            var stack = new Stack<FunctionContext>();

            for (int i = 0; i < index && i < formula.Length; i++)
            {
                char ch = formula[i];
                if (ch == '"')
                {
                    if (inString)
                    {
                        if (i + 1 < formula.Length && formula[i + 1] == '"')
                        {
                            i++;
                        }
                        else
                        {
                            inString = false;
                        }
                    }
                    else
                    {
                        inString = true;
                    }
                    continue;
                }

                if (inString)
                {
                    continue;
                }

                if (ch == '(')
                {
                    string funcName = ReadFunctionName(formula, i);
                    stack.Push(new FunctionContext(funcName, 1));
                    continue;
                }

                if (ch == ',')
                {
                    foreach (var ctx in stack)
                    {
                        if (!string.IsNullOrEmpty(ctx.Name))
                        {
                            ctx.ArgIndex++;
                            break;
                        }
                    }
                    continue;
                }

                if (ch == ')')
                {
                    if (stack.Count > 0)
                    {
                        stack.Pop();
                    }
                }
            }

            var active = stack.FirstOrDefault(ctx => !string.IsNullOrEmpty(ctx.Name));
            if (active == null || string.IsNullOrEmpty(active.Name))
            {
                return string.Empty;
            }

            string func = active.Name;
            int argIndex = active.ArgIndex;
            if (FunctionArguments.TryGetValue(func, out var args) && argIndex >= 1 && argIndex <= args.Length)
            {
                return func + "." + args[argIndex - 1];
            }

            return func + ".arg" + argIndex;
        }

        private static string ReadFunctionName(string formula, int parenIndex)
        {
            int i = parenIndex - 1;
            while (i >= 0 && char.IsWhiteSpace(formula[i]))
            {
                i--;
            }

            int end = i;
            while (i >= 0 && (char.IsLetterOrDigit(formula[i]) || formula[i] == '_' || formula[i] == '.'))
            {
                i--;
            }

            int start = i + 1;
            if (end < start)
            {
                return string.Empty;
            }

            var name = formula.Substring(start, end - start + 1);
            if (name.Length == 0 || char.IsDigit(name[0]))
            {
                return string.Empty;
            }

            return name.ToUpperInvariant();
        }

        private static bool TryParseQualifiedReference(string token, out string workbookName, out string sheetName, out string address)
        {
            workbookName = null;
            sheetName = null;
            address = null;

            int bang = token.LastIndexOf('!');
            if (bang <= 0 || bang >= token.Length - 1)
            {
                return false;
            }

            var left = token.Substring(0, bang);
            address = token.Substring(bang + 1);

            if (left.Length >= 2 && left[0] == '\'' && left[left.Length - 1] == '\'')
            {
                left = left.Substring(1, left.Length - 2).Replace("''", "'");
            }

            int wbStart = left.IndexOf('[');
            int wbEnd = left.IndexOf(']');
            if (wbStart >= 0 && wbEnd > wbStart)
            {
                workbookName = left.Substring(wbStart + 1, wbEnd - wbStart - 1);
                sheetName = left.Substring(wbEnd + 1);
            }
            else
            {
                sheetName = left;
            }

            return !string.IsNullOrWhiteSpace(sheetName) && !string.IsNullOrWhiteSpace(address);
        }

        private List<ResolvedRange> CollectDependents(Excel.Range selection)
        {
            var results = new List<ResolvedRange>();
            if (selection == null)
            {
                return results;
            }

            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var app = selection.Application;
            var hiddenSheets = _unhideEnabled ? UnhideHiddenSheets(_app) : null;
            Excel.Range originalSelection = null;
            Excel.Range originalActiveCell = null;

            using (new UiGuard(_app, hideStatusBar: true))
            {
                try
                {
                    originalSelection = _app.Selection as Excel.Range;
                    originalActiveCell = _app.ActiveCell;
                }
                catch
                {
                    originalSelection = null;
                    originalActiveCell = null;
                }

                try
                {
                    if (selection.Cells != null)
                    {
                        int processed = 0;
                        foreach (Excel.Range target in selection.Cells)
                        {
                            if (CheckTimeout())
                            {
                                break;
                            }

                            if (target == null)
                            {
                                continue;
                            }

                            if (++processed > MaxTraceCells)
                            {
                                break;
                            }

                            RangeHelpers.SafeActivateSheet(target.Worksheet);
                            RangeHelpers.SafeSelect(target);
                            CollectDependentsFromCell(target, results, seen);
                        }
                    }
                }
                finally
                {
                    ClearTraceArrows();

                    if (hiddenSheets != null)
                    {
                        HidePreviouslyHiddenSheets(_app, hiddenSheets);
                    }

                    try
                    {
                        if (originalSelection != null)
                        {
                            originalSelection.Select();
                            if (originalActiveCell != null)
                            {
                                originalActiveCell.Activate();
                            }
                        }
                    }
                    catch
                    {
                        // ignore
                    }
                }
            }

            return results;
        }

        private void CollectDependentsFromCell(Excel.Range cell, List<ResolvedRange> results, HashSet<string> seen)
        {
            if (cell == null)
            {
                return;
            }

            try
            {
                cell.ShowDependents(Type.Missing);
            }
            catch
            {
                // ignore
            }

            for (int arrow = 1; arrow < MaxArrowLinks; arrow++)
            {
                if (CheckTimeout())
                {
                    break;
                }

                Excel.Range dep = null;
                try
                {
                    dep = cell.NavigateArrow(false, arrow, 1);
                }
                catch
                {
                    dep = null;
                }

                if (dep == null || IsSameAddress(dep, cell))
                {
                    break;
                }

                AddResolvedRange(results, seen, dep);

                for (int link = 2; link < MaxArrowLinks; link++)
                {
                    if (CheckTimeout())
                    {
                        break;
                    }

                    Excel.Range dep2 = null;
                    try
                    {
                        dep2 = cell.NavigateArrow(false, arrow, link);
                    }
                    catch
                    {
                        dep2 = null;
                    }

                    if (dep2 == null || IsSameAddress(dep2, cell))
                    {
                        break;
                    }

                    AddResolvedRange(results, seen, dep2);
                }
            }

            try
            {
                cell.ShowDependents(true);
            }
            catch
            {
                // ignore
            }

            ClearTraceArrows();
        }

        private List<ResolvedRange> CollectPrecedents(Excel.Range cell)
        {
            var results = new List<ResolvedRange>();
            if (cell == null)
            {
                return results;
            }

            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (TryCollectDirectPrecedents(cell, results, seen))
            {
                return results;
            }

            if (!UseArrowPrecedents)
            {
                return results;
            }

            var hiddenSheets = _unhideEnabled ? UnhideHiddenSheets(_app) : null;
            Excel.Range originalSelection = null;
            Excel.Range originalActiveCell = null;
            bool timedOut = false;

            using (new UiGuard(_app, hideStatusBar: true))
            {
                try
                {
                    try
                    {
                        originalSelection = _app.Selection as Excel.Range;
                        originalActiveCell = _app.ActiveCell;
                    }
                    catch
                    {
                        originalSelection = null;
                        originalActiveCell = null;
                    }

                    RangeHelpers.SafeActivateSheet(cell.Worksheet);
                    RangeHelpers.SafeSelect(cell);

                    try
                    {
                        if (CheckTimeout())
                        {
                            timedOut = true;
                            return results;
                        }

                        cell.ShowPrecedents(Type.Missing);
                    }
                    catch
                    {
                        // ignore
                    }

                    for (int arrow = 1; arrow < MaxArrowLinks; arrow++)
                    {
                        if (CheckTimeout())
                        {
                            timedOut = true;
                            break;
                        }

                        Excel.Range prec = null;
                        try
                        {
                            prec = cell.NavigateArrow(true, arrow, 1);
                        }
                        catch
                        {
                            prec = null;
                        }

                        if (prec == null || IsSameAddress(prec, cell))
                        {
                            break;
                        }

                        AddResolvedRange(results, seen, prec);

                        for (int link = 2; link < MaxArrowLinks; link++)
                        {
                            if (CheckTimeout())
                            {
                                timedOut = true;
                                break;
                            }

                            Excel.Range prec2 = null;
                            try
                            {
                                prec2 = cell.NavigateArrow(true, arrow, link);
                            }
                            catch
                            {
                                prec2 = null;
                            }

                            if (prec2 == null || IsSameAddress(prec2, cell))
                            {
                                break;
                            }

                            AddResolvedRange(results, seen, prec2);
                        }

                        if (timedOut)
                        {
                            break;
                        }
                    }

                    if (!timedOut)
                    {
                        try
                        {
                            cell.ShowPrecedents(true);
                        }
                        catch
                        {
                            // ignore
                        }
                    }
                }
                finally
                {
                    ClearTraceArrows();

                    if (hiddenSheets != null)
                    {
                        HidePreviouslyHiddenSheets(_app, hiddenSheets);
                    }

                    try
                    {
                        if (originalSelection != null)
                        {
                            originalSelection.Select();
                            if (originalActiveCell != null)
                            {
                                originalActiveCell.Activate();
                            }
                        }
                    }
                    catch
                    {
                        // ignore
                    }
                }
            }

            return results;
        }

        private static void AddResolvedRange(List<ResolvedRange> results, HashSet<string> seen, Excel.Range range)
        {
            if (range == null)
            {
                return;
            }

            var key = RangeHelpers.BuildRangeKey(range);
            if (!string.IsNullOrEmpty(key) && !seen.Add(key))
            {
                return;
            }

            var resolved = BuildResolved(range);
            if (string.IsNullOrEmpty(key))
            {
                var fallbackKey = resolved.DisplayAddress ?? resolved.Token;
                if (!string.IsNullOrEmpty(fallbackKey) && !seen.Add(fallbackKey))
                {
                    return;
                }
            }

            results.Add(resolved);
        }

        private bool TryCollectDirectPrecedents(Excel.Range cell, List<ResolvedRange> results, HashSet<string> seen)
        {
            if (cell == null || results == null || seen == null)
            {
                return false;
            }

            Excel.Range direct = null;
            try
            {
                direct = cell.DirectPrecedents;
            }
            catch
            {
                direct = null;
            }

            if (direct == null)
            {
                return false;
            }

            try
            {
                var areas = direct.Areas;
                if (areas != null && areas.Count > 0)
                {
                    for (int i = 1; i <= areas.Count; i++)
                    {
                        if (CheckTimeout())
                        {
                            break;
                        }

                        Excel.Range area = null;
                        try
                        {
                            area = areas.Item[i];
                        }
                        catch
                        {
                            area = null;
                        }

                        if (area != null)
                        {
                            AddResolvedRange(results, seen, area);
                        }
                    }

                    return results.Count > 0;
                }
            }
            catch
            {
                // fall back below
            }

            AddResolvedRange(results, seen, direct);
            return results.Count > 0;
        }

        private static ResolvedRange BuildResolved(Excel.Range range)
        {
            var resolved = new ResolvedRange
            {
                Range = range
            };

            if (range == null)
            {
                return resolved;
            }

            try
            {
                resolved.DisplayAddress = range.Address[false, false, Excel.XlReferenceStyle.xlA1, true];
                resolved.ValueText = GetRangeValueText(range);
                resolved.Token = resolved.DisplayAddress;
            }
            catch
            {
                // ignore
            }

            return resolved;
        }

        private static Color[] BuildPalette()
        {
            return new[]
            {
                Color.FromArgb(68, 114, 196),
                Color.FromArgb(237, 125, 49),
                Color.FromArgb(165, 165, 165),
                Color.FromArgb(255, 192, 0),
                Color.FromArgb(91, 155, 213),
                Color.FromArgb(112, 173, 71),
                Color.FromArgb(192, 0, 0),
                Color.FromArgb(112, 48, 160)
            };
        }

        private static readonly Regex ReferenceRegex = new Regex(
            @"(?:(?:'[^']+'|\[[^\]]+\][^!]+|[A-Za-z0-9_.:]+)!)?(?:\$?[A-Za-z]{1,3}\$?\d{1,7}(?::\$?[A-Za-z]{1,3}\$?\d{1,7})?|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}|\$?\d{1,7}:\$?\d{1,7}|[A-Za-z_][A-Za-z0-9_]*\[[^\]]+\]|[A-Za-z_][A-Za-z0-9_.]*)",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static readonly Regex PlainAddressRegex = new Regex(
            @"^\$?[A-Za-z]{1,3}\$?\d{1,7}(:\$?[A-Za-z]{1,3}\$?\d{1,7})?$|^\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}$|^\$?\d{1,7}:\$?\d{1,7}$",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static readonly Dictionary<string, string[]> FunctionArguments = new Dictionary<string, string[]>
        {
            { "INDEX", new[] { "array", "row_num", "column_num", "area_num" } },
            { "MATCH", new[] { "lookup_value", "lookup_array", "match_type" } },
            { "XLOOKUP", new[] { "lookup_value", "lookup_array", "return_array", "if_not_found", "match_mode", "search_mode" } },
            { "XMATCH", new[] { "lookup_value", "lookup_array", "match_mode", "search_mode" } },
            { "VLOOKUP", new[] { "lookup_value", "table_array", "col_index_num", "range_lookup" } },
            { "HLOOKUP", new[] { "lookup_value", "table_array", "row_index_num", "range_lookup" } },
            { "IF", new[] { "logical_test", "value_if_true", "value_if_false" } },
            { "IFS", new[] { "logical_test1", "value_if_true1" } },
            { "SUMIFS", new[] { "sum_range", "criteria_range1", "criteria1" } },
            { "COUNTIFS", new[] { "criteria_range1", "criteria1" } },
            { "AVERAGEIFS", new[] { "average_range", "criteria_range1", "criteria1" } },
            { "MINIFS", new[] { "min_range", "criteria_range1", "criteria1" } },
            { "MAXIFS", new[] { "max_range", "criteria_range1", "criteria1" } },
            { "OFFSET", new[] { "reference", "rows", "cols", "height", "width" } },
            { "CHOOSE", new[] { "index_num", "value1", "value2" } },
            { "INDIRECT", new[] { "ref_text", "a1" } },
            { "TEXT", new[] { "value", "format_text" } },
            { "ROUND", new[] { "number", "num_digits" } },
            { "ROUNDUP", new[] { "number", "num_digits" } },
            { "ROUNDDOWN", new[] { "number", "num_digits" } }
        };

        private List<TraceItem> MergePrecedents(List<TraceItem> refs, List<ResolvedRange> arrowRefs)
        {
            var merged = new List<TraceItem>();
            merged.AddRange(refs);

            var formulaTokens = BuildFormulaTokenSet(refs);
            var lookup = new Dictionary<string, TraceItem>(StringComparer.OrdinalIgnoreCase);
            foreach (var item in merged)
            {
                foreach (var key in GetItemKeys(item))
                {
                    if (!string.IsNullOrEmpty(key))
                    {
                        lookup[key] = item;
                    }
                }
            }

            foreach (var resolved in arrowRefs)
            {
                foreach (var expanded in ExpandResolvedRangeIfNeeded(resolved, formulaTokens))
                {
                    TraceItem existing = null;
                    foreach (var key in GetResolvedKeys(expanded))
                    {
                        if (!string.IsNullOrEmpty(key) && lookup.TryGetValue(key, out existing))
                        {
                            break;
                        }
                    }

                    if (existing != null)
                    {
                        if (existing.Range == null)
                        {
                            existing.Range = expanded.Range;
                        }
                        if (string.IsNullOrWhiteSpace(existing.DisplayAddress))
                        {
                            existing.DisplayAddress = expanded.DisplayAddress;
                        }
                        if (string.IsNullOrWhiteSpace(existing.ValueText))
                        {
                            existing.ValueText = expanded.ValueText;
                        }
                        continue;
                    }

                    var item = new TraceItem
                    {
                        Token = expanded.Token,
                        StartIndex = -1,
                        Length = 0,
                        Range = expanded.Range,
                        DisplayAddress = expanded.DisplayAddress,
                        ValueText = expanded.ValueText
                    };
                    merged.Add(item);
                    foreach (var key in GetItemKeys(item))
                    {
                        if (!string.IsNullOrEmpty(key))
                        {
                            lookup[key] = item;
                        }
                    }
                }
            }

            return merged;
        }

        private static HashSet<string> BuildFormulaTokenSet(List<TraceItem> refs)
        {
            var tokens = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (refs == null)
            {
                return tokens;
            }

            foreach (var item in refs)
            {
                if (item == null || string.IsNullOrWhiteSpace(item.Token))
                {
                    continue;
                }

                tokens.Add(item.Token);
                if (TryParseQualifiedReference(item.Token, out _, out _, out var address)
                    && !string.IsNullOrWhiteSpace(address))
                {
                    tokens.Add(address);
                }
            }

            return tokens;
        }

        private static bool ContainsFormulaToken(HashSet<string> tokens, ResolvedRange resolved)
        {
            if (tokens == null || tokens.Count == 0 || resolved == null)
            {
                return false;
            }

            if (!string.IsNullOrWhiteSpace(resolved.Token) && tokens.Contains(resolved.Token))
            {
                return true;
            }

            if (!string.IsNullOrWhiteSpace(resolved.DisplayAddress) && tokens.Contains(resolved.DisplayAddress))
            {
                return true;
            }

            if (!string.IsNullOrWhiteSpace(resolved.Token)
                && TryParseQualifiedReference(resolved.Token, out _, out _, out var address)
                && !string.IsNullOrWhiteSpace(address)
                && tokens.Contains(address))
            {
                return true;
            }

            return false;
        }

        private List<ResolvedRange> ExpandResolvedRangeIfNeeded(ResolvedRange resolved, HashSet<string> formulaTokens)
        {
            var expanded = new List<ResolvedRange>();
            if (resolved == null)
            {
                return expanded;
            }

            if (!RangeHelpers.IsRangeValid(resolved.Range))
            {
                expanded.Add(resolved);
                return expanded;
            }

            long count;
            try
            {
                count = Convert.ToInt64(resolved.Range.CountLarge);
            }
            catch
            {
                expanded.Add(resolved);
                return expanded;
            }

            if (count <= 1)
            {
                expanded.Add(resolved);
                return expanded;
            }

            if (formulaTokens == null || formulaTokens.Count == 0)
            {
                expanded.Add(resolved);
                return expanded;
            }

            if (ContainsFormulaToken(formulaTokens, resolved))
            {
                expanded.Add(resolved);
                return expanded;
            }

            if (count > MaxExpandCells)
            {
                expanded.Add(resolved);
                return expanded;
            }

            int added = 0;
            try
            {
                foreach (Excel.Range cell in resolved.Range.Cells)
                {
                    if (CheckTimeout())
                    {
                        break;
                    }

                    if (cell == null)
                    {
                        continue;
                    }

                    expanded.Add(BuildResolved(cell));
                    if (++added >= MaxExpandCells)
                    {
                        break;
                    }
                }
            }
            catch
            {
                if (added == 0)
                {
                    expanded.Add(resolved);
                }
            }

            if (expanded.Count == 0)
            {
                expanded.Add(resolved);
            }

            return expanded;
        }

        private IEnumerable<string> GetItemKeys(TraceItem item)
        {
            var keys = new List<string>();
            if (item == null)
            {
                return keys;
            }

            if (item.Range != null && RangeHelpers.IsRangeValid(item.Range))
            {
                var key = RangeHelpers.BuildRangeKey(item.Range);
                if (!string.IsNullOrEmpty(key))
                {
                    keys.Add(key);
                }

                var local = TryGetAddress(item.Range, false);
                if (!string.IsNullOrEmpty(local))
                {
                    keys.Add(local);
                }

                var external = TryGetAddress(item.Range, true);
                if (!string.IsNullOrEmpty(external))
                {
                    keys.Add(external);
                }
            }

            if (!string.IsNullOrWhiteSpace(item.DisplayAddress))
            {
                keys.Add(item.DisplayAddress);
            }

            if (!string.IsNullOrWhiteSpace(item.Token))
            {
                keys.Add(item.Token);
                if (TryParseQualifiedReference(item.Token, out _, out _, out var address))
                {
                    if (!string.IsNullOrWhiteSpace(address))
                    {
                        keys.Add(address);
                    }
                }
            }

            return keys;
        }

        private IEnumerable<string> GetResolvedKeys(ResolvedRange resolved)
        {
            var keys = new List<string>();
            if (resolved == null)
            {
                return keys;
            }

            if (resolved.Range != null && RangeHelpers.IsRangeValid(resolved.Range))
            {
                var key = RangeHelpers.BuildRangeKey(resolved.Range);
                if (!string.IsNullOrEmpty(key))
                {
                    keys.Add(key);
                }

                var local = TryGetAddress(resolved.Range, false);
                if (!string.IsNullOrEmpty(local))
                {
                    keys.Add(local);
                }

                var external = TryGetAddress(resolved.Range, true);
                if (!string.IsNullOrEmpty(external))
                {
                    keys.Add(external);
                }
            }

            if (!string.IsNullOrWhiteSpace(resolved.DisplayAddress))
            {
                keys.Add(resolved.DisplayAddress);
            }

            if (!string.IsNullOrWhiteSpace(resolved.Token))
            {
                keys.Add(resolved.Token);
                if (TryParseQualifiedReference(resolved.Token, out _, out _, out var address))
                {
                    if (!string.IsNullOrWhiteSpace(address))
                    {
                        keys.Add(address);
                    }
                }
            }

            return keys;
        }

        private string GetDisplayAddress(Excel.Range range)
        {
            if (!RangeHelpers.IsRangeValid(range))
            {
                return null;
            }

            try
            {
                var anchorSheet = _anchorCell?.Worksheet;
                var rangeSheet = range.Worksheet;
                if (anchorSheet != null && rangeSheet != null)
                {
                    var anchorWb = anchorSheet.Parent as Excel.Workbook;
                    var rangeWb = rangeSheet.Parent as Excel.Workbook;
                    var anchorKey = anchorWb?.FullName ?? anchorWb?.Name ?? string.Empty;
                    var rangeKey = rangeWb?.FullName ?? rangeWb?.Name ?? string.Empty;
                    bool sameWorkbook = !string.IsNullOrEmpty(anchorKey)
                        && string.Equals(anchorKey, rangeKey, StringComparison.OrdinalIgnoreCase);
                    bool sameSheet = sameWorkbook
                        && string.Equals(anchorSheet.Name, rangeSheet.Name, StringComparison.OrdinalIgnoreCase);
                    if (sameSheet)
                    {
                        return TryGetAddress(range, false);
                    }

                    if (sameWorkbook)
                    {
                        return FormatSheetAddress(rangeSheet.Name, TryGetAddress(range, false));
                    }

                    return TryGetAddress(range, true);
                }
            }
            catch
            {
                // ignore
            }

            return TryGetAddress(range, true);
        }

        private static string FormatSheetAddress(string sheetName, string address)
        {
            if (string.IsNullOrWhiteSpace(sheetName) || string.IsNullOrWhiteSpace(address))
            {
                return address ?? string.Empty;
            }

            var safeName = sheetName.Replace("'", "''");
            return $"'{safeName}'!{address}";
        }

        private bool IsSameAnchorSheet(string sheetName, string workbookName)
        {
            if (string.IsNullOrWhiteSpace(sheetName))
            {
                return false;
            }

            var anchorSheet = _anchorCell?.Worksheet;
            if (anchorSheet == null)
            {
                return false;
            }

            if (!string.Equals(anchorSheet.Name, sheetName, StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            if (string.IsNullOrWhiteSpace(workbookName))
            {
                return true;
            }

            try
            {
                var anchorWb = anchorSheet.Parent as Excel.Workbook;
                var anchorKey = anchorWb?.FullName ?? anchorWb?.Name ?? string.Empty;
                return string.Equals(anchorKey, workbookName, StringComparison.OrdinalIgnoreCase);
            }
            catch
            {
                return false;
            }
        }

        private static string TryGetAddress(Excel.Range range, bool external)
        {
            if (range == null)
            {
                return null;
            }

            try
            {
                return range.Address[false, false, Excel.XlReferenceStyle.xlA1, external];
            }
            catch
            {
                return null;
            }
        }

        private List<GroupBucket> BuildGroupBuckets(IEnumerable<TraceItem> items)
        {
            var buckets = new List<GroupBucket>();
            var lookup = new Dictionary<string, GroupBucket>(StringComparer.OrdinalIgnoreCase);

            foreach (var item in items)
            {
                var key = GetGroupKey(item, out var label);
                if (!lookup.TryGetValue(key, out var bucket))
                {
                    bucket = new GroupBucket(key, string.IsNullOrWhiteSpace(label) ? "Unresolved" : label);
                    lookup[key] = bucket;
                    buckets.Add(bucket);
                }
                bucket.Items.Add(item);
            }

            return buckets;
        }

        private string GetGroupKey(TraceItem item, out string label)
        {
            label = string.Empty;

            if (item?.Kind == TraceItemKind.Chart && item.Chart != null)
            {
                var sheet = GetChartWorksheet(item.Chart);
                var wb = sheet?.Parent as Excel.Workbook;
                var wbName = wb?.Name ?? string.Empty;
                var sheetName = sheet?.Name ?? string.Empty;
                if (!string.IsNullOrEmpty(sheetName))
                {
                    label = string.IsNullOrEmpty(wbName) ? sheetName : $"[{wbName}]{sheetName}";
                    return $"{wbName}|{sheetName}";
                }
            }

            if (item?.Range != null && RangeHelpers.IsRangeValid(item.Range))
            {
                var ws = item.Range.Worksheet;
                var wb = ws?.Parent as Excel.Workbook;
                var wbName = wb?.Name ?? string.Empty;
                var sheetName = ws?.Name ?? string.Empty;
                if (!string.IsNullOrEmpty(sheetName))
                {
                    label = string.IsNullOrEmpty(wbName) ? sheetName : $"[{wbName}]{sheetName}";
                    return $"{wbName}|{sheetName}";
                }
            }

            var token = item?.DisplayAddress ?? item?.Token;
            if (!string.IsNullOrWhiteSpace(token))
            {
                if (TryParseQualifiedReference(token, out var wbName, out var sheetName, out _))
                {
                    label = string.IsNullOrEmpty(wbName) ? sheetName : $"[{wbName}]{sheetName}";
                    return $"{wbName}|{sheetName}";
                }
            }

            label = "Unresolved";
            return "Unresolved";
        }

        private Excel.Worksheet GetChartWorksheet(Excel.Chart chart)
        {
            if (chart == null)
            {
                return null;
            }

            try
            {
                var chartObject = chart.Parent as Excel.ChartObject;
                if (chartObject != null)
                {
                    return chartObject.Parent as Excel.Worksheet;
                }

                var shape = chart.Parent as Excel.Shape;
                if (shape != null)
                {
                    return shape.Parent as Excel.Worksheet;
                }

                var worksheet = chart.Parent as Excel.Worksheet;
                if (worksheet != null)
                {
                    return worksheet;
                }

                return chart.Parent as Excel.Worksheet;
            }
            catch
            {
                return null;
            }
        }

        private Excel.Range TryResolveNameReference(string token)
        {
            if (string.IsNullOrWhiteSpace(token))
            {
                return null;
            }

            if (token.IndexOf('!') >= 0 || token.IndexOf(':') >= 0 || token.IndexOf('[') >= 0)
            {
                return null;
            }

            try
            {
                var wb = _app.ActiveWorkbook;
                if (wb != null)
                {
                    var name = wb.Names.Item(token, Type.Missing, Type.Missing);
                    var range = name?.RefersToRange;
                    if (RangeHelpers.IsRangeValid(range))
                    {
                        return range;
                    }
                }
            }
            catch
            {
                // ignore
            }

            try
            {
                var ws = _app.ActiveSheet as Excel.Worksheet;
                if (ws != null)
                {
                    var name = ws.Names.Item(token, Type.Missing, Type.Missing);
                    var range = name?.RefersToRange;
                    if (RangeHelpers.IsRangeValid(range))
                    {
                        return range;
                    }
                }
            }
            catch
            {
                // ignore
            }

            return null;
        }

        private static List<string> UnhideHiddenSheets(Excel.Application app)
        {
            var names = new List<string>();
            if (app == null)
            {
                return names;
            }

            var wb = app.ActiveWorkbook;
            if (wb == null)
            {
                return names;
            }

            foreach (Excel.Worksheet sheet in wb.Worksheets)
            {
                try
                {
                    if (sheet.Visible == Excel.XlSheetVisibility.xlSheetHidden)
                    {
                        sheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;
                        names.Add(sheet.Name);
                    }
                }
                catch
                {
                    // ignore
                }
            }

            return names;
        }

        private static void HidePreviouslyHiddenSheets(Excel.Application app, List<string> names)
        {
            if (app == null || names == null || names.Count == 0)
            {
                return;
            }

            var wb = app.ActiveWorkbook;
            if (wb == null)
            {
                return;
            }

            foreach (var name in names)
            {
                try
                {
                    var sheet = wb.Worksheets[name] as Excel.Worksheet;
                    if (sheet != null)
                    {
                        sheet.Visible = Excel.XlSheetVisibility.xlSheetHidden;
                    }
                }
                catch
                {
                    // ignore
                }
            }
        }

        private void RestoreHiddenSheets()
        {
            if (_navUnhiddenSheets.Count == 0)
            {
                return;
            }

            var activeSheet = _app.ActiveSheet as Excel.Worksheet;
            foreach (var sheet in _navUnhiddenSheets.ToArray())
            {
                try
                {
                    if (sheet == null)
                    {
                        continue;
                    }

                    if (activeSheet != null && sheet == activeSheet)
                    {
                        continue;
                    }

                    sheet.Visible = Excel.XlSheetVisibility.xlSheetHidden;
                }
                catch
                {
                    // ignore
                }
            }

            _navUnhiddenSheets.Clear();
        }

        private static bool IsSameAddress(Excel.Range left, Excel.Range right)
        {
            if (left == null || right == null)
            {
                return false;
            }

            try
            {
                var leftAddr = left.Address[true, true, Excel.XlReferenceStyle.xlA1, true];
                var rightAddr = right.Address[true, true, Excel.XlReferenceStyle.xlA1, true];
                return string.Equals(leftAddr, rightAddr, StringComparison.OrdinalIgnoreCase);
            }
            catch
            {
                return false;
            }
        }

        private sealed class TraceItem
        {
            public TraceItemKind Kind { get; set; } = TraceItemKind.Cell;
            public string Token { get; set; }
            public int StartIndex { get; set; } = -1;
            public int Length { get; set; }
            public Excel.Range Range { get; set; }
            public string DisplayAddress { get; set; }
            public string ValueText { get; set; }
            public string ArgumentLabel { get; set; }
            public Color Color { get; set; }
            public string Label { get; set; }
            public Excel.Chart Chart { get; set; }
            public List<TraceItem> Children { get; } = new List<TraceItem>();
            public bool HasDeferredChildren { get; set; }
            public bool ChildrenLoaded { get; set; }
        }

        private sealed class GroupBucket
        {
            public GroupBucket(string key, string label)
            {
                Key = key;
                Label = label;
                Items = new List<TraceItem>();
            }

            public string Key { get; }
            public string Label { get; }
            public List<TraceItem> Items { get; }
        }

        private sealed class FunctionContext
        {
            public FunctionContext(string name, int argIndex)
            {
                Name = name;
                ArgIndex = argIndex;
            }

            public string Name { get; }
            public int ArgIndex { get; set; }
        }

        private sealed class ResolvedRange
        {
            public string Token { get; set; }
            public Excel.Range Range { get; set; }
            public string DisplayAddress { get; set; }
            public string ValueText { get; set; }
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool GetWindowRect(IntPtr hwnd, out RECT rect);
    }

}
