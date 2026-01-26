using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace VantagePackageHolder
{
    internal sealed class NameScrubberForm : Form
    {
        private const string SearchPlaceholder = "Search names...";
        private static readonly Color AccentColor = Color.FromArgb(0, 150, 255);
        private static readonly Color AccentLight = Color.FromArgb(218, 240, 255);
        private static readonly Color PanelBack = Color.FromArgb(245, 247, 250);
        private static readonly Color BorderColor = Color.FromArgb(210, 210, 210);
        private const bool DebugNameNavigation = false;

        private readonly Excel.Application _app;
        private readonly List<NameScrubberItem> _allNames = new List<NameScrubberItem>();
        private List<NameScrubberItem> _filteredNames = new List<NameScrubberItem>();

        private readonly ComboBox _filterCombo;
        private readonly ComboBox _searchModeCombo;
        private readonly TextBox _searchBox;
        private readonly CheckBox _checkAll;
        private readonly CheckBox _showDependents;
        private readonly ListView _namesList;
        private readonly ListView _dependentsList;
        private readonly Label _countLabel;
        private readonly ProgressBar _progress;
        private readonly Button _cancelButton;
        private readonly Button _hideButton;
        private readonly Button _unhideButton;
        private readonly Button _applyButton;
        private readonly Button _unapplyButton;
        private readonly Button _deleteButton;
        private readonly Button _editNameButton;
        private readonly Button _editRefButton;
        private readonly Button _cleanButton;
        private readonly Button _closeButton;
        private readonly Panel _dependentsPanel;
        private readonly TableLayoutPanel _layout;
        private readonly ToolTip _toolTip;

        private bool _suppressCheckEvents;
        private bool _suppressSelectionEvents;
        private bool _cancelRequested;
        private bool _placeholderActive;
        private bool _busy;
        private bool _suppressDeactivate;

        public NameScrubberForm(Excel.Application app)
        {
            _app = app ?? throw new ArgumentNullException(nameof(app));

            Text = "Name Scrubber";
            StartPosition = FormStartPosition.CenterScreen;
            Size = new Size(860, 560);
            MinimumSize = new Size(720, 420);
            FormBorderStyle = FormBorderStyle.SizableToolWindow;
            ShowInTaskbar = false;
            KeyPreview = true;
            Font = new Font("Segoe UI", 9f, FontStyle.Regular);
            BackColor = PanelBack;

            _filterCombo = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Width = 140
            };
            _filterCombo.Items.AddRange(new object[]
            {
                "All",
                "Visible",
                "Hidden",
                "Erroneous",
                "Linked",
                "Unused",
                "Lambda",
                "Scope: Workbook (Global)",
                "Scope: Sheet"
            });

            _searchModeCombo = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Width = 110
            };
            _searchModeCombo.Items.AddRange(new object[] { "Starts with", "Contains" });

            _searchBox = new TextBox { Width = 220 };

            _checkAll = new CheckBox { Text = "Check all", AutoSize = true };

            _showDependents = new CheckBox { Text = "Show dependents", AutoSize = true, Checked = true };

            _namesList = new ListView
            {
                View = View.Details,
                FullRowSelect = true,
                MultiSelect = true,
                CheckBoxes = true,
                HideSelection = false,
                Activation = ItemActivation.Standard,
                Dock = DockStyle.Fill,
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                GridLines = true,
                HeaderStyle = ColumnHeaderStyle.Nonclickable
            };
            _namesList.Columns.Add("Name", 160);
            _namesList.Columns.Add("Value", 160);
            _namesList.Columns.Add("Scope", 120);
            _namesList.Columns.Add("Refers To", 320);

            _dependentsList = new ListView
            {
                View = View.Details,
                FullRowSelect = true,
                MultiSelect = false,
                HideSelection = false,
                Dock = DockStyle.Fill,
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                GridLines = true,
                HeaderStyle = ColumnHeaderStyle.Nonclickable
            };
            _dependentsList.Columns.Add("Cell", 140);
            _dependentsList.Columns.Add("Formula", 520);

            _countLabel = new Label { AutoSize = true, Text = "0 names", ForeColor = Color.FromArgb(90, 90, 90) };

            _progress = new ProgressBar { Visible = false, Width = 200 };

            _cancelButton = new Button { Text = "Cancel", Visible = false, AutoSize = true };

            _hideButton = new Button { Text = "Hide" };
            _unhideButton = new Button { Text = "Unhide" };
            _applyButton = new Button { Text = "Apply" };
            _unapplyButton = new Button { Text = "Unapply" };
            _deleteButton = new Button { Text = "Delete" };
            _editNameButton = new Button { Text = "Edit Name" };
            _editRefButton = new Button { Text = "\u2206 Ref" };
            _cleanButton = new Button { Text = "Clean" };
            _closeButton = new Button { Text = "Close" };
            _toolTip = new ToolTip();
            _toolTip.SetToolTip(_applyButton, "Replace cell references in dependent formulas with the selected name.");
            _toolTip.SetToolTip(_unapplyButton, "Replace the selected name in dependent formulas with cell references.");

            StyleActionButton(_hideButton);
            StyleActionButton(_unhideButton);
            StyleActionButton(_applyButton);
            StyleActionButton(_unapplyButton);
            StyleActionButton(_deleteButton);
            StyleActionButton(_editNameButton);
            StyleActionButton(_editRefButton);
            StyleActionButton(_cleanButton);
            StyleActionButton(_closeButton);
            CancelButton = _closeButton;

            var topPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                AutoSize = true,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                BackColor = PanelBack
            };
            topPanel.Controls.Add(new Label { Text = "Filter:", AutoSize = true, Padding = new Padding(0, 6, 0, 0) });
            topPanel.Controls.Add(_filterCombo);
            topPanel.Controls.Add(new Label { Text = "Search:", AutoSize = true, Padding = new Padding(8, 6, 0, 0) });
            topPanel.Controls.Add(_searchModeCombo);
            topPanel.Controls.Add(_searchBox);
            topPanel.Controls.Add(_checkAll);
            topPanel.Controls.Add(_showDependents);

            _dependentsPanel = new Panel { Dock = DockStyle.Fill, BackColor = PanelBack, Padding = new Padding(0, 4, 0, 0) };
            var dependentsLabel = new Label
            {
                Text = "Dependents",
                AutoSize = true,
                Dock = DockStyle.Top,
                ForeColor = AccentColor
            };
            _dependentsPanel.Controls.Add(_dependentsList);
            _dependentsPanel.Controls.Add(dependentsLabel);

            var bottomPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                AutoSize = true,
                ColumnCount = 2,
                RowCount = 1
            };
            bottomPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            bottomPanel.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

            var leftPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                AutoSize = true
            };
            leftPanel.Controls.Add(_countLabel);
            leftPanel.Controls.Add(new Label { Width = 20 });
            leftPanel.Controls.Add(_progress);
            leftPanel.Controls.Add(_cancelButton);

            var rightPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.RightToLeft,
                WrapContents = false,
                AutoSize = true
            };
            rightPanel.Controls.AddRange(new Control[]
            {
                _closeButton, _cleanButton, _editRefButton, _editNameButton, _deleteButton,
                _unapplyButton, _applyButton, _unhideButton, _hideButton
            });

            bottomPanel.Controls.Add(leftPanel, 0, 0);
            bottomPanel.Controls.Add(rightPanel, 1, 0);

            _layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                RowCount = 4,
                ColumnCount = 1
            };
            _layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            _layout.RowStyles.Add(new RowStyle(SizeType.Percent, 60));
            _layout.RowStyles.Add(new RowStyle(SizeType.Percent, 40));
            _layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            _layout.Controls.Add(topPanel, 0, 0);
            _layout.Controls.Add(_namesList, 0, 1);
            _layout.Controls.Add(_dependentsPanel, 0, 2);
            _layout.Controls.Add(bottomPanel, 0, 3);

            Controls.Add(_layout);

            _filterCombo.SelectedIndexChanged += (_, __) => ApplyFilterAndSearch();
            _searchModeCombo.SelectedIndexChanged += (_, __) => ApplyFilterAndSearch();
            _searchBox.TextChanged += (_, __) =>
            {
                if (_placeholderActive)
                {
                    return;
                }
                ApplyFilterAndSearch();
            };
            _searchBox.GotFocus += (_, __) => ClearSearchPlaceholder();
            _searchBox.LostFocus += (_, __) => EnsureSearchPlaceholder();

            _checkAll.CheckedChanged += (_, __) => ToggleCheckAll();
            _showDependents.CheckedChanged += (_, __) => ToggleDependentsPanel();

            _namesList.ItemChecked += NamesList_ItemChecked;
            _namesList.SelectedIndexChanged += NamesList_SelectedIndexChanged;
            _namesList.DoubleClick += (_, __) => NavigateToSelectedName();
            _namesList.MouseDown += NamesList_MouseDown;
            _dependentsList.SelectedIndexChanged += DependentsList_SelectedIndexChanged;

            _hideButton.Click += (_, __) => RunCommand(() => HideSelected(true));
            _unhideButton.Click += (_, __) => RunCommand(() => HideSelected(false));
            _applyButton.Click += (_, __) => RunCommand(ApplySelected);
            _unapplyButton.Click += (_, __) => RunCommand(UnapplySelected);
            _deleteButton.Click += (_, __) => RunCommand(DeleteSelected);
            _editNameButton.Click += (_, __) => RunCommand(EditSelectedName);
            _editRefButton.Click += (_, __) => RunCommand(EditSelectedRefersTo);
            _cleanButton.Click += (_, __) => RunCommand(CleanNames);
            _closeButton.Click += (_, __) => Close();
            _cancelButton.Click += (_, __) => _cancelRequested = true;

            Load += (_, __) =>
            {
                _filterCombo.SelectedIndex = 1;
                _searchModeCombo.SelectedIndex = 0;
                SetSearchPlaceholder();
                ReloadNames();
            };
        }

        private void SetSearchPlaceholder()
        {
            _placeholderActive = true;
            _searchBox.ForeColor = SystemColors.GrayText;
            _searchBox.Text = SearchPlaceholder;
        }

        private void ClearSearchPlaceholder()
        {
            if (!_placeholderActive)
            {
                return;
            }

            _placeholderActive = false;
            _searchBox.ForeColor = SystemColors.WindowText;
            _searchBox.Text = string.Empty;
        }

        private void EnsureSearchPlaceholder()
        {
            if (!string.IsNullOrWhiteSpace(_searchBox.Text))
            {
                return;
            }

            SetSearchPlaceholder();
        }

        private void ToggleDependentsPanel()
        {
            _dependentsPanel.Visible = _showDependents.Checked;
            _layout.RowStyles[2].Height = _showDependents.Checked ? 40 : 0;
            _layout.RowStyles[2].SizeType = _showDependents.Checked ? SizeType.Percent : SizeType.Absolute;
            if (_showDependents.Checked)
            {
                LoadDependentsFromSelection();
            }
            else
            {
                _dependentsList.Items.Clear();
            }
        }

        private void NamesList_ItemChecked(object sender, ItemCheckedEventArgs e)
        {
            if (_suppressCheckEvents)
            {
                return;
            }

            if (e.Item?.Tag is NameScrubberItem item)
            {
                item.IsChecked = e.Item.Checked;
            }

            UpdateCheckAllState();
            UpdateActionButtons();
        }

        private void NamesList_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (_suppressSelectionEvents)
            {
                return;
            }

            LoadDependentsFromSelection();
        }

        private void NamesList_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Left)
            {
                return;
            }

            if (e.Clicks != 1)
            {
                return;
            }

            var hit = _namesList.HitTest(e.Location);
            if (hit.Item == null)
            {
                return;
            }

            if ((ModifierKeys & (Keys.Control | Keys.Shift)) != Keys.None)
            {
                return;
            }

            int checkboxWidth = 18;
            if (e.X <= checkboxWidth)
            {
                return;
            }

            _suppressCheckEvents = true;
            hit.Item.Checked = !hit.Item.Checked;
            if (hit.Item.Tag is NameScrubberItem nameItem)
            {
                nameItem.IsChecked = hit.Item.Checked;
            }
            _suppressCheckEvents = false;
            UpdateCheckAllState();
            UpdateActionButtons();
        }

        private void DependentsList_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (_dependentsList.SelectedItems.Count == 0)
            {
                return;
            }

            if (!(_dependentsList.SelectedItems[0].Tag is NameDependentItem dep))
            {
                return;
            }

            SelectRange(dep.Range, true);
        }

        private void ReloadNames()
        {
            _cancelRequested = false;
            SetBusy(true, "Loading names...");

            _allNames.Clear();
            _filteredNames.Clear();
            _namesList.Items.Clear();

            var wb = _app.ActiveWorkbook;
            if (wb == null)
            {
                SetBusy(false, null);
                return;
            }

            int count = 0;
            try { count = wb.Names.Count; } catch { count = 0; }
            _progress.Maximum = Math.Max(1, count);
            _progress.Value = 0;

            for (int i = 1; i <= count; i++)
            {
                if (_cancelRequested)
                {
                    break;
                }

                Excel.Name name = null;
                try
                {
                    name = wb.Names.Item(i);
                }
                catch
                {
                    continue;
                }

                if (name == null)
                {
                    continue;
                }

                var item = BuildItem(name);
                _allNames.Add(item);

                if (i % 25 == 0 || i == count)
                {
                    ReportProgress(i, count);
                }
            }

            SetBusy(false, null);
            ApplyFilterAndSearch();
        }

        private NameScrubberItem BuildItem(Excel.Name name)
        {
            var item = new NameScrubberItem(name);
            var label = name.Name ?? string.Empty;
            int bang = label.LastIndexOf('!');
            if (bang >= 0 && bang + 1 < label.Length)
            {
                label = label.Substring(bang + 1);
            }

            item.Label = label;
            item.RefersTo = NameScrubberUtil.SafeRefersTo(name);
            item.ParentName = name.Parent is Excel.Workbook ? "Workbook" : (name.Parent as Excel.Worksheet)?.Name ?? string.Empty;

            item.Text = BuildValueText(name);

            return item;
        }

        private string BuildValueText(Excel.Name name)
        {
            if (name == null)
            {
                return string.Empty;
            }

            if (!name.Visible)
            {
                return "(hidden)";
            }

            if (NameScrubberUtil.IsLinked(name))
            {
                return string.Empty;
            }

            Excel.Range refersToRange = null;
            try
            {
                refersToRange = name.RefersToRange;
            }
            catch
            {
                refersToRange = null;
            }

            if (refersToRange == null)
            {
                return string.Empty;
            }

            bool isArray = false;
            try
            {
                isArray = Convert.ToInt64(refersToRange.Cells.CountLarge) > 1;
            }
            catch
            {
                isArray = false;
            }

            if (!isArray)
            {
                try
                {
                    isArray = Convert.ToBoolean(refersToRange.HasArray);
                }
                catch
                {
                    isArray = false;
                }
            }

            if (isArray)
            {
                return "<array>";
            }

            try
            {
                return refersToRange.Text?.ToString() ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private void ApplyFilterAndSearch()
        {
            if (_busy)
            {
                return;
            }

            _cancelRequested = false;
            var filtered = new List<NameScrubberItem>();
            int filterIndex = _filterCombo.SelectedIndex;

            bool useProgress = filterIndex == 5;
            if (useProgress)
            {
                SetBusy(true, "Checking dependents...");
                _progress.Maximum = Math.Max(1, _allNames.Count);
                _progress.Value = 0;
            }

            int index = 0;
            foreach (var item in _allNames)
            {
                index++;
                if (_cancelRequested)
                {
                    break;
                }

                if (useProgress && (index % 10 == 0 || index == _allNames.Count))
                {
                    ReportProgress(index, _allNames.Count);
                }

                if (!FilterMatch(item, filterIndex))
                {
                    continue;
                }

                filtered.Add(item);
            }

            if (useProgress)
            {
                SetBusy(false, null);
            }

            string searchText = _placeholderActive ? string.Empty : _searchBox.Text.Trim().ToLowerInvariant();
            if (!string.IsNullOrEmpty(searchText))
            {
                bool contains = _searchModeCombo.SelectedIndex == 1;
                filtered = filtered.Where(item =>
                {
                    var label = (item.Label ?? string.Empty).ToLowerInvariant();
                    bool match = contains ? label.Contains(searchText) : label.StartsWith(searchText);
                    if (!match)
                    {
                        item.IsChecked = false;
                    }
                    return match;
                }).ToList();
            }

            _filteredNames = filtered;
            RebuildList();
        }

        private bool FilterMatch(NameScrubberItem item, int filterIndex)
        {
            switch (filterIndex)
            {
                case 1:
                    return item.Name != null && item.Name.Visible;
                case 2:
                    return item.Name != null && !item.Name.Visible;
                case 3:
                    if (!item.CachedIsErroneous.HasValue)
                    {
                        item.CachedIsErroneous = NameScrubberUtil.IsErroneous(item.Name);
                    }
                    return item.CachedIsErroneous.Value;
                case 4:
                    if (!item.CachedIsLinked.HasValue)
                    {
                        item.CachedIsLinked = NameScrubberUtil.IsLinked(item.Name);
                    }
                    return item.CachedIsLinked.Value;
                case 5:
                    if (!item.CachedIsLambda.HasValue)
                    {
                        item.CachedIsLambda = NameScrubberUtil.IsLambda(item.Name);
                    }
                    if (item.CachedIsLambda.Value)
                    {
                        return false;
                    }
                    if (!item.CachedHasDependents.HasValue)
                    {
                        item.CachedHasDependents = NameScrubberUtil.HasDependents(item.Name, () => _cancelRequested);
                    }
                    return !item.CachedHasDependents.Value;
                case 6:
                    if (!item.CachedIsLambda.HasValue)
                    {
                        item.CachedIsLambda = NameScrubberUtil.IsLambda(item.Name);
                    }
                    return item.CachedIsLambda.Value;
                case 7:
                    return item.Name?.Parent is Excel.Workbook;
                case 8:
                    return item.Name?.Parent is Excel.Worksheet;
                default:
                    return true;
            }
        }

        private void RebuildList()
        {
            _suppressCheckEvents = true;
            _suppressSelectionEvents = true;

            _namesList.BeginUpdate();
            _namesList.Items.Clear();

            foreach (var item in _filteredNames)
            {
                var listItem = new ListViewItem(item.Label ?? string.Empty)
                {
                    Tag = item,
                    Checked = item.IsChecked
                };
                listItem.SubItems.Add(item.Text ?? string.Empty);
                listItem.SubItems.Add(item.ParentName ?? string.Empty);
                listItem.SubItems.Add(item.RefersTo ?? string.Empty);
                if (item.Name != null && !item.Name.Visible)
                {
                    listItem.ForeColor = Color.Gray;
                }
                _namesList.Items.Add(listItem);
            }

            _namesList.EndUpdate();
            _suppressCheckEvents = false;
            _suppressSelectionEvents = false;

            UpdateCheckAllState();
            UpdateActionButtons();
            UpdateCountLabel();
            LoadDependentsFromSelection();
        }

        private void UpdateCountLabel()
        {
            _countLabel.Text = string.Format("{0:N0} names", _filteredNames.Count);
        }

        private void UpdateCheckAllState()
        {
            if (_suppressCheckEvents)
            {
                return;
            }

            _suppressCheckEvents = true;
            if (_filteredNames.Count == 0)
            {
                _checkAll.CheckState = CheckState.Unchecked;
            }
            else
            {
                int checkedCount = _filteredNames.Count(item => item.IsChecked);
                if (checkedCount == 0)
                {
                    _checkAll.CheckState = CheckState.Unchecked;
                }
                else if (checkedCount == _filteredNames.Count)
                {
                    _checkAll.CheckState = CheckState.Checked;
                }
                else
                {
                    _checkAll.CheckState = CheckState.Indeterminate;
                }
            }
            _suppressCheckEvents = false;
        }

        private void ToggleCheckAll()
        {
            if (_suppressCheckEvents)
            {
                return;
            }

            bool target = _checkAll.CheckState == CheckState.Checked;
            _suppressCheckEvents = true;
            foreach (ListViewItem item in _namesList.Items)
            {
                item.Checked = target;
                if (item.Tag is NameScrubberItem nameItem)
                {
                    nameItem.IsChecked = target;
                }
            }
            _suppressCheckEvents = false;
            UpdateActionButtons();
        }

        private void UpdateActionButtons()
        {
            if (_busy)
            {
                DisableActionButtons();
                return;
            }

            var checkedItems = GetCheckedItems();
            var actionItems = checkedItems.Count > 0 ? checkedItems : GetSelectedItems();
            int count = actionItems.Count;
            if (count == 0)
            {
                DisableActionButtons();
                return;
            }

            if (count == 1)
            {
                var item = actionItems[0];
                bool isNative = NameScrubberUtil.IsNative(item.Name);
                if (isNative)
                {
                    _hideButton.Enabled = false;
                    _unhideButton.Enabled = false;
                    _applyButton.Enabled = false;
                    _unapplyButton.Enabled = false;
                }
                else
                {
                    _hideButton.Enabled = item.Name != null && item.Name.Visible;
                    _unhideButton.Enabled = item.Name != null && !item.Name.Visible;
                    _applyButton.Enabled = true;
                    _unapplyButton.Enabled = true;
                }

                _deleteButton.Enabled = true;
                _editNameButton.Enabled = true;
                _editRefButton.Enabled = true;
            }
            else
            {
                _hideButton.Enabled = true;
                _unhideButton.Enabled = true;
                _applyButton.Enabled = true;
                _unapplyButton.Enabled = true;
                _deleteButton.Enabled = true;
                _editNameButton.Enabled = false;
                _editRefButton.Enabled = false;
            }
        }

        private void DisableActionButtons()
        {
            _hideButton.Enabled = false;
            _unhideButton.Enabled = false;
            _applyButton.Enabled = false;
            _unapplyButton.Enabled = false;
            _deleteButton.Enabled = false;
            _editNameButton.Enabled = false;
            _editRefButton.Enabled = false;
        }

        private List<NameScrubberItem> GetCheckedItems()
        {
            return _filteredNames.Where(item => item.IsChecked).ToList();
        }

        private List<NameScrubberItem> GetSelectedItems()
        {
            var items = new List<NameScrubberItem>();
            foreach (ListViewItem listItem in _namesList.SelectedItems)
            {
                if (listItem.Tag is NameScrubberItem item)
                {
                    items.Add(item);
                }
            }
            return items;
        }

        private List<NameScrubberItem> GetActionItems()
        {
            var checkedItems = GetCheckedItems();
            if (checkedItems.Count > 0)
            {
                return checkedItems;
            }

            return GetSelectedItems();
        }

        private void LoadDependentsFromSelection()
        {
            if (!_showDependents.Checked)
            {
                _dependentsList.Items.Clear();
                return;
            }

            if (_namesList.SelectedItems.Count != 1)
            {
                _dependentsList.Items.Clear();
                return;
            }

            if (!(_namesList.SelectedItems[0].Tag is NameScrubberItem item))
            {
                _dependentsList.Items.Clear();
                return;
            }

            if (item.Name == null || NameScrubberUtil.IsLinked(item.Name))
            {
                _dependentsList.Items.Clear();
                return;
            }

            var refersToRange = NameScrubberUtil.TryGetRefersToRange(item.Name);
            if (refersToRange == null)
            {
                _dependentsList.Items.Clear();
                return;
            }

            var deps = new List<NameDependentItem>();
            Excel.Range originalSelection = null;
            Excel.Window activeWindow = null;
            int? originalScrollRow = null;
            int? originalScrollCol = null;
            try
            {
                originalSelection = _app.Selection as Excel.Range;
            }
            catch
            {
                originalSelection = null;
            }

            try
            {
                activeWindow = _app.ActiveWindow;
                if (activeWindow != null)
                {
                    originalScrollRow = activeWindow.ScrollRow;
                    originalScrollCol = activeWindow.ScrollColumn;
                }
            }
            catch
            {
                activeWindow = null;
            }

            using (new UiGuard(_app, hideStatusBar: true))
            {
                var ranges = NameScrubberUtil.GetDependents(refersToRange, item.Name, () => _cancelRequested);
                foreach (var range in ranges)
                {
                    if (!NameScrubberUtil.FormulaReferencesName(range, item.Name))
                    {
                        continue;
                    }

                    var label = GetRangeLabel(range);
                    string formula = string.Empty;
                    try
                    {
                        formula = range.Formula?.ToString() ?? string.Empty;
                    }
                    catch
                    {
                        formula = string.Empty;
                    }

                    deps.Add(new NameDependentItem(range, label, formula));
                }
            }

            RestoreSelection(originalSelection, activeWindow, originalScrollRow, originalScrollCol);

            _dependentsList.BeginUpdate();
            _dependentsList.Items.Clear();
            foreach (var dep in deps)
            {
                var listItem = new ListViewItem(dep.Label) { Tag = dep };
                listItem.SubItems.Add(dep.Formula ?? string.Empty);
                _dependentsList.Items.Add(listItem);
            }
            _dependentsList.EndUpdate();
        }

        private void RestoreSelection(Excel.Range selection, Excel.Window window, int? scrollRow, int? scrollCol)
        {
            if (selection == null || !RangeHelpers.IsRangeValid(selection))
            {
                return;
            }

            bool prevSuppress = _suppressDeactivate;
            _suppressDeactivate = true;
            try
            {
                RangeHelpers.SafeActivateSheet(selection.Worksheet);
                RangeHelpers.SafeSelect(selection);

                if (window != null)
                {
                    if (scrollRow.HasValue)
                    {
                        window.ScrollRow = scrollRow.Value;
                    }
                    if (scrollCol.HasValue)
                    {
                        window.ScrollColumn = scrollCol.Value;
                    }
                }
            }
            catch
            {
                // ignore
            }
            finally
            {
                _suppressDeactivate = prevSuppress;
            }
        }

        private void RunCommand(Action action)
        {
            if (action == null || IsDisposed || Disposing)
            {
                return;
            }

            bool prevSuppress = _suppressDeactivate;
            _suppressDeactivate = true;
            try
            {
                action();
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message, "Name Scrubber", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                _suppressDeactivate = prevSuppress;
            }
        }

        private void NavigateToSelectedName()
        {
            if (IsDisposed || Disposing)
            {
                return;
            }

            bool prevSuppress = _suppressDeactivate;
            _suppressDeactivate = true;
            try
            {
                if (_namesList.SelectedItems.Count != 1)
                {
                    return;
                }

                if (!(_namesList.SelectedItems[0].Tag is NameScrubberItem item))
                {
                    return;
                }

                if (item.Name == null || NameScrubberUtil.IsLinked(item.Name))
                {
                    return;
                }

                var refersTo = NameScrubberUtil.SafeRefersTo(item.Name);
                if (string.IsNullOrWhiteSpace(refersTo))
                {
                    MessageBox.Show(this, "No attributes to name.", "Name Scrubber", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                var range = NameScrubberUtil.TryGetRefersToRange(item.Name);
                if (range == null)
                {
                    MessageBox.Show(this, "No attributes to name.", "Name Scrubber", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                DebugNavigate("Navigate name: " + item.Label + "\nRefersTo: " + refersTo + "\nRefersToRange: " + DescribeRange(range));
                if (TryGotoRefersTo(refersTo, true, range))
                {
                    return;
                }

                if (!RangeHelpers.IsRangeValid(range))
                {
                    return;
                }

                SelectRange(range, true);
            }
            finally
            {
                _suppressDeactivate = prevSuppress;
            }
        }

        private void NavigateToRange(Excel.Range range)
        {
            if (range == null)
            {
                return;
            }

            SelectRange(range, true);
        }

        private void SelectRange(Excel.Range range, bool center)
        {
            bool prevSuppress = _suppressDeactivate;
            _suppressDeactivate = true;
            try
            {
                if (range == null)
                {
                    return;
                }

                var targetRange = GetSelectionRange(range) ?? range;
                if (!RangeHelpers.IsRangeValid(targetRange))
                {
                    return;
                }
                Excel.Workbook wb = null;
                try
                {
                    wb = targetRange.Worksheet?.Parent as Excel.Workbook;
                }
                catch
                {
                    wb = null;
                }

                if (wb != null)
                {
                    try
                    {
                        wb.Activate();
                    }
                    catch
                    {
                        // ignore
                    }
                }

                RangeHelpers.SafeActivateSheet(targetRange.Worksheet);
                bool selected = false;
                try
                {
                    targetRange.Application.Goto(targetRange, true);
                    selected = true;
                }
                catch
                {
                    selected = false;
                }

                if (!selected)
                {
                    try
                    {
                        targetRange.Select();
                        selected = true;
                    }
                    catch
                    {
                        selected = false;
                    }
                }

                if (!selected)
                {
                    RangeHelpers.SafeSelect(targetRange);
                }

                Excel.Range selectedRange = null;
                try
                {
                    selectedRange = _app.Selection as Excel.Range;
                }
                catch
                {
                    selectedRange = null;
                }

                if (selectedRange != null && RangeHelpers.IsRangeValid(selectedRange))
                {
                    var expandedSelection = GetSelectionRange(selectedRange) ?? selectedRange;
                    var finalRange = targetRange;
                    if (IsRangeLarger(expandedSelection, finalRange))
                    {
                        finalRange = expandedSelection;
                    }

                    if (IsRangeLarger(selectedRange, finalRange))
                    {
                        finalRange = selectedRange;
                    }

                    if (IsRangeLarger(finalRange, selectedRange))
                    {
                        RangeHelpers.SafeSelect(finalRange);
                    }

                    targetRange = finalRange;
                }

                if (!center)
                {
                    return;
                }

                Application.DoEvents();
                CenterOnRange(targetRange);
            }
            catch
            {
                // ignore
            }
            finally
            {
                _suppressDeactivate = prevSuppress;
            }
        }

        private Excel.Range GetSelectionRange(Excel.Range range)
        {
            if (range == null)
            {
                return null;
            }

            var selection = range;
            selection = TryGetSpillRange(selection) ?? selection;
            try
            {
                if (Convert.ToBoolean(selection.MergeCells))
                {
                    selection = selection.MergeArea;
                }
            }
            catch
            {
                // ignore
            }

            bool hasArray = false;
            try
            {
                hasArray = Convert.ToBoolean(selection.HasArray);
            }
            catch
            {
                hasArray = false;
            }

            if (hasArray)
            {
                try
                {
                    selection = selection.CurrentArray;
                }
                catch
                {
                    // ignore
                }
            }

            return selection;
        }

        private Excel.Range TryGetSpillRange(Excel.Range range)
        {
            if (range == null)
            {
                return null;
            }

            try
            {
                var obj = range.GetType().InvokeMember("SpillRange", BindingFlags.GetProperty, null, range, null);
                return obj as Excel.Range;
            }
            catch
            {
                return null;
            }
        }

        private bool IsRangeLarger(Excel.Range candidate, Excel.Range baseline)
        {
            if (candidate == null || baseline == null)
            {
                return false;
            }

            try
            {
                long candCount = Convert.ToInt64(candidate.Cells.CountLarge);
                long baseCount = Convert.ToInt64(baseline.Cells.CountLarge);
                return candCount > baseCount;
            }
            catch
            {
                return false;
            }
        }

        private void CenterOnRange(Excel.Range range)
        {
            try
            {
                if (range == null)
                {
                    return;
                }

                var win = _app.ActiveWindow;
                if (win == null)
                {
                    return;
                }

                Excel.Range anchor = null;
                try
                {
                    int minRow = int.MaxValue;
                    int minCol = int.MaxValue;
                    int maxRow = 1;
                    int maxCol = 1;
                    foreach (Excel.Range area in range.Areas)
                    {
                        int areaRow = area.Row;
                        int areaCol = area.Column;
                        int areaMaxRow = areaRow + area.Rows.Count - 1;
                        int areaMaxCol = areaCol + area.Columns.Count - 1;
                        if (areaRow < minRow) minRow = areaRow;
                        if (areaCol < minCol) minCol = areaCol;
                        if (areaMaxRow > maxRow) maxRow = areaMaxRow;
                        if (areaMaxCol > maxCol) maxCol = areaMaxCol;
                    }

                    int centerRow = minRow + Math.Max(0, (maxRow - minRow) / 2);
                    int centerCol = minCol + Math.Max(0, (maxCol - minCol) / 2);
                    anchor = range.Worksheet.Cells[centerRow, centerCol] as Excel.Range;
                }
                catch
                {
                    anchor = null;
                }

                if (anchor == null)
                {
                    return;
                }

                var visible = win.VisibleRange;
                int visibleRows = visible.Rows.Count;
                int visibleCols = visible.Columns.Count;
                int targetRow = anchor.Row - visibleRows / 2;
                int targetCol = anchor.Column - visibleCols / 2;

                if (targetRow < 1) targetRow = 1;
                if (targetCol < 1) targetCol = 1;

                win.ScrollRow = targetRow;
                win.ScrollColumn = targetCol;
            }
            catch
            {
                // ignore
            }
        }

        private bool TryGotoRefersTo(string refersTo, bool center)
        {
            return TryGotoRefersTo(refersTo, center, null);
        }

        private bool TryGotoRefersTo(string refersTo, bool center, Excel.Range expectedRange)
        {
            if (string.IsNullOrWhiteSpace(refersTo))
            {
                return false;
            }

            var target = refersTo.Trim();
            if (target.StartsWith("=", StringComparison.Ordinal))
            {
                target = target.Substring(1);
            }

            var resolvedRange = ResolveRangeTarget(target, expectedRange);
            if (resolvedRange != null)
            {
                try
                {
                    var wb = resolvedRange.Worksheet?.Parent as Excel.Workbook;
                    if (wb != null)
                    {
                        wb.Activate();
                    }
                    RangeHelpers.SafeActivateSheet(resolvedRange.Worksheet);
                    DebugNavigate("Goto range target: " + DescribeRange(resolvedRange));
                    _app.Goto(resolvedRange, true);
                }
                catch (Exception ex)
                {
                    DebugNavigate("Goto range failed: " + ex.Message);
                    resolvedRange = null;
                }
            }

            if (resolvedRange == null)
            {
                try
                {
                    DebugNavigate("Goto target: " + target);
                    _app.Goto(target, true);
                }
                catch (Exception ex)
                {
                    DebugNavigate("Goto failed: " + ex.Message);
                    return false;
                }
            }

            Excel.Range selection = null;
            try
            {
                selection = _app.Selection as Excel.Range;
            }
            catch
            {
                selection = null;
            }

            if (selection == null || !RangeHelpers.IsRangeValid(selection))
            {
                DebugNavigate("Goto selection invalid.");
                return false;
            }

            var expanded = GetSelectionRange(selection) ?? selection;
            if (IsRangeLarger(expanded, selection))
            {
                RangeHelpers.SafeSelect(expanded);
                selection = expanded;
            }

            if (expectedRange != null && RangeHelpers.IsRangeValid(expectedRange))
            {
                var expectedExpanded = GetSelectionRange(expectedRange) ?? expectedRange;
                if (IsRangeLarger(expectedExpanded, selection))
                {
                    RangeHelpers.SafeSelect(expectedExpanded);
                    selection = expectedExpanded;
                }
            }

            DebugNavigate("Goto selection: " + DescribeRange(selection) + "\nExpected: " + DescribeRange(expectedRange));
            if (center)
            {
                Application.DoEvents();
                CenterOnRange(selection);
            }
            return true;
        }

        private Excel.Range ResolveRangeTarget(string target, Excel.Range expectedRange)
        {
            if (expectedRange != null && RangeHelpers.IsRangeValid(expectedRange))
            {
                return GetSelectionRange(expectedRange) ?? expectedRange;
            }

            if (string.IsNullOrWhiteSpace(target))
            {
                return null;
            }

            try
            {
                var eval = _app.Evaluate(target);
                if (eval is Excel.Range evalRange)
                {
                    return GetSelectionRange(evalRange) ?? evalRange;
                }
            }
            catch
            {
                // ignore
            }

            try
            {
                var range = _app.Range[target];
                return GetSelectionRange(range) ?? range;
            }
            catch
            {
                return null;
            }
        }

        private void DebugNavigate(string message)
        {
            if (!DebugNameNavigation)
            {
                return;
            }

            if (IsDisposed || Disposing)
            {
                return;
            }

            MessageBox.Show(this, message, "Name Scrubber Debug", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private string DescribeRange(Excel.Range range)
        {
            if (range == null)
            {
                return "<null>";
            }

            try
            {
                var sheet = range.Worksheet?.Name ?? "?";
                var addr = range.Address[false, false, Excel.XlReferenceStyle.xlA1];
                var rows = Convert.ToInt64(range.Rows.Count);
                var cols = Convert.ToInt64(range.Columns.Count);
                return sheet + "!" + addr + " (" + rows + "x" + cols + ")";
            }
            catch
            {
                return "<invalid>";
            }
        }

        private string GetRangeLabel(Excel.Range range)
        {
            if (range == null)
            {
                return string.Empty;
            }

            try
            {
                var wb = range.Worksheet?.Parent as Excel.Workbook;
                var activeWb = _app.ActiveWorkbook;
                bool external = wb != null && activeWb != null &&
                                !string.Equals(wb.FullName, activeWb.FullName, StringComparison.OrdinalIgnoreCase);

                if (external)
                {
                    return range.Address[false, false, Excel.XlReferenceStyle.xlA1, true];
                }

                var activeSheet = _app.ActiveSheet as Excel.Worksheet;
                if (activeSheet != null && !string.Equals(activeSheet.Name, range.Worksheet.Name, StringComparison.OrdinalIgnoreCase))
                {
                    return $"{range.Worksheet.Name}!{range.Address[false, false]}";
                }

                return range.Address[false, false];
            }
            catch
            {
                return string.Empty;
            }
        }

        private void HideSelected(bool hide)
        {
            var items = GetActionItems();
            if (items.Count == 0)
            {
                return;
            }

            foreach (var item in items)
            {
                try
                {
                    if (item.Name != null)
                    {
                        item.Name.Visible = !hide;
                    }
                }
                catch
                {
                    // ignore
                }
            }

            ApplyFilterAndSearch();
        }

        private void ApplySelected()
        {
            var items = GetActionItems();
            if (items.Count == 0)
            {
                return;
            }

            _cancelRequested = false;
            SetBusy(true, "Applying names...");

            var changed = new List<Excel.Range>();
            var prevCalc = _app.Calculation;
            var prevAlerts = _app.DisplayAlerts;
            _app.DisplayAlerts = false;
            _app.Calculation = Excel.XlCalculation.xlCalculationManual;

            using (new UiGuard(_app, hideStatusBar: true))
            {
                int total = items.Count;
                _progress.Maximum = Math.Max(1, total);
                int idx = 0;
                foreach (var item in items)
                {
                    idx++;
                    if (_cancelRequested)
                    {
                        break;
                    }
                    ReportProgress(idx, total);
                    try
                    {
                        changed.AddRange(NameScrubberUtil.ApplyNameToDependents(item.Name, () => _cancelRequested, null));
                    }
                    catch
                    {
                        // ignore
                    }
                }
            }

            _app.DisplayAlerts = prevAlerts;
            _app.Calculation = prevCalc;

            SetBusy(false, null);
            UpdateActionButtons();
            if (changed.Count > 0)
            {
                PromptGoTo(changed, "Some formulas were updated.");
            }
        }

        private void UnapplySelected()
        {
            var items = GetActionItems();
            if (items.Count == 0)
            {
                return;
            }

            _cancelRequested = false;
            SetBusy(true, "Unapplying names...");

            var changed = new List<Excel.Range>();
            var prevCalc = _app.Calculation;
            var prevAlerts = _app.DisplayAlerts;
            _app.DisplayAlerts = false;
            _app.Calculation = Excel.XlCalculation.xlCalculationManual;

            using (new UiGuard(_app, hideStatusBar: true))
            {
                int total = items.Count;
                _progress.Maximum = Math.Max(1, total);
                int idx = 0;
                foreach (var item in items)
                {
                    idx++;
                    if (_cancelRequested)
                    {
                        break;
                    }
                    ReportProgress(idx, total);
                    try
                    {
                        changed.AddRange(NameScrubberUtil.UnapplyNameFromDependents(item.Name, () => _cancelRequested, null));
                    }
                    catch
                    {
                        // ignore
                    }
                }
            }

            _app.DisplayAlerts = prevAlerts;
            _app.Calculation = prevCalc;

            SetBusy(false, null);
            UpdateActionButtons();
            if (changed.Count > 0)
            {
                PromptGoTo(changed, "Some formulas were updated.");
            }
        }

        private void DeleteSelected()
        {
            var items = GetActionItems();
            if (items.Count == 0)
            {
                return;
            }

            var response = MessageBox.Show(this,
                "Delete selected names?\n\nYes = replace formulas with cell references before deleting.\nNo = delete names only.",
                "Name Scrubber",
                MessageBoxButtons.YesNoCancel,
                MessageBoxIcon.Question);

            if (response == DialogResult.Cancel)
            {
                return;
            }

            bool unapplyFirst = response == DialogResult.Yes;
            _cancelRequested = false;
            SetBusy(true, "Deleting names...");

            var changed = new List<Excel.Range>();
            var prevCalc = _app.Calculation;
            var prevAlerts = _app.DisplayAlerts;
            _app.DisplayAlerts = false;
            _app.Calculation = Excel.XlCalculation.xlCalculationManual;

            using (new UiGuard(_app, hideStatusBar: true))
            {
                int total = items.Count;
                _progress.Maximum = Math.Max(1, total);
                int idx = 0;

                foreach (var item in items)
                {
                    idx++;
                    if (_cancelRequested)
                    {
                        break;
                    }
                    ReportProgress(idx, total);

                    if (unapplyFirst)
                    {
                        try
                        {
                            changed.AddRange(NameScrubberUtil.UnapplyNameFromDependents(item.Name, () => _cancelRequested, null));
                        }
                        catch
                        {
                            // ignore
                        }
                    }

                    try
                    {
                        item.Name.Visible = true;
                        item.Name.Delete();
                    }
                    catch
                    {
                        // ignore
                    }
                }
            }

            _app.DisplayAlerts = prevAlerts;
            _app.Calculation = prevCalc;

            SetBusy(false, null);
            ReloadNames();

            if (unapplyFirst && changed.Count > 0)
            {
                PromptGoTo(changed, "Formulas were updated before deleting names.");
            }
        }

        private void EditSelectedName()
        {
            var items = GetActionItems();
            if (items.Count != 1)
            {
                return;
            }

            var item = items[0];
            var current = item.Name?.Name ?? string.Empty;
            var input = Microsoft.VisualBasic.Interaction.InputBox("Enter new name:", "Edit Name", current);
            if (string.IsNullOrWhiteSpace(input) || string.Equals(input, current, StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            if (!IsValidName(input))
            {
                MessageBox.Show(this, "The name is not valid.", "Name Scrubber", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                item.Name.Name = input;
                item.Label = NameScrubberUtil.StripSheetPrefix(input);
                item.RefersTo = NameScrubberUtil.SafeRefersTo(item.Name);
                ApplyFilterAndSearch();
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message, "Name Scrubber", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void EditSelectedRefersTo()
        {
            var items = GetActionItems();
            if (items.Count != 1)
            {
                return;
            }

            var item = items[0];
            var current = NameScrubberUtil.SafeRefersTo(item.Name);
            var input = Microsoft.VisualBasic.Interaction.InputBox("Edit RefersTo:", "Edit RefersTo", current);
            input = input?.Trim() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(input))
            {
                return;
            }

            if (!input.StartsWith("=", StringComparison.Ordinal))
            {
                input = "=" + input;
            }

            if (string.Equals(input, current, StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            try
            {
                item.Name.RefersTo = input;
                item.RefersTo = NameScrubberUtil.SafeRefersTo(item.Name);
                item.Text = BuildValueText(item.Name);
                ApplyFilterAndSearch();
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message, "Name Scrubber", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void CleanNames()
        {
            var dialog = new CleanNamesDialog(_app.ActiveWorkbook);
            if (dialog.ShowDialog(this) != DialogResult.OK)
            {
                return;
            }

            var options = dialog.Options;
            RunCleanNames(options);
        }

        private void RunCleanNames(CleanNamesOptions options)
        {
            var wb = _app.ActiveWorkbook;
            if (wb == null)
            {
                return;
            }

            _cancelRequested = false;
            SetBusy(true, "Cleaning names...");

            int removed = 0;
            int count = 0;
            try { count = wb.Names.Count; } catch { count = 0; }
            _progress.Maximum = Math.Max(1, count);
            _progress.Value = 0;

            var prevCalc = _app.Calculation;
            var prevAlerts = _app.DisplayAlerts;
            _app.DisplayAlerts = false;
            _app.Calculation = Excel.XlCalculation.xlCalculationManual;

            using (new UiGuard(_app, hideStatusBar: true))
            {
                for (int i = count; i >= 1; i--)
                {
                    if (_cancelRequested)
                    {
                        break;
                    }

                    ReportProgress(count - i + 1, count);

                    Excel.Name name = null;
                    try
                    {
                        name = wb.Names.Item(i);
                    }
                    catch
                    {
                        continue;
                    }

                    if (name == null)
                    {
                        continue;
                    }

                    if (!options.IncludeHidden && !name.Visible)
                    {
                        continue;
                    }

                    if (NameScrubberUtil.IsNative(name))
                    {
                        continue;
                    }

                    var refersTo = NameScrubberUtil.SafeRefersTo(name);
                    bool shouldDelete = NameScrubberUtil.IsErroneous(name);
                    if (!shouldDelete && options.DeepClean)
                    {
                        shouldDelete = NameScrubberUtil.ExternalLinkMissing(refersTo, wb.Path);
                    }

                    if (!shouldDelete)
                    {
                        continue;
                    }

                    try
                    {
                        name.Visible = true;
                        name.Delete();
                        removed++;
                    }
                    catch
                    {
                        // ignore
                    }
                }
            }

            _app.DisplayAlerts = prevAlerts;
            _app.Calculation = prevCalc;

            SetBusy(false, null);
            MessageBox.Show(this,
                string.Format("Removed {0:N0} of {1:N0} names.", removed, count),
                "Name Scrubber",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);

            ReloadNames();
        }

        private void PromptGoTo(List<Excel.Range> ranges, string message)
        {
            if (ranges == null || ranges.Count == 0)
            {
                return;
            }

            var response = MessageBox.Show(this,
                message + "\n\nOpen list of updated cells?",
                "Name Scrubber",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (response != DialogResult.Yes)
            {
                return;
            }

            using (var form = new GoToForm(_app, ranges))
            {
                form.ShowDialog(this);
            }
        }

        private static bool IsValidName(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                return false;
            }

            if (NameScrubberUtil.IsNative(name))
            {
                return false;
            }

            if (Regex.IsMatch(name, "^[A-Za-z]{1,3}[0-9]+$", RegexOptions.IgnoreCase))
            {
                return false;
            }

            if (Regex.IsMatch(name, "^R[0-9]+C[0-9]+$", RegexOptions.IgnoreCase))
            {
                return false;
            }

            return Regex.IsMatch(name, "^[A-Za-z_\\\\][A-Za-z0-9_\\\\.]*$", RegexOptions.IgnoreCase);
        }

        private static void StyleActionButton(Button button)
        {
            if (button == null)
            {
                return;
            }

            button.FlatStyle = FlatStyle.Flat;
            button.BackColor = Color.White;
            button.ForeColor = Color.FromArgb(60, 60, 60);
            button.AutoSize = true;
            button.FlatAppearance.BorderColor = BorderColor;
            button.FlatAppearance.BorderSize = 1;
            button.Margin = new Padding(6, 2, 0, 2);
        }

        private void SetBusy(bool busy, string status)
        {
            _busy = busy;
            _progress.Visible = busy;
            _cancelButton.Visible = busy;
            _cancelButton.Enabled = busy;

            _namesList.Enabled = !busy;
            _dependentsList.Enabled = !busy;
            _filterCombo.Enabled = !busy;
            _searchModeCombo.Enabled = !busy;
            _searchBox.Enabled = !busy;
            _checkAll.Enabled = !busy;
            _showDependents.Enabled = !busy;

            if (busy)
            {
                DisableActionButtons();
                _progress.Value = 0;
                _cancelRequested = false;
            }
            else
            {
                UpdateActionButtons();
            }

            if (!string.IsNullOrEmpty(status))
            {
                _countLabel.Text = status;
            }
        }

        private void ReportProgress(int current, int total)
        {
            if (!_progress.Visible)
            {
                return;
            }

            if (total <= 0)
            {
                return;
            }

            int value = Math.Min(_progress.Maximum, Math.Max(0, current));
            _progress.Value = value;
            Application.DoEvents();
        }

        protected override void OnDeactivate(EventArgs e)
        {
            base.OnDeactivate(e);

            if (_busy || _suppressDeactivate)
            {
                return;
            }

            Close();
        }
    }
}
