using System;
using System.CodeDom.Compiler;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using A;
using ExcelAddIn1.Audit.TraceDialogs;
using ExcelAddIn1.Charts;
using ExcelAddIn1.Shapes;
using ExcelAddIn1.Sheets;
using ExcelAddIn1.SuperFind2.Queries;
using ExcelAddIn1.SuperFind2.Results;
using ExcelAddIn1.View;
using MacabacusMacros;
using MacabacusMacros.ExcelHelpers;
using MacabacusMacros.Explorer;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.SuperFind2.UI;

[DesignerGenerated]
public sealed class wpfResults : System.Windows.Controls.UserControl, INotifyPropertyChanged, IComponentConnector, IStyleConnector
{
	[Serializable]
	[CompilerGenerated]
	internal sealed class _Closure_0024__
	{
		public static readonly _Closure_0024__ A;

		public static ThreadStart A;

		public static Func<ResultItem, int> A;

		public static Func<ResultItem, string> A;

		public static Func<ResultItem, object> A;

		public static Func<ResultItem, object> B;

		static _Closure_0024__()
		{
			_Closure_0024__.A = new _Closure_0024__();
		}

		[SpecialName]
		internal void A()
		{
		}

		[SpecialName]
		internal int A(ResultItem A)
		{
			return A.UiIndex;
		}

		[SpecialName]
		internal string A(ResultItem A)
		{
			return ((object)A).GetType().ToString();
		}

		[SpecialName]
		internal object A(ResultItem A)
		{
			Range range = A.Range;
			if (range == null)
			{
				return null;
			}
			return NewLateBinding.LateGet(range.Cells[1, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(41347), new object[0], null, null, null);
		}

		[SpecialName]
		internal object B(ResultItem A)
		{
			Range range = A.Range;
			if (range == null)
			{
				while (true)
				{
					switch (2)
					{
					case 0:
						break;
					default:
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						return null;
					}
				}
			}
			return NewLateBinding.LateGet(range.Cells[1, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(41354), new object[0], null, null, null);
		}
	}

	[CompilerGenerated]
	internal sealed class NF
	{
		public bool A;

		public wpfResults A;

		public NF(NF A)
		{
			if (A == null)
			{
				return;
			}
			while (true)
			{
				switch (1)
				{
				case 0:
					continue;
				}
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				this.A = A.A;
				return;
			}
		}

		[SpecialName]
		internal void A()
		{
			IEnumerator<BaseItem> enumerator = default(IEnumerator<BaseItem>);
			try
			{
				enumerator = this.A.SearchResults.GetEnumerator();
				IEnumerator<SheetItem> enumerator2 = default(IEnumerator<SheetItem>);
				while (enumerator.MoveNext())
				{
					BaseItem current = enumerator.Current;
					((BaseItem)current).IsExpanded = this.A;
					if (!(current is WorkbookItem))
					{
						continue;
					}
					while (true)
					{
						switch (5)
						{
						case 0:
							continue;
						}
						break;
					}
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					try
					{
						enumerator2 = ((WorkbookItem)current).Sheets.GetEnumerator();
						while (enumerator2.MoveNext())
						{
							((BaseItem)enumerator2.Current).IsExpanded = true;
						}
					}
					finally
					{
						if (enumerator2 != null)
						{
							while (true)
							{
								switch (4)
								{
								case 0:
									continue;
								}
								enumerator2.Dispose();
								break;
							}
						}
					}
				}
			}
			finally
			{
				if (enumerator != null)
				{
					while (true)
					{
						switch (4)
						{
						case 0:
							continue;
						}
						enumerator.Dispose();
						break;
					}
				}
			}
			this.A.Dispatcher.BeginInvoke((global::A.A)([SpecialName] () =>
			{
				this.A.K();
			}));
		}
	}

	[CompilerGenerated]
	internal sealed class OF
	{
		public BaseItem A;

		public BitmapSource A;

		public wpfResults A;

		[SpecialName]
		internal void A()
		{
			Preview.PrepareWindow(this.A.PreviewWindow, (BaseItem)(object)this.A, this.A, this.A.PointToScreen(new System.Windows.Point(0.0, 0.0)));
		}
	}

	[CompilerGenerated]
	internal sealed class PF
	{
		public DoWorkEventArgs A;

		public wpfResults A;

		[SpecialName]
		internal void A()
		{
			this.A.A(ref this.A);
		}
	}

	[CompilerGenerated]
	internal sealed class QF
	{
		public RunWorkerCompletedEventArgs A;

		public wpfResults A;

		[SpecialName]
		internal void A()
		{
			this.A.A(this.A);
		}
	}

	[CompilerGenerated]
	internal sealed class RF
	{
		public string A;

		public wpfResults A;

		[SpecialName]
		internal void A()
		{
			this.A.txtLoading.Text = this.A;
		}
	}

	[CompilerGenerated]
	private PropertyChangedEventHandler m_A;

	private static FieldInfo m_A;

	private RoutedPropertyChangedEventHandler<object> m_A;

	[CompilerGenerated]
	private ctpSuperFind2 m_A;

	private DispatcherFrame m_A;

	private DispatcherTimer m_A;

	private DispatcherTimer m_B;

	private ICollectionView m_A;

	[CompilerGenerated]
	private ObservableCollection<BaseItem> m_A;

	[CompilerGenerated]
	private List<BaseQuery> m_A;

	[CompilerGenerated]
	private SearchScope m_A;

	[CompilerGenerated]
	private bool m_A;

	[CompilerGenerated]
	private wpfPane m_A;

	[CompilerGenerated]
	private wpfPreview m_A;

	[CompilerGenerated]
	private object m_A;

	[CompilerGenerated]
	private SheetItem m_A;

	[CompilerGenerated]
	private MsoCTPDockPosition m_A;

	[CompilerGenerated]
	private BackgroundWorker m_A;

	[CompilerGenerated]
	private List<Worksheet> m_A;

	private double m_A;

	private bool m_B;

	private ScrollViewer m_A;

	private bool m_C;

	[CompilerGenerated]
	[AccessedThroughProperty("btnBack")]
	private System.Windows.Controls.Button m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("btnRefresh")]
	private System.Windows.Controls.Button m_B;

	[CompilerGenerated]
	[AccessedThroughProperty("chkExpandCollapse")]
	private System.Windows.Controls.CheckBox m_A;

	[AccessedThroughProperty("chkHighlight")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox m_B;

	[CompilerGenerated]
	[AccessedThroughProperty("chkPreview")]
	private System.Windows.Controls.CheckBox m_C;

	[CompilerGenerated]
	[AccessedThroughProperty("ccOutdated")]
	private ContentControl m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("ccNoResults")]
	private ContentControl m_B;

	[CompilerGenerated]
	[AccessedThroughProperty("ccCanceled")]
	private ContentControl m_C;

	[AccessedThroughProperty("trvResults")]
	[CompilerGenerated]
	private System.Windows.Controls.TreeView m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("grdLoading")]
	private Grid m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("txtLoading")]
	private TextBlock m_A;

	[AccessedThroughProperty("btnCancel")]
	[CompilerGenerated]
	private System.Windows.Controls.Button m_C;

	[CompilerGenerated]
	[AccessedThroughProperty("radCell")]
	private System.Windows.Controls.RadioButton m_A;

	[AccessedThroughProperty("radRow")]
	[CompilerGenerated]
	private System.Windows.Controls.RadioButton m_B;

	[CompilerGenerated]
	[AccessedThroughProperty("radColumn")]
	private System.Windows.Controls.RadioButton m_C;

	[CompilerGenerated]
	[AccessedThroughProperty("btnSelect")]
	private System.Windows.Controls.Button m_D;

	private bool m_D;

	public ctpSuperFind2 thisCTP
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[CompilerGenerated]
		set
		{
			this.m_A = value;
		}
	}

	public ICollectionView SourceCollection
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(VH.A(36261));
		}
	}

	internal ObservableCollection<BaseItem> SearchResults
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[CompilerGenerated]
		set
		{
			this.m_A = value;
		}
	}

	private List<BaseQuery> SearchQueries
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[CompilerGenerated]
		set
		{
			this.m_A = value;
		}
	}

	private SearchScope SearchScope
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[CompilerGenerated]
		set
		{
			this.m_A = value;
		}
	}

	private bool LookInPrintAreasOnly
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[CompilerGenerated]
		set
		{
			this.m_A = value;
		}
	}

	private wpfPane ParentView
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[CompilerGenerated]
		set
		{
			this.m_A = value;
		}
	}

	private wpfPreview PreviewWindow
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[CompilerGenerated]
		set
		{
			this.m_A = value;
		}
	}

	private object ActiveMenu
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[CompilerGenerated]
		set
		{
			this.m_A = RuntimeHelpers.GetObjectValue(value);
		}
	}

	private SheetItem CopiedSheet
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[CompilerGenerated]
		set
		{
			this.m_A = value;
		}
	}

	internal MsoCTPDockPosition DockPosition
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[CompilerGenerated]
		set
		{
			this.m_A = value;
		}
	}

	private BackgroundWorker bgw
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[CompilerGenerated]
		set
		{
			this.m_A = value;
		}
	}

	private List<Worksheet> HighlightedSheets
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[CompilerGenerated]
		set
		{
			this.m_A = value;
		}
	}

	internal virtual System.Windows.Controls.Button btnBack
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			RoutedEventHandler value2 = BackToSearchView;
			System.Windows.Controls.Button button = this.m_A;
			if (button != null)
			{
				while (true)
				{
					switch (7)
					{
					case 0:
						continue;
					}
					break;
				}
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				button.Click -= value2;
			}
			this.m_A = value;
			button = this.m_A;
			if (button == null)
			{
				return;
			}
			while (true)
			{
				switch (4)
				{
				case 0:
					continue;
				}
				button.Click += value2;
				return;
			}
		}
	}

	internal virtual System.Windows.Controls.Button btnRefresh
	{
		[CompilerGenerated]
		get
		{
			return this.m_B;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			RoutedEventHandler value2 = RefreshResults;
			System.Windows.Controls.Button button = this.m_B;
			if (button != null)
			{
				button.Click -= value2;
			}
			this.m_B = value;
			button = this.m_B;
			if (button != null)
			{
				button.Click += value2;
			}
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkExpandCollapse
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_A = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkHighlight
	{
		[CompilerGenerated]
		get
		{
			return this.m_B;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_B = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkPreview
	{
		[CompilerGenerated]
		get
		{
			return this.m_C;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_C = value;
		}
	}

	internal virtual ContentControl ccOutdated
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_A = value;
		}
	}

	internal virtual ContentControl ccNoResults
	{
		[CompilerGenerated]
		get
		{
			return this.m_B;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_B = value;
		}
	}

	internal virtual ContentControl ccCanceled
	{
		[CompilerGenerated]
		get
		{
			return this.m_C;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_C = value;
		}
	}

	internal virtual System.Windows.Controls.TreeView trvResults
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			System.Windows.Input.KeyEventHandler value2 = TreeViewPreviewKeyDown;
			MouseWheelEventHandler value3 = trvExplorer_PreviewMouseWheel;
			System.Windows.Controls.TreeView treeView = this.m_A;
			if (treeView != null)
			{
				while (true)
				{
					switch (1)
					{
					case 0:
						continue;
					}
					break;
				}
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				treeView.PreviewKeyDown -= value2;
				treeView.PreviewMouseWheel -= value3;
			}
			this.m_A = value;
			treeView = this.m_A;
			if (treeView == null)
			{
				return;
			}
			while (true)
			{
				switch (2)
				{
				case 0:
					continue;
				}
				treeView.PreviewKeyDown += value2;
				treeView.PreviewMouseWheel += value3;
				return;
			}
		}
	}

	internal virtual Grid grdLoading
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_A = value;
		}
	}

	internal virtual TextBlock txtLoading
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_A = value;
		}
	}

	internal virtual System.Windows.Controls.Button btnCancel
	{
		[CompilerGenerated]
		get
		{
			return this.m_C;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			RoutedEventHandler value2 = CancelButtonClicked;
			System.Windows.Controls.Button button = this.m_C;
			if (button != null)
			{
				while (true)
				{
					switch (6)
					{
					case 0:
						continue;
					}
					break;
				}
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				button.Click -= value2;
			}
			this.m_C = value;
			button = this.m_C;
			if (button == null)
			{
				return;
			}
			while (true)
			{
				switch (6)
				{
				case 0:
					continue;
				}
				button.Click += value2;
				return;
			}
		}
	}

	internal virtual System.Windows.Controls.RadioButton radCell
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_A = value;
		}
	}

	internal virtual System.Windows.Controls.RadioButton radRow
	{
		[CompilerGenerated]
		get
		{
			return this.m_B;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_B = value;
		}
	}

	internal virtual System.Windows.Controls.RadioButton radColumn
	{
		[CompilerGenerated]
		get
		{
			return this.m_C;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_C = value;
		}
	}

	internal virtual System.Windows.Controls.Button btnSelect
	{
		[CompilerGenerated]
		get
		{
			return this.m_D;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			RoutedEventHandler value2 = btnSelect_Click;
			System.Windows.Controls.Button button = this.m_D;
			if (button != null)
			{
				while (true)
				{
					switch (3)
					{
					case 0:
						continue;
					}
					break;
				}
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				button.Click -= value2;
			}
			this.m_D = value;
			button = this.m_D;
			if (button == null)
			{
				return;
			}
			while (true)
			{
				switch (3)
				{
				case 0:
					continue;
				}
				button.Click += value2;
				return;
			}
		}
	}

	public event PropertyChangedEventHandler PropertyChanged
	{
		[CompilerGenerated]
		add
		{
			PropertyChangedEventHandler propertyChangedEventHandler = this.m_A;
			PropertyChangedEventHandler propertyChangedEventHandler2;
			do
			{
				propertyChangedEventHandler2 = propertyChangedEventHandler;
				PropertyChangedEventHandler value2 = (PropertyChangedEventHandler)Delegate.Combine(propertyChangedEventHandler2, value);
				propertyChangedEventHandler = Interlocked.CompareExchange(ref this.m_A, value2, propertyChangedEventHandler2);
			}
			while ((object)propertyChangedEventHandler != propertyChangedEventHandler2);
			while (true)
			{
				switch (3)
				{
				case 0:
					continue;
				}
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				return;
			}
		}
		[CompilerGenerated]
		remove
		{
			PropertyChangedEventHandler propertyChangedEventHandler = this.m_A;
			PropertyChangedEventHandler propertyChangedEventHandler2;
			do
			{
				propertyChangedEventHandler2 = propertyChangedEventHandler;
				PropertyChangedEventHandler value2 = (PropertyChangedEventHandler)Delegate.Remove(propertyChangedEventHandler2, value);
				propertyChangedEventHandler = Interlocked.CompareExchange(ref this.m_A, value2, propertyChangedEventHandler2);
			}
			while ((object)propertyChangedEventHandler != propertyChangedEventHandler2);
			while (true)
			{
				switch (6)
				{
				case 0:
					continue;
				}
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				return;
			}
		}
	}

	public wpfResults(wpfPane parent, List<BaseQuery> queries, SearchScope scope, bool blnPrintAreasOnly, bool blnExploreMode)
	{
		//IL_00cb: Unknown result type (might be due to invalid IL or missing references)
		//IL_00d5: Expected O, but got Unknown
		base.Loaded += ViewLoaded;
		base.Unloaded += ViewUnloaded;
		base.KeyDown += ViewKeyDown;
		this.m_A = [SpecialName] (object a0, RoutedPropertyChangedEventArgs<object> a1) =>
		{
			TreeViewSelectionChanged((System.Windows.Controls.TreeView)a0, a1);
		};
		this.m_A = null;
		this.m_B = null;
		this.m_A = null;
		PreviewWindow = null;
		CopiedSheet = null;
		this.m_A = 0.0;
		this.m_B = false;
		this.m_A = null;
		InitializeComponent();
		ParentView = parent;
		SearchQueries = queries;
		SearchScope = scope;
		LookInPrintAreasOnly = blnPrintAreasOnly;
		if (Constants.ColorPalette == null)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					continue;
				}
				break;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			Constants.ColorPalette = new Palette();
		}
		MySettings settings = global::A.K.Settings;
		System.Windows.Controls.CheckBox checkBox = chkHighlight;
		int value;
		if (settings.SuperFindHighlightResults)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					continue;
				}
				break;
			}
			value = ((!blnExploreMode) ? 1 : 0);
		}
		else
		{
			value = 0;
		}
		checkBox.IsChecked = (byte)value != 0;
		chkPreview.IsChecked = settings.ExplorerPreviews;
		switch (settings.AdvancedFindSelectMode)
		{
		case 0:
			radCell.IsChecked = true;
			break;
		case 1:
			radRow.IsChecked = true;
			break;
		case 2:
			radColumn.IsChecked = true;
			break;
		}
		settings = null;
		if (blnExploreMode)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					continue;
				}
				break;
			}
			chkHighlight.Visibility = System.Windows.Visibility.Collapsed;
		}
		S();
		bgw = new BackgroundWorker();
		BackgroundWorker backgroundWorker = bgw;
		backgroundWorker.WorkerSupportsCancellation = true;
		backgroundWorker.WorkerReportsProgress = false;
		backgroundWorker.DoWork += StartBackgroundSearch;
		backgroundWorker.RunWorkerCompleted += BackgroundSearchCompleted;
		_ = null;
	}

	private void A(string A)
	{
		PropertyChangedEventHandler propertyChangedEventHandler = this.m_A;
		if (propertyChangedEventHandler == null)
		{
			return;
		}
		while (true)
		{
			switch (1)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			propertyChangedEventHandler(this, new PropertyChangedEventArgs(A));
			return;
		}
	}

	private static void A()
	{
		wpfResults.m_A = typeof(SystemParameters).GetField(VH.A(42651), BindingFlags.Static | BindingFlags.NonPublic);
		B();
		SystemParameters.StaticPropertyChanged += SystemParameters_StaticPropertyChanged;
	}

	private static void SystemParameters_StaticPropertyChanged(object sender, PropertyChangedEventArgs e)
	{
		B();
	}

	private static void B()
	{
		if (!SystemParameters.MenuDropAlignment || (object)wpfResults.m_A == null)
		{
			return;
		}
		while (true)
		{
			switch (7)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			wpfResults.m_A.SetValue(null, false);
			return;
		}
	}

	private void ViewLoaded(object sender, RoutedEventArgs e)
	{
		Props.SingleCellIcon = (ControlTemplate)FindResource(VH.A(124857));
		Props.MultiCellIcon = (ControlTemplate)FindResource(VH.A(124878));
		DockPosition = Pane.PanesCollection[MH.A.Application.ActiveWindow.Hwnd].DockPosition;
		C();
		E();
	}

	private void ViewUnloaded(object sender, RoutedEventArgs e)
	{
		I();
		D();
		RemoveTreeViewSelectionChangedHandler();
		Q();
		SourceCollection = null;
		SearchResults = null;
		SearchQueries = null;
		ParentView = null;
		PreviewWindow = null;
		this.m_A = null;
		this.m_B = null;
		this.m_A = null;
		ActiveMenu = null;
		this.m_A = null;
		Props.SingleCellIcon = null;
		Props.MultiCellIcon = null;
		Props.SearchForm = null;
	}

	private void ViewKeyDown(object sender, System.Windows.Input.KeyEventArgs e)
	{
		if (e.Key == Key.Escape)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					Pane.Toggle(blnPressed: false);
					return;
				}
			}
		}
		if (e.Key != Key.F5)
		{
			if (e.Key != Key.R)
			{
				return;
			}
			while (true)
			{
				switch (2)
				{
				case 0:
					continue;
				}
				break;
			}
			if (System.Windows.Input.Keyboard.Modifiers != ModifierKeys.Control)
			{
				return;
			}
			while (true)
			{
				switch (1)
				{
				case 0:
					continue;
				}
				break;
			}
		}
		if (bgw != null)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					continue;
				}
				break;
			}
			if (bgw.IsBusy)
			{
				return;
			}
		}
		E();
	}

	private void C()
	{
		chkHighlight.Checked += HighlightToggle;
		chkHighlight.Unchecked += HighlightToggle;
		chkPreview.Checked += PreviewToggle;
		chkPreview.Unchecked += PreviewToggle;
		radCell.Checked += CellModeChecked;
		radRow.Checked += RowModeChecked;
		radColumn.Checked += ColumnModeChecked;
		chkExpandCollapse.Checked += ExpandCollapseChanged;
		chkExpandCollapse.Unchecked += ExpandCollapseChanged;
		trvResults.LostKeyboardFocus += TreeViewLostKeyboardFocus;
		trvResults.GotKeyboardFocus += TreeViewGotKeyboardFocus;
	}

	private void D()
	{
		chkHighlight.Checked -= HighlightToggle;
		chkHighlight.Unchecked -= HighlightToggle;
		chkPreview.Checked -= PreviewToggle;
		chkPreview.Unchecked -= PreviewToggle;
		radCell.Checked -= CellModeChecked;
		radRow.Checked -= RowModeChecked;
		radColumn.Checked -= ColumnModeChecked;
		chkExpandCollapse.Checked -= ExpandCollapseChanged;
		chkExpandCollapse.Unchecked -= ExpandCollapseChanged;
		trvResults.LostKeyboardFocus -= TreeViewLostKeyboardFocus;
		trvResults.GotKeyboardFocus -= TreeViewGotKeyboardFocus;
	}

	private void StartBackgroundSearch(object sender, DoWorkEventArgs e)
	{
		base.Dispatcher.Invoke([SpecialName] () =>
		{
			A(ref e);
		});
	}

	private void BackgroundSearchCompleted(object sender, RunWorkerCompletedEventArgs e)
	{
		base.Dispatcher.Invoke([SpecialName] () =>
		{
			A(e);
		});
	}

	private void E()
	{
		S();
		txtLoading.Text = VH.A(124897);
		grdLoading.Visibility = System.Windows.Visibility.Visible;
		btnRefresh.IsEnabled = false;
		chkExpandCollapse.IsEnabled = false;
		System.Windows.Forms.Application.DoEvents();
		I();
		Q();
		RemoveTreeViewSelectionChangedHandler();
		bgw.RunWorkerAsync();
	}

	private void B(string A)
	{
		try
		{
			base.Dispatcher.Invoke([SpecialName] () =>
			{
				txtLoading.Text = A;
			});
			Dispatcher dispatcher = base.Dispatcher;
			ThreadStart method;
			if (_Closure_0024__.A == null)
			{
				method = (_Closure_0024__.A = [SpecialName] () =>
				{
				});
			}
			else
			{
				while (true)
				{
					switch (1)
					{
					case 0:
						continue;
					}
					break;
				}
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				method = _Closure_0024__.A;
			}
			dispatcher.Invoke(DispatcherPriority.Background, method);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	private void A(RunWorkerCompletedEventArgs A)
	{
		grdLoading.Visibility = System.Windows.Visibility.Collapsed;
		btnRefresh.IsEnabled = true;
		chkExpandCollapse.IsEnabled = true;
		if (!A.Cancelled)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					continue;
				}
				break;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			if (SearchResults.Count > 0)
			{
				while (true)
				{
					switch (4)
					{
					case 0:
						continue;
					}
					break;
				}
				this.A(MH.A.Application);
				trvResults.Focus();
			}
			else
			{
				ccNoResults.Visibility = System.Windows.Visibility.Visible;
			}
		}
		else
		{
			try
			{
				SearchResults.Clear();
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			ccCanceled.Visibility = System.Windows.Visibility.Visible;
		}
		K();
	}

	private void A(ref DoWorkEventArgs A)
	{
		Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
		int count = application.ActiveWindow.SelectedSheets.Count;
		Range rng = default(Range);
		try
		{
			rng = (Range)application.Selection;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		application.ScreenUpdating = false;
		if (SearchResults == null)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					continue;
				}
				break;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			SearchResults = new ObservableCollection<BaseItem>();
		}
		else
		{
			SearchResults.Clear();
		}
		try
		{
			switch (SearchScope)
			{
			case SearchScope.Selection:
			{
				WorksheetItem worksheetItem = new WorksheetItem(null, (Worksheet)application.ActiveSheet, application.ActiveWorkbook, 0);
				if (!(application.Selection is Range))
				{
					using List<Chart>.Enumerator enumerator5 = Helpers.SelectedCharts(null).ChartList.GetEnumerator();
					while (enumerator5.MoveNext())
					{
						Chart current = enumerator5.Current;
						using List<BaseQuery>.Enumerator enumerator6 = SearchQueries.GetEnumerator();
						while (enumerator6.MoveNext())
						{
							BaseQuery current2 = enumerator6.Current;
							if (bgw.CancellationPending)
							{
								while (true)
								{
									switch (1)
									{
									case 0:
										break;
									default:
										return;
									}
								}
							}
							if (current2 is TF)
							{
								((TF)current2).Query(worksheetItem, current);
							}
						}
						while (true)
						{
							switch (7)
							{
							case 0:
								break;
							default:
								goto end_IL_025e;
							}
							continue;
							end_IL_025e:
							break;
						}
					}
					while (true)
					{
						switch (7)
						{
						case 0:
							break;
						default:
							goto end_IL_0284;
						}
						continue;
						end_IL_0284:
						break;
					}
				}
				else
				{
					while (true)
					{
						switch (7)
						{
						case 0:
							continue;
						}
						break;
					}
					Range range = JH.A((Range)null);
					if (Operators.ConditionalCompareObjectGreater(range.Cells.CountLarge, 1, TextCompare: false))
					{
						using List<BaseQuery>.Enumerator enumerator7 = SearchQueries.GetEnumerator();
						while (enumerator7.MoveNext())
						{
							BaseQuery current3 = enumerator7.Current;
							if (bgw.CancellationPending)
							{
								while (true)
								{
									switch (6)
									{
									case 0:
										break;
									default:
										return;
									}
								}
							}
							if (current3 is VF)
							{
								Props.SearchForm = current3;
								((VF)current3).Query(worksheetItem, range);
							}
							else
							{
								if (!(current3 is UF))
								{
									continue;
								}
								while (true)
								{
									switch (2)
									{
									case 0:
										continue;
									}
									break;
								}
								Props.SearchForm = current3;
								((UF)current3).Query(worksheetItem, range);
							}
						}
						while (true)
						{
							switch (2)
							{
							case 0:
								break;
							default:
								goto end_IL_01b0;
							}
							continue;
							end_IL_01b0:
							break;
						}
					}
					else
					{
						F();
					}
				}
				this.A(worksheetItem, (WorkbookItem)null);
				worksheetItem = null;
				break;
			}
			case SearchScope.SelectedSheets:
			{
				IEnumerator enumerator4 = default(IEnumerator);
				try
				{
					enumerator4 = application.ActiveWindow.SelectedSheets.GetEnumerator();
					while (enumerator4.MoveNext())
					{
						object objectValue3 = RuntimeHelpers.GetObjectValue(enumerator4.Current);
						if (bgw.CancellationPending)
						{
							while (true)
							{
								switch (2)
								{
								case 0:
									break;
								default:
									return;
								}
							}
						}
						B(Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject(VH.A(124922), NewLateBinding.LateGet(objectValue3, null, VH.A(19019), new object[0], null, null, null)), VH.A(39851))));
						if (!(objectValue3 is Worksheet))
						{
							continue;
						}
						Worksheet worksheet = (Worksheet)objectValue3;
						if (count == 1)
						{
							while (true)
							{
								switch (1)
								{
								case 0:
									continue;
								}
								break;
							}
							ExcelAddIn1.Sheets.Protection.Unprotect(worksheet);
						}
						this.A(worksheet);
					}
					while (true)
					{
						switch (2)
						{
						case 0:
							break;
						default:
							return;
						}
					}
				}
				finally
				{
					if (enumerator4 is IDisposable)
					{
						while (true)
						{
							switch (4)
							{
							case 0:
								continue;
							}
							(enumerator4 as IDisposable).Dispose();
							break;
						}
					}
				}
			}
			case SearchScope.ThisWorkbook:
			{
				IEnumerator enumerator3 = default(IEnumerator);
				try
				{
					enumerator3 = application.ActiveWorkbook.Sheets.GetEnumerator();
					while (enumerator3.MoveNext())
					{
						object objectValue2 = RuntimeHelpers.GetObjectValue(enumerator3.Current);
						if (bgw.CancellationPending)
						{
							break;
						}
						B(Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject(VH.A(124922), NewLateBinding.LateGet(objectValue2, null, VH.A(19019), new object[0], null, null, null)), VH.A(39851))));
						if (!(objectValue2 is Worksheet))
						{
							continue;
						}
						Worksheet worksheet = (Worksheet)objectValue2;
						if (count == 1)
						{
							while (true)
							{
								switch (5)
								{
								case 0:
									continue;
								}
								break;
							}
							ExcelAddIn1.Sheets.Protection.Unprotect(worksheet);
						}
						this.A(worksheet);
					}
					break;
				}
				finally
				{
					if (enumerator3 is IDisposable)
					{
						while (true)
						{
							switch (4)
							{
							case 0:
								continue;
							}
							(enumerator3 as IDisposable).Dispose();
							break;
						}
					}
				}
			}
			case SearchScope.AllWorkbooks:
			{
				IEnumerator enumerator = default(IEnumerator);
				try
				{
					enumerator = application.Workbooks.GetEnumerator();
					IEnumerator enumerator2 = default(IEnumerator);
					while (enumerator.MoveNext())
					{
						Microsoft.Office.Interop.Excel.Workbook workbook = (Microsoft.Office.Interop.Excel.Workbook)enumerator.Current;
						if (bgw.CancellationPending)
						{
							while (true)
							{
								switch (2)
								{
								case 0:
									break;
								default:
									return;
								}
							}
						}
						if (!clsPanes.IsValidWorkbook(workbook))
						{
							continue;
						}
						WorkbookItem workbookItem = new WorkbookItem(workbook);
						if (workbook != application.ActiveWorkbook)
						{
							while (true)
							{
								switch (1)
								{
								case 0:
									continue;
								}
								break;
							}
							workbookItem.Opacity = Constants.INACTIVE_FILE_OPACITY;
						}
						try
						{
							enumerator2 = workbook.Sheets.GetEnumerator();
							while (enumerator2.MoveNext())
							{
								object objectValue = RuntimeHelpers.GetObjectValue(enumerator2.Current);
								if (bgw.CancellationPending)
								{
									return;
								}
								B(Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject(VH.A(124922), NewLateBinding.LateGet(objectValue, null, VH.A(19019), new object[0], null, null, null)), VH.A(39851))));
								if (!(objectValue is Worksheet))
								{
									continue;
								}
								while (true)
								{
									switch (7)
									{
									case 0:
										continue;
									}
									break;
								}
								Worksheet worksheet = (Worksheet)objectValue;
								if (worksheet.ProtectContents)
								{
									try
									{
										worksheet.Unprotect(RuntimeHelpers.GetObjectValue(Missing.Value));
									}
									catch (Exception ex3)
									{
										ProjectData.SetProjectError(ex3);
										Exception ex4 = ex3;
										ProjectData.ClearProjectError();
									}
								}
								this.A(worksheet, workbookItem);
							}
							while (true)
							{
								switch (3)
								{
								case 0:
									break;
								default:
									goto end_IL_0608;
								}
								continue;
								end_IL_0608:
								break;
							}
						}
						finally
						{
							if (enumerator2 is IDisposable)
							{
								while (true)
								{
									switch (5)
									{
									case 0:
										continue;
									}
									(enumerator2 as IDisposable).Dispose();
									break;
								}
							}
						}
						if (workbookItem.Sheets.Count > 0)
						{
							SearchResults.Add(workbookItem);
						}
					}
					while (true)
					{
						switch (2)
						{
						case 0:
							break;
						default:
							return;
						}
					}
				}
				finally
				{
					if (enumerator is IDisposable)
					{
						while (true)
						{
							switch (5)
							{
							case 0:
								continue;
							}
							(enumerator as IDisposable).Dispose();
							break;
						}
					}
				}
			}
			}
		}
		catch (Exception ex5)
		{
			ProjectData.SetProjectError(ex5);
			Exception ex6 = ex5;
			if (!bgw.CancellationPending)
			{
				while (true)
				{
					switch (7)
					{
					case 0:
						continue;
					}
					break;
				}
				C(ex6.Message);
			}
			ProjectData.ClearProjectError();
		}
		finally
		{
			Base.ReturnToPreviousRange(rng);
			application.ScreenUpdating = true;
			rng = null;
			application = null;
			if (!bgw.CancellationPending)
			{
				while (true)
				{
					switch (5)
					{
					case 0:
						continue;
					}
					B(VH.A(124945));
					SourceCollection = CollectionViewSource.GetDefaultView(SearchResults);
					break;
				}
			}
			else
			{
				B(VH.A(124982));
				A.Cancel = true;
			}
		}
	}

	private void A(Worksheet A, WorkbookItem B = null)
	{
		if (A.ProtectContents)
		{
			return;
		}
		IEnumerator enumerator2 = default(IEnumerator);
		while (true)
		{
			switch (7)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			WorksheetItem worksheetItem = new WorksheetItem(B, A, (Microsoft.Office.Interop.Excel.Workbook)A.Parent, 0);
			using (List<BaseQuery>.Enumerator enumerator = SearchQueries.GetEnumerator())
			{
				while (true)
				{
					if (enumerator.MoveNext())
					{
						BaseQuery current = enumerator.Current;
						if (bgw.CancellationPending)
						{
							break;
						}
						if (current is VF)
						{
							while (true)
							{
								switch (4)
								{
								case 0:
									continue;
								}
								break;
							}
							Props.SearchForm = current;
							VF vF = (VF)current;
							if (!LookInPrintAreasOnly)
							{
								while (true)
								{
									switch (2)
									{
									case 0:
										continue;
									}
									break;
								}
								vF.Query(worksheetItem, A.UsedRange);
							}
							else if (Operators.CompareString(A.PageSetup.PrintArea, string.Empty, TextCompare: false) != 0)
							{
								string[] array = Strings.Split(A.PageSetup.PrintArea, CultureInfo.CurrentCulture.TextInfo.ListSeparator, -1, CompareMethod.Text);
								foreach (string cell in array)
								{
									vF.Query(worksheetItem, ((_Worksheet)A).get_Range((object)cell, RuntimeHelpers.GetObjectValue(Missing.Value)));
								}
							}
							vF = null;
							continue;
						}
						if (current is WorksheetQuery)
						{
							((WorksheetQuery)current).Query(worksheetItem, A);
							continue;
						}
						if (current is UF)
						{
							Props.SearchForm = current;
							((UF)current).Query(worksheetItem, A);
							continue;
						}
						if (!(current is TF))
						{
							continue;
						}
						TF tF = (TF)current;
						try
						{
							enumerator2 = ((IEnumerable)A.ChartObjects(RuntimeHelpers.GetObjectValue(Missing.Value))).GetEnumerator();
							while (enumerator2.MoveNext())
							{
								ChartObject chartObject = (ChartObject)enumerator2.Current;
								tF.Query(worksheetItem, chartObject.Chart);
							}
							while (true)
							{
								switch (1)
								{
								case 0:
									break;
								default:
									goto end_IL_01e6;
								}
								continue;
								end_IL_01e6:
								break;
							}
						}
						finally
						{
							if (enumerator2 is IDisposable)
							{
								while (true)
								{
									switch (4)
									{
									case 0:
										continue;
									}
									(enumerator2 as IDisposable).Dispose();
									break;
								}
							}
						}
						tF = null;
						continue;
					}
					while (true)
					{
						switch (4)
						{
						case 0:
							break;
						default:
							goto end_IL_0221;
						}
						continue;
						end_IL_0221:
						break;
					}
					break;
				}
			}
			if (!bgw.CancellationPending)
			{
				while (true)
				{
					switch (5)
					{
					case 0:
						continue;
					}
					break;
				}
				this.A(worksheetItem, B);
			}
			worksheetItem = null;
			return;
		}
	}

	private void A(WorksheetItem A, WorkbookItem B)
	{
		int count = A.Children.Count;
		if (count <= 0)
		{
			return;
		}
		A.ResultsCount = count;
		IOrderedEnumerable<ResultItem> source = A.Children.OrderBy([SpecialName] (ResultItem resultItem) => resultItem.UiIndex);
		Func<ResultItem, string> keySelector;
		if (_Closure_0024__.A == null)
		{
			keySelector = (_Closure_0024__.A = [SpecialName] (ResultItem resultItem) => ((object)resultItem).GetType().ToString());
		}
		else
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					continue;
				}
				break;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			keySelector = _Closure_0024__.A;
		}
		IOrderedEnumerable<ResultItem> source2 = source.ThenBy(keySelector).ThenBy([SpecialName] (ResultItem resultItem) =>
		{
			Range range = resultItem.Range;
			return (range == null) ? null : NewLateBinding.LateGet(range.Cells[1, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(41347), new object[0], null, null, null);
		});
		Func<ResultItem, object> keySelector2;
		if (_Closure_0024__.B == null)
		{
			keySelector2 = (_Closure_0024__.B = [SpecialName] (ResultItem resultItem) =>
			{
				Range range = resultItem.Range;
				if (range == null)
				{
					while (true)
					{
						switch (2)
						{
						case 0:
							break;
						default:
							if (1 == 0)
							{
								/*OpCode not supported: LdMemberToken*/;
							}
							return (object)null;
						}
					}
				}
				return NewLateBinding.LateGet(range.Cells[1, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(41354), new object[0], null, null, null);
			});
		}
		else
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					continue;
				}
				break;
			}
			keySelector2 = _Closure_0024__.B;
		}
		A.Children = new ObservableCollection<ResultItem>(source2.ThenBy(keySelector2));
		if (B == null)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					SearchResults.Add(A);
					return;
				}
			}
		}
		B.Sheets.Add(A);
	}

	internal void A(Microsoft.Office.Interop.Excel.Application A)
	{
		this.A(RuntimeHelpers.GetObjectValue(A.ActiveSheet));
		P();
	}

	private void F()
	{
		Forms.WarningMessage(VH.A(125009));
	}

	private void G()
	{
		if (SearchResults != null)
		{
			SearchResults.Clear();
		}
	}

	private void RefreshResults(object sender, RoutedEventArgs e)
	{
		if (bgw != null)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					continue;
				}
				break;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			if (bgw.IsBusy)
			{
				return;
			}
			while (true)
			{
				switch (4)
				{
				case 0:
					continue;
				}
				break;
			}
		}
		E();
	}

	private void CancelButtonClicked(object sender, RoutedEventArgs e)
	{
		H();
	}

	private void H()
	{
		if (bgw == null)
		{
			return;
		}
		while (true)
		{
			switch (5)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			if (bgw.IsBusy)
			{
				bgw.CancelAsync();
			}
			return;
		}
	}

	private void A(object A)
	{
		bool? isChecked = chkHighlight.IsChecked;
		bool? flag;
		if (!isChecked.HasValue)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					continue;
				}
				break;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			flag = isChecked;
		}
		else
		{
			flag = isChecked != true;
		}
		isChecked = flag;
		if (isChecked == true)
		{
			return;
		}
		while (true)
		{
			switch (2)
			{
			case 0:
				continue;
			}
			if (!(A is Worksheet))
			{
				return;
			}
			while (true)
			{
				switch (5)
				{
				case 0:
					continue;
				}
				Worksheet worksheet = (Worksheet)A;
				if (HighlightedSheets == null)
				{
					while (true)
					{
						switch (6)
						{
						case 0:
							continue;
						}
						break;
					}
					HighlightedSheets = new List<Worksheet>();
				}
				else if (HighlightedSheets.Contains(worksheet))
				{
					worksheet = null;
					return;
				}
				worksheet.Application.ScreenUpdating = false;
				try
				{
					using IEnumerator<BaseItem> enumerator = SearchResults.GetEnumerator();
					while (enumerator.MoveNext())
					{
						BaseItem current = enumerator.Current;
						if (current is WorksheetItem)
						{
							while (true)
							{
								switch (6)
								{
								case 0:
									continue;
								}
								break;
							}
							this.A((WorksheetItem)current, worksheet);
							continue;
						}
						if (current is ExploreItem)
						{
							while (true)
							{
								switch (7)
								{
								case 0:
									continue;
								}
								break;
							}
							this.A((ExploreItem)current, worksheet);
							continue;
						}
						if (current is SearchItem)
						{
							while (true)
							{
								switch (3)
								{
								case 0:
									continue;
								}
								break;
							}
							this.A((SearchItem)current, worksheet);
							continue;
						}
						if (!(current is WorkbookItem))
						{
							continue;
						}
						while (true)
						{
							switch (4)
							{
							case 0:
								continue;
							}
							break;
						}
						B((WorkbookItem)current, worksheet);
					}
					while (true)
					{
						switch (1)
						{
						case 0:
							break;
						default:
							goto end_IL_0164;
						}
						continue;
						end_IL_0164:
						break;
					}
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					ProjectData.ClearProjectError();
				}
				worksheet.Application.ScreenUpdating = true;
				HighlightedSheets.Add(worksheet);
				worksheet = null;
				return;
			}
		}
	}

	private void B(WorkbookItem A, Worksheet B)
	{
		using IEnumerator<SheetItem> enumerator = A.Sheets.GetEnumerator();
		while (enumerator.MoveNext())
		{
			SheetItem current = enumerator.Current;
			if (!(current is WorksheetItem))
			{
				continue;
			}
			while (true)
			{
				switch (1)
				{
				case 0:
					continue;
				}
				break;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			this.A((WorksheetItem)current, B);
		}
		while (true)
		{
			switch (5)
			{
			case 0:
				break;
			default:
				return;
			}
		}
	}

	private void A(WorksheetItem A, Worksheet B)
	{
		IEnumerator<ResultItem> enumerator = default(IEnumerator<ResultItem>);
		try
		{
			enumerator = A.Children.GetEnumerator();
			while (enumerator.MoveNext())
			{
				ResultItem current = enumerator.Current;
				if (current is SearchItem)
				{
					this.A((SearchItem)current, B);
				}
				else
				{
					if (!(current is ExploreItem))
					{
						continue;
					}
					while (true)
					{
						switch (1)
						{
						case 0:
							continue;
						}
						break;
					}
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					this.A((ExploreItem)current, B);
				}
			}
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					return;
				}
			}
		}
		finally
		{
			if (enumerator != null)
			{
				while (true)
				{
					switch (7)
					{
					case 0:
						continue;
					}
					enumerator.Dispose();
					break;
				}
			}
		}
	}

	private void A(ExploreItem A, Worksheet B)
	{
		try
		{
			if (A.Worksheet != B)
			{
				return;
			}
			while (true)
			{
				switch (4)
				{
				case 0:
					continue;
				}
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				Highlighter.A(A.Range);
				return;
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	private void A(SearchItem A, Worksheet B)
	{
		try
		{
			if (A.Worksheet == B)
			{
				Highlighter.A(A.Range);
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	internal void I()
	{
		Highlighter.A();
		HighlightedSheets = null;
	}

	private void HighlightToggle(object sender, RoutedEventArgs e)
	{
		bool value = ((System.Windows.Controls.CheckBox)sender).IsChecked.Value;
		global::A.K.Settings.SuperFindHighlightResults = value;
		if (value)
		{
			A(RuntimeHelpers.GetObjectValue(MH.A.Application.ActiveSheet));
		}
		else
		{
			I();
		}
	}

	private void OnSelected(object sender, RoutedEventArgs e)
	{
		try
		{
			if (!trvResults.IsKeyboardFocusWithin)
			{
				while (true)
				{
					switch (3)
					{
					case 0:
						continue;
					}
					break;
				}
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				((TreeViewItem)sender).BringIntoView();
			}
			e.Handled = true;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	private void OnExpanded(object sender, RoutedEventArgs e)
	{
	}

	private void NodePreviewMouseDown(object sender, MouseButtonEventArgs e)
	{
		if (chkPreview.IsChecked != true)
		{
			return;
		}
		while (true)
		{
			switch (6)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			N();
			return;
		}
	}

	private void NodePreviewMouseRightButtonDown(object sender, MouseButtonEventArgs e)
	{
		BaseItem baseItem = (BaseItem)((TreeViewItem)sender).DataContext;
		if (!((BaseItem)baseItem).IsSelected)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					continue;
				}
				break;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			((BaseItem)baseItem).IsSelected = true;
		}
		baseItem = null;
	}

	private void NodeKeyDown(object sender, System.Windows.Input.KeyEventArgs e)
	{
		TreeViewItem treeViewItem = (TreeViewItem)sender;
		if (e.Key == Key.F5)
		{
			goto IL_0057;
		}
		while (true)
		{
			switch (3)
			{
			case 0:
				continue;
			}
			break;
		}
		if (1 == 0)
		{
			/*OpCode not supported: LdMemberToken*/;
		}
		if (e.Key == Key.R)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					continue;
				}
				break;
			}
			if (System.Windows.Input.Keyboard.Modifiers == ModifierKeys.Control)
			{
				while (true)
				{
					switch (7)
					{
					case 0:
						continue;
					}
					break;
				}
				goto IL_0057;
			}
		}
		if (e.Key == Key.Up)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					continue;
				}
				break;
			}
			if (e.KeyboardDevice.Modifiers == ModifierKeys.Control)
			{
				goto IL_036c;
			}
			while (true)
			{
				switch (3)
				{
				case 0:
					continue;
				}
				break;
			}
		}
		if (e.Key == Key.Down)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					continue;
				}
				break;
			}
			if (e.KeyboardDevice.Modifiers == ModifierKeys.Control)
			{
				goto IL_036c;
			}
			while (true)
			{
				switch (6)
				{
				case 0:
					continue;
				}
				break;
			}
		}
		Key key = e.Key;
		if (key <= Key.Escape)
		{
			if (key == Key.Return)
			{
				goto IL_0333;
			}
			while (true)
			{
				switch (4)
				{
				case 0:
					continue;
				}
				break;
			}
			if (key == Key.Escape)
			{
				SendKeys.Send(VH.A(124594));
				e.Handled = true;
			}
		}
		else
		{
			if (key == Key.Space)
			{
				goto IL_0333;
			}
			while (true)
			{
				switch (5)
				{
				case 0:
					continue;
				}
				break;
			}
			if (key != Key.Delete)
			{
				while (true)
				{
					switch (7)
					{
					case 0:
						continue;
					}
					break;
				}
				if (key == Key.F2)
				{
					while (true)
					{
						switch (7)
						{
						case 0:
							continue;
						}
						break;
					}
					if (treeViewItem.DataContext is SheetItem)
					{
						while (true)
						{
							switch (7)
							{
							case 0:
								continue;
							}
							break;
						}
						((SheetItem)treeViewItem.DataContext).B();
					}
					else if (treeViewItem.DataContext is HyperlinkItem)
					{
						while (true)
						{
							switch (6)
							{
							case 0:
								continue;
							}
							break;
						}
						((HyperlinkItem)treeViewItem.DataContext).Edit();
					}
					else if (treeViewItem.DataContext is ValidationItem)
					{
						while (true)
						{
							switch (4)
							{
							case 0:
								continue;
							}
							break;
						}
						((ValidationItem)treeViewItem.DataContext).Modify();
					}
					else if (treeViewItem.DataContext is ConditionalFormatItem)
					{
						while (true)
						{
							switch (2)
							{
							case 0:
								continue;
							}
							break;
						}
						((ConditionalFormatItem)treeViewItem.DataContext).Modify();
					}
					else if (treeViewItem.DataContext is ChartItem)
					{
						while (true)
						{
							switch (2)
							{
							case 0:
								continue;
							}
							break;
						}
						((ChartItem)treeViewItem.DataContext).SelectData();
					}
					else
					{
						J();
					}
					e.Handled = true;
				}
			}
			else
			{
				try
				{
					if (treeViewItem.DataContext is SheetItem)
					{
						while (true)
						{
							switch (6)
							{
							case 0:
								continue;
							}
							((SheetItem)treeViewItem.DataContext).A();
							M();
							break;
						}
					}
					else
					{
						if (!(treeViewItem.DataContext is ExploreItem))
						{
							throw new NotImplementedException();
						}
						while (true)
						{
							switch (7)
							{
							case 0:
								continue;
							}
							((ExploreItem)treeViewItem.DataContext).Delete();
							M();
							break;
						}
					}
				}
				catch (NotImplementedException ex)
				{
					ProjectData.SetProjectError(ex);
					NotImplementedException ex2 = ex;
					J();
					ProjectData.ClearProjectError();
				}
				e.Handled = true;
			}
		}
		goto IL_036c;
		IL_0333:
		if (treeViewItem.DataContext is HyperlinkItem)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					continue;
				}
				break;
			}
			B((HyperlinkItem)treeViewItem.DataContext);
		}
		else
		{
			J();
		}
		e.Handled = true;
		goto IL_036c;
		IL_0057:
		try
		{
			if (!(treeViewItem.DataContext is WorkbookItem))
			{
				while (true)
				{
					switch (7)
					{
					case 0:
						continue;
					}
					if (treeViewItem.DataContext is WorksheetItem)
					{
						break;
					}
					while (true)
					{
						switch (1)
						{
						case 0:
							continue;
						}
						if (treeViewItem.DataContext is ExploreItem)
						{
							while (true)
							{
								switch (3)
								{
								case 0:
									continue;
								}
								((ExploreItem)treeViewItem.DataContext).Refresh();
								break;
							}
							break;
						}
						throw new NotImplementedException();
					}
					break;
				}
			}
		}
		catch (NotImplementedException ex3)
		{
			ProjectData.SetProjectError(ex3);
			NotImplementedException ex4 = ex3;
			J();
			ProjectData.ClearProjectError();
		}
		e.Handled = true;
		goto IL_036c;
		IL_036c:
		treeViewItem = null;
	}

	private void J()
	{
		D(VH.A(125094));
		M();
	}

	private void OnRequestBringIntoView(object sender, RequestBringIntoViewEventArgs e)
	{
		if (this.m_A == null)
		{
			this.m_A = (ScrollViewer)trvResults.Template.FindName(VH.A(43250), trvResults);
			if (this.m_A != null)
			{
				this.m_A.ScrollChanged -= TreeViewScrollViewerScrollChanged;
				this.m_A.ScrollChanged += TreeViewScrollViewerScrollChanged;
			}
		}
		this.m_B = true;
		this.m_A = this.m_A.HorizontalOffset;
	}

	private void TreeViewScrollViewerScrollChanged(object sender, ScrollChangedEventArgs e)
	{
		if (this.m_B)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					continue;
				}
				break;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			this.m_A.ScrollToHorizontalOffset(this.m_A);
		}
		this.m_B = false;
	}

	private DependencyObject A(DependencyObject A)
	{
		DependencyObject parent = VisualTreeHelper.GetParent(A);
		if (!(parent is System.Windows.Controls.TreeView))
		{
			return this.A(parent);
		}
		return parent;
	}

	private void TreeViewGotKeyboardFocus(object sender, KeyboardFocusChangedEventArgs e)
	{
	}

	private void TreeViewLostKeyboardFocus(object sender, KeyboardFocusChangedEventArgs e)
	{
	}

	private void TreeViewPreviewKeyDown(object sender, System.Windows.Input.KeyEventArgs e)
	{
		if (!e.IsRepeat)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					continue;
				}
				break;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			N();
		}
		Key key = e.Key;
		if (key != Key.Up)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					continue;
				}
				break;
			}
			if (key != Key.Down)
			{
				return;
			}
		}
		if (!e.IsRepeat)
		{
			RemoveTreeViewSelectionChangedHandler();
			trvResults.KeyUp += TreeViewNavKeyUp;
		}
	}

	private void TreeViewNavKeyUp(object sender, System.Windows.Input.KeyEventArgs e)
	{
		trvResults.KeyUp -= TreeViewNavKeyUp;
		K();
		L();
		e.Handled = true;
	}

	private void K()
	{
		trvResults.SelectedItemChanged += this.m_A;
	}

	public void RemoveTreeViewSelectionChangedHandler()
	{
		trvResults.SelectedItemChanged -= this.m_A;
	}

	private void TreeViewSelectionChanged(System.Windows.Controls.TreeView sender, RoutedPropertyChangedEventArgs<object> e)
	{
		if (!trvResults.IsKeyboardFocusWithin)
		{
			return;
		}
		while (true)
		{
			switch (1)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			base.Dispatcher.BeginInvoke((global::A.A)([SpecialName] () =>
			{
				L();
			}), DispatcherPriority.Normal);
			e.Handled = true;
			return;
		}
	}

	private void L()
	{
		if (trvResults.SelectedItem == null)
		{
			return;
		}
		BaseItem baseItem = (BaseItem)trvResults.SelectedItem;
		Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
		Microsoft.Office.Interop.Excel.Workbook activeWorkbook = application.ActiveWorkbook;
		try
		{
			if (baseItem is WorksheetItem)
			{
				while (true)
				{
					switch (5)
					{
					case 0:
						continue;
					}
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					A((SheetItem)baseItem, application);
					break;
				}
			}
			else if (baseItem is WorkbookItem)
			{
				A((WorkbookItem)baseItem);
			}
			else
			{
				application.ScreenUpdating = false;
				if (baseItem is SearchItem)
				{
					SearchItem searchItem = (SearchItem)baseItem;
					searchItem.Workbook.Activate();
					if (ExcelAddIn1.Sheets.Visibility.A(searchItem.Worksheet))
					{
						while (true)
						{
							switch (5)
							{
							case 0:
								continue;
							}
							break;
						}
						searchItem.Worksheet.Activate();
						try
						{
							A(searchItem.Range);
						}
						catch (Exception ex)
						{
							ProjectData.SetProjectError(ex);
							Exception ex2 = ex;
							D(VH.A(37357));
							A((ResultItem)baseItem);
							ProjectData.ClearProjectError();
						}
					}
					searchItem = null;
				}
				else if (baseItem is ExploreItem)
				{
					while (true)
					{
						switch (4)
						{
						case 0:
							continue;
						}
						ExploreItem exploreItem = (ExploreItem)baseItem;
						exploreItem.Workbook.Activate();
						if (ExcelAddIn1.Sheets.Visibility.A(exploreItem.Worksheet))
						{
							exploreItem.Worksheet.Activate();
							if (baseItem is ChartItem)
							{
								while (true)
								{
									switch (2)
									{
									case 0:
										continue;
									}
									break;
								}
								A((ChartItem)baseItem, activeWorkbook);
							}
							else if (baseItem is ShapeItem)
							{
								while (true)
								{
									switch (1)
									{
									case 0:
										continue;
									}
									break;
								}
								A((ShapeItem)baseItem, activeWorkbook);
							}
							else if (baseItem is NoteItem)
							{
								while (true)
								{
									switch (5)
									{
									case 0:
										continue;
									}
									break;
								}
								A((NoteItem)baseItem);
							}
							else if (baseItem is CommentItem)
							{
								A((CommentItem)baseItem);
							}
							else if (baseItem is AutoFilterItem)
							{
								while (true)
								{
									switch (7)
									{
									case 0:
										continue;
									}
									break;
								}
								A((AutoFilterItem)baseItem);
							}
							else if (baseItem is TableItem)
							{
								A((TableItem)baseItem);
							}
							else if (baseItem is QueryTableItem)
							{
								while (true)
								{
									switch (7)
									{
									case 0:
										continue;
									}
									break;
								}
								A((QueryTableItem)baseItem);
							}
							else if (baseItem is PivotTableItem)
							{
								while (true)
								{
									switch (4)
									{
									case 0:
										continue;
									}
									break;
								}
								A((PivotTableItem)baseItem);
							}
							else if (baseItem is SparklineItem)
							{
								A((SparklineItem)baseItem);
							}
							else if (baseItem is WatchItem)
							{
								while (true)
								{
									switch (7)
									{
									case 0:
										continue;
									}
									break;
								}
								A((WatchItem)baseItem);
							}
							else if (baseItem is HyperlinkItem)
							{
								while (true)
								{
									switch (3)
									{
									case 0:
										continue;
									}
									break;
								}
								A((HyperlinkItem)baseItem);
							}
							else if (exploreItem.Range != null)
							{
								while (true)
								{
									switch (6)
									{
									case 0:
										continue;
									}
									break;
								}
								try
								{
									A(exploreItem.Range);
								}
								catch (Exception ex3)
								{
									ProjectData.SetProjectError(ex3);
									Exception ex4 = ex3;
									D(VH.A(37357));
									A((ResultItem)baseItem);
									ProjectData.ClearProjectError();
								}
							}
						}
						exploreItem = null;
						break;
					}
				}
			}
		}
		catch (Exception ex5)
		{
			ProjectData.SetProjectError(ex5);
			Exception ex6 = ex5;
			C(ex6.Message);
			M();
			ProjectData.ClearProjectError();
		}
		application.ScreenUpdating = true;
		baseItem = null;
		activeWorkbook = null;
		application = null;
	}

	private void A(WorkbookItem A)
	{
		A.Workbook.Activate();
	}

	private void A(SheetItem A, Microsoft.Office.Interop.Excel.Application B)
	{
		Microsoft.Office.Interop.Excel.Window window = null;
		SheetItem sheetItem = A;
		Microsoft.Office.Interop.Excel.Workbook workbook;
		try
		{
			if (!ExcelAddIn1.Sheets.Visibility.A(RuntimeHelpers.GetObjectValue(sheetItem.Sheet)))
			{
				return;
			}
			workbook = sheetItem.Workbook;
			if (workbook == B.ActiveWorkbook)
			{
				while (true)
				{
					switch (6)
					{
					case 0:
						continue;
					}
					break;
				}
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				if (workbook.Windows.Count > 1)
				{
					while (true)
					{
						switch (2)
						{
						case 0:
							continue;
						}
						break;
					}
					window = B.ActiveWindow;
					B.ScreenUpdating = false;
				}
			}
			workbook.Activate();
			if (window != null)
			{
				window.Activate();
				M();
			}
			NewLateBinding.LateCall(sheetItem.Sheet, null, VH.A(39985), new object[0], null, null, null, IgnoreReturn: true);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			sheetItem.Parent.Sheets.Remove(A);
			ProjectData.ClearProjectError();
		}
		sheetItem = null;
		workbook = null;
		window = null;
	}

	private void A(Range A)
	{
		Ranges.ScrollIntoView(A);
		B(A);
	}

	private void A(ChartItem A, Microsoft.Office.Interop.Excel.Workbook B)
	{
		try
		{
			ChartObject chartObject = A.ChartObject;
			Ranges.ScrollIntoView(((_Application)chartObject.Application).get_Range((object)chartObject.TopLeftCell, (object)chartObject.BottomRightCell));
			chartObject.Chart.ChartArea.Select();
			chartObject = null;
			M();
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			try
			{
				if (A.Worksheet.ProtectContents)
				{
					while (true)
					{
						switch (5)
						{
						case 0:
							break;
						default:
							if (1 == 0)
							{
								/*OpCode not supported: LdMemberToken*/;
							}
							ProjectData.ClearProjectError();
							return;
						}
					}
				}
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				ProjectData.ClearProjectError();
			}
			D(VH.A(125185));
			this.A(A);
			ProjectData.ClearProjectError();
		}
	}

	private void A(ShapeItem A, Microsoft.Office.Interop.Excel.Workbook B)
	{
		try
		{
			ExcelAddIn1.Shapes.Navigate.A(A.Shape);
			M();
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			try
			{
				if (A.Worksheet.ProtectContents)
				{
					while (true)
					{
						switch (5)
						{
						case 0:
							break;
						default:
							if (1 == 0)
							{
								/*OpCode not supported: LdMemberToken*/;
							}
							ProjectData.ClearProjectError();
							return;
						}
					}
				}
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				ProjectData.ClearProjectError();
			}
			D(VH.A(125218));
			this.A(A);
			ProjectData.ClearProjectError();
		}
	}

	private void A(NoteItem A)
	{
		try
		{
			_ = A.Note.Shape.Height;
			this.A((Range)A.Note.Parent);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			D(VH.A(125251));
			this.A((ResultItem)A);
			ProjectData.ClearProjectError();
		}
	}

	private void A(CommentItem A)
	{
		try
		{
			this.A((Range)NewLateBinding.LateGet(A.Comment, null, VH.A(8701), new object[0], null, null, null));
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			D(VH.A(125282));
			this.A((ResultItem)A);
			ProjectData.ClearProjectError();
		}
	}

	private void A(AutoFilterItem A)
	{
		try
		{
			this.A(A.AutoFilter.Range);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			D(VH.A(125319));
			this.A((ResultItem)A);
			ProjectData.ClearProjectError();
		}
	}

	private void A(TableItem A)
	{
		try
		{
			this.A(A.Table.Range);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			D(VH.A(125354));
			this.A((ResultItem)A);
			ProjectData.ClearProjectError();
		}
	}

	private void A(QueryTableItem A)
	{
		try
		{
			this.A(A.QueryTable.ResultRange);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			D(VH.A(125387));
			this.A((ResultItem)A);
			ProjectData.ClearProjectError();
		}
	}

	private void A(PivotTableItem A)
	{
		try
		{
			this.A(A.PivotTable.TableRange1);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			D(VH.A(125432));
			this.A((ResultItem)A);
			ProjectData.ClearProjectError();
		}
	}

	private void A(SparklineItem A)
	{
		try
		{
			this.A(A.SparklineGroup.Location);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			D(VH.A(125477));
			this.A((ResultItem)A);
			ProjectData.ClearProjectError();
		}
	}

	private void A(WatchItem A)
	{
		try
		{
			this.A((Range)A.Watch.Source);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			D(VH.A(125518));
			this.A((ResultItem)A);
			ProjectData.ClearProjectError();
		}
	}

	private void A(HyperlinkItem A)
	{
		try
		{
			this.A(A.Hyperlink.Range);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			D(VH.A(125551));
			this.A((ResultItem)A);
			ProjectData.ClearProjectError();
		}
	}

	private void A(ResultItem A)
	{
		if (chkHighlight.IsChecked == true)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					continue;
				}
				break;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			try
			{
				Highlighter.B(A.Range);
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
		}
		A.Parent.A(A);
		M();
	}

	private void M()
	{
		try
		{
			thisCTP.BringToFront();
			thisCTP.Focus();
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		trvResults.Focus();
	}

	private void ExpandCollapseChanged(object sender, RoutedEventArgs e)
	{
		A(chkExpandCollapse.IsChecked.Value);
		trvResults.Focus();
	}

	private void A(bool A)
	{
		NF a = default(NF);
		NF CS_0024_003C_003E8__locals6 = new NF(a);
		CS_0024_003C_003E8__locals6.A = this;
		CS_0024_003C_003E8__locals6.A = A;
		RemoveTreeViewSelectionChangedHandler();
		Task.Factory.StartNew([SpecialName] () =>
		{
			IEnumerator<BaseItem> enumerator = default(IEnumerator<BaseItem>);
			try
			{
				enumerator = CS_0024_003C_003E8__locals6.A.SearchResults.GetEnumerator();
				IEnumerator<SheetItem> enumerator2 = default(IEnumerator<SheetItem>);
				while (enumerator.MoveNext())
				{
					BaseItem current = enumerator.Current;
					((BaseItem)current).IsExpanded = CS_0024_003C_003E8__locals6.A;
					if (current is WorkbookItem)
					{
						while (true)
						{
							switch (5)
							{
							case 0:
								continue;
							}
							break;
						}
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						try
						{
							enumerator2 = ((WorkbookItem)current).Sheets.GetEnumerator();
							while (enumerator2.MoveNext())
							{
								((BaseItem)enumerator2.Current).IsExpanded = true;
							}
						}
						finally
						{
							if (enumerator2 != null)
							{
								while (true)
								{
									switch (4)
									{
									case 0:
										break;
									default:
										enumerator2.Dispose();
										goto end_IL_007c;
									}
									continue;
									end_IL_007c:
									break;
								}
							}
						}
					}
				}
			}
			finally
			{
				if (enumerator != null)
				{
					while (true)
					{
						switch (4)
						{
						case 0:
							break;
						default:
							enumerator.Dispose();
							goto end_IL_009f;
						}
						continue;
						end_IL_009f:
						break;
					}
				}
			}
			CS_0024_003C_003E8__locals6.A.Dispatcher.BeginInvoke((global::A.A)([SpecialName] () =>
			{
				CS_0024_003C_003E8__locals6.A.K();
			}));
		});
	}

	private void MouseEnterNode(object sender, System.Windows.Input.MouseEventArgs e)
	{
		//IL_029f: Unknown result type (might be due to invalid IL or missing references)
		//IL_02a9: Expected O, but got Unknown
		bool? isChecked;
		bool? flag = (isChecked = chkPreview.IsChecked);
		bool? flag2;
		if (flag.HasValue)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					continue;
				}
				break;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			if (isChecked != true)
			{
				flag2 = false;
				goto IL_0068;
			}
			while (true)
			{
				switch (6)
				{
				case 0:
					continue;
				}
				break;
			}
		}
		if (this.m_B != null)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					continue;
				}
				break;
			}
			flag2 = false;
		}
		else
		{
			flag2 = isChecked;
		}
		goto IL_0068;
		IL_0068:
		bool? flag3 = flag2;
		if (flag3.HasValue)
		{
			if (flag3 != true)
			{
				return;
			}
			while (true)
			{
				switch (1)
				{
				case 0:
					continue;
				}
				break;
			}
		}
		if (System.Windows.Input.Keyboard.Modifiers != ModifierKeys.None)
		{
			return;
		}
		while (true)
		{
			switch (5)
			{
			case 0:
				continue;
			}
			if (!flag3.HasValue)
			{
				return;
			}
			if (((TreeViewItem)sender).HasItems)
			{
				return;
			}
			BaseItem A;
			BitmapSource A2;
			while (true)
			{
				switch (5)
				{
				case 0:
					continue;
				}
				A = (BaseItem)((TreeViewItem)sender).DataContext;
				A2 = null;
				if (this.m_A != null)
				{
					while (true)
					{
						switch (6)
						{
						case 0:
							continue;
						}
						break;
					}
					this.m_A.Stop();
					this.m_A = null;
				}
				if (A is ChartItem)
				{
					while (true)
					{
						switch (7)
						{
						case 0:
							continue;
						}
						break;
					}
					ChartItem chartItem = (ChartItem)A;
					if (chartItem.PreviewImage != null)
					{
						while (true)
						{
							switch (7)
							{
							case 0:
								continue;
							}
							break;
						}
						A2 = chartItem.PreviewImage;
					}
					else
					{
						try
						{
							string text = modFunctionsIO.PathGetTempFileName();
							chartItem.ChartObject.Chart.Export(text, VH.A(125592), RuntimeHelpers.GetObjectValue(Missing.Value));
							A2 = Preview.GetBitmapImage(text);
							chartItem.PreviewImage = A2;
						}
						catch (Exception ex)
						{
							ProjectData.SetProjectError(ex);
							Exception ex2 = ex;
							ProjectData.ClearProjectError();
						}
					}
					chartItem = null;
				}
				else if (A is ShapeItem)
				{
					while (true)
					{
						switch (6)
						{
						case 0:
							continue;
						}
						break;
					}
					ShapeItem shapeItem = (ShapeItem)A;
					if (shapeItem.PreviewImage != null)
					{
						while (true)
						{
							switch (3)
							{
							case 0:
								continue;
							}
							break;
						}
						A2 = shapeItem.PreviewImage;
					}
					else
					{
						Shape shape;
						try
						{
							shape = ((ShapeItem)A).Shape;
							this.A(shape);
							A2 = this.A();
							shapeItem.PreviewImage = A2;
						}
						catch (Exception ex3)
						{
							ProjectData.SetProjectError(ex3);
							Exception ex4 = ex3;
							ProjectData.ClearProjectError();
						}
						shape = null;
					}
					shapeItem = null;
				}
				else
				{
					_ = A is NoteItem;
				}
				if (A2 != null)
				{
					while (true)
					{
						switch (6)
						{
						case 0:
							continue;
						}
						break;
					}
					try
					{
						if (PreviewWindow == null)
						{
							while (true)
							{
								switch (1)
								{
								case 0:
									continue;
								}
								break;
							}
							PreviewWindow = new wpfPreview();
						}
						base.Dispatcher.Invoke([SpecialName] () =>
						{
							Preview.PrepareWindow(PreviewWindow, (BaseItem)(object)A, A2, PointToScreen(new System.Windows.Point(0.0, 0.0)));
						}, DispatcherPriority.Background);
					}
					catch (Exception ex5)
					{
						ProjectData.SetProjectError(ex5);
						Exception ex6 = ex5;
						ProjectData.ClearProjectError();
					}
					A2 = null;
					e.Handled = true;
				}
				else
				{
					N();
				}
				A = null;
				return;
			}
		}
	}

	private void MouseLeaveNode(object sender, System.Windows.Input.MouseEventArgs e)
	{
		bool? isChecked;
		bool? flag = (isChecked = chkPreview.IsChecked);
		bool? obj;
		if (flag.HasValue)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					continue;
				}
				break;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			if (isChecked != true)
			{
				obj = false;
				goto IL_0056;
			}
		}
		obj = ((PreviewWindow != null) ? isChecked : new bool?(false));
		goto IL_0056;
		IL_0056:
		bool? flag2 = obj;
		if (flag2.HasValue)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					continue;
				}
				break;
			}
			if (flag2 != true)
			{
				return;
			}
			while (true)
			{
				switch (2)
				{
				case 0:
					continue;
				}
				break;
			}
		}
		if (!((UIElement)(object)PreviewWindow).IsVisible || !flag2.HasValue)
		{
			return;
		}
		while (true)
		{
			switch (3)
			{
			case 0:
				continue;
			}
			this.m_A = new DispatcherTimer();
			this.m_A.Interval = TimeSpan.FromMilliseconds(100.0);
			this.m_A.Tick += [SpecialName] (object obj2, EventArgs ev) =>
			{
				try
				{
					this.m_A.Stop();
					N();
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					ProjectData.ClearProjectError();
				}
			};
			this.m_A.Start();
			e.Handled = true;
			return;
		}
	}

	private void MouseMoveNode(object sender, System.Windows.Input.MouseEventArgs e)
	{
		if (chkPreview.IsChecked != true)
		{
			return;
		}
		while (true)
		{
			switch (1)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			if (PreviewWindow != null)
			{
				while (true)
				{
					switch (5)
					{
					case 0:
						continue;
					}
					break;
				}
				if (!((UIElement)(object)PreviewWindow).IsVisible)
				{
					while (true)
					{
						switch (4)
						{
						case 0:
							continue;
						}
						break;
					}
					((System.Windows.Window)(object)PreviewWindow).Show();
				}
				A(e);
			}
			e.Handled = true;
			return;
		}
	}

	private void A(System.Windows.Input.MouseEventArgs A)
	{
		System.Windows.Point point = trvResults.PointToScreen(new System.Windows.Point(0.0, 0.0));
		wpfPreview previewWindow = PreviewWindow;
		if (DockPosition == MsoCTPDockPosition.msoCTPDockPositionRight)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					continue;
				}
				break;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			((System.Windows.Window)(object)previewWindow).Left = point.X - ((FrameworkElement)(object)previewWindow).ActualWidth;
		}
		else
		{
			((System.Windows.Window)(object)previewWindow).Left = point.X + trvResults.ActualWidth + 4.0;
		}
		((System.Windows.Window)(object)previewWindow).Top = point.Y;
		previewWindow = null;
	}

	private void N()
	{
		if (PreviewWindow == null)
		{
			return;
		}
		while (true)
		{
			switch (7)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			((System.Windows.Window)(object)PreviewWindow).Close();
			PreviewWindow = null;
			return;
		}
	}

	private void PreviewToggle(object sender, RoutedEventArgs e)
	{
		N();
		global::A.K.Settings.ExplorerPreviews = chkPreview.IsChecked.Value;
	}

	private void A(Shape A)
	{
		A.CopyPicture(XlPictureAppearance.xlScreen, XlCopyPictureFormat.xlBitmap);
	}

	private BitmapSource A()
	{
		Bitmap bitmap = (Bitmap)System.Windows.Forms.Clipboard.GetImage();
		if (bitmap != null)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					return Imaging.CreateBitmapSourceFromHBitmap(bitmap.GetHbitmap(), IntPtr.Zero, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions());
				}
			}
		}
		return null;
	}

	private void trvExplorer_PreviewMouseWheel(object sender, MouseWheelEventArgs e)
	{
		N();
		if (this.m_B != null)
		{
			this.m_B.Stop();
		}
		this.m_B = new DispatcherTimer();
		this.m_B.Interval = TimeSpan.FromMilliseconds(100.0);
		this.m_B.Tick += ScrollingComplete;
		this.m_B.Start();
	}

	private void ScrollingComplete(object sender, EventArgs e)
	{
		this.m_B.Stop();
		this.m_B = null;
	}

	private void CellModeChecked(object sender, RoutedEventArgs e)
	{
		global::A.K.Settings.AdvancedFindSelectMode = 0;
		O();
	}

	private void RowModeChecked(object sender, RoutedEventArgs e)
	{
		global::A.K.Settings.AdvancedFindSelectMode = 1;
		O();
	}

	private void ColumnModeChecked(object sender, RoutedEventArgs e)
	{
		global::A.K.Settings.AdvancedFindSelectMode = 2;
		O();
	}

	private void O()
	{
		System.Windows.Controls.TreeView treeView = trvResults;
		treeView.Focus();
		try
		{
			if (treeView.SelectedItem is SearchItem)
			{
				while (true)
				{
					switch (6)
					{
					case 0:
						continue;
					}
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					B(((SearchItem)treeView.SelectedItem).Range);
					break;
				}
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		treeView = null;
	}

	private void btnSelect_Click(object sender, RoutedEventArgs e)
	{
		Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
		Dictionary<Worksheet, Range> B = new Dictionary<Worksheet, Range>();
		if (application.ActiveSheet is Worksheet)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					continue;
				}
				break;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			if (SearchResults.Any())
			{
				Worksheet worksheet = (Worksheet)application.ActiveSheet;
				application.ScreenUpdating = false;
				try
				{
					using (IEnumerator<BaseItem> enumerator = SearchResults.GetEnumerator())
					{
						while (enumerator.MoveNext())
						{
							BaseItem current = enumerator.Current;
							if (current is WorksheetItem)
							{
								while (true)
								{
									switch (2)
									{
									case 0:
										continue;
									}
									break;
								}
								A((WorksheetItem)current, ref B);
							}
							else if (current is SearchItem)
							{
								while (true)
								{
									switch (6)
									{
									case 0:
										continue;
									}
									break;
								}
								A((SearchItem)current, ref B);
							}
							else if (current is ExploreItem)
							{
								while (true)
								{
									switch (4)
									{
									case 0:
										continue;
									}
									break;
								}
								A((ExploreItem)current, ref B);
							}
							else
							{
								if (!(current is WorkbookItem))
								{
									continue;
								}
								using IEnumerator<SheetItem> enumerator2 = ((WorkbookItem)current).Sheets.GetEnumerator();
								while (enumerator2.MoveNext())
								{
									SheetItem current2 = enumerator2.Current;
									if (!(current2 is WorksheetItem))
									{
										continue;
									}
									while (true)
									{
										switch (4)
										{
										case 0:
											continue;
										}
										break;
									}
									A((WorksheetItem)current2, ref B);
								}
								while (true)
								{
									switch (7)
									{
									case 0:
										break;
									default:
										goto end_IL_0147;
									}
									continue;
									end_IL_0147:
									break;
								}
							}
						}
						while (true)
						{
							switch (4)
							{
							case 0:
								break;
							default:
								goto end_IL_016c;
							}
							continue;
							end_IL_016c:
							break;
						}
					}
					foreach (KeyValuePair<Worksheet, Range> item in B)
					{
						item.Key.Activate();
						this.B(item.Value);
					}
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					ProjectData.ClearProjectError();
				}
				worksheet.Activate();
				worksheet = null;
				application.ScreenUpdating = true;
			}
		}
		application = null;
		B = null;
	}

	private void A(WorksheetItem A, ref Dictionary<Worksheet, Range> B)
	{
		using IEnumerator<ResultItem> enumerator = A.Children.GetEnumerator();
		while (enumerator.MoveNext())
		{
			ResultItem current = enumerator.Current;
			if (current is SearchItem)
			{
				while (true)
				{
					switch (5)
					{
					case 0:
						continue;
					}
					break;
				}
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				this.A((SearchItem)current, ref B);
			}
			else
			{
				if (!(current is ExploreItem))
				{
					continue;
				}
				while (true)
				{
					switch (3)
					{
					case 0:
						continue;
					}
					break;
				}
				this.A((ExploreItem)current, ref B);
			}
		}
		while (true)
		{
			switch (4)
			{
			case 0:
				break;
			default:
				return;
			}
		}
	}

	private void A(SearchItem A, ref Dictionary<Worksheet, Range> B)
	{
		SearchItem searchItem = A;
		this.A(ref B, searchItem.Worksheet, searchItem.Range);
		searchItem = null;
	}

	private void A(ExploreItem A, ref Dictionary<Worksheet, Range> B)
	{
		ExploreItem exploreItem = A;
		this.A(ref B, exploreItem.Worksheet, exploreItem.Range);
		exploreItem = null;
	}

	private void A(ref Dictionary<Worksheet, Range> A, Worksheet B, Range C)
	{
		if (C == null)
		{
			return;
		}
		while (true)
		{
			switch (1)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			if (A.ContainsKey(B))
			{
				Dictionary<Worksheet, Range> obj = A;
				Worksheet key;
				Range A2 = obj[key = B];
				RangeHelpers.A(ref A2, C);
				obj[key] = A2;
			}
			else
			{
				A.Add(B, C);
			}
			return;
		}
	}

	private void B(Range A)
	{
		if (radCell.IsChecked == true)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					A.Select();
					return;
				}
			}
		}
		if (radRow.IsChecked == true)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					A.EntireRow.Select();
					A.Activate();
					return;
				}
			}
		}
		if (radColumn.IsChecked != true)
		{
			return;
		}
		while (true)
		{
			switch (6)
			{
			case 0:
				continue;
			}
			A.EntireColumn.Select();
			A.Activate();
			return;
		}
	}

	private void BackToSearchView(object sender, RoutedEventArgs e)
	{
		H();
		I();
		Q();
		ParentView.A();
	}

	private void P()
	{
		Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
		new ComAwareEventInfo(typeof(AppEvents_Event), VH.A(125599)).AddEventHandler(application, new AppEvents_AfterCalculateEventHandler(R));
		new ComAwareEventInfo(typeof(AppEvents_Event), VH.A(125628)).AddEventHandler(application, new AppEvents_SheetActivateEventHandler(B));
		new ComAwareEventInfo(typeof(AppEvents_Event), VH.A(125655)).AddEventHandler(application, new AppEvents_WorkbookBeforeSaveEventHandler(A));
		application = null;
	}

	internal void Q()
	{
		Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
		new ComAwareEventInfo(typeof(AppEvents_Event), VH.A(125599)).RemoveEventHandler(application, new AppEvents_AfterCalculateEventHandler(R));
		new ComAwareEventInfo(typeof(AppEvents_Event), VH.A(125628)).RemoveEventHandler(application, new AppEvents_SheetActivateEventHandler(B));
		new ComAwareEventInfo(typeof(AppEvents_Event), VH.A(125655)).RemoveEventHandler(application, new AppEvents_WorkbookBeforeSaveEventHandler(A));
		application = null;
	}

	private void R()
	{
		IEnumerator<BaseItem> enumerator = default(IEnumerator<BaseItem>);
		try
		{
			enumerator = SearchResults.GetEnumerator();
			while (enumerator.MoveNext())
			{
				BaseItem current = enumerator.Current;
				try
				{
					if (current is WorksheetItem)
					{
						while (true)
						{
							switch (6)
							{
							case 0:
								continue;
							}
							if (1 == 0)
							{
								/*OpCode not supported: LdMemberToken*/;
							}
							A((WorksheetItem)current);
							break;
						}
					}
					else if (current is ExploreItem)
					{
						((ExploreItem)current).Refresh();
					}
					else if (current is SearchItem)
					{
						while (true)
						{
							switch (2)
							{
							case 0:
								continue;
							}
							((SearchItem)current).Refresh();
							break;
						}
					}
					else
					{
						if (!(current is WorkbookItem))
						{
							continue;
						}
						while (true)
						{
							switch (5)
							{
							case 0:
								continue;
							}
							B((WorkbookItem)current);
							break;
						}
						continue;
					}
				}
				catch (NotImplementedException ex)
				{
					ProjectData.SetProjectError(ex);
					NotImplementedException ex2 = ex;
					ProjectData.ClearProjectError();
				}
				catch (Exception ex3)
				{
					ProjectData.SetProjectError(ex3);
					Exception ex4 = ex3;
					T();
					ProjectData.ClearProjectError();
					return;
				}
			}
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					return;
				}
			}
		}
		finally
		{
			if (enumerator != null)
			{
				while (true)
				{
					switch (5)
					{
					case 0:
						continue;
					}
					enumerator.Dispose();
					break;
				}
			}
		}
	}

	private void B(WorkbookItem A)
	{
		A.A();
		IEnumerator<SheetItem> enumerator = default(IEnumerator<SheetItem>);
		try
		{
			enumerator = A.Sheets.GetEnumerator();
			while (enumerator.MoveNext())
			{
				SheetItem current = enumerator.Current;
				if (!(current is WorksheetItem))
				{
					continue;
				}
				while (true)
				{
					switch (5)
					{
					case 0:
						continue;
					}
					break;
				}
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				this.A((WorksheetItem)current);
			}
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					return;
				}
			}
		}
		finally
		{
			if (enumerator != null)
			{
				while (true)
				{
					switch (2)
					{
					case 0:
						continue;
					}
					enumerator.Dispose();
					break;
				}
			}
		}
	}

	private void A(WorksheetItem A)
	{
		A.B();
		using IEnumerator<ResultItem> enumerator = A.Children.GetEnumerator();
		while (enumerator.MoveNext())
		{
			ResultItem current = enumerator.Current;
			ExploreItem obj = current as ExploreItem;
			if (obj == null)
			{
				while (true)
				{
					switch (1)
					{
					case 0:
						continue;
					}
					break;
				}
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
			}
			else
			{
				obj.Refresh();
			}
			(current as SearchItem)?.Refresh();
		}
		while (true)
		{
			switch (5)
			{
			case 0:
				break;
			default:
				return;
			}
		}
	}

	private void B(object A)
	{
		this.A(RuntimeHelpers.GetObjectValue(A));
	}

	private void A(Microsoft.Office.Interop.Excel.Workbook A, bool B, ref bool C)
	{
		chkHighlight.IsChecked = false;
	}

	private void OnGotFocusEvent(object sender, RoutedEventArgs e)
	{
		if (LogicalTreeHelper.GetParent((DependencyObject)e.OriginalSource) != ActiveMenu)
		{
			return;
		}
		while (true)
		{
			switch (4)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			base.Dispatcher.BeginInvoke((D<bool>)([SpecialName] () =>
			{
				this.m_A = new DispatcherFrame();
				Dispatcher.PushFrame(this.m_A);
				return false;
			}), DispatcherPriority.Normal);
			return;
		}
	}

	private void OnLostFocusEvent(object sender, RoutedEventArgs e)
	{
		if (LogicalTreeHelper.GetParent((DependencyObject)e.OriginalSource) == ActiveMenu)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					continue;
				}
				break;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			this.m_A.Continue = false;
		}
		Dispatcher.ExitAllFrames();
	}

	private void A(System.Windows.Controls.ContextMenu A)
	{
		ActiveMenu = A;
		A.RemoveHandler(UIElement.GotFocusEvent, new RoutedEventHandler(OnGotFocusEvent));
		A.RemoveHandler(UIElement.LostFocusEvent, new RoutedEventHandler(OnLostFocusEvent));
		A.AddHandler(UIElement.GotFocusEvent, new RoutedEventHandler(OnGotFocusEvent));
		A.AddHandler(UIElement.LostFocusEvent, new RoutedEventHandler(OnLostFocusEvent));
		_ = null;
	}

	private void MenuOpened(object sender, RoutedEventArgs e)
	{
		A((System.Windows.Controls.ContextMenu)sender);
	}

	private void SubMenuOpened(object sender, RoutedEventArgs e)
	{
		ActiveMenu = RuntimeHelpers.GetObjectValue(sender);
		System.Windows.Controls.MenuItem obj = (System.Windows.Controls.MenuItem)sender;
		obj.RemoveHandler(UIElement.GotFocusEvent, new RoutedEventHandler(OnGotFocusEvent));
		obj.RemoveHandler(UIElement.LostFocusEvent, new RoutedEventHandler(OnLostFocusEvent));
		obj.AddHandler(UIElement.GotFocusEvent, new RoutedEventHandler(OnGotFocusEvent));
		obj.AddHandler(UIElement.LostFocusEvent, new RoutedEventHandler(OnLostFocusEvent));
		_ = null;
	}

	private void SheetMenuOpened(object sender, RoutedEventArgs e)
	{
		System.Windows.Controls.ContextMenu contextMenu = (System.Windows.Controls.ContextMenu)sender;
		A(contextMenu);
		IEnumerator enumerator = default(IEnumerator);
		SheetItem sheetItem;
		try
		{
			enumerator = ((IEnumerable)contextMenu.Items).GetEnumerator();
			while (enumerator.MoveNext())
			{
				object objectValue = RuntimeHelpers.GetObjectValue(enumerator.Current);
				if (!(objectValue is System.Windows.Controls.MenuItem))
				{
					continue;
				}
				while (true)
				{
					switch (5)
					{
					case 0:
						continue;
					}
					break;
				}
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				sheetItem = (SheetItem)contextMenu.DataContext;
				System.Windows.Controls.MenuItem menuItem = (System.Windows.Controls.MenuItem)objectValue;
				string left = menuItem.Header.ToString().ToLower();
				if (Operators.CompareString(left, VH.A(125692), TextCompare: false) != 0)
				{
					while (true)
					{
						switch (7)
						{
						case 0:
							continue;
						}
						break;
					}
					if (Operators.CompareString(left, VH.A(125713), TextCompare: false) != 0)
					{
						while (true)
						{
							switch (4)
							{
							case 0:
								continue;
							}
							break;
						}
						if (Operators.CompareString(left, VH.A(125738), TextCompare: false) == 0)
						{
							menuItem.IsEnabled = Operators.ConditionalCompareObjectNotEqual(NewLateBinding.LateGet(sheetItem.Sheet, null, VH.A(41367), new object[0], null, null, null), XlSheetVisibility.xlSheetVeryHidden, TextCompare: false);
						}
					}
					else
					{
						menuItem.IsEnabled = Operators.ConditionalCompareObjectNotEqual(NewLateBinding.LateGet(sheetItem.Sheet, null, VH.A(41367), new object[0], null, null, null), XlSheetVisibility.xlSheetVisible, TextCompare: false);
					}
				}
				else
				{
					menuItem.IsEnabled = Operators.ConditionalCompareObjectNotEqual(NewLateBinding.LateGet(sheetItem.Sheet, null, VH.A(41367), new object[0], null, null, null), XlSheetVisibility.xlSheetHidden, TextCompare: false);
				}
				A((System.Windows.Controls.MenuItem)objectValue);
				menuItem = null;
			}
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					goto end_IL_01a1;
				}
				continue;
				end_IL_01a1:
				break;
			}
		}
		finally
		{
			if (enumerator is IDisposable)
			{
				while (true)
				{
					switch (4)
					{
					case 0:
						continue;
					}
					(enumerator as IDisposable).Dispose();
					break;
				}
			}
		}
		contextMenu = null;
		sheetItem = null;
	}

	private void NameMenuOpened(object sender, RoutedEventArgs e)
	{
		System.Windows.Controls.ContextMenu contextMenu = (System.Windows.Controls.ContextMenu)sender;
		A(contextMenu);
		IEnumerator enumerator = default(IEnumerator);
		NameItem nameItem;
		try
		{
			enumerator = ((IEnumerable)contextMenu.Items).GetEnumerator();
			while (enumerator.MoveNext())
			{
				object objectValue = RuntimeHelpers.GetObjectValue(enumerator.Current);
				if (!(objectValue is System.Windows.Controls.MenuItem))
				{
					continue;
				}
				while (true)
				{
					switch (5)
					{
					case 0:
						continue;
					}
					break;
				}
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				nameItem = (NameItem)contextMenu.DataContext;
				System.Windows.Controls.MenuItem menuItem = (System.Windows.Controls.MenuItem)objectValue;
				string left = menuItem.Header.ToString().ToLower();
				if (Operators.CompareString(left, VH.A(125759), TextCompare: false) != 0)
				{
					while (true)
					{
						switch (2)
						{
						case 0:
							continue;
						}
						break;
					}
					if (Operators.CompareString(left, VH.A(125778), TextCompare: false) != 0)
					{
						while (true)
						{
							switch (3)
							{
							case 0:
								continue;
							}
							break;
						}
					}
					else
					{
						menuItem.IsEnabled = !nameItem.Name.Visible;
						A((System.Windows.Controls.MenuItem)objectValue);
					}
				}
				else
				{
					menuItem.IsEnabled = nameItem.Name.Visible;
					A((System.Windows.Controls.MenuItem)objectValue);
				}
				menuItem = null;
			}
		}
		finally
		{
			if (enumerator is IDisposable)
			{
				while (true)
				{
					switch (6)
					{
					case 0:
						continue;
					}
					(enumerator as IDisposable).Dispose();
					break;
				}
			}
		}
		contextMenu = null;
		nameItem = null;
	}

	private void A(System.Windows.Controls.MenuItem A)
	{
		SolidColorBrush solidColorBrush = new SolidColorBrush(Colors.Gray);
		try
		{
			if (A.IsEnabled)
			{
				while (true)
				{
					switch (3)
					{
					case 0:
						continue;
					}
					break;
				}
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				solidColorBrush.Opacity = 1.0;
			}
			else
			{
				solidColorBrush.Opacity = Constants.INACTIVE_FILE_OPACITY;
			}
			((GeometryDrawing)((DrawingImage)((System.Windows.Controls.Image)A.Icon).Source).Drawing).Brush = solidColorBrush;
			_ = null;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		solidColorBrush = null;
	}

	private void Unhide(object sender, RoutedEventArgs e)
	{
		A((System.Windows.Controls.MenuItem)sender).F();
	}

	private void Hide(object sender, RoutedEventArgs e)
	{
		A((System.Windows.Controls.MenuItem)sender).G();
	}

	private void Bury(object sender, RoutedEventArgs e)
	{
		A((System.Windows.Controls.MenuItem)sender).H();
	}

	private void InsertSheet(object sender, RoutedEventArgs e)
	{
		A((System.Windows.Controls.MenuItem)sender).C();
	}

	private void DeleteSheet(object sender, RoutedEventArgs e)
	{
		A((System.Windows.Controls.MenuItem)sender).A();
		M();
	}

	private void RenameSheet(object sender, RoutedEventArgs e)
	{
		A((System.Windows.Controls.MenuItem)sender).B();
	}

	private void CopySheet(object sender, RoutedEventArgs e)
	{
		CopiedSheet = A((System.Windows.Controls.MenuItem)sender);
	}

	private void PasteSheet(object sender, RoutedEventArgs e)
	{
		if (CopiedSheet != null)
		{
			Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
			application.ScreenUpdating = false;
			bool copyObjectsWithCells = application.CopyObjectsWithCells;
			application.CopyObjectsWithCells = true;
			RemoveTreeViewSelectionChangedHandler();
			try
			{
				A((System.Windows.Controls.MenuItem)sender).A(CopiedSheet);
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				C(ex2.Message);
				ProjectData.ClearProjectError();
			}
			K();
			application.CopyObjectsWithCells = copyObjectsWithCells;
			application.ScreenUpdating = true;
			application = null;
		}
		else
		{
			D(VH.A(93593));
		}
		trvResults.Focus();
	}

	private void ClearAll(object sender, RoutedEventArgs e)
	{
		A((System.Windows.Controls.MenuItem)sender).D();
	}

	private void ClearFormats(object sender, RoutedEventArgs e)
	{
		A((System.Windows.Controls.MenuItem)sender).F();
	}

	private void ClearContents(object sender, RoutedEventArgs e)
	{
		A((System.Windows.Controls.MenuItem)sender).E();
	}

	private void ClearCommentsNotes(object sender, RoutedEventArgs e)
	{
		A((System.Windows.Controls.MenuItem)sender).G();
	}

	private void ClearHyperlinks(object sender, RoutedEventArgs e)
	{
		A((System.Windows.Controls.MenuItem)sender).H();
	}

	private void ClearPrintAreas(object sender, RoutedEventArgs e)
	{
		A((System.Windows.Controls.MenuItem)sender).I();
	}

	private void ClearShapes(object sender, RoutedEventArgs e)
	{
		A((System.Windows.Controls.MenuItem)sender).J();
	}

	private void ClearCharts(object sender, RoutedEventArgs e)
	{
		A((System.Windows.Controls.MenuItem)sender).K();
	}

	private void ClearSparklines(object sender, RoutedEventArgs e)
	{
		A((System.Windows.Controls.MenuItem)sender).L();
	}

	private void ClearConditionalFormats(object sender, RoutedEventArgs e)
	{
		A((System.Windows.Controls.MenuItem)sender).N();
	}

	private void ClearWatches(object sender, RoutedEventArgs e)
	{
		A((System.Windows.Controls.MenuItem)sender).M();
	}

	private void ViewNormal(object sender, RoutedEventArgs e)
	{
		A((System.Windows.Controls.MenuItem)sender).Q();
	}

	private void ViewPageBreakPreview(object sender, RoutedEventArgs e)
	{
		A((System.Windows.Controls.MenuItem)sender).R();
	}

	private void ScrollAreaLock(object sender, RoutedEventArgs e)
	{
		ScrollArea.Lock();
	}

	private void ScrollAreaUnlock(object sender, RoutedEventArgs e)
	{
		ScrollArea.Unlock();
	}

	private void ProtectSheet(object sender, RoutedEventArgs e)
	{
		A((System.Windows.Controls.MenuItem)sender).O();
	}

	private void UnprotectSheet(object sender, RoutedEventArgs e)
	{
		A((System.Windows.Controls.MenuItem)sender).P();
	}

	private SheetItem A(System.Windows.Controls.MenuItem A)
	{
		return (SheetItem)A.DataContext;
	}

	private WorksheetItem A(System.Windows.Controls.MenuItem A)
	{
		return (WorksheetItem)A.DataContext;
	}

	private void HideName(object sender, RoutedEventArgs e)
	{
		A((System.Windows.Controls.MenuItem)sender).A();
	}

	private void UnhideName(object sender, RoutedEventArgs e)
	{
		A((System.Windows.Controls.MenuItem)sender).B();
	}

	private void DeleteName(object sender, RoutedEventArgs e)
	{
		A((System.Windows.Controls.MenuItem)sender).Delete();
		M();
	}

	private void ApplyName(object sender, RoutedEventArgs e)
	{
		A((System.Windows.Controls.MenuItem)sender).D();
	}

	private void UnapplyName(object sender, RoutedEventArgs e)
	{
		A((System.Windows.Controls.MenuItem)sender).C();
		M();
	}

	private NameItem A(System.Windows.Controls.MenuItem A)
	{
		return (NameItem)A.DataContext;
	}

	private void ConvertToComment(object sender, RoutedEventArgs e)
	{
		bool flag = false;
		NoteItem noteItem = A((System.Windows.Controls.MenuItem)sender);
		NoteItem noteItem2 = noteItem;
		WorksheetItem worksheetItem = (WorksheetItem)noteItem2.Parent;
		Range range = noteItem2.Range;
		try
		{
			noteItem2.C();
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			C(ex2.Message);
			worksheetItem = null;
			noteItem = null;
			range = null;
			ProjectData.ClearProjectError();
			return;
		}
		noteItem2 = null;
		worksheetItem.A(noteItem);
		using (List<BaseQuery>.Enumerator enumerator = SearchQueries.GetEnumerator())
		{
			while (true)
			{
				if (enumerator.MoveNext())
				{
					string uniqueId = enumerator.Current.UniqueId;
					if (Operators.CompareString(uniqueId, VH.A(114992), TextCompare: false) != 0 && Operators.CompareString(uniqueId, VH.A(115014), TextCompare: false) != 0)
					{
						while (true)
						{
							switch (7)
							{
							case 0:
								continue;
							}
							break;
						}
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						if (Operators.CompareString(uniqueId, VH.A(115212), TextCompare: false) != 0 && Operators.CompareString(uniqueId, VH.A(115340), TextCompare: false) != 0)
						{
							while (true)
							{
								switch (4)
								{
								case 0:
									continue;
								}
								break;
							}
							if (Operators.CompareString(uniqueId, VH.A(115425), TextCompare: false) != 0)
							{
								while (true)
								{
									switch (5)
									{
									case 0:
										continue;
									}
									break;
								}
							}
							else
							{
								if (!RangeHelpers.B(range))
								{
									continue;
								}
								while (true)
								{
									switch (3)
									{
									case 0:
										continue;
									}
									flag = true;
									break;
								}
								break;
							}
							continue;
						}
					}
					flag = true;
					break;
				}
				while (true)
				{
					switch (5)
					{
					case 0:
						break;
					default:
						goto end_IL_015a;
					}
					continue;
					end_IL_015a:
					break;
				}
				break;
			}
		}
		if (flag)
		{
			worksheetItem.M(range);
		}
		else if (chkHighlight.IsChecked == true)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					continue;
				}
				break;
			}
			try
			{
				Highlighter.B(range);
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				ProjectData.ClearProjectError();
			}
		}
		worksheetItem = null;
		noteItem = null;
		range = null;
	}

	private void FixNote(object sender, RoutedEventArgs e)
	{
		A((System.Windows.Controls.MenuItem)sender).A();
	}

	private void RemoveAuthor(object sender, RoutedEventArgs e)
	{
		A((System.Windows.Controls.MenuItem)sender).B();
	}

	private void DeleteNote(object sender, RoutedEventArgs e)
	{
		A((System.Windows.Controls.MenuItem)sender).Delete();
		M();
	}

	private NoteItem A(System.Windows.Controls.MenuItem A)
	{
		return (NoteItem)A.DataContext;
	}

	private void DeleteComment(object sender, RoutedEventArgs e)
	{
		A((System.Windows.Controls.MenuItem)sender).Delete();
		M();
	}

	private void ResolveComment(object sender, RoutedEventArgs e)
	{
		A((System.Windows.Controls.MenuItem)sender).A();
	}

	private void UnresolveComment(object sender, RoutedEventArgs e)
	{
		A((System.Windows.Controls.MenuItem)sender).B();
	}

	private CommentItem A(System.Windows.Controls.MenuItem A)
	{
		return (CommentItem)A.DataContext;
	}

	private void RotateRight90(object sender, RoutedEventArgs e)
	{
		ShapeItem shapeItem = A((System.Windows.Controls.MenuItem)sender);
		shapeItem.Shape.IncrementRotation(-90f);
		shapeItem.PreviewImage = null;
		_ = null;
	}

	private void RotateLeft90(object sender, RoutedEventArgs e)
	{
		ShapeItem shapeItem = A((System.Windows.Controls.MenuItem)sender);
		shapeItem.Shape.IncrementRotation(90f);
		shapeItem.PreviewImage = null;
		_ = null;
	}

	private void FlipVertical(object sender, RoutedEventArgs e)
	{
		ShapeItem shapeItem = A((System.Windows.Controls.MenuItem)sender);
		shapeItem.Shape.Flip(MsoFlipCmd.msoFlipVertical);
		shapeItem.PreviewImage = null;
		_ = null;
	}

	private void FlipHorizontal(object sender, RoutedEventArgs e)
	{
		ShapeItem shapeItem = A((System.Windows.Controls.MenuItem)sender);
		shapeItem.Shape.Flip(MsoFlipCmd.msoFlipHorizontal);
		shapeItem.PreviewImage = null;
		_ = null;
	}

	private void BringForward(object sender, RoutedEventArgs e)
	{
		A((System.Windows.Controls.MenuItem)sender).Shape.ZOrder(MsoZOrderCmd.msoBringForward);
	}

	private void BringToFront(object sender, RoutedEventArgs e)
	{
		A((System.Windows.Controls.MenuItem)sender).Shape.ZOrder(MsoZOrderCmd.msoBringToFront);
	}

	private void SendBackward(object sender, RoutedEventArgs e)
	{
		A((System.Windows.Controls.MenuItem)sender).Shape.ZOrder(MsoZOrderCmd.msoSendBackward);
	}

	private void SendToback(object sender, RoutedEventArgs e)
	{
		A((System.Windows.Controls.MenuItem)sender).Shape.ZOrder(MsoZOrderCmd.msoSendToBack);
	}

	private void DeleteShape(object sender, RoutedEventArgs e)
	{
		A((System.Windows.Controls.MenuItem)sender).Delete();
		M();
	}

	private ShapeItem A(System.Windows.Controls.MenuItem A)
	{
		return (ShapeItem)A.DataContext;
	}

	private void FollowHyperlink(object sender, RoutedEventArgs e)
	{
		B((HyperlinkItem)((System.Windows.Controls.MenuItem)sender).DataContext);
	}

	private void HyperlinkClick(object sender, MouseButtonEventArgs e)
	{
		HyperlinkItem hyperlinkItem = (HyperlinkItem)((TextBlock)sender).DataContext;
		if (hyperlinkItem.IsSelected)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					continue;
				}
				break;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			B(hyperlinkItem);
		}
		hyperlinkItem = null;
	}

	private void B(HyperlinkItem A)
	{
		try
		{
			A.Follow();
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			M();
			ProjectData.ClearProjectError();
		}
	}

	private void RemoveHyperlink(object sender, RoutedEventArgs e)
	{
		((HyperlinkItem)((System.Windows.Controls.MenuItem)sender).DataContext).Delete();
		M();
	}

	private void EditHyperlink(object sender, RoutedEventArgs e)
	{
		((HyperlinkItem)((System.Windows.Controls.MenuItem)sender).DataContext).Edit();
	}

	private void SelectChartData(object sender, RoutedEventArgs e)
	{
		A((System.Windows.Controls.MenuItem)sender).SelectData();
	}

	private void ChangeChartType(object sender, RoutedEventArgs e)
	{
		A((System.Windows.Controls.MenuItem)sender).ChangeChartType();
	}

	private void ChartPlotSize(object sender, RoutedEventArgs e)
	{
		ChartAndPlotSize.ShowDialog();
	}

	private void MemorizeChart(object sender, RoutedEventArgs e)
	{
		MemorizeApply.Memorize();
	}

	private void ApplyChartSize(object sender, RoutedEventArgs e)
	{
		MemorizeApply.SetToMemorized(MemorizeApply.MemorizedProperty.ChartSize);
	}

	private void ApplyPlotSize(object sender, RoutedEventArgs e)
	{
		MemorizeApply.SetToMemorized(MemorizeApply.MemorizedProperty.PlotSize);
	}

	private void ApplyPlotPosition(object sender, RoutedEventArgs e)
	{
		MemorizeApply.SetToMemorized(MemorizeApply.MemorizedProperty.PlotPosition);
	}

	private void ApplyAll(object sender, RoutedEventArgs e)
	{
		MemorizeApply.SetToMemorized(MemorizeApply.MemorizedProperty.All);
	}

	private void ResizeToExcel(object sender, RoutedEventArgs e)
	{
		ResizeTo.ExcelSelection();
	}

	private void ResizeToPowerPoint(object sender, RoutedEventArgs e)
	{
		ResizeTo.PowerPointSelection();
	}

	private void ResizeToWord(object sender, RoutedEventArgs e)
	{
		ResizeTo.WordSelection();
	}

	private void RecolorChartDefaults(object sender, RoutedEventArgs e)
	{
		Recolor.SeriesToDefaults();
	}

	private void RecolorChartSource(object sender, RoutedEventArgs e)
	{
		Recolor.PointsToSource();
	}

	private void RecolorChartPoints(object sender, RoutedEventArgs e)
	{
		Recolor.LabelsToPoints();
	}

	private void SaveAsPicture(object sender, RoutedEventArgs e)
	{
		SaveAsImage.Initiate();
	}

	private void DeleteChart(object sender, RoutedEventArgs e)
	{
		A((System.Windows.Controls.MenuItem)sender).Delete();
		M();
	}

	private ChartItem A(System.Windows.Controls.MenuItem A)
	{
		return (ChartItem)A.DataContext;
	}

	private void AutoFilterClear(object sender, RoutedEventArgs e)
	{
		((AutoFilterItem)((System.Windows.Controls.MenuItem)sender).DataContext).A();
	}

	private void AutoFilterReapply(object sender, RoutedEventArgs e)
	{
		((AutoFilterItem)((System.Windows.Controls.MenuItem)sender).DataContext).B();
	}

	private void RemoveValidation(object sender, RoutedEventArgs e)
	{
		((ValidationItem)((System.Windows.Controls.MenuItem)sender).DataContext).Delete();
		M();
	}

	private void ModifyValidation(object sender, RoutedEventArgs e)
	{
		((ValidationItem)((System.Windows.Controls.MenuItem)sender).DataContext).Modify();
	}

	private void RemoveConditionalFormat(object sender, RoutedEventArgs e)
	{
		((ConditionalFormatItem)((System.Windows.Controls.MenuItem)sender).DataContext).Delete();
		M();
	}

	private void ModifyConditionalFormat(object sender, RoutedEventArgs e)
	{
		((ConditionalFormatItem)((System.Windows.Controls.MenuItem)sender).DataContext).Modify();
	}

	private void RemovePrintArea(object sender, RoutedEventArgs e)
	{
		((PrintAreaItem)((System.Windows.Controls.MenuItem)sender).DataContext).Delete();
		M();
	}

	private void ErrorWrap(object sender, RoutedEventArgs e)
	{
		((ErrorItem)((System.Windows.Controls.MenuItem)sender).DataContext).ErrorWrap();
		M();
	}

	private void DeleteWatch(object sender, RoutedEventArgs e)
	{
		((WatchItem)((System.Windows.Controls.MenuItem)sender).DataContext).Delete();
		M();
	}

	private void UnmergeCells(object sender, RoutedEventArgs e)
	{
		((MergedCellsItem)((System.Windows.Controls.MenuItem)sender).DataContext).Delete();
	}

	private void ConvertToRange(object sender, RoutedEventArgs e)
	{
		((TableItem)((System.Windows.Controls.MenuItem)sender).DataContext).Delete();
		M();
	}

	private void RefreshQueryTable(object sender, RoutedEventArgs e)
	{
		((QueryTableItem)((System.Windows.Controls.MenuItem)sender).DataContext).A();
	}

	private void DeleteQuery(object sender, RoutedEventArgs e)
	{
		((QueryTableItem)((System.Windows.Controls.MenuItem)sender).DataContext).Delete();
	}

	private void DeleteSparklines(object sender, RoutedEventArgs e)
	{
		((SparklineItem)((System.Windows.Controls.MenuItem)sender).DataContext).Delete();
		M();
	}

	private void RefreshTable(object sender, RoutedEventArgs e)
	{
		((PivotTableItem)((System.Windows.Controls.MenuItem)sender).DataContext).RefreshTable();
	}

	private void ClearTable(object sender, RoutedEventArgs e)
	{
		((PivotTableItem)((System.Windows.Controls.MenuItem)sender).DataContext).Clear();
	}

	private void CleanUsedRange(object sender, RoutedEventArgs e)
	{
		((UsedRangeItem)((System.Windows.Controls.MenuItem)sender).DataContext).A();
	}

	private void S()
	{
		ccOutdated.Visibility = System.Windows.Visibility.Collapsed;
		ccNoResults.Visibility = System.Windows.Visibility.Collapsed;
		ccCanceled.Visibility = System.Windows.Visibility.Collapsed;
	}

	private void T()
	{
		ccOutdated.Visibility = System.Windows.Visibility.Visible;
	}

	private void C(string A)
	{
		Forms.ErrorMessage(System.Windows.Window.GetWindow(this), A);
	}

	private void D(string A)
	{
		Forms.WarningMessage(System.Windows.Window.GetWindow(this), A);
	}

	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	[DebuggerNonUserCode]
	public void InitializeComponent()
	{
		if (!this.m_D)
		{
			this.m_D = true;
			Uri resourceLocator = new Uri(VH.A(125801), UriKind.Relative);
			System.Windows.Application.LoadComponent(this, resourceLocator);
		}
	}

	void IComponentConnector.InitializeComponent()
	{
		//ILSpy generated this explicit interface implementation from .override directive in InitializeComponent
		this.InitializeComponent();
	}

	[DebuggerNonUserCode]
	[EditorBrowsable(EditorBrowsableState.Never)]
	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	public void System_Windows_Markup_IComponentConnector_Connect(int connectionId, object target)
	{
		if (connectionId == 2)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					btnBack = (System.Windows.Controls.Button)target;
					return;
				}
			}
		}
		if (connectionId == 3)
		{
			btnRefresh = (System.Windows.Controls.Button)target;
			return;
		}
		if (connectionId == 4)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					chkExpandCollapse = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 5)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					chkHighlight = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 6)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					chkPreview = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 7)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					ccOutdated = (ContentControl)target;
					return;
				}
			}
		}
		if (connectionId == 8)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					ccNoResults = (ContentControl)target;
					return;
				}
			}
		}
		if (connectionId == 9)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					ccCanceled = (ContentControl)target;
					return;
				}
			}
		}
		if (connectionId == 10)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					trvResults = (System.Windows.Controls.TreeView)target;
					return;
				}
			}
		}
		if (connectionId == 117)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					grdLoading = (Grid)target;
					return;
				}
			}
		}
		if (connectionId == 118)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					txtLoading = (TextBlock)target;
					return;
				}
			}
		}
		if (connectionId == 119)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					btnCancel = (System.Windows.Controls.Button)target;
					return;
				}
			}
		}
		if (connectionId == 120)
		{
			radCell = (System.Windows.Controls.RadioButton)target;
			return;
		}
		if (connectionId == 121)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					radRow = (System.Windows.Controls.RadioButton)target;
					return;
				}
			}
		}
		switch (connectionId)
		{
		case 122:
			radColumn = (System.Windows.Controls.RadioButton)target;
			break;
		case 123:
			while (true)
			{
				switch (4)
				{
				case 0:
					continue;
				}
				btnSelect = (System.Windows.Controls.Button)target;
				return;
			}
		default:
			this.m_D = true;
			break;
		}
	}

	void IComponentConnector.Connect(int connectionId, object target)
	{
		//ILSpy generated this explicit interface implementation from .override directive in System_Windows_Markup_IComponentConnector_Connect
		this.System_Windows_Markup_IComponentConnector_Connect(connectionId, target);
	}

	[DebuggerNonUserCode]
	[EditorBrowsable(EditorBrowsableState.Never)]
	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	public void System_Windows_Markup_IStyleConnector_Connect(int connectionId, object target)
	{
		if (connectionId == 1)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					continue;
				}
				break;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			EventSetter eventSetter = new EventSetter();
			eventSetter.Event = FrameworkElement.RequestBringIntoViewEvent;
			eventSetter.Handler = new RequestBringIntoViewEventHandler(OnRequestBringIntoView);
			((System.Windows.Style)target).Setters.Add(eventSetter);
			eventSetter = new EventSetter();
			eventSetter.Event = TreeViewItem.SelectedEvent;
			eventSetter.Handler = new RoutedEventHandler(OnSelected);
			((System.Windows.Style)target).Setters.Add(eventSetter);
			eventSetter = new EventSetter();
			eventSetter.Event = TreeViewItem.ExpandedEvent;
			eventSetter.Handler = new RoutedEventHandler(OnExpanded);
			((System.Windows.Style)target).Setters.Add(eventSetter);
			eventSetter = new EventSetter();
			eventSetter.Event = UIElement.KeyDownEvent;
			eventSetter.Handler = new System.Windows.Input.KeyEventHandler(NodeKeyDown);
			((System.Windows.Style)target).Setters.Add(eventSetter);
			eventSetter = new EventSetter();
			eventSetter.Event = UIElement.PreviewMouseDownEvent;
			eventSetter.Handler = new MouseButtonEventHandler(NodePreviewMouseDown);
			((System.Windows.Style)target).Setters.Add(eventSetter);
			eventSetter = new EventSetter();
			eventSetter.Event = UIElement.PreviewMouseRightButtonDownEvent;
			eventSetter.Handler = new MouseButtonEventHandler(NodePreviewMouseRightButtonDown);
			((System.Windows.Style)target).Setters.Add(eventSetter);
			eventSetter = new EventSetter();
			eventSetter.Event = UIElement.MouseEnterEvent;
			eventSetter.Handler = new System.Windows.Input.MouseEventHandler(MouseEnterNode);
			((System.Windows.Style)target).Setters.Add(eventSetter);
			eventSetter = new EventSetter();
			eventSetter.Event = UIElement.MouseLeaveEvent;
			eventSetter.Handler = new System.Windows.Input.MouseEventHandler(MouseLeaveNode);
			((System.Windows.Style)target).Setters.Add(eventSetter);
			eventSetter = new EventSetter();
			eventSetter.Event = UIElement.MouseMoveEvent;
			eventSetter.Handler = new System.Windows.Input.MouseEventHandler(MouseMoveNode);
			((System.Windows.Style)target).Setters.Add(eventSetter);
		}
		if (connectionId == 11)
		{
			((System.Windows.Controls.ContextMenu)target).Opened += SheetMenuOpened;
		}
		if (connectionId == 12)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					continue;
				}
				break;
			}
			((System.Windows.Controls.MenuItem)target).Click += Unhide;
		}
		if (connectionId == 13)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					continue;
				}
				break;
			}
			((System.Windows.Controls.MenuItem)target).Click += Hide;
		}
		if (connectionId == 14)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					continue;
				}
				break;
			}
			((System.Windows.Controls.MenuItem)target).Click += Bury;
		}
		if (connectionId == 15)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					continue;
				}
				break;
			}
			((System.Windows.Controls.MenuItem)target).Click += DeleteSheet;
		}
		if (connectionId == 16)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					continue;
				}
				break;
			}
			((System.Windows.Controls.MenuItem)target).Click += RenameSheet;
		}
		if (connectionId == 17)
		{
			((System.Windows.Controls.MenuItem)target).SubmenuOpened += SubMenuOpened;
		}
		if (connectionId == 18)
		{
			((System.Windows.Controls.MenuItem)target).Click += ClearAll;
		}
		if (connectionId == 19)
		{
			((System.Windows.Controls.MenuItem)target).Click += ClearFormats;
		}
		if (connectionId == 20)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					continue;
				}
				break;
			}
			((System.Windows.Controls.MenuItem)target).Click += ClearContents;
		}
		if (connectionId == 21)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					continue;
				}
				break;
			}
			((System.Windows.Controls.MenuItem)target).Click += ClearCommentsNotes;
		}
		if (connectionId == 22)
		{
			((System.Windows.Controls.MenuItem)target).Click += ClearHyperlinks;
		}
		if (connectionId == 23)
		{
			((System.Windows.Controls.MenuItem)target).Click += ClearPrintAreas;
		}
		if (connectionId == 24)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					continue;
				}
				break;
			}
			((System.Windows.Controls.MenuItem)target).Click += ClearShapes;
		}
		if (connectionId == 25)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					continue;
				}
				break;
			}
			((System.Windows.Controls.MenuItem)target).Click += ClearCharts;
		}
		if (connectionId == 26)
		{
			((System.Windows.Controls.MenuItem)target).Click += ClearSparklines;
		}
		if (connectionId == 27)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					continue;
				}
				break;
			}
			((System.Windows.Controls.MenuItem)target).Click += ClearConditionalFormats;
		}
		if (connectionId == 28)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					continue;
				}
				break;
			}
			((System.Windows.Controls.MenuItem)target).Click += ClearWatches;
		}
		if (connectionId == 29)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					continue;
				}
				break;
			}
			((System.Windows.Controls.MenuItem)target).SubmenuOpened += SubMenuOpened;
		}
		if (connectionId == 30)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					continue;
				}
				break;
			}
			((System.Windows.Controls.MenuItem)target).Click += ViewNormal;
		}
		if (connectionId == 31)
		{
			((System.Windows.Controls.MenuItem)target).Click += ViewPageBreakPreview;
		}
		if (connectionId == 32)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					continue;
				}
				break;
			}
			((System.Windows.Controls.MenuItem)target).SubmenuOpened += SubMenuOpened;
		}
		if (connectionId == 33)
		{
			((System.Windows.Controls.MenuItem)target).Click += ScrollAreaLock;
		}
		if (connectionId == 34)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					continue;
				}
				break;
			}
			((System.Windows.Controls.MenuItem)target).Click += ScrollAreaUnlock;
		}
		if (connectionId == 35)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					continue;
				}
				break;
			}
			((System.Windows.Controls.ContextMenu)target).Opened += MenuOpened;
		}
		if (connectionId == 36)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					continue;
				}
				break;
			}
			((System.Windows.Controls.MenuItem)target).Click += ErrorWrap;
		}
		if (connectionId == 37)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					continue;
				}
				break;
			}
			((System.Windows.Controls.ContextMenu)target).Opened += MenuOpened;
		}
		if (connectionId == 38)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					continue;
				}
				break;
			}
			((System.Windows.Controls.MenuItem)target).Click += ModifyValidation;
		}
		if (connectionId == 39)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					continue;
				}
				break;
			}
			((System.Windows.Controls.MenuItem)target).Click += RemoveValidation;
		}
		if (connectionId == 40)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					continue;
				}
				break;
			}
			((System.Windows.Controls.ContextMenu)target).Opened += MenuOpened;
		}
		if (connectionId == 41)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					continue;
				}
				break;
			}
			((System.Windows.Controls.MenuItem)target).Click += ModifyConditionalFormat;
		}
		if (connectionId == 42)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					continue;
				}
				break;
			}
			((System.Windows.Controls.MenuItem)target).Click += RemoveConditionalFormat;
		}
		if (connectionId == 43)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					continue;
				}
				break;
			}
			((System.Windows.Controls.ContextMenu)target).Opened += MenuOpened;
		}
		if (connectionId == 44)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					continue;
				}
				break;
			}
			((System.Windows.Controls.MenuItem)target).Click += DeleteWatch;
		}
		if (connectionId == 45)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					continue;
				}
				break;
			}
			((System.Windows.Controls.ContextMenu)target).Opened += MenuOpened;
		}
		if (connectionId == 46)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					continue;
				}
				break;
			}
			((System.Windows.Controls.MenuItem)target).Click += DeleteSparklines;
		}
		if (connectionId == 47)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					continue;
				}
				break;
			}
			((System.Windows.Controls.ContextMenu)target).Opened += MenuOpened;
		}
		if (connectionId == 48)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					continue;
				}
				break;
			}
			((System.Windows.Controls.MenuItem)target).Click += AutoFilterClear;
		}
		if (connectionId == 49)
		{
			((System.Windows.Controls.MenuItem)target).Click += AutoFilterReapply;
		}
		if (connectionId == 50)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					continue;
				}
				break;
			}
			((System.Windows.Controls.ContextMenu)target).Opened += MenuOpened;
		}
		if (connectionId == 51)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					continue;
				}
				break;
			}
			((System.Windows.Controls.MenuItem)target).Click += ConvertToRange;
		}
		if (connectionId == 52)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					continue;
				}
				break;
			}
			((System.Windows.Controls.ContextMenu)target).Opened += MenuOpened;
		}
		if (connectionId == 53)
		{
			((System.Windows.Controls.MenuItem)target).Click += RefreshQueryTable;
		}
		if (connectionId == 54)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					continue;
				}
				break;
			}
			((System.Windows.Controls.MenuItem)target).Click += DeleteQuery;
		}
		if (connectionId == 55)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					continue;
				}
				break;
			}
			((System.Windows.Controls.ContextMenu)target).Opened += MenuOpened;
		}
		if (connectionId == 56)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					continue;
				}
				break;
			}
			((System.Windows.Controls.MenuItem)target).Click += RefreshTable;
		}
		if (connectionId == 57)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					continue;
				}
				break;
			}
			((System.Windows.Controls.MenuItem)target).Click += ClearTable;
		}
		if (connectionId == 58)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					continue;
				}
				break;
			}
			((System.Windows.Controls.ContextMenu)target).Opened += MenuOpened;
		}
		if (connectionId == 59)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					continue;
				}
				break;
			}
			((System.Windows.Controls.MenuItem)target).Click += CleanUsedRange;
		}
		if (connectionId == 60)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					continue;
				}
				break;
			}
			((System.Windows.Controls.ContextMenu)target).Opened += MenuOpened;
		}
		if (connectionId == 61)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					continue;
				}
				break;
			}
			((System.Windows.Controls.MenuItem)target).Click += RemovePrintArea;
		}
		if (connectionId == 62)
		{
			((System.Windows.Controls.ContextMenu)target).Opened += MenuOpened;
		}
		if (connectionId == 63)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					continue;
				}
				break;
			}
			((System.Windows.Controls.MenuItem)target).Click += UnmergeCells;
		}
		if (connectionId == 64)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					continue;
				}
				break;
			}
			((System.Windows.Controls.ContextMenu)target).Opened += MenuOpened;
		}
		if (connectionId == 65)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					continue;
				}
				break;
			}
			((System.Windows.Controls.MenuItem)target).Click += ResolveComment;
		}
		if (connectionId == 66)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					continue;
				}
				break;
			}
			((System.Windows.Controls.MenuItem)target).Click += UnresolveComment;
		}
		if (connectionId == 67)
		{
			((System.Windows.Controls.MenuItem)target).Click += DeleteComment;
		}
		if (connectionId == 68)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					continue;
				}
				break;
			}
			((System.Windows.Controls.ContextMenu)target).Opened += MenuOpened;
		}
		if (connectionId == 69)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					continue;
				}
				break;
			}
			((System.Windows.Controls.MenuItem)target).Click += ConvertToComment;
		}
		if (connectionId == 70)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					continue;
				}
				break;
			}
			((System.Windows.Controls.MenuItem)target).Click += FixNote;
		}
		if (connectionId == 71)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					continue;
				}
				break;
			}
			((System.Windows.Controls.MenuItem)target).Click += RemoveAuthor;
		}
		if (connectionId == 72)
		{
			((System.Windows.Controls.MenuItem)target).Click += DeleteNote;
		}
		if (connectionId == 73)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					continue;
				}
				break;
			}
			((System.Windows.Controls.ContextMenu)target).Opened += NameMenuOpened;
		}
		if (connectionId == 74)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					continue;
				}
				break;
			}
			((System.Windows.Controls.MenuItem)target).Click += ApplyName;
		}
		if (connectionId == 75)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					continue;
				}
				break;
			}
			((System.Windows.Controls.MenuItem)target).Click += UnapplyName;
		}
		if (connectionId == 76)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					continue;
				}
				break;
			}
			((System.Windows.Controls.MenuItem)target).Click += HideName;
		}
		if (connectionId == 77)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					continue;
				}
				break;
			}
			((System.Windows.Controls.MenuItem)target).Click += UnhideName;
		}
		if (connectionId == 78)
		{
			((System.Windows.Controls.MenuItem)target).Click += DeleteName;
		}
		if (connectionId == 79)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					continue;
				}
				break;
			}
			((System.Windows.Controls.ContextMenu)target).Opened += MenuOpened;
		}
		if (connectionId == 80)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					continue;
				}
				break;
			}
			((System.Windows.Controls.MenuItem)target).Click += SelectChartData;
		}
		if (connectionId == 81)
		{
			((System.Windows.Controls.MenuItem)target).Click += ChangeChartType;
		}
		if (connectionId == 82)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					continue;
				}
				break;
			}
			((System.Windows.Controls.MenuItem)target).Click += ChartPlotSize;
		}
		if (connectionId == 83)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					continue;
				}
				break;
			}
			((System.Windows.Controls.MenuItem)target).SubmenuOpened += SubMenuOpened;
		}
		if (connectionId == 84)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					continue;
				}
				break;
			}
			((System.Windows.Controls.MenuItem)target).Click += ResizeToExcel;
		}
		if (connectionId == 85)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					continue;
				}
				break;
			}
			((System.Windows.Controls.MenuItem)target).Click += ResizeToPowerPoint;
		}
		if (connectionId == 86)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					continue;
				}
				break;
			}
			((System.Windows.Controls.MenuItem)target).Click += ResizeToWord;
		}
		if (connectionId == 87)
		{
			((System.Windows.Controls.MenuItem)target).Click += MemorizeChart;
		}
		if (connectionId == 88)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					continue;
				}
				break;
			}
			((System.Windows.Controls.MenuItem)target).SubmenuOpened += SubMenuOpened;
		}
		if (connectionId == 89)
		{
			((System.Windows.Controls.MenuItem)target).Click += ApplyChartSize;
		}
		if (connectionId == 90)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					continue;
				}
				break;
			}
			((System.Windows.Controls.MenuItem)target).Click += ApplyPlotSize;
		}
		if (connectionId == 91)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					continue;
				}
				break;
			}
			((System.Windows.Controls.MenuItem)target).Click += ApplyPlotPosition;
		}
		if (connectionId == 92)
		{
			((System.Windows.Controls.MenuItem)target).Click += ApplyAll;
		}
		if (connectionId == 93)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					continue;
				}
				break;
			}
			((System.Windows.Controls.MenuItem)target).SubmenuOpened += SubMenuOpened;
		}
		if (connectionId == 94)
		{
			((System.Windows.Controls.MenuItem)target).Click += RecolorChartDefaults;
		}
		if (connectionId == 95)
		{
			((System.Windows.Controls.MenuItem)target).Click += RecolorChartSource;
		}
		if (connectionId == 96)
		{
			((System.Windows.Controls.MenuItem)target).Click += RecolorChartPoints;
		}
		if (connectionId == 97)
		{
			((System.Windows.Controls.MenuItem)target).Click += SaveAsPicture;
		}
		if (connectionId == 98)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					continue;
				}
				break;
			}
			((System.Windows.Controls.MenuItem)target).Click += DeleteChart;
		}
		if (connectionId == 99)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					continue;
				}
				break;
			}
			((System.Windows.Controls.ContextMenu)target).Opened += MenuOpened;
		}
		if (connectionId == 100)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					continue;
				}
				break;
			}
			((System.Windows.Controls.MenuItem)target).SubmenuOpened += SubMenuOpened;
		}
		if (connectionId == 101)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					continue;
				}
				break;
			}
			((System.Windows.Controls.MenuItem)target).Click += RotateRight90;
		}
		if (connectionId == 102)
		{
			((System.Windows.Controls.MenuItem)target).Click += RotateLeft90;
		}
		if (connectionId == 103)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					continue;
				}
				break;
			}
			((System.Windows.Controls.MenuItem)target).Click += FlipVertical;
		}
		if (connectionId == 104)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					continue;
				}
				break;
			}
			((System.Windows.Controls.MenuItem)target).Click += FlipHorizontal;
		}
		if (connectionId == 105)
		{
			((System.Windows.Controls.MenuItem)target).SubmenuOpened += SubMenuOpened;
		}
		if (connectionId == 106)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					continue;
				}
				break;
			}
			((System.Windows.Controls.MenuItem)target).Click += BringForward;
		}
		if (connectionId == 107)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					continue;
				}
				break;
			}
			((System.Windows.Controls.MenuItem)target).Click += BringToFront;
		}
		if (connectionId == 108)
		{
			((System.Windows.Controls.MenuItem)target).SubmenuOpened += SubMenuOpened;
		}
		if (connectionId == 109)
		{
			((System.Windows.Controls.MenuItem)target).Click += SendBackward;
		}
		if (connectionId == 110)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					continue;
				}
				break;
			}
			((System.Windows.Controls.MenuItem)target).Click += SendToback;
		}
		if (connectionId == 111)
		{
			((System.Windows.Controls.MenuItem)target).Click += DeleteShape;
		}
		if (connectionId == 112)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					continue;
				}
				break;
			}
			((System.Windows.Controls.ContextMenu)target).Opened += MenuOpened;
		}
		if (connectionId == 113)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					continue;
				}
				break;
			}
			((System.Windows.Controls.MenuItem)target).Click += FollowHyperlink;
		}
		if (connectionId == 114)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					continue;
				}
				break;
			}
			((System.Windows.Controls.MenuItem)target).Click += EditHyperlink;
		}
		if (connectionId == 115)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					continue;
				}
				break;
			}
			((System.Windows.Controls.MenuItem)target).Click += RemoveHyperlink;
		}
		if (connectionId != 116)
		{
			return;
		}
		while (true)
		{
			switch (3)
			{
			case 0:
				continue;
			}
			((TextBlock)target).MouseLeftButtonDown += HyperlinkClick;
			return;
		}
	}

	void IStyleConnector.Connect(int connectionId, object target)
	{
		//ILSpy generated this explicit interface implementation from .override directive in System_Windows_Markup_IStyleConnector_Connect
		this.System_Windows_Markup_IStyleConnector_Connect(connectionId, target);
	}

	[SpecialName]
	[CompilerGenerated]
	private void U()
	{
		L();
	}

	[SpecialName]
	[CompilerGenerated]
	private void V()
	{
		K();
	}

	[SpecialName]
	[CompilerGenerated]
	private bool A()
	{
		this.m_A = new DispatcherFrame();
		Dispatcher.PushFrame(this.m_A);
		return false;
	}
}
