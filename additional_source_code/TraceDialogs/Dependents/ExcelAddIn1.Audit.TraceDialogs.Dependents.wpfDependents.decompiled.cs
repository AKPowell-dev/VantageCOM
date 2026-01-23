using System;
using System.CodeDom.Compiler;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Media;
using A;
using ExcelAddIn1.Charts;
using ExcelAddIn1.Formulas;
using MacabacusMacros;
using MacabacusMacros.Auth;
using MacabacusMacros.ExcelHelpers;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Audit.TraceDialogs.Dependents;

[DesignerGenerated]
public sealed class wpfDependents : System.Windows.Controls.UserControl, INotifyPropertyChanged, IComponentConnector, IStyleConnector
{
	private struct QC
	{
		public Range A;

		public string A;
	}

	[Serializable]
	[CompilerGenerated]
	internal sealed class _Closure$__
	{
		public static readonly _Closure$__ A;

		public static Func<QC, int> A;

		public static Func<QC, int> B;

		public static Func<QC, int> C;

		public static Func<QC, int> D;

		public static Func<QC, int> E;

		public static Func<QC, int> F;

		public static Func<QC, int> G;

		public static Func<QC, int> H;

		public static Func<QC, int> I;

		public static Func<QC, int> J;

		public static Func<KeyValuePair<Worksheet, Range>, int> A;

		public static Func<KeyValuePair<Worksheet, Range>, Worksheet> A;

		public static Func<KeyValuePair<Worksheet, Range>, Range> A;

		public static Func<QC, int> K;

		public static Func<QC, int> L;

		public static Func<KeyValuePair<Worksheet, Range>, int> B;

		public static Func<KeyValuePair<Worksheet, Range>, Worksheet> B;

		public static Func<KeyValuePair<Worksheet, Range>, Range> B;

		public static Func<QC, int> M;

		public static Func<QC, int> N;

		static _Closure$__()
		{
			_Closure$__.A = new _Closure$__();
		}

		[SpecialName]
		internal int A(QC A)
		{
			return A.A.Column;
		}

		[SpecialName]
		internal int B(QC A)
		{
			return A.A.Row;
		}

		[SpecialName]
		internal int C(QC A)
		{
			return A.A.Worksheet.Index;
		}

		[SpecialName]
		internal int D(QC A)
		{
			return A.A.Column;
		}

		[SpecialName]
		internal int E(QC A)
		{
			return A.A.Row;
		}

		[SpecialName]
		internal int F(QC A)
		{
			return A.A.Worksheet.Index;
		}

		[SpecialName]
		internal int G(QC A)
		{
			return A.A.Column;
		}

		[SpecialName]
		internal int H(QC A)
		{
			return A.A.Row;
		}

		[SpecialName]
		internal int I(QC A)
		{
			return A.A.Column;
		}

		[SpecialName]
		internal int J(QC A)
		{
			return A.A.Row;
		}

		[SpecialName]
		internal int A(KeyValuePair<Worksheet, Range> A)
		{
			return A.Key.Index;
		}

		[SpecialName]
		internal Worksheet A(KeyValuePair<Worksheet, Range> A)
		{
			return A.Key;
		}

		[SpecialName]
		internal Range A(KeyValuePair<Worksheet, Range> A)
		{
			return A.Value;
		}

		[SpecialName]
		internal int K(QC A)
		{
			return A.A.Column;
		}

		[SpecialName]
		internal int L(QC A)
		{
			return A.A.Row;
		}

		[SpecialName]
		internal int B(KeyValuePair<Worksheet, Range> A)
		{
			return A.Key.Index;
		}

		[SpecialName]
		internal Worksheet B(KeyValuePair<Worksheet, Range> A)
		{
			return A.Key;
		}

		[SpecialName]
		internal Range B(KeyValuePair<Worksheet, Range> A)
		{
			return A.Value;
		}

		[SpecialName]
		internal int M(QC A)
		{
			return A.A.Column;
		}

		[SpecialName]
		internal int N(QC A)
		{
			return A.A.Row;
		}
	}

	[CompilerGenerated]
	internal sealed class RC
	{
		public BaseItem A;

		public List<BaseItem> A;

		public int A;

		public RC(RC A)
		{
			if (A == null)
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				this.A = A.A;
				this.A = A.A;
				this.A = A.A;
				return;
			}
		}

		[SpecialName]
		internal void A()
		{
			this.A.Items.AddRange(this.A.GetRange(99, checked(this.A - 100)));
			JH.A((object)this.A);
			this.A = null;
		}
	}

	[CompilerGenerated]
	private PropertyChangedEventHandler m_A;

	private static FieldInfo m_A;

	private Microsoft.Office.Interop.Excel.Application m_A;

	private Range m_A;

	private List<Worksheet> m_A;

	private Microsoft.Office.Interop.Excel.Workbook m_A;

	private Microsoft.Office.Interop.Excel.Workbook m_B;

	private Worksheet m_A;

	private Range m_B;

	private RoutedPropertyChangedEventHandler<object> m_A;

	private bool m_A;

	private bool m_B;

	private bool m_C;

	private ScrollViewer m_A;

	private Visibility m_A;

	private ObservableCollection<BaseItem> m_A;

	private double m_A;

	[CompilerGenerated]
	private frmDependentsHost m_A;

	private double m_B;

	private bool m_D;

	private ScrollViewer m_B;

	private bool m_E;

	private string m_A;

	private bool m_F;

	[CompilerGenerated]
	[AccessedThroughProperty("ThisWindow")]
	private wpfDependents m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("grdMain")]
	private Grid m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("scroller")]
	private ScrollViewer m_C;

	[AccessedThroughProperty("tbFormula")]
	[CompilerGenerated]
	private TextBlock m_A;

	[AccessedThroughProperty("tbDummy")]
	[CompilerGenerated]
	private TextBlock m_B;

	[AccessedThroughProperty("txtFormula")]
	[CompilerGenerated]
	private System.Windows.Controls.TextBox m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("chkWrap")]
	private System.Windows.Controls.CheckBox m_A;

	[AccessedThroughProperty("masterColumn0")]
	[CompilerGenerated]
	private ColumnDefinition m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("masterColumn2")]
	private ColumnDefinition m_B;

	[AccessedThroughProperty("trvTrace")]
	[CompilerGenerated]
	private System.Windows.Controls.TreeView m_A;

	[AccessedThroughProperty("chkSettings")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox m_B;

	[AccessedThroughProperty("btnOk")]
	[CompilerGenerated]
	private System.Windows.Controls.Button m_A;

	[AccessedThroughProperty("btnCancel")]
	[CompilerGenerated]
	private System.Windows.Controls.Button m_B;

	[AccessedThroughProperty("popSettings")]
	[CompilerGenerated]
	private Popup m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("chkGroup")]
	private System.Windows.Controls.CheckBox m_C;

	[CompilerGenerated]
	[AccessedThroughProperty("chkArrows")]
	private System.Windows.Controls.CheckBox m_D;

	[AccessedThroughProperty("chkUnhide")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox m_E;

	[CompilerGenerated]
	[AccessedThroughProperty("tbHighlight")]
	private TextBlock m_C;

	[AccessedThroughProperty("chkHighlight")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox m_F;

	[AccessedThroughProperty("chkMove")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox m_G;

	private bool m_G;

	public ObservableCollection<BaseItem> RootItems
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(VH.A(42688));
		}
	}

	public double TreeViewViewportWidth
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(VH.A(42707));
		}
	}

	private frmDependentsHost Host
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

	internal virtual wpfDependents ThisWindow
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

	internal virtual Grid grdMain
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

	internal virtual ScrollViewer scroller
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

	internal virtual TextBlock tbFormula
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
			MouseButtonEventHandler value2 = FormulaMouseDown;
			TextBlock textBlock = this.m_A;
			if (textBlock != null)
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
				textBlock.PreviewMouseDown -= value2;
			}
			this.m_A = value;
			textBlock = this.m_A;
			if (textBlock == null)
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
				textBlock.PreviewMouseDown += value2;
				return;
			}
		}
	}

	internal virtual TextBlock tbDummy
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

	internal virtual System.Windows.Controls.TextBox txtFormula
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
			System.Windows.Input.KeyEventHandler value2 = FormulaKeyUp;
			System.Windows.Input.KeyEventHandler value3 = rtbFormula_KeyDown;
			System.Windows.Controls.TextBox textBox = this.m_A;
			if (textBox != null)
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
				textBox.PreviewKeyDown -= value2;
				textBox.KeyDown -= value3;
			}
			this.m_A = value;
			textBox = this.m_A;
			if (textBox != null)
			{
				textBox.PreviewKeyDown += value2;
				textBox.KeyDown += value3;
			}
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkWrap
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

	internal virtual ColumnDefinition masterColumn0
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

	internal virtual ColumnDefinition masterColumn2
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

	internal virtual System.Windows.Controls.TreeView trvTrace
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
			System.Windows.Controls.TreeView treeView = this.m_A;
			if (treeView != null)
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
				treeView.PreviewKeyDown -= value2;
			}
			this.m_A = value;
			treeView = this.m_A;
			if (treeView == null)
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
				treeView.PreviewKeyDown += value2;
				return;
			}
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkSettings
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

	internal virtual System.Windows.Controls.Button btnOk
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
			RoutedEventHandler value2 = btnOk_Click;
			System.Windows.Controls.Button button = this.m_A;
			if (button != null)
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
				button.Click -= value2;
			}
			this.m_A = value;
			button = this.m_A;
			if (button != null)
			{
				button.Click += value2;
			}
		}
	}

	internal virtual System.Windows.Controls.Button btnCancel
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
			RoutedEventHandler value2 = btnCancel_Click;
			System.Windows.Controls.Button button = this.m_B;
			if (button != null)
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
				button.Click -= value2;
			}
			this.m_B = value;
			button = this.m_B;
			if (button == null)
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
				button.Click += value2;
				return;
			}
		}
	}

	internal virtual Popup popSettings
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
			EventHandler value2 = SettingsPopupOpened;
			EventHandler value3 = SettingsPopupClosed;
			System.Windows.Input.KeyEventHandler value4 = CloseSettingsPopup;
			Popup popup = this.m_A;
			if (popup != null)
			{
				popup.Opened -= value2;
				popup.Closed -= value3;
				popup.PreviewKeyDown -= value4;
			}
			this.m_A = value;
			popup = this.m_A;
			if (popup != null)
			{
				popup.Opened += value2;
				popup.Closed += value3;
				popup.PreviewKeyDown += value4;
			}
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkGroup
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

	internal virtual System.Windows.Controls.CheckBox chkArrows
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
			this.m_D = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkUnhide
	{
		[CompilerGenerated]
		get
		{
			return this.m_E;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_E = value;
		}
	}

	internal virtual TextBlock tbHighlight
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

	internal virtual System.Windows.Controls.CheckBox chkHighlight
	{
		[CompilerGenerated]
		get
		{
			return this.m_F;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_F = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkMove
	{
		[CompilerGenerated]
		get
		{
			return this.m_G;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_G = value;
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
				switch (1)
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
				switch (7)
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

	public wpfDependents(frmDependentsHost frm)
	{
		base.Loaded += wpfChild_Loaded;
		base.PreviewKeyDown += wpfChild_KeyUp;
		this.m_A = [SpecialName] (object a0, RoutedPropertyChangedEventArgs<object> a1) =>
		{
			TreeViewSelectionChanged((System.Windows.Controls.TreeView)a0, a1);
		};
		this.m_B = false;
		this.m_C = false;
		this.m_A = null;
		this.m_B = 0.0;
		this.m_D = false;
		this.m_B = null;
		InitializeComponent();
		Host = frm;
		A();
		this.m_A = MH.A.Application;
		this.m_A = new List<Worksheet>();
		try
		{
			this.m_A = this.m_A.ActiveWindow.ActivePane.VisibleRange;
			this.m_B = (Range)this.m_A.Selection;
			this.m_A = this.m_B.Worksheet;
			this.m_B = (Microsoft.Office.Interop.Excel.Workbook)this.m_A.Parent;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			Host.Close();
			ProjectData.ClearProjectError();
		}
		if (Operators.ConditionalCompareObjectGreater(this.m_B.Cells.CountLarge, 5, TextCompare: false))
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
					C(VH.A(43969));
					Close();
					return;
				}
			}
		}
		if (!Access.IsLegacyPlan())
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
			if (!Access.AllowExcelOperation((PlanType)5, (Restriction)0, true))
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
				K.Settings.AuditHighlightCells = false;
				chkHighlight.IsEnabled = false;
				tbHighlight.IsEnabled = false;
			}
		}
		MySettings settings = K.Settings;
		chkWrap.IsChecked = settings.AuditFormulaWrap;
		chkGroup.IsChecked = settings.AuditGroupDependents;
		chkUnhide.IsChecked = settings.AuditUnhideRowsColumns;
		chkHighlight.IsChecked = settings.AuditHighlightCells;
		settings = null;
		chkWrap.Checked += EnableFormulaWrap;
		chkWrap.Unchecked += DisableFormulaWrap;
		if (chkWrap.IsChecked == true)
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
			tbFormula.TextWrapping = TextWrapping.Wrap;
		}
		chkGroup.Checked += GroupingCheckedChanged;
		chkGroup.Unchecked += GroupingCheckedChanged;
		chkHighlight.Checked += HighlightCheckedChanged;
		chkHighlight.Unchecked += HighlightCheckedChanged;
		chkUnhide.Checked += UnhideCheckedChanged;
		chkUnhide.Unchecked += UnhideCheckedChanged;
		txtFormula.PreviewLostKeyboardFocus += FormulaLostKeyboardFocus;
		new ComAwareEventInfo(typeof(AppEvents_Event), VH.A(42971)).AddEventHandler(this.m_A, new AppEvents_WorkbookBeforeCloseEventHandler(A));
	}

	private void A(string A)
	{
		this.m_A?.Invoke(this, new PropertyChangedEventArgs(A));
	}

	private static void A()
	{
		wpfDependents.m_A = typeof(SystemParameters).GetField(VH.A(42651), BindingFlags.Static | BindingFlags.NonPublic);
		B();
		SystemParameters.StaticPropertyChanged += SystemParameters_StaticPropertyChanged;
	}

	private static void SystemParameters_StaticPropertyChanged(object sender, PropertyChangedEventArgs e)
	{
		B();
	}

	private static void B()
	{
		if (!SystemParameters.MenuDropAlignment)
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
			if ((object)wpfDependents.m_A == null)
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
				wpfDependents.m_A.SetValue(null, false);
				return;
			}
		}
	}

	private void wpfChild_Loaded(object sender, RoutedEventArgs e)
	{
		this.m_A = Visibility.Collapsed;
		this.m_A = (ScrollViewer)Forms.GetScrollViewer((DependencyObject)trvTrace);
		TreeViewViewportWidth = this.m_A.ViewportWidth;
		this.m_A.ScrollChanged += TreeViewScrollChanged;
		this.m_A.SizeChanged += TreeViewSizeChanged;
		C();
		Base.CheckForWorkshare();
		trvTrace.Focus();
	}

	private void wpfChild_KeyUp(object sender, System.Windows.Input.KeyEventArgs e)
	{
		System.Windows.Input.KeyEventArgs e2 = e;
		if (e2.Key == Key.Escape)
		{
			if (!popSettings.IsOpen)
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
				if (!txtFormula.IsKeyboardFocused)
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
					this.m_A = true;
					D();
					e2.Handled = true;
				}
			}
		}
		else if (e2.Key == Key.Return && !txtFormula.IsKeyboardFocused)
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
			this.m_A = false;
			Close();
			e2.Handled = true;
		}
		else if (Base.ProcessShortcut(Host, e))
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
			e2.Handled = true;
		}
		else if (e2.Key == Key.F2 && tbFormula.IsEnabled)
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
			G();
			e2.Handled = true;
		}
		else
		{
			if (e2.Key == Key.H)
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
				if (System.Windows.Input.Keyboard.Modifiers == ModifierKeys.Control)
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
					System.Windows.Controls.CheckBox checkBox = chkHighlight;
					bool? isChecked = chkHighlight.IsChecked;
					bool? isChecked2;
					if (!isChecked.HasValue)
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
						isChecked2 = isChecked;
					}
					else
					{
						isChecked2 = isChecked != true;
					}
					checkBox.IsChecked = isChecked2;
					e2.Handled = true;
					goto IL_01ec;
				}
			}
			if (e2.Key == Key.W)
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
				if (System.Windows.Input.Keyboard.Modifiers == ModifierKeys.Control)
				{
					System.Windows.Controls.CheckBox checkBox2 = chkWrap;
					bool? isChecked = chkWrap.IsChecked;
					bool? isChecked3;
					if (!isChecked.HasValue)
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
						isChecked3 = isChecked;
					}
					else
					{
						isChecked3 = isChecked != true;
					}
					checkBox2.IsChecked = isChecked3;
					e2.Handled = true;
				}
			}
		}
		goto IL_01ec;
		IL_01ec:
		e2 = null;
	}

	private void C()
	{
		bool flag = false;
		RemoveTreeViewSelectionChangedHandler();
		RootItems = new ObservableCollection<BaseItem>();
		RootItem rootItem = new RootItem(this.m_B);
		B(rootItem);
		if (rootItem.Items.Count > 0)
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
			rootItem.IsExpanded = true;
			rootItem.IsSelected = true;
			flag = true;
		}
		RootItems.Add(rootItem);
		rootItem = null;
		E();
		if (flag)
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
			if (Operators.ConditionalCompareObjectEqual(this.m_B.Cells.CountLarge, 1, TextCompare: false))
			{
				this.m_A.ScreenUpdating = false;
				try
				{
					Range rng = (Range)this.m_A.Selection;
					List<string> hiddenSheetNames = Base.UnhideHiddenSheets();
					if (Core.HasDependents(this.m_B))
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
						C(VH.A(42750));
					}
					else
					{
						B(VH.A(42829));
					}
					Base.HidePreviouslyHiddenSheets(hiddenSheetNames);
					Base.ReturnToPreviousRange(rng);
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					B(VH.A(42829));
					ProjectData.ClearProjectError();
				}
				finally
				{
				}
				this.m_A.ScreenUpdating = true;
			}
			else
			{
				B(VH.A(42898));
			}
			Close();
			return;
		}
	}

	public void Close()
	{
		int try0000_dispatch = -1;
		int num3 = default(int);
		int num = default(int);
		int num2 = default(int);
		while (true)
		{
			try
			{
				/*Note: ILSpy has introduced the following switch to emulate a goto from catch-block to try-block*/;
				switch (try0000_dispatch)
				{
				default:
					ProjectData.ClearProjectError();
					num3 = 1;
					goto IL_0007;
				case 646:
					{
						num = num2;
						switch (num3)
						{
						case 1:
							break;
						default:
							goto end_IL_0000;
						}
						int num4 = num + 1;
						num = 0;
						switch (num4)
						{
						case 1:
							break;
						case 2:
							goto IL_0007;
						case 3:
							goto IL_0014;
						case 4:
							goto IL_0030;
						case 5:
							goto IL_0037;
						case 6:
							goto IL_003f;
						case 7:
							goto IL_0058;
						case 8:
							goto IL_0071;
						case 9:
							goto IL_008c;
						case 10:
							goto IL_00a8;
						case 11:
							goto IL_00c2;
						case 12:
							goto IL_00de;
						case 13:
							goto IL_00fa;
						case 14:
							goto IL_0114;
						case 15:
							goto IL_0130;
						case 16:
							goto IL_014c;
						case 17:
							goto IL_0160;
						case 18:
							goto IL_0197;
						case 19:
							goto IL_01a1;
						case 20:
							goto IL_01ab;
						case 21:
							goto IL_01b5;
						case 22:
							goto IL_01bf;
						case 23:
							goto IL_01c9;
						case 24:
							goto IL_01d3;
						case 25:
							goto IL_01dd;
						case 26:
							goto IL_01e7;
						case 27:
							goto end_IL_0000_2;
						default:
							goto end_IL_0000;
						case 28:
							goto end_IL_0000_3;
						}
						goto default;
					}
					IL_01e7:
					num2 = 26;
					this.m_A = null;
					break;
					IL_0007:
					num2 = 2;
					this.m_A.ClearArrows();
					goto IL_0014;
					IL_0014:
					num2 = 3;
					Base.HideSheets(this.m_A, this.m_A, !this.m_A);
					goto IL_0030;
					IL_0030:
					num2 = 4;
					Base.RemoveHighlight();
					goto IL_0037;
					IL_0037:
					num2 = 5;
					RemoveTreeViewSelectionChangedHandler();
					goto IL_003f;
					IL_003f:
					num2 = 6;
					this.m_A.ScrollChanged -= TreeViewScrollChanged;
					goto IL_0058;
					IL_0058:
					num2 = 7;
					this.m_A.SizeChanged -= TreeViewSizeChanged;
					goto IL_0071;
					IL_0071:
					num2 = 8;
					chkWrap.Checked -= EnableFormulaWrap;
					goto IL_008c;
					IL_008c:
					num2 = 9;
					chkWrap.Unchecked -= DisableFormulaWrap;
					goto IL_00a8;
					IL_00a8:
					num2 = 10;
					chkGroup.Checked -= GroupingCheckedChanged;
					goto IL_00c2;
					IL_00c2:
					num2 = 11;
					chkGroup.Unchecked -= GroupingCheckedChanged;
					goto IL_00de;
					IL_00de:
					num2 = 12;
					chkHighlight.Checked -= HighlightCheckedChanged;
					goto IL_00fa;
					IL_00fa:
					num2 = 13;
					chkHighlight.Unchecked -= HighlightCheckedChanged;
					goto IL_0114;
					IL_0114:
					num2 = 14;
					chkUnhide.Checked -= UnhideCheckedChanged;
					goto IL_0130;
					IL_0130:
					num2 = 15;
					chkUnhide.Unchecked -= UnhideCheckedChanged;
					goto IL_014c;
					IL_014c:
					num2 = 16;
					SystemParameters.StaticPropertyChanged -= SystemParameters_StaticPropertyChanged;
					goto IL_0160;
					IL_0160:
					num2 = 17;
					new ComAwareEventInfo(typeof(AppEvents_Event), VH.A(42971)).RemoveEventHandler(this.m_A, new AppEvents_WorkbookBeforeCloseEventHandler(A));
					goto IL_0197;
					IL_0197:
					num2 = 18;
					this.m_A = null;
					goto IL_01a1;
					IL_01a1:
					num2 = 19;
					this.m_A = null;
					goto IL_01ab;
					IL_01ab:
					num2 = 20;
					this.m_A = null;
					goto IL_01b5;
					IL_01b5:
					num2 = 21;
					this.m_A = null;
					goto IL_01bf;
					IL_01bf:
					num2 = 22;
					this.m_B = null;
					goto IL_01c9;
					IL_01c9:
					num2 = 23;
					this.m_A = null;
					goto IL_01d3;
					IL_01d3:
					num2 = 24;
					this.m_B = null;
					goto IL_01dd;
					IL_01dd:
					num2 = 25;
					RootItems = null;
					goto IL_01e7;
					end_IL_0000_2:
					break;
				}
				num2 = 27;
				Host.Close();
				break;
				end_IL_0000:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 646;
				continue;
			}
			throw ProjectData.CreateProjectError(-2146828237);
			continue;
			end_IL_0000_3:
			break;
		}
		if (num != 0)
		{
			ProjectData.ClearProjectError();
		}
	}

	private void D()
	{
		this.m_A.ScreenUpdating = false;
		this.m_A.EnableEvents = false;
		Base.CollapseExpandedCells();
		try
		{
			this.m_B.Activate();
			this.m_A.Select(RuntimeHelpers.GetObjectValue(Missing.Value));
			this.m_A.Goto(this.m_B, false);
			this.m_A.Goto(this.m_B, false);
			Base.ScrollTo(this.m_A);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			clsReporting.LogException(ex2);
			ProjectData.ClearProjectError();
		}
		this.m_A.EnableEvents = true;
		this.m_A.ScreenUpdating = true;
		Close();
	}

	private void btnOk_Click(object sender, RoutedEventArgs e)
	{
		this.m_A = false;
		Close();
	}

	private void btnCancel_Click(object sender, RoutedEventArgs e)
	{
		this.m_A = true;
		D();
	}

	private void TreeViewScrollChanged(object sender, ScrollChangedEventArgs e)
	{
		if (this.m_A.ComputedVerticalScrollBarVisibility == this.m_A)
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
			this.m_A = this.m_A.ComputedVerticalScrollBarVisibility;
			TreeViewViewportWidth = this.m_A.ViewportWidth;
			return;
		}
	}

	private void TreeViewSizeChanged(object sender, SizeChangedEventArgs e)
	{
		TreeViewViewportWidth = Base.TreeViewSizeChanged(e, this.m_A);
	}

	private void TreeViewPreviewKeyDown(object sender, System.Windows.Input.KeyEventArgs e)
	{
		Key key = e.Key;
		if ((uint)(key - 19) > 1u)
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
			if (key != Key.Up)
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
				if (key != Key.Down)
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
		}
		if (!e.IsRepeat)
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
			RemoveTreeViewSelectionChangedHandler();
			trvTrace.KeyUp += NavKeyUp;
		}
		BaseItem baseItem = (BaseItem)trvTrace.SelectedItem;
		BaseItem baseItem2 = baseItem;
		checked
		{
			if (e.Key == Key.Down)
			{
				try
				{
					if (baseItem2.IsExpanded)
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
						baseItem2.Items.First().IsSelected = true;
					}
					else
					{
						TraceItem parent = baseItem2.Parent;
						if (baseItem.Index == parent.Items.Count - 1)
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
							parent.Parent.Items[parent.Index + 1].IsSelected = true;
						}
						else
						{
							parent.Items[baseItem.Index + 1].IsSelected = true;
						}
						parent = null;
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
			else if (e.Key == Key.Up)
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
				try
				{
					if (baseItem2.Parent != null)
					{
						if (baseItem2.Index == 0)
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
							baseItem2.Parent.IsSelected = true;
						}
						else
						{
							baseItem2.Parent.Items[baseItem2.Index - 1].IsSelected = true;
						}
					}
					e.Handled = true;
				}
				catch (Exception ex3)
				{
					ProjectData.SetProjectError(ex3);
					Exception ex4 = ex3;
					ProjectData.ClearProjectError();
				}
			}
			baseItem2 = null;
			baseItem = null;
		}
	}

	private void NavKeyUp(object sender, System.Windows.Input.KeyEventArgs e)
	{
		trvTrace.KeyUp -= NavKeyUp;
		E();
		F();
		e.Handled = true;
	}

	private void E()
	{
		trvTrace.SelectedItemChanged += this.m_A;
	}

	public void RemoveTreeViewSelectionChangedHandler()
	{
		trvTrace.SelectedItemChanged -= this.m_A;
	}

	private void TreeViewSelectionChanged(System.Windows.Controls.TreeView sender, RoutedPropertyChangedEventArgs<object> e)
	{
		F();
	}

	private void F()
	{
		BaseItem baseItem = (BaseItem)trvTrace.SelectedItem;
		if (baseItem == null)
		{
			return;
		}
		Range range;
		try
		{
			if (!(baseItem is ChartItem))
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
					range = baseItem.Range;
					Host.IgnoreDeactivate = true;
					try
					{
						A(range);
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						ProjectData.ClearProjectError();
					}
					Host.IgnoreDeactivate = false;
					if (!Host.Focused)
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
						trvTrace.Focus();
					}
					string strFormula;
					if (baseItem.Level > 0 && !(baseItem is MultiCellItem))
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
						tbFormula.IsEnabled = true;
						try
						{
							strFormula = NewLateBinding.LateGet(range, null, VH.A(8714), new object[0], null, null, null).ToString();
						}
						catch (Exception ex3)
						{
							ProjectData.SetProjectError(ex3);
							Exception ex4 = ex3;
							strFormula = range.FormulaLocal.ToString();
							ProjectData.ClearProjectError();
						}
						strFormula = ExcelAddIn1.Formulas.Helpers.RemoveExtraneousSheetName(strFormula, range.Worksheet.Name);
						bool? isChecked = chkWrap.IsChecked;
						bool? flag;
						if (!isChecked.HasValue)
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
							flag = isChecked;
						}
						else
						{
							flag = isChecked != true;
						}
						isChecked = flag;
						if (isChecked == true)
						{
							strFormula = Base.RemoveNewlines(strFormula);
						}
					}
					else
					{
						tbFormula.IsEnabled = false;
						strFormula = "";
					}
					if (strFormula.Length > 0)
					{
						while (true)
						{
							switch (3)
							{
							case 0:
								continue;
							}
							A(baseItem, range, strFormula);
							break;
						}
					}
					else
					{
						tbFormula.Text = "";
					}
					break;
				}
			}
			else
			{
				tbFormula.IsEnabled = false;
				tbFormula.Text = "";
				Charts.GoToChart(((ChartItem)baseItem).Chart, this.m_A);
			}
		}
		catch (Exception ex5)
		{
			ProjectData.SetProjectError(ex5);
			Exception ex6 = ex5;
			ProjectData.ClearProjectError();
		}
		baseItem = null;
		range = null;
	}

	private void A(Range A)
	{
		Base.GoToCell(A, this.m_A, trvTrace, chkUnhide.IsChecked.Value, ref this.m_A);
		try
		{
			Base.MoveFormAsNeeded(Host, A);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	private void A(BaseItem A, Range B, string C)
	{
		string text = "";
		checked
		{
			List<Run> list2;
			MatchCollection matchCollection;
			Run run;
			try
			{
				BaseItem baseItem = (BaseItem)A.Parent;
				if (baseItem is MultiCellItem)
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
					baseItem = (BaseItem)baseItem.Parent;
				}
				try
				{
					text = ((Name)baseItem.Range.Name).Name;
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					ProjectData.ClearProjectError();
				}
				string text2 = Regex.Replace(baseItem.Label, VH.A(43010), "");
				string text3;
				if (B.Worksheet != baseItem.Range.Worksheet)
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
					text3 = VH.A(43025) + baseItem.Range.Worksheet.Name + VH.A(43030);
				}
				else
				{
					text3 = "";
				}
				string text4 = default(string);
				if (Operators.ConditionalCompareObjectEqual(baseItem.Range.Cells.CountLarge, 1, TextCompare: false))
				{
					Match match = Regex.Match(text2, VH.A(43037));
					string value = match.Groups[1].Value;
					string value2 = match.Groups[2].Value;
					_ = null;
					text4 = text3 + VH.A(43088) + value + VH.A(43088) + value2;
				}
				else
				{
					List<string> list = new List<string>();
					Range range;
					try
					{
						range = ((_Worksheet)B.Worksheet).get_Range((object)text2, RuntimeHelpers.GetObjectValue(Missing.Value));
						if (range.Areas.Count == 1)
						{
							Match match2 = Regex.Match(range.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)), VH.A(43095));
							string value = match2.Groups[1].Value;
							string value2 = match2.Groups[2].Value;
							string value3 = match2.Groups[3].Value;
							string value4 = match2.Groups[4].Value;
							_ = null;
							list.Add(text3 + VH.A(43088) + value + VH.A(43088) + value2 + VH.A(43194) + value3 + VH.A(43088) + value4);
						}
						IEnumerator enumerator = range.Cells.GetEnumerator();
						try
						{
							while (enumerator.MoveNext())
							{
								Match match3 = Regex.Match(((Range)enumerator.Current).get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)), VH.A(43037));
								string value = match3.Groups[1].Value;
								string value2 = match3.Groups[2].Value;
								_ = null;
								list.Add(text3 + VH.A(43088) + value + VH.A(43088) + value2);
							}
							while (true)
							{
								switch (7)
								{
								case 0:
									break;
								default:
									goto end_IL_03a1;
								}
								continue;
								end_IL_03a1:
								break;
							}
						}
						finally
						{
							IDisposable disposable = enumerator as IDisposable;
							if (disposable != null)
							{
								disposable.Dispose();
							}
						}
						text4 = Strings.Join(list.ToArray(), VH.A(43203));
					}
					catch (Exception ex3)
					{
						ProjectData.SetProjectError(ex3);
						Exception ex4 = ex3;
						ProjectData.ClearProjectError();
					}
					range = null;
					list = null;
				}
				matchCollection = Regex.Matches(C, VH.A(43206) + text4 + VH.A(43203) + text + VH.A(43243), RegexOptions.None);
				list2 = new List<Run>();
				int num = 0;
				run = null;
				IEnumerator enumerator2 = default(IEnumerator);
				try
				{
					enumerator2 = matchCollection.GetEnumerator();
					while (enumerator2.MoveNext())
					{
						Group obj = ((Match)enumerator2.Current).Groups[1];
						if (obj.Value.Length > 0)
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
							list2.Add(new Run(C.Substring(num, obj.Index - num)));
							Run run2 = new Run(C.Substring(obj.Index, obj.Value.Length));
							Run run3 = run2;
							run3.Background = run3.Foreground.Clone();
							run3.Background.Opacity = 0.12;
							run3.Background.Freeze();
							run3.FontWeight = FontWeights.SemiBold;
							run3.Foreground = Brushes.Black;
							run3.Foreground.Freeze();
							_ = null;
							if (run == null)
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
								run = run2;
							}
							list2.Add(run2);
							run2 = null;
							num = obj.Index + obj.Value.Length;
						}
						obj = null;
					}
					while (true)
					{
						switch (2)
						{
						case 0:
							break;
						default:
							goto end_IL_0582;
						}
						continue;
						end_IL_0582:
						break;
					}
				}
				finally
				{
					if (enumerator2 is IDisposable)
					{
						while (true)
						{
							switch (7)
							{
							case 0:
								continue;
							}
							(enumerator2 as IDisposable).Dispose();
							break;
						}
					}
				}
				if (num < C.Length)
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
					list2.Add(new Run(C.Substring(num, C.Length - num)));
				}
				InlineCollection inlines = tbFormula.Inlines;
				inlines.Clear();
				scroller.ScrollToHorizontalOffset(0.0);
				inlines.AddRange(list2.ToArray());
				if (Operators.CompareString(tbFormula.Text, C, TextCompare: false) != 0)
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
					tbFormula.Text = C;
				}
				else
				{
					bool? isChecked = chkWrap.IsChecked;
					bool? flag;
					if (!isChecked.HasValue)
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
						flag = isChecked;
					}
					else
					{
						flag = isChecked != true;
					}
					bool? flag2 = flag;
					if (flag2.HasValue)
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
						if (flag2 != true)
						{
							goto IL_070f;
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
					if (run != null)
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
						if (flag2.HasValue)
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
							if (Base.ScrollRightRequired(run.ElementEnd, scroller))
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
								Base.ScrollFormula(run.ElementStart, scroller);
							}
						}
					}
				}
				goto IL_070f;
				IL_070f:
				_ = null;
			}
			catch (Exception ex5)
			{
				ProjectData.SetProjectError(ex5);
				Exception ex6 = ex5;
				tbFormula.Text = C;
				ProjectData.ClearProjectError();
			}
			list2 = null;
			matchCollection = null;
			run = null;
		}
	}

	private void OnSelected(object sender, RoutedEventArgs e)
	{
		if (!object.ReferenceEquals(RuntimeHelpers.GetObjectValue(sender), RuntimeHelpers.GetObjectValue(e.OriginalSource)))
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
			if (!(e.OriginalSource is TreeViewItem treeViewItem))
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
				treeViewItem.BringIntoView();
				TreeViewItem treeViewItem2 = null;
				return;
			}
		}
	}

	private void OnExpanded(object sender, RoutedEventArgs e)
	{
		BaseItem baseItem = (BaseItem)((TreeViewItem)sender).DataContext;
		if (baseItem is SingleCellItem || baseItem is RootItem)
		{
			if (baseItem.Items.Count == 1)
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
				if (baseItem.Items[0] is DummyItem)
				{
					B(baseItem);
					baseItem.IsSelected = true;
				}
			}
		}
		else if (baseItem is MultiCellItem)
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
			if (baseItem.Items.Count == 1 && baseItem.Items[0] is DummyItem)
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
				A(baseItem);
				baseItem.IsSelected = true;
			}
		}
		e.Handled = true;
		baseItem = null;
	}

	private void OnCollapsed(object sender, RoutedEventArgs e)
	{
		if (this.m_B)
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
			e.Handled = true;
			return;
		}
	}

	private BaseItem A(BaseItem A)
	{
		while (true)
		{
			if (A is SingleCellItem)
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
				if (A.IsExpanded)
				{
					break;
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
			if (A is RootItem)
			{
				break;
			}
			A = (BaseItem)A.Parent;
		}
		return A;
	}

	private void A(Microsoft.Office.Interop.Excel.Workbook A, ref bool B)
	{
		Close();
	}

	private void EnableFormulaWrap(object sender, RoutedEventArgs e)
	{
		tbFormula.TextWrapping = TextWrapping.Wrap;
		K.Settings.AuditFormulaWrap = true;
	}

	private void DisableFormulaWrap(object sender, RoutedEventArgs e)
	{
		tbFormula.TextWrapping = TextWrapping.NoWrap;
		K.Settings.AuditFormulaWrap = false;
	}

	private void SettingsPopupOpened(object sender, EventArgs e)
	{
		chkSettings.IsHitTestVisible = false;
		chkUnhide.Focus();
	}

	private void SettingsPopupClosed(object sender, EventArgs e)
	{
		chkSettings.IsChecked = false;
		chkSettings.IsHitTestVisible = true;
	}

	private void CloseSettingsPopup(object sender, System.Windows.Input.KeyEventArgs e)
	{
		if (e.Key != Key.Escape)
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
			chkSettings.IsChecked = false;
			e.Handled = true;
			return;
		}
	}

	private void HighlightCheckedChanged(object sender, RoutedEventArgs e)
	{
		K.Settings.AuditHighlightCells = chkHighlight.IsChecked.Value;
		bool? isChecked = chkHighlight.IsChecked;
		if (((!isChecked) ?? isChecked) != true)
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
			Base.RemoveHighlight();
			return;
		}
	}

	private void UnhideCheckedChanged(object sender, RoutedEventArgs e)
	{
		if (chkUnhide.IsChecked == true)
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
			Range range = (Range)this.m_A.Selection;
			if (Ranges.HasHiddenCells(range))
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
				Base.D(range);
			}
			range.Select();
			range = null;
		}
		trvTrace.Focus();
		K.Settings.AuditUnhideRowsColumns = chkUnhide.IsChecked.Value;
	}

	private void GroupingCheckedChanged(object sender, RoutedEventArgs e)
	{
		C();
		K.Settings.AuditGroupDependents = chkGroup.IsChecked.Value;
	}

	private void ArrowsCheckedChanged(object sender, RoutedEventArgs e)
	{
	}

	private void MoveCheckedChanged(object sender, RoutedEventArgs e)
	{
		K.Settings.AuditFormMoveOnNavigate = chkMove.IsChecked.Value;
	}

	private void OnRequestBringIntoView(object sender, RequestBringIntoViewEventArgs e)
	{
		if (this.m_B == null)
		{
			this.m_B = (ScrollViewer)trvTrace.Template.FindName(VH.A(43250), trvTrace);
			if (this.m_B != null)
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
				this.m_B.ScrollChanged -= TreeViewScrollViewerScrollChanged;
				this.m_B.ScrollChanged += TreeViewScrollViewerScrollChanged;
			}
		}
		this.m_D = true;
		this.m_B = this.m_B.HorizontalOffset;
	}

	private void TreeViewScrollViewerScrollChanged(object sender, ScrollChangedEventArgs e)
	{
		if (this.m_D)
		{
			this.m_B.ScrollToHorizontalOffset(this.m_B);
		}
		this.m_D = false;
	}

	private DependencyObject B(DependencyObject A)
	{
		DependencyObject parent = VisualTreeHelper.GetParent(A);
		if (!(parent is System.Windows.Controls.TreeView))
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					return B(parent);
				}
			}
		}
		return parent;
	}

	private void A(BaseItem A)
	{
		int num = 0;
		A.Items.Clear();
		Range rng = (Range)this.m_A.Selection;
		this.m_A.ScreenUpdating = false;
		List<string> hiddenSheetNames = Base.UnhideHiddenSheets();
		BaseItem baseItem;
		try
		{
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = A.Range.GetEnumerator();
				while (enumerator.MoveNext())
				{
					Range range = (Range)enumerator.Current;
					baseItem = new SingleCellItem(A, range, range.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)), num);
					if (Core.HasDependents(range))
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
						baseItem.Items.Add(new DummyItem(baseItem));
					}
					A.Items.Add(baseItem);
					num = checked(num + 1);
				}
				while (true)
				{
					switch (7)
					{
					case 0:
						break;
					default:
						goto end_IL_00dd;
					}
					continue;
					end_IL_00dd:
					break;
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
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		Base.HidePreviouslyHiddenSheets(hiddenSheetNames);
		Base.ReturnToPreviousRange(rng);
		this.m_A.ScreenUpdating = true;
		rng = null;
		baseItem = null;
		hiddenSheetNames = null;
	}

	private void B(BaseItem A)
	{
		RC a = default(RC);
		RC CS$<>8__locals34 = new RC(a);
		CS$<>8__locals34.A = A;
		Range range = null;
		List<string> list = new List<string>();
		CS$<>8__locals34.A = new List<BaseItem>();
		bool flag = false;
		Range range2 = CS$<>8__locals34.A.Range;
		Worksheet worksheet = range2.Worksheet;
		Microsoft.Office.Interop.Excel.Workbook workbook = (Microsoft.Office.Interop.Excel.Workbook)worksheet.Parent;
		Regex c = new Regex(VH.A(43285));
		bool value = chkGroup.IsChecked.Value;
		Microsoft.Office.Interop.Excel.Application a2 = this.m_A;
		XlCalculation calculation = a2.Calculation;
		bool iteration = a2.Iteration;
		a2.Calculation = XlCalculation.xlCalculationManual;
		a2.Iteration = false;
		a2.ScreenUpdating = false;
		_ = null;
		XlDisplayDrawingObjects displayDrawingObjects = workbook.DisplayDrawingObjects;
		workbook.DisplayDrawingObjects = XlDisplayDrawingObjects.xlHide;
		List<QC> list2 = new List<QC>();
		List<QC> list3 = new List<QC>();
		List<QC> list4 = new List<QC>();
		Dictionary<Worksheet, Range> dictionary = default(Dictionary<Worksheet, Range>);
		Dictionary<Worksheet, Range> dictionary2 = default(Dictionary<Worksheet, Range>);
		if (value)
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
			dictionary = new Dictionary<Worksheet, Range>();
			dictionary2 = new Dictionary<Worksheet, Range>();
		}
		List<string> hiddenSheetNames = new List<string>();
		Host.IgnoreDeactivate = true;
		Stopwatch A2 = new Stopwatch();
		A2.Start();
		checked
		{
			Range activeCell;
			Worksheet worksheet2;
			try
			{
				hiddenSheetNames = Base.UnhideHiddenSheets();
				IEnumerator enumerator = default(IEnumerator);
				try
				{
					enumerator = range2.Cells.GetEnumerator();
					while (enumerator.MoveNext())
					{
						Range range3 = (Range)enumerator.Current;
						try
						{
							range3.ShowDependents(RuntimeHelpers.GetObjectValue(Missing.Value));
						}
						catch (Exception ex)
						{
							ProjectData.SetProjectError(ex);
							Exception ex2 = ex;
							if (worksheet.ProtectContents)
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
								System.Windows.Forms.MessageBox.Show(ex2.Message, VH.A(43304), MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
							}
							else
							{
								D(ex2.Message);
								clsReporting.LogException(ex2);
							}
							ProjectData.ClearProjectError();
							break;
						}
						long num = 1L;
						long num2 = 1L;
						bool flag2 = true;
						string right = range3.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value));
						if (!value)
						{
							while (true)
							{
								this.m_A.Goto(range3, RuntimeHelpers.GetObjectValue(Missing.Value));
								try
								{
									this.m_A.ActiveCell.NavigateArrow(false, num, num2);
								}
								catch (Exception ex3)
								{
									ProjectData.SetProjectError(ex3);
									Exception ex4 = ex3;
									ProjectData.ClearProjectError();
									goto IL_04ce;
								}
								activeCell = this.m_A.ActiveCell;
								string text = activeCell.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value));
								if (Operators.CompareString(text, right, TextCompare: false) == 0)
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
									try
									{
										range3.NavigateArrow(false, num + 1, num2);
										if (Operators.CompareString(this.m_A.ActiveCell.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value)), right, TextCompare: false) == 0)
										{
											while (true)
											{
												switch (5)
												{
												case 0:
													break;
												default:
													goto end_IL_031d;
												}
												continue;
												end_IL_031d:
												break;
											}
											goto IL_04ce;
										}
									}
									catch (Exception ex5)
									{
										ProjectData.SetProjectError(ex5);
										Exception ex6 = ex5;
										ProjectData.ClearProjectError();
										goto IL_04ce;
									}
								}
								if (!list.Contains(text))
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
									list.Add(text);
									string text2 = activeCell.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
									worksheet2 = activeCell.Worksheet;
									bool flag3 = worksheet2.Parent == workbook;
									if (flag3 && worksheet2 == worksheet)
									{
										list2.Add(new QC
										{
											A = activeCell,
											A = text2
										});
									}
									else if (flag3)
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
										list3.Add(new QC
										{
											A = activeCell,
											A = this.A(text2, worksheet2, c)
										});
									}
									else
									{
										list4.Add(new QC
										{
											A = activeCell,
											A = Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(VH.A(43335), NewLateBinding.LateGet(worksheet2.Parent, null, VH.A(19019), new object[0], null, null, null)), VH.A(43340)), worksheet2.Name), VH.A(43343)), text2))
										});
									}
								}
								flag2 = false;
								num2++;
								this.A(ref A2);
								continue;
								IL_04ce:
								if (flag2)
								{
									break;
								}
								num2 = 1L;
								flag2 = true;
								num++;
								this.A(ref A2);
							}
							continue;
						}
						while (true)
						{
							this.m_A.Goto(range3, RuntimeHelpers.GetObjectValue(Missing.Value));
							try
							{
								this.m_A.ActiveCell.NavigateArrow(false, num, num2);
							}
							catch (Exception ex7)
							{
								ProjectData.SetProjectError(ex7);
								Exception ex8 = ex7;
								ProjectData.ClearProjectError();
								goto IL_0b51;
							}
							activeCell = this.m_A.ActiveCell;
							string text = activeCell.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value));
							if (Operators.CompareString(text, right, TextCompare: false) == 0)
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
								try
								{
									range3.NavigateArrow(false, num + 1, num2);
									if (Operators.CompareString(this.m_A.ActiveCell.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value)), right, TextCompare: false) == 0)
									{
										goto IL_0b51;
									}
								}
								catch (Exception ex9)
								{
									ProjectData.SetProjectError(ex9);
									Exception ex10 = ex9;
									ProjectData.ClearProjectError();
									goto IL_0b51;
								}
							}
							if (!list.Contains(text))
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
								list.Add(text);
								string text2 = activeCell.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
								worksheet2 = activeCell.Worksheet;
								bool flag3 = worksheet2.Parent == workbook;
								if (flag3)
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
									if (worksheet2 == worksheet)
									{
										if (range != null)
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
											range = this.m_A.Union(range, activeCell, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
										}
										else
										{
											range = activeCell;
										}
										goto IL_0b3a;
									}
								}
								if (flag3)
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
									Dictionary<Worksheet, Range> dictionary3 = dictionary;
									if (dictionary3.ContainsKey(worksheet2))
									{
										dictionary3[worksheet2] = this.m_A.Union(dictionary3[worksheet2], activeCell, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
									}
									else
									{
										dictionary3.Add(worksheet2, activeCell);
									}
									dictionary3 = null;
								}
								else
								{
									Dictionary<Worksheet, Range> dictionary4 = dictionary2;
									if (dictionary4.ContainsKey(worksheet2))
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
										dictionary4[worksheet2] = this.m_A.Union(dictionary4[worksheet2], activeCell, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
									}
									else
									{
										dictionary4.Add(worksheet2, activeCell);
									}
									dictionary4 = null;
								}
							}
							goto IL_0b3a;
							IL_0b51:
							if (flag2)
							{
								break;
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
							num2 = 1L;
							flag2 = true;
							num++;
							this.A(ref A2);
							continue;
							IL_0b3a:
							flag2 = false;
							num2++;
							this.A(ref A2);
						}
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
			}
			catch (TimeoutException ex11)
			{
				ProjectData.SetProjectError(ex11);
				TimeoutException ex12 = ex11;
				C(VH.A(43348));
				ProjectData.ClearProjectError();
			}
			catch (Exception ex13)
			{
				ProjectData.SetProjectError(ex13);
				Exception ex14 = ex13;
				ProjectData.ClearProjectError();
			}
			A2.Stop();
			A2 = null;
			Host.IgnoreDeactivate = false;
			workbook.DisplayDrawingObjects = displayDrawingObjects;
			try
			{
				worksheet.ClearArrows();
			}
			catch (Exception ex15)
			{
				ProjectData.SetProjectError(ex15);
				Exception ex16 = ex15;
				ProjectData.ClearProjectError();
			}
			Base.HidePreviouslyHiddenSheets(hiddenSheetNames);
			this.m_A.ScreenUpdating = true;
			List<QC> list5 = new List<QC>();
			if (!value)
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
				if (list2.Any())
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
					List<QC> list6 = list5;
					IOrderedEnumerable<QC> source = list2.OrderBy([SpecialName] (QC qC3) => qC3.A.Column);
					Func<QC, int> keySelector;
					if (_Closure$__.B == null)
					{
						keySelector = (_Closure$__.B = [SpecialName] (QC qC3) => qC3.A.Row);
					}
					else
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
						keySelector = _Closure$__.B;
					}
					list6.AddRange(source.ThenBy(keySelector));
				}
				if (list3.Any())
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
					List<QC> list7 = list5;
					IOrderedEnumerable<QC> source2 = from qC3 in list3
						orderby qC3.A.Worksheet.Index, qC3.A.Column
						select qC3;
					Func<QC, int> keySelector2;
					if (_Closure$__.E == null)
					{
						keySelector2 = (_Closure$__.E = [SpecialName] (QC qC3) => qC3.A.Row);
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
						keySelector2 = _Closure$__.E;
					}
					list7.AddRange(source2.ThenBy(keySelector2));
				}
				if (list4.Any())
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
					List<QC> list8 = list5;
					List<QC> source3 = list4;
					Func<QC, int> keySelector3;
					if (_Closure$__.F == null)
					{
						keySelector3 = (_Closure$__.F = [SpecialName] (QC qC3) => qC3.A.Worksheet.Index);
					}
					else
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
						keySelector3 = _Closure$__.F;
					}
					IOrderedEnumerable<QC> source4 = source3.OrderBy(keySelector3);
					Func<QC, int> keySelector4;
					if (_Closure$__.G == null)
					{
						keySelector4 = (_Closure$__.G = [SpecialName] (QC qC3) => qC3.A.Column);
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
						keySelector4 = _Closure$__.G;
					}
					IOrderedEnumerable<QC> source5 = source4.ThenBy(keySelector4);
					Func<QC, int> keySelector5;
					if (_Closure$__.H == null)
					{
						keySelector5 = (_Closure$__.H = [SpecialName] (QC qC3) => qC3.A.Row);
					}
					else
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
						keySelector5 = _Closure$__.H;
					}
					list8.AddRange(source5.ThenBy(keySelector5));
				}
			}
			else
			{
				List<QC> list9;
				if (range != null)
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
					list9 = new List<QC>();
					IEnumerator enumerator2 = default(IEnumerator);
					try
					{
						enumerator2 = range.Areas.GetEnumerator();
						while (enumerator2.MoveNext())
						{
							Range range4 = (Range)enumerator2.Current;
							list9.Add(new QC
							{
								A = range4,
								A = range4.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value))
							});
						}
						while (true)
						{
							switch (4)
							{
							case 0:
								break;
							default:
								goto end_IL_0eb5;
							}
							continue;
							end_IL_0eb5:
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
					List<QC> list10 = list5;
					List<QC> source6 = list9;
					Func<QC, int> keySelector6;
					if (_Closure$__.I == null)
					{
						keySelector6 = (_Closure$__.I = [SpecialName] (QC qC3) => qC3.A.Column);
					}
					else
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
						keySelector6 = _Closure$__.I;
					}
					IOrderedEnumerable<QC> source7 = source6.OrderBy(keySelector6);
					Func<QC, int> keySelector7;
					if (_Closure$__.J == null)
					{
						keySelector7 = (_Closure$__.J = [SpecialName] (QC qC3) => qC3.A.Row);
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
						keySelector7 = _Closure$__.J;
					}
					list10.AddRange(source7.ThenBy(keySelector7));
				}
				Dictionary<Worksheet, Range> dictionary5;
				if (dictionary.Any())
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
					Dictionary<Worksheet, Range> source8 = dictionary;
					Func<KeyValuePair<Worksheet, Range>, int> keySelector8;
					if (_Closure$__.A == null)
					{
						keySelector8 = (_Closure$__.A = [SpecialName] (KeyValuePair<Worksheet, Range> keyValuePair) => keyValuePair.Key.Index);
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
						keySelector8 = _Closure$__.A;
					}
					IOrderedEnumerable<KeyValuePair<Worksheet, Range>> source9 = source8.OrderBy(keySelector8);
					Func<KeyValuePair<Worksheet, Range>, Worksheet> keySelector9 = [SpecialName] (KeyValuePair<Worksheet, Range> keyValuePair) => keyValuePair.Key;
					Func<KeyValuePair<Worksheet, Range>, Range> elementSelector;
					if (_Closure$__.A == null)
					{
						elementSelector = (_Closure$__.A = [SpecialName] (KeyValuePair<Worksheet, Range> keyValuePair) => keyValuePair.Value);
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
						elementSelector = _Closure$__.A;
					}
					dictionary5 = source9.ToDictionary(keySelector9, elementSelector);
					IEnumerator enumerator4 = default(IEnumerator);
					foreach (KeyValuePair<Worksheet, Range> item in dictionary5)
					{
						list9 = new List<QC>();
						try
						{
							enumerator4 = item.Value.Areas.GetEnumerator();
							while (enumerator4.MoveNext())
							{
								Range range5 = (Range)enumerator4.Current;
								list9.Add(new QC
								{
									A = range5,
									A = this.A(range5.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)), item.Key, c)
								});
							}
							while (true)
							{
								switch (6)
								{
								case 0:
									break;
								default:
									goto end_IL_10ac;
								}
								continue;
								end_IL_10ac:
								break;
							}
						}
						finally
						{
							if (enumerator4 is IDisposable)
							{
								while (true)
								{
									switch (7)
									{
									case 0:
										continue;
									}
									(enumerator4 as IDisposable).Dispose();
									break;
								}
							}
						}
						List<QC> list11 = list5;
						List<QC> source10 = list9;
						Func<QC, int> keySelector10;
						if (_Closure$__.K != null)
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
							keySelector10 = _Closure$__.K;
						}
						else
						{
							keySelector10 = (_Closure$__.K = [SpecialName] (QC qC3) => qC3.A.Column);
						}
						list11.AddRange(source10.OrderBy(keySelector10).ThenBy((_Closure$__.L == null) ? (_Closure$__.L = [SpecialName] (QC qC3) => qC3.A.Row) : _Closure$__.L));
					}
				}
				if (dictionary2.Any())
				{
					Dictionary<Worksheet, Range> source11 = dictionary2;
					Func<KeyValuePair<Worksheet, Range>, int> keySelector11;
					if (_Closure$__.B == null)
					{
						keySelector11 = (_Closure$__.B = [SpecialName] (KeyValuePair<Worksheet, Range> keyValuePair) => keyValuePair.Key.Index);
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
						keySelector11 = _Closure$__.B;
					}
					IOrderedEnumerable<KeyValuePair<Worksheet, Range>> source12 = source11.OrderBy(keySelector11);
					Func<KeyValuePair<Worksheet, Range>, Worksheet> keySelector12;
					if (_Closure$__.B == null)
					{
						keySelector12 = (_Closure$__.B = [SpecialName] (KeyValuePair<Worksheet, Range> keyValuePair) => keyValuePair.Key);
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
						keySelector12 = _Closure$__.B;
					}
					Func<KeyValuePair<Worksheet, Range>, Range> elementSelector2;
					if (_Closure$__.B == null)
					{
						elementSelector2 = (_Closure$__.B = [SpecialName] (KeyValuePair<Worksheet, Range> keyValuePair) => keyValuePair.Value);
					}
					else
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
						elementSelector2 = _Closure$__.B;
					}
					dictionary5 = source12.ToDictionary(keySelector12, elementSelector2);
					using Dictionary<Worksheet, Range>.Enumerator enumerator5 = dictionary5.GetEnumerator();
					IEnumerator enumerator6 = default(IEnumerator);
					while (enumerator5.MoveNext())
					{
						KeyValuePair<Worksheet, Range> current2 = enumerator5.Current;
						list9 = new List<QC>();
						try
						{
							enumerator6 = current2.Value.Areas.GetEnumerator();
							while (enumerator6.MoveNext())
							{
								Range range6 = (Range)enumerator6.Current;
								list9.Add(new QC
								{
									A = range6,
									A = Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(VH.A(43335), NewLateBinding.LateGet(current2.Key.Parent, null, VH.A(19019), new object[0], null, null, null)), VH.A(43340)), current2.Key.Name), VH.A(43343)), range6.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value))))
								});
							}
						}
						finally
						{
							if (enumerator6 is IDisposable)
							{
								while (true)
								{
									switch (7)
									{
									case 0:
										continue;
									}
									(enumerator6 as IDisposable).Dispose();
									break;
								}
							}
						}
						List<QC> list12 = list5;
						IOrderedEnumerable<QC> source13 = list9.OrderBy([SpecialName] (QC qC3) => qC3.A.Column);
						Func<QC, int> keySelector13;
						if (_Closure$__.N != null)
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
							keySelector13 = _Closure$__.N;
						}
						else
						{
							keySelector13 = (_Closure$__.N = [SpecialName] (QC qC3) => qC3.A.Row);
						}
						list12.AddRange(source13.ThenBy(keySelector13));
					}
					while (true)
					{
						switch (1)
						{
						case 0:
							break;
						default:
							goto end_IL_13c2;
						}
						continue;
						end_IL_13c2:
						break;
					}
				}
				list9 = null;
				dictionary5 = null;
			}
			CS$<>8__locals34.A = list5.Count - 1;
			this.m_A.ScreenUpdating = false;
			BaseItem baseItem;
			Range a3;
			Range rng;
			try
			{
				if (CS$<>8__locals34.A <= 999)
				{
					while (true)
					{
						switch (1)
						{
						case 0:
							continue;
						}
						rng = (Range)this.m_A.Selection;
						bool flag4 = false;
						int num3 = CS$<>8__locals34.A;
						for (int num4 = 0; num4 <= num3; num4++)
						{
							QC qC = list5[num4];
							a3 = qC.A;
							if (Operators.ConditionalCompareObjectEqual(a3.Cells.CountLarge, 1, TextCompare: false))
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
								baseItem = new SingleCellItem(CS$<>8__locals34.A, a3, qC.A, num4);
								flag4 = true;
								if (Core.HasDependents(a3))
								{
									baseItem.Items.Add(new DummyItem(baseItem));
								}
							}
							else
							{
								baseItem = new MultiCellItem(CS$<>8__locals34.A, a3, qC.A, num4);
							}
							qC = default(QC);
							CS$<>8__locals34.A.Add(baseItem);
						}
						while (true)
						{
							switch (3)
							{
							case 0:
								continue;
							}
							if (!flag4)
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
								Base.ReturnToPreviousRange(rng);
								break;
							}
							break;
						}
						break;
					}
				}
				else
				{
					int num5 = CS$<>8__locals34.A;
					for (int num6 = 0; num6 <= num5; num6++)
					{
						QC qC2 = list5[num6];
						if (Operators.ConditionalCompareObjectEqual(qC2.A.Cells.CountLarge, 1, TextCompare: false))
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
							baseItem = new SingleCellItem(CS$<>8__locals34.A, qC2.A, qC2.A, num6);
						}
						else
						{
							baseItem = new MultiCellItem(CS$<>8__locals34.A, qC2.A, qC2.A, num6);
						}
						qC2 = default(QC);
						CS$<>8__locals34.A.Add(baseItem);
					}
					while (true)
					{
						switch (3)
						{
						case 0:
							continue;
						}
						C(VH.A(43451));
						break;
					}
				}
			}
			catch (Exception ex17)
			{
				ProjectData.SetProjectError(ex17);
				Exception ex18 = ex17;
				D(VH.A(43576));
				clsReporting.LogException(ex18);
				flag = true;
				ProjectData.ClearProjectError();
			}
			Microsoft.Office.Interop.Excel.Application a4 = this.m_A;
			a4.ScreenUpdating = true;
			a4.Iteration = iteration;
			a4.Calculation = calculation;
			_ = null;
			CS$<>8__locals34.A.Items.Clear();
			if (!flag)
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
				if (chkArrows.IsChecked == true)
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
					range2.ShowDependents(RuntimeHelpers.GetObjectValue(Missing.Value));
				}
				this.A(range2, ref CS$<>8__locals34.A, ref CS$<>8__locals34.A);
				this.A(range2, CS$<>8__locals34.A, ref CS$<>8__locals34.A);
				CS$<>8__locals34.A = CS$<>8__locals34.A.Count;
				if (CS$<>8__locals34.A <= 100)
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
					CS$<>8__locals34.A.Items.AddRange(CS$<>8__locals34.A);
					JH.A((object)CS$<>8__locals34.A);
					CS$<>8__locals34.A = null;
				}
				else
				{
					CS$<>8__locals34.A.Items.AddRange(CS$<>8__locals34.A.GetRange(0, 100));
					Task.Factory.StartNew([SpecialName] () =>
					{
						CS$<>8__locals34.A.Items.AddRange(CS$<>8__locals34.A.GetRange(99, CS$<>8__locals34.A - 100));
						JH.A((object)CS$<>8__locals34.A);
						CS$<>8__locals34.A = null;
					});
				}
			}
			else
			{
				JH.A((object)CS$<>8__locals34.A);
				CS$<>8__locals34.A = null;
			}
			JH.A((object)list5);
			JH.A((object)list2);
			JH.A((object)list4);
			JH.A((object)list3);
			list5 = null;
			list2 = null;
			list4 = null;
			list3 = null;
			hiddenSheetNames = null;
			baseItem = null;
			a3 = null;
			range2 = null;
			rng = null;
			activeCell = null;
			worksheet = null;
			worksheet2 = null;
			workbook = null;
			list = null;
			c = null;
			dictionary = null;
			dictionary2 = null;
		}
	}

	private void A(ref Stopwatch A)
	{
		if (A.Elapsed.Seconds <= KH.A.DependentsTimeout)
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
			if (Forms.OkCancelMessage(VH.A(43635)) == DialogResult.Cancel)
			{
				throw new TimeoutException();
			}
			A.Restart();
			return;
		}
	}

	private string A(string A, Worksheet B, Regex C)
	{
		string name = B.Name;
		if (C.IsMatch(name))
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					return name + VH.A(7827) + A;
				}
			}
		}
		return VH.A(39851) + name + VH.A(43343) + A;
	}

	private void FormulaMouseDown(object sender, MouseButtonEventArgs e)
	{
		G();
		e.Handled = true;
	}

	private void FormulaKeyUp(object sender, System.Windows.Input.KeyEventArgs e)
	{
		System.Windows.Input.KeyEventArgs e2 = e;
		if (e2.Key == Key.Escape || e2.Key == Key.Return)
		{
			this.m_F = e2.Key == Key.Escape;
			trvTrace.Focus();
			e2.Handled = true;
		}
		e2 = null;
	}

	private void FormulaLostKeyboardFocus(object sender, KeyboardFocusChangedEventArgs e)
	{
		BaseItem baseItem = (BaseItem)trvTrace.SelectedItem;
		checked
		{
			if (!this.m_F && Operators.CompareString(txtFormula.Text, this.m_A, TextCompare: false) != 0)
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
				Range range = baseItem.Range;
				try
				{
					try
					{
						NewLateBinding.LateSet(range, null, VH.A(8714), new object[1] { txtFormula.Text }, null, null);
						txtFormula.Text = NewLateBinding.LateGet(range, null, VH.A(8714), new object[0], null, null, null).ToString();
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						range.FormulaLocal = txtFormula.Text;
						txtFormula.Text = range.FormulaLocal.ToString();
						ProjectData.ClearProjectError();
					}
					baseItem.Value = Base.CleanValueText(Conversions.ToString(range.Text));
					if (range != this.m_B)
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
						try
						{
							if (!A(range))
							{
								while (true)
								{
									switch (7)
									{
									case 0:
										continue;
									}
									baseItem.Parent.Items.Remove(baseItem);
									RangeObservableCollection<TraceItem> items = baseItem.Parent.Items;
									int num = items.Count - 1;
									for (int i = 0; i <= num; i++)
									{
										items[i].Index = i;
									}
									while (true)
									{
										switch (4)
										{
										case 0:
											continue;
										}
										items = null;
										tbFormula.IsEnabled = false;
										tbFormula.Text = "";
										break;
									}
									break;
								}
							}
							else
							{
								A(baseItem, range, txtFormula.Text);
							}
						}
						catch (Exception ex3)
						{
							ProjectData.SetProjectError(ex3);
							Exception ex4 = ex3;
							ProjectData.ClearProjectError();
						}
					}
					H();
				}
				catch (Exception ex5)
				{
					ProjectData.SetProjectError(ex5);
					Exception ex6 = ex5;
					System.Windows.Controls.TextBox textBox = txtFormula;
					textBox.PreviewLostKeyboardFocus -= FormulaLostKeyboardFocus;
					C(VH.A(43769));
					textBox.PreviewLostKeyboardFocus += FormulaLostKeyboardFocus;
					textBox.Text = this.m_A;
					textBox.Focus();
					textBox.SelectAll();
					_ = null;
					e.Handled = true;
					ProjectData.ClearProjectError();
				}
				range = null;
			}
			else
			{
				H();
			}
			baseItem = null;
		}
	}

	private void G()
	{
		System.Windows.Controls.TextBox textBox = txtFormula;
		textBox.Text = tbFormula.Text;
		tbFormula.Visibility = Visibility.Hidden;
		textBox.Visibility = Visibility.Visible;
		textBox.Focus();
		textBox.SelectAll();
		this.m_A = textBox.Text;
		textBox = null;
		this.m_F = false;
	}

	private void H()
	{
		txtFormula.Visibility = Visibility.Hidden;
		tbFormula.Visibility = Visibility.Visible;
	}

	private void rtbFormula_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
	{
		System.Windows.Input.KeyEventArgs e2 = e;
		if (e2.Key != Key.Return)
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
			if (e2.Key != Key.Escape)
			{
				goto IL_007c;
			}
		}
		this.m_F = e2.Key == Key.Escape;
		if (e2.Key == Key.Escape)
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
			txtFormula.Text = this.m_A;
		}
		trvTrace.Focus();
		e2.Handled = true;
		goto IL_007c;
		IL_007c:
		e2 = null;
	}

	private bool A(Range A)
	{
		bool result = false;
		int num = 0;
		BaseItem baseItem = (BaseItem)trvTrace.SelectedItem;
		if (baseItem.Level == 0)
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
					return true;
				}
			}
		}
		Range range = baseItem.Parent.Range;
		try
		{
			num = range.DirectDependents.Count;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		checked
		{
			if (num > 0)
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
				if (range != null)
				{
					XlCalculation calculation = this.m_A.Calculation;
					bool iteration = this.m_A.Iteration;
					this.m_A.Calculation = XlCalculation.xlCalculationManual;
					this.m_A.Iteration = false;
					this.m_A.ScreenUpdating = false;
					this.m_A.EnableEvents = false;
					Microsoft.Office.Interop.Excel.Workbook obj = (Microsoft.Office.Interop.Excel.Workbook)range.Worksheet.Parent;
					XlDisplayDrawingObjects displayDrawingObjects = obj.DisplayDrawingObjects;
					obj.DisplayDrawingObjects = XlDisplayDrawingObjects.xlHide;
					_ = null;
					List<string> hiddenSheetNames = Base.UnhideHiddenSheets();
					try
					{
						range.ShowDependents(RuntimeHelpers.GetObjectValue(Missing.Value));
						long num2 = 1L;
						long num3 = 1L;
						bool flag = true;
						while (true)
						{
							this.m_A.Goto(range, RuntimeHelpers.GetObjectValue(Missing.Value));
							try
							{
								this.m_A.ActiveCell.NavigateArrow(false, num2, num3);
							}
							catch (Exception ex3)
							{
								ProjectData.SetProjectError(ex3);
								Exception ex4 = ex3;
								ProjectData.ClearProjectError();
								goto IL_0256;
							}
							string left = this.m_A.ActiveCell.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value));
							if (Operators.CompareString(left, range.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value)), TextCompare: false) != 0)
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
								if (Operators.CompareString(left, A.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value)), TextCompare: false) != 0)
								{
									flag = false;
									num3++;
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
								result = true;
							}
							goto IL_0256;
							IL_0256:
							if (flag)
							{
								break;
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
							num3 = 1L;
							flag = true;
							num2++;
						}
					}
					catch (Exception ex5)
					{
						ProjectData.SetProjectError(ex5);
						Exception ex6 = ex5;
						ProjectData.ClearProjectError();
					}
					Base.ReturnToPreviousRange(A);
					((Microsoft.Office.Interop.Excel.Workbook)range.Worksheet.Parent).DisplayDrawingObjects = displayDrawingObjects;
					Dialog.ClearArrowsOnActiveSheet(this.m_A);
					Base.HidePreviouslyHiddenSheets(hiddenSheetNames);
					this.m_A.Iteration = iteration;
					this.m_A.Calculation = calculation;
					this.m_A.ScreenUpdating = true;
					this.m_A.EnableEvents = true;
					hiddenSheetNames = null;
				}
			}
			range = null;
			return result;
		}
	}

	private void A(Range A, ref BaseItem B, ref List<BaseItem> C)
	{
		this.m_A.ScreenUpdating = false;
		this.m_A.EnableEvents = false;
		object objectValue = RuntimeHelpers.GetObjectValue(this.m_A.ActiveSheet);
		try
		{
			IEnumerator enumerator = this.m_A.ActiveWorkbook.Sheets.GetEnumerator();
			try
			{
				IEnumerator enumerator2 = default(IEnumerator);
				while (enumerator.MoveNext())
				{
					object objectValue2 = RuntimeHelpers.GetObjectValue(enumerator.Current);
					if (objectValue2 is Worksheet)
					{
						Worksheet worksheet = (Worksheet)objectValue2;
						worksheet.Activate();
						try
						{
							enumerator2 = ((IEnumerable)worksheet.ChartObjects(RuntimeHelpers.GetObjectValue(Missing.Value))).GetEnumerator();
							while (enumerator2.MoveNext())
							{
								ChartObject chartObject = (ChartObject)enumerator2.Current;
								this.A(A, B, chartObject.Chart, ref C);
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
								break;
							}
						}
						finally
						{
							if (enumerator2 is IDisposable)
							{
								while (true)
								{
									switch (6)
									{
									case 0:
										continue;
									}
									(enumerator2 as IDisposable).Dispose();
									break;
								}
							}
						}
					}
					else
					{
						Chart chart = (Chart)objectValue2;
						chart.Activate();
						this.A(A, B, chart, ref C);
					}
				}
				while (true)
				{
					switch (2)
					{
					case 0:
						break;
					default:
						goto end_IL_011e;
					}
					continue;
					end_IL_011e:
					break;
				}
			}
			finally
			{
				IDisposable disposable = enumerator as IDisposable;
				if (disposable != null)
				{
					disposable.Dispose();
				}
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		NewLateBinding.LateCall(objectValue, null, VH.A(39985), new object[0], null, null, null, IgnoreReturn: true);
		objectValue = null;
		this.m_A.EnableEvents = true;
		this.m_A.ScreenUpdating = true;
	}

	private void A(Range A, BaseItem B, Chart C, ref List<BaseItem> D)
	{
		IEnumerator enumerator = default(IEnumerator);
		Range range;
		try
		{
			enumerator = ((IEnumerable)C.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value))).GetEnumerator();
			while (enumerator.MoveNext())
			{
				string[] array = ExcelAddIn1.Charts.Helpers.A((Series)enumerator.Current);
				foreach (string cell in array)
				{
					range = null;
					try
					{
						range = ((_Application)this.m_A).get_Range((object)cell, RuntimeHelpers.GetObjectValue(Missing.Value));
						if (this.m_A.Intersect(range, A, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) != null)
						{
							D.Add(new ChartItem(B, C, C.Name, VH.A(43802), D.Count));
							range = null;
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
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					goto end_IL_021d;
				}
				continue;
				end_IL_021d:
				break;
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
		range = null;
	}

	private void A(Range A, BaseItem B, ref List<BaseItem> C)
	{
		try
		{
			C.Add(new NameItem(B, A, C.Count));
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	private void B(string A)
	{
		Forms.InfoMessage(System.Windows.Window.GetWindow(this), A);
	}

	private void C(string A)
	{
		Forms.WarningMessage(System.Windows.Window.GetWindow(this), A);
	}

	private void D(string A)
	{
		Forms.ErrorMessage(System.Windows.Window.GetWindow(this), A);
	}

	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	[DebuggerNonUserCode]
	public void InitializeComponent()
	{
		if (!this.m_G)
		{
			this.m_G = true;
			Uri resourceLocator = new Uri(VH.A(43817), UriKind.Relative);
			System.Windows.Application.LoadComponent(this, resourceLocator);
		}
	}

	void IComponentConnector.InitializeComponent()
	{
		//ILSpy generated this explicit interface implementation from .override directive in InitializeComponent
		this.InitializeComponent();
	}

	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	[EditorBrowsable(EditorBrowsableState.Never)]
	[DebuggerNonUserCode]
	public void System_Windows_Markup_IComponentConnector_Connect(int connectionId, object target)
	{
		if (connectionId == 1)
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
					ThisWindow = (wpfDependents)target;
					return;
				}
			}
		}
		if (connectionId == 3)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					grdMain = (Grid)target;
					return;
				}
			}
		}
		if (connectionId == 4)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					scroller = (ScrollViewer)target;
					return;
				}
			}
		}
		if (connectionId == 5)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					tbFormula = (TextBlock)target;
					return;
				}
			}
		}
		if (connectionId == 6)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					tbDummy = (TextBlock)target;
					return;
				}
			}
		}
		if (connectionId == 7)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					txtFormula = (System.Windows.Controls.TextBox)target;
					return;
				}
			}
		}
		if (connectionId == 8)
		{
			chkWrap = (System.Windows.Controls.CheckBox)target;
			return;
		}
		if (connectionId == 9)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					masterColumn0 = (ColumnDefinition)target;
					return;
				}
			}
		}
		if (connectionId == 10)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					masterColumn2 = (ColumnDefinition)target;
					return;
				}
			}
		}
		if (connectionId == 11)
		{
			trvTrace = (System.Windows.Controls.TreeView)target;
			return;
		}
		if (connectionId == 12)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					chkSettings = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 13)
		{
			btnOk = (System.Windows.Controls.Button)target;
			return;
		}
		if (connectionId == 14)
		{
			btnCancel = (System.Windows.Controls.Button)target;
			return;
		}
		if (connectionId == 15)
		{
			popSettings = (Popup)target;
			return;
		}
		if (connectionId == 16)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					chkGroup = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 17)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					chkArrows = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 18)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					chkUnhide = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 19)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					tbHighlight = (TextBlock)target;
					return;
				}
			}
		}
		switch (connectionId)
		{
		case 20:
			chkHighlight = (System.Windows.Controls.CheckBox)target;
			break;
		case 21:
			while (true)
			{
				switch (7)
				{
				case 0:
					continue;
				}
				chkMove = (System.Windows.Controls.CheckBox)target;
				return;
			}
		default:
			this.m_G = true;
			break;
		}
	}

	void IComponentConnector.Connect(int connectionId, object target)
	{
		//ILSpy generated this explicit interface implementation from .override directive in System_Windows_Markup_IComponentConnector_Connect
		this.System_Windows_Markup_IComponentConnector_Connect(connectionId, target);
	}

	[EditorBrowsable(EditorBrowsableState.Never)]
	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	[DebuggerNonUserCode]
	public void System_Windows_Markup_IStyleConnector_Connect(int connectionId, object target)
	{
		if (connectionId != 2)
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
			eventSetter.Event = TreeViewItem.CollapsedEvent;
			eventSetter.Handler = new RoutedEventHandler(OnCollapsed);
			((System.Windows.Style)target).Setters.Add(eventSetter);
			return;
		}
	}

	void IStyleConnector.Connect(int connectionId, object target)
	{
		//ILSpy generated this explicit interface implementation from .override directive in System_Windows_Markup_IStyleConnector_Connect
		this.System_Windows_Markup_IStyleConnector_Connect(connectionId, target);
	}
}
