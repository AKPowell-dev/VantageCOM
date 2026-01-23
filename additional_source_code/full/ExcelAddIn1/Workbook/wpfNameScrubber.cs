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
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Threading;
using A;
using ExcelAddIn1.Formulas;
using MacabacusMacros;
using MacabacusMacros.ExcelHelpers;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Workbook;

[DesignerGenerated]
public sealed class wpfNameScrubber : System.Windows.Window, INotifyPropertyChanged, IComponentConnector, IStyleConnector
{
	public enum Filter
	{
		All,
		Visible,
		Hidden,
		Erroneous,
		Linked,
		Unused,
		Lambda,
		ScopeWorkbook,
		ScopeWorksheet
	}

	[Serializable]
	[CompilerGenerated]
	internal sealed class _Closure_0024__
	{
		public static readonly _Closure_0024__ A;

		public static Func<DefinedName, bool> A;

		public static Func<DefinedName, bool> B;

		public static Func<DefinedName, bool> C;

		public static Func<DefinedName, bool> D;

		public static Func<DefinedName, bool> E;

		public static Func<DefinedName, bool> F;

		public static Func<DefinedName, bool> G;

		public static Func<DefinedName, bool> H;

		public static Func<DefinedName, bool> I;

		static _Closure_0024__()
		{
			_Closure_0024__.A = new _Closure_0024__();
		}

		[SpecialName]
		internal bool A(DefinedName A)
		{
			return A.Name.Visible;
		}

		[SpecialName]
		internal bool B(DefinedName A)
		{
			return !A.Name.Visible;
		}

		[SpecialName]
		internal bool C(DefinedName A)
		{
			return ExcelAddIn1.Formulas.Names.A(A.Name);
		}

		[SpecialName]
		internal bool D(DefinedName A)
		{
			return ExcelAddIn1.Formulas.Names.B(A.Name);
		}

		[SpecialName]
		internal bool E(DefinedName A)
		{
			if (!ExcelAddIn1.Formulas.Names.C(A.Name))
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
						return !ExcelAddIn1.Formulas.Names.A(A.Name, B: true);
					}
				}
			}
			return false;
		}

		[SpecialName]
		internal bool F(DefinedName A)
		{
			return ExcelAddIn1.Formulas.Names.C(A.Name);
		}

		[SpecialName]
		internal bool G(DefinedName A)
		{
			return A.Name.Parent is Microsoft.Office.Interop.Excel.Workbook;
		}

		[SpecialName]
		internal bool H(DefinedName A)
		{
			return !(A.Name.Parent is Microsoft.Office.Interop.Excel.Workbook);
		}

		[SpecialName]
		internal bool I(DefinedName A)
		{
			return A.IsChecked;
		}
	}

	[CompilerGenerated]
	internal sealed class PG
	{
		public List<DefinedName> A;

		public wpfNameScrubber A;

		public PG(PG A)
		{
			if (A == null)
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
				this.A = A.A;
				return;
			}
		}

		[SpecialName]
		internal void A()
		{
			this.A.lvNames.SelectionChanged -= this.A.lvNames_SelectionChanged;
			this.A.m_B = new ObservableCollection<DefinedName>(this.A);
			this.A.SourceCollection = CollectionViewSource.GetDefaultView(this.A.m_B);
			this.A.SourceCollection.Filter = this.A.A;
			this.A.F();
			this.A.lvNames.SelectionChanged += this.A.lvNames_SelectionChanged;
		}
	}

	[CompilerGenerated]
	internal sealed class QG
	{
		public DefinedName A;

		public wpfNameScrubber A;

		public Action A;

		public QG(QG A)
		{
			if (A == null)
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
				this.A = A.A;
				return;
			}
		}

		[SpecialName]
		internal void A()
		{
			this.A.m_A.Remove(this.A);
			this.A.m_B.Remove(this.A);
		}
	}

	[CompilerGenerated]
	internal sealed class RG
	{
		public string A;

		public RG(RG A)
		{
			if (A == null)
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
				this.A = A.A;
				return;
			}
		}

		[SpecialName]
		internal bool A(Range A)
		{
			return Operators.CompareString(A.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value)), this.A, TextCompare: false) == 0;
		}
	}

	[CompilerGenerated]
	internal sealed class SG
	{
		public int A;

		public wpfNameScrubber A;

		public SG(SG A)
		{
			if (A == null)
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
				this.A = A.A;
				return;
			}
		}

		[SpecialName]
		internal void A()
		{
			if (this.A.lvNames.SelectedItems.Count == 0)
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
				System.Windows.Controls.ListView lvNames = this.A.lvNames;
				if (((DefinedName)lvNames.SelectedItems[0]).IsChecked)
				{
					this.A = lvNames.Items.IndexOf(RuntimeHelpers.GetObjectValue(lvNames.SelectedItems[0]));
					lvNames.SelectedItems.Clear();
					this.A.A(this.A);
				}
				lvNames = null;
				return;
			}
		}
	}

	[CompilerGenerated]
	internal sealed class TG
	{
		public int A;

		public wpfNameScrubber A;

		[SpecialName]
		internal void A()
		{
			this.A.lblCount.Text = this.A.ToString(this.A.m_C) + VH.A(212129);
		}
	}

	[CompilerGenerated]
	internal sealed class UG
	{
		public string A;

		public wpfNameScrubber A;

		[SpecialName]
		internal void A()
		{
			Forms.InfoMessage(System.Windows.Window.GetWindow(this.A), this.A);
		}
	}

	[CompilerGenerated]
	internal sealed class VG
	{
		public string A;

		public wpfNameScrubber A;

		[SpecialName]
		internal void A()
		{
			Forms.WarningMessage(System.Windows.Window.GetWindow(this.A), this.A);
		}
	}

	[CompilerGenerated]
	internal sealed class WG
	{
		public string A;

		public wpfNameScrubber A;

		[SpecialName]
		internal void A()
		{
			Forms.ErrorMessage(System.Windows.Window.GetWindow(this.A), this.A);
		}
	}

	private Microsoft.Office.Interop.Excel.Application m_A;

	private ObservableCollection<DefinedName> m_A;

	private ObservableCollection<DefinedName> m_B;

	private bool m_A;

	private bool m_B;

	private List<Range> m_A;

	private string m_A;

	private int m_A;

	private BackgroundWorker m_A;

	private BackgroundWorker m_B;

	private readonly string m_B;

	private readonly string m_C;

	private readonly float m_A;

	private readonly Color m_A;

	private readonly Color m_B;

	private ICollectionView m_A;

	[CompilerGenerated]
	private PropertyChangedEventHandler m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("grdMain")]
	private Grid m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("cbxFilter")]
	private System.Windows.Controls.ComboBox m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("cbxSearch")]
	private System.Windows.Controls.ComboBox m_B;

	[CompilerGenerated]
	[AccessedThroughProperty("txtSearch")]
	private System.Windows.Controls.TextBox m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("lvNames")]
	private System.Windows.Controls.ListView m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("gvNames")]
	private GridView m_A;

	[AccessedThroughProperty("chkCheckAll")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox m_A;

	[AccessedThroughProperty("grdLoading")]
	[CompilerGenerated]
	private Grid m_B;

	[AccessedThroughProperty("pbLoading")]
	[CompilerGenerated]
	private System.Windows.Controls.ProgressBar m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("bdrAbort")]
	private System.Windows.Controls.Border m_A;

	[AccessedThroughProperty("stkControls")]
	[CompilerGenerated]
	private StackPanel m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("btnDelete")]
	private System.Windows.Controls.Button m_A;

	[AccessedThroughProperty("btnHide")]
	[CompilerGenerated]
	private System.Windows.Controls.Button m_B;

	[CompilerGenerated]
	[AccessedThroughProperty("btnUnhide")]
	private System.Windows.Controls.Button m_C;

	[CompilerGenerated]
	[AccessedThroughProperty("btnApply")]
	private System.Windows.Controls.Button m_D;

	[CompilerGenerated]
	[AccessedThroughProperty("btnUnapply")]
	private System.Windows.Controls.Button m_E;

	[CompilerGenerated]
	[AccessedThroughProperty("btnEditName")]
	private System.Windows.Controls.Button m_F;

	[AccessedThroughProperty("btnEditRefersTo")]
	[CompilerGenerated]
	private System.Windows.Controls.Button m_G;

	[AccessedThroughProperty("lblCount")]
	[CompilerGenerated]
	private TextBlock m_A;

	[AccessedThroughProperty("lvDependencies")]
	[CompilerGenerated]
	private System.Windows.Controls.ListView m_B;

	[AccessedThroughProperty("gvDependents")]
	[CompilerGenerated]
	private GridView m_B;

	[AccessedThroughProperty("chkShowDependents")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox m_B;

	[CompilerGenerated]
	[AccessedThroughProperty("btnClean")]
	private System.Windows.Controls.Button m_H;

	[AccessedThroughProperty("btnClose")]
	[CompilerGenerated]
	private System.Windows.Controls.Button m_I;

	private bool m_C;

	public ICollectionView SourceCollection
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			D(VH.A(36261));
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

	internal virtual System.Windows.Controls.ComboBox cbxFilter
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

	internal virtual System.Windows.Controls.ComboBox cbxSearch
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

	internal virtual System.Windows.Controls.TextBox txtSearch
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
			RoutedEventHandler value2 = txtSearch_GotFocus;
			RoutedEventHandler value3 = txtSearch_LostFocus;
			System.Windows.Controls.TextBox textBox = this.m_A;
			if (textBox != null)
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
				textBox.GotFocus -= value2;
				textBox.LostFocus -= value3;
			}
			this.m_A = value;
			textBox = this.m_A;
			if (textBox != null)
			{
				textBox.GotFocus += value2;
				textBox.LostFocus += value3;
			}
		}
	}

	internal virtual System.Windows.Controls.ListView lvNames
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
			System.Windows.Input.KeyEventHandler value2 = NamesPreviewKeyDown;
			System.Windows.Input.KeyEventHandler value3 = SpacebarToggle;
			System.Windows.Controls.ListView listView = this.m_A;
			if (listView != null)
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
				listView.PreviewKeyDown -= value2;
				listView.PreviewKeyDown -= value3;
			}
			this.m_A = value;
			listView = this.m_A;
			if (listView != null)
			{
				listView.PreviewKeyDown += value2;
				listView.PreviewKeyDown += value3;
			}
		}
	}

	internal virtual GridView gvNames
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

	internal virtual System.Windows.Controls.CheckBox chkCheckAll
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

	internal virtual Grid grdLoading
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
			DependencyPropertyChangedEventHandler value2 = grdLoading_IsVisibleChanged;
			Grid grid = this.m_B;
			if (grid != null)
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
				grid.IsVisibleChanged -= value2;
			}
			this.m_B = value;
			grid = this.m_B;
			if (grid == null)
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
				grid.IsVisibleChanged += value2;
				return;
			}
		}
	}

	internal virtual System.Windows.Controls.ProgressBar pbLoading
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

	internal virtual System.Windows.Controls.Border bdrAbort
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
			MouseButtonEventHandler value2 = StopLoading;
			System.Windows.Controls.Border border = this.m_A;
			if (border != null)
			{
				border.PreviewMouseDown -= value2;
			}
			this.m_A = value;
			border = this.m_A;
			if (border != null)
			{
				border.PreviewMouseDown += value2;
			}
		}
	}

	internal virtual StackPanel stkControls
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

	internal virtual System.Windows.Controls.Button btnDelete
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
			RoutedEventHandler value2 = StartBackgroundWorker;
			System.Windows.Controls.Button button = this.m_A;
			if (button != null)
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

	internal virtual System.Windows.Controls.Button btnHide
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
			RoutedEventHandler value2 = btnHide_Click;
			System.Windows.Controls.Button button = this.m_B;
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
			this.m_B = value;
			button = this.m_B;
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

	internal virtual System.Windows.Controls.Button btnUnhide
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
			RoutedEventHandler value2 = btnUnhide_Click;
			System.Windows.Controls.Button button = this.m_C;
			if (button != null)
			{
				button.Click -= value2;
			}
			this.m_C = value;
			button = this.m_C;
			if (button != null)
			{
				button.Click += value2;
			}
		}
	}

	internal virtual System.Windows.Controls.Button btnApply
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
			RoutedEventHandler value2 = StartBackgroundWorker;
			System.Windows.Controls.Button button = this.m_D;
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

	internal virtual System.Windows.Controls.Button btnUnapply
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
			RoutedEventHandler value2 = StartBackgroundWorker;
			System.Windows.Controls.Button button = this.m_E;
			if (button != null)
			{
				button.Click -= value2;
			}
			this.m_E = value;
			button = this.m_E;
			if (button == null)
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
				button.Click += value2;
				return;
			}
		}
	}

	internal virtual System.Windows.Controls.Button btnEditName
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
			RoutedEventHandler value2 = btnEditname_Click;
			System.Windows.Controls.Button button = this.m_F;
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
			this.m_F = value;
			button = this.m_F;
			if (button != null)
			{
				button.Click += value2;
			}
		}
	}

	internal virtual System.Windows.Controls.Button btnEditRefersTo
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
			RoutedEventHandler value2 = btnEditRefersTo_Click;
			System.Windows.Controls.Button button = this.m_G;
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
			this.m_G = value;
			button = this.m_G;
			if (button != null)
			{
				button.Click += value2;
			}
		}
	}

	internal virtual TextBlock lblCount
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

	internal virtual System.Windows.Controls.ListView lvDependencies
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
			System.Windows.Input.KeyEventHandler value2 = DependentsPreviewKeyDown;
			System.Windows.Controls.ListView listView = this.m_B;
			if (listView != null)
			{
				listView.PreviewKeyDown -= value2;
			}
			this.m_B = value;
			listView = this.m_B;
			if (listView == null)
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
				listView.PreviewKeyDown += value2;
				return;
			}
		}
	}

	internal virtual GridView gvDependents
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

	internal virtual System.Windows.Controls.CheckBox chkShowDependents
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
			RoutedEventHandler value2 = chkShowDependents_CheckedChanged;
			RoutedEventHandler value3 = chkShowDependents_CheckedChanged;
			System.Windows.Controls.CheckBox checkBox = this.m_B;
			if (checkBox != null)
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
				checkBox.Checked -= value2;
				checkBox.Unchecked -= value3;
			}
			this.m_B = value;
			checkBox = this.m_B;
			if (checkBox != null)
			{
				checkBox.Checked += value2;
				checkBox.Unchecked += value3;
			}
		}
	}

	internal virtual System.Windows.Controls.Button btnClean
	{
		[CompilerGenerated]
		get
		{
			return this.m_H;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			RoutedEventHandler value2 = btnClean_Click;
			System.Windows.Controls.Button button = this.m_H;
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
			this.m_H = value;
			button = this.m_H;
			if (button != null)
			{
				button.Click += value2;
			}
		}
	}

	internal virtual System.Windows.Controls.Button btnClose
	{
		[CompilerGenerated]
		get
		{
			return this.m_I;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			RoutedEventHandler value2 = btnClose_Click;
			System.Windows.Controls.Button button = this.m_I;
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
			this.m_I = value;
			button = this.m_I;
			if (button == null)
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

	public wpfNameScrubber()
	{
		base.Loaded += wpfNameScrubber_Loaded;
		base.Closing += wpfNameScrubber_Closing;
		base.PreviewKeyDown += wpfNameScrubber_PreviewKeyDown;
		this.m_A = null;
		this.m_B = false;
		this.m_A = new List<Range>();
		this.m_A = -1;
		this.m_B = VH.A(102604);
		this.m_C = VH.A(49303) + ((_Application)MH.A.Application).get_International((object)XlApplicationInternational.xlThousandsSeparator).ToString() + VH.A(52500);
		this.m_A = 0.6f;
		this.m_A = Colors.IndianRed;
		this.m_B = Colors.ForestGreen;
		InitializeComponent();
		base.Icon = Forms.GetIcon();
		this.m_A = MH.A.Application;
		this.m_A = new BackgroundWorker();
		BackgroundWorker a = this.m_A;
		a.WorkerSupportsCancellation = true;
		a.WorkerReportsProgress = true;
		a.DoWork += bw_DoWork;
		a.ProgressChanged += ProgressChanged;
		a.RunWorkerCompleted += bw_RunWorkerCompleted;
		_ = null;
		this.m_B = new BackgroundWorker();
		BackgroundWorker b = this.m_B;
		b.WorkerSupportsCancellation = true;
		b.WorkerReportsProgress = true;
		b.DoWork += bgw_DoWork;
		b.ProgressChanged += ProgressChanged;
		b.RunWorkerCompleted += bgw_RunWorkerCompleted;
		_ = null;
	}

	private void wpfNameScrubber_Loaded(object sender, RoutedEventArgs e)
	{
		chkShowDependents.IsChecked = global::A.K.Settings.NameScrubberShowDependents;
		B(chkShowDependents.IsChecked.Value);
		cbxSearch.SelectedIndex = 0;
		cbxFilter.SelectionChanged += cbxFilter_SelectionChanged;
		cbxFilter.SelectedIndex = 0;
		cbxSearch.SelectionChanged += cbxSearch_SelectionChanged;
		lvNames.SelectionChanged += lvNames_SelectionChanged;
		lvDependencies.SelectionChanged += lvDependencies_SelectionChanged;
	}

	private void wpfNameScrubber_Closing(object sender, CancelEventArgs e)
	{
		this.m_A.ScreenUpdating = true;
		this.m_A.EnableEvents = true;
		this.m_A.DisplayAlerts = true;
		A(A: true);
		this.m_A = null;
		this.m_B = null;
		this.m_A = null;
		this.m_A = null;
		this.m_B = null;
		this.m_A = null;
		cbxFilter.SelectionChanged -= cbxFilter_SelectionChanged;
		cbxSearch.SelectionChanged -= cbxSearch_SelectionChanged;
		txtSearch.TextChanged -= SearchTextChanged;
		lvNames.SelectionChanged -= lvNames_SelectionChanged;
		lvDependencies.SelectionChanged -= lvDependencies_SelectionChanged;
		chkCheckAll.Checked -= chkCheckAll_CheckedChanged;
		chkCheckAll.Unchecked -= chkCheckAll_CheckedChanged;
	}

	private void A()
	{
		pbLoading.Value = 0.0;
		grdLoading.Visibility = Visibility.Visible;
	}

	private void ProgressChanged(object sender, ProgressChangedEventArgs e)
	{
		pbLoading.Value = e.ProgressPercentage;
	}

	private void grdLoading_IsVisibleChanged(object sender, DependencyPropertyChangedEventArgs e)
	{
		if (grdLoading.Visibility == Visibility.Visible)
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
					stkControls.Visibility = Visibility.Hidden;
					return;
				}
			}
		}
		stkControls.Visibility = Visibility.Visible;
	}

	private void StopLoading(object sender, MouseButtonEventArgs e)
	{
		A(A: false);
	}

	private void btnClose_Click(object sender, RoutedEventArgs e)
	{
		Close();
	}

	private void A(bool A = false)
	{
		if (this.m_A.IsBusy)
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
					try
					{
						this.m_A.CancelAsync();
						if (A)
						{
							JH.A(this.m_A, 15000, 200);
						}
						return;
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						ProjectData.ClearProjectError();
						return;
					}
				}
			}
		}
		if (!this.m_B.IsBusy)
		{
			return;
		}
		try
		{
			this.m_B.CancelAsync();
			if (A)
			{
				JH.A(this.m_B, 15000, 200);
			}
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			ProjectData.ClearProjectError();
		}
	}

	private void NamesPreviewKeyDown(object sender, System.Windows.Input.KeyEventArgs e)
	{
		Key key = e.Key;
		if (key != Key.Up)
		{
			if (key != Key.Down)
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
				break;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
		}
		if (!e.IsRepeat)
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
			lvNames.SelectionChanged -= lvNames_SelectionChanged;
			lvNames.KeyUp += NamesNavKeyUp;
		}
		if (e.Key == Key.Up)
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
			if (lvNames.SelectedIndex == 0)
			{
				while (true)
				{
					switch (6)
					{
					case 0:
						break;
					default:
						e.Handled = true;
						return;
					}
				}
			}
		}
		if (e.Key == Key.Down && lvNames.SelectedIndex == checked(lvNames.Items.Count - 1))
		{
			e.Handled = true;
		}
	}

	private void NamesNavKeyUp(object sender, System.Windows.Input.KeyEventArgs e)
	{
		lvNames.KeyUp -= NamesNavKeyUp;
		lvNames.SelectionChanged += lvNames_SelectionChanged;
		B();
		e.Handled = true;
	}

	private void lvNames_SelectionChanged(object sender, SelectionChangedEventArgs e)
	{
		if (!lvNames.IsKeyboardFocusWithin)
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
			base.Dispatcher.BeginInvoke((global::A.A)([SpecialName] () =>
			{
				B();
			}), DispatcherPriority.Normal);
			e.Handled = true;
			return;
		}
	}

	private void B()
	{
		if (lvNames.SelectedItems.Count == 1)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
				{
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					Name name;
					Range refersToRange;
					try
					{
						name = ((DefinedName)lvNames.SelectedItems[0]).Name;
						if (!ExcelAddIn1.Formulas.Names.B(name))
						{
							while (true)
							{
								switch (7)
								{
								case 0:
									break;
								default:
									refersToRange = name.RefersToRange;
									if (chkShowDependents.IsChecked == true)
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
										List<Range> list = new List<Range>();
										List<NameDependent> list2 = new List<NameDependent>();
										lvDependencies.SelectionChanged -= lvDependencies_SelectionChanged;
										lvDependencies.ItemsSource = null;
										try
										{
											Microsoft.Office.Interop.Excel.Application application = refersToRange.Application;
											application.ScreenUpdating = false;
											application.EnableEvents = false;
											_ = application.ActiveCell;
											list = ExcelAddIn1.Formulas.Names.GetDependents(refersToRange, name);
											application.ScreenUpdating = true;
											application.EnableEvents = true;
											_ = null;
											string fullName = this.m_A.ActiveWorkbook.FullName;
											using List<Range>.Enumerator enumerator = list.GetEnumerator();
											while (enumerator.MoveNext())
											{
												Range current = enumerator.Current;
												try
												{
													if (ExcelAddIn1.Formulas.Names.A(current, name))
													{
														string strLabel = ((Operators.CompareString(((Microsoft.Office.Interop.Excel.Workbook)NewLateBinding.LateGet(current.Parent, null, VH.A(8701), new object[0], null, null, null)).FullName, fullName, TextCompare: false) == 0) ? Ranges.RangeAddress(current) : current.get_Address((object)false, (object)false, XlReferenceStyle.xlA1, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value)));
														string strFormula;
														try
														{
															strFormula = current.Formula.ToString();
														}
														catch (Exception ex)
														{
															ProjectData.SetProjectError(ex);
															Exception ex2 = ex;
															strFormula = "";
															ProjectData.ClearProjectError();
														}
														list2.Add(new NameDependent(current, strLabel, strFormula));
													}
												}
												catch (Exception ex3)
												{
													ProjectData.SetProjectError(ex3);
													Exception ex4 = ex3;
													ProjectData.ClearProjectError();
												}
											}
											while (true)
											{
												switch (2)
												{
												case 0:
													break;
												default:
													goto end_IL_0208;
												}
												continue;
												end_IL_0208:
												break;
											}
										}
										catch (Exception ex5)
										{
											ProjectData.SetProjectError(ex5);
											Exception ex6 = ex5;
											C(VH.A(176871) + ex6.Message);
											ProjectData.ClearProjectError();
										}
										lvDependencies.ItemsSource = list2;
										Q();
										lvDependencies.SelectionChanged += lvDependencies_SelectionChanged;
										list = null;
										list2 = null;
									}
									A(refersToRange);
									goto end_IL_005e;
								}
								continue;
								end_IL_005e:
								break;
							}
						}
					}
					catch (Exception ex7)
					{
						ProjectData.SetProjectError(ex7);
						Exception ex8 = ex7;
						ProjectData.ClearProjectError();
					}
					name = null;
					refersToRange = null;
					return;
				}
				}
			}
		}
		lvDependencies.ItemsSource = null;
	}

	private void DependentsPreviewKeyDown(object sender, System.Windows.Input.KeyEventArgs e)
	{
		Key key = e.Key;
		if (key != Key.Up)
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
			if (key != Key.Down)
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
		if (!e.IsRepeat)
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
			lvDependencies.SelectionChanged -= lvDependencies_SelectionChanged;
			lvDependencies.KeyUp += DependentsNavKeyUp;
		}
		if (e.Key == Key.Up)
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
			if (lvDependencies.SelectedIndex == 0)
			{
				while (true)
				{
					switch (3)
					{
					case 0:
						break;
					default:
						e.Handled = true;
						return;
					}
				}
			}
		}
		if (e.Key != Key.Down)
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
			if (lvDependencies.SelectedIndex == checked(lvDependencies.Items.Count - 1))
			{
				e.Handled = true;
			}
			return;
		}
	}

	private void DependentsNavKeyUp(object sender, System.Windows.Input.KeyEventArgs e)
	{
		lvDependencies.KeyUp -= DependentsNavKeyUp;
		lvDependencies.SelectionChanged += lvDependencies_SelectionChanged;
		C();
		e.Handled = true;
	}

	private void lvDependencies_SelectionChanged(object sender, SelectionChangedEventArgs e)
	{
		if (!lvDependencies.IsKeyboardFocusWithin)
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
			base.Dispatcher.BeginInvoke((global::A.A)([SpecialName] () =>
			{
				C();
			}), DispatcherPriority.Normal);
			e.Handled = true;
			return;
		}
	}

	private void C()
	{
		A(((NameDependent)lvDependencies.SelectedItems[0]).Range);
	}

	private void A(Range A)
	{
		this.m_A.ScreenUpdating = false;
		this.m_A.EnableEvents = false;
		try
		{
			A.Worksheet.Activate();
			Ranges.ScrollIntoView(A);
			A.Select();
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		this.m_A.ScreenUpdating = true;
		this.m_A.EnableEvents = true;
	}

	private void chkShowDependents_CheckedChanged(object sender, RoutedEventArgs e)
	{
		bool value = chkShowDependents.IsChecked.Value;
		if (value && lvNames.SelectedItems.Count > 0)
		{
			try
			{
				System.Windows.Controls.ListView listView = lvNames;
				int a = listView.Items.IndexOf(RuntimeHelpers.GetObjectValue(listView.SelectedItems[0]));
				listView.SelectedItems.Clear();
				A(a);
				listView = null;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
		}
		B(value);
		global::A.K.Settings.NameScrubberShowDependents = value;
	}

	private void A(int A)
	{
		lvNames.SelectedIndex = A;
		B();
	}

	private void B(bool A)
	{
		if (A)
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
					grdMain.RowDefinitions[2].Height = new GridLength(1.0, GridUnitType.Auto);
					grdMain.RowDefinitions[1].Height = new GridLength(1.0, GridUnitType.Star);
					grdMain.RowDefinitions[3].Height = new GridLength(1.0, GridUnitType.Star);
					return;
				}
			}
		}
		grdMain.RowDefinitions[2].Height = new GridLength(0.0);
		grdMain.RowDefinitions[3].Height = new GridLength(0.0);
	}

	private void cbxFilter_SelectionChanged(object sender, SelectionChangedEventArgs e)
	{
		if (this.m_A.IsBusy)
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
			E();
			if (this.m_A != null)
			{
				while (true)
				{
					switch (2)
					{
					case 0:
						break;
					default:
						lvNames.SelectionChanged -= lvNames_SelectionChanged;
						switch (cbxFilter.SelectedIndex)
						{
						case 0:
							this.m_B = this.m_A;
							break;
						case 1:
						{
							ObservableCollection<DefinedName> a4 = this.m_A;
							Func<DefinedName, bool> predicate4;
							if (_Closure_0024__.A == null)
							{
								predicate4 = (_Closure_0024__.A = [SpecialName] (DefinedName A) => A.Name.Visible);
							}
							else
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
								predicate4 = _Closure_0024__.A;
							}
							this.m_B = new ObservableCollection<DefinedName>(a4.Where(predicate4));
							break;
						}
						case 2:
						{
							ObservableCollection<DefinedName> a3 = this.m_A;
							Func<DefinedName, bool> predicate3;
							if (_Closure_0024__.B == null)
							{
								predicate3 = (_Closure_0024__.B = [SpecialName] (DefinedName A) => !A.Name.Visible);
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
								predicate3 = _Closure_0024__.B;
							}
							this.m_B = new ObservableCollection<DefinedName>(a3.Where(predicate3));
							break;
						}
						case 3:
						{
							ObservableCollection<DefinedName> a7 = this.m_A;
							Func<DefinedName, bool> predicate7;
							if (_Closure_0024__.C == null)
							{
								predicate7 = (_Closure_0024__.C = [SpecialName] (DefinedName A) => ExcelAddIn1.Formulas.Names.A(A.Name));
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
								predicate7 = _Closure_0024__.C;
							}
							this.m_B = new ObservableCollection<DefinedName>(a7.Where(predicate7));
							break;
						}
						case 4:
						{
							ObservableCollection<DefinedName> a5 = this.m_A;
							Func<DefinedName, bool> predicate5;
							if (_Closure_0024__.D == null)
							{
								predicate5 = (_Closure_0024__.D = [SpecialName] (DefinedName A) => ExcelAddIn1.Formulas.Names.B(A.Name));
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
								predicate5 = _Closure_0024__.D;
							}
							this.m_B = new ObservableCollection<DefinedName>(a5.Where(predicate5));
							break;
						}
						case 5:
						{
							this.m_A.ScreenUpdating = false;
							this.m_A.EnableEvents = false;
							Range activeCell = this.m_A.ActiveCell;
							this.m_B = new ObservableCollection<DefinedName>(this.m_A.Where([SpecialName] (DefinedName A) =>
							{
								if (!ExcelAddIn1.Formulas.Names.C(A.Name))
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
											return !ExcelAddIn1.Formulas.Names.A(A.Name, B: true);
										}
									}
								}
								return false;
							}));
							ExcelAddIn1.Formulas.Names.A(activeCell);
							this.m_A.ScreenUpdating = true;
							this.m_A.EnableEvents = true;
							break;
						}
						case 6:
						{
							ObservableCollection<DefinedName> a2 = this.m_A;
							Func<DefinedName, bool> predicate2;
							if (_Closure_0024__.F == null)
							{
								predicate2 = (_Closure_0024__.F = [SpecialName] (DefinedName A) => ExcelAddIn1.Formulas.Names.C(A.Name));
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
								predicate2 = _Closure_0024__.F;
							}
							this.m_B = new ObservableCollection<DefinedName>(a2.Where(predicate2));
							break;
						}
						case 7:
						{
							ObservableCollection<DefinedName> a6 = this.m_A;
							Func<DefinedName, bool> predicate6;
							if (_Closure_0024__.G == null)
							{
								predicate6 = (_Closure_0024__.G = [SpecialName] (DefinedName A) => A.Name.Parent is Microsoft.Office.Interop.Excel.Workbook);
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
								predicate6 = _Closure_0024__.G;
							}
							this.m_B = new ObservableCollection<DefinedName>(a6.Where(predicate6));
							break;
						}
						case 8:
						{
							ObservableCollection<DefinedName> a = this.m_A;
							Func<DefinedName, bool> predicate;
							if (_Closure_0024__.H == null)
							{
								predicate = (_Closure_0024__.H = [SpecialName] (DefinedName A) => !(A.Name.Parent is Microsoft.Office.Interop.Excel.Workbook));
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
								predicate = _Closure_0024__.H;
							}
							this.m_B = new ObservableCollection<DefinedName>(a.Where(predicate));
							break;
						}
						}
						SourceCollection = CollectionViewSource.GetDefaultView(this.m_B);
						SourceCollection.Filter = A;
						lvNames.SelectionChanged += lvNames_SelectionChanged;
						P();
						N();
						O();
						F();
						H();
						return;
					}
				}
			}
			A();
			btnClose.Focus();
			lblCount.Text = VH.A(177005);
			bdrAbort.ToolTip = VH.A(177038);
			G();
			this.m_B.RunWorkerAsync();
			return;
		}
	}

	private void bgw_DoWork(object sender, DoWorkEventArgs e)
	{
		PG a = default(PG);
		PG CS_0024_003C_003E8__locals15 = new PG(a);
		CS_0024_003C_003E8__locals15.A = this;
		CS_0024_003C_003E8__locals15.A = new List<DefinedName>();
		Microsoft.Office.Interop.Excel.Names names = this.m_A.ActiveWorkbook.Names;
		int count = names.Count;
		float num = (float)((double)count / 100.0);
		int num2 = count;
		int num3 = 1;
		checked
		{
			while (true)
			{
				if (num3 <= num2)
				{
					if (this.m_B.CancellationPending)
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
						e.Cancel = true;
						break;
					}
					CS_0024_003C_003E8__locals15.A.Add(A(names.Item(num3, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value))));
					this.m_B.ReportProgress((int)Math.Round((float)num3 / num));
					B(num3);
					num3++;
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
				break;
			}
			base.Dispatcher.Invoke([SpecialName] () =>
			{
				CS_0024_003C_003E8__locals15.A.lvNames.SelectionChanged -= CS_0024_003C_003E8__locals15.A.lvNames_SelectionChanged;
				CS_0024_003C_003E8__locals15.A.m_B = new ObservableCollection<DefinedName>(CS_0024_003C_003E8__locals15.A);
				CS_0024_003C_003E8__locals15.A.SourceCollection = CollectionViewSource.GetDefaultView(CS_0024_003C_003E8__locals15.A.m_B);
				CS_0024_003C_003E8__locals15.A.SourceCollection.Filter = CS_0024_003C_003E8__locals15.A.A;
				CS_0024_003C_003E8__locals15.A.F();
				CS_0024_003C_003E8__locals15.A.lvNames.SelectionChanged += CS_0024_003C_003E8__locals15.A.lvNames_SelectionChanged;
			});
			names = null;
			CS_0024_003C_003E8__locals15.A = null;
		}
	}

	private DefinedName A(Name A)
	{
		Name name = A;
		string strLabel = (name.Name.Contains(VH.A(7827)) ? name.Name.Split('!')[1] : name.Name);
		string strRefersTo;
		try
		{
			strRefersTo = Conversions.ToString(name.RefersTo);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			strRefersTo = "";
			ProjectData.ClearProjectError();
		}
		string strParentName;
		if (name.Parent is Microsoft.Office.Interop.Excel.Workbook)
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
			strParentName = VH.A(175987);
		}
		else
		{
			strParentName = Conversions.ToString(NewLateBinding.LateGet(name.Parent, null, VH.A(19019), new object[0], null, null, null));
		}
		float sngOpacity;
		string strText;
		if (!name.Visible)
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
			sngOpacity = this.m_A;
			strText = VH.A(177077);
		}
		else
		{
			sngOpacity = 1f;
			if (!ExcelAddIn1.Formulas.Names.B(A))
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
					strText = Conversions.ToString(name.RefersToRange.Text);
				}
				catch (Exception ex3)
				{
					ProjectData.SetProjectError(ex3);
					Exception ex4 = ex3;
					strText = VH.A(177077);
					ProjectData.ClearProjectError();
				}
			}
			else
			{
				strText = VH.A(177077);
			}
		}
		name = null;
		return new DefinedName(A, A.Index, blnChecked: false, strLabel, strText, strParentName, strRefersTo, sngOpacity);
	}

	private void bgw_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
	{
		if (e.Error != null)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					{
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						if (e.Error is InvalidCastException)
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
							if (e.Error.Message.Contains(VH.A(177084)))
							{
								C(VH.A(177131) + e.Error.Message);
							}
							else
							{
								C(e.Error.Message);
							}
							clsReporting.LogException(e.Error);
						}
						else if (e.Error is OutOfMemoryException)
						{
							C(e.Error.Message + VH.A(177395) + this.m_A.ActiveWorkbook.Names.Count.ToString(this.m_C) + VH.A(177765));
						}
						else if (this.m_A != null)
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
							int count = this.m_A.ActiveWorkbook.Names.Count;
							if (e.Error.HResult == -2147024882)
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
								if (count > 0)
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
									C(e.Error.Message + VH.A(177780) + count.ToString(this.m_C) + VH.A(177765));
									goto IL_01c3;
								}
							}
							C(e.Error.Message);
							goto IL_01c3;
						}
						goto IL_01d0;
					}
					IL_01c3:
					clsReporting.LogException(e.Error);
					goto IL_01d0;
					IL_01d0:
					Close();
					return;
				}
			}
		}
		N();
		O();
		H();
		grdLoading.Visibility = Visibility.Hidden;
		lvNames.Focus();
		if (cbxFilter.SelectedIndex != 0)
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
			if (this.m_A == null)
			{
				this.m_A = this.m_B;
			}
			return;
		}
	}

	private bool A(object A)
	{
		string text = txtSearch.Text.ToLower();
		if (text.Length != 0)
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
			if (Operators.CompareString(text, this.m_B.ToLower(), TextCompare: false) != 0)
			{
				DefinedName definedName = (DefinedName)A;
				if (cbxSearch.SelectedIndex == 1)
				{
					while (true)
					{
						switch (6)
						{
						case 0:
							break;
						default:
							if (definedName.Label.ToLower().Contains(text))
							{
								while (true)
								{
									switch (7)
									{
									case 0:
										break;
									default:
										return true;
									}
								}
							}
							definedName.IsChecked = false;
							return false;
						}
					}
				}
				if (definedName.Label.ToLower().StartsWith(text))
				{
					return true;
				}
				definedName.IsChecked = false;
				return false;
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
		}
		return true;
	}

	private void SearchTextChanged(object sender, TextChangedEventArgs e)
	{
		D();
	}

	private void cbxSearch_SelectionChanged(object sender, SelectionChangedEventArgs e)
	{
		D();
	}

	private void D()
	{
		try
		{
			this.m_A.CancelAsync();
			this.m_B.CancelAsync();
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		this.m_B = true;
		CollectionViewSource.GetDefaultView(this.m_B).Refresh();
		this.m_B = false;
		P();
		N();
		O();
		F();
		H();
	}

	private void txtSearch_GotFocus(object sender, RoutedEventArgs e)
	{
		System.Windows.Controls.TextBox textBox = txtSearch;
		if (Operators.CompareString(textBox.Text, this.m_B, TextCompare: false) == 0)
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
			textBox.Foreground = new SolidColorBrush(SystemColors.ControlTextColor);
			textBox.TextChanged -= SearchTextChanged;
			textBox.Clear();
			textBox.TextChanged += SearchTextChanged;
		}
		textBox = null;
	}

	private void txtSearch_LostFocus(object sender, RoutedEventArgs e)
	{
		if (txtSearch.Text.Length == 0)
		{
			E();
		}
	}

	private void E()
	{
		System.Windows.Controls.TextBox textBox = txtSearch;
		textBox.Foreground = new SolidColorBrush(Colors.DarkGray);
		textBox.TextChanged -= SearchTextChanged;
		textBox.Text = this.m_B;
		textBox.TextChanged += SearchTextChanged;
		_ = null;
	}

	private void wpfNameScrubber_PreviewKeyDown(object sender, System.Windows.Input.KeyEventArgs e)
	{
		if (e.Key != Key.Escape || !txtSearch.IsFocused)
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
			lvNames.Focus();
			e.Handled = true;
			return;
		}
	}

	private void NameCheckedChanged(object sender, RoutedEventArgs e)
	{
		if (this.m_B)
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
			F();
			H();
			return;
		}
	}

	private void F()
	{
		int count = A().Count;
		System.Windows.Controls.CheckBox checkBox = chkCheckAll;
		checkBox.Checked -= chkCheckAll_CheckedChanged;
		checkBox.Unchecked -= chkCheckAll_CheckedChanged;
		if (count == 0)
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
			checkBox.IsChecked = false;
		}
		else if (count == lvNames.Items.Count)
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
			checkBox.IsChecked = true;
		}
		else
		{
			checkBox.IsChecked = null;
		}
		checkBox.Checked += chkCheckAll_CheckedChanged;
		checkBox.Unchecked += chkCheckAll_CheckedChanged;
		checkBox = null;
	}

	private void chkCheckAll_CheckedChanged(object sender, RoutedEventArgs e)
	{
		bool value = chkCheckAll.IsChecked.Value;
		this.m_B = true;
		checked
		{
			int num = lvNames.Items.Count - 1;
			for (int i = 0; i <= num; i++)
			{
				((DefinedName)lvNames.Items[i]).IsChecked = value;
			}
			this.m_B = false;
			lvNames.Focus();
			H();
		}
	}

	private void G()
	{
		System.Windows.Controls.CheckBox checkBox = chkCheckAll;
		checkBox.Checked -= chkCheckAll_CheckedChanged;
		checkBox.Unchecked -= chkCheckAll_CheckedChanged;
		checkBox.IsChecked = false;
		checkBox.Checked += chkCheckAll_CheckedChanged;
		checkBox.Unchecked += chkCheckAll_CheckedChanged;
		_ = null;
	}

	private void SpacebarToggle(object sender, System.Windows.Input.KeyEventArgs e)
	{
		if (e.Key != Key.Space)
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
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
			if (lvNames.SelectedItems.Count <= 0)
			{
				return;
			}
			this.m_B = true;
			try
			{
				bool isChecked = !((DefinedName)lvNames.SelectedItems[0]).IsChecked;
				enumerator = lvNames.SelectedItems.GetEnumerator();
				try
				{
					while (enumerator.MoveNext())
					{
						((DefinedName)enumerator.Current).IsChecked = isChecked;
					}
					while (true)
					{
						switch (1)
						{
						case 0:
							break;
						default:
							goto end_IL_0098;
						}
						continue;
						end_IL_0098:
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
				F();
				H();
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			this.m_B = false;
			e.Handled = true;
			return;
		}
	}

	private ObservableCollection<DefinedName> A()
	{
		ObservableCollection<DefinedName> b = this.m_B;
		Func<DefinedName, bool> predicate;
		if (_Closure_0024__.I == null)
		{
			predicate = (_Closure_0024__.I = [SpecialName] (DefinedName A) => A.IsChecked);
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			predicate = _Closure_0024__.I;
		}
		return new ObservableCollection<DefinedName>(b.Where(predicate));
	}

	private void H()
	{
		ObservableCollection<DefinedName> observableCollection;
		try
		{
			observableCollection = A();
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
			return;
		}
		int count = observableCollection.Count;
		if (count != 0)
		{
			if (count != 1)
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
				btnHide.IsEnabled = true;
				btnUnhide.IsEnabled = true;
				btnApply.IsEnabled = true;
				btnUnapply.IsEnabled = true;
				btnDelete.IsEnabled = true;
				btnEditName.IsEnabled = false;
				btnEditRefersTo.IsEnabled = false;
			}
			else
			{
				if (ExcelAddIn1.Formulas.Names.A(observableCollection[0].Label))
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
					btnHide.IsEnabled = false;
					btnUnhide.IsEnabled = false;
					btnApply.IsEnabled = false;
					btnUnapply.IsEnabled = false;
				}
				else
				{
					btnHide.IsEnabled = observableCollection[0].Name.Visible;
					btnUnhide.IsEnabled = !btnHide.IsEnabled;
					btnApply.IsEnabled = true;
					btnUnapply.IsEnabled = true;
				}
				btnDelete.IsEnabled = true;
				btnEditName.IsEnabled = true;
				btnEditRefersTo.IsEnabled = true;
			}
		}
		else
		{
			btnHide.IsEnabled = false;
			btnUnhide.IsEnabled = false;
			btnApply.IsEnabled = false;
			btnUnapply.IsEnabled = false;
			btnDelete.IsEnabled = false;
			btnEditName.IsEnabled = false;
			btnEditRefersTo.IsEnabled = false;
		}
		observableCollection = null;
	}

	private void btnHide_Click(object sender, RoutedEventArgs e)
	{
		ObservableCollection<DefinedName> observableCollection = A();
		bool flag = cbxFilter.SelectedIndex == 1;
		using (IEnumerator<DefinedName> enumerator = observableCollection.GetEnumerator())
		{
			while (enumerator.MoveNext())
			{
				DefinedName current = enumerator.Current;
				try
				{
					current.Name.Visible = false;
					current.Opacity = this.m_A;
					if (!flag)
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
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						current.IsChecked = false;
						this.m_B.Remove(current);
						break;
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
					break;
				default:
					goto end_IL_008d;
				}
				continue;
				end_IL_008d:
				break;
			}
		}
		observableCollection = null;
		I();
		if (flag)
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
			H();
		}
		else
		{
			btnHide.IsEnabled = false;
			btnUnhide.IsEnabled = true;
		}
		lvNames.Focus();
	}

	private void btnUnhide_Click(object sender, RoutedEventArgs e)
	{
		ObservableCollection<DefinedName> observableCollection = A();
		bool flag = cbxFilter.SelectedIndex == 2;
		IEnumerator<DefinedName> enumerator = default(IEnumerator<DefinedName>);
		try
		{
			enumerator = observableCollection.GetEnumerator();
			while (enumerator.MoveNext())
			{
				DefinedName current = enumerator.Current;
				try
				{
					current.Name.Visible = true;
					current.Opacity = 1f;
					if (!flag)
					{
						continue;
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
						current.IsChecked = false;
						this.m_B.Remove(current);
						break;
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
				switch (4)
				{
				case 0:
					break;
				default:
					goto end_IL_008c;
				}
				continue;
				end_IL_008c:
				break;
			}
		}
		finally
		{
			if (enumerator != null)
			{
				while (true)
				{
					switch (3)
					{
					case 0:
						continue;
					}
					enumerator.Dispose();
					break;
				}
			}
		}
		observableCollection = null;
		I();
		if (flag)
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
			H();
		}
		else
		{
			btnUnhide.IsEnabled = false;
			btnHide.IsEnabled = true;
		}
		lvNames.Focus();
	}

	private void I()
	{
		N();
		O();
		F();
	}

	private void btnEditname_Click(object sender, RoutedEventArgs e)
	{
		ObservableCollection<DefinedName> observableCollection = A();
		if (observableCollection.Count == 1)
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
			DefinedName definedName = observableCollection[0];
			string text = Forms.InputBox(System.Windows.Window.GetWindow(this), VH.A(178082), VH.A(178099), definedName.Name.Name);
			if (Operators.CompareString(text, string.Empty, TextCompare: false) != 0)
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
				if (text.Length > 0)
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
					if (Operators.CompareString(text, definedName.Name.Name, TextCompare: false) != 0)
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
							if (ExcelAddIn1.Formulas.Names.A().IsMatch(text))
							{
								while (true)
								{
									switch (6)
									{
									case 0:
										continue;
									}
									try
									{
										definedName.Name.Name = text;
										definedName.Label = text;
										P();
									}
									catch (Exception ex)
									{
										ProjectData.SetProjectError(ex);
										Exception ex2 = ex;
										C(ex2.Message);
										clsReporting.LogException(ex2);
										ProjectData.ClearProjectError();
									}
									break;
								}
							}
							else
							{
								Forms.WarningMessage(VH.A(178134));
							}
						}
						catch (Exception ex3)
						{
							ProjectData.SetProjectError(ex3);
							Exception ex4 = ex3;
							ProjectData.ClearProjectError();
						}
					}
				}
			}
			definedName = null;
		}
		observableCollection = null;
	}

	private void btnEditRefersTo_Click(object sender, RoutedEventArgs e)
	{
		ObservableCollection<DefinedName> observableCollection = A();
		if (observableCollection.Count == 1)
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
			DefinedName definedName = observableCollection[0];
			wpfEditRefersTo wpfEditRefersTo2 = new wpfEditRefersTo();
			wpfEditRefersTo2.txtRefersTo.Text = definedName.Name.RefersTo.ToString();
			wpfEditRefersTo2.Owner = System.Windows.Window.GetWindow(this);
			wpfEditRefersTo2.ShowDialog();
			if (wpfEditRefersTo2.DialogResult.HasValue)
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
				if (wpfEditRefersTo2.DialogResult.Value)
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
						definedName.Name.RefersTo = wpfEditRefersTo2.txtRefersTo.Text;
						definedName.RefersTo = wpfEditRefersTo2.txtRefersTo.Text;
						P();
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						Forms.ErrorMessage(ex2.Message);
						clsReporting.LogException(ex2);
						ProjectData.ClearProjectError();
					}
				}
			}
			wpfEditRefersTo2 = null;
			definedName = null;
		}
		observableCollection = null;
	}

	private void StartBackgroundWorker(object sender, RoutedEventArgs e)
	{
		if (this.m_A.IsBusy)
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
			A();
			btnClose.Focus();
			System.Windows.Controls.Button button = (System.Windows.Controls.Button)sender;
			TextBlock textBlock = lblCount;
			if (button == btnDelete)
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
				textBlock.Text = VH.A(178702);
				bdrAbort.ToolTip = VH.A(178725);
			}
			else if (button == btnUnapply)
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
				textBlock.Text = VH.A(178766);
				bdrAbort.ToolTip = VH.A(178793);
			}
			else if (button == btnApply)
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
				textBlock.Text = VH.A(178838);
				bdrAbort.ToolTip = VH.A(178861);
			}
			textBlock = null;
			this.m_A.Clear();
			this.m_A.RunWorkerAsync(button);
			button = null;
			return;
		}
	}

	private void bw_DoWork(object sender, DoWorkEventArgs e)
	{
		base.Dispatcher.Invoke([SpecialName] () =>
		{
			btnClean.IsEnabled = false;
		});
		if (e.Argument == btnDelete)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					J();
					return;
				}
			}
		}
		if (e.Argument == btnUnapply)
		{
			L();
		}
		else if (e.Argument == btnApply)
		{
			K();
		}
	}

	private void bw_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
	{
		btnClean.IsEnabled = true;
		N();
		O();
		H();
		grdLoading.Visibility = Visibility.Hidden;
		if (this.m_A == null || !this.m_A.Any())
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
			A(null, "");
			return;
		}
	}

	private void A(List<Range> A = null, string B = "")
	{
		if (A == null)
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
			A = this.m_A;
		}
		if (B.Length == 0)
		{
			B = this.m_A;
		}
		if (Forms.YesNoMessage(B) != System.Windows.Forms.DialogResult.Yes)
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
			wpfGoTo obj = new wpfGoTo(A);
			obj.Owner = System.Windows.Window.GetWindow(this);
			obj.WindowStartupLocation = WindowStartupLocation.CenterOwner;
			obj.ShowDialog();
			_ = null;
			return;
		}
	}

	private void J()
	{
		QG a = default(QG);
		QG CS_0024_003C_003E8__locals11 = new QG(a);
		CS_0024_003C_003E8__locals11.A = this;
		ObservableCollection<DefinedName> observableCollection = A();
		int count = observableCollection.Count;
		if (count == 0)
		{
			observableCollection = null;
			return;
		}
		int num = 1;
		string text;
		if (count <= 1)
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
			text = VH.A(178902);
		}
		else
		{
			text = VH.A(179001);
		}
		DialogResult dialogResult = Forms.YesNoCancelMessage(text + VH.A(179104));
		checked
		{
			bool flag = default(bool);
			bool flag2 = default(bool);
			bool visible = default(bool);
			int num3 = default(int);
			if (dialogResult != System.Windows.Forms.DialogResult.Cancel)
			{
				while (true)
				{
					switch (5)
					{
					case 0:
						break;
					default:
					{
						if (dialogResult != System.Windows.Forms.DialogResult.Yes)
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
							if (dialogResult != System.Windows.Forms.DialogResult.No)
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
							}
							else
							{
								flag = false;
							}
						}
						else
						{
							flag = true;
						}
						Microsoft.Office.Interop.Excel.Application a2 = this.m_A;
						a2.ScreenUpdating = false;
						a2.EnableEvents = false;
						a2.DisplayAlerts = false;
						XlCalculation calculation = a2.Calculation;
						a2.Calculation = XlCalculation.xlCalculationManual;
						_ = a2.ActiveCell;
						_ = null;
						List<Range> list = new List<Range>();
						try
						{
							int num2 = observableCollection.Count - 1;
							while (true)
							{
								if (num2 < 0)
								{
									while (true)
									{
										switch (6)
										{
										case 0:
											break;
										default:
											goto end_IL_0220;
										}
										continue;
										end_IL_0220:
										break;
									}
									break;
								}
								if (this.m_A.CancellationPending)
								{
									flag2 = true;
									break;
								}
								this.m_A.ReportProgress((int)Math.Round((double)num / (double)count * 100.0));
								CS_0024_003C_003E8__locals11.A = observableCollection[num2];
								Name name;
								try
								{
									name = CS_0024_003C_003E8__locals11.A.Name;
									if (flag)
									{
										while (true)
										{
											switch (6)
											{
											case 0:
												break;
											default:
												try
												{
													list = ExcelAddIn1.Formulas.Names.B(list, name);
												}
												catch (Exception ex)
												{
													ProjectData.SetProjectError(ex);
													Exception ex2 = ex;
													ProjectData.ClearProjectError();
												}
												goto end_IL_015b;
											}
											continue;
											end_IL_015b:
											break;
										}
									}
								}
								catch (Exception ex3)
								{
									ProjectData.SetProjectError(ex3);
									Exception ex4 = ex3;
									ProjectData.ClearProjectError();
									goto IL_0212;
								}
								try
								{
									visible = name.Visible;
									name.Visible = true;
									name.Delete();
									Dispatcher dispatcher = base.Dispatcher;
									Action callback;
									if (CS_0024_003C_003E8__locals11.A != null)
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
										callback = CS_0024_003C_003E8__locals11.A;
									}
									else
									{
										callback = (CS_0024_003C_003E8__locals11.A = [SpecialName] () =>
										{
											CS_0024_003C_003E8__locals11.A.m_A.Remove(CS_0024_003C_003E8__locals11.A);
											CS_0024_003C_003E8__locals11.A.m_B.Remove(CS_0024_003C_003E8__locals11.A);
										});
									}
									dispatcher.Invoke(callback);
									num3++;
								}
								catch (Exception ex5)
								{
									ProjectData.SetProjectError(ex5);
									Exception ex6 = ex5;
									name.Visible = visible;
									ProjectData.ClearProjectError();
								}
								CS_0024_003C_003E8__locals11.A = null;
								name = null;
								num++;
								goto IL_0212;
								IL_0212:
								num2 += -1;
							}
						}
						catch (Exception ex7)
						{
							ProjectData.SetProjectError(ex7);
							Exception ex8 = ex7;
							ProjectData.ClearProjectError();
						}
						Microsoft.Office.Interop.Excel.Application a3 = this.m_A;
						a3.Calculation = calculation;
						a3.ScreenUpdating = true;
						a3.EnableEvents = true;
						a3.DisplayAlerts = true;
						_ = null;
						if (flag2)
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
							A(VH.A(179370) + num3.ToString(this.m_C) + VH.A(179455));
						}
						else if (num3 > 0)
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
							if (num3 != count)
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
								num = count - num3;
								string[] obj = new string[7]
								{
									VH.A(179474),
									Conversions.ToString(num),
									VH.A(41385),
									(num == 1) ? VH.A(67336) : VH.A(179509),
									VH.A(179520),
									null,
									null
								};
								string text2;
								if (num != 1)
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
									text2 = VH.A(179509);
								}
								else
								{
									text2 = VH.A(67336);
								}
								obj[5] = text2;
								obj[6] = VH.A(179561);
								B(string.Concat(obj));
							}
						}
						else if (count == 1)
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
							C(VH.A(179666));
						}
						else
						{
							C(VH.A(179856));
						}
						base.Dispatcher.Invoke([SpecialName] () =>
						{
							F();
							P();
							B();
						});
						if (flag)
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
							if (!list.Any())
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
								A(VH.A(180070));
							}
							else
							{
								this.m_A = list;
								this.m_A = A(list, count);
							}
						}
						list = null;
						observableCollection = null;
						return;
					}
					}
				}
			}
			observableCollection = null;
		}
	}

	private void K()
	{
		SG a = default(SG);
		SG CS_0024_003C_003E8__locals10 = new SG(a);
		CS_0024_003C_003E8__locals10.A = this;
		XlCalculation calculation = this.m_A.Calculation;
		List<Range> list = new List<Range>();
		List<Range> list2 = new List<Range>();
		ObservableCollection<DefinedName> observableCollection = A();
		int count = observableCollection.Count;
		if (count == 0)
		{
			observableCollection = null;
			return;
		}
		int num = 1;
		checked
		{
			try
			{
				this.m_A.ScreenUpdating = false;
				this.m_A.EnableEvents = false;
				this.m_A.Calculation = XlCalculation.xlCalculationManual;
				Range activeCell = this.m_A.ActiveCell;
				IEnumerator<DefinedName> enumerator = default(IEnumerator<DefinedName>);
				Name name = default(Name);
				try
				{
					enumerator = observableCollection.GetEnumerator();
					RG rG = default(RG);
					while (enumerator.MoveNext())
					{
						DefinedName current = enumerator.Current;
						if (this.m_A.CancellationPending)
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
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						this.m_A.ReportProgress((int)Math.Round((double)num / (double)count * 100.0));
						Range refersToRange;
						try
						{
							name = current.Name;
							refersToRange = name.RefersToRange;
						}
						catch (Exception ex)
						{
							ProjectData.SetProjectError(ex);
							Exception ex2 = ex;
							ProjectData.ClearProjectError();
							continue;
						}
						list = ExcelAddIn1.Formulas.Names.GetDependents(refersToRange, name);
						using (List<Range>.Enumerator enumerator2 = list.GetEnumerator())
						{
							while (enumerator2.MoveNext())
							{
								Range current2 = enumerator2.Current;
								if (Operators.CompareString(current2.Worksheet.Name, refersToRange.Worksheet.Name, TextCompare: false) != 0)
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
								string formula = current2.Formula.ToString();
								object objectValue = RuntimeHelpers.GetObjectValue(current2.get_Value(RuntimeHelpers.GetObjectValue(Missing.Value)));
								try
								{
									current2.ApplyNames(new string[1] { name.Name }, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlApplyNamesOrder.xlRowThenColumn, RuntimeHelpers.GetObjectValue(Missing.Value));
								}
								catch (Exception ex3)
								{
									ProjectData.SetProjectError(ex3);
									Exception ex4 = ex3;
									ProjectData.ClearProjectError();
								}
								if (Operators.ConditionalCompareObjectNotEqual(current2.get_Value(RuntimeHelpers.GetObjectValue(Missing.Value)), objectValue, TextCompare: false))
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
									current2.Formula = formula;
								}
								try
								{
									rG = new RG(rG);
									rG.A = current2.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value));
									if (list2.Find(rG.A) == null)
									{
										list2.Add(current2);
									}
								}
								catch (Exception ex5)
								{
									ProjectData.SetProjectError(ex5);
									Exception ex6 = ex5;
									ProjectData.ClearProjectError();
								}
							}
							while (true)
							{
								switch (3)
								{
								case 0:
									break;
								default:
									goto end_IL_02a6;
								}
								continue;
								end_IL_02a6:
								break;
							}
						}
						num++;
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
				try
				{
					base.Dispatcher.Invoke([SpecialName] () =>
					{
						if (CS_0024_003C_003E8__locals10.A.lvNames.SelectedItems.Count != 0)
						{
							while (true)
							{
								switch (4)
								{
								case 0:
									break;
								default:
								{
									if (1 == 0)
									{
										/*OpCode not supported: LdMemberToken*/;
									}
									System.Windows.Controls.ListView listView = CS_0024_003C_003E8__locals10.A.lvNames;
									if (((DefinedName)listView.SelectedItems[0]).IsChecked)
									{
										CS_0024_003C_003E8__locals10.A = listView.Items.IndexOf(RuntimeHelpers.GetObjectValue(listView.SelectedItems[0]));
										listView.SelectedItems.Clear();
										CS_0024_003C_003E8__locals10.A.A(CS_0024_003C_003E8__locals10.A);
									}
									listView = null;
									return;
								}
								}
							}
						}
					});
				}
				catch (Exception ex7)
				{
					ProjectData.SetProjectError(ex7);
					Exception ex8 = ex7;
					ProjectData.ClearProjectError();
				}
				try
				{
					ExcelAddIn1.Formulas.Names.A(activeCell);
				}
				catch (Exception ex9)
				{
					ProjectData.SetProjectError(ex9);
					Exception ex10 = ex9;
					ProjectData.ClearProjectError();
				}
				finally
				{
					this.m_A.Calculation = calculation;
					this.m_A.ScreenUpdating = true;
					this.m_A.EnableEvents = true;
				}
				try
				{
					CS_0024_003C_003E8__locals10.A = list2.Count;
					string text;
					if (count != 1)
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
						text = VH.A(180224);
					}
					else
					{
						text = VH.A(180269) + name.Name + VH.A(180282);
					}
					string text2 = text;
					text2 = text2 + VH.A(180293) + CS_0024_003C_003E8__locals10.A + VH.A(41385) + ((CS_0024_003C_003E8__locals10.A == 1) ? VH.A(116966) : VH.A(116955)) + VH.A(117100) + ((CS_0024_003C_003E8__locals10.A == 1) ? VH.A(117049) : VH.A(117026)) + VH.A(117068);
					this.m_A = text2;
					this.m_A = list2;
				}
				catch (Exception ex11)
				{
					ProjectData.SetProjectError(ex11);
					Exception ex12 = ex11;
					ProjectData.ClearProjectError();
				}
			}
			catch (Exception ex13)
			{
				ProjectData.SetProjectError(ex13);
				Exception ex14 = ex13;
				ProjectData.ClearProjectError();
			}
			finally
			{
				Range refersToRange = null;
				Range activeCell = null;
				list2 = null;
				list = null;
				Name name = null;
			}
		}
	}

	private void L()
	{
		List<Range> list = new List<Range>();
		ObservableCollection<DefinedName> observableCollection = A();
		int count = observableCollection.Count;
		if (count == 0)
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
					observableCollection = null;
					return;
				}
			}
		}
		int num = 1;
		this.m_A.ScreenUpdating = false;
		this.m_A.EnableEvents = false;
		Range activeCell = this.m_A.ActiveCell;
		checked
		{
			IEnumerator<DefinedName> enumerator = default(IEnumerator<DefinedName>);
			Name name;
			try
			{
				enumerator = observableCollection.GetEnumerator();
				while (true)
				{
					if (enumerator.MoveNext())
					{
						DefinedName current = enumerator.Current;
						if (this.m_A.CancellationPending)
						{
							while (true)
							{
								switch (2)
								{
								case 0:
									break;
								default:
									goto end_IL_0086;
								}
								continue;
								end_IL_0086:
								break;
							}
							break;
						}
						this.m_A.ReportProgress((int)Math.Round((double)num / (double)count * 100.0));
						try
						{
							name = current.Name;
							try
							{
								list = ExcelAddIn1.Formulas.Names.B(list, name);
							}
							catch (Exception ex)
							{
								ProjectData.SetProjectError(ex);
								Exception ex2 = ex;
								ProjectData.ClearProjectError();
							}
						}
						catch (Exception ex3)
						{
							ProjectData.SetProjectError(ex3);
							Exception ex4 = ex3;
							ProjectData.ClearProjectError();
							continue;
						}
						num++;
						continue;
					}
					while (true)
					{
						switch (2)
						{
						case 0:
							break;
						default:
							goto end_IL_0100;
						}
						continue;
						end_IL_0100:
						break;
					}
					break;
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
			observableCollection = null;
			ExcelAddIn1.Formulas.Names.A(activeCell);
			this.m_A.ScreenUpdating = true;
			this.m_A.EnableEvents = true;
			if (!list.Any())
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
				A(VH.A(180070));
			}
			else
			{
				base.Dispatcher.Invoke([SpecialName] () =>
				{
					lvDependencies.ItemsSource = null;
				});
				this.m_A = list;
				this.m_A = A(list, count);
			}
			activeCell = null;
			list = null;
			name = null;
		}
	}

	private string A(List<Range> A, int B)
	{
		int num = 0;
		try
		{
			num = A.Count;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		if (num > 0)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
				{
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					if (B == 1)
					{
						while (true)
						{
							switch (2)
							{
							case 0:
								break;
							default:
							{
								string[] obj = new string[7]
								{
									VH.A(180318),
									Conversions.ToString(num),
									VH.A(41385),
									null,
									null,
									null,
									null
								};
								string text;
								if (num != 1)
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
									text = VH.A(116955);
								}
								else
								{
									text = VH.A(116966);
								}
								obj[3] = text;
								obj[4] = VH.A(117100);
								string text2;
								if (num != 1)
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
									text2 = VH.A(117026);
								}
								else
								{
									text2 = VH.A(117049);
								}
								obj[5] = text2;
								obj[6] = VH.A(117068);
								return string.Concat(obj);
							}
							}
						}
					}
					string[] obj2 = new string[9]
					{
						VH.A(180359),
						Conversions.ToString(B),
						VH.A(180380),
						Conversions.ToString(num),
						VH.A(41385),
						null,
						null,
						null,
						null
					};
					string text3;
					if (num != 1)
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
						text3 = VH.A(116955);
					}
					else
					{
						text3 = VH.A(116966);
					}
					obj2[5] = text3;
					obj2[6] = VH.A(116975);
					string text4;
					if (num != 1)
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
						text4 = VH.A(117026);
					}
					else
					{
						text4 = VH.A(117049);
					}
					obj2[7] = text4;
					obj2[8] = VH.A(117068);
					return string.Concat(obj2);
				}
				}
			}
		}
		string result = default(string);
		return result;
	}

	private void btnClean_Click(object sender, RoutedEventArgs e)
	{
		wpfCleanNames wpfCleanNames2 = new wpfCleanNames([SpecialName] () =>
		{
			A(A: true);
		});
		wpfCleanNames2.Owner = System.Windows.Window.GetWindow(this);
		wpfCleanNames2.ShowDialog();
		if (wpfCleanNames2.DialogResult.HasValue)
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
			if (wpfCleanNames2.DialogResult.Value)
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
				M();
			}
		}
		wpfCleanNames2 = null;
	}

	private void M()
	{
		cbxFilter.SelectionChanged -= cbxFilter_SelectionChanged;
		cbxFilter.SelectedIndex = -1;
		cbxFilter.SelectionChanged += cbxFilter_SelectionChanged;
		this.m_A = null;
		cbxFilter.SelectedIndex = 0;
	}

	private void N()
	{
		B(lvNames.Items.Count);
	}

	private void B(int A)
	{
		base.Dispatcher.Invoke([SpecialName] () =>
		{
			lblCount.Text = A.ToString(this.m_C) + VH.A(212129);
		});
	}

	private void O()
	{
		if (Operators.CompareString(lblCount.Text, VH.A(180401), TextCompare: false) == 0)
		{
			lblCount.Text = VH.A(180416);
		}
	}

	private void P()
	{
		A(gvNames);
	}

	private void Q()
	{
		A(gvDependents);
	}

	private void A(GridView A)
	{
		Forms.AutoResizeGridView(A);
	}

	private void A(string A)
	{
		base.Dispatcher.Invoke([SpecialName] () =>
		{
			Forms.InfoMessage(System.Windows.Window.GetWindow(this), A);
		});
	}

	private void B(string A)
	{
		base.Dispatcher.Invoke([SpecialName] () =>
		{
			Forms.WarningMessage(System.Windows.Window.GetWindow(this), A);
		});
	}

	private void C(string A)
	{
		base.Dispatcher.Invoke([SpecialName] () =>
		{
			Forms.ErrorMessage(System.Windows.Window.GetWindow(this), A);
		});
	}

	private void D(string A)
	{
		PropertyChangedEventHandler propertyChangedEventHandler = this.m_A;
		if (propertyChangedEventHandler == null)
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
			propertyChangedEventHandler(this, new PropertyChangedEventArgs(A));
			return;
		}
	}

	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	[DebuggerNonUserCode]
	public void InitializeComponent()
	{
		if (this.m_C)
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
			this.m_C = true;
			Uri resourceLocator = new Uri(VH.A(180429), UriKind.Relative);
			System.Windows.Application.LoadComponent(this, resourceLocator);
			return;
		}
	}

	void IComponentConnector.InitializeComponent()
	{
		//ILSpy generated this explicit interface implementation from .override directive in InitializeComponent
		this.InitializeComponent();
	}

	[EditorBrowsable(EditorBrowsableState.Never)]
	[DebuggerNonUserCode]
	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	public void System_Windows_Markup_IComponentConnector_Connect(int connectionId, object target)
	{
		if (connectionId == 1)
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
					grdMain = (Grid)target;
					return;
				}
			}
		}
		if (connectionId == 2)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					cbxFilter = (System.Windows.Controls.ComboBox)target;
					return;
				}
			}
		}
		if (connectionId == 3)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					cbxSearch = (System.Windows.Controls.ComboBox)target;
					return;
				}
			}
		}
		if (connectionId == 4)
		{
			txtSearch = (System.Windows.Controls.TextBox)target;
			return;
		}
		if (connectionId == 5)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					lvNames = (System.Windows.Controls.ListView)target;
					return;
				}
			}
		}
		if (connectionId == 6)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					gvNames = (GridView)target;
					return;
				}
			}
		}
		if (connectionId == 7)
		{
			chkCheckAll = (System.Windows.Controls.CheckBox)target;
			return;
		}
		if (connectionId == 9)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					grdLoading = (Grid)target;
					return;
				}
			}
		}
		if (connectionId == 10)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					pbLoading = (System.Windows.Controls.ProgressBar)target;
					return;
				}
			}
		}
		if (connectionId == 11)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					bdrAbort = (System.Windows.Controls.Border)target;
					return;
				}
			}
		}
		if (connectionId == 12)
		{
			stkControls = (StackPanel)target;
			return;
		}
		if (connectionId == 13)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					btnDelete = (System.Windows.Controls.Button)target;
					return;
				}
			}
		}
		if (connectionId == 14)
		{
			btnHide = (System.Windows.Controls.Button)target;
			return;
		}
		if (connectionId == 15)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					btnUnhide = (System.Windows.Controls.Button)target;
					return;
				}
			}
		}
		if (connectionId == 16)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					btnApply = (System.Windows.Controls.Button)target;
					return;
				}
			}
		}
		if (connectionId == 17)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					btnUnapply = (System.Windows.Controls.Button)target;
					return;
				}
			}
		}
		if (connectionId == 18)
		{
			btnEditName = (System.Windows.Controls.Button)target;
			return;
		}
		if (connectionId == 19)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					btnEditRefersTo = (System.Windows.Controls.Button)target;
					return;
				}
			}
		}
		if (connectionId == 20)
		{
			lblCount = (TextBlock)target;
			return;
		}
		if (connectionId == 21)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					lvDependencies = (System.Windows.Controls.ListView)target;
					return;
				}
			}
		}
		if (connectionId == 22)
		{
			gvDependents = (GridView)target;
			return;
		}
		if (connectionId == 23)
		{
			chkShowDependents = (System.Windows.Controls.CheckBox)target;
			return;
		}
		if (connectionId == 24)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					btnClean = (System.Windows.Controls.Button)target;
					return;
				}
			}
		}
		if (connectionId == 25)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					btnClose = (System.Windows.Controls.Button)target;
					return;
				}
			}
		}
		this.m_C = true;
	}

	void IComponentConnector.Connect(int connectionId, object target)
	{
		//ILSpy generated this explicit interface implementation from .override directive in System_Windows_Markup_IComponentConnector_Connect
		this.System_Windows_Markup_IComponentConnector_Connect(connectionId, target);
	}

	[EditorBrowsable(EditorBrowsableState.Never)]
	[DebuggerNonUserCode]
	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	public void System_Windows_Markup_IStyleConnector_Connect(int connectionId, object target)
	{
		if (connectionId != 8)
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
			((System.Windows.Controls.CheckBox)target).Checked += NameCheckedChanged;
			((System.Windows.Controls.CheckBox)target).Unchecked += NameCheckedChanged;
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
	private void R()
	{
		B();
	}

	[SpecialName]
	[CompilerGenerated]
	private void S()
	{
		C();
	}

	[SpecialName]
	[CompilerGenerated]
	private void T()
	{
		btnClean.IsEnabled = false;
	}

	[SpecialName]
	[CompilerGenerated]
	private void U()
	{
		F();
		P();
		B();
	}

	[SpecialName]
	[CompilerGenerated]
	private void V()
	{
		lvDependencies.ItemsSource = null;
	}

	[SpecialName]
	[CompilerGenerated]
	private void W()
	{
		A(A: true);
	}
}
