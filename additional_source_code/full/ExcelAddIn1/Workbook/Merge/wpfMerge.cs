using System;
using System.CodeDom.Compiler;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Media;
using A;
using MacabacusMacros;
using MacabacusMacros.ExcelHelpers;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Workbook.Merge;

[DesignerGenerated]
public sealed class wpfMerge : System.Windows.Window, INotifyPropertyChanged, IComponentConnector, IStyleConnector
{
	[Serializable]
	[CompilerGenerated]
	internal sealed class _Closure_0024__
	{
		public static readonly _Closure_0024__ A;

		public static Func<SourceWorkbook, bool> A;

		public static Func<SourceSheet, bool> A;

		public static Func<FileInfo, string> A;

		public static Func<SourceWorkbook, string> A;

		public static Func<SourceWorkbook, string> B;

		public static Func<SourceWorkbook, string> C;

		public static Func<SourceWorkbook, string> D;

		public static Func<SourceWorkbook, DateTime> A;

		public static Func<SourceWorkbook, DateTime> B;

		public static Func<SourceWorkbook, DateTime> C;

		public static Func<SourceWorkbook, DateTime> D;

		public static Func<DestinationSheet, bool> A;

		static _Closure_0024__()
		{
			_Closure_0024__.A = new _Closure_0024__();
		}

		[SpecialName]
		internal bool A(SourceWorkbook A)
		{
			bool? isChecked = A.IsChecked;
			bool? obj;
			if (!isChecked.HasValue)
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
				obj = null;
			}
			else
			{
				obj = isChecked == true;
			}
			isChecked = obj;
			return isChecked.Value;
		}

		[SpecialName]
		internal bool A(SourceSheet A)
		{
			return A.IsChecked;
		}

		[SpecialName]
		internal string A(FileInfo A)
		{
			return A.Name;
		}

		[SpecialName]
		internal string A(SourceWorkbook A)
		{
			return A.Workbook.Name;
		}

		[SpecialName]
		internal string B(SourceWorkbook A)
		{
			return A.Workbook.Name;
		}

		[SpecialName]
		internal string C(SourceWorkbook A)
		{
			return A.Workbook.Path;
		}

		[SpecialName]
		internal string D(SourceWorkbook A)
		{
			return A.Workbook.Path;
		}

		[SpecialName]
		internal DateTime A(SourceWorkbook A)
		{
			if (A.Workbook.Path.Length > 0)
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
						return new FileInfo(A.Workbook.FullName).LastWriteTime;
					}
				}
			}
			return DateTime.MaxValue;
		}

		[SpecialName]
		internal DateTime B(SourceWorkbook A)
		{
			if (A.Workbook.Path.Length > 0)
			{
				return new FileInfo(A.Workbook.FullName).LastWriteTime;
			}
			return DateTime.MaxValue;
		}

		[SpecialName]
		internal DateTime C(SourceWorkbook A)
		{
			if (A.Workbook.Path.Length > 0)
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
						return new FileInfo(A.Workbook.FullName).CreationTime;
					}
				}
			}
			return DateTime.MaxValue;
		}

		[SpecialName]
		internal DateTime D(SourceWorkbook A)
		{
			if (A.Workbook.Path.Length > 0)
			{
				return new FileInfo(A.Workbook.FullName).CreationTime;
			}
			return DateTime.MaxValue;
		}

		[SpecialName]
		internal bool A(DestinationSheet A)
		{
			return A.Visibility == Visibility.Visible;
		}
	}

	[CompilerGenerated]
	internal sealed class NG
	{
		public SourceSheet A;

		[SpecialName]
		internal bool A(DestinationSheet A)
		{
			return A.Source == this.A;
		}
	}

	[CompilerGenerated]
	internal sealed class OG
	{
		public SourceSheet A;

		[SpecialName]
		internal bool A(DestinationSheet A)
		{
			return A.Source == this.A;
		}
	}

	[CompilerGenerated]
	private PropertyChangedEventHandler m_A;

	private Microsoft.Office.Interop.Excel.Application m_A;

	private BackgroundWorker m_A;

	private List<Microsoft.Office.Interop.Excel.Workbook> m_A;

	private bool m_A;

	private ObservableCollection<SourceWorkbook> m_A;

	private ObservableCollection<DestinationWorkbook> m_A;

	[CompilerGenerated]
	private static Microsoft.Office.Interop.Excel.Workbook m_A;

	private bool m_B;

	[CompilerGenerated]
	[AccessedThroughProperty("tbCounter")]
	private TextBlock m_A;

	[AccessedThroughProperty("trvSources")]
	[CompilerGenerated]
	private System.Windows.Controls.TreeView m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("btnAdd")]
	private System.Windows.Controls.Button m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("btnAddFiles")]
	private System.Windows.Controls.MenuItem m_A;

	[AccessedThroughProperty("imgAddFiles")]
	[CompilerGenerated]
	private Image m_A;

	[AccessedThroughProperty("btnAddFolder")]
	[CompilerGenerated]
	private System.Windows.Controls.MenuItem m_B;

	[AccessedThroughProperty("imgAddFolder")]
	[CompilerGenerated]
	private Image m_B;

	[AccessedThroughProperty("btnRemove")]
	[CompilerGenerated]
	private System.Windows.Controls.Button m_B;

	[AccessedThroughProperty("btnWorkbooksSort")]
	[CompilerGenerated]
	private System.Windows.Controls.Button m_C;

	[CompilerGenerated]
	[AccessedThroughProperty("cbxSort")]
	private System.Windows.Controls.ComboBox m_A;

	[AccessedThroughProperty("btnDown")]
	[CompilerGenerated]
	private System.Windows.Controls.Button m_D;

	[CompilerGenerated]
	[AccessedThroughProperty("btnUp")]
	private System.Windows.Controls.Button m_E;

	[CompilerGenerated]
	[AccessedThroughProperty("chkNew")]
	private System.Windows.Controls.CheckBox m_A;

	[AccessedThroughProperty("trvDestination")]
	[CompilerGenerated]
	private System.Windows.Controls.TreeView m_B;

	[CompilerGenerated]
	[AccessedThroughProperty("cbxString1")]
	private System.Windows.Controls.ComboBox m_B;

	[AccessedThroughProperty("cbxString2")]
	[CompilerGenerated]
	private System.Windows.Controls.ComboBox m_C;

	[CompilerGenerated]
	[AccessedThroughProperty("cbxString3")]
	private System.Windows.Controls.ComboBox m_D;

	[AccessedThroughProperty("cbxString4")]
	[CompilerGenerated]
	private System.Windows.Controls.ComboBox m_E;

	[CompilerGenerated]
	[AccessedThroughProperty("btnOk")]
	private System.Windows.Controls.Button m_F;

	private bool m_C;

	public ObservableCollection<SourceWorkbook> SourceWorkbooks
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(VH.A(176034));
		}
	}

	public ObservableCollection<DestinationWorkbook> DestinationWorkbooks
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(VH.A(176065));
		}
	}

	private static Microsoft.Office.Interop.Excel.Workbook ActiveWorkbook
	{
		[CompilerGenerated]
		get
		{
			return wpfMerge.m_A;
		}
		[CompilerGenerated]
		set
		{
			wpfMerge.m_A = value;
		}
	}

	internal virtual TextBlock tbCounter
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

	internal virtual System.Windows.Controls.TreeView trvSources
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
			RoutedPropertyChangedEventHandler<object> value2 = trvSources_SelectedItemChanged;
			System.Windows.Input.KeyEventHandler value3 = SpacebarToggleWorkbooks;
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
				treeView.SelectedItemChanged -= value2;
				treeView.PreviewKeyDown -= value3;
			}
			this.m_A = value;
			treeView = this.m_A;
			if (treeView == null)
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
				treeView.SelectedItemChanged += value2;
				treeView.PreviewKeyDown += value3;
				return;
			}
		}
	}

	internal virtual System.Windows.Controls.Button btnAdd
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
			RoutedEventHandler value2 = ShowAddMenu;
			System.Windows.Controls.Button button = this.m_A;
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
			this.m_A = value;
			button = this.m_A;
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

	internal virtual System.Windows.Controls.MenuItem btnAddFiles
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

	internal virtual Image imgAddFiles
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

	internal virtual System.Windows.Controls.MenuItem btnAddFolder
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

	internal virtual Image imgAddFolder
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

	internal virtual System.Windows.Controls.Button btnRemove
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
			RoutedEventHandler value2 = btnRemove_Click;
			System.Windows.Controls.Button button = this.m_B;
			if (button != null)
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

	internal virtual System.Windows.Controls.Button btnWorkbooksSort
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

	internal virtual System.Windows.Controls.ComboBox cbxSort
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

	internal virtual System.Windows.Controls.Button btnDown
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
			RoutedEventHandler value2 = MoveSourceDown;
			System.Windows.Controls.Button button = this.m_D;
			if (button != null)
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
				button.Click -= value2;
			}
			this.m_D = value;
			button = this.m_D;
			if (button != null)
			{
				button.Click += value2;
			}
		}
	}

	internal virtual System.Windows.Controls.Button btnUp
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
			RoutedEventHandler value2 = MoveSourceUp;
			System.Windows.Controls.Button button = this.m_E;
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
			this.m_E = value;
			button = this.m_E;
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

	internal virtual System.Windows.Controls.CheckBox chkNew
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

	internal virtual System.Windows.Controls.TreeView trvDestination
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
			RoutedPropertyChangedEventHandler<object> value2 = trvDestination_SelectedItemChanged;
			System.Windows.Controls.TreeView treeView = this.m_B;
			if (treeView != null)
			{
				treeView.SelectedItemChanged -= value2;
			}
			this.m_B = value;
			treeView = this.m_B;
			if (treeView == null)
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
				treeView.SelectedItemChanged += value2;
				return;
			}
		}
	}

	internal virtual System.Windows.Controls.ComboBox cbxString1
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

	internal virtual System.Windows.Controls.ComboBox cbxString2
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

	internal virtual System.Windows.Controls.ComboBox cbxString3
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

	internal virtual System.Windows.Controls.ComboBox cbxString4
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

	internal virtual System.Windows.Controls.Button btnOk
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
			RoutedEventHandler value2 = btnOk_Click;
			System.Windows.Controls.Button button = this.m_F;
			if (button != null)
			{
				button.Click -= value2;
			}
			this.m_F = value;
			button = this.m_F;
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
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
		}
	}

	public wpfMerge()
	{
		base.Loaded += wpfMerge_Loaded;
		base.Closing += wpfMerge_Closing;
		this.m_A = false;
		InitializeComponent();
		base.Icon = Forms.GetIcon();
		Forms.SetImageSource(imgAddFiles, J.ExcelSmall);
		Forms.SetImageSource(imgAddFolder, J.FolderOpen);
		this.m_A = MH.A.Application;
		ActiveWorkbook = this.m_A.ActiveWorkbook;
	}

	private void A(string A)
	{
		this.m_A?.Invoke(this, new PropertyChangedEventArgs(A));
	}

	private void wpfMerge_Loaded(object sender, RoutedEventArgs e)
	{
		DestinationWorkbooks = new ObservableCollection<DestinationWorkbook>();
		DestinationWorkbooks.Add(new DestinationWorkbook());
		chkNew.IsChecked = true;
		trvDestination.ItemsSource = DestinationWorkbooks;
		this.m_A = true;
		SourceWorkbooks = new ObservableCollection<SourceWorkbook>();
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = MH.A.Application.Workbooks.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Microsoft.Office.Interop.Excel.Workbook workbook = (Microsoft.Office.Interop.Excel.Workbook)enumerator.Current;
				if (clsPanes.IsValidWorkbook(workbook))
				{
					A(workbook);
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					(enumerator as IDisposable).Dispose();
					break;
				}
			}
		}
		this.m_A = false;
		trvSources.ItemsSource = SourceWorkbooks;
		cbxSort.SelectionChanged += SortWorkbooks;
		cbxString1.SelectedIndex = 0;
		cbxString1.SelectionChanged += NameSelectionChanged;
		cbxString2.SelectionChanged += NameSelectionChanged;
		cbxString3.SelectionChanged += NameSelectionChanged;
		cbxString4.SelectionChanged += NameSelectionChanged;
		E();
	}

	private void A(Microsoft.Office.Interop.Excel.Workbook A)
	{
		SourceWorkbook sourceWorkbook = new SourceWorkbook(A);
		SourceWorkbooks.Add(sourceWorkbook);
		IEnumerator enumerator = default(IEnumerator);
		SourceSheet sourceSheet;
		try
		{
			enumerator = A.Sheets.GetEnumerator();
			while (enumerator.MoveNext())
			{
				object objectValue = RuntimeHelpers.GetObjectValue(enumerator.Current);
				sourceSheet = new SourceSheet(RuntimeHelpers.GetObjectValue(objectValue), sourceWorkbook);
				sourceWorkbook.Sheets.Add(sourceSheet);
				DestinationWorkbooks[0].Sheets.Add(new DestinationSheet(sourceSheet));
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
				break;
			}
		}
		finally
		{
			if (enumerator is IDisposable)
			{
				while (true)
				{
					switch (3)
					{
					case 0:
						continue;
					}
					(enumerator as IDisposable).Dispose();
					break;
				}
			}
		}
		sourceWorkbook = null;
		sourceSheet = null;
	}

	private ObservableCollection<SourceWorkbook> A()
	{
		ObservableCollection<SourceWorkbook> sourceWorkbooks = SourceWorkbooks;
		Func<SourceWorkbook, bool> predicate;
		if (_Closure_0024__.A == null)
		{
			predicate = (_Closure_0024__.A = [SpecialName] (SourceWorkbook A) =>
			{
				bool? isChecked = A.IsChecked;
				bool? obj;
				if (!isChecked.HasValue)
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
					obj = null;
				}
				else
				{
					obj = isChecked == true;
				}
				isChecked = obj;
				return isChecked.Value;
			});
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			predicate = _Closure_0024__.A;
		}
		return new ObservableCollection<SourceWorkbook>(sourceWorkbooks.Where(predicate));
	}

	private ObservableCollection<SourceSheet> A(SourceWorkbook A)
	{
		return new ObservableCollection<SourceSheet>(A.Sheets.Where([SpecialName] (SourceSheet sourceSheet) => sourceSheet.IsChecked));
	}

	private void A()
	{
	}

	private void trvSources_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
	{
		if (e.NewValue is SourceSheet)
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
					btnRemove.IsEnabled = false;
					btnUp.IsEnabled = true;
					btnDown.IsEnabled = true;
					return;
				}
			}
		}
		if (e.NewValue is SourceWorkbook)
		{
			btnRemove.IsEnabled = true;
			btnUp.IsEnabled = true;
			btnDown.IsEnabled = true;
		}
		else
		{
			btnUp.IsEnabled = false;
			btnDown.IsEnabled = false;
		}
	}

	private void trvDestination_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
	{
		if (!(e.NewValue is DestinationSheet))
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
			try
			{
				SourceSheet source = ((DestinationSheet)e.NewValue).Source;
				trvSources.Focus();
				if (source.Parent.IsExpanded)
				{
					source.IsSelected = true;
				}
				else
				{
					source.Parent.IsSelected = true;
				}
				trvDestination.Focus();
				source = null;
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

	private void OnRequestBringIntoView(object sender, RequestBringIntoViewEventArgs e)
	{
		if (this.m_B)
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
			e.Handled = true;
			this.m_B = true;
			if (sender is TreeViewItem treeViewItem)
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
				Rect targetRectangle = new Rect(-1000.0, 0.0, treeViewItem.ActualWidth + 1000.0, treeViewItem.ActualHeight);
				treeViewItem.BringIntoView(targetRectangle);
			}
			this.m_B = false;
			return;
		}
	}

	private DependencyObject A(DependencyObject A)
	{
		DependencyObject parent = VisualTreeHelper.GetParent(A);
		if (!(parent is System.Windows.Controls.TreeView))
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
					return this.A(parent);
				}
			}
		}
		return parent;
	}

	private void OnSelected(object sender, RoutedEventArgs e)
	{
		((TreeViewItem)sender).BringIntoView();
		e.Handled = true;
	}

	private void WorkbookCheckedChanged(object sender, RoutedEventArgs e)
	{
		if (this.m_A)
		{
			return;
		}
		IEnumerator<SourceSheet> enumerator = default(IEnumerator<SourceSheet>);
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
			System.Windows.Controls.CheckBox checkBox = (System.Windows.Controls.CheckBox)sender;
			if (checkBox.IsChecked.HasValue)
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
				bool value = checkBox.IsChecked.Value;
				SourceWorkbook sourceWorkbook = (SourceWorkbook)checkBox.DataContext;
				this.m_A = true;
				try
				{
					enumerator = sourceWorkbook.Sheets.GetEnumerator();
					while (enumerator.MoveNext())
					{
						SourceSheet current = enumerator.Current;
						current.IsChecked = value;
						A(current, value);
					}
					while (true)
					{
						switch (7)
						{
						case 0:
							break;
						default:
							goto end_IL_00a2;
						}
						continue;
						end_IL_00a2:
						break;
					}
				}
				finally
				{
					if (enumerator != null)
					{
						while (true)
						{
							switch (1)
							{
							case 0:
								continue;
							}
							enumerator.Dispose();
							break;
						}
					}
				}
				this.m_A = false;
				sourceWorkbook = null;
			}
			checkBox = null;
			E();
			F();
			return;
		}
	}

	private void SheetCheckedChanged(object sender, RoutedEventArgs e)
	{
		if (this.m_A)
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
			System.Windows.Controls.CheckBox checkBox = (System.Windows.Controls.CheckBox)sender;
			SourceSheet sourceSheet = (SourceSheet)checkBox.DataContext;
			SourceWorkbook parent = sourceSheet.Parent;
			A(sourceSheet, checkBox.IsChecked.Value);
			int count = A(parent).Count;
			this.m_A = true;
			if (count == 0)
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
				parent.IsChecked = false;
			}
			else if (count == parent.Sheets.Count)
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
				parent.IsChecked = true;
			}
			else
			{
				parent.IsChecked = null;
			}
			this.m_A = false;
			parent = null;
			checkBox = null;
			sourceSheet = null;
			E();
			F();
			return;
		}
	}

	private void A(SourceSheet A, bool B)
	{
		DestinationSheet destinationSheet = DestinationWorkbooks[0].Sheets.Where([SpecialName] (DestinationSheet destinationSheet2) => destinationSheet2.Source == A).ElementAt(0);
		if (B)
		{
			destinationSheet.Visibility = Visibility.Visible;
		}
		else
		{
			destinationSheet.Visibility = Visibility.Collapsed;
		}
		destinationSheet = null;
	}

	private void SpacebarToggleWorkbooks(object sender, System.Windows.Input.KeyEventArgs e)
	{
		if (e.Key != Key.Space)
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
			this.m_A = true;
			try
			{
				if (trvSources.SelectedItem is SourceSheet)
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
					SourceSheet obj = (SourceSheet)trvSources.SelectedItem;
					obj.IsChecked = !obj.IsChecked;
					_ = null;
				}
				else if (trvSources.SelectedItem is SourceWorkbook)
				{
					SourceWorkbook sourceWorkbook = (SourceWorkbook)trvSources.SelectedItem;
					bool? isChecked = sourceWorkbook.IsChecked;
					bool? flag;
					if (!isChecked.HasValue)
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
						flag = isChecked;
					}
					else
					{
						flag = isChecked != true;
					}
					isChecked = flag;
					bool value = isChecked.Value;
					using (IEnumerator<SourceSheet> enumerator = sourceWorkbook.Sheets.GetEnumerator())
					{
						while (enumerator.MoveNext())
						{
							enumerator.Current.IsChecked = value;
						}
						while (true)
						{
							switch (4)
							{
							case 0:
								break;
							default:
								goto end_IL_00fd;
							}
							continue;
							end_IL_00fd:
							break;
						}
					}
					sourceWorkbook = null;
				}
				A();
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			this.m_A = false;
			e.Handled = true;
			return;
		}
	}

	private void ShowAddMenu(object sender, RoutedEventArgs e)
	{
		System.Windows.Controls.ContextMenu contextMenu = btnAdd.ContextMenu;
		contextMenu.PlacementTarget = btnAdd;
		contextMenu.Placement = PlacementMode.Top;
		contextMenu.IsOpen = true;
		_ = null;
	}

	private void AddFiles(object sender, RoutedEventArgs e)
	{
		Microsoft.Office.Interop.Excel.Workbook workbook = null;
		OpenFileDialog openFileDialog = new OpenFileDialog();
		openFileDialog.DefaultExt = VH.A(176106);
		openFileDialog.Filter = VH.A(176115);
		openFileDialog.Title = VH.A(176230);
		openFileDialog.Multiselect = true;
		if (openFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
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
			List<string> list = openFileDialog.FileNames.ToList();
			if (this.m_A == null)
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
				this.m_A = new List<Microsoft.Office.Interop.Excel.Workbook>();
			}
			this.m_A = MH.A.Application;
			this.m_A.ScreenUpdating = false;
			this.m_A.EnableEvents = false;
			B();
			try
			{
				using (List<string>.Enumerator enumerator = list.GetEnumerator())
				{
					while (enumerator.MoveNext())
					{
						string current = enumerator.Current;
						if (workbook == null)
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
							workbook = this.m_A.Workbooks.Open(current, false, true, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), false, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
							this.m_A.Add(workbook);
						}
						A(workbook);
						workbook = null;
					}
					while (true)
					{
						switch (3)
						{
						case 0:
							break;
						default:
							goto end_IL_01b4;
						}
						continue;
						end_IL_01b4:
						break;
					}
				}
				ObservableCollection<SourceWorkbook> sourceWorkbooks = SourceWorkbooks;
				sourceWorkbooks[checked(sourceWorkbooks.Count - 1)].IsSelected = true;
				_ = null;
				C();
				E();
				F();
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				Forms.ErrorMessage(System.Windows.Window.GetWindow(this), ex2.Message);
				clsReporting.LogException(ex2);
				ProjectData.ClearProjectError();
			}
			this.m_A.ScreenUpdating = true;
			this.m_A.EnableEvents = true;
			list = null;
			Focus();
		}
		openFileDialog = null;
	}

	private void AddFolder(object sender, RoutedEventArgs e)
	{
		Microsoft.Office.Interop.Excel.Workbook workbook = null;
		base.Topmost = false;
		Microsoft.Office.Core.FileDialog fileDialog = ((_Application)this.m_A).get_FileDialog(MsoFileDialogType.msoFileDialogFolderPicker);
		fileDialog.Title = VH.A(176285);
		fileDialog.Filters.Clear();
		fileDialog.AllowMultiSelect = false;
		fileDialog.Show();
		FileDialogSelectedItems selectedItems = fileDialog.SelectedItems;
		if (selectedItems.Count > 0)
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
			if (this.m_A == null)
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
				this.m_A = new List<Microsoft.Office.Interop.Excel.Workbook>();
			}
			this.m_A = MH.A.Application;
			this.m_A.ScreenUpdating = false;
			this.m_A.EnableEvents = false;
			B();
			try
			{
				FileInfo[] files = new DirectoryInfo(selectedItems.Cast<object>().ElementAtOrDefault(0).ToString() + VH.A(48247)).GetFiles(VH.A(153184), SearchOption.TopDirectoryOnly);
				files.OrderBy([SpecialName] (FileInfo A) => A.Name);
				FileInfo[] array = files;
				IEnumerator enumerator = default(IEnumerator);
				foreach (FileInfo fileInfo in array)
				{
					try
					{
						enumerator = this.m_A.Workbooks.GetEnumerator();
						while (true)
						{
							if (enumerator.MoveNext())
							{
								Microsoft.Office.Interop.Excel.Workbook workbook2 = (Microsoft.Office.Interop.Excel.Workbook)enumerator.Current;
								if (Operators.CompareString(workbook2.FullName, fileInfo.FullName, TextCompare: false) == 0)
								{
									workbook = workbook2;
									break;
								}
								continue;
							}
							while (true)
							{
								switch (1)
								{
								case 0:
									break;
								default:
									goto end_IL_0187;
								}
								continue;
								end_IL_0187:
								break;
							}
							break;
						}
					}
					finally
					{
						if (enumerator is IDisposable)
						{
							while (true)
							{
								switch (7)
								{
								case 0:
									continue;
								}
								(enumerator as IDisposable).Dispose();
								break;
							}
						}
					}
					if (workbook == null)
					{
						workbook = this.m_A.Workbooks.Open(fileInfo.FullName, false, true, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), false, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
						this.m_A.Add(workbook);
					}
					A(workbook);
					workbook = null;
				}
				while (true)
				{
					switch (6)
					{
					case 0:
						continue;
					}
					ObservableCollection<SourceWorkbook> sourceWorkbooks = SourceWorkbooks;
					sourceWorkbooks[checked(sourceWorkbooks.Count - 1)].IsSelected = true;
					_ = null;
					break;
				}
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				Forms.ErrorMessage(System.Windows.Window.GetWindow(this), ex2.Message);
				clsReporting.LogException(ex2);
				ProjectData.ClearProjectError();
			}
			this.m_A.ScreenUpdating = true;
			this.m_A.EnableEvents = true;
			C();
			E();
			F();
		}
		_ = null;
		base.Topmost = true;
		Focus();
	}

	private void B()
	{
		if (this.m_A.Any())
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
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
			try
			{
				enumerator = this.m_A.Workbooks.GetEnumerator();
				while (enumerator.MoveNext())
				{
					Microsoft.Office.Interop.Excel.Workbook workbook = (Microsoft.Office.Interop.Excel.Workbook)enumerator.Current;
					if (workbook.Path.Length == 0)
					{
						workbook.Saved = false;
					}
				}
				return;
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
	}

	private void btnRemove_Click(object sender, RoutedEventArgs e)
	{
		if (!(trvSources.SelectedItem is SourceWorkbook))
		{
			return;
		}
		checked
		{
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
				SourceWorkbook sourceWorkbook = (SourceWorkbook)trvSources.SelectedItem;
				ObservableCollection<DestinationSheet> sheets = DestinationWorkbooks[0].Sheets;
				for (int i = sheets.Count - 1; i >= 0; i += -1)
				{
					if (sheets[i].Source.Parent != sourceWorkbook)
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
					sheets.RemoveAt(i);
				}
				while (true)
				{
					switch (4)
					{
					case 0:
						continue;
					}
					sheets = null;
					try
					{
						if (this.m_A.Contains(sourceWorkbook.Workbook))
						{
							this.m_A.Remove(sourceWorkbook.Workbook);
							Workbooks.ForceClose(sourceWorkbook.Workbook, false);
							Focus();
						}
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						ProjectData.ClearProjectError();
					}
					SourceWorkbooks.Remove(sourceWorkbook);
					sourceWorkbook = null;
					E();
					F();
					return;
				}
			}
		}
	}

	private void MoveSourceUp(object sender, RoutedEventArgs e)
	{
		A(-1);
	}

	private void MoveSourceDown(object sender, RoutedEventArgs e)
	{
		A(1);
	}

	private void A(int A)
	{
		if (trvSources.SelectedItem == null)
		{
			return;
		}
		checked
		{
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
				if (trvSources.SelectedItem is SourceSheet)
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
					SourceSheet A2;
					SourceSheet value;
					try
					{
						ObservableCollection<SourceSheet> sheets = ((SourceSheet)trvSources.SelectedItem).Parent.Sheets;
						int num = sheets.IndexOf((SourceSheet)trvSources.SelectedItem);
						A2 = sheets[num];
						value = sheets[num + A];
						sheets[num] = value;
						sheets[num + A] = A2;
						sheets = null;
						A2.IsSelected = true;
						ObservableCollection<DestinationSheet> sheets2 = DestinationWorkbooks[0].Sheets;
						DestinationSheet destinationSheet = sheets2.Where([SpecialName] (DestinationSheet destinationSheet2) => destinationSheet2.Source == A2).ElementAt(0);
						num = sheets2.IndexOf(destinationSheet);
						DestinationSheet value2 = sheets2[num + A];
						sheets2[num] = value2;
						sheets2[num + A] = destinationSheet;
						_ = null;
						destinationSheet = null;
						value2 = null;
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						ProjectData.ClearProjectError();
					}
					A2 = null;
					value = null;
				}
				else if (trvSources.SelectedItem is SourceWorkbook)
				{
					C();
					SourceWorkbook sourceWorkbook;
					SourceWorkbook value3;
					try
					{
						ObservableCollection<SourceWorkbook> sourceWorkbooks = SourceWorkbooks;
						int num2 = sourceWorkbooks.IndexOf((SourceWorkbook)trvSources.SelectedItem);
						sourceWorkbook = sourceWorkbooks[num2];
						value3 = sourceWorkbooks[num2 + A];
						sourceWorkbooks[num2] = value3;
						sourceWorkbooks[num2 + A] = sourceWorkbook;
						_ = null;
						sourceWorkbook.IsSelected = true;
						D();
					}
					catch (Exception ex3)
					{
						ProjectData.SetProjectError(ex3);
						Exception ex4 = ex3;
						ProjectData.ClearProjectError();
					}
					sourceWorkbook = null;
					value3 = null;
				}
				F();
				return;
			}
		}
	}

	private void SortWorkbooks(object sender, SelectionChangedEventArgs e)
	{
		ObservableCollection<SourceWorkbook> observableCollection = null;
		switch (((System.Windows.Controls.ComboBox)sender).SelectedIndex)
		{
		case 2:
			observableCollection = new ObservableCollection<SourceWorkbook>(SourceWorkbooks.OrderBy([SpecialName] (SourceWorkbook A) => A.Workbook.Name));
			break;
		case 3:
			observableCollection = new ObservableCollection<SourceWorkbook>(SourceWorkbooks.OrderByDescending([SpecialName] (SourceWorkbook A) => A.Workbook.Name));
			break;
		case 5:
			observableCollection = new ObservableCollection<SourceWorkbook>(SourceWorkbooks.OrderBy([SpecialName] (SourceWorkbook A) => A.Workbook.Path));
			break;
		case 6:
			observableCollection = new ObservableCollection<SourceWorkbook>(SourceWorkbooks.OrderByDescending([SpecialName] (SourceWorkbook A) => A.Workbook.Path));
			break;
		case 8:
		{
			ObservableCollection<SourceWorkbook> sourceWorkbooks2 = SourceWorkbooks;
			Func<SourceWorkbook, DateTime> keySelector2;
			if (_Closure_0024__.A == null)
			{
				keySelector2 = (_Closure_0024__.A = [SpecialName] (SourceWorkbook A) =>
				{
					if (A.Workbook.Path.Length > 0)
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
								return new FileInfo(A.Workbook.FullName).LastWriteTime;
							}
						}
					}
					return DateTime.MaxValue;
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
				keySelector2 = _Closure_0024__.A;
			}
			observableCollection = new ObservableCollection<SourceWorkbook>(sourceWorkbooks2.OrderBy(keySelector2));
			break;
		}
		case 9:
			observableCollection = new ObservableCollection<SourceWorkbook>(SourceWorkbooks.OrderByDescending([SpecialName] (SourceWorkbook A) => (A.Workbook.Path.Length > 0) ? new FileInfo(A.Workbook.FullName).LastWriteTime : DateTime.MaxValue));
			break;
		case 11:
		{
			ObservableCollection<SourceWorkbook> sourceWorkbooks3 = SourceWorkbooks;
			Func<SourceWorkbook, DateTime> keySelector3;
			if (_Closure_0024__.C == null)
			{
				keySelector3 = (_Closure_0024__.C = [SpecialName] (SourceWorkbook A) =>
				{
					if (A.Workbook.Path.Length > 0)
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
								return new FileInfo(A.Workbook.FullName).CreationTime;
							}
						}
					}
					return DateTime.MaxValue;
				});
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
				keySelector3 = _Closure_0024__.C;
			}
			observableCollection = new ObservableCollection<SourceWorkbook>(sourceWorkbooks3.OrderBy(keySelector3));
			break;
		}
		case 12:
		{
			ObservableCollection<SourceWorkbook> sourceWorkbooks = SourceWorkbooks;
			Func<SourceWorkbook, DateTime> keySelector;
			if (_Closure_0024__.D == null)
			{
				keySelector = (_Closure_0024__.D = [SpecialName] (SourceWorkbook A) => (A.Workbook.Path.Length > 0) ? new FileInfo(A.Workbook.FullName).CreationTime : DateTime.MaxValue);
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
				keySelector = _Closure_0024__.D;
			}
			observableCollection = new ObservableCollection<SourceWorkbook>(sourceWorkbooks.OrderByDescending(keySelector));
			break;
		}
		}
		if (observableCollection == null)
		{
			return;
		}
		SourceWorkbooks.Clear();
		IEnumerator<SourceWorkbook> enumerator = default(IEnumerator<SourceWorkbook>);
		try
		{
			enumerator = observableCollection.GetEnumerator();
			while (enumerator.MoveNext())
			{
				SourceWorkbook current = enumerator.Current;
				SourceWorkbooks.Add(current);
			}
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					goto end_IL_0290;
				}
				continue;
				end_IL_0290:
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
		D();
	}

	private void C()
	{
		cbxSort.SelectionChanged -= SortWorkbooks;
		cbxSort.SelectedIndex = 0;
		cbxSort.SelectionChanged += SortWorkbooks;
	}

	private void D()
	{
		DestinationWorkbooks[0].Sheets = new ObservableCollection<DestinationSheet>(DestinationWorkbooks[0].Sheets.OrderBy([SpecialName] (DestinationSheet A) => SourceWorkbooks.IndexOf(A.Source.Parent)));
	}

	private void E()
	{
		ObservableCollection<DestinationSheet> sheets = DestinationWorkbooks[0].Sheets;
		Func<DestinationSheet, bool> predicate;
		if (_Closure_0024__.A == null)
		{
			predicate = (_Closure_0024__.A = [SpecialName] (DestinationSheet A) => A.Visibility == Visibility.Visible);
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			predicate = _Closure_0024__.A;
		}
		int num = sheets.Where(predicate).Count();
		if (num != 1)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					tbCounter.Text = VH.A(39848) + num + VH.A(176328);
					return;
				}
			}
		}
		tbCounter.Text = VH.A(176363);
	}

	private void TargetWorkbookChanged(object sender, RoutedEventArgs e)
	{
		if (chkNew.IsChecked == true)
		{
			DestinationWorkbooks[0].Name = VH.A(176400);
		}
		else
		{
			DestinationWorkbooks[0].Name = ActiveWorkbook.Name;
		}
	}

	private void NameSelectionChanged(object sender, SelectionChangedEventArgs e)
	{
		F();
	}

	private void NameTextChanged(object sender, TextChangedEventArgs e)
	{
		F();
	}

	private void F()
	{
		string text = "";
		string text2 = "";
		string text3 = "";
		string text4 = "";
		int num = 1;
		System.Windows.Controls.ComboBox comboBox = cbxString1;
		int selectedIndex = comboBox.SelectedIndex;
		if (selectedIndex == -1)
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
			text = comboBox.Text;
		}
		comboBox = null;
		System.Windows.Controls.ComboBox comboBox2 = cbxString2;
		int selectedIndex2 = comboBox2.SelectedIndex;
		if (selectedIndex2 == -1)
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
			text2 = comboBox2.Text;
		}
		comboBox2 = null;
		System.Windows.Controls.ComboBox comboBox3 = cbxString3;
		int selectedIndex3 = comboBox3.SelectedIndex;
		if (selectedIndex3 == -1)
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
			text3 = comboBox3.Text;
		}
		comboBox3 = null;
		System.Windows.Controls.ComboBox comboBox4 = cbxString4;
		int selectedIndex4 = comboBox4.SelectedIndex;
		if (selectedIndex4 == -1)
		{
			text4 = comboBox4.Text;
		}
		comboBox4 = null;
		IEnumerator<DestinationSheet> enumerator = default(IEnumerator<DestinationSheet>);
		try
		{
			enumerator = DestinationWorkbooks[0].Sheets.GetEnumerator();
			while (enumerator.MoveNext())
			{
				DestinationSheet current = enumerator.Current;
				if (current.Visibility != Visibility.Visible)
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
				string text5 = "";
				SourceSheet source = current.Source;
				text5 = selectedIndex switch
				{
					-1 => text5 + text, 
					0 => Conversions.ToString(Operators.ConcatenateObject(text5, NewLateBinding.LateGet(source.Sheet, null, VH.A(19019), new object[0], null, null, null))), 
					1 => text5 + Path.GetFileNameWithoutExtension(source.Parent.Workbook.Name), 
					_ => text5 + num, 
				};
				text5 = selectedIndex2 switch
				{
					-1 => text5 + text2, 
					0 => Conversions.ToString(Operators.ConcatenateObject(text5, NewLateBinding.LateGet(source.Sheet, null, VH.A(19019), new object[0], null, null, null))), 
					1 => text5 + Path.GetFileNameWithoutExtension(source.Parent.Workbook.Name), 
					_ => text5 + num, 
				};
				text5 = selectedIndex3 switch
				{
					-1 => text5 + text3, 
					0 => Conversions.ToString(Operators.ConcatenateObject(text5, NewLateBinding.LateGet(source.Sheet, null, VH.A(19019), new object[0], null, null, null))), 
					1 => text5 + Path.GetFileNameWithoutExtension(source.Parent.Workbook.Name), 
					_ => text5 + num, 
				};
				text5 = selectedIndex4 switch
				{
					-1 => text5 + text4, 
					0 => Conversions.ToString(Operators.ConcatenateObject(text5, NewLateBinding.LateGet(source.Sheet, null, VH.A(19019), new object[0], null, null, null))), 
					1 => text5 + Path.GetFileNameWithoutExtension(source.Parent.Workbook.Name), 
					_ => text5 + num, 
				};
				if (text5.Length == 0)
				{
					while (true)
					{
						switch (7)
						{
						case 0:
							break;
						default:
							return;
						}
					}
				}
				current.Name = text5;
				num = checked(num + 1);
				source = null;
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
					switch (6)
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

	private void btnOk_Click(object sender, RoutedEventArgs e)
	{
		base.DialogResult = true;
	}

	private void wpfMerge_Closing(object sender, CancelEventArgs e)
	{
		bool flag = false;
		int num = 1;
		int num2 = 0;
		List<string> list = new List<string>();
		Microsoft.Office.Interop.Excel.Application a = this.m_A;
		a.ScreenUpdating = false;
		a.EnableEvents = false;
		bool copyObjectsWithCells = a.CopyObjectsWithCells;
		a.CopyObjectsWithCells = true;
		_ = null;
		checked
		{
			Microsoft.Office.Interop.Excel.Workbook workbook;
			if (base.DialogResult == true)
			{
				try
				{
					IEnumerator<DestinationSheet> enumerator = default(IEnumerator<DestinationSheet>);
					try
					{
						enumerator = DestinationWorkbooks[0].Sheets.GetEnumerator();
						while (true)
						{
							if (enumerator.MoveNext())
							{
								DestinationSheet current = enumerator.Current;
								if (current.Visibility == Visibility.Visible)
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
									if (list.Contains(current.Name))
									{
										Forms.WarningMessage(VH.A(176429));
										e.Cancel = true;
										break;
									}
									if (current.Name.Length > 31)
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
										current.Name = Strings.Left(current.Name, 31);
										flag = true;
									}
									list.Add(current.Name);
								}
								current = null;
								continue;
							}
							while (true)
							{
								switch (2)
								{
								case 0:
									break;
								default:
									goto end_IL_011c;
								}
								continue;
								end_IL_011c:
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
								switch (6)
								{
								case 0:
									continue;
								}
								enumerator.Dispose();
								break;
							}
						}
					}
					list = null;
					if (!e.Cancel)
					{
						while (true)
						{
							switch (4)
							{
							case 0:
								continue;
							}
							if (flag)
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
								Forms.WarningMessage(VH.A(176508));
							}
							if (chkNew.IsChecked == true)
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
								workbook = this.m_A.Workbooks.Add(RuntimeHelpers.GetObjectValue(Missing.Value));
								num = workbook.Sheets.Count;
								num2 = num;
							}
							else
							{
								workbook = ActiveWorkbook;
								num = workbook.Sheets.Count;
							}
							foreach (DestinationSheet sheet in DestinationWorkbooks[0].Sheets)
							{
								if (sheet.Visibility == Visibility.Visible)
								{
									sheet.Source.Parent.Workbook.Activate();
									NewLateBinding.LateCall(sheet.Source.Sheet, null, VH.A(224), new object[1] { workbook.Sheets[num] }, new string[1] { VH.A(80163) }, null, null, IgnoreReturn: true);
									num++;
									try
									{
										NewLateBinding.LateSetComplex(workbook.Sheets[workbook.Sheets.Count], null, VH.A(19019), new object[1] { sheet.Name }, null, null, OptimisticSet: false, RValueBase: true);
									}
									catch (Exception ex)
									{
										ProjectData.SetProjectError(ex);
										Exception ex2 = ex;
										ProjectData.ClearProjectError();
									}
								}
							}
							int num3 = num2;
							for (int i = 1; i <= num3; i++)
							{
								try
								{
									NewLateBinding.LateCall(workbook.Sheets[1], null, VH.A(60691), new object[0], null, null, null, IgnoreReturn: true);
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
									continue;
								}
								clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)8, VH.A(176708));
								break;
							}
							break;
						}
					}
				}
				catch (Exception ex5)
				{
					ProjectData.SetProjectError(ex5);
					Exception ex6 = ex5;
					Forms.ErrorMessage(ex6.Message);
					clsReporting.LogException(ex6);
					ProjectData.ClearProjectError();
				}
			}
			if (!e.Cancel)
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
				if (this.m_A != null)
				{
					this.m_A.DisplayAlerts = false;
					foreach (Microsoft.Office.Interop.Excel.Workbook item in this.m_A)
					{
						try
						{
							item.Close(false, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
						}
						catch (Exception ex7)
						{
							ProjectData.SetProjectError(ex7);
							Exception ex8 = ex7;
							ProjectData.ClearProjectError();
						}
					}
					this.m_A.DisplayAlerts = true;
				}
				this.m_A.CopyObjectsWithCells = copyObjectsWithCells;
				this.m_A.ScreenUpdating = true;
				this.m_A.EnableEvents = true;
				SourceWorkbooks = null;
				this.m_A = null;
				cbxSort.SelectionChanged -= SortWorkbooks;
				cbxString1.SelectionChanged -= NameSelectionChanged;
				cbxString2.SelectionChanged -= NameSelectionChanged;
				cbxString3.SelectionChanged -= NameSelectionChanged;
				cbxString4.SelectionChanged -= NameSelectionChanged;
			}
			else
			{
				this.m_A.CopyObjectsWithCells = copyObjectsWithCells;
				this.m_A.ScreenUpdating = true;
				this.m_A.EnableEvents = true;
			}
			workbook = null;
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
			switch (2)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			this.m_C = true;
			Uri resourceLocator = new Uri(VH.A(176739), UriKind.Relative);
			System.Windows.Application.LoadComponent(this, resourceLocator);
			return;
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
				switch (7)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					tbCounter = (TextBlock)target;
					return;
				}
			}
		}
		if (connectionId == 3)
		{
			trvSources = (System.Windows.Controls.TreeView)target;
			return;
		}
		if (connectionId == 6)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					btnAdd = (System.Windows.Controls.Button)target;
					btnAdd.Click += ShowAddMenu;
					return;
				}
			}
		}
		if (connectionId == 7)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					btnAddFiles = (System.Windows.Controls.MenuItem)target;
					btnAddFiles.Click += AddFiles;
					return;
				}
			}
		}
		if (connectionId == 8)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					imgAddFiles = (Image)target;
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
					btnAddFolder = (System.Windows.Controls.MenuItem)target;
					btnAddFolder.Click += AddFolder;
					return;
				}
			}
		}
		if (connectionId == 10)
		{
			imgAddFolder = (Image)target;
			return;
		}
		if (connectionId == 11)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					btnRemove = (System.Windows.Controls.Button)target;
					return;
				}
			}
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
					btnWorkbooksSort = (System.Windows.Controls.Button)target;
					return;
				}
			}
		}
		if (connectionId == 13)
		{
			cbxSort = (System.Windows.Controls.ComboBox)target;
			return;
		}
		if (connectionId == 14)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					btnDown = (System.Windows.Controls.Button)target;
					return;
				}
			}
		}
		if (connectionId == 15)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					btnUp = (System.Windows.Controls.Button)target;
					return;
				}
			}
		}
		if (connectionId == 16)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					chkNew = (System.Windows.Controls.CheckBox)target;
					chkNew.Checked += TargetWorkbookChanged;
					chkNew.Unchecked += TargetWorkbookChanged;
					return;
				}
			}
		}
		if (connectionId == 17)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					trvDestination = (System.Windows.Controls.TreeView)target;
					return;
				}
			}
		}
		if (connectionId == 18)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					cbxString1 = (System.Windows.Controls.ComboBox)target;
					cbxString1.AddHandler(System.Windows.Controls.Primitives.TextBoxBase.TextChangedEvent, new TextChangedEventHandler(NameTextChanged));
					return;
				}
			}
		}
		if (connectionId == 19)
		{
			cbxString2 = (System.Windows.Controls.ComboBox)target;
			cbxString2.AddHandler(System.Windows.Controls.Primitives.TextBoxBase.TextChangedEvent, new TextChangedEventHandler(NameTextChanged));
			return;
		}
		if (connectionId == 20)
		{
			cbxString3 = (System.Windows.Controls.ComboBox)target;
			cbxString3.AddHandler(System.Windows.Controls.Primitives.TextBoxBase.TextChangedEvent, new TextChangedEventHandler(NameTextChanged));
			return;
		}
		if (connectionId == 21)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					cbxString4 = (System.Windows.Controls.ComboBox)target;
					cbxString4.AddHandler(System.Windows.Controls.Primitives.TextBoxBase.TextChangedEvent, new TextChangedEventHandler(NameTextChanged));
					return;
				}
			}
		}
		if (connectionId == 22)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					btnOk = (System.Windows.Controls.Button)target;
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

	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	[EditorBrowsable(EditorBrowsableState.Never)]
	[DebuggerNonUserCode]
	public void System_Windows_Markup_IStyleConnector_Connect(int connectionId, object target)
	{
		if (connectionId == 1)
		{
			EventSetter eventSetter = new EventSetter();
			eventSetter.Event = FrameworkElement.RequestBringIntoViewEvent;
			eventSetter.Handler = new RequestBringIntoViewEventHandler(OnRequestBringIntoView);
			((System.Windows.Style)target).Setters.Add(eventSetter);
			eventSetter = new EventSetter();
			eventSetter.Event = TreeViewItem.SelectedEvent;
			eventSetter.Handler = new RoutedEventHandler(OnSelected);
			((System.Windows.Style)target).Setters.Add(eventSetter);
		}
		if (connectionId == 4)
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
			((System.Windows.Controls.CheckBox)target).Checked += WorkbookCheckedChanged;
			((System.Windows.Controls.CheckBox)target).Unchecked += WorkbookCheckedChanged;
		}
		if (connectionId != 5)
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
			((System.Windows.Controls.CheckBox)target).Checked += SheetCheckedChanged;
			((System.Windows.Controls.CheckBox)target).Unchecked += SheetCheckedChanged;
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
	private int A(DestinationSheet A)
	{
		return SourceWorkbooks.IndexOf(A.Source.Parent);
	}
}
