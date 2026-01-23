using System;
using System.CodeDom.Compiler;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using A;
using MacabacusMacros;
using MacabacusMacros.Explorer;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.Links;
using PowerPointAddIn1.Pagination;
using PowerPointAddIn1.Publishing;
using PowerPointAddIn1.Publishing.Share;
using PowerPointAddIn1.Shapes;

namespace PowerPointAddIn1.Explorer;

[DesignerGenerated]
public sealed class wpfExplorer : System.Windows.Controls.UserControl, INotifyPropertyChanged, IComponentConnector, IStyleConnector
{
	[CompilerGenerated]
	internal sealed class MF
	{
		public BaseItem A;

		public BitmapSource A;

		public wpfExplorer A;

		[SpecialName]
		internal void A()
		{
			Preview.PrepareWindow(this.A.PreviewWindow, this.A, this.A, this.A.PointToScreen(new System.Windows.Point(0.0, 0.0)));
		}
	}

	[Serializable]
	[CompilerGenerated]
	internal sealed class _Closure_0024__66<_0024CLS0>
	{
		public static readonly _Closure_0024__66<_0024CLS0> A;

		public static Func<ContentItem, bool> A;

		static _Closure_0024__66()
		{
			_Closure_0024__66<_0024CLS0>.A = new _Closure_0024__66<_0024CLS0>();
		}

		[SpecialName]
		internal bool A(ContentItem A)
		{
			return A is _0024CLS0;
		}
	}

	[CompilerGenerated]
	internal sealed class NF
	{
		public ObservableCollection<PresentationItem> A;

		public bool A;

		public wpfExplorer A;

		public NF(NF A)
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
				this.A = A.A;
				return;
			}
		}

		[SpecialName]
		internal void A()
		{
			IEnumerator<PresentationItem> enumerator = default(IEnumerator<PresentationItem>);
			try
			{
				enumerator = this.A.GetEnumerator();
				while (enumerator.MoveNext())
				{
					PresentationItem current = enumerator.Current;
					((BaseItem)current).IsExpanded = this.A;
					foreach (SlideItem slide in current.Slides)
					{
						((BaseItem)slide).IsExpanded = this.A;
					}
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
			this.A = null;
			this.A.Dispatcher.BeginInvoke((B)([SpecialName] () =>
			{
				this.A.H();
			}));
		}
	}

	[CompilerGenerated]
	private PropertyChangedEventHandler m_A;

	private static FieldInfo m_A;

	private Microsoft.Office.Interop.PowerPoint.Application m_A;

	private DispatcherTimer m_A;

	private DispatcherTimer m_B;

	private RoutedPropertyChangedEventHandler<object> m_A;

	private DispatcherFrame m_A;

	private object m_A;

	[CompilerGenerated]
	private ctpExplorer2 m_A;

	[CompilerGenerated]
	private wpfPreview m_A;

	private ICollectionView m_A;

	private ObservableCollection<PresentationItem> m_A;

	[CompilerGenerated]
	private Microsoft.Office.Interop.PowerPoint.Presentation m_A;

	[CompilerGenerated]
	private PresentationItem m_A;

	private double m_A;

	private bool m_A;

	private ScrollViewer m_A;

	private bool m_B;

	private System.Windows.Point m_A;

	private System.Windows.Point m_B;

	private bool m_C;

	[CompilerGenerated]
	[AccessedThroughProperty("chkSettings")]
	private System.Windows.Controls.CheckBox m_A;

	[AccessedThroughProperty("chkPreview")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox m_B;

	[CompilerGenerated]
	[AccessedThroughProperty("chkShowAll")]
	private System.Windows.Controls.CheckBox m_C;

	[CompilerGenerated]
	[AccessedThroughProperty("btnRefresh")]
	private System.Windows.Controls.Button m_A;

	[AccessedThroughProperty("btnExpand")]
	[CompilerGenerated]
	private System.Windows.Controls.Button m_B;

	[CompilerGenerated]
	[AccessedThroughProperty("btnCollapse")]
	private System.Windows.Controls.Button m_C;

	[AccessedThroughProperty("btnSpacing")]
	[CompilerGenerated]
	private System.Windows.Controls.Button m_D;

	[AccessedThroughProperty("popSettings")]
	[CompilerGenerated]
	private Popup m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("chkCharts")]
	private System.Windows.Controls.CheckBox m_D;

	[CompilerGenerated]
	[AccessedThroughProperty("chkTables")]
	private System.Windows.Controls.CheckBox m_E;

	[AccessedThroughProperty("chkExcel")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox m_F;

	[CompilerGenerated]
	[AccessedThroughProperty("chkWord")]
	private System.Windows.Controls.CheckBox m_G;

	[CompilerGenerated]
	[AccessedThroughProperty("chkSmartArt")]
	private System.Windows.Controls.CheckBox m_H;

	[CompilerGenerated]
	[AccessedThroughProperty("chkImages")]
	private System.Windows.Controls.CheckBox m_I;

	[CompilerGenerated]
	[AccessedThroughProperty("chkMedia")]
	private System.Windows.Controls.CheckBox m_J;

	[CompilerGenerated]
	[AccessedThroughProperty("chkInk")]
	private System.Windows.Controls.CheckBox m_K;

	[AccessedThroughProperty("chkHyperlinks")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox m_L;

	[CompilerGenerated]
	[AccessedThroughProperty("chkComments")]
	private System.Windows.Controls.CheckBox m_M;

	[AccessedThroughProperty("chkNotes")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox m_N;

	[CompilerGenerated]
	[AccessedThroughProperty("trvExplorer")]
	private System.Windows.Controls.TreeView m_A;

	[AccessedThroughProperty("txtSearch")]
	[CompilerGenerated]
	private System.Windows.Controls.TextBox m_A;

	private bool m_D;

	public ctpExplorer2 thisCTP
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

	public ICollectionView SourceCollection
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(AH.A(10961));
		}
	}

	private ObservableCollection<PresentationItem> SinglePresentation
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(AH.A(116267));
		}
	}

	private Microsoft.Office.Interop.PowerPoint.Presentation ThisPresentation
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

	public PresentationItem ThisPresentationItem
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

	internal virtual System.Windows.Controls.CheckBox chkSettings
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

	internal virtual System.Windows.Controls.CheckBox chkPreview
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

	internal virtual System.Windows.Controls.CheckBox chkShowAll
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

	internal virtual System.Windows.Controls.Button btnRefresh
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
			RoutedEventHandler value2 = btnRefresh_Click;
			System.Windows.Controls.Button button = this.m_A;
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
			this.m_A = value;
			button = this.m_A;
			if (button != null)
			{
				button.Click += value2;
			}
		}
	}

	internal virtual System.Windows.Controls.Button btnExpand
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
			RoutedEventHandler value2 = ExpandAll;
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

	internal virtual System.Windows.Controls.Button btnCollapse
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
			RoutedEventHandler value2 = CollapseAll;
			System.Windows.Controls.Button button = this.m_C;
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
			this.m_C = value;
			button = this.m_C;
			if (button != null)
			{
				button.Click += value2;
			}
		}
	}

	internal virtual System.Windows.Controls.Button btnSpacing
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

	internal virtual System.Windows.Controls.CheckBox chkCharts
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

	internal virtual System.Windows.Controls.CheckBox chkTables
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

	internal virtual System.Windows.Controls.CheckBox chkExcel
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

	internal virtual System.Windows.Controls.CheckBox chkWord
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

	internal virtual System.Windows.Controls.CheckBox chkSmartArt
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
			this.m_H = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkImages
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
			this.m_I = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkMedia
	{
		[CompilerGenerated]
		get
		{
			return this.m_J;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_J = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkInk
	{
		[CompilerGenerated]
		get
		{
			return this.m_K;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_K = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkHyperlinks
	{
		[CompilerGenerated]
		get
		{
			return this.m_L;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_L = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkComments
	{
		[CompilerGenerated]
		get
		{
			return this.m_M;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_M = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkNotes
	{
		[CompilerGenerated]
		get
		{
			return this.m_N;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_N = value;
		}
	}

	internal virtual System.Windows.Controls.TreeView trvExplorer
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
			MouseButtonEventHandler value4 = treeView_PreviewMouseLeftButtonDown;
			System.Windows.Input.MouseEventHandler value5 = treeView_PreviewMouseMove;
			System.Windows.DragEventHandler value6 = treeView_DragEnter;
			System.Windows.DragEventHandler value7 = treeView_Drop;
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
				treeView.PreviewMouseLeftButtonDown -= value4;
				treeView.PreviewMouseMove -= value5;
				treeView.DragEnter -= value6;
				treeView.Drop -= value7;
			}
			this.m_A = value;
			treeView = this.m_A;
			if (treeView == null)
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
				treeView.PreviewKeyDown += value2;
				treeView.PreviewMouseWheel += value3;
				treeView.PreviewMouseLeftButtonDown += value4;
				treeView.PreviewMouseMove += value5;
				treeView.DragEnter += value6;
				treeView.Drop += value7;
				return;
			}
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
			this.m_A = value;
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
		}
	}

	public wpfExplorer()
	{
		base.Loaded += wpfExplorer_Loaded;
		base.Unloaded += wpfExplorer_Unloaded;
		this.m_A = null;
		this.m_B = null;
		this.m_A = [SpecialName] (object a0, RoutedPropertyChangedEventArgs<object> a1) =>
		{
			TreeViewSelectionChanged((System.Windows.Controls.TreeView)a0, a1);
		};
		PreviewWindow = null;
		this.m_A = null;
		this.m_A = null;
		this.m_A = 0.0;
		this.m_A = false;
		this.m_A = null;
		this.m_C = false;
		InitializeComponent();
		A();
		this.m_A = NG.A.Application;
		MySettings settings = PB.Settings;
		Options.ShowCharts = settings.ExplorerShowCharts;
		Options.ShowTables = settings.ExplorerShowTables;
		Options.ShowEmbeddedExcel = settings.ExplorerShowEmbeddedExcel;
		Options.ShowEmbeddedWord = settings.ExplorerShowEmbeddedWord;
		Options.ShowSmartArt = settings.ExplorerShowSmartArt;
		Options.ShowImages = settings.ExplorerShowImages;
		Options.ShowMedia = settings.ExplorerShowMedia;
		Options.ShowInk = settings.ExplorerShowInk;
		Options.ShowComments = settings.ExplorerShowComments;
		Options.ShowNotes = settings.ExplorerShowNotes;
		Options.ShowHyperlinks = settings.ExplorerShowHyperlinks;
		Options.ShowPreviews = settings.ExplorerPreviews;
		Options.ShowAll = settings.ExplorerShowAllPresentations;
		_ = null;
		AddFilterHandlers();
	}

	private void A(string A)
	{
		this.m_A?.Invoke(this, new PropertyChangedEventArgs(A));
	}

	private static void A()
	{
		wpfExplorer.m_A = typeof(SystemParameters).GetField(AH.A(57499), BindingFlags.Static | BindingFlags.NonPublic);
		B();
		SystemParameters.StaticPropertyChanged += SystemParameters_StaticPropertyChanged;
	}

	private static void SystemParameters_StaticPropertyChanged(object sender, PropertyChangedEventArgs e)
	{
		B();
	}

	private static void B()
	{
		if (!SystemParameters.MenuDropAlignment || (object)wpfExplorer.m_A == null)
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
			wpfExplorer.m_A.SetValue(null, false);
			return;
		}
	}

	public void Initialize(Microsoft.Office.Interop.PowerPoint.Presentation pres)
	{
		ThisPresentation = pres;
	}

	private void wpfExplorer_Loaded(object sender, RoutedEventArgs e)
	{
		trvExplorer.LostKeyboardFocus += TreeViewLostKeyboardFocus;
		trvExplorer.GotKeyboardFocus += TreeViewGotKeyboardFocus;
		chkShowAll.Checked += ShowAllToggle;
		chkShowAll.Unchecked += ShowAllToggle;
		chkPreview.Checked += PreviewToggle;
		chkPreview.Unchecked += PreviewToggle;
		txtSearch.TextChanged += SearchTextChanged;
		H();
	}

	private void wpfExplorer_Unloaded(object sender, RoutedEventArgs e)
	{
		trvExplorer.LostKeyboardFocus -= TreeViewLostKeyboardFocus;
		trvExplorer.GotKeyboardFocus -= TreeViewGotKeyboardFocus;
		chkShowAll.Checked -= ShowAllToggle;
		chkShowAll.Unchecked -= ShowAllToggle;
		chkPreview.Checked -= PreviewToggle;
		chkPreview.Unchecked -= PreviewToggle;
		txtSearch.TextChanged -= SearchTextChanged;
		RemoveTreeViewSelectionChangedHandler();
		E();
		this.m_A = null;
		this.m_A = null;
		PreviewWindow = null;
		this.m_A = null;
		this.m_B = null;
		this.m_A = null;
		this.m_A = null;
		this.m_A = null;
		SinglePresentation = null;
		ThisPresentation = null;
	}

	private void btnRefresh_Click(object sender, RoutedEventArgs e)
	{
		K();
		if (chkShowAll.IsChecked == true)
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
			Pane.LoadPresentations(this.m_A);
		}
		else
		{
			SinglePresentation.Clear();
			D();
		}
		trvExplorer.Focus();
	}

	private void ShowAllToggle(object sender, RoutedEventArgs e)
	{
		base.Dispatcher.BeginInvoke((B)([SpecialName] () =>
		{
			bool value = chkShowAll.IsChecked.Value;
			PB.Settings.ExplorerShowAllPresentations = value;
			RemoveTreeViewSelectionChangedHandler();
			if (value)
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
				Pane.LoadPresentations(this.m_A);
				A(Pane.AllPresentations);
				C();
			}
			else
			{
				if (SinglePresentation == null)
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
					SinglePresentation = new ObservableCollection<PresentationItem>();
				}
				else
				{
					SinglePresentation.Clear();
				}
				D();
				A(SinglePresentation);
			}
			H();
			trvExplorer.Focus();
		}), DispatcherPriority.Normal);
	}

	public void Connect()
	{
		base.Dispatcher.BeginInvoke((B)([SpecialName] () =>
		{
			Events.Reset(this.m_A);
			RemoveTreeViewSelectionChangedHandler();
			if (chkShowAll.IsChecked == true)
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
				A(Pane.AllPresentations);
				C();
			}
			else
			{
				if (SinglePresentation == null)
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
					SinglePresentation = new ObservableCollection<PresentationItem>();
				}
				else
				{
					SinglePresentation.Clear();
				}
				D();
				A(SinglePresentation);
			}
			H();
		}), DispatcherPriority.Normal);
	}

	private void A(ObservableCollection<PresentationItem> A)
	{
		SourceCollection = CollectionViewSource.GetDefaultView(A);
	}

	private void C()
	{
		ThisPresentationItem = Pane.AllPresentations.Where([SpecialName] (PresentationItem A) => A.Presentation == ThisPresentation).ElementAtOrDefault(0);
	}

	private void D()
	{
		ThisPresentationItem = new PresentationItem(ThisPresentation);
		SinglePresentation.Add(ThisPresentationItem);
	}

	private bool A(object A)
	{
		return true;
	}

	private void SettingsPopupOpened(object sender, EventArgs e)
	{
		chkSettings.IsHitTestVisible = false;
		chkTables.Focus();
	}

	private void SettingsPopupClosed(object sender, EventArgs e)
	{
		chkSettings.IsChecked = false;
		chkSettings.IsHitTestVisible = true;
	}

	private void CloseSettingsPopup(object sender, System.Windows.Input.KeyEventArgs e)
	{
		if (e.Key == Key.Escape)
		{
			chkSettings.IsChecked = false;
			e.Handled = true;
		}
	}

	public void AddFilterHandlers()
	{
		chkCharts.Checked += ToggleCharts;
		chkCharts.Unchecked += ToggleCharts;
		chkTables.Checked += ToggleTables;
		chkTables.Unchecked += ToggleTables;
		chkExcel.Checked += ToggleExcel;
		chkExcel.Unchecked += ToggleExcel;
		chkWord.Checked += ToggleWord;
		chkWord.Unchecked += ToggleWord;
		chkSmartArt.Checked += ToggleSmartArt;
		chkSmartArt.Unchecked += ToggleSmartArt;
		chkImages.Checked += ToggleImages;
		chkImages.Unchecked += ToggleImages;
		chkMedia.Checked += ToggleMedia;
		chkMedia.Unchecked += ToggleMedia;
		chkInk.Checked += ToggleInk;
		chkInk.Unchecked += ToggleInk;
		chkComments.Checked += ToggleComments;
		chkComments.Unchecked += ToggleComments;
		chkNotes.Checked += ToggleNotes;
		chkNotes.Unchecked += ToggleNotes;
		chkHyperlinks.Checked += ToggleHyperlinks;
		chkHyperlinks.Unchecked += ToggleHyperlinks;
	}

	private void E()
	{
		chkCharts.Checked -= ToggleCharts;
		chkCharts.Unchecked -= ToggleCharts;
		chkTables.Checked -= ToggleTables;
		chkTables.Unchecked -= ToggleTables;
		chkExcel.Checked -= ToggleExcel;
		chkExcel.Unchecked -= ToggleExcel;
		chkWord.Checked -= ToggleWord;
		chkWord.Unchecked -= ToggleWord;
		chkSmartArt.Checked -= ToggleSmartArt;
		chkSmartArt.Unchecked -= ToggleSmartArt;
		chkImages.Checked -= ToggleImages;
		chkImages.Unchecked -= ToggleImages;
		chkMedia.Checked -= ToggleMedia;
		chkMedia.Unchecked -= ToggleMedia;
		chkInk.Checked -= ToggleInk;
		chkInk.Unchecked -= ToggleInk;
		chkComments.Checked -= ToggleComments;
		chkComments.Unchecked -= ToggleComments;
		chkNotes.Checked -= ToggleNotes;
		chkNotes.Unchecked -= ToggleNotes;
		chkHyperlinks.Checked -= ToggleHyperlinks;
		chkHyperlinks.Unchecked -= ToggleHyperlinks;
	}

	private void ToggleCharts(object sender, RoutedEventArgs e)
	{
		if (!popSettings.IsOpen)
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
			bool value = chkCharts.IsChecked.Value;
			PB.Settings.ExplorerShowCharts = value;
			A<ChartItem>(value);
			return;
		}
	}

	private void ToggleTables(object sender, RoutedEventArgs e)
	{
		if (!popSettings.IsOpen)
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
			bool value = chkTables.IsChecked.Value;
			PB.Settings.ExplorerShowTables = value;
			A<TableItem>(value);
			return;
		}
	}

	private void ToggleExcel(object sender, RoutedEventArgs e)
	{
		if (!popSettings.IsOpen)
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
			bool value = chkExcel.IsChecked.Value;
			PB.Settings.ExplorerShowEmbeddedExcel = value;
			A<EmbeddedExcelItem>(value);
			return;
		}
	}

	private void ToggleWord(object sender, RoutedEventArgs e)
	{
		if (popSettings.IsOpen)
		{
			bool value = chkWord.IsChecked.Value;
			PB.Settings.ExplorerShowEmbeddedWord = value;
			A<EmbeddedWordItem>(value);
		}
	}

	private void ToggleSmartArt(object sender, RoutedEventArgs e)
	{
		if (popSettings.IsOpen)
		{
			bool value = chkSmartArt.IsChecked.Value;
			PB.Settings.ExplorerShowSmartArt = value;
			A<SmartArtItem>(value);
		}
	}

	private void ToggleImages(object sender, RoutedEventArgs e)
	{
		if (!popSettings.IsOpen)
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
			bool value = chkImages.IsChecked.Value;
			PB.Settings.ExplorerShowImages = value;
			A<ImageItem>(value);
			return;
		}
	}

	private void ToggleMedia(object sender, RoutedEventArgs e)
	{
		if (!popSettings.IsOpen)
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
			bool value = chkMedia.IsChecked.Value;
			PB.Settings.ExplorerShowMedia = value;
			A<MediaItem>(value);
			return;
		}
	}

	private void ToggleInk(object sender, RoutedEventArgs e)
	{
		if (!popSettings.IsOpen)
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
			bool value = chkInk.IsChecked.Value;
			PB.Settings.ExplorerShowInk = value;
			A<InkItem>(value);
			return;
		}
	}

	private void ToggleComments(object sender, RoutedEventArgs e)
	{
		if (!popSettings.IsOpen)
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
			bool value = chkComments.IsChecked.Value;
			PB.Settings.ExplorerShowComments = value;
			A<CommentItem>(value);
			return;
		}
	}

	private void ToggleNotes(object sender, RoutedEventArgs e)
	{
		if (!popSettings.IsOpen)
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
			bool value = chkNotes.IsChecked.Value;
			PB.Settings.ExplorerShowNotes = value;
			A<NotesItem>(value);
			return;
		}
	}

	private void ToggleHyperlinks(object sender, RoutedEventArgs e)
	{
		if (!popSettings.IsOpen)
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
			bool value = chkHyperlinks.IsChecked.Value;
			PB.Settings.ExplorerShowHyperlinks = value;
			A<HyperlinkItem>(value);
			return;
		}
	}

	private void A<A>(bool A)
	{
		ObservableCollection<PresentationItem> observableCollection = this.A();
		checked
		{
			if (A)
			{
				using IEnumerator<PresentationItem> enumerator = observableCollection.GetEnumerator();
				IEnumerator<SlideItem> enumerator2 = default(IEnumerator<SlideItem>);
				while (enumerator.MoveNext())
				{
					PresentationItem current = enumerator.Current;
					try
					{
						SlideItem current2;
						for (enumerator2 = current.Slides.GetEnumerator(); enumerator2.MoveNext(); current2 = null)
						{
							current2 = enumerator2.Current;
							if (!((BaseItem)current2).IsExpanded)
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
								if (!current2.IsPopulated())
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
							}
							current2.Populate();
						}
						while (true)
						{
							switch (5)
							{
							case 0:
								break;
							default:
								goto end_IL_0083;
							}
							continue;
							end_IL_0083:
							break;
						}
					}
					finally
					{
						if (enumerator2 != null)
						{
							while (true)
							{
								switch (7)
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
				while (true)
				{
					switch (1)
					{
					case 0:
						break;
					default:
						goto end_IL_00b0;
					}
					continue;
					end_IL_00b0:
					break;
				}
			}
			else
			{
				IEnumerable<ContentItem> source;
				using (IEnumerator<PresentationItem> enumerator3 = observableCollection.GetEnumerator())
				{
					IEnumerator<SlideItem> enumerator4 = default(IEnumerator<SlideItem>);
					while (enumerator3.MoveNext())
					{
						PresentationItem current3 = enumerator3.Current;
						try
						{
							enumerator4 = current3.Slides.GetEnumerator();
							while (enumerator4.MoveNext())
							{
								SlideItem current4 = enumerator4.Current;
								ObservableCollection<ContentItem> children = current4.Children;
								Func<ContentItem, bool> predicate;
								if (_Closure_0024__66<A>.A == null)
								{
									predicate = (_Closure_0024__66<A>.A = [SpecialName] (ContentItem contentItem) => contentItem is A);
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
									predicate = _Closure_0024__66<A>.A;
								}
								source = children.Where(predicate);
								for (int num = source.Count() - 1; num >= 0; num += -1)
								{
									current4.RemoveChild(source.ElementAtOrDefault(num));
								}
							}
							while (true)
							{
								switch (2)
								{
								case 0:
									break;
								default:
									goto end_IL_017c;
								}
								continue;
								end_IL_017c:
								break;
							}
						}
						finally
						{
							if (enumerator4 != null)
							{
								while (true)
								{
									switch (7)
									{
									case 0:
										continue;
									}
									enumerator4.Dispose();
									break;
								}
							}
						}
					}
					while (true)
					{
						switch (2)
						{
						case 0:
							break;
						default:
							goto end_IL_01ac;
						}
						continue;
						end_IL_01ac:
						break;
					}
				}
				source = null;
			}
			observableCollection = null;
		}
	}

	private void F()
	{
	}

	private void ClearSearch(object sender, RoutedEventArgs e)
	{
		txtSearch.TextChanged -= SearchTextChanged;
		txtSearch.Clear();
		txtSearch.TextChanged += SearchTextChanged;
		G();
	}

	private void SearchTextChanged(object sender, TextChangedEventArgs e)
	{
		if (!txtSearch.IsKeyboardFocused)
		{
			return;
		}
		IEnumerator<SlideItem> enumerator2 = default(IEnumerator<SlideItem>);
		IEnumerator<ContentItem> enumerator3 = default(IEnumerator<ContentItem>);
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
			string text = txtSearch.Text;
			if (text.Length > 0)
			{
				text = text.ToLower();
				ObservableCollection<PresentationItem> observableCollection = A();
				foreach (PresentationItem item in observableCollection)
				{
					try
					{
						enumerator2 = item.Slides.GetEnumerator();
						while (enumerator2.MoveNext())
						{
							SlideItem current2 = enumerator2.Current;
							current2.Search(text);
							if (!(current2 is SlideItem))
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
							SlideItem slideItem = current2;
							if (!((BaseItem)slideItem).IsExpanded)
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
								if (!slideItem.IsPopulated())
								{
									goto IL_011e;
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
							try
							{
								enumerator3 = slideItem.Children.GetEnumerator();
								while (enumerator3.MoveNext())
								{
									enumerator3.Current.Search(text);
								}
								while (true)
								{
									switch (6)
									{
									case 0:
										break;
									default:
										goto end_IL_00fc;
									}
									continue;
									end_IL_00fc:
									break;
								}
							}
							finally
							{
								if (enumerator3 != null)
								{
									while (true)
									{
										switch (7)
										{
										case 0:
											continue;
										}
										enumerator3.Dispose();
										break;
									}
								}
							}
							goto IL_011e;
							IL_011e:
							slideItem = null;
						}
						while (true)
						{
							switch (5)
							{
							case 0:
								break;
							default:
								goto end_IL_012f;
							}
							continue;
							end_IL_012f:
							break;
						}
					}
					finally
					{
						if (enumerator2 != null)
						{
							while (true)
							{
								switch (1)
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
				observableCollection = null;
			}
			else
			{
				G();
			}
			return;
		}
	}

	private void G()
	{
		ObservableCollection<PresentationItem> observableCollection = A();
		IEnumerator<PresentationItem> enumerator = default(IEnumerator<PresentationItem>);
		try
		{
			enumerator = observableCollection.GetEnumerator();
			IEnumerator<ContentItem> enumerator3 = default(IEnumerator<ContentItem>);
			while (enumerator.MoveNext())
			{
				PresentationItem current = enumerator.Current;
				using IEnumerator<SlideItem> enumerator2 = current.Slides.GetEnumerator();
				SlideItem slideItem;
				for (; enumerator2.MoveNext(); slideItem = null)
				{
					SlideItem current2 = enumerator2.Current;
					((BaseItem)current2).IsHighlighted = false;
					slideItem = current2;
					if (!((BaseItem)slideItem).IsExpanded)
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
						if (!slideItem.IsPopulated())
						{
							continue;
						}
					}
					try
					{
						enumerator3 = slideItem.Children.GetEnumerator();
						while (enumerator3.MoveNext())
						{
							((BaseItem)enumerator3.Current).IsHighlighted = false;
						}
					}
					finally
					{
						if (enumerator3 != null)
						{
							while (true)
							{
								switch (3)
								{
								case 0:
									continue;
								}
								enumerator3.Dispose();
								break;
							}
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
						goto end_IL_00be;
					}
					continue;
					end_IL_00be:
					break;
				}
			}
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					goto end_IL_00e1;
				}
				continue;
				end_IL_00e1:
				break;
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
		observableCollection = null;
	}

	private void ExpandAll(object sender, RoutedEventArgs e)
	{
		A(A: true);
		trvExplorer.Focus();
	}

	private void CollapseAll(object sender, RoutedEventArgs e)
	{
		A(A: false);
		trvExplorer.Focus();
	}

	private void A(bool A)
	{
		NF a = default(NF);
		NF CS_0024_003C_003E8__locals9 = new NF(a);
		CS_0024_003C_003E8__locals9.A = this;
		CS_0024_003C_003E8__locals9.A = A;
		CS_0024_003C_003E8__locals9.A = this.A();
		RemoveTreeViewSelectionChangedHandler();
		Task.Factory.StartNew([SpecialName] () =>
		{
			IEnumerator<PresentationItem> enumerator = default(IEnumerator<PresentationItem>);
			try
			{
				enumerator = CS_0024_003C_003E8__locals9.A.GetEnumerator();
				while (enumerator.MoveNext())
				{
					PresentationItem current = enumerator.Current;
					((BaseItem)current).IsExpanded = CS_0024_003C_003E8__locals9.A;
					foreach (SlideItem slide in current.Slides)
					{
						((BaseItem)slide).IsExpanded = CS_0024_003C_003E8__locals9.A;
					}
				}
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
						goto end_IL_0062;
					}
					continue;
					end_IL_0062:
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
							break;
						default:
							enumerator.Dispose();
							goto end_IL_007a;
						}
						continue;
						end_IL_007a:
						break;
					}
				}
			}
			CS_0024_003C_003E8__locals9.A = null;
			CS_0024_003C_003E8__locals9.A.Dispatcher.BeginInvoke((B)([SpecialName] () =>
			{
				CS_0024_003C_003E8__locals9.A.H();
			}));
		});
	}

	private void TreeViewGotKeyboardFocus(object sender, KeyboardFocusChangedEventArgs e)
	{
		Events.Disable(this.m_A);
	}

	private void TreeViewLostKeyboardFocus(object sender, KeyboardFocusChangedEventArgs e)
	{
		Events.Reset(this.m_A);
	}

	private void TreeViewPreviewKeyDown(object sender, System.Windows.Input.KeyEventArgs e)
	{
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			K();
		}
		Key key = e.Key;
		if (key != Key.Up)
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
		}
		if (e.IsRepeat)
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
			RemoveTreeViewSelectionChangedHandler();
			trvExplorer.KeyUp += NavKeyUp;
			return;
		}
	}

	private void NavKeyUp(object sender, System.Windows.Input.KeyEventArgs e)
	{
		trvExplorer.KeyUp -= NavKeyUp;
		H();
		I();
		e.Handled = true;
	}

	private void H()
	{
		trvExplorer.SelectedItemChanged += this.m_A;
	}

	public void RemoveTreeViewSelectionChangedHandler()
	{
		trvExplorer.SelectedItemChanged -= this.m_A;
	}

	private void TreeViewSelectionChanged(System.Windows.Controls.TreeView sender, RoutedPropertyChangedEventArgs<object> e)
	{
		if (Pane.IsPresentationClosing)
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
			if (!trvExplorer.IsKeyboardFocusWithin)
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
				base.Dispatcher.BeginInvoke((B)([SpecialName] () =>
				{
					I();
				}), DispatcherPriority.Normal);
				e.Handled = true;
				return;
			}
		}
	}

	private void I()
	{
		//IL_0032: Unknown result type (might be due to invalid IL or missing references)
		//IL_0038: Expected O, but got Unknown
		if (trvExplorer.SelectedItem == null)
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
			BaseItem val = (BaseItem)trvExplorer.SelectedItem;
			Microsoft.Office.Interop.PowerPoint.Presentation activePresentation = this.m_A.ActivePresentation;
			trvExplorer.LostKeyboardFocus -= TreeViewLostKeyboardFocus;
			try
			{
				if (val is SlideItem)
				{
					while (true)
					{
						switch (1)
						{
						case 0:
							continue;
						}
						A((SlideItem)(object)val);
						break;
					}
				}
				else if (val is PresentationItem)
				{
					while (true)
					{
						switch (6)
						{
						case 0:
							continue;
						}
						A((PresentationItem)(object)val);
						break;
					}
				}
				else
				{
					SlideItem parent = ((ContentItem)(object)val).Parent;
					parent.Parent.Presentation.Windows[1].View.GotoSlide(parent.Slide.SlideIndex);
					A(parent.Parent);
					parent = null;
					if (!(val is CommentItem))
					{
						while (true)
						{
							switch (5)
							{
							case 0:
								continue;
							}
							if (val is NotesItem)
							{
								break;
							}
							while (true)
							{
								switch (2)
								{
								case 0:
									continue;
								}
								if (val is HyperlinkItem)
								{
									while (true)
									{
										switch (5)
										{
										case 0:
											continue;
										}
										A((HyperlinkItem)(object)val);
										break;
									}
									break;
								}
								if (!(val is ContentItem))
								{
									break;
								}
								while (true)
								{
									switch (2)
									{
									case 0:
										continue;
									}
									A((ContentItem)(object)val);
									break;
								}
								break;
							}
							break;
						}
					}
				}
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				B(ex2.Message);
				L();
				clsReporting.LogException(ex2);
				ProjectData.ClearProjectError();
			}
			trvExplorer.LostKeyboardFocus += TreeViewLostKeyboardFocus;
			if (chkShowAll.IsChecked == true)
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
				Microsoft.Office.Interop.PowerPoint.Presentation activePresentation2 = this.m_A.ActivePresentation;
				if (activePresentation2 != activePresentation)
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
					Events.Reset(this.m_A);
					Pane.A(activePresentation2);
				}
				activePresentation2 = null;
			}
			activePresentation = null;
			val = null;
			return;
		}
	}

	private void A(PresentationItem A)
	{
		A.Presentation.Windows[1].Activate();
	}

	private void A(SlideItem A)
	{
		PresentationItem parent = A.Parent;
		DocumentWindow documentWindow = parent.Presentation.Windows[1];
		try
		{
			documentWindow.View.GotoSlide(A.Slide.SlideIndex);
			documentWindow.Activate();
			try
			{
				documentWindow.Panes[1].Activate();
				A.Slide.Select();
				if (parent.Presentation == ThisPresentation)
				{
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
						L();
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
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			try
			{
				_ = A.Slide.Name;
			}
			catch (Exception ex5)
			{
				ProjectData.SetProjectError(ex5);
				Exception ex6 = ex5;
				parent.Slides.Remove(A);
				try
				{
					parent.Slides[Conversions.ToInteger(Operators.SubtractObject(NewLateBinding.LateGet(parent.Presentation.Windows[1].Selection.SlideRange[1], null, AH.A(116304), new object[0], null, null, null), 1))].IsSelected = true;
				}
				catch (Exception ex7)
				{
					ProjectData.SetProjectError(ex7);
					Exception ex8 = ex7;
					ProjectData.ClearProjectError();
				}
				ProjectData.ClearProjectError();
			}
			ProjectData.ClearProjectError();
		}
		parent = null;
		documentWindow = null;
	}

	private void A(ContentItem A)
	{
		if (!clsPowerPoint.IsNormalView(NG.A.Application, false))
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
			try
			{
				Microsoft.Office.Interop.PowerPoint.Shape shape = A.Shape;
				if (shape.Visible == MsoTriState.msoTrue)
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
					shape.Select();
				}
				shape = null;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				C(AH.A(116315));
				B(A);
				ProjectData.ClearProjectError();
			}
			if (A.Parent.Parent.Presentation == ThisPresentation)
			{
				L();
			}
			return;
		}
	}

	private void A(HyperlinkItem A)
	{
		if (A.Shape == null)
		{
			A.Shape = PowerPointAddIn1.Links.Hyperlinks.GetParentShape(A.Hyperlink, blnIgnoreTables: false);
		}
		this.A((ContentItem)A);
	}

	private void B(ContentItem A)
	{
		A.Parent.RemoveChild(A);
	}

	private void OnSelected(object sender, RoutedEventArgs e)
	{
		try
		{
			if (!trvExplorer.IsKeyboardFocusWithin)
			{
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
		TreeViewItem treeViewItem = (TreeViewItem)sender;
		if (treeViewItem.DataContext is PresentationItem)
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
			PresentationItem presentationItem = (PresentationItem)treeViewItem.DataContext;
			if (presentationItem.Slides.Count != presentationItem.Presentation.Slides.Count)
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
				presentationItem.RefreshSlides();
			}
			presentationItem = null;
		}
		else if (treeViewItem.DataContext is SlideItem)
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
			SlideItem slideItem = (SlideItem)treeViewItem.DataContext;
			if (!slideItem.IsPopulated())
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
				slideItem.Populate();
			}
			slideItem = null;
		}
		treeViewItem = null;
	}

	private void NodePreviewMouseDown(object sender, MouseButtonEventArgs e)
	{
		if (chkPreview.IsChecked != true)
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
			K();
			return;
		}
	}

	private void NodePreviewMouseRightButtonDown(object sender, MouseButtonEventArgs e)
	{
		//IL_000d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0013: Expected O, but got Unknown
		BaseItem val = (BaseItem)((TreeViewItem)sender).DataContext;
		if (!val.IsSelected)
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
			val.IsSelected = true;
		}
		val = null;
	}

	private void NodeKeyDown(object sender, System.Windows.Input.KeyEventArgs e)
	{
		TreeViewItem treeViewItem = (TreeViewItem)sender;
		if (e.Key == Key.X)
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
			if (e.KeyboardDevice.Modifiers == ModifierKeys.Control)
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
				if (trvExplorer.SelectedItem is SlideItem)
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
					B((SlideItem)trvExplorer.SelectedItem);
					e.Handled = true;
				}
				goto IL_04cf;
			}
		}
		if (e.Key == Key.C && e.KeyboardDevice.Modifiers == ModifierKeys.Control)
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
			if (trvExplorer.SelectedItem is SlideItem)
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
				C((SlideItem)trvExplorer.SelectedItem);
				e.Handled = true;
			}
		}
		else
		{
			if (e.Key == Key.V)
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
					while (true)
					{
						switch (7)
						{
						case 0:
							continue;
						}
						break;
					}
					if (trvExplorer.SelectedItem is SlideItem)
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
						D((SlideItem)trvExplorer.SelectedItem);
						e.Handled = true;
					}
					goto IL_04cf;
				}
			}
			if (e.Key == Key.F5)
			{
				goto IL_01b6;
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
			if (e.Key == Key.R)
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
				if (Keyboard.Modifiers == ModifierKeys.Control)
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
					goto IL_01b6;
				}
			}
			if (e.Key == Key.Up)
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
				if (e.KeyboardDevice.Modifiers == ModifierKeys.Control)
				{
					if (treeViewItem.DataContext is SlideItem)
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
						((SlideItem)treeViewItem.DataContext).MoveUp();
						e.Handled = true;
					}
					goto IL_04cf;
				}
			}
			if (e.Key == Key.Down && e.KeyboardDevice.Modifiers == ModifierKeys.Control)
			{
				if (treeViewItem.DataContext is SlideItem)
				{
					((SlideItem)treeViewItem.DataContext).MoveDown();
					e.Handled = true;
				}
			}
			else
			{
				Key key = e.Key;
				if (key <= Key.Escape)
				{
					if (key == Key.Return)
					{
						goto IL_0498;
					}
					if (key != Key.Escape)
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
					}
					else
					{
						SendKeys.Send(AH.A(116348));
						e.Handled = true;
					}
				}
				else
				{
					if (key == Key.Space)
					{
						goto IL_0498;
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
					if (key != Key.Delete)
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
							if (treeViewItem.DataContext is SlideItem)
							{
								((SlideItem)treeViewItem.DataContext).Rename();
								L();
							}
							else if (treeViewItem.DataContext is HyperlinkItem)
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
								((HyperlinkItem)treeViewItem.DataContext).Edit();
							}
							else if (treeViewItem.DataContext is ChartItem)
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
							if (treeViewItem.DataContext is SlideItem)
							{
								((SlideItem)treeViewItem.DataContext).Delete();
								L();
							}
							else
							{
								if (!(treeViewItem.DataContext is ContentItem))
								{
									throw new NotImplementedException();
								}
								while (true)
								{
									switch (1)
									{
									case 0:
										continue;
									}
									((ContentItem)treeViewItem.DataContext).Delete();
									L();
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
			}
		}
		goto IL_04cf;
		IL_0498:
		if (treeViewItem.DataContext is HyperlinkItem)
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
			B((HyperlinkItem)treeViewItem.DataContext);
		}
		else
		{
			J();
		}
		e.Handled = true;
		goto IL_04cf;
		IL_04cf:
		treeViewItem = null;
		return;
		IL_01b6:
		try
		{
			if (treeViewItem.DataContext is PresentationItem)
			{
				while (true)
				{
					switch (2)
					{
					case 0:
						continue;
					}
					((PresentationItem)treeViewItem.DataContext).Refresh();
					break;
				}
			}
			else if (treeViewItem.DataContext is SlideItem)
			{
				((SlideItem)treeViewItem.DataContext).Refresh();
			}
			else
			{
				if (!(treeViewItem.DataContext is ContentItem))
				{
					throw new NotImplementedException();
				}
				while (true)
				{
					switch (3)
					{
					case 0:
						continue;
					}
					((ContentItem)treeViewItem.DataContext).Refresh();
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
		goto IL_04cf;
	}

	private void J()
	{
		C(AH.A(116357));
		L();
		trvExplorer.Focus();
	}

	private void OnRequestBringIntoView(object sender, RequestBringIntoViewEventArgs e)
	{
		if (this.m_A == null)
		{
			this.m_A = (ScrollViewer)trvExplorer.Template.FindName(AH.A(116448), trvExplorer);
			if (this.m_A != null)
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
				this.m_A.ScrollChanged -= TreeViewScrollViewerScrollChanged;
				this.m_A.ScrollChanged += TreeViewScrollViewerScrollChanged;
			}
		}
		this.m_A = true;
		this.m_A = this.m_A.HorizontalOffset;
	}

	private void TreeViewScrollViewerScrollChanged(object sender, ScrollChangedEventArgs e)
	{
		if (this.m_A)
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
			this.m_A.ScrollToHorizontalOffset(this.m_A);
		}
		this.m_A = false;
	}

	private DependencyObject A(DependencyObject A)
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
					return this.A(parent);
				}
			}
		}
		return parent;
	}

	private void OnGotFocusEvent(object sender, RoutedEventArgs e)
	{
		if (LogicalTreeHelper.GetParent((DependencyObject)e.OriginalSource) == this.m_A)
		{
			base.Dispatcher.BeginInvoke((C<bool>)([SpecialName] () =>
			{
				this.m_A = new DispatcherFrame();
				Dispatcher.PushFrame(this.m_A);
				return false;
			}), DispatcherPriority.Normal);
		}
	}

	private void OnLostFocusEvent(object sender, RoutedEventArgs e)
	{
		if (LogicalTreeHelper.GetParent((DependencyObject)e.OriginalSource) == this.m_A)
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
			this.m_A.Continue = false;
		}
		Dispatcher.ExitAllFrames();
	}

	private void A(System.Windows.Controls.ContextMenu A)
	{
		this.m_A = A;
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
		this.m_A = RuntimeHelpers.GetObjectValue(sender);
		System.Windows.Controls.MenuItem obj = (System.Windows.Controls.MenuItem)sender;
		obj.RemoveHandler(UIElement.GotFocusEvent, new RoutedEventHandler(OnGotFocusEvent));
		obj.RemoveHandler(UIElement.LostFocusEvent, new RoutedEventHandler(OnLostFocusEvent));
		obj.AddHandler(UIElement.GotFocusEvent, new RoutedEventHandler(OnGotFocusEvent));
		obj.AddHandler(UIElement.LostFocusEvent, new RoutedEventHandler(OnLostFocusEvent));
		_ = null;
	}

	private void SlideMenuOpened(object sender, RoutedEventArgs e)
	{
		System.Windows.Controls.ContextMenu contextMenu = (System.Windows.Controls.ContextMenu)sender;
		A(contextMenu);
		IEnumerator enumerator = default(IEnumerator);
		SlideItem slideItem;
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
				slideItem = (SlideItem)contextMenu.DataContext;
				System.Windows.Controls.MenuItem menuItem = (System.Windows.Controls.MenuItem)objectValue;
				string left = menuItem.Header.ToString().ToLower();
				if (Operators.CompareString(left, AH.A(116483), TextCompare: false) != 0)
				{
					if (Operators.CompareString(left, AH.A(116504), TextCompare: false) != 0)
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
						menuItem.IsEnabled = slideItem.Slide.SlideShowTransition.Hidden == MsoTriState.msoTrue;
					}
				}
				else
				{
					menuItem.IsEnabled = slideItem.Slide.SlideShowTransition.Hidden == MsoTriState.msoFalse;
				}
				A((System.Windows.Controls.MenuItem)objectValue);
				menuItem = null;
			}
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					goto end_IL_010c;
				}
				continue;
				end_IL_010c:
				break;
			}
		}
		finally
		{
			if (enumerator is IDisposable)
			{
				while (true)
				{
					switch (2)
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
		slideItem = null;
	}

	private void A(System.Windows.Controls.MenuItem A)
	{
		SolidColorBrush solidColorBrush = new SolidColorBrush(clsColors.GrayFontColor());
		try
		{
			SolidColorBrush solidColorBrush2 = solidColorBrush;
			double opacity;
			if (!A.IsEnabled)
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
				opacity = Constants.INACTIVE_FILE_OPACITY;
			}
			else
			{
				opacity = 1.0;
			}
			solidColorBrush2.Opacity = opacity;
			((GeometryDrawing)((DrawingImage)((System.Windows.Controls.Image)A.Icon).Source).Drawing).Brush = solidColorBrush;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		solidColorBrush = null;
	}

	private void CloseOthers(object sender, RoutedEventArgs e)
	{
		A((System.Windows.Controls.MenuItem)sender).CloseOthers();
	}

	private void Reopen(object sender, RoutedEventArgs e)
	{
		A((System.Windows.Controls.MenuItem)sender).Reopen();
	}

	private void Duplicate(object sender, RoutedEventArgs e)
	{
		A((System.Windows.Controls.MenuItem)sender).Duplicate();
	}

	private void RenamePresentation(object sender, RoutedEventArgs e)
	{
		A((System.Windows.Controls.MenuItem)sender).Rename();
	}

	private void ShowInFolder(object sender, RoutedEventArgs e)
	{
		A((System.Windows.Controls.MenuItem)sender).ShowInFolder();
	}

	private void CopyPath(object sender, RoutedEventArgs e)
	{
		A((System.Windows.Controls.MenuItem)sender).CopyPath();
	}

	private void FinalizePagination(object sender, RoutedEventArgs e)
	{
		if (!PowerPointAddIn1.Pagination.Pane.IsVisible())
		{
			PowerPointAddIn1.Pagination.Pane.Toggle(blnPressed: true);
			KG.A.InvalidateControl(PowerPointAddIn1.Pagination.Pane.RIBBON_CONTROL);
		}
	}

	private void PrepareToShare(object sender, RoutedEventArgs e)
	{
		PowerPointAddIn1.Publishing.Share.Pane.Toggle(blnPressed: true);
	}

	private void SendPresentation(object sender, RoutedEventArgs e)
	{
		Send.ShowDialog();
	}

	private PresentationItem A(System.Windows.Controls.MenuItem A)
	{
		return (PresentationItem)A.DataContext;
	}

	private void MarkAsFacing(object sender, RoutedEventArgs e)
	{
		FacingSlides.MarkSlide();
	}

	private void UnmarkAsFacing(object sender, RoutedEventArgs e)
	{
		FacingSlides.UnmarkSlide();
	}

	private void Unhide(object sender, RoutedEventArgs e)
	{
		A((System.Windows.Controls.MenuItem)sender).Unhide();
	}

	private void Hide(object sender, RoutedEventArgs e)
	{
		A((System.Windows.Controls.MenuItem)sender).Hide();
	}

	private void InsertSlide(object sender, RoutedEventArgs e)
	{
		A((System.Windows.Controls.MenuItem)sender).Insert();
	}

	private void DeleteSlide(object sender, RoutedEventArgs e)
	{
		A((System.Windows.Controls.MenuItem)sender).Delete();
		L();
	}

	private void RenameSlide(object sender, RoutedEventArgs e)
	{
		A((System.Windows.Controls.MenuItem)sender).Rename();
		L();
	}

	private void DuplicateSlide(object sender, RoutedEventArgs e)
	{
		SlideItem slideItem = A((System.Windows.Controls.MenuItem)sender);
		Slide slide;
		try
		{
			slide = slideItem.Slide.Duplicate()[1];
			slide.Select();
			L();
			A(slideItem, slide, checked(slide.SlideIndex - 1));
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			B(ex2.Message);
			trvExplorer.Focus();
			ProjectData.ClearProjectError();
		}
		H();
		slide = null;
		slideItem = null;
	}

	private void CutSlide(object sender, RoutedEventArgs e)
	{
		B(A((System.Windows.Controls.MenuItem)sender));
	}

	private void B(SlideItem A)
	{
		RemoveTreeViewSelectionChangedHandler();
		A.Slide.Select();
		A.IsSelected = false;
		this.m_A.ActiveWindow.Selection.Cut();
		A.Parent.Slides.Remove(A);
		H();
		L();
	}

	private void CopySlide(object sender, RoutedEventArgs e)
	{
		C(A((System.Windows.Controls.MenuItem)sender));
	}

	private void C(SlideItem A)
	{
		A.Slide.Select();
		this.m_A.ActiveWindow.Selection.Copy();
		L();
	}

	private void PasteSlide(object sender, RoutedEventArgs e)
	{
		D(A((System.Windows.Controls.MenuItem)sender));
	}

	private void D(SlideItem A)
	{
		int num = 0;
		Slide b;
		try
		{
			num = this.m_A.ActiveWindow.Selection.SlideRange[1].SlideIndex;
			b = A.Parent.Presentation.Slides.Paste(num)[1];
			this.A(A, b, checked(num - 1));
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			B(ex2.Message);
			trvExplorer.Focus();
			ProjectData.ClearProjectError();
		}
		H();
		b = null;
	}

	private void A(SlideItem A, Slide B, int C)
	{
		RemoveTreeViewSelectionChangedHandler();
		ObservableCollection<SlideItem> slides = A.Parent.Slides;
		slides.Add(new SlideItem(A.Parent, B));
		slides.Move(checked(slides.Count - 1), C);
		SlideItem slideItem = slides[C];
		if (((BaseItem)A).IsExpanded)
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
			slideItem.Populate();
			((BaseItem)slideItem).IsExpanded = true;
		}
		slideItem.IsSelected = true;
		slideItem = null;
		_ = null;
	}

	private void SendToEnd(object sender, RoutedEventArgs e)
	{
		SlideItem slideItem = A((System.Windows.Controls.MenuItem)sender);
		slideItem.IsSelected = false;
		slideItem.SendToEnd();
		RemoveTreeViewSelectionChangedHandler();
		try
		{
			slideItem.Parent.Slides.Where([SpecialName] (SlideItem A) => A.Slide == this.m_A.ActivePresentation.Windows[1].Selection.SlideRange[1]).ElementAtOrDefault(0).IsSelected = true;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		H();
		slideItem = null;
		L();
	}

	private SlideItem A(System.Windows.Controls.MenuItem A)
	{
		return (SlideItem)A.DataContext;
	}

	private void RefreshLink(object sender, RoutedEventArgs e)
	{
		Ribbon.RefreshLinks();
	}

	private void ViewLinkSource(object sender, RoutedEventArgs e)
	{
		Ribbon.ViewSource();
	}

	private void EditLink(object sender, RoutedEventArgs e)
	{
		Ribbon.EditLinks();
	}

	private void BreakLink(object sender, RoutedEventArgs e)
	{
		this.m_A.StartNewUndoEntry();
		ContentItem obj = (ContentItem)((System.Windows.Controls.MenuItem)sender).DataContext;
		PowerPointAddIn1.Links.Shapes.BreakLink(obj.Shape);
		obj.IsLinked = false;
		_ = null;
	}

	private void RotateRight90(object sender, RoutedEventArgs e)
	{
		ContentItem contentItem = A((System.Windows.Controls.MenuItem)sender);
		contentItem.Shape.IncrementRotation(-90f);
		contentItem.PreviewImage = null;
		_ = null;
	}

	private void RotateLeft90(object sender, RoutedEventArgs e)
	{
		ContentItem contentItem = A((System.Windows.Controls.MenuItem)sender);
		contentItem.Shape.IncrementRotation(90f);
		contentItem.PreviewImage = null;
		_ = null;
	}

	private void FlipVertical(object sender, RoutedEventArgs e)
	{
		ContentItem contentItem = A((System.Windows.Controls.MenuItem)sender);
		contentItem.Shape.Flip(MsoFlipCmd.msoFlipVertical);
		contentItem.PreviewImage = null;
		_ = null;
	}

	private void FlipHorizontal(object sender, RoutedEventArgs e)
	{
		ContentItem contentItem = A((System.Windows.Controls.MenuItem)sender);
		contentItem.Shape.Flip(MsoFlipCmd.msoFlipHorizontal);
		contentItem.PreviewImage = null;
		_ = null;
	}

	private void BringForward(object sender, RoutedEventArgs e)
	{
		A((System.Windows.Controls.MenuItem)sender).ZOrder(MsoZOrderCmd.msoBringForward);
	}

	private void BringToFront(object sender, RoutedEventArgs e)
	{
		A((System.Windows.Controls.MenuItem)sender).ZOrder(MsoZOrderCmd.msoBringToFront);
	}

	private void SendBackward(object sender, RoutedEventArgs e)
	{
		A((System.Windows.Controls.MenuItem)sender).ZOrder(MsoZOrderCmd.msoSendBackward);
	}

	private void SendToback(object sender, RoutedEventArgs e)
	{
		A((System.Windows.Controls.MenuItem)sender).ZOrder(MsoZOrderCmd.msoSendToBack);
	}

	private void HideShape(object sender, RoutedEventArgs e)
	{
		A((System.Windows.Controls.MenuItem)sender).Hide();
	}

	private void UnhideShape(object sender, RoutedEventArgs e)
	{
		A((System.Windows.Controls.MenuItem)sender).Unhide();
	}

	private void DeleteShape(object sender, RoutedEventArgs e)
	{
		A((System.Windows.Controls.MenuItem)sender).Delete();
		L();
	}

	private ContentItem A(System.Windows.Controls.MenuItem A)
	{
		return (ContentItem)A.DataContext;
	}

	private Microsoft.Office.Interop.PowerPoint.Shape A(System.Windows.Controls.MenuItem A)
	{
		return ((ContentItem)A.DataContext).Shape;
	}

	private void FixImageScaling(object sender, RoutedEventArgs e)
	{
		A((System.Windows.Controls.MenuItem)sender).Select();
		Images.FixScale();
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
			L();
			ProjectData.ClearProjectError();
		}
	}

	private void RemoveHyperlink(object sender, RoutedEventArgs e)
	{
		((HyperlinkItem)((System.Windows.Controls.MenuItem)sender).DataContext).Delete();
		L();
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
	}

	private void MemorizeChart(object sender, RoutedEventArgs e)
	{
	}

	private void ApplyChartSize(object sender, RoutedEventArgs e)
	{
	}

	private void ApplyPlotSize(object sender, RoutedEventArgs e)
	{
	}

	private void ApplyPlotPosition(object sender, RoutedEventArgs e)
	{
	}

	private void ApplyAll(object sender, RoutedEventArgs e)
	{
	}

	private void ResizeToExcel(object sender, RoutedEventArgs e)
	{
	}

	private void ResizeToPowerPoint(object sender, RoutedEventArgs e)
	{
	}

	private void ResizeToWord(object sender, RoutedEventArgs e)
	{
	}

	private void RecolorChartDefaults(object sender, RoutedEventArgs e)
	{
	}

	private void RecolorChartPoints(object sender, RoutedEventArgs e)
	{
	}

	private void SaveAsPicture(object sender, RoutedEventArgs e)
	{
	}

	private void DeleteChart(object sender, RoutedEventArgs e)
	{
		A((System.Windows.Controls.MenuItem)sender).Delete();
		L();
	}

	private ChartItem A(System.Windows.Controls.MenuItem A)
	{
		return (ChartItem)A.DataContext;
	}

	private void ConvertToPicture(object sender, RoutedEventArgs e)
	{
		ContentItem contentItem = (ContentItem)((System.Windows.Controls.MenuItem)sender).DataContext;
		SlideItem parent = contentItem.Parent;
		Microsoft.Office.Interop.PowerPoint.Shape shape = null;
		PowerPointAddIn1.Shapes.ConvertToPicture.SelectedObjects();
		try
		{
			shape = this.m_A.ActiveWindow.Selection.ShapeRange[1];
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		((BaseItem)contentItem).IsSelected = false;
		ImageItem imageItem;
		if (shape != null)
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
			imageItem = parent.NewImageItem(shape);
			SlideItem slideItem = parent;
			ObservableCollection<ContentItem> children = slideItem.Children;
			children.Add(imageItem);
			children.Move(checked(children.Count - 1), children.IndexOf(contentItem));
			children = null;
			slideItem.RemoveChild(contentItem);
			_ = null;
			imageItem.IsSelected = true;
		}
		else
		{
			parent.Refresh();
		}
		parent = null;
		contentItem = null;
		imageItem = null;
	}

	private void DeleteTable(object sender, RoutedEventArgs e)
	{
		((TableItem)((System.Windows.Controls.MenuItem)sender).DataContext).Delete();
		L();
	}

	private void DeleteComment(object sender, RoutedEventArgs e)
	{
		((CommentItem)((System.Windows.Controls.MenuItem)sender).DataContext).Delete();
		L();
	}

	private void DeleteNotes(object sender, RoutedEventArgs e)
	{
		((NotesItem)((System.Windows.Controls.MenuItem)sender).DataContext).Delete();
		L();
	}

	private void MouseEnterNode(object sender, System.Windows.Input.MouseEventArgs e)
	{
		//IL_0273: Unknown result type (might be due to invalid IL or missing references)
		//IL_027d: Expected O, but got Unknown
		//IL_00eb: Unknown result type (might be due to invalid IL or missing references)
		//IL_00f5: Expected O, but got Unknown
		bool? isChecked;
		bool? flag = (isChecked = chkPreview.IsChecked);
		bool? flag2;
		if (flag.HasValue)
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
			if (isChecked != true)
			{
				flag2 = false;
				goto IL_0066;
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
		if (this.m_B != null)
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
			flag2 = false;
		}
		else
		{
			flag2 = isChecked;
		}
		goto IL_0066;
		IL_0066:
		bool? flag3 = flag2;
		if (flag3.HasValue)
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
		if (Keyboard.Modifiers != ModifierKeys.None)
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
			if (!flag3.HasValue)
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
				if (((TreeViewItem)sender).HasItems)
				{
					return;
				}
				BaseItem A = (BaseItem)((TreeViewItem)sender).DataContext;
				BitmapSource A2 = null;
				if (this.m_A != null)
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
					this.m_A.Stop();
					this.m_A = null;
				}
				if (!(A is CommentItem))
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
					if (!(A is NotesItem) && !(A is HyperlinkItem))
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
						if (A is ContentItem)
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
							ContentItem contentItem = (ContentItem)(object)A;
							if (contentItem.PreviewImage != null)
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
								A2 = contentItem.PreviewImage;
							}
							else
							{
								string text = null;
								try
								{
									text = modFunctionsIO.PathGetTempFileName();
									contentItem.Shape.Export(text, PpShapeFormat.ppShapeFormatPNG);
									Bitmap bitmap = new Bitmap(text);
									try
									{
										A2 = Forms.GetImageSource(bitmap);
									}
									finally
									{
										if (bitmap != null)
										{
											while (true)
											{
												switch (5)
												{
												case 0:
													continue;
												}
												((IDisposable)bitmap).Dispose();
												break;
											}
										}
									}
									A2.Freeze();
									contentItem.PreviewImage = A2;
								}
								catch (Exception projectError)
								{
									ProjectData.SetProjectError(projectError);
									ProjectData.ClearProjectError();
								}
								finally
								{
									try
									{
										if (text != null)
										{
											while (true)
											{
												switch (6)
												{
												case 0:
													continue;
												}
												File.Delete(text);
												break;
											}
										}
									}
									catch (Exception projectError2)
									{
										ProjectData.SetProjectError(projectError2);
										ProjectData.ClearProjectError();
									}
								}
							}
							contentItem = null;
						}
					}
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
								switch (6)
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
							Preview.PrepareWindow(PreviewWindow, A, A2, PointToScreen(new System.Windows.Point(0.0, 0.0)));
						}, DispatcherPriority.Background);
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						ProjectData.ClearProjectError();
					}
					A2 = null;
					e.Handled = true;
				}
				else
				{
					K();
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
			if (isChecked != true)
			{
				obj = false;
				goto IL_0054;
			}
		}
		obj = ((PreviewWindow != null) ? isChecked : new bool?(false));
		goto IL_0054;
		IL_0054:
		bool? flag2 = obj;
		if (flag2.HasValue)
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
			if (flag2 != true)
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
				break;
			}
		}
		if (!((UIElement)(object)PreviewWindow).IsVisible || !flag2.HasValue)
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
			this.m_A = new DispatcherTimer();
			this.m_A.Interval = TimeSpan.FromMilliseconds(100.0);
			this.m_A.Tick += [SpecialName] (object obj2, EventArgs ev) =>
			{
				try
				{
					this.m_A.Stop();
					K();
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
			switch (2)
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
					switch (3)
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
						switch (5)
						{
						case 0:
							continue;
						}
						break;
					}
					((Window)(object)PreviewWindow).Show();
				}
				System.Windows.Point position = e.GetPosition((IInputElement)PreviewWindow);
				((Window)(object)PreviewWindow).Top += position.Y + 0.0;
				((Window)(object)PreviewWindow).Left += position.X + 25.0;
			}
			e.Handled = true;
			return;
		}
	}

	private void K()
	{
		if (PreviewWindow == null)
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
			((Window)(object)PreviewWindow).Close();
			PreviewWindow = null;
			return;
		}
	}

	private void PreviewToggle(object sender, RoutedEventArgs e)
	{
		K();
		PB.Settings.ExplorerPreviews = chkPreview.IsChecked.Value;
	}

	private void trvExplorer_PreviewMouseWheel(object sender, MouseWheelEventArgs e)
	{
		K();
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

	private void treeView_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
	{
		this.m_A = e.GetPosition((IInputElement)sender);
	}

	private void treeView_PreviewMouseMove(object sender, System.Windows.Input.MouseEventArgs e)
	{
		if (e.LeftButton != MouseButtonState.Pressed)
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
			if (this.m_C)
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
				System.Windows.Point position = e.GetPosition((IInputElement)sender);
				if (!(Math.Abs(position.X - this.m_A.X) > SystemParameters.MinimumHorizontalDragDistance))
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
					if (!(Math.Abs(position.Y - this.m_A.Y) > SystemParameters.MinimumVerticalDragDistance))
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
						break;
					}
				}
				A(e);
				return;
			}
		}
	}

	private void A(System.Windows.Input.MouseEventArgs A)
	{
		TreeViewItem treeViewItem = DragAndDrop.FindAncestor<TreeViewItem>((DependencyObject)A.OriginalSource);
		if (treeViewItem == null)
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
			BaseItem itemAtLocation = DragAndDrop.GetItemAtLocation(A.GetPosition(trvExplorer), trvExplorer);
			if (itemAtLocation == null || !(itemAtLocation is SlideItem))
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
				this.m_C = true;
				try
				{
					System.Windows.DataObject data = new System.Windows.DataObject(typeof(BaseItem), itemAtLocation);
					DragDrop.DoDragDrop(treeViewItem, data, System.Windows.DragDropEffects.Move);
				}
				finally
				{
					this.m_C = false;
				}
				treeViewItem = null;
				itemAtLocation = null;
				return;
			}
		}
	}

	private void treeView_QueryContinueDrag(object sender, System.Windows.QueryContinueDragEventArgs e)
	{
	}

	private void treeView_DragEnter(object sender, System.Windows.DragEventArgs e)
	{
		this.m_B = e.GetPosition(trvExplorer);
	}

	private void treeView_Drop(object sender, System.Windows.DragEventArgs e)
	{
		if (!e.Data.GetDataPresent(typeof(BaseItem)))
		{
			return;
		}
		SlideItem slideItem = (SlideItem)e.Data.GetData(typeof(BaseItem));
		BaseItem itemAtLocation = DragAndDrop.GetItemAtLocation(this.m_B, trvExplorer);
		bool flag;
		PresentationItem parent;
		if (itemAtLocation != null)
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
			if ((object)itemAtLocation != slideItem)
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
				if (chkShowAll.IsChecked == true)
				{
					parent = slideItem.Parent;
					if (itemAtLocation is SlideItem)
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
						if (parent.Presentation == ((SlideItem)(object)itemAtLocation).Parent.Presentation)
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
							flag = true;
							goto IL_0116;
						}
					}
					if (itemAtLocation is PresentationItem)
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
						if (parent.Presentation == ((PresentationItem)(object)itemAtLocation).Presentation)
						{
							flag = true;
							goto IL_0116;
						}
					}
					flag = false;
					goto IL_0116;
				}
				flag = true;
				goto IL_011d;
			}
		}
		goto IL_03ae;
		IL_0116:
		parent = null;
		goto IL_011d;
		IL_011d:
		checked
		{
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
				if (itemAtLocation is SlideItem)
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
						SlideItem slideItem2 = slideItem;
						slideItem2.Slide.MoveTo(((SlideItem)(object)itemAtLocation).Slide.SlideIndex);
						ObservableCollection<SlideItem> slides = slideItem2.Parent.Slides;
						int oldIndex = slides.IndexOf(slideItem);
						int newIndex = slides.IndexOf((SlideItem)(object)itemAtLocation);
						slides.Move(oldIndex, newIndex);
						_ = null;
						_ = null;
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						B(ex2.Message);
						trvExplorer.Focus();
						ProjectData.ClearProjectError();
					}
				}
			}
			else
			{
				Slide slide = null;
				PresentationItem presentationItem = default(PresentationItem);
				if (itemAtLocation is SlideItem)
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
					slide = ((SlideItem)(object)itemAtLocation).Slide;
					presentationItem = ((SlideItem)(object)itemAtLocation).Parent;
				}
				else if (itemAtLocation is PresentationItem)
				{
					slide = ((PresentationItem)(object)itemAtLocation).Presentation.Slides[1];
					presentationItem = (PresentationItem)(object)itemAtLocation;
				}
				if (slide != null)
				{
					PresentationItem parent2 = slideItem.Parent;
					_ = parent2.Presentation;
					RemoveTreeViewSelectionChangedHandler();
					trvExplorer.LostKeyboardFocus -= TreeViewLostKeyboardFocus;
					Slide slide2;
					try
					{
						slideItem.Slide.Copy();
						slide2 = presentationItem.Presentation.Slides.Paste()[1];
						int slideIndex = slide.SlideIndex;
						slide2.MoveTo(slideIndex);
						presentationItem.Presentation.Windows[1].Activate();
						bool isExpanded = ((BaseItem)slideItem).IsExpanded;
						parent2.Slides.Remove(slideItem);
						ObservableCollection<SlideItem> slides2 = presentationItem.Slides;
						slides2.Add(new SlideItem(presentationItem, slide2));
						SlideItem slideItem3 = slides2[slides2.Count - 1];
						if (isExpanded)
						{
							slideItem3.Populate();
							((BaseItem)slideItem3).IsExpanded = true;
						}
						slideItem3.IsSelected = true;
						slideItem3 = null;
						slides2.Move(slides2.Count - 1, slideIndex - 1);
						_ = null;
						Pane.A(presentationItem.Presentation);
					}
					catch (Exception ex3)
					{
						ProjectData.SetProjectError(ex3);
						Exception ex4 = ex3;
						B(ex4.Message);
						trvExplorer.Focus();
						ProjectData.ClearProjectError();
					}
					slide2 = null;
					H();
					trvExplorer.LostKeyboardFocus += TreeViewLostKeyboardFocus;
					slide = null;
					parent2 = null;
				}
				presentationItem = null;
			}
			goto IL_03ae;
		}
		IL_03ae:
		slideItem = null;
		itemAtLocation = null;
	}

	private ObservableCollection<PresentationItem> A()
	{
		if (chkShowAll.IsChecked == true)
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
					return Pane.AllPresentations;
				}
			}
		}
		return SinglePresentation;
	}

	private void L()
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
		trvExplorer.Focus();
	}

	private void B(string A)
	{
		Forms.ErrorMessage(Window.GetWindow(this), A);
	}

	private void C(string A)
	{
		Forms.WarningMessage(Window.GetWindow(this), A);
	}

	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	[DebuggerNonUserCode]
	public void InitializeComponent()
	{
		if (!this.m_D)
		{
			this.m_D = true;
			Uri resourceLocator = new Uri(AH.A(116529), UriKind.Relative);
			System.Windows.Application.LoadComponent(this, resourceLocator);
		}
	}

	void IComponentConnector.InitializeComponent()
	{
		//ILSpy generated this explicit interface implementation from .override directive in InitializeComponent
		this.InitializeComponent();
	}

	[EditorBrowsable(EditorBrowsableState.Never)]
	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	[DebuggerNonUserCode]
	public void System_Windows_Markup_IComponentConnector_Connect(int connectionId, object target)
	{
		if (connectionId == 1)
		{
			((System.Windows.Controls.ContextMenu)target).Opened += MenuOpened;
			return;
		}
		if (connectionId == 2)
		{
			((System.Windows.Controls.MenuItem)target).Click += HideShape;
			return;
		}
		if (connectionId == 3)
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
					((System.Windows.Controls.MenuItem)target).Click += UnhideShape;
					return;
				}
			}
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
					((System.Windows.Controls.MenuItem)target).Click += DeleteShape;
					return;
				}
			}
		}
		if (connectionId == 6)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					chkSettings = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 7)
		{
			chkPreview = (System.Windows.Controls.CheckBox)target;
			return;
		}
		if (connectionId == 8)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					chkShowAll = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 9)
		{
			btnRefresh = (System.Windows.Controls.Button)target;
			return;
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
					btnExpand = (System.Windows.Controls.Button)target;
					return;
				}
			}
		}
		if (connectionId == 11)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					btnCollapse = (System.Windows.Controls.Button)target;
					return;
				}
			}
		}
		if (connectionId == 12)
		{
			btnSpacing = (System.Windows.Controls.Button)target;
			return;
		}
		if (connectionId == 13)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					popSettings = (Popup)target;
					return;
				}
			}
		}
		if (connectionId == 14)
		{
			chkCharts = (System.Windows.Controls.CheckBox)target;
			return;
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
					chkTables = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 16)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					chkExcel = (System.Windows.Controls.CheckBox)target;
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
					chkWord = (System.Windows.Controls.CheckBox)target;
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
					chkSmartArt = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 19)
		{
			chkImages = (System.Windows.Controls.CheckBox)target;
			return;
		}
		if (connectionId == 20)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					chkMedia = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 21)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					chkInk = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 22)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					chkHyperlinks = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 23)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					chkComments = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 24)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					chkNotes = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 25)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					trvExplorer = (System.Windows.Controls.TreeView)target;
					return;
				}
			}
		}
		switch (connectionId)
		{
		case 111:
			txtSearch = (System.Windows.Controls.TextBox)target;
			break;
		case 112:
			while (true)
			{
				switch (4)
				{
				case 0:
					continue;
				}
				((System.Windows.Controls.Button)target).Click += ClearSearch;
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

	[EditorBrowsable(EditorBrowsableState.Never)]
	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	[DebuggerNonUserCode]
	public void System_Windows_Markup_IStyleConnector_Connect(int connectionId, object target)
	{
		if (connectionId == 5)
		{
			EventSetter eventSetter = new EventSetter();
			eventSetter.Event = FrameworkElement.RequestBringIntoViewEvent;
			eventSetter.Handler = new RequestBringIntoViewEventHandler(OnRequestBringIntoView);
			((Style)target).Setters.Add(eventSetter);
			eventSetter = new EventSetter();
			eventSetter.Event = TreeViewItem.SelectedEvent;
			eventSetter.Handler = new RoutedEventHandler(OnSelected);
			((Style)target).Setters.Add(eventSetter);
			eventSetter = new EventSetter();
			eventSetter.Event = TreeViewItem.ExpandedEvent;
			eventSetter.Handler = new RoutedEventHandler(OnExpanded);
			((Style)target).Setters.Add(eventSetter);
			eventSetter = new EventSetter();
			eventSetter.Event = UIElement.KeyDownEvent;
			eventSetter.Handler = new System.Windows.Input.KeyEventHandler(NodeKeyDown);
			((Style)target).Setters.Add(eventSetter);
			eventSetter = new EventSetter();
			eventSetter.Event = UIElement.PreviewMouseDownEvent;
			eventSetter.Handler = new MouseButtonEventHandler(NodePreviewMouseDown);
			((Style)target).Setters.Add(eventSetter);
			eventSetter = new EventSetter();
			eventSetter.Event = UIElement.PreviewMouseRightButtonDownEvent;
			eventSetter.Handler = new MouseButtonEventHandler(NodePreviewMouseRightButtonDown);
			((Style)target).Setters.Add(eventSetter);
			eventSetter = new EventSetter();
			eventSetter.Event = UIElement.MouseEnterEvent;
			eventSetter.Handler = new System.Windows.Input.MouseEventHandler(MouseEnterNode);
			((Style)target).Setters.Add(eventSetter);
			eventSetter = new EventSetter();
			eventSetter.Event = UIElement.MouseLeaveEvent;
			eventSetter.Handler = new System.Windows.Input.MouseEventHandler(MouseLeaveNode);
			((Style)target).Setters.Add(eventSetter);
			eventSetter = new EventSetter();
			eventSetter.Event = UIElement.MouseMoveEvent;
			eventSetter.Handler = new System.Windows.Input.MouseEventHandler(MouseMoveNode);
			((Style)target).Setters.Add(eventSetter);
		}
		if (connectionId == 26)
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
			((System.Windows.Controls.ContextMenu)target).Opened += MenuOpened;
		}
		if (connectionId == 27)
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
			((System.Windows.Controls.MenuItem)target).Click += CloseOthers;
		}
		if (connectionId == 28)
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
			((System.Windows.Controls.MenuItem)target).Click += Reopen;
		}
		if (connectionId == 29)
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
			((System.Windows.Controls.MenuItem)target).Click += Duplicate;
		}
		if (connectionId == 30)
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
			((System.Windows.Controls.MenuItem)target).Click += RenamePresentation;
		}
		if (connectionId == 31)
		{
			((System.Windows.Controls.MenuItem)target).Click += ShowInFolder;
		}
		if (connectionId == 32)
		{
			((System.Windows.Controls.MenuItem)target).Click += CopyPath;
		}
		if (connectionId == 33)
		{
			((System.Windows.Controls.MenuItem)target).Click += FinalizePagination;
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
			((System.Windows.Controls.MenuItem)target).Click += PrepareToShare;
		}
		if (connectionId == 35)
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
			((System.Windows.Controls.MenuItem)target).Click += SendPresentation;
		}
		if (connectionId == 36)
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
			((System.Windows.Controls.ContextMenu)target).Opened += SlideMenuOpened;
		}
		if (connectionId == 37)
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
		if (connectionId == 38)
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
			((System.Windows.Controls.MenuItem)target).Click += Hide;
		}
		if (connectionId == 39)
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
		if (connectionId == 40)
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
			((System.Windows.Controls.MenuItem)target).Click += MarkAsFacing;
		}
		if (connectionId == 41)
		{
			((System.Windows.Controls.MenuItem)target).Click += UnmarkAsFacing;
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
			((System.Windows.Controls.MenuItem)target).Click += InsertSlide;
		}
		if (connectionId == 43)
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
			((System.Windows.Controls.MenuItem)target).Click += DuplicateSlide;
		}
		if (connectionId == 44)
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
			((System.Windows.Controls.MenuItem)target).Click += DeleteSlide;
		}
		if (connectionId == 45)
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
			((System.Windows.Controls.MenuItem)target).Click += SendToEnd;
		}
		if (connectionId == 46)
		{
			((System.Windows.Controls.MenuItem)target).Click += RenameSlide;
		}
		if (connectionId == 47)
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
		if (connectionId == 48)
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
		if (connectionId == 49)
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
			((System.Windows.Controls.MenuItem)target).Click += EditHyperlink;
		}
		if (connectionId == 50)
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
			((System.Windows.Controls.MenuItem)target).Click += RemoveHyperlink;
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
			((TextBlock)target).PreviewMouseLeftButtonDown += HyperlinkClick;
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
			while (true)
			{
				switch (1)
				{
				case 0:
					continue;
				}
				break;
			}
			((System.Windows.Controls.MenuItem)target).Click += DeleteComment;
		}
		if (connectionId == 54)
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
		if (connectionId == 55)
		{
			((System.Windows.Controls.MenuItem)target).Click += DeleteNotes;
		}
		if (connectionId == 56)
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
			((System.Windows.Controls.MenuItem)target).SubmenuOpened += SubMenuOpened;
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
			((System.Windows.Controls.MenuItem)target).Click += RefreshLink;
		}
		if (connectionId == 59)
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
			((System.Windows.Controls.MenuItem)target).Click += ViewLinkSource;
		}
		if (connectionId == 60)
		{
			((System.Windows.Controls.MenuItem)target).Click += EditLink;
		}
		if (connectionId == 61)
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
			((System.Windows.Controls.MenuItem)target).Click += BreakLink;
		}
		if (connectionId == 62)
		{
			((System.Windows.Controls.MenuItem)target).Click += ConvertToPicture;
		}
		if (connectionId == 63)
		{
			((System.Windows.Controls.MenuItem)target).Click += HideShape;
		}
		if (connectionId == 64)
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
			((System.Windows.Controls.MenuItem)target).Click += UnhideShape;
		}
		if (connectionId == 65)
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
			((System.Windows.Controls.MenuItem)target).Click += DeleteShape;
		}
		if (connectionId == 66)
		{
			((System.Windows.Controls.ContextMenu)target).Opened += MenuOpened;
		}
		if (connectionId == 67)
		{
			((System.Windows.Controls.MenuItem)target).SubmenuOpened += SubMenuOpened;
		}
		if (connectionId == 68)
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
			((System.Windows.Controls.MenuItem)target).Click += RefreshLink;
		}
		if (connectionId == 69)
		{
			((System.Windows.Controls.MenuItem)target).Click += ViewLinkSource;
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
			((System.Windows.Controls.MenuItem)target).Click += EditLink;
		}
		if (connectionId == 71)
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
			((System.Windows.Controls.MenuItem)target).Click += BreakLink;
		}
		if (connectionId == 72)
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
			((System.Windows.Controls.MenuItem)target).Click += FixImageScaling;
		}
		if (connectionId == 73)
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
		if (connectionId == 74)
		{
			((System.Windows.Controls.MenuItem)target).Click += RotateRight90;
		}
		if (connectionId == 75)
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
			((System.Windows.Controls.MenuItem)target).Click += RotateLeft90;
		}
		if (connectionId == 76)
		{
			((System.Windows.Controls.MenuItem)target).Click += FlipVertical;
		}
		if (connectionId == 77)
		{
			((System.Windows.Controls.MenuItem)target).Click += FlipHorizontal;
		}
		if (connectionId == 78)
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
			((System.Windows.Controls.MenuItem)target).SubmenuOpened += SubMenuOpened;
		}
		if (connectionId == 79)
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
			((System.Windows.Controls.MenuItem)target).Click += BringForward;
		}
		if (connectionId == 80)
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
			((System.Windows.Controls.MenuItem)target).Click += BringToFront;
		}
		if (connectionId == 81)
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
		if (connectionId == 82)
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
			((System.Windows.Controls.MenuItem)target).Click += SendBackward;
		}
		if (connectionId == 83)
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
			((System.Windows.Controls.MenuItem)target).Click += SendToback;
		}
		if (connectionId == 84)
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
			((System.Windows.Controls.MenuItem)target).Click += HideShape;
		}
		if (connectionId == 85)
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
			((System.Windows.Controls.MenuItem)target).Click += UnhideShape;
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
			((System.Windows.Controls.MenuItem)target).Click += DeleteShape;
		}
		if (connectionId == 87)
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
		if (connectionId == 88)
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
		if (connectionId == 89)
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
			((System.Windows.Controls.MenuItem)target).Click += RefreshLink;
		}
		if (connectionId == 90)
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
			((System.Windows.Controls.MenuItem)target).Click += ViewLinkSource;
		}
		if (connectionId == 91)
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
			((System.Windows.Controls.MenuItem)target).Click += EditLink;
		}
		if (connectionId == 92)
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
			((System.Windows.Controls.MenuItem)target).Click += BreakLink;
		}
		if (connectionId == 93)
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
			((System.Windows.Controls.MenuItem)target).Click += HideShape;
		}
		if (connectionId == 94)
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
			((System.Windows.Controls.MenuItem)target).Click += UnhideShape;
		}
		if (connectionId == 95)
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
			((System.Windows.Controls.MenuItem)target).Click += DeleteShape;
		}
		if (connectionId == 96)
		{
			((System.Windows.Controls.ContextMenu)target).Opened += MenuOpened;
		}
		if (connectionId == 97)
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
			((System.Windows.Controls.MenuItem)target).SubmenuOpened += SubMenuOpened;
		}
		if (connectionId == 98)
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
			((System.Windows.Controls.MenuItem)target).Click += RefreshLink;
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
			((System.Windows.Controls.MenuItem)target).Click += ViewLinkSource;
		}
		if (connectionId == 100)
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
			((System.Windows.Controls.MenuItem)target).Click += EditLink;
		}
		if (connectionId == 101)
		{
			((System.Windows.Controls.MenuItem)target).Click += BreakLink;
		}
		if (connectionId == 102)
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
			((System.Windows.Controls.MenuItem)target).Click += ConvertToPicture;
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
			((System.Windows.Controls.MenuItem)target).Click += HideShape;
		}
		if (connectionId == 104)
		{
			((System.Windows.Controls.MenuItem)target).Click += UnhideShape;
		}
		if (connectionId == 105)
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
			((System.Windows.Controls.MenuItem)target).Click += DeleteShape;
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
			((System.Windows.Controls.ContextMenu)target).Opened += MenuOpened;
		}
		if (connectionId == 107)
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
			((System.Windows.Controls.MenuItem)target).Click += ConvertToPicture;
		}
		if (connectionId == 108)
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
			((System.Windows.Controls.MenuItem)target).Click += HideShape;
		}
		if (connectionId == 109)
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
			((System.Windows.Controls.MenuItem)target).Click += UnhideShape;
		}
		if (connectionId == 110)
		{
			((System.Windows.Controls.MenuItem)target).Click += DeleteShape;
		}
	}

	void IStyleConnector.Connect(int connectionId, object target)
	{
		//ILSpy generated this explicit interface implementation from .override directive in System_Windows_Markup_IStyleConnector_Connect
		this.System_Windows_Markup_IStyleConnector_Connect(connectionId, target);
	}

	[SpecialName]
	[CompilerGenerated]
	private void M()
	{
		bool value = chkShowAll.IsChecked.Value;
		PB.Settings.ExplorerShowAllPresentations = value;
		RemoveTreeViewSelectionChangedHandler();
		if (value)
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
			Pane.LoadPresentations(this.m_A);
			A(Pane.AllPresentations);
			C();
		}
		else
		{
			if (SinglePresentation == null)
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
				SinglePresentation = new ObservableCollection<PresentationItem>();
			}
			else
			{
				SinglePresentation.Clear();
			}
			D();
			A(SinglePresentation);
		}
		H();
		trvExplorer.Focus();
	}

	[SpecialName]
	[CompilerGenerated]
	private void N()
	{
		Events.Reset(this.m_A);
		RemoveTreeViewSelectionChangedHandler();
		if (chkShowAll.IsChecked == true)
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
			A(Pane.AllPresentations);
			C();
		}
		else
		{
			if (SinglePresentation == null)
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
				SinglePresentation = new ObservableCollection<PresentationItem>();
			}
			else
			{
				SinglePresentation.Clear();
			}
			D();
			A(SinglePresentation);
		}
		H();
	}

	[SpecialName]
	[CompilerGenerated]
	private bool A(PresentationItem A)
	{
		return A.Presentation == ThisPresentation;
	}

	[SpecialName]
	[CompilerGenerated]
	private void O()
	{
		H();
	}

	[SpecialName]
	[CompilerGenerated]
	private void P()
	{
		I();
	}

	[SpecialName]
	[CompilerGenerated]
	private bool A()
	{
		this.m_A = new DispatcherFrame();
		Dispatcher.PushFrame(this.m_A);
		return false;
	}

	[SpecialName]
	[CompilerGenerated]
	private bool A(SlideItem A)
	{
		return A.Slide == this.m_A.ActivePresentation.Windows[1].Selection.SlideRange[1];
	}
}
