using System;
using System.CodeDom.Compiler;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Media.Animation;
using A;
using ExcelAddIn1.Library2.Versioning.Replace;
using ExcelAddIn1.Shapes;
using MacabacusMacros;
using MacabacusMacros.Libraries;
using MacabacusMacros.Libraries.Manage.Publish;
using MacabacusMacros.Libraries.Versioning;
using MacabacusMacros.Proofing.UI;
using MacabacusMacros.UI;
using MacabacusMacros.UI.FormsExtensions;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Library2.Versioning;

[DesignerGenerated]
public sealed class wpfVersions : UserControl, INotifyPropertyChanged, IComponentConnector, IStyleConnector
{
	[Serializable]
	[CompilerGenerated]
	internal sealed class _Closure_0024__
	{
		public static readonly _Closure_0024__ A;

		public static Func<ContentItem, bool> A;

		static _Closure_0024__()
		{
			_Closure_0024__.A = new _Closure_0024__();
		}

		[SpecialName]
		internal bool A(ContentItem A)
		{
			return A.IsOutdated;
		}
	}

	[CompilerGenerated]
	private PropertyChangedEventHandler m_A;

	private Microsoft.Office.Interop.PowerPoint.Application m_A;

	private bool m_A;

	private ICollectionView m_A;

	[CompilerGenerated]
	private ObservableCollection<ContentItem> m_A;

	[CompilerGenerated]
	private bool m_B;

	[CompilerGenerated]
	private ContentItem m_A;

	[CompilerGenerated]
	private List<string> m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("lbxResults")]
	private System.Windows.Controls.ListBox m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("grdFooter")]
	private Grid m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("chkOutdated")]
	private System.Windows.Controls.CheckBox m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("btnReplaceAll")]
	private Button m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("grdNoContent")]
	private Grid m_B;

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
			A(VH.A(36261));
		}
	}

	private ObservableCollection<ContentItem> ContentItems
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

	private bool CollapseAnimationRunning
	{
		[CompilerGenerated]
		get
		{
			return this.m_B;
		}
		[CompilerGenerated]
		set
		{
			this.m_B = value;
		}
	}

	private ContentItem ItemQueuedToRemove
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

	private List<string> HiddenFiles
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

	internal virtual System.Windows.Controls.ListBox lbxResults
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
			KeyEventHandler value2 = lbxResults_KeyDown;
			System.Windows.Controls.ListBox listBox = this.m_A;
			if (listBox != null)
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
				listBox.PreviewKeyDown -= value2;
			}
			this.m_A = value;
			listBox = this.m_A;
			if (listBox == null)
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
				listBox.PreviewKeyDown += value2;
				return;
			}
		}
	}

	internal virtual Grid grdFooter
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

	internal virtual System.Windows.Controls.CheckBox chkOutdated
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

	internal virtual Button btnReplaceAll
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
			RoutedEventHandler value2 = ReplaceAll;
			Button button = this.m_A;
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
			this.m_A = value;
			button = this.m_A;
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

	internal virtual Grid grdNoContent
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
				switch (4)
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

	public wpfVersions()
	{
		base.Unloaded += wpfVersions_Unloaded;
		this.m_A = null;
		InitializeComponent();
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

	private void wpfVersions_Unloaded(object sender, RoutedEventArgs e)
	{
		F();
	}

	internal void A()
	{
		ContentItems = new ObservableCollection<ContentItem>();
		foreach (ShapeItem libraryShape in Check.LibraryShapes)
		{
			ContentItems.Add(((ContentItem)libraryShape).Clone());
		}
		Check.LibraryShapes = null;
		SourceCollection = CollectionViewSource.GetDefaultView(ContentItems);
		SourceCollection.GroupDescriptions.Add(new PropertyGroupDescription(VH.A(84167)));
		SourceCollection.Filter = A;
		grdFooter.Visibility = ((ContentItems.Count <= 1) ? Visibility.Collapsed : Visibility.Visible);
		E();
		chkOutdated.IsChecked = global::A.K.Settings.ContentShowOnlyOutdated;
		Button button = btnReplaceAll;
		ObservableCollection<ContentItem> contentItems = ContentItems;
		Func<ContentItem, bool> predicate;
		if (_Closure_0024__.A == null)
		{
			predicate = (_Closure_0024__.A = [SpecialName] (ContentItem A) => A.IsOutdated);
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			predicate = _Closure_0024__.A;
		}
		button.IsEnabled = contentItems.Where(predicate).Any();
	}

	internal void B()
	{
		F();
		H();
	}

	private bool A(object A)
	{
		//IL_002f: Unknown result type (might be due to invalid IL or missing references)
		if (chkOutdated.IsChecked == true)
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
					return ((ContentItem)A).IsOutdated;
				}
			}
		}
		return true;
	}

	private void C()
	{
		if (SourceCollection == null)
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
			SourceCollection.Refresh();
			return;
		}
	}

	private void OutdatedOnlyCheckChanged(object sender, RoutedEventArgs e)
	{
		C();
		global::A.K.Settings.ContentShowOnlyOutdated = chkOutdated.IsChecked.Value;
	}

	private void ListBoxGotKeyboardFocus(object sender, KeyboardFocusChangedEventArgs e)
	{
		H();
	}

	private void ListBoxLostKeyboardFocus(object sender, KeyboardFocusChangedEventArgs e)
	{
		G();
	}

	private void lbxResults_KeyDown(object sender, KeyEventArgs e)
	{
		//IL_007a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0084: Expected O, but got Unknown
		switch (e.Key)
		{
		case Key.Left:
		case Key.Up:
		case Key.Right:
		case Key.Down:
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
				if (e.IsRepeat)
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
					lbxResults.KeyUp += NavKeyUp;
					return;
				}
			}
		case Key.Space:
			if (lbxResults.SelectedIndex <= -1)
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
				A((ContentItem)lbxResults.SelectedItem);
				return;
			}
		}
	}

	private void NavKeyUp(object sender, KeyEventArgs e)
	{
		lbxResults.KeyUp -= NavKeyUp;
		D();
		e.Handled = true;
	}

	private void ListBoxSelectionChanged(object sender, SelectionChangedEventArgs e)
	{
		D();
	}

	private void D()
	{
		if (lbxResults.SelectedIndex <= -1)
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
			ShapeItem shapeItem = (ShapeItem)lbxResults.SelectedItem;
			Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
			application.ScreenUpdating = false;
			application.EnableEvents = false;
			try
			{
				Navigate.A(shapeItem.Shape);
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				try
				{
					if (shapeItem.Shape.TopLeftCell.Worksheet.ProtectContents)
					{
						while (true)
						{
							switch (4)
							{
							case 0:
								break;
							default:
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
				C(ex2.Message);
				C((ContentItem)(object)shapeItem);
				ProjectData.ClearProjectError();
			}
			application.ScreenUpdating = true;
			application.EnableEvents = true;
			application = null;
			shapeItem = null;
			M();
			return;
		}
	}

	private void E()
	{
		System.Windows.Controls.ListBox listBox = lbxResults;
		listBox.SelectionChanged += ListBoxSelectionChanged;
		listBox.GotKeyboardFocus += ListBoxGotKeyboardFocus;
		listBox.LostKeyboardFocus += ListBoxLostKeyboardFocus;
		_ = null;
		chkOutdated.Checked += OutdatedOnlyCheckChanged;
		chkOutdated.Unchecked += OutdatedOnlyCheckChanged;
	}

	private void F()
	{
		System.Windows.Controls.ListBox listBox = lbxResults;
		listBox.SelectionChanged -= ListBoxSelectionChanged;
		listBox.GotKeyboardFocus -= ListBoxGotKeyboardFocus;
		listBox.LostKeyboardFocus -= ListBoxLostKeyboardFocus;
		_ = null;
		chkOutdated.Checked -= OutdatedOnlyCheckChanged;
		chkOutdated.Unchecked -= OutdatedOnlyCheckChanged;
	}

	private void G()
	{
		new ComAwareEventInfo(typeof(AppEvents_Event), VH.A(1700)).AddEventHandler(MH.A.Application, new AppEvents_SheetSelectionChangeEventHandler(A));
	}

	private void H()
	{
		new ComAwareEventInfo(typeof(AppEvents_Event), VH.A(1700)).RemoveEventHandler(MH.A.Application, new AppEvents_SheetSelectionChangeEventHandler(A));
	}

	private void A(object A, Range B)
	{
		//IL_003f: Unknown result type (might be due to invalid IL or missing references)
		System.Windows.Controls.ListBox listBox = lbxResults;
		if (listBox.SelectedIndex > -1)
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
			listBox.SelectionChanged -= ListBoxSelectionChanged;
			((ContentItem)listBox.SelectedItem).IsSelected = false;
			listBox.SelectionChanged += ListBoxSelectionChanged;
		}
		listBox = null;
	}

	private void ReplaceContent(object sender, RoutedEventArgs e)
	{
		//IL_000c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0016: Expected O, but got Unknown
		A((ContentItem)((Button)sender).DataContext);
	}

	private void A(ContentItem A)
	{
		if (this.A())
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
			if (!Prompts.AskReplace(System.Windows.Window.GetWindow(this), A))
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
				HiddenFiles = new List<string>();
				this.m_A = false;
				bool flag = false;
				Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
				application.ScreenUpdating = false;
				try
				{
					B(A);
					flag = true;
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					C(ex2.Message);
					ProjectData.ClearProjectError();
				}
				application.ScreenUpdating = true;
				I();
				if (flag)
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
					F(VH.A(84180));
				}
				application = null;
				return;
			}
		}
	}

	private void ReplaceAll(object sender, RoutedEventArgs e)
	{
		if (A())
		{
			return;
		}
		checked
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
				if (!A(VH.A(84213)))
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
					HiddenFiles = new List<string>();
					this.m_A = false;
					bool flag = false;
					Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
					application.ScreenUpdating = false;
					try
					{
						for (int i = ContentItems.Count - 1; i >= 0; i += -1)
						{
							if (!ContentItems[i].IsOutdated)
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
							B(ContentItems[i]);
						}
						while (true)
						{
							switch (1)
							{
							case 0:
								continue;
							}
							flag = true;
							B(VH.A(84407));
							break;
						}
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						C(ex2.Message);
						ProjectData.ClearProjectError();
					}
					application.ScreenUpdating = true;
					I();
					if (flag)
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
						F(VH.A(84180));
					}
					application = null;
					return;
				}
			}
		}
	}

	private void B(ContentItem A)
	{
		//IL_0010: Unknown result type (might be due to invalid IL or missing references)
		//IL_0015: Unknown result type (might be due to invalid IL or missing references)
		//IL_001a: Unknown result type (might be due to invalid IL or missing references)
		//IL_001c: Unknown result type (might be due to invalid IL or missing references)
		//IL_001f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0031: Expected I4, but got Unknown
		//IL_0031: Unknown result type (might be due to invalid IL or missing references)
		//IL_0035: Invalid comparison between Unknown and I4
		Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
		ContentType contentType = A.ContentInfo.ContentType;
		switch (contentType - 3)
		{
		default:
			if ((int)contentType != 16)
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
			}
			else
			{
				Videos.A((ShapeItem)(object)A);
				B(VH.A(84556));
			}
			break;
		case 0:
		{
			ShapeItem a = (ShapeItem)(object)A;
			Microsoft.Office.Interop.Excel.Application b = application;
			ref Microsoft.Office.Interop.PowerPoint.Application a2 = ref this.m_A;
			ref bool a3 = ref this.m_A;
			List<string> E = HiddenFiles;
			ExcelAddIn1.Library2.Versioning.Replace.Shapes.A(a, b, ref a2, ref a3, ref E);
			HiddenFiles = E;
			B(VH.A(84470));
			break;
		}
		case 1:
			ExcelAddIn1.Library2.Versioning.Replace.Images.A((ShapeItem)(object)A, application);
			B(VH.A(84513));
			break;
		case 2:
			ExcelAddIn1.Library2.Versioning.Replace.Charts.A((ShapeItem)(object)A);
			B(VH.A(84599));
			break;
		}
		string author = Core.GetAuthor(application.ActiveWorkbook);
		string c = Content.RightNow();
		Microsoft.Office.Interop.Excel.Shape shape = ((ShapeItem)(object)A).Shape;
		string alternativeText = shape.AlternativeText;
		Tagging.UpdateCurrentVersion(shape, A.LatestVersion);
		Tagging.UpdateIgnoredVersion(shape, 0);
		Tagging.A(alternativeText, Content.XML_MODIFIED_BY, author);
		Tagging.A(alternativeText, Content.XML_MODIFIED_AT, c);
		A.RebaseVersion();
		application = null;
	}

	private void I()
	{
		if (this.m_A != null)
		{
			J();
			K();
		}
	}

	private void J()
	{
		using (List<string>.Enumerator enumerator = HiddenFiles.GetEnumerator())
		{
			while (enumerator.MoveNext())
			{
				string current = enumerator.Current;
				try
				{
					this.m_A.Presentations[Path.GetFileName(current)].Close();
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
					continue;
				}
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				break;
			}
		}
		HiddenFiles = null;
	}

	private void K()
	{
		if (this.m_A)
		{
			this.m_A.Quit();
			this.m_A = null;
		}
	}

	private void IgnoreContent(object sender, RoutedEventArgs e)
	{
		if (A())
		{
			return;
		}
		checked
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
				ShapeItem shapeItem = (ShapeItem)((Button)sender).DataContext;
				if (Prompts.AskIgnore(System.Windows.Window.GetWindow(this), (ContentItem)(object)shapeItem))
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
					object objectValue = RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(MH.A.Application, null, VH.A(84642), new object[0], null, null, null));
					NewLateBinding.LateCall(objectValue, null, VH.A(84663), new object[1] { VH.A(84698) }, null, null, null, IgnoreReturn: true);
					try
					{
						Tagging.UpdateIgnoredVersion(shapeItem.Shape, ((ContentItem)shapeItem).LatestVersion);
						((ContentItem)shapeItem).IgnoreVersion();
						if (chkOutdated.IsChecked == true)
						{
							((ContentItem)shapeItem).IsSelected = false;
						}
						string alternativeText = shapeItem.Shape.AlternativeText;
						for (int i = ContentItems.Count - 1; i >= 0; i += -1)
						{
							if ((object)ContentItems[i] == shapeItem)
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
							if (Operators.CompareString(((ShapeItem)(object)ContentItems[i]).Shape.AlternativeText, alternativeText, TextCompare: false) != 0)
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
							ContentItems[i].IgnoreVersion();
						}
						while (true)
						{
							switch (7)
							{
							case 0:
								continue;
							}
							C();
							N();
							break;
						}
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						C(ex2.Message);
						ProjectData.ClearProjectError();
					}
					NewLateBinding.LateCall(objectValue, null, VH.A(84745), new object[0], null, null, null, IgnoreReturn: true);
					objectValue = null;
				}
				shapeItem = null;
				return;
			}
		}
	}

	private void UnlinkContent(object sender, RoutedEventArgs e)
	{
		ShapeItem shapeItem = (ShapeItem)((Button)sender).DataContext;
		if (!A())
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
			if (Prompts.AskUnlink(System.Windows.Window.GetWindow(this), (ContentItem)(object)shapeItem))
			{
				Tagging.UnlinkContent(shapeItem.Shape);
				C((ContentItem)(object)shapeItem);
			}
		}
		shapeItem = null;
	}

	private void C(ContentItem A)
	{
		ItemQueuedToRemove = A;
		ListBoxItem listBoxItem;
		try
		{
			listBoxItem = (ListBoxItem)lbxResults.ItemContainerGenerator.ContainerFromItem(A);
			if (listBoxItem != null)
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
					this.A(listBoxItem);
					break;
				}
			}
			else
			{
				L();
			}
		}
		catch (NullReferenceException ex)
		{
			ProjectData.SetProjectError(ex);
			NullReferenceException ex2 = ex;
			clsReporting.LogException((Exception)ex2);
			ProjectData.ClearProjectError();
		}
		listBoxItem = null;
	}

	private void A(ListBoxItem A)
	{
		CollapseAnimationRunning = true;
		DoubleAnimation collapseAnimation = Pane.GetCollapseAnimation(A);
		collapseAnimation.Completed += CollapseComplete;
		Pane.CollapseListBoxItem(A, collapseAnimation);
		collapseAnimation = null;
	}

	private void CollapseComplete(object sender, EventArgs e)
	{
		L();
		CollapseAnimationRunning = false;
	}

	private void L()
	{
		ContentItems.Remove(ItemQueuedToRemove);
		ItemQueuedToRemove = null;
	}

	private bool A()
	{
		if (MH.A.Application.ActiveWorkbook.Final)
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
					D(VH.A(84776));
					return true;
				}
			}
		}
		return false;
	}

	private void M()
	{
		Forms.FocusListBoxItem(lbxResults, true);
	}

	private void B(string A)
	{
		clsReporting.LogActivity((ActivityApp)3, (ActivityCategory)6, A);
	}

	private void N()
	{
		clsReporting.LogActivity((ActivityApp)3, (ActivityCategory)6, VH.A(84906));
	}

	private bool A(string A)
	{
		return UIFormsExtensions.AskOkCancel(System.Windows.Window.GetWindow(this), A);
	}

	private void C(string A)
	{
		Forms.ErrorMessage(System.Windows.Window.GetWindow(this), A);
	}

	private void D(string A)
	{
		Forms.WarningMessage(System.Windows.Window.GetWindow(this), A);
	}

	private void E(string A)
	{
		Forms.InfoMessage(System.Windows.Window.GetWindow(this), A);
	}

	private void F(string A)
	{
		Forms.SuccessMessage(System.Windows.Window.GetWindow(this), A);
	}

	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	[DebuggerNonUserCode]
	public void InitializeComponent()
	{
		if (!this.m_C)
		{
			this.m_C = true;
			Uri resourceLocator = new Uri(VH.A(84951), UriKind.Relative);
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
					lbxResults = (System.Windows.Controls.ListBox)target;
					return;
				}
			}
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
					grdFooter = (Grid)target;
					return;
				}
			}
		}
		if (connectionId == 6)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					chkOutdated = (System.Windows.Controls.CheckBox)target;
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
					btnReplaceAll = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 8)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					grdNoContent = (Grid)target;
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

	[DebuggerNonUserCode]
	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	[EditorBrowsable(EditorBrowsableState.Never)]
	public void System_Windows_Markup_IStyleConnector_Connect(int connectionId, object target)
	{
		if (connectionId == 2)
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
			((Button)target).Click += IgnoreContent;
		}
		if (connectionId == 3)
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
			((Button)target).Click += UnlinkContent;
		}
		if (connectionId != 4)
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
			((Button)target).Click += ReplaceContent;
			return;
		}
	}

	void IStyleConnector.Connect(int connectionId, object target)
	{
		//ILSpy generated this explicit interface implementation from .override directive in System_Windows_Markup_IStyleConnector_Connect
		this.System_Windows_Markup_IStyleConnector_Connect(connectionId, target);
	}
}
