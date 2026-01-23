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
using MacabacusMacros;
using MacabacusMacros.Libraries;
using MacabacusMacros.Libraries.Manage.Publish;
using MacabacusMacros.Libraries.Versioning;
using MacabacusMacros.Proofing.UI;
using MacabacusMacros.UI;
using MacabacusMacros.UI.FormsExtensions;
using Macabacus_Word.Library2.Versioning.Replace;
using Macabacus_Word.Shapes;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic.CompilerServices;

namespace Macabacus_Word.Library2.Versioning;

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
	private ListBox m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("grdFooter")]
	private Grid m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("chkOutdated")]
	private CheckBox m_A;

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
			A(XC.A(3391));
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

	internal virtual ListBox lbxResults
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
			ListBox listBox = this.m_A;
			if (listBox != null)
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

	internal virtual CheckBox chkOutdated
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
			this.m_A = value;
			button = this.m_A;
			if (button != null)
			{
				button.Click += value2;
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
				switch (2)
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
			switch (5)
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
		using (List<ShapeItem>.Enumerator enumerator = Check.LibraryObjects.GetEnumerator())
		{
			while (enumerator.MoveNext())
			{
				ShapeItem current = enumerator.Current;
				ContentItems.Add(((ContentItem)current).Clone());
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
		Check.LibraryObjects = null;
		SourceCollection = CollectionViewSource.GetDefaultView(ContentItems);
		SourceCollection.GroupDescriptions.Add(new PropertyGroupDescription(XC.A(8561)));
		SourceCollection.Filter = A;
		Grid grid = grdFooter;
		int visibility;
		if (ContentItems.Count <= 1)
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
			visibility = 2;
		}
		else
		{
			visibility = 0;
		}
		grid.Visibility = (Visibility)visibility;
		E();
		chkOutdated.IsChecked = global::A.N.Settings.ContentShowOnlyOutdated;
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
				switch (3)
				{
				case 0:
					continue;
				}
				break;
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
		//IL_001a: Unknown result type (might be due to invalid IL or missing references)
		if (chkOutdated.IsChecked == true)
		{
			return ((ContentItem)A).IsOutdated;
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
			switch (6)
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
		global::A.N.Settings.ContentShowOnlyOutdated = chkOutdated.IsChecked.Value;
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
		//IL_0074: Unknown result type (might be due to invalid IL or missing references)
		//IL_007e: Expected O, but got Unknown
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
				if (!e.IsRepeat)
				{
					lbxResults.KeyUp += NavKeyUp;
				}
				return;
			}
		case Key.Space:
			if (lbxResults.SelectedIndex <= -1)
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
			switch (2)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			ShapeItem shapeItem = (ShapeItem)lbxResults.SelectedItem;
			try
			{
				Navigate.A(RuntimeHelpers.GetObjectValue(shapeItem.Shape));
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				C(ex2.Message);
				C((ContentItem)(object)shapeItem);
				ProjectData.ClearProjectError();
			}
			shapeItem = null;
			M();
			return;
		}
	}

	private void E()
	{
		ListBox listBox = lbxResults;
		listBox.SelectionChanged += ListBoxSelectionChanged;
		listBox.GotKeyboardFocus += ListBoxGotKeyboardFocus;
		listBox.LostKeyboardFocus += ListBoxLostKeyboardFocus;
		_ = null;
		chkOutdated.Checked += OutdatedOnlyCheckChanged;
		chkOutdated.Unchecked += OutdatedOnlyCheckChanged;
	}

	private void F()
	{
		ListBox listBox = lbxResults;
		listBox.SelectionChanged -= ListBoxSelectionChanged;
		listBox.GotKeyboardFocus -= ListBoxGotKeyboardFocus;
		listBox.LostKeyboardFocus -= ListBoxLostKeyboardFocus;
		_ = null;
		chkOutdated.Checked -= OutdatedOnlyCheckChanged;
		chkOutdated.Unchecked -= OutdatedOnlyCheckChanged;
	}

	private void G()
	{
		new ComAwareEventInfo(typeof(ApplicationEvents4_Event), XC.A(1839)).AddEventHandler(PC.A.Application, new ApplicationEvents4_WindowSelectionChangeEventHandler(A));
	}

	private void H()
	{
		new ComAwareEventInfo(typeof(ApplicationEvents4_Event), XC.A(1839)).RemoveEventHandler(PC.A.Application, new ApplicationEvents4_WindowSelectionChangeEventHandler(A));
	}

	private void A(Microsoft.Office.Interop.Word.Selection A)
	{
		//IL_003f: Unknown result type (might be due to invalid IL or missing references)
		ListBox listBox = lbxResults;
		if (listBox.SelectedIndex > -1)
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
		if (this.A() || !Prompts.AskReplace(System.Windows.Window.GetWindow(this), A))
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
			HiddenFiles = new List<string>();
			this.m_A = false;
			bool flag = false;
			Microsoft.Office.Interop.Word.Application application = PC.A.Application;
			UndoRecord undoRecord = application.UndoRecord;
			undoRecord.StartCustomRecord(XC.A(8574));
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
			undoRecord.EndCustomRecord();
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
				application.ScreenRefresh();
				F(XC.A(8619));
			}
			undoRecord = null;
			application = null;
			return;
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
				switch (6)
				{
				case 0:
					continue;
				}
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				if (!A(XC.A(8652)))
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
					HiddenFiles = new List<string>();
					this.m_A = false;
					bool flag = false;
					Microsoft.Office.Interop.Word.Application application = PC.A.Application;
					UndoRecord undoRecord = application.UndoRecord;
					undoRecord.StartCustomRecord(XC.A(8846));
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
								switch (5)
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
							B(XC.A(8846));
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
					undoRecord.EndCustomRecord();
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
						application.ScreenRefresh();
						F(XC.A(8619));
					}
					undoRecord = null;
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
		Microsoft.Office.Interop.Word.Application application = PC.A.Application;
		ContentType contentType = A.ContentInfo.ContentType;
		switch (contentType - 3)
		{
		default:
			if ((int)contentType != 16)
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
			}
			else
			{
				Videos.A((ShapeItem)(object)A);
				B(XC.A(8995));
			}
			break;
		case 0:
		{
			ShapeItem a = (ShapeItem)(object)A;
			Microsoft.Office.Interop.Word.Application b = application;
			ref Microsoft.Office.Interop.PowerPoint.Application a2 = ref this.m_A;
			ref bool a3 = ref this.m_A;
			List<string> E = HiddenFiles;
			Macabacus_Word.Library2.Versioning.Replace.Shapes.A(a, b, ref a2, ref a3, ref E);
			HiddenFiles = E;
			B(XC.A(8909));
			break;
		}
		case 1:
			Macabacus_Word.Library2.Versioning.Replace.Images.A((ShapeItem)(object)A, application);
			B(XC.A(8952));
			break;
		case 2:
			Charts.A((ShapeItem)(object)A);
			B(XC.A(9038));
			break;
		}
		string author = Core.GetAuthor(application.ActiveDocument);
		string c = Content.RightNow();
		object objectValue = RuntimeHelpers.GetObjectValue(((ShapeItem)(object)A).Shape);
		string alternativeText;
		if (objectValue is Microsoft.Office.Interop.Word.Shape)
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
			alternativeText = ((Microsoft.Office.Interop.Word.Shape)objectValue).AlternativeText;
		}
		else
		{
			alternativeText = ((InlineShape)objectValue).AlternativeText;
		}
		ContentItem val;
		object[] array;
		bool[] array2;
		NewLateBinding.LateCall(null, typeof(Tagging), XC.A(9081), array = new object[2]
		{
			objectValue,
			(val = A).LatestVersion
		}, null, null, array2 = new bool[2] { true, true }, IgnoreReturn: true);
		if (array2[0])
		{
			objectValue = RuntimeHelpers.GetObjectValue(array[0]);
		}
		if (array2[1])
		{
			val.LatestVersion = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[1]), typeof(int));
		}
		NewLateBinding.LateCall(null, typeof(Tagging), XC.A(9122), array = new object[2] { objectValue, 0 }, null, null, array2 = new bool[2] { true, false }, IgnoreReturn: true);
		if (array2[0])
		{
			objectValue = RuntimeHelpers.GetObjectValue(array[0]);
		}
		Tagging.A(alternativeText, Content.XML_MODIFIED_BY, author);
		Tagging.A(alternativeText, Content.XML_MODIFIED_AT, c);
		A.RebaseVersion();
		application = null;
	}

	private void I()
	{
		if (this.m_A == null)
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
			J();
			K();
			return;
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
				switch (5)
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
			string alternativeText = default(string);
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
				ShapeItem shapeItem = (ShapeItem)((Button)sender).DataContext;
				if (Prompts.AskIgnore(System.Windows.Window.GetWindow(this), (ContentItem)(object)shapeItem))
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
					UndoRecord undoRecord = PC.A.Application.UndoRecord;
					undoRecord.StartCustomRecord(XC.A(9163));
					try
					{
						Type typeFromHandle = typeof(Tagging);
						string memberName = XC.A(9122);
						ShapeItem shapeItem2;
						ShapeItem shapeItem3;
						object[] obj = new object[2]
						{
							(shapeItem2 = shapeItem).Shape,
							((ContentItem)(shapeItem3 = shapeItem)).LatestVersion
						};
						object[] array = obj;
						bool[] obj2 = new bool[2] { true, true };
						bool[] array2 = obj2;
						NewLateBinding.LateCall(null, typeFromHandle, memberName, obj, null, null, obj2, IgnoreReturn: true);
						if (array2[0])
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
							shapeItem2.Shape = RuntimeHelpers.GetObjectValue(RuntimeHelpers.GetObjectValue(array[0]));
						}
						if (array2[1])
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
							((ContentItem)shapeItem3).LatestVersion = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[1]), typeof(int));
						}
						((ContentItem)shapeItem).IgnoreVersion();
						if (chkOutdated.IsChecked == true)
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
							((ContentItem)shapeItem).IsSelected = false;
						}
						if (shapeItem.Shape is InlineShape)
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
							alternativeText = ((InlineShape)shapeItem.Shape).AlternativeText;
						}
						else if (shapeItem.Shape is Microsoft.Office.Interop.Word.Shape)
						{
							alternativeText = ((Microsoft.Office.Interop.Word.Shape)shapeItem.Shape).AlternativeText;
						}
						for (int i = ContentItems.Count - 1; i >= 0; i += -1)
						{
							if ((object)ContentItems[i] == shapeItem)
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
							ShapeItem shapeItem4 = (ShapeItem)(object)ContentItems[i];
							if (shapeItem4.Shape is InlineShape)
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
								if (Operators.CompareString(((InlineShape)shapeItem4.Shape).AlternativeText, alternativeText, TextCompare: false) == 0)
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
									ContentItems[i].IgnoreVersion();
								}
							}
							else if (shapeItem4.Shape is Microsoft.Office.Interop.Word.Shape)
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
								if (Operators.CompareString(((Microsoft.Office.Interop.Word.Shape)shapeItem4.Shape).AlternativeText, alternativeText, TextCompare: false) == 0)
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
									ContentItems[i].IgnoreVersion();
								}
							}
							shapeItem4 = null;
						}
						while (true)
						{
							switch (3)
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
					undoRecord.EndCustomRecord();
					undoRecord = null;
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
				Type typeFromHandle = typeof(Tagging);
				string memberName = XC.A(9210);
				ShapeItem shapeItem2;
				object[] obj = new object[1] { (shapeItem2 = shapeItem).Shape };
				object[] array = obj;
				bool[] obj2 = new bool[1] { true };
				bool[] array2 = obj2;
				NewLateBinding.LateCall(null, typeFromHandle, memberName, obj, null, null, obj2, IgnoreReturn: true);
				if (array2[0])
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
					shapeItem2.Shape = RuntimeHelpers.GetObjectValue(RuntimeHelpers.GetObjectValue(array[0]));
				}
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
					switch (6)
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
		if (PC.A.Application.ActiveDocument.Final)
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
					D(XC.A(9237));
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
		clsReporting.LogActivity((ActivityApp)3, (ActivityCategory)6, XC.A(9367));
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

	[DebuggerNonUserCode]
	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
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
			Uri resourceLocator = new Uri(XC.A(9412), UriKind.Relative);
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
	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	[DebuggerNonUserCode]
	public void System_Windows_Markup_IComponentConnector_Connect(int connectionId, object target)
	{
		if (connectionId == 1)
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
					lbxResults = (ListBox)target;
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
					grdFooter = (Grid)target;
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
					chkOutdated = (CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 7)
		{
			while (true)
			{
				switch (2)
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
				switch (5)
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

	[EditorBrowsable(EditorBrowsableState.Never)]
	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	[DebuggerNonUserCode]
	public void System_Windows_Markup_IStyleConnector_Connect(int connectionId, object target)
	{
		if (connectionId == 2)
		{
			((Button)target).Click += IgnoreContent;
		}
		if (connectionId == 3)
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
			((Button)target).Click += UnlinkContent;
		}
		if (connectionId != 4)
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
