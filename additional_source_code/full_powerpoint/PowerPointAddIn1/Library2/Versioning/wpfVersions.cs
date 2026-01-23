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
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.Library2.Versioning.Replace;
using PowerPointAddIn1.Links;
using PowerPointAddIn1.Slides;

namespace PowerPointAddIn1.Library2.Versioning;

[DesignerGenerated]
public sealed class wpfVersions : UserControl, INotifyPropertyChanged, IComponentConnector, IStyleConnector
{
	[Serializable]
	[CompilerGenerated]
	internal sealed class _Closure_0024__
	{
		public static readonly _Closure_0024__ A;

		public static Func<ContentItem, bool> A;

		public static Func<Slide, int> A;

		static _Closure_0024__()
		{
			_Closure_0024__.A = new _Closure_0024__();
		}

		[SpecialName]
		internal bool A(ContentItem A)
		{
			return A.IsOutdated;
		}

		[SpecialName]
		internal int A(Slide A)
		{
			return A.SlideIndex;
		}
	}

	[CompilerGenerated]
	private PropertyChangedEventHandler m_A;

	private ICollectionView m_A;

	[CompilerGenerated]
	private ObservableCollection<ContentItem> m_A;

	[CompilerGenerated]
	private bool m_A;

	[CompilerGenerated]
	private ContentItem m_A;

	[CompilerGenerated]
	private List<string> m_A;

	[AccessedThroughProperty("lbxResults")]
	[CompilerGenerated]
	private ListBox m_A;

	[AccessedThroughProperty("grdFooter")]
	[CompilerGenerated]
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

	private bool m_B;

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
			return this.m_A;
		}
		[CompilerGenerated]
		set
		{
			this.m_A = value;
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
				switch (6)
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
			switch (4)
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
		using (List<SlideItem>.Enumerator enumerator = Check.LibrarySlides.GetEnumerator())
		{
			while (enumerator.MoveNext())
			{
				SlideItem current = enumerator.Current;
				ContentItems.Add(((ContentItem)current).Clone());
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
				break;
			}
		}
		Check.LibrarySlides = null;
		foreach (ShapeItem libraryShape in Check.LibraryShapes)
		{
			ContentItems.Add(((ContentItem)libraryShape).Clone());
		}
		Check.LibraryShapes = null;
		SourceCollection = CollectionViewSource.GetDefaultView(ContentItems);
		SourceCollection.GroupDescriptions.Add(new PropertyGroupDescription(AH.A(52438)));
		SourceCollection.Filter = A;
		grdFooter.Visibility = ((ContentItems.Count <= 1) ? Visibility.Collapsed : Visibility.Visible);
		E();
		chkOutdated.IsChecked = PB.Settings.ContentShowOnlyOutdated;
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
				switch (6)
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
		//IL_002f: Unknown result type (might be due to invalid IL or missing references)
		if (chkOutdated.IsChecked == true)
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
			switch (4)
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
		PB.Settings.ContentShowOnlyOutdated = chkOutdated.IsChecked.Value;
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
		//IL_0084: Unknown result type (might be due to invalid IL or missing references)
		//IL_008e: Expected O, but got Unknown
		Key key = e.Key;
		if (key != Key.Space)
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
					if ((uint)(key - 23) <= 3u)
					{
						while (true)
						{
							switch (3)
							{
							case 0:
								break;
							default:
								if (!e.IsRepeat)
								{
									while (true)
									{
										switch (5)
										{
										case 0:
											break;
										default:
											lbxResults.KeyUp += NavKeyUp;
											return;
										}
									}
								}
								return;
							}
						}
					}
					return;
				}
			}
		}
		if (lbxResults.SelectedIndex <= -1)
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
			A((ContentItem)lbxResults.SelectedItem);
			return;
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
		//IL_0039: Unknown result type (might be due to invalid IL or missing references)
		//IL_003f: Expected O, but got Unknown
		if (lbxResults.SelectedIndex <= -1)
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
			K();
			ContentItem val = (ContentItem)lbxResults.SelectedItem;
			try
			{
				if (val is SlideItem)
				{
					while (true)
					{
						switch (3)
						{
						case 0:
							continue;
						}
						try
						{
							Microsoft.Office.Interop.PowerPoint.Presentation activePresentation = NG.A.Application.ActivePresentation;
							List<Slide> slides = ((SlideItem)(object)val).Slides;
							Func<Slide, int> selector;
							if (_Closure_0024__.A == null)
							{
								selector = (_Closure_0024__.A = [SpecialName] (Slide A) => A.SlideIndex);
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
								selector = _Closure_0024__.A;
							}
							Helpers.SelectMultipleSlides(activePresentation, slides.Select(selector));
						}
						catch (Exception ex)
						{
							ProjectData.SetProjectError(ex);
							Exception ex2 = ex;
							ProjectData.ClearProjectError();
						}
						break;
					}
				}
				else if (NG.A.Application.ActiveWindow.Panes[1].ViewType != PpViewType.ppViewNormal)
				{
					while (true)
					{
						switch (4)
						{
						case 0:
							continue;
						}
						try
						{
							A(((ShapeItem)(object)val).Shape);
						}
						catch (Exception ex3)
						{
							ProjectData.SetProjectError(ex3);
							Exception ex4 = ex3;
							C(ex4.Message);
							D(val);
							ProjectData.ClearProjectError();
						}
						break;
					}
				}
			}
			catch (Exception ex5)
			{
				ProjectData.SetProjectError(ex5);
				Exception ex6 = ex5;
				C(ex6.Message);
				ProjectData.ClearProjectError();
			}
			val = null;
			L();
			return;
		}
	}

	private void A(Slide A)
	{
		((Microsoft.Office.Interop.PowerPoint.Presentation)A.Parent).Windows[1].View.GotoSlide(A.SlideIndex);
	}

	private void A(Shape A)
	{
		Slide slideFromShape = clsPowerPoint.GetSlideFromShape(A);
		Slide slide = null;
		Microsoft.Office.Interop.PowerPoint.Application application = NG.A.Application;
		try
		{
			slide = application.ActiveWindow.Selection.SlideRange[1];
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		if (slide != slideFromShape)
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
			application.ActivePresentation.Windows[1].Activate();
			application.ActiveWindow.View.GotoSlide(slideFromShape.SlideIndex);
		}
		application = null;
		A.Select();
		slideFromShape = null;
		slide = null;
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
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(12762)).AddEventHandler(NG.A.Application, new EApplication_WindowSelectionChangeEventHandler(A));
	}

	private void H()
	{
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(12762)).RemoveEventHandler(NG.A.Application, new EApplication_WindowSelectionChangeEventHandler(A));
	}

	private void A(Selection A)
	{
		//IL_002c: Unknown result type (might be due to invalid IL or missing references)
		ListBox listBox = lbxResults;
		if (listBox.SelectedIndex > -1)
		{
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
			switch (5)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			if (!Prompts.AskReplace(Window.GetWindow(this), A))
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
				HiddenFiles = new List<string>();
				H();
				NG.A.Application.StartNewUndoEntry();
				try
				{
					B(A);
					F(AH.A(64058));
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					C(ex2.Message);
					ProjectData.ClearProjectError();
				}
				I();
				G();
				return;
			}
		}
	}

	private void ReplaceAll(object sender, RoutedEventArgs e)
	{
		if (A() || !A(AH.A(64091)))
		{
			return;
		}
		checked
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
				HiddenFiles = new List<string>();
				H();
				NG.A.Application.StartNewUndoEntry();
				try
				{
					for (int i = ContentItems.Count - 1; i >= 0; i += -1)
					{
						if (ContentItems[i].IsOutdated)
						{
							B(ContentItems[i]);
						}
					}
					while (true)
					{
						switch (7)
						{
						case 0:
							continue;
						}
						B(AH.A(64293));
						F(AH.A(64058));
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
				I();
				G();
				return;
			}
		}
	}

	private void B(ContentItem A)
	{
		//IL_000e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0013: Unknown result type (might be due to invalid IL or missing references)
		//IL_0015: Unknown result type (might be due to invalid IL or missing references)
		//IL_001a: Unknown result type (might be due to invalid IL or missing references)
		//IL_001b: Unknown result type (might be due to invalid IL or missing references)
		//IL_001d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0043: Expected I4, but got Unknown
		//IL_0043: Unknown result type (might be due to invalid IL or missing references)
		//IL_0046: Invalid comparison between Unknown and I4
		Microsoft.Office.Interop.PowerPoint.Application application = NG.A.Application;
		ContentType contentType = A.ContentInfo.ContentType;
		switch (contentType - 1)
		{
		default:
			if ((int)contentType == 16)
			{
				Videos.A((ShapeItem)(object)A);
				B(AH.A(64485));
			}
			break;
		case 0:
		case 1:
		{
			SlideItem a3 = (SlideItem)(object)A;
			Microsoft.Office.Interop.PowerPoint.Application b3 = application;
			List<string> C = HiddenFiles;
			PowerPointAddIn1.Library2.Versioning.Replace.Slides.A(a3, b3, ref C);
			HiddenFiles = C;
			B(AH.A(64356));
			break;
		}
		case 2:
		{
			ShapeItem a2 = (ShapeItem)(object)A;
			Microsoft.Office.Interop.PowerPoint.Application b2 = application;
			List<string> C = HiddenFiles;
			PowerPointAddIn1.Library2.Versioning.Replace.Shapes.A(a2, b2, ref C);
			HiddenFiles = C;
			B(AH.A(64399));
			break;
		}
		case 7:
		{
			ShapeItem a = (ShapeItem)(object)A;
			Microsoft.Office.Interop.PowerPoint.Application b = application;
			List<string> C = HiddenFiles;
			PowerPointAddIn1.Library2.Versioning.Replace.Shapes.A(a, b, ref C);
			HiddenFiles = C;
			B(AH.A(64399));
			break;
		}
		case 3:
			Images.A((ShapeItem)(object)A, application);
			B(AH.A(64442));
			break;
		case 4:
			Charts.A((ShapeItem)(object)A);
			B(AH.A(64528));
			break;
		case 5:
		case 6:
			break;
		}
		string author = Core.GetAuthor(application.ActivePresentation);
		string c = Content.RightNow();
		if (A is SlideItem)
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
			using List<Slide>.Enumerator enumerator = ((SlideItem)(object)A).Slides.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Slide current = enumerator.Current;
				Tagging.A(current.Tags, A.LatestVersion);
				Tagging.B(current.Tags, 0);
				Tagging.A(current.Tags, Content.XML_MODIFIED_BY, author);
				Tagging.A(current.Tags, Content.XML_MODIFIED_AT, c);
			}
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					goto end_IL_01ee;
				}
				continue;
				end_IL_01ee:
				break;
			}
		}
		else
		{
			Shape shape = ((ShapeItem)(object)A).Shape;
			Tagging.A(shape.Tags, A.LatestVersion);
			Tagging.B(shape.Tags, 0);
			Tagging.A(shape.Tags, Content.XML_MODIFIED_BY, author);
			Tagging.A(shape.Tags, Content.XML_MODIFIED_AT, c);
			_ = null;
		}
		A.RebaseVersion();
		application = null;
	}

	private void I()
	{
		Microsoft.Office.Interop.PowerPoint.Application application = NG.A.Application;
		using (List<string>.Enumerator enumerator = HiddenFiles.GetEnumerator())
		{
			while (enumerator.MoveNext())
			{
				string current = enumerator.Current;
				try
				{
					application.Presentations[Path.GetFileName(current)].Close();
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
		application = null;
	}

	private void IgnoreContent(object sender, RoutedEventArgs e)
	{
		//IL_000e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0018: Expected O, but got Unknown
		C((ContentItem)((Button)sender).DataContext);
	}

	private void C(ContentItem A)
	{
		if (this.A())
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
			if (!Prompts.AskIgnore(Window.GetWindow(this), A))
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
				NG.A.Application.StartNewUndoEntry();
				try
				{
					if (A is SlideItem)
					{
						foreach (Slide slide in ((SlideItem)(object)A).Slides)
						{
							Tagging.B(slide.Tags, A.LatestVersion);
						}
					}
					else
					{
						Tagging.B(((ShapeItem)(object)A).Shape.Tags, A.LatestVersion);
					}
					A.IgnoreVersion();
					if (chkOutdated.IsChecked == true)
					{
						A.IsSelected = false;
					}
					C();
					M();
					return;
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					C(ex2.Message);
					ProjectData.ClearProjectError();
					return;
				}
			}
		}
	}

	private void UnlinkContent(object sender, RoutedEventArgs e)
	{
		//IL_000d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0013: Expected O, but got Unknown
		ContentItem val = (ContentItem)((Button)sender).DataContext;
		if (!A())
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
			if (Prompts.AskUnlink(Window.GetWindow(this), val))
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
				if (val is SlideItem)
				{
					using List<Slide>.Enumerator enumerator = ((SlideItem)(object)val).Slides.GetEnumerator();
					while (enumerator.MoveNext())
					{
						Tagging.B(enumerator.Current.Tags);
					}
					while (true)
					{
						switch (5)
						{
						case 0:
							break;
						default:
							goto end_IL_008b;
						}
						continue;
						end_IL_008b:
						break;
					}
				}
				else
				{
					Tagging.B(((ShapeItem)(object)val).Shape.Tags);
				}
				D(val);
			}
		}
		val = null;
	}

	private void D(ContentItem A)
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
					switch (3)
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
				J();
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
		J();
		CollapseAnimationRunning = false;
	}

	private void J()
	{
		ContentItems.Remove(ItemQueuedToRemove);
		ItemQueuedToRemove = null;
	}

	private void PreviewSlideContent(object sender, RoutedEventArgs e)
	{
		new wpfSlideUpdates((SlideItem)((Button)sender).DataContext, A).Show();
		_ = null;
	}

	private void A(SlideItem A, bool B, bool C)
	{
		if (B)
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
					this.A((ContentItem)(object)A);
					return;
				}
			}
		}
		if (!C)
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
			this.C((ContentItem)(object)A);
			return;
		}
	}

	private void K()
	{
		try
		{
			IEnumerable<wpfSlideUpdates> source = System.Windows.Application.Current.Windows.OfType<wpfSlideUpdates>();
			if (source.Any())
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
				source.ElementAt(0).Close();
			}
			source = null;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	private bool A()
	{
		if (NG.A.Application.ActivePresentation.Final)
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
					D(PowerPointAddIn1.Links.Common.A);
					return true;
				}
			}
		}
		return false;
	}

	private void L()
	{
		Forms.FocusListBoxItem(lbxResults, true);
	}

	private void B(string A)
	{
		clsReporting.LogActivity((ActivityApp)2, (ActivityCategory)6, A);
	}

	private void M()
	{
		clsReporting.LogActivity((ActivityApp)2, (ActivityCategory)6, AH.A(64571));
	}

	private bool A(string A)
	{
		return UIFormsExtensions.AskOkCancel(Window.GetWindow(this), A);
	}

	private void C(string A)
	{
		Forms.ErrorMessage(Window.GetWindow(this), A);
	}

	private void D(string A)
	{
		Forms.WarningMessage(Window.GetWindow(this), A);
	}

	private void E(string A)
	{
		Forms.InfoMessage(Window.GetWindow(this), A);
	}

	private void F(string A)
	{
		Forms.SuccessMessage(Window.GetWindow(this), A);
	}

	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	[DebuggerNonUserCode]
	public void InitializeComponent()
	{
		if (this.m_B)
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
			this.m_B = true;
			Uri resourceLocator = new Uri(AH.A(64616), UriKind.Relative);
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
				switch (4)
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
		if (connectionId == 6)
		{
			grdFooter = (Grid)target;
			return;
		}
		if (connectionId == 7)
		{
			chkOutdated = (CheckBox)target;
			return;
		}
		if (connectionId == 8)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					btnReplaceAll = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 9)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					grdNoContent = (Grid)target;
					return;
				}
			}
		}
		this.m_B = true;
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
			((Button)target).Click += PreviewSlideContent;
		}
		if (connectionId == 3)
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
			((Button)target).Click += IgnoreContent;
		}
		if (connectionId == 4)
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
		if (connectionId != 5)
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
