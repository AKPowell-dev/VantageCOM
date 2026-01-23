using System;
using System.CodeDom.Compiler;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Markup;
using A;
using MacabacusMacros;
using MacabacusMacros.Libraries.Versioning;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.Library2.Versioning;

[DesignerGenerated]
public sealed class wpfSlideUpdates : Window, IComponentConnector, IStyleConnector
{
	private ObservableCollection<LD> m_A;

	private bool m_A;

	private bool B;

	private Action<SlideItem, bool, bool> m_A;

	[CompilerGenerated]
	private SlideItem m_A;

	[AccessedThroughProperty("lvSlides")]
	[CompilerGenerated]
	private ListView m_A;

	[AccessedThroughProperty("btnViewSource")]
	[CompilerGenerated]
	private Button m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("btnIgnore")]
	private Button B;

	[AccessedThroughProperty("btnUpdate")]
	[CompilerGenerated]
	private Button C;

	[CompilerGenerated]
	[AccessedThroughProperty("btnClose")]
	private Button D;

	private bool C;

	private SlideItem SlideItem
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

	internal virtual ListView lvSlides
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

	internal virtual Button btnViewSource
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
			RoutedEventHandler value2 = ViewSource;
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

	internal virtual Button btnIgnore
	{
		[CompilerGenerated]
		get
		{
			return B;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			RoutedEventHandler value2 = btnIgnore_Click;
			Button button = B;
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
			B = value;
			button = B;
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

	internal virtual Button btnUpdate
	{
		[CompilerGenerated]
		get
		{
			return this.C;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			RoutedEventHandler value2 = btnUpdate_Click;
			Button button = this.C;
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
			this.C = value;
			button = this.C;
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
				button.Click += value2;
				return;
			}
		}
	}

	internal virtual Button btnClose
	{
		[CompilerGenerated]
		get
		{
			return D;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			RoutedEventHandler value2 = btnClose_Click;
			Button button = D;
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
			D = value;
			button = D;
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

	public wpfSlideUpdates(SlideItem itm, Action<SlideItem, bool, bool> act)
	{
		base.Closing += wpfSlideUpdates_Closing;
		base.Closed += wpfSlideUpdates_Closed;
		this.m_A = false;
		this.B = false;
		InitializeComponent();
		base.Icon = Forms.GetIcon();
		this.m_A = new ObservableCollection<LD>();
		this.m_A.Add(new LD(itm));
		lvSlides.ItemsSource = this.m_A;
		SlideItem = itm;
		this.m_A = act;
		btnViewSource.Visibility = Visibility.Collapsed;
		if (((ContentItem)itm).IsLegacySlideLink)
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
					btnIgnore.Visibility = Visibility.Collapsed;
					return;
				}
			}
		}
		btnIgnore.Visibility = Visibility.Visible;
	}

	private void wpfSlideUpdates_Closing(object sender, CancelEventArgs e)
	{
	}

	private void wpfSlideUpdates_Closed(object sender, EventArgs e)
	{
		this.m_A(SlideItem, this.m_A, this.B);
		IEnumerator<LD> enumerator = default(IEnumerator<LD>);
		try
		{
			enumerator = this.m_A.GetEnumerator();
			while (enumerator.MoveNext())
			{
				enumerator.Current.A(A: false);
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					enumerator.Dispose();
					break;
				}
			}
		}
		this.m_A.Clear();
		this.m_A = null;
		ReleaseHelper.DoGarbageCollection();
	}

	private void btnUpdate_Click(object sender, RoutedEventArgs e)
	{
		this.m_A = true;
		Close();
	}

	private void btnIgnore_Click(object sender, RoutedEventArgs e)
	{
		this.B = true;
		Close();
	}

	private void btnClose_Click(object sender, RoutedEventArgs e)
	{
		Close();
	}

	private void GoToSlide(object sender, RoutedEventArgs e)
	{
		try
		{
			((Slide)((Button)sender).DataContext).Select();
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			Forms.ErrorMessage(Window.GetWindow(this), AH.A(63385));
			ProjectData.ClearProjectError();
		}
	}

	private void ShowArrows(object sender, MouseEventArgs e)
	{
		LD lD = (LD)((Grid)sender).DataContext;
		if (lD.SlideCount > 1)
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
			lD.D();
		}
		lD = null;
	}

	private void HideArrows(object sender, MouseEventArgs e)
	{
		LD obj = (LD)((Grid)sender).DataContext;
		obj.LeftArrowVisibility = Visibility.Hidden;
		obj.RightArrowVisibility = Visibility.Hidden;
		_ = null;
	}

	private void PreviewNext(object sender, RoutedEventArgs e)
	{
		A((Button)sender).A();
	}

	private void PreviewPrevious(object sender, RoutedEventArgs e)
	{
		A((Button)sender).B();
	}

	private LD A(Button A)
	{
		return (LD)A.DataContext;
	}

	private void ViewSource(object sender, RoutedEventArgs e)
	{
		//IL_01fc: Unknown result type (might be due to invalid IL or missing references)
		//IL_007b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0080: Unknown result type (might be due to invalid IL or missing references)
		//IL_0056: Unknown result type (might be due to invalid IL or missing references)
		//IL_005b: Unknown result type (might be due to invalid IL or missing references)
		//IL_01d6: Unknown result type (might be due to invalid IL or missing references)
		//IL_0226: Unknown result type (might be due to invalid IL or missing references)
		Microsoft.Office.Interop.PowerPoint.Application application = NG.A.Application;
		Microsoft.Office.Interop.PowerPoint.Presentation presentation = null;
		Slide a = this.m_A[0].OldSlides[0];
		string text = default(string);
		if (Tagging.A(a))
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
			text = Tagging.A(a).Value.ContentPath;
		}
		else if (Check.A(a))
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
			text = Check.A(a, Check.A(a));
		}
		try
		{
			presentation = application.Presentations[Path.GetFileName(text)];
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		if (presentation != null)
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
			if (Operators.CompareString(presentation.FullName.ToLower(), text.ToLower(), TextCompare: false) == 0)
			{
				goto IL_0202;
			}
		}
		try
		{
			new ComAwareEventInfo(typeof(EApplication_Event), AH.A(63418)).RemoveEventHandler(application, new EApplication_AfterPresentationOpenEventHandler(Access.AfterPresentationOpen));
			presentation = application.Presentations.Open(text, MsoTriState.msoTrue);
			new ComAwareEventInfo(typeof(EApplication_Event), AH.A(63418)).AddEventHandler(application, new EApplication_AfterPresentationOpenEventHandler(Access.AfterPresentationOpen));
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			new ComAwareEventInfo(typeof(EApplication_Event), AH.A(63418)).AddEventHandler(application, new EApplication_AfterPresentationOpenEventHandler(Access.AfterPresentationOpen));
			if (!clsFile.IsPathUrl(text))
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
				if (File.Exists(text))
				{
					clsReporting.LogException(ex4);
					application = null;
					throw new UpdateLinkException(AH.A(63461) + text + AH.A(63591) + ex4.Message);
				}
			}
			application = null;
			throw new UpdateLinkException(AH.A(63602) + text + AH.A(63695));
		}
		goto IL_0202;
		IL_0202:
		application = null;
		if (presentation != null)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					A(presentation);
					presentation = null;
					return;
				}
			}
		}
		throw new UpdateLinkException(AH.A(63859));
	}

	private static void A(Microsoft.Office.Interop.PowerPoint.Presentation A)
	{
		Microsoft.Office.Interop.PowerPoint.Presentation presentation = A;
		if (presentation.Windows.Count == 0)
		{
			presentation.NewWindow();
		}
		DocumentWindow documentWindow = presentation.Windows[1];
		documentWindow.Activate();
		if (documentWindow.View.Type != PpViewType.ppViewNormal)
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
			documentWindow.ViewType = PpViewType.ppViewNormal;
		}
		documentWindow = null;
		presentation = null;
	}

	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	[DebuggerNonUserCode]
	public void InitializeComponent()
	{
		if (C)
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
			C = true;
			Uri resourceLocator = new Uri(AH.A(63912), UriKind.Relative);
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
				switch (2)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					lvSlides = (ListView)target;
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
					btnViewSource = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 7)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					btnIgnore = (Button)target;
					return;
				}
			}
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
					btnUpdate = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 9)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					btnClose = (Button)target;
					return;
				}
			}
		}
		C = true;
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
			((Button)target).Click += GoToSlide;
		}
		if (connectionId == 3)
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
			((Grid)target).MouseEnter += ShowArrows;
			((Grid)target).MouseLeave += HideArrows;
		}
		if (connectionId == 4)
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
			((Button)target).Click += PreviewPrevious;
		}
		if (connectionId == 5)
		{
			((Button)target).Click += PreviewNext;
		}
	}

	void IStyleConnector.Connect(int connectionId, object target)
	{
		//ILSpy generated this explicit interface implementation from .override directive in System_Windows_Markup_IStyleConnector_Connect
		this.System_Windows_Markup_IStyleConnector_Connect(connectionId, target);
	}
}
