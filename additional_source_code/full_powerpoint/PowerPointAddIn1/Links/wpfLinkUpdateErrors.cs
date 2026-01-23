using System;
using System.CodeDom.Compiler;
using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Markup;
using A;
using MacabacusMacros;
using MacabacusMacros.Links;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.Links;

[DesignerGenerated]
public sealed class wpfLinkUpdateErrors : Window, IComponentConnector
{
	[AccessedThroughProperty("lvErrors")]
	[CompilerGenerated]
	private ListView A;

	[CompilerGenerated]
	[AccessedThroughProperty("colShapeOrSlide")]
	private GridViewColumn A;

	[CompilerGenerated]
	[AccessedThroughProperty("btnClose")]
	private Button A;

	private bool A;

	internal virtual ListView lvErrors
	{
		[CompilerGenerated]
		get
		{
			return this.A;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			SelectionChangedEventHandler value2 = lvErrors_SelectionChanged;
			ListView listView = this.A;
			if (listView != null)
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
				listView.SelectionChanged -= value2;
			}
			this.A = value;
			listView = this.A;
			if (listView == null)
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
				listView.SelectionChanged += value2;
				return;
			}
		}
	}

	internal virtual GridViewColumn colShapeOrSlide
	{
		[CompilerGenerated]
		get
		{
			return this.A;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.A = value;
		}
	}

	internal virtual Button btnClose
	{
		[CompilerGenerated]
		get
		{
			return this.A;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			RoutedEventHandler value2 = btnClose_Click;
			Button button = this.A;
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
			this.A = value;
			button = this.A;
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

	public wpfLinkUpdateErrors()
	{
		InitializeComponent();
		base.Icon = Forms.GetIcon();
	}

	private void lvErrors_SelectionChanged(object sender, SelectionChangedEventArgs e)
	{
		//IL_0107: Unknown result type (might be due to invalid IL or missing references)
		Slide slide = null;
		Microsoft.Office.Interop.PowerPoint.Shape shape = null;
		Hyperlink hyperlink = null;
		TextRange2 textRange = null;
		if (lvErrors.SelectedItems.Count <= 0)
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
			LinkError linkError = (LinkError)lvErrors.SelectedItems[0];
			if (linkError.LinkedObject is Slide)
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
				slide = (Slide)linkError.LinkedObject;
			}
			else if (linkError.LinkedObject is Microsoft.Office.Interop.PowerPoint.Shape)
			{
				shape = (Microsoft.Office.Interop.PowerPoint.Shape)linkError.LinkedObject;
				slide = clsPowerPoint.GetSlideFromShape(shape);
			}
			else if (linkError.LinkedObject is Hyperlink)
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
				hyperlink = (Hyperlink)linkError.LinkedObject;
				shape = Hyperlinks.GetParentShape(hyperlink, blnIgnoreTables: false);
				slide = clsPowerPoint.GetSlideFromShape(shape);
			}
			else if (linkError.LinkedObject is TextLink)
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
				textRange = ((TextLink)linkError.LinkedObject).TextRange;
				shape = Text.TextRangeParentShape(textRange);
				slide = clsPowerPoint.GetSlideFromShape(shape);
			}
			if (slide != null)
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
				NG.A.Application.ActiveWindow.View.GotoSlide(slide.SlideIndex);
			}
			if (shape != null)
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
				if (shape.Visible == MsoTriState.msoFalse)
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
					Forms.WarningMessage(Window.GetWindow(this), AH.A(93507));
					goto IL_01c9;
				}
			}
			if (hyperlink != null)
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
				Hyperlinks.HyperlinkParentTextRange(hyperlink).Select();
			}
			else if (textRange != null)
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
				textRange.Select();
			}
			else
			{
				shape?.Select();
			}
			goto IL_01c9;
			IL_01c9:
			shape = null;
			slide = null;
			hyperlink = null;
			textRange = null;
			linkError = null;
			return;
		}
	}

	private void btnClose_Click(object sender, RoutedEventArgs e)
	{
		Close();
	}

	[DebuggerNonUserCode]
	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	public void InitializeComponent()
	{
		if (A)
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
			A = true;
			Uri resourceLocator = new Uri(AH.A(93586), UriKind.Relative);
			System.Windows.Application.LoadComponent(this, resourceLocator);
			return;
		}
	}

	void IComponentConnector.InitializeComponent()
	{
		//ILSpy generated this explicit interface implementation from .override directive in InitializeComponent
		this.InitializeComponent();
	}

	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	[DebuggerNonUserCode]
	[EditorBrowsable(EditorBrowsableState.Never)]
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
					lvErrors = (ListView)target;
					return;
				}
			}
		}
		if (connectionId == 2)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					colShapeOrSlide = (GridViewColumn)target;
					return;
				}
			}
		}
		if (connectionId == 3)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					btnClose = (Button)target;
					return;
				}
			}
		}
		A = true;
	}

	void IComponentConnector.Connect(int connectionId, object target)
	{
		//ILSpy generated this explicit interface implementation from .override directive in System_Windows_Markup_IComponentConnector_Connect
		this.System_Windows_Markup_IComponentConnector_Connect(connectionId, target);
	}
}
