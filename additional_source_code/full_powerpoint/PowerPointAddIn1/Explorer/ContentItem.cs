using System;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using A;
using MacabacusMacros.Explorer;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.Library2.Versioning;

namespace PowerPointAddIn1.Explorer;

public abstract class ContentItem : BaseItem
{
	[CompilerGenerated]
	private SlideItem m_A;

	[CompilerGenerated]
	private Microsoft.Office.Interop.PowerPoint.Shape m_A;

	[CompilerGenerated]
	private BitmapSource m_A;

	private string m_A;

	private Brush m_A;

	private Brush B;

	[CompilerGenerated]
	private ContextMenu m_A;

	private bool m_A;

	private Visibility m_A;

	private bool B;

	private Visibility B;

	public SlideItem Parent
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

	public Microsoft.Office.Interop.PowerPoint.Shape Shape
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

	public BitmapSource PreviewImage
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

	public string Tooltip
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			((BaseItem)this).NotifyPropertyChanged(AH.A(113591));
		}
	}

	public Brush FontColor
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			((BaseItem)this).NotifyPropertyChanged(AH.A(113606));
		}
	}

	public Brush IconColor
	{
		get
		{
			return this.B;
		}
		set
		{
			this.B = value;
			((BaseItem)this).NotifyPropertyChanged(AH.A(113625));
		}
	}

	public ContextMenu ContextMenu
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

	public bool IsLinked
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			int linkAdornerVisibility;
			if (!value)
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
				linkAdornerVisibility = 2;
			}
			else
			{
				linkAdornerVisibility = 0;
			}
			LinkAdornerVisibility = (Visibility)linkAdornerVisibility;
		}
	}

	public Visibility LinkAdornerVisibility
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			((BaseItem)this).NotifyPropertyChanged(AH.A(113644));
		}
	}

	public bool IsLibraryContent
	{
		get
		{
			return this.B;
		}
		set
		{
			this.B = value;
			int libraryAdornerVisibility;
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				libraryAdornerVisibility = 2;
			}
			else
			{
				libraryAdornerVisibility = 0;
			}
			LibraryAdornerVisibility = (Visibility)libraryAdornerVisibility;
		}
	}

	public Visibility LibraryAdornerVisibility
	{
		get
		{
			return B;
		}
		set
		{
			B = value;
			((BaseItem)this).NotifyPropertyChanged(AH.A(113687));
		}
	}

	public ContentItem(SlideItem wsi, string strLabel, SolidColorBrush brush, Geometry geo)
		: base(strLabel, geo)
	{
		PreviewImage = null;
		this.m_A = Visibility.Collapsed;
		B = Visibility.Collapsed;
		Parent = wsi;
		FontColor = new SolidColorBrush(base.DEFAULT_FONT_COLOR);
		IconColor = brush;
	}

	public abstract void Refresh();

	public abstract void Delete();

	public abstract void Search(string strQuery);

	public void SearchOnInstantiate()
	{
		if (Options.SearchQuery.Length <= 0)
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
				Search(Options.SearchQuery);
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

	public void Hide()
	{
		if (Shape == null)
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
			Shape.Visible = MsoTriState.msoFalse;
			UpdateColors(MsoTriState.msoFalse);
			return;
		}
	}

	public void Unhide()
	{
		if (Shape != null)
		{
			Shape.Visible = MsoTriState.msoTrue;
			UpdateColors(MsoTriState.msoTrue);
		}
	}

	public void UpdateColors(MsoTriState visible)
	{
		double opacity;
		if (visible == MsoTriState.msoTrue)
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
			opacity = 1.0;
		}
		else
		{
			opacity = base.HIDDEN_OPACITY;
		}
		FontColor.Opacity = opacity;
		IconColor.Opacity = opacity;
	}

	internal bool A(Microsoft.Office.Interop.PowerPoint.Shape A)
	{
		return Tagging.A(A);
	}
}
