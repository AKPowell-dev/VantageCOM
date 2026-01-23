using System;
using System.Collections;
using System.Collections.ObjectModel;
using System.Runtime.CompilerServices;
using System.Windows.Media;
using A;
using MacabacusMacros.Explorer;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.Presentation;

namespace PowerPointAddIn1.Explorer;

public sealed class PresentationItem : BaseItem
{
	private readonly Color A;

	private readonly Color B;

	private readonly Color C;

	private bool A;

	private double A;

	[CompilerGenerated]
	private Microsoft.Office.Interop.PowerPoint.Presentation A;

	private ObservableCollection<SlideItem> A;

	public override bool IsSelected
	{
		get
		{
			return this.A;
		}
		set
		{
			this.A = value;
			((BaseItem)this).NotifyPropertyChanged(AH.A(62846));
			RefreshLabel();
		}
	}

	public double Opacity
	{
		get
		{
			return this.A;
		}
		set
		{
			this.A = value;
			((BaseItem)this).NotifyPropertyChanged(AH.A(114575));
		}
	}

	public Microsoft.Office.Interop.PowerPoint.Presentation Presentation
	{
		[CompilerGenerated]
		get
		{
			return this.A;
		}
		[CompilerGenerated]
		set
		{
			this.A = value;
		}
	}

	public ObservableCollection<SlideItem> Slides
	{
		get
		{
			return A;
		}
		set
		{
			A = value;
			((BaseItem)this).NotifyPropertyChanged(AH.A(114590));
		}
	}

	public PresentationItem(Microsoft.Office.Interop.PowerPoint.Presentation pres)
		: base(pres.Name, AH.A(114603))
	{
		//IL_006a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0074: Expected O, but got Unknown
		this.A = Color.FromRgb(175, 175, 175);
		B = System.Windows.Media.Colors.Firebrick;
		C = System.Windows.Media.Colors.LightCoral;
		if (Constants.ColorPalette == null)
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
			Constants.ColorPalette = new Palette();
		}
		Opacity = 1.0;
		Slides = new ObservableCollection<SlideItem>();
		Presentation = pres;
		((BaseItem)this).IsExpanded = true;
		Slide slide;
		try
		{
			slide = pres.Application.ActiveWindow.Selection.SlideRange[1];
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			slide = null;
			ProjectData.ClearProjectError();
		}
		foreach (Slide slide2 in pres.Slides)
		{
			SlideItem slideItem = new SlideItem(this, slide2);
			if (slide2 == slide)
			{
				slideItem.IsSelected = true;
			}
			Slides.Add(slideItem);
			slideItem = null;
		}
		slide = null;
		if (pres.Path.Length == 0)
		{
			pres.Saved = MsoTriState.msoFalse;
		}
	}

	public PresentationItem Clone()
	{
		return (PresentationItem)((object)this).MemberwiseClone();
	}

	public void Refresh()
	{
		RefreshSlides();
		RefreshLabel();
	}

	public void RefreshSlides()
	{
		Slides.Clear();
		IEnumerator enumerator = Presentation.Slides.GetEnumerator();
		try
		{
			while (enumerator.MoveNext())
			{
				Slide sld = (Slide)enumerator.Current;
				Slides.Add(new SlideItem(this, sld));
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
				return;
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

	public void RefreshLabel()
	{
		((BaseItem)this).Label = Presentation.Name;
	}

	public void Close()
	{
		Presentation.Close();
	}

	public void CloseOthers()
	{
		Miscellaneous.CloseOthers(Presentation);
	}

	public void Reopen()
	{
		Presentation.Windows[1].Activate();
		Miscellaneous.Reopen();
	}

	public void Duplicate()
	{
		Pane.CreatePane(Miscellaneous.Duplicate(Presentation));
	}

	public void Rename()
	{
	}

	public void ShowInFolder()
	{
		Miscellaneous.OpenFolder(Presentation);
	}

	public void CopyPath()
	{
		Miscellaneous.CopyPath(Presentation);
	}
}
