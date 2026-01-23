using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Media;
using A;
using MacabacusMacros;
using MacabacusMacros.Libraries.Versioning;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.Library2.Versioning;

public sealed class SlideItem : ContentItem
{
	[CompilerGenerated]
	private List<Slide> m_A;

	private bool m_A;

	public List<Slide> Slides
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

	public override bool IsSelected
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			((ContentItem)this).NotifyPropertyChanged(AH.A(62846));
			int ignoreButtonVisibility;
			if (value && ((ContentItem)this).IsOutdated)
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
				if (!((ContentItem)this).IsLegacySlideLink)
				{
					ignoreButtonVisibility = 0;
					goto IL_0050;
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
			}
			ignoreButtonVisibility = 2;
			goto IL_0050;
			IL_007b:
			int previewButtonVisibility;
			((ContentItem)this).PreviewButtonVisibility = (Visibility)previewButtonVisibility;
			return;
			IL_0050:
			((ContentItem)this).IgnoreButtonVisibility = (Visibility)ignoreButtonVisibility;
			if (value)
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
				if (((ContentItem)this).IsOutdated)
				{
					previewButtonVisibility = 0;
					goto IL_007b;
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
			previewButtonVisibility = 2;
			goto IL_007b;
		}
	}

	public SlideItem(Slide sld, ContentInfo ci, ManifestInfo mi)
		: base(ci, mi)
	{
		//IL_0001: Unknown result type (might be due to invalid IL or missing references)
		//IL_0002: Unknown result type (might be due to invalid IL or missing references)
		this.m_A = false;
		Slides = new List<Slide>();
		Slides.Add(sld);
		((ContentItem)this).IconData = Geometry.Parse(AH.A(62867));
		((ContentItem)this).IconPadding = new Thickness(3.0);
	}

	internal void A(Slide A)
	{
		Slides.Add(A);
		((ContentItem)this).ItemsCount = checked(((ContentItem)this).ItemsCount + 1);
	}

	public void ClearReferences()
	{
		List<Slide> slides = Slides;
		ReleaseHelper.ClearListReferences<Slide>(ref slides, false, (Action<Slide>)null);
		Slides = slides;
	}
}
