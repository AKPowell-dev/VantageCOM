using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using MacabacusMacros;
using MacabacusMacros.Links;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.Links;

public sealed class SlideLink
{
	[CompilerGenerated]
	private string A;

	[CompilerGenerated]
	private string B;

	[CompilerGenerated]
	private bool A;

	[CompilerGenerated]
	private List<Slide> A;

	public string LinkId
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

	public string SourcePath
	{
		[CompilerGenerated]
		get
		{
			return B;
		}
		[CompilerGenerated]
		set
		{
			B = value;
		}
	}

	public bool KeepSourceFormatting
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

	public List<Slide> OldSlides
	{
		[CompilerGenerated]
		get
		{
			return A;
		}
		[CompilerGenerated]
		set
		{
			A = value;
		}
	}

	public SlideLink(Link link, Slide sld)
	{
		//IL_0007: Unknown result type (might be due to invalid IL or missing references)
		//IL_0013: Unknown result type (might be due to invalid IL or missing references)
		//IL_001f: Unknown result type (might be due to invalid IL or missing references)
		LinkId = link.Name;
		SourcePath = link.Source;
		KeepSourceFormatting = link.KeepSourceFormatting;
		OldSlides = new List<Slide>();
		OldSlides.Add(sld);
	}

	public void ClearReferences()
	{
		List<Slide> oldSlides = OldSlides;
		ReleaseHelper.ClearListReferences<Slide>(ref oldSlides, false, (Action<Slide>)null);
		OldSlides = oldSlides;
	}
}
