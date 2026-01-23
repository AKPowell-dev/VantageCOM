using System.Runtime.CompilerServices;
using A;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.DeckCheck.Errors;
using PowerPointAddIn1.Slides;

namespace PowerPointAddIn1.DeckCheck.Check;

public sealed class SlideNumbers
{
	[CompilerGenerated]
	private bool A;

	private bool RequireSlideNumbers
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

	public SlideNumbers(Conventions conv)
	{
		RequireSlideNumbers = (double)conv.SlideNumbersCount >= (double)conv.ContentSlidesCount * 0.5;
	}

	public void Check(Slide sld)
	{
		if (sld.SlideIndex <= 1 || Helpers.IsSpecialSlide(sld))
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
			if (RequireSlideNumbers)
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
				if (sld.HeadersFooters.SlideNumber.Visible == MsoTriState.msoFalse)
				{
					while (true)
					{
						switch (3)
						{
						case 0:
							break;
						default:
							Main.Analysis.Errors.Add(new MissingSlideNumber(sld, AH.A(14753) + Conversions.ToString(sld.SlideIndex) + AH.A(14780)));
							return;
						}
					}
				}
			}
			if (RequireSlideNumbers)
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
				if (sld.HeadersFooters.SlideNumber.Visible == MsoTriState.msoTrue)
				{
					Main.Analysis.Errors.Add(new ExcessSlideNumber(sld, AH.A(14753) + Conversions.ToString(sld.SlideIndex) + AH.A(14780)));
				}
				return;
			}
		}
	}
}
