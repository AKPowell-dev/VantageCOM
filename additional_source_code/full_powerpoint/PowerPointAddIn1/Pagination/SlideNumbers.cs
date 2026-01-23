using System;
using System.Collections;
using A;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.Slides;

namespace PowerPointAddIn1.Pagination;

public sealed class SlideNumbers
{
	public static void Reset()
	{
		try
		{
			Application application = NG.A.Application;
			application.StartNewUndoEntry();
			A(application.ActivePresentation);
			_ = null;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	private static void A(Microsoft.Office.Interop.PowerPoint.Presentation A)
	{
		IEnumerator enumerator = A.Slides.GetEnumerator();
		try
		{
			while (enumerator.MoveNext())
			{
				Reset((Slide)enumerator.Current);
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

	public static void Reset(Slide sld)
	{
		HeaderFooter slideNumber = sld.HeadersFooters.SlideNumber;
		if (slideNumber.Visible == MsoTriState.msoTrue)
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
			slideNumber.Visible = MsoTriState.msoFalse;
			slideNumber.Visible = MsoTriState.msoTrue;
		}
		slideNumber = null;
	}

	public static void Freeze(Slide sld)
	{
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = sld.Shapes.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Microsoft.Office.Interop.PowerPoint.Shape shape = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current;
				try
				{
					if (!Numbers.IsSlideNumberPlaceholder(shape))
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
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						TextRange2 textRange = shape.TextFrame2.TextRange;
						textRange.Text = textRange.Text;
						_ = null;
						return;
					}
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
				switch (2)
				{
				case 0:
					break;
				default:
					return;
				}
			}
		}
		finally
		{
			if (enumerator is IDisposable)
			{
				while (true)
				{
					switch (5)
					{
					case 0:
						continue;
					}
					(enumerator as IDisposable).Dispose();
					break;
				}
			}
		}
	}
}
