using System;
using System.Collections;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic.CompilerServices;

namespace Macabacus_Word.TextOps.Redaction.Redactors;

public sealed class TextHighlighter
{
	public static void AddHighlight(Range rngWord, int textColor)
	{
		Font font = rngWord.Font;
		if (rngWord.ListFormat.ListType == WdListType.wdListNoNumbering)
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
			if (textColor == 0)
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
				B(rngWord);
				goto IL_007a;
			}
		}
		try
		{
			rngWord.HighlightColorIndex = WdColorIndex.wdAuto;
			font.Shading.Texture = WdTextureIndex.wdTextureNone;
			font.Shading.BackgroundPatternColor = (WdColor)textColor;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			A(rngWord);
			ProjectData.ClearProjectError();
		}
		goto IL_007a;
		IL_007a:
		font.TextColor.RGB = textColor;
		font = null;
	}

	private static void A(Range A)
	{
		if (A.ListFormat.ListType == WdListType.wdListNoNumbering)
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
					A.HighlightColorIndex = WdColorIndex.wdBlack;
					return;
				}
			}
		}
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = A.Characters.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Range obj = (Range)enumerator.Current;
				_ = obj.Text;
				obj.HighlightColorIndex = WdColorIndex.wdBlack;
			}
		}
		finally
		{
			if (enumerator is IDisposable)
			{
				while (true)
				{
					switch (6)
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

	private static void B(Range A)
	{
		A.HighlightColorIndex = WdColorIndex.wdBlack;
	}
}
