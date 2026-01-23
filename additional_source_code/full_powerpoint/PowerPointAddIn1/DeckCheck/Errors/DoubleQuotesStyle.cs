using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using A;
using MacabacusMacros.Proofing;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.DeckCheck.UI;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class DoubleQuotesStyle : BaseTextError
{
	[CompilerGenerated]
	private new TextRange2 A;

	private TextRange2 ParentTextRange
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

	public DoubleQuotesStyle(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, List<TextRange2> listRanges, TextRange2 parent)
		: base(ErrorType.Text, ((Settings)Main.Analysis.Options).QuotesStyle, sld, shp, listRanges, blnHasFix: true, blnCanFixMultiple: true)
	{
		//IL_000e: Unknown result type (might be due to invalid IL or missing references)
		//IL_002d: Unknown result type (might be due to invalid IL or missing references)
		//IL_007e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0083: Unknown result type (might be due to invalid IL or missing references)
		//IL_0086: Invalid comparison between Unknown and I4
		BaseError val = (BaseError)(object)this;
		Errors.DoubleQuotesStyle(ref val, ((Settings)Main.Analysis.Options).QuotesStyleConvention);
		int count = listRanges.Count;
		string subtitle;
		if (count != 1)
		{
			subtitle = (((int)((Settings)Main.Analysis.Options).QuotesStyleConvention != 1) ? (AH.A(42825) + count + AH.A(43245)) : (AH.A(42825) + count + AH.A(43156)));
		}
		else
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
			subtitle = A((List<TextRange2>)((BaseError)this).TextRanges, shp);
		}
		((BaseError)this).Subtitle = subtitle;
		ParentTextRange = parent;
	}

	public override void FixAction(int i)
	{
		//IL_001f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0024: Unknown result type (might be due to invalid IL or missing references)
		//IL_0027: Invalid comparison between Unknown and I4
		NG.A.Application.StartNewUndoEntry();
		IEnumerator<TextRange2> enumerator = default(IEnumerator<TextRange2>);
		if ((int)((Settings)Main.Analysis.Options).QuotesStyleConvention == 1)
		{
			try
			{
				enumerator = ((BaseError)this).TextRanges.GetEnumerator();
				while (enumerator.MoveNext())
				{
					enumerator.Current.Text = AH.A(15132);
				}
				while (true)
				{
					switch (6)
					{
					case 0:
						break;
					default:
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						return;
					}
				}
			}
			finally
			{
				if (enumerator != null)
				{
					while (true)
					{
						switch (6)
						{
						case 0:
							break;
						default:
							enumerator.Dispose();
							goto end_IL_0079;
						}
						continue;
						end_IL_0079:
						break;
					}
				}
			}
		}
		int num = 0;
		checked
		{
			foreach (TextRange2 textRange in ((BaseError)this).TextRanges)
			{
				string text = string.Empty;
				string text2 = string.Empty;
				int start = textRange.Start;
				if (start == 1)
				{
					textRange.Text = Constants.DOUBLE_QUOTE_OPEN;
					continue;
				}
				if (start == ParentTextRange.get_Characters(-1, -1).Count)
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
					textRange.Text = Constants.DOUBLE_QUOTE_CLOSE;
					continue;
				}
				try
				{
					text2 = ParentTextRange.get_Characters(start - 1, 1).Text;
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					ProjectData.ClearProjectError();
				}
				try
				{
					text = ParentTextRange.get_Characters(start + 1, 1).Text;
				}
				catch (Exception ex3)
				{
					ProjectData.SetProjectError(ex3);
					Exception ex4 = ex3;
					ProjectData.ClearProjectError();
				}
				if (Operators.CompareString(text2, string.Empty, TextCompare: false) != 0)
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
					if (Regex.IsMatch(text2, AH.A(42976)))
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
						if (Operators.CompareString(text, string.Empty, TextCompare: false) != 0 && Regex.IsMatch(text, AH.A(42976)))
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
							num++;
							continue;
						}
					}
				}
				if (Operators.CompareString(text2, string.Empty, TextCompare: false) != 0)
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
					if (Regex.IsMatch(text2, AH.A(42976)))
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
						textRange.Text = Constants.DOUBLE_QUOTE_OPEN;
						continue;
					}
				}
				if (Operators.CompareString(text, string.Empty, TextCompare: false) != 0)
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
					if (Regex.IsMatch(text, AH.A(42976)))
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
						textRange.Text = Constants.DOUBLE_QUOTE_CLOSE;
						continue;
					}
				}
				num++;
			}
			if (num <= 0)
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
				Callout.DoNotClose = true;
				Forms.WarningMessage(AH.A(42981) + num + AH.A(43018));
				Callout.DoNotClose = false;
				return;
			}
		}
	}
}
