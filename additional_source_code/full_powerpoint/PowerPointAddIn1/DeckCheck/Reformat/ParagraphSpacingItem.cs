using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Windows;
using A;
using MacabacusMacros;
using MacabacusMacros.Proofing.UI.Reformat;
using Microsoft.Office.Core;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.DeckCheck.Reformat;

public sealed class ParagraphSpacingItem : BaseItem
{
	private readonly string A;

	private readonly string B;

	[CompilerGenerated]
	private LineSpacing A;

	public LineSpacing Spacing
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

	public ParagraphSpacingItem(LineSpacing spacing, List<IndexedObject> listObjects, DataTemplate template, DataTemplate navItemTemplate, int intTotal, int intIndex)
		: base(intTotal, intIndex, listObjects, template, navItemTemplate, AH.A(50823))
	{
		this.A = AH.A(50090);
		B = AH.A(50099);
		string text;
		if (Operators.CompareString(clsPublish.SystemDecimalSeparator(), AH.A(14417), TextCompare: false) == 0)
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
			text = this.A;
		}
		else
		{
			text = B;
		}
		Spacing = spacing;
		string text2 = spacing.Before.ToString(text) + AH.A(14600) + spacing.After.ToString(text) + AH.A(14600);
		if (spacing.LineRuleWithin == MsoTriState.msoTrue)
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
			float within = spacing.Within;
			if (within == 1f)
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
				text2 += AH.A(50858);
			}
			else if (within == 2f)
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
				text2 += AH.A(50871);
			}
			else
			{
				text2 = text2 + spacing.Within.ToString(text) + AH.A(50884);
			}
		}
		else
		{
			text2 = text2 + spacing.Within.ToString(text) + AH.A(17909);
		}
		((BaseItem)this).Label = text2;
	}

	public void Reformat(ParagraphSpacingOption opt, ref List<string> listErrors)
	{
		NG.A.Application.StartNewUndoEntry();
		foreach (NavigationItem @object in base.Objects)
		{
			IndexedObject indexedObject = @object.IndexedObject;
			try
			{
				ParagraphFormat2 paragraphFormat = ((TextRange2)indexedObject.Child).ParagraphFormat;
				paragraphFormat.SpaceBefore = opt.Spacing.Before;
				paragraphFormat.SpaceAfter = opt.Spacing.After;
				paragraphFormat.LineRuleWithin = opt.Spacing.LineRuleWithin;
				paragraphFormat.SpaceWithin = opt.Spacing.Within;
				_ = null;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				listErrors.Add(ex2.Message);
				ProjectData.ClearProjectError();
			}
			indexedObject = null;
		}
	}
}
