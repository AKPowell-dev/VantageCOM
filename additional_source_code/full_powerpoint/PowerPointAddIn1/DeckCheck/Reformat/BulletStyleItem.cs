using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Windows;
using A;
using MacabacusMacros.Proofing.UI.Reformat;
using Microsoft.Office.Core;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.DeckCheck.Reformat;

public sealed class BulletStyleItem : BaseItem
{
	[CompilerGenerated]
	private BulletStyle A;

	public BulletStyle Style
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

	public BulletStyleItem(BulletStyle bs, List<IndexedObject> listObjects, DataTemplate template, DataTemplate navItemTemplate, int intTotal, int intIndex)
		: base(intTotal, intIndex, listObjects, template, navItemTemplate, AH.A(49345))
	{
		Style = bs;
		switch (bs.Style)
		{
		case MsoNumberedBulletStyle.msoBulletArabicParenBoth:
			((BaseItem)this).Label = AH.A(49390);
			break;
		case MsoNumberedBulletStyle.msoBulletArabicParenRight:
			((BaseItem)this).Label = AH.A(49413);
			break;
		case MsoNumberedBulletStyle.msoBulletArabicPeriod:
			((BaseItem)this).Label = AH.A(49430);
			break;
		case MsoNumberedBulletStyle.msoBulletAlphaUCParenBoth:
			((BaseItem)this).Label = AH.A(49447);
			break;
		case MsoNumberedBulletStyle.msoBulletAlphaUCParenRight:
			((BaseItem)this).Label = AH.A(49470);
			break;
		case MsoNumberedBulletStyle.msoBulletAlphaUCPeriod:
			((BaseItem)this).Label = AH.A(49487);
			break;
		case MsoNumberedBulletStyle.msoBulletAlphaLCParenBoth:
			((BaseItem)this).Label = AH.A(49504);
			break;
		case MsoNumberedBulletStyle.msoBulletAlphaLCParenRight:
			((BaseItem)this).Label = AH.A(49527);
			break;
		case MsoNumberedBulletStyle.msoBulletAlphaLCPeriod:
			((BaseItem)this).Label = AH.A(49544);
			break;
		case MsoNumberedBulletStyle.msoBulletRomanUCParenBoth:
			((BaseItem)this).Label = AH.A(49561);
			break;
		case MsoNumberedBulletStyle.msoBulletRomanUCParenRight:
			((BaseItem)this).Label = AH.A(49590);
			break;
		case MsoNumberedBulletStyle.msoBulletRomanUCPeriod:
			((BaseItem)this).Label = AH.A(49613);
			break;
		case MsoNumberedBulletStyle.msoBulletRomanLCParenBoth:
			((BaseItem)this).Label = AH.A(49636);
			break;
		case MsoNumberedBulletStyle.msoBulletRomanLCParenRight:
			((BaseItem)this).Label = AH.A(49665);
			break;
		case MsoNumberedBulletStyle.msoBulletRomanLCPeriod:
			((BaseItem)this).Label = AH.A(49688);
			break;
		case MsoNumberedBulletStyle.msoBulletCircleNumDBPlain:
			((BaseItem)this).Label = AH.A(49711);
			break;
		case MsoNumberedBulletStyle.msoBulletCircleNumWDWhitePlain:
			((BaseItem)this).Label = AH.A(49750);
			break;
		case MsoNumberedBulletStyle.msoBulletCircleNumWDBlackPlain:
			((BaseItem)this).Label = AH.A(49801);
			break;
		case MsoNumberedBulletStyle.msoBulletArabicPlain:
		case MsoNumberedBulletStyle.msoBulletSimpChinPlain:
		case MsoNumberedBulletStyle.msoBulletSimpChinPeriod:
			break;
		}
	}

	public void Reformat(BulletStyleOption opt, ref List<string> listErrors)
	{
		NG.A.Application.StartNewUndoEntry();
		foreach (NavigationItem @object in base.Objects)
		{
			IndexedObject indexedObject = @object.IndexedObject;
			try
			{
				if (indexedObject.Child is BulletFormat2)
				{
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
						((ParagraphFormat2)((BulletFormat2)indexedObject.Child).Parent).Bullet.Style = opt.Style.Style;
						_ = null;
						break;
					}
				}
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
