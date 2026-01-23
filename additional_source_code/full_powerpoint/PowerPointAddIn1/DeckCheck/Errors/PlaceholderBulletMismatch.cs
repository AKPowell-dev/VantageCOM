using System.Collections.Generic;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class PlaceholderBulletMismatch : BaseError
{
	[CompilerGenerated]
	private new BulletFormat2 A;

	private BulletFormat2 MasterBullet
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

	public PlaceholderBulletMismatch(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, List<TextRange2> listRanges, int intLevel, BulletFormat2 bullet)
		: base(ErrorType.PlaceholderBulletMismatch, Main.Analysis.Options.CheckPlaceholderBulletMismatch, sld, shp, blnHasFix: true)
	{
		//IL_0011: Unknown result type (might be due to invalid IL or missing references)
		//IL_0016: Unknown result type (might be due to invalid IL or missing references)
		((BaseError)this).Title = AH.A(37946);
		((BaseError)this).Subtitle = AH.A(38015) + intLevel + AH.A(38028);
		((BaseError)this).Tooltip = AH.A(38015) + intLevel + AH.A(38162);
		((BaseError)this).TextRanges = listRanges;
		MasterBullet = bullet;
	}

	public override void FixAction()
	{
		NG.A.Application.StartNewUndoEntry();
		using IEnumerator<TextRange2> enumerator = ((BaseError)this).TextRanges.GetEnumerator();
		BulletFormat2 bullet;
		for (; enumerator.MoveNext(); bullet = null)
		{
			bullet = enumerator.Current.ParagraphFormat.Bullet;
			if (bullet.Type != MsoBulletType.msoBulletNumbered)
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
				if (bullet.Type != MsoBulletType.msoBulletUnnumbered)
				{
					continue;
				}
				while (true)
				{
					switch (2)
					{
					case 0:
						continue;
					}
					break;
				}
			}
			if (bullet.Type == MsoBulletType.msoBulletUnnumbered && MasterBullet.Type == MsoBulletType.msoBulletUnnumbered)
			{
				if (bullet.UseTextFont == MsoTriState.msoFalse)
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
					bullet.Font.Name = MasterBullet.Font.Name;
				}
				bullet.Character = MasterBullet.Character;
			}
			else if (bullet.Type == MsoBulletType.msoBulletNumbered && MasterBullet.Type == MsoBulletType.msoBulletNumbered)
			{
				bullet.Style = MasterBullet.Style;
			}
			if (bullet.UseTextColor == MsoTriState.msoFalse)
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
				bullet.Font.Fill.ForeColor.RGB = MasterBullet.Font.Fill.ForeColor.RGB;
			}
			bullet.RelativeSize = MasterBullet.RelativeSize;
		}
		while (true)
		{
			switch (5)
			{
			case 0:
				break;
			default:
				return;
			}
		}
	}
}
