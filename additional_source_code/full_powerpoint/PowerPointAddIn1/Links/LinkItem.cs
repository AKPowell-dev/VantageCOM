using System.Drawing;
using A;
using MacabacusMacros.ImportExport;
using MacabacusMacros.Links;
using MacabacusMacros.UI;

namespace PowerPointAddIn1.Links;

public class LinkItem : LinkItem
{
	private Link m_A;

	public override Link Link
	{
		get
		{
			//IL_0001: Unknown result type (might be due to invalid IL or missing references)
			return this.m_A;
		}
		set
		{
			//IL_0001: Unknown result type (might be due to invalid IL or missing references)
			//IL_0002: Unknown result type (might be due to invalid IL or missing references)
			//IL_0008: Unknown result type (might be due to invalid IL or missing references)
			//IL_0014: Unknown result type (might be due to invalid IL or missing references)
			//IL_0027: Unknown result type (might be due to invalid IL or missing references)
			//IL_0034: Unknown result type (might be due to invalid IL or missing references)
			//IL_0046: Unknown result type (might be due to invalid IL or missing references)
			//IL_0047: Unknown result type (might be due to invalid IL or missing references)
			//IL_005a: Unknown result type (might be due to invalid IL or missing references)
			//IL_005b: Unknown result type (might be due to invalid IL or missing references)
			//IL_0073: Unknown result type (might be due to invalid IL or missing references)
			//IL_0074: Unknown result type (might be due to invalid IL or missing references)
			this.m_A = value;
			((LinkItem)this).SourcePath = value.Source;
			((LinkItem)this).LastUpdate = Base.FormatTime(value.LastUpdate);
			((LinkItem)this).ModifiedBy = value.LastUser;
			((LinkItem)this).SourceRange = Manage2.SourceRangeName(Link);
			((LinkItem)this).LinkTypeToolTip = Manage2.LinkTypeToolTip(value.Type);
			((LinkItem)this).LinkTypeImage = Forms.GetImageSource(A(value.Type));
			((LinkItem)this).SourceTypeImage = Forms.GetImageSource(B(value.Type));
			((LinkItem)this).NotifyPropertyChanged(AH.A(59442));
		}
	}

	public LinkItem(Link lnk, string strGroup)
	{
		//IL_0007: Unknown result type (might be due to invalid IL or missing references)
		Link = lnk;
		((LinkItem)this).Group = strGroup;
	}

	private Bitmap A(ImportType A)
	{
		//IL_0000: Unknown result type (might be due to invalid IL or missing references)
		//IL_0001: Unknown result type (might be due to invalid IL or missing references)
		//IL_0002: Unknown result type (might be due to invalid IL or missing references)
		//IL_0004: Unknown result type (might be due to invalid IL or missing references)
		//IL_003a: Expected I4, but got Unknown
		switch (A - 1)
		{
		case 0:
		case 5:
		case 10:
		case 11:
			return OB.Picture;
		case 1:
			return OB.Table;
		case 2:
		case 7:
			return OB.EmbeddedExcel;
		case 4:
		case 6:
			return OB.ChartSmall;
		default:
			return OB.GroupFormatText;
		}
	}

	private Bitmap B(ImportType A)
	{
		//IL_0000: Unknown result type (might be due to invalid IL or missing references)
		//IL_0001: Unknown result type (might be due to invalid IL or missing references)
		//IL_0002: Unknown result type (might be due to invalid IL or missing references)
		//IL_0004: Unknown result type (might be due to invalid IL or missing references)
		//IL_003a: Expected I4, but got Unknown
		switch (A - 1)
		{
		case 0:
		case 1:
		case 2:
		case 3:
		case 4:
		case 10:
			return OB.Table;
		case 5:
		case 6:
		case 7:
		case 11:
			return OB.ChartSmall;
		default:
			return OB.ContentTypeShapes;
		}
	}
}
