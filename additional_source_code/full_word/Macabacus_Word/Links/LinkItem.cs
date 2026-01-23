using System.Drawing;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros.ImportExport;
using MacabacusMacros.Links;
using MacabacusMacros.UI;

namespace Macabacus_Word.Links;

public sealed class LinkItem : LinkItem
{
	private Link m_A;

	private object m_A;

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
			//IL_0039: Unknown result type (might be due to invalid IL or missing references)
			//IL_0048: Unknown result type (might be due to invalid IL or missing references)
			//IL_0049: Unknown result type (might be due to invalid IL or missing references)
			//IL_005c: Unknown result type (might be due to invalid IL or missing references)
			//IL_005d: Unknown result type (might be due to invalid IL or missing references)
			//IL_0077: Unknown result type (might be due to invalid IL or missing references)
			//IL_0078: Unknown result type (might be due to invalid IL or missing references)
			this.m_A = value;
			((LinkItem)this).SourcePath = value.Source;
			((LinkItem)this).LastUpdate = Base.FormatTime(value.LastUpdate);
			((LinkItem)this).ModifiedBy = value.LastUser;
			((LinkItem)this).SourceRange = Manage2.SourceRangeName(Link);
			((LinkItem)this).LinkTypeToolTip = Manage2.LinkTypeToolTip(value.Type);
			((LinkItem)this).LinkTypeImage = Forms.GetImageSource(A(value.Type));
			((LinkItem)this).SourceTypeImage = Forms.GetImageSource(B(value.Type));
			((LinkItem)this).NotifyPropertyChanged(XC.A(17893));
		}
	}

	public object LinkedObject
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = RuntimeHelpers.GetObjectValue(value);
			((LinkItem)this).NotifyPropertyChanged(XC.A(17902));
		}
	}

	public LinkItem(Link link, object objLinked)
	{
		//IL_0007: Unknown result type (might be due to invalid IL or missing references)
		Link = link;
		LinkedObject = RuntimeHelpers.GetObjectValue(objLinked);
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
			return M.Picture;
		case 1:
			return M.Table;
		case 2:
		case 7:
			return M.EmbeddedExcel;
		case 4:
		case 6:
			return M.ChartSmall;
		default:
			return M.GroupFormatText;
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
			return M.Table;
		case 5:
		case 6:
		case 7:
		case 11:
			return M.ChartSmall;
		default:
			return M.ContentTypeShapes;
		}
	}
}
