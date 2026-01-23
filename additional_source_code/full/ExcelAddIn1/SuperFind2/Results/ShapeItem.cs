using System.Runtime.CompilerServices;
using System.Windows.Forms;
using System.Windows.Media;
using A;
using ExcelAddIn1.SuperFind2.UI;
using MacabacusMacros.Explorer;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.SuperFind2.Results;

public sealed class ShapeItem : ExploreItem
{
	private bool m_A;

	[CompilerGenerated]
	private Shape m_A;

	public override bool IsSelected
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			((BaseItem)this).NotifyPropertyChanged(VH.A(21693));
			A();
		}
	}

	internal Shape Shape
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[CompilerGenerated]
		set
		{
			this.m_A = value;
		}
	}

	public ShapeItem(WorksheetItem wsi, Shape shp)
		: base(wsi, Constants.ColorPalette.Cyan.Clone(), Props.Icons.GeoShape, 12)
	{
		Shape = shp;
		MsoShapeType type = shp.Type;
		if (type <= MsoShapeType.msoTextBox)
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			switch (type)
			{
			case MsoShapeType.msoPicture:
				break;
			case MsoShapeType.msoTextBox:
				goto IL_00b6;
			case MsoShapeType.msoFormControl:
				goto IL_00cf;
			case MsoShapeType.msoMedia:
				goto IL_0103;
			default:
				goto IL_011e;
			}
			((BaseItem)this).Icon = Geometry.Parse(VH.A(83164));
		}
		else
		{
			if ((uint)(type - 22) > 1u)
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
				if (type == MsoShapeType.msoWebVideo)
				{
					goto IL_0103;
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
				goto IL_011e;
			}
			((BaseItem)this).Icon = Geometry.Parse(VH.A(121115));
		}
		goto IL_0141;
		IL_00cf:
		((BaseItem)this).Icon = Geometry.Parse(VH.A(120699));
		goto IL_0141;
		IL_00b6:
		((BaseItem)this).Icon = Geometry.Parse(VH.A(120339));
		goto IL_0141;
		IL_0103:
		((BaseItem)this).Icon = Geometry.Parse(VH.A(122369));
		goto IL_0141;
		IL_011e:
		if (shp.Type == (MsoShapeType)28)
		{
			((BaseItem)this).Icon = Geometry.Parse(VH.A(83164));
		}
		goto IL_0141;
		IL_0141:
		if (shp.Visible == MsoTriState.msoFalse)
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
			base.FontColor.Opacity = ((BaseItem)this).HIDDEN_OPACITY;
			base.IconColor.Opacity = ((BaseItem)this).HIDDEN_OPACITY;
		}
		A();
	}

	public override void Refresh()
	{
		A();
		base.PreviewImage = null;
	}

	public override void Delete()
	{
		if (MessageBox.Show(VH.A(120239), VH.A(40448), MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation) != DialogResult.OK)
		{
			return;
		}
		while (true)
		{
			switch (3)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			Shape.Delete();
			base.Parent.A(this);
			return;
		}
	}

	public override void Search(string strQuery)
	{
		int isHighlighted;
		if (!((BaseItem)this).Label.ToLower().Contains(strQuery))
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			isHighlighted = ((Operators.CompareString(strQuery, VH.A(120326), TextCompare: false) == 0) ? 1 : 0);
		}
		else
		{
			isHighlighted = 1;
		}
		((BaseItem)this).IsHighlighted = (byte)isHighlighted != 0;
	}

	private void A()
	{
		((BaseItem)this).Label = Shape.Name;
	}
}
