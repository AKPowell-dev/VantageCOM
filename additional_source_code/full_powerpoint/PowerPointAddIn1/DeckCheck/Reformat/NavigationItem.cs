using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Windows;
using A;
using MacabacusMacros.Proofing.UI;
using MacabacusMacros.Proofing.UI.Reformat;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.DeckCheck.Reformat;

public sealed class NavigationItem : BaseItem
{
	[CompilerGenerated]
	private string A;

	[CompilerGenerated]
	private IndexedObject A;

	public string IconPath
	{
		[CompilerGenerated]
		get
		{
			return this.A;
		}
		[CompilerGenerated]
		set
		{
			this.A = value;
		}
	}

	public IndexedObject IndexedObject
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

	public NavigationItem(BaseItem parent, IndexedObject obj, DataTemplate template, string strHeader, string strDefaultLabel = "")
		: base(0, 0, new List<IndexedObject>(), template, template, strHeader)
	{
		((BaseItem)this).IsVisible = false;
		IndexedObject = obj;
		IndexedObject indexedObject = obj;
		if (indexedObject.Child is Cell)
		{
			((BaseItem)this).Label = ((Microsoft.Office.Interop.PowerPoint.Shape)NewLateBinding.LateGet(NewLateBinding.LateGet(indexedObject.Child, null, AH.A(28234), new object[0], null, null, null), null, AH.A(28234), new object[0], null, null, null)).Name + AH.A(50122);
			IconPath = Icons.TABLE;
		}
		else if (indexedObject.Child is TextRange2)
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			string label = Strings.Trim(((TextRange2)indexedObject.Child).Text).Replace(AH.A(7894), AH.A(14625)).Replace(AH.A(47331), AH.A(14625))
				.Replace(AH.A(47334), AH.A(14625));
			_ = null;
			((BaseItem)this).Label = label;
		}
		else if (indexedObject.Child is BulletFormat2)
		{
			((BaseItem)this).Label = indexedObject.Shape.Name + AH.A(49166);
		}
		else
		{
			string label = indexedObject.Shape.Name;
			if (indexedObject.Child is Microsoft.Office.Interop.PowerPoint.Shape)
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
				((BaseItem)this).Label = label;
			}
			else if (indexedObject.Shape.HasChart == MsoTriState.msoTrue)
			{
				if (indexedObject.Child is Chart)
				{
					label += AH.A(50137);
				}
				else if (indexedObject.Child is ChartArea)
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
					label += AH.A(50154);
				}
				else if (indexedObject.Child is PlotArea)
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
					label += AH.A(50181);
				}
				else if (indexedObject.Child is Legend)
				{
					label += AH.A(50206);
				}
				else if (indexedObject.Child is Microsoft.Office.Core.LegendEntry)
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
					label += AH.A(50225);
				}
				else if (indexedObject.Child is ChartTitle)
				{
					label += AH.A(50256);
				}
				else if (indexedObject.Child is DataTable)
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
					label += AH.A(50285);
				}
				else if (indexedObject.Child is AxisTitle)
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
					label += AH.A(50312);
				}
				else if (indexedObject.Child is TickLabels)
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
					label += AH.A(50339);
				}
				else if (indexedObject.Child is Axis)
				{
					label += AH.A(50368);
				}
				else if (indexedObject.Child is Gridlines)
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
					label += AH.A(50383);
				}
				else if (indexedObject.Child is HiLoLines)
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
					label += AH.A(50408);
				}
				else if (indexedObject.Child is DropLines)
				{
					label += AH.A(50443);
				}
				else if (indexedObject.Child is UpBars)
				{
					label += AH.A(50470);
				}
				else if (indexedObject.Child is DownBars)
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
					label += AH.A(50491);
				}
				else if (indexedObject.Child is IMsoErrorBars)
				{
					label += AH.A(50516);
				}
				else if (indexedObject.Child is IMsoLeaderLines)
				{
					label += AH.A(50543);
				}
				else if (indexedObject.Child is IMsoTrendline)
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
					label += AH.A(50574);
				}
				else if (indexedObject.Child is IMsoSeries)
				{
					label += AH.A(50601);
				}
				else if (indexedObject.Child is ChartPoint)
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
					label += AH.A(50620);
				}
				else if (indexedObject.Child is IMsoDataLabels)
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
					label += AH.A(50637);
				}
				else if (indexedObject.Child is IMsoDataLabel)
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
					label += AH.A(50666);
				}
				else if (indexedObject.Child is IMsoLegendKey)
				{
					string text = label;
					string text2;
					if (indexedObject.IsMarker)
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
						text2 = AH.A(50693);
					}
					else
					{
						text2 = AH.A(50734);
					}
					label = text + text2;
				}
				else
				{
					label = label + AH.A(14625) + Versioned.TypeName(RuntimeHelpers.GetObjectValue(indexedObject.Child));
				}
				((BaseItem)this).Label = label;
			}
			else if (indexedObject.Child is SmartArt)
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
				((BaseItem)this).Label = label + AH.A(50761);
			}
			else if (indexedObject.Child is Microsoft.Office.Core.Shape)
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
				((BaseItem)this).Label = label + AH.A(50786);
			}
			else
			{
				((BaseItem)this).Label = Versioned.TypeName(RuntimeHelpers.GetObjectValue(indexedObject.Child));
			}
		}
		indexedObject = null;
		if (strDefaultLabel.Length > 0)
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
			((BaseItem)this).Label = strDefaultLabel;
		}
		if (Operators.CompareString(IconPath, string.Empty, TextCompare: false) != 0)
		{
			return;
		}
		while (true)
		{
			switch (1)
			{
			case 0:
				continue;
			}
			if (obj.Child is Cell)
			{
				return;
			}
			while (true)
			{
				switch (2)
				{
				case 0:
					continue;
				}
				Microsoft.Office.Interop.PowerPoint.Shape shape = obj.Shape;
				MsoShapeType type = shape.Type;
				if (type != MsoShapeType.msoAutoShape)
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
					if (type == MsoShapeType.msoTextBox)
					{
						IconPath = Icons.TEXT;
					}
					else if (shape.HasChart == MsoTriState.msoTrue)
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
						IconPath = Icons.CHART;
					}
					else if (shape.HasTable == MsoTriState.msoTrue)
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
						IconPath = Icons.TABLE;
					}
					else if (shape.HasSmartArt == MsoTriState.msoTrue)
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
						IconPath = Icons.SMART_ART;
					}
					else if (shape.Type == MsoShapeType.msoPlaceholder)
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
						IconPath = Icons.PLACEHOLDER;
					}
					else
					{
						IconPath = Icons.SHAPE;
					}
				}
				else
				{
					IconPath = Icons.SHAPE;
				}
				shape = null;
				return;
			}
		}
	}

	public void SetBulletIcon(BulletFormat2 bf)
	{
		string iconPath;
		if (bf.Type != MsoBulletType.msoBulletNumbered)
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
			iconPath = Icons.BULLET;
		}
		else
		{
			iconPath = Icons.LIST;
		}
		IconPath = iconPath;
	}
}
