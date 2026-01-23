using System;
using System.Windows.Forms;
using System.Windows.Media;
using A;
using MacabacusMacros.Explorer;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.Links;

namespace PowerPointAddIn1.Explorer;

public sealed class ChartItem : ContentItem
{
	private new bool m_A;

	public override bool IsSelected
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			((BaseItem)this).NotifyPropertyChanged(AH.A(62846));
			A();
		}
	}

	public ChartItem(SlideItem si, string strLabel, Microsoft.Office.Interop.PowerPoint.Shape shp, SolidColorBrush brush)
		: base(si, strLabel, brush, Pane.CachedObjects.GeoChart)
	{
		base.Shape = shp;
		base.IsLinked = PowerPointAddIn1.Links.Shapes.IsLinked(shp);
		base.IsLibraryContent = A(shp);
		UpdateColors(shp.Visible);
		string text = string.Empty;
		try
		{
			XlChartType chartType = shp.Chart.ChartType;
			if (chartType <= XlChartType.xl3DArea)
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
				if (chartType == XlChartType.xlXYScatter)
				{
					goto IL_019b;
				}
				while (true)
				{
					switch (6)
					{
					case 0:
						continue;
					}
					break;
				}
				if (chartType == XlChartType.xlDoughnut)
				{
					goto IL_01a3;
				}
				while (true)
				{
					switch (1)
					{
					case 0:
						continue;
					}
					break;
				}
				switch (chartType)
				{
				case XlChartType.xl3DPie:
					goto IL_0183;
				case XlChartType.xl3DArea:
					goto IL_018b;
				case XlChartType.xl3DLine:
					goto IL_0193;
				case XlChartType.xl3DColumn:
				case (XlChartType)(-4099):
					break;
				}
			}
			else
			{
				switch (chartType)
				{
				case XlChartType.xlPie:
					goto IL_0183;
				case XlChartType.xlArea:
					goto IL_018b;
				case XlChartType.xlLine:
					goto IL_0193;
				case XlChartType.xlBubble:
					goto IL_01ab;
				case (XlChartType)2:
				case (XlChartType)3:
					goto end_IL_004c;
				}
				while (true)
				{
					switch (7)
					{
					case 0:
						continue;
					}
					break;
				}
				switch (chartType)
				{
				case XlChartType.xlBarClustered:
				case XlChartType.xlBarStacked:
				case XlChartType.xlBarStacked100:
				case XlChartType.xl3DBarClustered:
				case XlChartType.xl3DBarStacked:
				case XlChartType.xl3DBarStacked100:
				case XlChartType.xlBarOfPie:
					text = Constants.DATA_CHART_BASIC;
					break;
				case XlChartType.xlPieOfPie:
				case XlChartType.xlPieExploded:
				case XlChartType.xl3DPieExploded:
					goto IL_0183;
				case XlChartType.xlAreaStacked:
				case XlChartType.xlAreaStacked100:
				case XlChartType.xl3DAreaStacked:
				case XlChartType.xl3DAreaStacked100:
					goto IL_018b;
				case XlChartType.xlLineStacked:
				case XlChartType.xlLineStacked100:
				case XlChartType.xlLineMarkers:
				case XlChartType.xlLineMarkersStacked:
				case XlChartType.xlLineMarkersStacked100:
					goto IL_0193;
				case XlChartType.xlXYScatterSmooth:
				case XlChartType.xlXYScatterSmoothNoMarkers:
				case XlChartType.xlXYScatterLines:
				case XlChartType.xlXYScatterLinesNoMarkers:
					goto IL_019b;
				case XlChartType.xlDoughnutExploded:
					goto IL_01a3;
				case XlChartType.xlBubble3DEffect:
					goto IL_01ab;
				case XlChartType.xlRadarMarkers:
				case XlChartType.xlRadarFilled:
				case XlChartType.xlSurface:
				case XlChartType.xlSurfaceWireframe:
				case XlChartType.xlSurfaceTopView:
				case XlChartType.xlSurfaceTopViewWireframe:
					break;
				}
			}
			goto end_IL_004c;
			IL_01a3:
			text = Constants.DATA_CHART_DONUT;
			goto end_IL_004c;
			IL_0183:
			text = Constants.DATA_CHART_PIE;
			goto end_IL_004c;
			IL_0193:
			text = Constants.DATA_CHART_LINE;
			goto end_IL_004c;
			IL_018b:
			text = Constants.DATA_CHART_BASIC;
			goto end_IL_004c;
			IL_01ab:
			text = Constants.DATA_CHART_BUBBLE;
			goto end_IL_004c;
			IL_019b:
			text = Constants.DATA_CHART_BASIC;
			end_IL_004c:;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		if (Operators.CompareString(text, string.Empty, TextCompare: false) != 0)
		{
			((BaseItem)this).Icon = Geometry.Parse(text);
		}
		SearchOnInstantiate();
	}

	public override void Refresh()
	{
		A();
		base.PreviewImage = null;
	}

	public override void Delete()
	{
		if (MessageBox.Show(AH.A(113231), AH.A(5874), MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation) != DialogResult.OK)
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
			base.Shape.Delete();
			base.Parent.RemoveChild(this);
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
			if (Operators.CompareString(strQuery, AH.A(113318), TextCompare: false) != 0)
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
				if (Operators.CompareString(strQuery, AH.A(113331), TextCompare: false) == 0)
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
					if (base.IsLinked)
					{
						goto IL_00a0;
					}
					while (true)
					{
						switch (6)
						{
						case 0:
							continue;
						}
						break;
					}
				}
				isHighlighted = ((Operators.CompareString(strQuery, AH.A(113342), TextCompare: false) == 0 && base.IsLibraryContent) ? 1 : 0);
				goto IL_00a1;
			}
		}
		goto IL_00a0;
		IL_00a0:
		isHighlighted = 1;
		goto IL_00a1;
		IL_00a1:
		((BaseItem)this).IsHighlighted = (byte)isHighlighted != 0;
	}

	public void SelectData()
	{
		try
		{
			NewLateBinding.LateCall(NewLateBinding.LateGet(base.Shape.Application, null, AH.A(113351), new object[0], null, null, null), null, AH.A(113374), new object[1] { AH.A(113395) }, null, null, null, IgnoreReturn: true);
			System.Windows.Forms.Application.DoEvents();
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	public void ChangeChartType()
	{
		try
		{
			NewLateBinding.LateCall(NewLateBinding.LateGet(base.Shape.Application, null, AH.A(113351), new object[0], null, null, null), null, AH.A(113374), new object[1] { AH.A(113434) }, null, null, null, IgnoreReturn: true);
			System.Windows.Forms.Application.DoEvents();
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	private void A()
	{
		Microsoft.Office.Interop.PowerPoint.Shape shape = base.Shape;
		string label = shape.Name;
		if (shape.Chart.HasTitle)
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
			label = shape.Chart.ChartTitle.Text;
		}
		shape = null;
		((BaseItem)this).Label = label;
	}
}
