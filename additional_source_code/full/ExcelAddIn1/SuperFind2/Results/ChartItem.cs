using System;
using System.Runtime.CompilerServices;
using System.Windows.Forms;
using System.Windows.Media;
using A;
using ExcelAddIn1.SuperFind2.UI;
using MacabacusMacros.Explorer;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.SuperFind2.Results;

public sealed class ChartItem : ExploreItem
{
	private bool m_A;

	[CompilerGenerated]
	private ChartObject m_A;

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
			Refresh();
		}
	}

	internal ChartObject ChartObject
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

	public ChartItem(WorksheetItem wsi, ChartObject chtObj)
		: base(wsi, Constants.ColorPalette.Green.Clone(), Props.Icons.GeoChart, 10)
	{
		ChartObject = chtObj;
		Refresh();
		string text = string.Empty;
		try
		{
			XlChartType chartType = chtObj.Chart.ChartType;
			if (chartType <= XlChartType.xl3DArea)
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
				switch (chartType)
				{
				case XlChartType.xl3DPie:
					break;
				case XlChartType.xl3DArea:
					goto IL_0166;
				case XlChartType.xl3DLine:
					goto IL_016e;
				case XlChartType.xlXYScatter:
					goto IL_0176;
				case XlChartType.xlDoughnut:
					goto IL_017e;
				default:
					goto end_IL_003d;
				}
				goto IL_015e;
			}
			switch (chartType)
			{
			case XlChartType.xlPie:
				goto IL_015e;
			case XlChartType.xlArea:
				goto IL_0166;
			case XlChartType.xlLine:
				goto IL_016e;
			case XlChartType.xlBubble:
				goto IL_0186;
			case (XlChartType)2:
			case (XlChartType)3:
				goto end_IL_003d;
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
				goto IL_015e;
			case XlChartType.xlAreaStacked:
			case XlChartType.xlAreaStacked100:
			case XlChartType.xl3DAreaStacked:
			case XlChartType.xl3DAreaStacked100:
				goto IL_0166;
			case XlChartType.xlLineStacked:
			case XlChartType.xlLineStacked100:
			case XlChartType.xlLineMarkers:
			case XlChartType.xlLineMarkersStacked:
			case XlChartType.xlLineMarkersStacked100:
				goto IL_016e;
			case XlChartType.xlXYScatterSmooth:
			case XlChartType.xlXYScatterSmoothNoMarkers:
			case XlChartType.xlXYScatterLines:
			case XlChartType.xlXYScatterLinesNoMarkers:
				goto IL_0176;
			case XlChartType.xlDoughnutExploded:
				goto IL_017e;
			case XlChartType.xlBubble3DEffect:
				goto IL_0186;
			case XlChartType.xlRadarMarkers:
			case XlChartType.xlRadarFilled:
			case XlChartType.xlSurface:
			case XlChartType.xlSurfaceWireframe:
			case XlChartType.xlSurfaceTopView:
			case XlChartType.xlSurfaceTopViewWireframe:
				break;
			}
			goto end_IL_003d;
			IL_017e:
			text = Constants.DATA_CHART_DONUT;
			goto end_IL_003d;
			IL_015e:
			text = Constants.DATA_CHART_PIE;
			goto end_IL_003d;
			IL_016e:
			text = Constants.DATA_CHART_LINE;
			goto end_IL_003d;
			IL_0166:
			text = Constants.DATA_CHART_BASIC;
			goto end_IL_003d;
			IL_0186:
			text = Constants.DATA_CHART_BUBBLE;
			goto end_IL_003d;
			IL_0176:
			text = Constants.DATA_CHART_BASIC;
			end_IL_003d:;
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
	}

	public override void Refresh()
	{
		A();
		base.PreviewImage = null;
	}

	public override void Delete()
	{
		if (MessageBox.Show(VH.A(118513), VH.A(40448), MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation) != DialogResult.OK)
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			ChartObject.Delete();
			base.Parent.A(this);
			return;
		}
	}

	public override void Search(string strQuery)
	{
		((BaseItem)this).IsHighlighted = ((BaseItem)this).Label.ToLower().Contains(strQuery) || Operators.CompareString(strQuery, VH.A(118600), TextCompare: false) == 0;
	}

	public void SelectData()
	{
		try
		{
			ChartObject.Application.CommandBars.ExecuteMso(VH.A(118613));
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
			ChartObject.Application.CommandBars.ExecuteMso(VH.A(118652));
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
		ChartObject chartObject = ChartObject;
		((BaseItem)this).Label = chartObject.Name;
		if (chartObject.Chart.HasTitle)
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
			((BaseItem)this).Label = chartObject.Chart.ChartTitle.Text;
		}
		chartObject = null;
	}
}
