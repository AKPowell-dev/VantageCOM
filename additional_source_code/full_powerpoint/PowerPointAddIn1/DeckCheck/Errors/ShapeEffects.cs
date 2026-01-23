using System;
using System.Collections.Generic;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class ShapeEffects : BaseError
{
	public ShapeEffects(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp)
		: base(ErrorType.ShapeEffects, Main.Analysis.Options.ShapeEffects, sld, shp, blnHasFix: true, blnCanFixMultiple: true)
	{
		//IL_0011: Unknown result type (might be due to invalid IL or missing references)
		//IL_0016: Unknown result type (might be due to invalid IL or missing references)
		A();
	}

	public ShapeEffects(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, List<Microsoft.Office.Interop.PowerPoint.Shape> listShapes)
		: base(ErrorType.ShapeEffects, Main.Analysis.Options.ShapeEffects, sld, shp, blnHasFix: true, blnCanFixMultiple: true)
	{
		//IL_0011: Unknown result type (might be due to invalid IL or missing references)
		//IL_0016: Unknown result type (might be due to invalid IL or missing references)
		A();
		base.Shapes = listShapes;
	}

	public ShapeEffects(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, List<Microsoft.Office.Core.Shape> listShapes)
		: base(ErrorType.ShapeEffects, Main.Analysis.Options.ShapeEffects, sld, shp, blnHasFix: true, blnCanFixMultiple: true)
	{
		//IL_0011: Unknown result type (might be due to invalid IL or missing references)
		A();
		((BaseError)this).OfficeShapes = listShapes;
	}

	public ShapeEffects(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, ChartTitle title)
		: base(ErrorType.ShapeEffects, Main.Analysis.Options.ShapeEffects, sld, shp, blnHasFix: true, blnCanFixMultiple: true)
	{
		//IL_0011: Unknown result type (might be due to invalid IL or missing references)
		A();
		base.ChartTitle = title;
	}

	public ShapeEffects(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, AxisTitle title)
		: base(ErrorType.ShapeEffects, Main.Analysis.Options.ShapeEffects, sld, shp, blnHasFix: true, blnCanFixMultiple: true)
	{
		//IL_000d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0012: Unknown result type (might be due to invalid IL or missing references)
		A();
		base.AxisTitle = title;
	}

	public ShapeEffects(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, Axis axis)
		: base(ErrorType.ShapeEffects, Main.Analysis.Options.ShapeEffects, sld, shp, blnHasFix: true, blnCanFixMultiple: true)
	{
		//IL_000f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0014: Unknown result type (might be due to invalid IL or missing references)
		A();
		base.Axis = axis;
	}

	public ShapeEffects(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, Legend legend)
		: base(ErrorType.ShapeEffects, Main.Analysis.Options.ShapeEffects, sld, shp, blnHasFix: true, blnCanFixMultiple: true)
	{
		//IL_0011: Unknown result type (might be due to invalid IL or missing references)
		A();
		base.Legend = legend;
	}

	public ShapeEffects(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, DataTable tbl)
		: base(ErrorType.ShapeEffects, Main.Analysis.Options.ShapeEffects, sld, shp, blnHasFix: true, blnCanFixMultiple: true)
	{
		//IL_000f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0014: Unknown result type (might be due to invalid IL or missing references)
		A();
		base.DataTable = tbl;
	}

	public ShapeEffects(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, PlotArea plotArea)
		: base(ErrorType.ShapeEffects, Main.Analysis.Options.ShapeEffects, sld, shp, blnHasFix: true, blnCanFixMultiple: true)
	{
		//IL_000f: Unknown result type (might be due to invalid IL or missing references)
		A();
		base.PlotArea = plotArea;
	}

	public ShapeEffects(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, IMsoSeries series)
		: base(ErrorType.ShapeEffects, Main.Analysis.Options.ShapeEffects, sld, shp, blnHasFix: true, blnCanFixMultiple: true)
	{
		//IL_000d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0012: Unknown result type (might be due to invalid IL or missing references)
		A();
		((BaseError)this).Series = series;
	}

	private void A()
	{
		BaseError val = (BaseError)(object)this;
		Errors.ShapeEffects(ref val);
	}

	public override void FixAction()
	{
		NG.A.Application.StartNewUndoEntry();
		if (((BaseError)this).OfficeShapes != null)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
				{
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					using IEnumerator<Microsoft.Office.Core.Shape> enumerator = ((BaseError)this).OfficeShapes.GetEnumerator();
					while (enumerator.MoveNext())
					{
						Microsoft.Office.Core.Shape current = enumerator.Current;
						try
						{
							current.Reflection.Type = MsoReflectionType.msoReflectionTypeNone;
							current.Reflection.Size = 0f;
						}
						catch (Exception ex)
						{
							ProjectData.SetProjectError(ex);
							Exception ex2 = ex;
							ProjectData.ClearProjectError();
						}
						try
						{
							current.Glow.Radius = 0f;
						}
						catch (Exception ex3)
						{
							ProjectData.SetProjectError(ex3);
							Exception ex4 = ex3;
							ProjectData.ClearProjectError();
						}
						try
						{
							Microsoft.Office.Core.ThreeDFormat threeD = current.ThreeD;
							threeD.ResetRotation();
							threeD.BevelTopType = MsoBevelType.msoBevelNone;
							threeD.BevelBottomType = MsoBevelType.msoBevelNone;
							threeD.Visible = MsoTriState.msoFalse;
							_ = null;
						}
						catch (Exception ex5)
						{
							ProjectData.SetProjectError(ex5);
							Exception ex6 = ex5;
							ProjectData.ClearProjectError();
						}
						try
						{
							current.SoftEdge.Type = MsoSoftEdgeType.msoSoftEdgeTypeNone;
						}
						catch (Exception ex7)
						{
							ProjectData.SetProjectError(ex7);
							Exception ex8 = ex7;
							ProjectData.ClearProjectError();
						}
						try
						{
							current.Shadow.Size = 0f;
							current.Shadow.Blur = 0f;
							current.Shadow.Transparency = 1f;
							current.Shadow.Visible = MsoTriState.msoFalse;
						}
						catch (Exception ex9)
						{
							ProjectData.SetProjectError(ex9);
							Exception ex10 = ex9;
							ProjectData.ClearProjectError();
						}
						current = null;
					}
					while (true)
					{
						switch (1)
						{
						case 0:
							break;
						default:
							return;
						}
					}
				}
				}
			}
		}
		if (base.ChartTitle != null)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					A(base.ChartTitle.Format);
					return;
				}
			}
		}
		if (base.AxisTitle != null)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					A(base.AxisTitle.Format);
					return;
				}
			}
		}
		if (base.Axis != null)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					A(base.Axis.Format);
					return;
				}
			}
		}
		if (base.Legend != null)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					A(base.Legend.Format);
					return;
				}
			}
		}
		if (base.PlotArea != null)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					A(base.PlotArea.Format);
					return;
				}
			}
		}
		if (base.DataTable != null)
		{
			A(base.DataTable.Format);
			return;
		}
		if (((BaseError)this).Series != null)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
				{
					IMsoChartFormat format = ((BaseError)this).Series.Format;
					try
					{
						format.Glow.Radius = 0f;
					}
					catch (Exception ex11)
					{
						ProjectData.SetProjectError(ex11);
						Exception ex12 = ex11;
						ProjectData.ClearProjectError();
					}
					try
					{
						Microsoft.Office.Core.ThreeDFormat threeD2 = format.ThreeD;
						threeD2.ResetRotation();
						threeD2.BevelTopType = MsoBevelType.msoBevelNone;
						threeD2.BevelBottomType = MsoBevelType.msoBevelNone;
						_ = null;
					}
					catch (Exception ex13)
					{
						ProjectData.SetProjectError(ex13);
						Exception ex14 = ex13;
						ProjectData.ClearProjectError();
					}
					try
					{
						format.SoftEdge.Type = MsoSoftEdgeType.msoSoftEdgeTypeNone;
					}
					catch (Exception ex15)
					{
						ProjectData.SetProjectError(ex15);
						Exception ex16 = ex15;
						ProjectData.ClearProjectError();
					}
					try
					{
						Microsoft.Office.Core.ShadowFormat shadow = format.Shadow;
						if (shadow.Transparency >= 0f)
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
							if (shadow.Transparency < 1f)
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
								shadow.Transparency = 1f;
							}
						}
						shadow = null;
					}
					catch (Exception ex17)
					{
						ProjectData.SetProjectError(ex17);
						Exception ex18 = ex17;
						ProjectData.ClearProjectError();
					}
					format = null;
					return;
				}
				}
			}
		}
		if (base.Shapes != null)
		{
			using (List<Microsoft.Office.Interop.PowerPoint.Shape>.Enumerator enumerator2 = base.Shapes.GetEnumerator())
			{
				while (enumerator2.MoveNext())
				{
					Microsoft.Office.Interop.PowerPoint.Shape current2 = enumerator2.Current;
					try
					{
						Microsoft.Office.Interop.PowerPoint.ThreeDFormat threeD3 = current2.ThreeD;
						threeD3.BevelTopType = MsoBevelType.msoBevelNone;
						threeD3.BevelBottomType = MsoBevelType.msoBevelNone;
						threeD3.Visible = MsoTriState.msoFalse;
						_ = null;
					}
					catch (Exception ex19)
					{
						ProjectData.SetProjectError(ex19);
						Exception ex20 = ex19;
						ProjectData.ClearProjectError();
					}
				}
				while (true)
				{
					switch (4)
					{
					case 0:
						break;
					default:
						return;
					}
				}
			}
		}
		Microsoft.Office.Interop.PowerPoint.Shape shape = base.Shape;
		try
		{
			shape.Reflection.Type = MsoReflectionType.msoReflectionTypeNone;
			shape.Reflection.Size = 0f;
		}
		catch (Exception ex21)
		{
			ProjectData.SetProjectError(ex21);
			Exception ex22 = ex21;
			ProjectData.ClearProjectError();
		}
		try
		{
			shape.Glow.Radius = 0f;
		}
		catch (Exception ex23)
		{
			ProjectData.SetProjectError(ex23);
			Exception ex24 = ex23;
			ProjectData.ClearProjectError();
		}
		try
		{
			Microsoft.Office.Interop.PowerPoint.ThreeDFormat threeD4 = shape.ThreeD;
			threeD4.ResetRotation();
			threeD4.BevelTopType = MsoBevelType.msoBevelNone;
			threeD4.BevelBottomType = MsoBevelType.msoBevelNone;
			threeD4.Visible = MsoTriState.msoFalse;
			_ = null;
		}
		catch (Exception ex25)
		{
			ProjectData.SetProjectError(ex25);
			Exception ex26 = ex25;
			ProjectData.ClearProjectError();
		}
		try
		{
			shape.SoftEdge.Type = MsoSoftEdgeType.msoSoftEdgeTypeNone;
		}
		catch (Exception ex27)
		{
			ProjectData.SetProjectError(ex27);
			Exception ex28 = ex27;
			ProjectData.ClearProjectError();
		}
		try
		{
			shape.Shadow.Size = 0f;
			shape.Shadow.Blur = 0f;
			shape.Shadow.Transparency = 1f;
			shape.Shadow.Visible = MsoTriState.msoFalse;
		}
		catch (Exception ex29)
		{
			ProjectData.SetProjectError(ex29);
			Exception ex30 = ex29;
			ProjectData.ClearProjectError();
		}
		shape = null;
	}

	private void A(ChartFormat A)
	{
		ChartFormat chartFormat = A;
		try
		{
			chartFormat.Glow.Radius = 0f;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		try
		{
			Microsoft.Office.Interop.PowerPoint.ThreeDFormat threeD = chartFormat.ThreeD;
			threeD.ResetRotation();
			threeD.BevelTopType = MsoBevelType.msoBevelNone;
			threeD.BevelBottomType = MsoBevelType.msoBevelNone;
			_ = null;
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			ProjectData.ClearProjectError();
		}
		try
		{
			chartFormat.SoftEdge.Type = MsoSoftEdgeType.msoSoftEdgeTypeNone;
		}
		catch (Exception ex5)
		{
			ProjectData.SetProjectError(ex5);
			Exception ex6 = ex5;
			ProjectData.ClearProjectError();
		}
		try
		{
			Microsoft.Office.Interop.PowerPoint.ShadowFormat shadow = chartFormat.Shadow;
			if (shadow.Transparency >= 0f)
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
				if (shadow.Transparency < 1f)
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
					shadow.Transparency = 1f;
				}
			}
			shadow = null;
		}
		catch (Exception ex7)
		{
			ProjectData.SetProjectError(ex7);
			Exception ex8 = ex7;
			ProjectData.ClearProjectError();
		}
		chartFormat = null;
	}
}
