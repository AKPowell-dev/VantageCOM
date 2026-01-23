using System;
using System.Drawing;
using System.Reflection;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace Macabacus_Word.Shapes;

public sealed class Guides
{
	private static readonly string m_A = XC.A(18392);

	public static void ShowExcel()
	{
		Microsoft.Office.Interop.Excel.Application application = null;
		try
		{
			application = InstanceManagement.GetExcelInstance(false);
			float a;
			float b;
			if (application.Selection is Microsoft.Office.Interop.Excel.Range)
			{
				Microsoft.Office.Interop.Excel.Range obj = (Microsoft.Office.Interop.Excel.Range)application.Selection;
				a = Conversions.ToSingle(obj.Width);
				b = Conversions.ToSingle(obj.Height);
				_ = null;
			}
			else
			{
				if (application.ActiveChart == null)
				{
					throw new Exception();
				}
				while (true)
				{
					switch (3)
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
				ChartObject obj2 = (ChartObject)application.ActiveChart.Parent;
				a = (float)obj2.Width;
				b = (float)obj2.Height;
				_ = null;
			}
			A(a, b);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			Forms.WarningMessage(XC.A(18013));
			ProjectData.ClearProjectError();
		}
		if (application == null)
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
			MC.A(application);
			application = null;
			return;
		}
	}

	public static void ShowPowerPoint()
	{
		Microsoft.Office.Interop.PowerPoint.Application application = (Microsoft.Office.Interop.PowerPoint.Application)PC.A.Application;
		Microsoft.Office.Interop.PowerPoint.Shape shape = null;
		try
		{
			Microsoft.Office.Interop.PowerPoint.ShapeRange shapeRange = application.ActiveWindow.Selection.ShapeRange;
			if (shapeRange.Count == 1)
			{
				shape = shapeRange[1];
			}
			shapeRange = null;
			if (shape == null)
			{
				throw new Exception();
			}
			A(shape.Width, shape.Height);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			Forms.WarningMessage(XC.A(18066));
			ProjectData.ClearProjectError();
		}
		application = null;
		shape = null;
	}

	public static void ShowWord()
	{
		try
		{
			float[] array = clsPublish.SelectedWordShapeSize((Microsoft.Office.Interop.Word.Application)Interaction.GetObject(null, XC.A(18181)));
			if (array[0] > 0f)
			{
				A(array[0], array[1]);
				return;
			}
			throw new Exception();
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			Forms.WarningMessage(XC.A(18214));
			ProjectData.ClearProjectError();
		}
	}

	public static void Show(int i)
	{
		//IL_0001: Unknown result type (might be due to invalid IL or missing references)
		//IL_0006: Unknown result type (might be due to invalid IL or missing references)
		//IL_0008: Unknown result type (might be due to invalid IL or missing references)
		//IL_0009: Unknown result type (might be due to invalid IL or missing references)
		//IL_0016: Unknown result type (might be due to invalid IL or missing references)
		//IL_004e: Unknown result type (might be due to invalid IL or missing references)
		StandardSize standardSize;
		try
		{
			standardSize = clsPublish.GetStandardSize(i);
			A(clsPublish.InchesToPoints(standardSize.Width), clsPublish.InchesToPoints(standardSize.Height));
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			Forms.ErrorMessage(XC.A(18317));
			ProjectData.ClearProjectError();
		}
		standardSize = default(StandardSize);
	}

	private static Microsoft.Office.Interop.Word.Shape A(float A, float B)
	{
		Color gUIDE_COLOR = clsGuides.GUIDE_COLOR;
		PageSetup pageSetup = PC.A.Application.ActiveDocument.PageSetup;
		float pageWidth = pageSetup.PageWidth;
		float pageHeight = pageSetup.PageHeight;
		_ = null;
		float left = pageWidth / 2f - A / 2f;
		float top = pageHeight / 2f - B / 2f;
		Microsoft.Office.Interop.Word.Shapes shapes = PC.A.Application.ActiveDocument.Shapes;
		object Anchor = RuntimeHelpers.GetObjectValue(Missing.Value);
		Microsoft.Office.Interop.Word.Shape shape = shapes.AddShape(1, left, top, A, B, ref Anchor);
		shape.Fill.Visible = MsoTriState.msoFalse;
		shape.Line.ForeColor.RGB = Information.RGB(gUIDE_COLOR.R, gUIDE_COLOR.G, gUIDE_COLOR.B);
		shape.Line.Weight = 1f;
		shape.ZOrder(MsoZOrderCmd.msoBringToFront);
		shape.Name = Guides.m_A;
		Anchor = RuntimeHelpers.GetObjectValue(Missing.Value);
		shape.Select(ref Anchor);
		_ = null;
		return shape;
	}

	public static void Remove()
	{
		try
		{
			Document activeDocument = PC.A.Application.ActiveDocument;
			for (int i = activeDocument.Shapes.Count; i >= 1; i = checked(i + -1))
			{
				Microsoft.Office.Interop.Word.Shapes shapes = activeDocument.Shapes;
				object Index = i;
				Microsoft.Office.Interop.Word.Shape shape = shapes[ref Index];
				if (Operators.CompareString(shape.Name, Guides.m_A, TextCompare: false) == 0)
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
					shape.Delete();
				}
				shape = null;
			}
			while (true)
			{
				switch (1)
				{
				case 0:
					continue;
				}
				activeDocument = null;
				return;
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}
}
