using System;
using System.Collections;
using System.Drawing;
using System.Text;
using System.Xml;
using A;
using MacabacusMacros;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.Shapes;

public sealed class Guides
{
	private static readonly string m_A = AH.A(82690);

	public static void ShowExcel()
	{
		if (!Licensing.AllowRestrictedMode())
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			Microsoft.Office.Interop.Excel.Application application = null;
			try
			{
				application = InstanceManagement.GetExcelInstance(false);
				float a;
				float b;
				if (application.Selection is Range)
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
					Range obj = (Range)application.Selection;
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
				Forms.WarningMessage(AH.A(82118));
				ProjectData.ClearProjectError();
			}
			if (application == null)
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
				JG.A(application);
				application = null;
				return;
			}
		}
	}

	public static void ShowPowerPoint()
	{
		if (!Licensing.AllowRestrictedMode())
		{
			return;
		}
		Microsoft.Office.Interop.PowerPoint.Application application = NG.A.Application;
		Microsoft.Office.Interop.PowerPoint.Shape shape = null;
		try
		{
			Microsoft.Office.Interop.PowerPoint.ShapeRange shapeRange = application.ActiveWindow.Selection.ShapeRange;
			if (shapeRange.Count == 1)
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
				shape = shapeRange[1];
			}
			shapeRange = null;
			if (shape == null)
			{
				while (true)
				{
					switch (6)
					{
					case 0:
						continue;
					}
					throw new Exception();
				}
			}
			A(shape.Width, shape.Height);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			Forms.WarningMessage(AH.A(82171));
			ProjectData.ClearProjectError();
		}
		application = null;
		shape = null;
	}

	public static void ShowWord()
	{
		if (!Licensing.AllowRestrictedMode())
		{
			return;
		}
		while (true)
		{
			switch (6)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			try
			{
				float[] array = clsPublish.SelectedWordShapeSize((Microsoft.Office.Interop.Word.Application)Interaction.GetObject(null, AH.A(82290)));
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
				Forms.WarningMessage(AH.A(82323));
				ProjectData.ClearProjectError();
				return;
			}
		}
	}

	public static void Show(int i)
	{
		//IL_0069: Unknown result type (might be due to invalid IL or missing references)
		//IL_001c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0021: Unknown result type (might be due to invalid IL or missing references)
		//IL_0023: Unknown result type (might be due to invalid IL or missing references)
		//IL_0024: Unknown result type (might be due to invalid IL or missing references)
		//IL_0031: Unknown result type (might be due to invalid IL or missing references)
		if (!Licensing.AllowRestrictedMode())
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
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
				Forms.ErrorMessage(AH.A(82430));
				ProjectData.ClearProjectError();
			}
			standardSize = default(StandardSize);
			return;
		}
	}

	private static Microsoft.Office.Interop.PowerPoint.Shape A(float A, float B)
	{
		Color gUIDE_COLOR = clsGuides.GUIDE_COLOR;
		Slide slide = NG.A.Application.ActiveWindow.Selection.SlideRange[1];
		CustomLayout customLayout = slide.CustomLayout;
		float width = customLayout.Width;
		float height = customLayout.Height;
		_ = null;
		float left = width / 2f - A / 2f;
		float top = height / 2f - B / 2f;
		Microsoft.Office.Interop.PowerPoint.Shape shape = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, left, top, A, B);
		shape.Fill.Visible = MsoTriState.msoFalse;
		shape.Line.ForeColor.RGB = Information.RGB(gUIDE_COLOR.R, gUIDE_COLOR.G, gUIDE_COLOR.B);
		shape.Line.Weight = 1f;
		shape.ZOrder(MsoZOrderCmd.msoBringToFront);
		shape.Name = Guides.m_A;
		shape.Select();
		_ = null;
		return shape;
	}

	public static void Remove()
	{
		Slide slide;
		try
		{
			slide = NG.A.Application.ActiveWindow.Selection.SlideRange[1];
			for (int i = slide.Shapes.Count; i >= 1; i = checked(i + -1))
			{
				Microsoft.Office.Interop.PowerPoint.Shape shape = slide.Shapes[i];
				if (Operators.CompareString(shape.Name, Guides.m_A, TextCompare: false) == 0)
				{
					shape.Delete();
				}
				shape = null;
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
				break;
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		slide = null;
	}

	public static string ShowGuideMenu()
	{
		int num = 1;
		StringBuilder stringBuilder = new StringBuilder(AH.A(47526));
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = clsPublish.StandardSizeNodes().GetEnumerator();
			while (enumerator.MoveNext())
			{
				string text = ((XmlNode)enumerator.Current).Attributes[AH.A(82505)].Value.Replace(AH.A(82514), AH.A(82517));
				string text2 = num + AH.A(82538) + text;
				if (num < 10)
				{
					text2 = AH.A(82543) + text2;
				}
				stringBuilder.Append(AH.A(82554) + num + AH.A(47705) + text2 + AH.A(82597) + num + AH.A(82654) + text + AH.A(82681));
				num = checked(num + 1);
			}
			while (true)
			{
				switch (1)
				{
				case 0:
					continue;
				}
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				break;
			}
		}
		finally
		{
			if (enumerator is IDisposable)
			{
				while (true)
				{
					switch (1)
					{
					case 0:
						continue;
					}
					(enumerator as IDisposable).Dispose();
					break;
				}
			}
		}
		stringBuilder.Append(AH.A(49007));
		return stringBuilder.ToString();
	}
}
