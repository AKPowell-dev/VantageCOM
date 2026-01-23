using System;
using System.Collections;
using A;
using MacabacusMacros;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.Shapes;

public sealed class Helpers
{
	public static readonly string TAG_SHAPE_TYPE = AH.A(74515);

	public static string GetShapeType(Microsoft.Office.Interop.PowerPoint.Shape shp)
	{
		string result;
		try
		{
			result = shp.Tags[TAG_SHAPE_TYPE];
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			result = "";
			ProjectData.ClearProjectError();
		}
		return result;
	}

	public static bool IsShapeType(Microsoft.Office.Interop.PowerPoint.Shape shp, string strType)
	{
		bool result;
		try
		{
			result = Operators.CompareString(GetShapeType(shp), strType, TextCompare: false) == 0;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			result = false;
			ProjectData.ClearProjectError();
		}
		return result;
	}

	public static bool IsShapeMatch(Microsoft.Office.Interop.PowerPoint.Shape shp1, Microsoft.Office.Interop.PowerPoint.Shape shp2)
	{
		if (shp1.Top == shp2.Top)
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
			if (shp1.Left == shp2.Left)
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
				if (shp1.Width == shp2.Width)
				{
					while (true)
					{
						switch (3)
						{
						case 0:
							break;
						default:
							return shp1.Height == shp2.Height;
						}
					}
				}
			}
		}
		return false;
	}

	public static Microsoft.Office.Interop.PowerPoint.Shape GetBodyPlaceholder(Microsoft.Office.Interop.PowerPoint.Presentation pres)
	{
		Microsoft.Office.Interop.PowerPoint.Shape result = null;
		IEnumerator enumerator = pres.Designs[1].SlideMaster.Shapes.GetEnumerator();
		try
		{
			while (true)
			{
				if (enumerator.MoveNext())
				{
					Microsoft.Office.Interop.PowerPoint.Shape shape = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current;
					if (shape.Type != MsoShapeType.msoPlaceholder)
					{
						continue;
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					if (shape.PlaceholderFormat.Type != PpPlaceholderType.ppPlaceholderBody)
					{
						continue;
					}
					while (true)
					{
						switch (3)
						{
						case 0:
							continue;
						}
						result = shape;
						break;
					}
					break;
				}
				while (true)
				{
					switch (3)
					{
					case 0:
						break;
					default:
						goto end_IL_0086;
					}
					continue;
					end_IL_0086:
					break;
				}
				break;
			}
		}
		finally
		{
			IDisposable disposable = enumerator as IDisposable;
			if (disposable != null)
			{
				disposable.Dispose();
			}
		}
		return result;
	}

	internal static int A(Slide A, Microsoft.Office.Interop.PowerPoint.Shape B)
	{
		return clsPowerPoint.GetShapeIndex(A.Shapes, B);
	}

	internal static int A(CustomLayout A, Microsoft.Office.Interop.PowerPoint.Shape B)
	{
		return clsPowerPoint.GetShapeIndex(A.Shapes, B);
	}

	internal static int A(Microsoft.Office.Interop.PowerPoint.Shape A)
	{
		Microsoft.Office.Interop.PowerPoint.Shapes shapes = null;
		try
		{
			try
			{
				shapes = clsPowerPoint.GetSlideFromShape(A).Shapes;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			if (shapes == null)
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
				try
				{
					shapes = clsPowerPoint.GetLayoutFromShape(A).Shapes;
				}
				catch (Exception ex3)
				{
					ProjectData.SetProjectError(ex3);
					Exception ex4 = ex3;
					ProjectData.ClearProjectError();
				}
			}
			return clsPowerPoint.GetShapeIndex(shapes, A);
		}
		finally
		{
			shapes = null;
		}
	}

	public static void SingleShapeRequiredError()
	{
		//IL_000d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0012: Unknown result type (might be due to invalid IL or missing references)
		//IL_0014: Unknown result type (might be due to invalid IL or missing references)
		//IL_0015: Unknown result type (might be due to invalid IL or missing references)
		//IL_002b: Expected I4, but got Unknown
		string text = AH.A(73830);
		Language applicationLanguage = clsEnvironment.ApplicationLanguage;
		switch ((int)applicationLanguage)
		{
		case 2:
			text = AH.A(73927);
			break;
		case 3:
			text = AH.A(74048);
			break;
		}
		Forms.WarningMessage(text);
	}

	public static void TwoOrMoreShapesRequiredError()
	{
		//IL_000d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0012: Unknown result type (might be due to invalid IL or missing references)
		//IL_0013: Unknown result type (might be due to invalid IL or missing references)
		//IL_0029: Expected I4, but got Unknown
		string text = AH.A(74161);
		Language applicationLanguage = clsEnvironment.ApplicationLanguage;
		switch ((int)applicationLanguage)
		{
		case 2:
			text = AH.A(74266);
			break;
		case 3:
			text = AH.A(74400);
			break;
		}
		Forms.WarningMessage(text);
	}
}
