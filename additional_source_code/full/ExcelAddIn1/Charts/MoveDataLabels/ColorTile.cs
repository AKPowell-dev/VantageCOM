using System;
using System.Drawing;
using System.Windows.Media;
using MacabacusMacros.Config;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Charts.MoveDataLabels;

public sealed class ColorTile
{
	internal static System.Windows.Media.Brush A(Microsoft.Office.Interop.Excel.FillFormat A)
	{
		System.Windows.Media.Brush result;
		try
		{
			object obj;
			if (A.Visible != MsoTriState.msoTrue)
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
				obj = ColorTile.A();
			}
			else
			{
				obj = ColorTile.A(A.ForeColor.RGB);
			}
			result = (System.Windows.Media.Brush)obj;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			result = ColorTile.A();
			ProjectData.ClearProjectError();
		}
		return result;
	}

	internal static System.Windows.Media.Brush A(LineFormat A)
	{
		System.Windows.Media.Brush result;
		try
		{
			object obj;
			if (A.Visible != MsoTriState.msoTrue)
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
				obj = ColorTile.A();
			}
			else
			{
				obj = ColorTile.A(A.ForeColor.RGB);
			}
			result = (System.Windows.Media.Brush)obj;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			result = ColorTile.A();
			ProjectData.ClearProjectError();
		}
		return result;
	}

	internal static System.Windows.Media.Brush A(int A)
	{
		System.Drawing.Color color = ColorTranslator.FromOle(A);
		return new SolidColorBrush(System.Windows.Media.Color.FromRgb(color.R, color.G, color.B));
	}

	internal static ImageBrush A()
	{
		return Base.Checkerboard(false);
	}
}
