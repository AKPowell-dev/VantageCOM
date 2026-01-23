using System;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Windows.Forms;
using A;
using MacabacusMacros;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Charts;

public sealed class ResizeTo
{
	public static void PowerPointSelection()
	{
		if (!Helpers.A())
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
			Chart chart;
			try
			{
				chart = Helpers.SelectedChart();
				if (chart == null)
				{
					while (true)
					{
						switch (5)
						{
						case 0:
							break;
						default:
							throw new Exception();
						}
					}
				}
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
				return;
			}
			Microsoft.Office.Interop.PowerPoint.Shape shape;
			try
			{
				Microsoft.Office.Interop.PowerPoint.ShapeRange shapeRange = ((Microsoft.Office.Interop.PowerPoint.Application)Interaction.GetObject(null, VH.A(62824))).ActiveWindow.Selection.ShapeRange;
				if (shapeRange.Count != 1)
				{
					while (true)
					{
						switch (4)
						{
						case 0:
							continue;
						}
						throw new Exception();
					}
				}
				shape = shapeRange[1];
				_ = null;
				ChartObject obj = (ChartObject)chart.Parent;
				obj.Width = shape.Width;
				obj.Height = shape.Height;
				_ = null;
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				Forms.WarningMessage(VH.A(62869));
				ProjectData.ClearProjectError();
			}
			shape = null;
			chart = null;
			return;
		}
	}

	public static void WordSelection()
	{
		if (!Helpers.A())
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
			Chart chart;
			try
			{
				chart = Helpers.SelectedChart();
				if (chart == null)
				{
					throw new Exception();
				}
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
				return;
			}
			try
			{
				float[] array = clsPublish.SelectedWordShapeSize((Microsoft.Office.Interop.Word.Application)Interaction.GetObject(null, VH.A(62984)));
				if (!(array[0] > 0f))
				{
					throw new Exception();
				}
				while (true)
				{
					switch (1)
					{
					case 0:
						continue;
					}
					ChartObject obj = (ChartObject)chart.Parent;
					obj.Width = array[0];
					obj.Height = array[1];
					_ = null;
					break;
				}
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				Forms.WarningMessage(VH.A(63017));
				ProjectData.ClearProjectError();
			}
			chart = null;
			return;
		}
	}

	public static void ExcelSelection()
	{
		if (!Helpers.A())
		{
			return;
		}
		Range range = null;
		string text = "";
		try
		{
			text = MH.A.Application.ActiveWindow.RangeSelection.get_Address((object)true, (object)true, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		try
		{
			ChartObject obj = (ChartObject)Helpers.SelectedChart().Parent;
			range = (Range)MH.A.Application.InputBox(VH.A(62623), VH.A(40448), text, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), 8);
			ChartObject chartObject = obj;
			chartObject.Width = Conversions.ToDouble(range.Width);
			chartObject.Height = Conversions.ToDouble(range.Height);
			chartObject.Placement = XlPlacement.xlFreeFloating;
			if (MessageBox.Show(VH.A(63120), VH.A(40448), MessageBoxButtons.YesNo, MessageBoxIcon.Asterisk) == DialogResult.Yes)
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
				chartObject.Top = Conversions.ToDouble(range.Top);
				chartObject.Left = Conversions.ToDouble(range.Left);
			}
			chartObject = null;
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			ProjectData.ClearProjectError();
		}
		range = null;
	}

	public static void StandardSize(int i)
	{
		//IL_0059: Unknown result type (might be due to invalid IL or missing references)
		//IL_005e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0060: Unknown result type (might be due to invalid IL or missing references)
		//IL_0070: Unknown result type (might be due to invalid IL or missing references)
		//IL_0084: Unknown result type (might be due to invalid IL or missing references)
		if (!Helpers.A())
		{
			return;
		}
		while (true)
		{
			switch (5)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
			if (application.ActiveSheet is Chart)
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
				Helpers.A();
			}
			else
			{
				try
				{
					Chart chart = Helpers.SelectedChart();
					if (chart != null)
					{
						StandardSize standardSize = clsPublish.GetStandardSize(i);
						ChartObject obj = (ChartObject)chart.Parent;
						obj.Height = application.InchesToPoints(standardSize.Height);
						obj.Width = application.InchesToPoints(standardSize.Width);
						_ = null;
						chart = null;
					}
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					ProjectData.ClearProjectError();
				}
			}
			application = null;
			return;
		}
	}
}
