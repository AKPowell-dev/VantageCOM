using System;
using System.Collections;
using System.IO;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Windows.Forms;
using A;
using MacabacusMacros;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Charts;

public sealed class SaveAsImage
{
	public static void Initiate()
	{
		if (!Helpers.A())
		{
			return;
		}
		Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
		string B = string.Empty;
		int D = 0;
		int F = 0;
		bool flag = false;
		try
		{
			Chart chart;
			if (application.ActiveSheet is Worksheet)
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
				chart = application.ActiveChart;
			}
			else
			{
				chart = (Chart)application.ActiveSheet;
			}
			string C = application.ActiveWorkbook.Path;
			string D2 = default(string);
			if (chart != null)
			{
				B = Path.Combine(C, chart.Name + VH.A(63217));
				B = A(application, B);
				if (Operators.CompareString(B, VH.A(63226), TextCompare: false) != 0)
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
					if (!clsFile.IsPathUrl(B))
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
						if (File.Exists(B))
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
							if (MessageBox.Show(VH.A(63237), VH.A(40448), MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation) == DialogResult.Cancel)
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
								flag = true;
							}
						}
					}
					if (!flag)
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
						A(chart, B, Path.GetExtension(B), ref D);
					}
				}
				chart = null;
			}
			else if (Operators.CompareString(Versioned.TypeName(RuntimeHelpers.GetObjectValue(application.Selection)), VH.A(56245), TextCompare: false) == 0)
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
				ShapeRange shapeRange = (ShapeRange)NewLateBinding.LateGet(application.Selection, null, VH.A(56274), new object[0], null, null, null);
				IEnumerator enumerator = default(IEnumerator);
				try
				{
					enumerator = shapeRange.GetEnumerator();
					Shape shape;
					do
					{
						if (enumerator.MoveNext())
						{
							shape = (Shape)enumerator.Current;
							continue;
						}
						while (true)
						{
							switch (5)
							{
							case 0:
								break;
							default:
								goto end_IL_01f7;
							}
							continue;
							end_IL_01f7:
							break;
						}
						break;
					}
					while (shape.HasChart != MsoTriState.msoTrue || A(shape.Chart, ref B, ref C, ref D2, ref D, ref F));
				}
				finally
				{
					if (enumerator is IDisposable)
					{
						while (true)
						{
							switch (4)
							{
							case 0:
								continue;
							}
							(enumerator as IDisposable).Dispose();
							break;
						}
					}
				}
				shapeRange = null;
			}
			else
			{
				ChartObjects chartObjects = (ChartObjects)((Worksheet)application.ActiveSheet).ChartObjects(RuntimeHelpers.GetObjectValue(Missing.Value));
				if (chartObjects.Count > 1)
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
					IEnumerator enumerator2 = default(IEnumerator);
					try
					{
						enumerator2 = chartObjects.GetEnumerator();
						while (enumerator2.MoveNext() && A(((ChartObject)enumerator2.Current).Chart, ref B, ref C, ref D2, ref D, ref F))
						{
						}
					}
					finally
					{
						if (enumerator2 is IDisposable)
						{
							while (true)
							{
								switch (2)
								{
								case 0:
									continue;
								}
								(enumerator2 as IDisposable).Dispose();
								break;
							}
						}
					}
				}
				else
				{
					Forms.WarningMessage(VH.A(63373));
				}
				chartObjects = null;
			}
			if (D > 0)
			{
				while (true)
				{
					switch (6)
					{
					case 0:
						continue;
					}
					clsFile.OpenExplorerToFile(B);
					clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)5, VH.A(63442));
					break;
				}
			}
		}
		catch (ArgumentException ex)
		{
			ProjectData.SetProjectError(ex);
			ArgumentException ex2 = ex;
			Forms.ErrorMessage(ex2.Message);
			ProjectData.ClearProjectError();
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			Forms.ErrorMessage(ex4.Message);
			clsReporting.LogException(ex4);
			ProjectData.ClearProjectError();
		}
		application = null;
	}

	private static bool A(Chart A, ref string B, ref string C, ref string D, ref int E, ref int F)
	{
		if (Operators.CompareString(B, string.Empty, TextCompare: false) == 0)
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
			B = Path.Combine(C, VH.A(63487));
			B = SaveAsImage.A(A.Application, B);
			if (Operators.CompareString(B, VH.A(63226), TextCompare: false) == 0)
			{
				while (true)
				{
					switch (4)
					{
					case 0:
						continue;
					}
					return false;
				}
			}
			C = Path.GetDirectoryName(B);
			D = Path.GetFileNameWithoutExtension(B);
		}
		string extension = Path.GetExtension(B);
		checked
		{
			if (!clsFile.IsPathUrl(B))
			{
				while (true)
				{
					F++;
					B = Path.Combine(C, D + VH.A(63506) + F.ToString(VH.A(63509)) + extension);
					if (F > 999)
					{
						break;
					}
					if (File.Exists(B))
					{
						continue;
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
					break;
				}
			}
			else
			{
				F++;
				B = Path.Combine(C, D + VH.A(63506) + F.ToString(VH.A(63509)) + extension);
			}
			SaveAsImage.A(A, B, extension, ref E);
			return true;
		}
	}

	private static void A(Chart A, string B, string C, ref int D)
	{
		if (Operators.CompareString(C, VH.A(63217), TextCompare: false) == 0)
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
			A.Export(B, VH.A(63516), RuntimeHelpers.GetObjectValue(Missing.Value));
		}
		else
		{
			try
			{
				A.Export(B, VH.A(63523), RuntimeHelpers.GetObjectValue(Missing.Value));
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				Forms.ErrorMessage(VH.A(63530));
				clsReporting.LogException(ex2);
				throw;
			}
		}
		checked
		{
			D++;
		}
	}

	private static string A(Microsoft.Office.Interop.Excel.Application A, string B)
	{
		return Conversions.ToString(A.GetSaveAsFilename(B, VH.A(63623), RuntimeHelpers.GetObjectValue(Missing.Value), VH.A(63724), RuntimeHelpers.GetObjectValue(Missing.Value)));
	}
}
