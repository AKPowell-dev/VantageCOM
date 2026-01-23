using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows.Forms;
using System.Xml;
using A;
using MacabacusMacros;
using MacabacusMacros.Config;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Format;

public sealed class Styles
{
	[CompilerGenerated]
	private static int m_A;

	[CompilerGenerated]
	private static int B;

	internal static int CycleNumber
	{
		[CompilerGenerated]
		get
		{
			return Styles.m_A;
		}
		[CompilerGenerated]
		set
		{
			Styles.m_A = value;
		}
	}

	internal static int CycleIndex
	{
		[CompilerGenerated]
		get
		{
			return B;
		}
		[CompilerGenerated]
		set
		{
			B = value;
		}
	}

	public static void CycleCustom1()
	{
		CustomCycle(1);
	}

	public static void CycleCustom2()
	{
		CustomCycle(2);
	}

	public static void CycleCustom3()
	{
		CustomCycle(3);
	}

	public static void CycleCustom4()
	{
		CustomCycle(4);
	}

	public static void CycleCustom5()
	{
		CustomCycle(5);
	}

	public static void CycleCustom6()
	{
		CustomCycle(6);
	}

	public static void CycleCustom7()
	{
		CustomCycle(7);
	}

	public static void CycleCustom8()
	{
		CustomCycle(8);
	}

	public static void CustomCycle(int i)
	{
		if (!Licensing.AllowRestrictedMode())
		{
			return;
		}
		checked
		{
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
				try
				{
					List<XmlNode> value = KH.A.CustomCycles.ElementAt(i - 1).Value;
					if (!value.Any())
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
						if (i == CycleNumber)
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
							if (CycleIndex == value.Count - 1)
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
								CycleIndex = 0;
							}
							else
							{
								CycleIndex++;
							}
						}
						else
						{
							CycleIndex = 0;
						}
						ApplyStyle(value[CycleIndex]);
						CycleNumber = i;
						if (CycleIndex != 0)
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
							Base.LogActivity(VH.A(151380));
							return;
						}
					}
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					Base.LogException(ex2);
					ProjectData.ClearProjectError();
					return;
				}
				finally
				{
					List<XmlNode> value = null;
				}
			}
		}
	}

	public static void DoStyle(IRibbonControl ctrl)
	{
		if (Licensing.AllowRestrictedMode())
		{
			string[] array = ctrl.Tag.Split(',');
			ApplyStyle(KH.A.CustomCycles.ElementAt(Conversions.ToInteger(array[0])).Value[Conversions.ToInteger(array[1])]);
		}
	}

	public static void ApplyStyle(XmlNode ndStyle)
	{
		Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
		try
		{
			if (application.Selection is Range)
			{
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
					Range range = JH.A((Range)null);
					if (!Base.IsWorksheetProtected(range.Worksheet))
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
						application.ScreenUpdating = false;
						try
						{
							bool num = JH.A(range);
							ApplyStyle(range, ndStyle);
							if (num)
							{
								JH.A(range, VH.A(148068));
							}
						}
						catch (Exception ex)
						{
							ProjectData.SetProjectError(ex);
							Exception ex2 = ex;
							Base.HandleFormattingException(ex2);
							ProjectData.ClearProjectError();
						}
						application.ScreenUpdating = true;
					}
					range = null;
					break;
				}
			}
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			Base.LogException(ex4);
			ProjectData.ClearProjectError();
		}
		application = null;
	}

	public static void ApplyStyle(Range oRng, XmlNode xmlStyle)
	{
		Styles.ApplyStyle(oRng, xmlStyle);
	}

	private Bitmap A(Range A, int B)
	{
		Bitmap result = null;
		Microsoft.Office.Interop.Excel.Application application = A.Application;
		application.ScreenUpdating = true;
		try
		{
			((_Application)application).get_Range((object)checked(A.get_Offset((object)(-B), (object)(-B))), (object)A.get_Offset((object)B, (object)B)).CopyPicture(XlPictureAppearance.xlScreen, XlCopyPictureFormat.xlBitmap);
			if (Clipboard.ContainsImage())
			{
				while (true)
				{
					switch (2)
					{
					case 0:
						continue;
					}
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					result = new Bitmap(Clipboard.GetImage());
					break;
				}
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			clsReporting.LogException(ex2);
			ProjectData.ClearProjectError();
		}
		application.ScreenUpdating = false;
		application = null;
		return result;
	}
}
