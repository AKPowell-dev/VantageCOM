using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros.Proofing;
using MacabacusMacros.Proofing.Check;
using Macabacus_Word.Proofing.Errors;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic.CompilerServices;

namespace Macabacus_Word.Proofing.Check;

public sealed class ShapeColors
{
	public static void FillColor(Microsoft.Office.Interop.Word.Shape shp, List<int> listColors, Severity sev)
	{
		//IL_0039: Unknown result type (might be due to invalid IL or missing references)
		//IL_0028: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			if (shp.HasSmartArt == MsoTriState.msoTrue)
			{
				while (true)
				{
					switch (1)
					{
					case 0:
						break;
					default:
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						A(shp, shp.SmartArt, listColors, sev);
						return;
					}
				}
			}
			A(shp, shp.Fill, sev, listColors);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	public static void FillColor(InlineShape shp, List<int> listColors, Severity sev)
	{
		//IL_0039: Unknown result type (might be due to invalid IL or missing references)
		//IL_0028: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			if (shp.HasSmartArt == MsoTriState.msoTrue)
			{
				while (true)
				{
					switch (1)
					{
					case 0:
						break;
					default:
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						A(shp, shp.SmartArt, listColors, sev);
						return;
					}
				}
			}
			A(shp, shp.Fill, sev, listColors);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	private static void A(object A, SmartArt B, List<int> C, Severity D)
	{
		//IL_0051: Unknown result type (might be due to invalid IL or missing references)
		Dictionary<int, List<Microsoft.Office.Core.Shape>> dictionary = (Dictionary<int, List<Microsoft.Office.Core.Shape>>)Color.SmartArtFill(B, C);
		if (dictionary.Count > 0)
		{
			foreach (KeyValuePair<int, List<Microsoft.Office.Core.Shape>> item in dictionary)
			{
				Main.Analysis.Errors.Add(new NonconformingSmartArtFillColor(RuntimeHelpers.GetObjectValue(A), item.Key, item.Value, D));
			}
		}
		dictionary = null;
	}

	private static void A(object A, Microsoft.Office.Interop.Word.FillFormat B, Severity C, List<int> D)
	{
		//IL_005a: Unknown result type (might be due to invalid IL or missing references)
		Microsoft.Office.Interop.Word.FillFormat fillFormat = B;
		if (fillFormat.Visible == MsoTriState.msoTrue)
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
			int rGB = fillFormat.ForeColor.RGB;
			if (Color.ColorNotInPalette(rGB, D))
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
				Main.Analysis.Errors.Add(new NonconformingFillColor(RuntimeHelpers.GetObjectValue(A), rGB, C));
			}
		}
		fillFormat = null;
	}

	public static void BorderColor(Microsoft.Office.Interop.Word.Shape shp, List<int> listColors, Severity sev)
	{
		//IL_0035: Unknown result type (might be due to invalid IL or missing references)
		//IL_0024: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			if (shp.HasSmartArt == MsoTriState.msoTrue)
			{
				while (true)
				{
					switch (4)
					{
					case 0:
						break;
					default:
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						B(shp, shp.SmartArt, listColors, sev);
						return;
					}
				}
			}
			A(shp, shp.Line, sev, listColors);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	public static void BorderColor(InlineShape shp, List<int> listColors, Severity sev)
	{
		//IL_0039: Unknown result type (might be due to invalid IL or missing references)
		//IL_0028: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			if (shp.HasSmartArt == MsoTriState.msoTrue)
			{
				while (true)
				{
					switch (6)
					{
					case 0:
						break;
					default:
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						B(shp, shp.SmartArt, listColors, sev);
						return;
					}
				}
			}
			A(shp, shp.Line, sev, listColors);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	private static void B(object A, SmartArt B, List<int> C, Severity D)
	{
		//IL_0073: Unknown result type (might be due to invalid IL or missing references)
		Dictionary<int, List<Microsoft.Office.Core.Shape>> dictionary = (Dictionary<int, List<Microsoft.Office.Core.Shape>>)Color.SmartArtBorder((SmartArt)NewLateBinding.LateGet(A, null, XC.A(22839), new object[0], null, null, null), C);
		if (dictionary.Count > 0)
		{
			foreach (KeyValuePair<int, List<Microsoft.Office.Core.Shape>> item in dictionary)
			{
				Main.Analysis.Errors.Add(new NonconformingSmartArtBorderColor(RuntimeHelpers.GetObjectValue(A), item.Key, item.Value, D));
			}
		}
		dictionary = null;
	}

	private static void A(object A, Microsoft.Office.Interop.Word.LineFormat B, Severity C, List<int> D)
	{
		//IL_004e: Unknown result type (might be due to invalid IL or missing references)
		Microsoft.Office.Interop.Word.LineFormat lineFormat = B;
		if (lineFormat.Visible == MsoTriState.msoTrue)
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
			int rGB = lineFormat.ForeColor.RGB;
			if (Color.ColorNotInPalette(rGB, D))
			{
				Main.Analysis.Errors.Add(new NonconformingBorderColor(RuntimeHelpers.GetObjectValue(A), rGB, C));
			}
		}
		lineFormat = null;
	}

	public static void FontColor(Microsoft.Office.Interop.Word.Shape shp, List<int> listColors, Severity sev)
	{
		//IL_0015: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			if (shp.HasSmartArt == MsoTriState.msoTrue)
			{
				C(shp, shp.SmartArt, listColors, sev);
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	public static void FontColor(InlineShape shp, List<int> listColors, Severity sev)
	{
		//IL_0028: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			if (shp.HasSmartArt != MsoTriState.msoTrue)
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
				C(shp, shp.SmartArt, listColors, sev);
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

	private static void C(object A, SmartArt B, List<int> C, Severity D)
	{
		//IL_0024: Unknown result type (might be due to invalid IL or missing references)
		//IL_0029: Unknown result type (might be due to invalid IL or missing references)
		//IL_002b: Unknown result type (might be due to invalid IL or missing references)
		//IL_002c: Unknown result type (might be due to invalid IL or missing references)
		//IL_00aa: Unknown result type (might be due to invalid IL or missing references)
		//IL_00ba: Unknown result type (might be due to invalid IL or missing references)
		//IL_00ca: Unknown result type (might be due to invalid IL or missing references)
		//IL_00da: Unknown result type (might be due to invalid IL or missing references)
		//IL_003c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0078: Unknown result type (might be due to invalid IL or missing references)
		SmartArtTextColors val = Color.SmartArtText((SmartArt)NewLateBinding.LateGet(A, null, XC.A(22839), new object[0], null, null, null), C);
		if (val.FontColors.Count > 0)
		{
			IEnumerator<KeyValuePair<int, IList<TextRange2>>> enumerator = default(IEnumerator<KeyValuePair<int, IList<TextRange2>>>);
			try
			{
				enumerator = val.FontColors.GetEnumerator();
				while (enumerator.MoveNext())
				{
					KeyValuePair<int, IList<TextRange2>> current = enumerator.Current;
					Main.Analysis.Errors.Add(new NonconformingSmartArtFontColor(RuntimeHelpers.GetObjectValue(A), current.Key, (List<TextRange2>)current.Value, D));
				}
			}
			finally
			{
				if (enumerator != null)
				{
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
						enumerator.Dispose();
						break;
					}
				}
			}
		}
		_ = val.UnderlineColors.Count;
		_ = 0;
		_ = val.HighlightColors.Count;
		_ = 0;
		_ = val.OutlineColors.Count;
		_ = 0;
		val = default(SmartArtTextColors);
	}

	private static void A(object A, Font2 B, Severity C, List<int> D)
	{
		//IL_0044: Unknown result type (might be due to invalid IL or missing references)
		int rGB = B.Fill.ForeColor.RGB;
		if (!Color.ColorNotInPalette(rGB, D))
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			Main.Analysis.Errors.Add(new NonconformingFontColor(RuntimeHelpers.GetObjectValue(A), rGB, C));
			return;
		}
	}

	public static void FillTransparency(Microsoft.Office.Interop.Word.Shape shp)
	{
		if (shp.Fill.Visible != MsoTriState.msoTrue)
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
			if (shp.Fill.Transparency > 0f)
			{
				Main.Analysis.Errors.Add(new FillTransparency(shp));
			}
			return;
		}
	}

	public static void FillTransparency(InlineShape shp)
	{
		if (shp.Fill.Visible != MsoTriState.msoTrue)
		{
			return;
		}
		while (true)
		{
			switch (4)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			if (!(shp.Fill.Transparency > 0f))
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
				Main.Analysis.Errors.Add(new FillTransparency((Microsoft.Office.Interop.Word.Shape)shp));
				return;
			}
		}
	}
}
