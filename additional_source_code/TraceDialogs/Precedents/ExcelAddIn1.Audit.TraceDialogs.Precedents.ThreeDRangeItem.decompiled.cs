using System;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using A;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Audit.TraceDialogs.Precedents;

public sealed class ThreeDRangeItem : BaseItem
{
	[CompilerGenerated]
	private List<Range> A;

	public List<Range> Ranges
	{
		[CompilerGenerated]
		get
		{
			return A;
		}
		[CompilerGenerated]
		set
		{
			A = value;
		}
	}

	public ThreeDRangeItem(BaseItem parent, string strLabel, Microsoft.Office.Interop.Excel.Workbook wb)
	{
		checked
		{
			base..ctor(parent, null, parent.Level + 1, VH.A(45981));
			base.Label = strLabel;
			base.Value = VH.A(41885);
			Ranges = new List<Range>();
			Match match = Regex.Match(strLabel, VH.A(46303) + Base.CELL_REF_PATTERN + VH.A(41262));
			Worksheet worksheet;
			Microsoft.Office.Interop.Excel.Sheets sheets;
			if (match.Success)
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
				try
				{
					sheets = wb.Sheets;
					worksheet = (Worksheet)sheets[match.Groups[1].Value.Replace(VH.A(39854), VH.A(39851))];
					Worksheet obj = (Worksheet)sheets[match.Groups[2].Value.Replace(VH.A(39854), VH.A(39851))];
					int index = worksheet.Index;
					int index2 = obj.Index;
					for (int i = index; i <= index2; i++)
					{
						if (sheets[i] is Worksheet)
						{
							Ranges.Add(((_Worksheet)(Worksheet)sheets[i]).get_Range((object)match.Groups[3].Value, RuntimeHelpers.GetObjectValue(Missing.Value)));
						}
					}
					while (true)
					{
						switch (1)
						{
						case 0:
							break;
						default:
							goto end_IL_0191;
						}
						continue;
						end_IL_0191:
						break;
					}
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					ProjectData.ClearProjectError();
				}
			}
			match = null;
			worksheet = null;
			sheets = null;
		}
	}
}
