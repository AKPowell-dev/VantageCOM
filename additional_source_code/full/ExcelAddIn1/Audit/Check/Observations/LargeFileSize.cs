using A;

namespace ExcelAddIn1.Audit.Check.Observations;

public sealed class LargeFileSize : Observation
{
	public LargeFileSize(Severity sev, long lngBtyes)
		: base(Category.Workbook, sev, VH.A(19282))
	{
		string subtitle;
		if (lngBtyes < 1000000)
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
			subtitle = ((double)lngBtyes / 1000.0).ToString(VH.A(19313)) + VH.A(19328);
		}
		else
		{
			subtitle = ((double)lngBtyes / 1000000.0).ToString(VH.A(19313)) + VH.A(19333);
		}
		base.Subtitle = subtitle;
		base.Explanation = VH.A(19338);
	}
}
