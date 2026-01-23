using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.DeckCheck.Check;

public sealed class FootnoteSequence
{
	public void Check(Slide sld, Shape shp, List<int> listFootnoteNumbers)
	{
		if (listFootnoteNumbers.Count <= 0)
		{
			return;
		}
		bool flag = false;
		listFootnoteNumbers = listFootnoteNumbers.Distinct().ToList();
		listFootnoteNumbers = listFootnoteNumbers.OrderBy([SpecialName] (int A) => A).ToList();
		int num = 1;
		using List<int>.Enumerator enumerator = listFootnoteNumbers.GetEnumerator();
		while (enumerator.MoveNext())
		{
			if (enumerator.Current != num)
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
						flag = true;
						return;
					}
				}
			}
			num = checked(num + 1);
		}
		while (true)
		{
			switch (5)
			{
			case 0:
				break;
			default:
				return;
			}
		}
	}
}
