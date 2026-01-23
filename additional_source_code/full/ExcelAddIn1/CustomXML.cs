using MacabacusMacros;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1;

public sealed class CustomXML
{
	internal static CustomXMLPart A(Shape A)
	{
		return CustomXML.A((Microsoft.Office.Interop.Excel.Workbook)A.TopLeftCell.Worksheet.Parent, A.AlternativeText);
	}

	internal static CustomXMLPart A(Microsoft.Office.Interop.Excel.Workbook A, string B)
	{
		return CustomXML.RetrievePart(A.CustomXMLParts, B);
	}

	internal static void A(Shape A)
	{
		CustomXMLPart customXMLPart = CustomXML.A(A);
		if (customXMLPart == null)
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
			customXMLPart.Delete();
			customXMLPart = null;
			return;
		}
	}
}
