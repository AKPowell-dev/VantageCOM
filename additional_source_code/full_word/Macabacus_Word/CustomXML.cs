using System;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic.CompilerServices;

namespace Macabacus_Word;

public sealed class CustomXML
{
	public static CustomXMLPart RetrievePart(InlineShape shp)
	{
		return RetrievePart(shp.Range.Document, shp.AlternativeText);
	}

	public static CustomXMLPart RetrievePart(Microsoft.Office.Interop.Word.Shape shp)
	{
		return RetrievePart(shp.Anchor.Document, shp.AlternativeText);
	}

	public static CustomXMLPart RetrievePart(Table shp)
	{
		return RetrievePart(shp.Range.Document, shp.Descr);
	}

	public static CustomXMLPart RetrievePart(ContentControl cc)
	{
		return RetrievePart(cc.Range.Document, cc.Tag);
	}

	public static CustomXMLPart RetrievePart(Document doc, string strId)
	{
		return CustomXML.RetrievePart(doc.CustomXMLParts, strId);
	}

	public static void RemoveCustomXMLPart(object obj)
	{
		Type typeFromHandle = typeof(CustomXML);
		string memberName = XC.A(3113);
		object[] obj2 = new object[1] { obj };
		object[] array = obj2;
		bool[] obj3 = new bool[1] { true };
		bool[] array2 = obj3;
		object obj4 = NewLateBinding.LateGet(null, typeFromHandle, memberName, obj2, null, null, obj3);
		if (array2[0])
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
			obj = RuntimeHelpers.GetObjectValue(array[0]);
		}
		object objectValue = RuntimeHelpers.GetObjectValue(obj4);
		if (objectValue == null)
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
			NewLateBinding.LateCall(objectValue, null, XC.A(3138), new object[0], null, null, null, IgnoreReturn: true);
			objectValue = null;
			return;
		}
	}
}
