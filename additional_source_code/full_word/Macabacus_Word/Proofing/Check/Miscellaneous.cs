using System;
using System.IO;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros;
using Macabacus_Word.Links;
using Macabacus_Word.Proofing.Errors;
using Microsoft.VisualBasic.CompilerServices;

namespace Macabacus_Word.Proofing.Check;

public sealed class Miscellaneous
{
	public static void CheckMacabacusLink(object obj)
	{
		Type typeFromHandle = typeof(Common);
		string memberName = XC.A(13018);
		object[] obj2 = new object[1] { obj };
		object[] array = obj2;
		bool[] obj3 = new bool[1] { true };
		bool[] array2 = obj3;
		object value = NewLateBinding.LateGet(null, typeFromHandle, memberName, obj2, null, null, obj3);
		if (array2[0])
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
			obj = RuntimeHelpers.GetObjectValue(array[0]);
		}
		if (!Conversions.ToBoolean(value))
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
			Type typeFromHandle2 = typeof(Common);
			string memberName2 = XC.A(11777);
			object[] obj4 = new object[1] { obj };
			array = obj4;
			bool[] obj5 = new bool[1] { true };
			array2 = obj5;
			object instance = NewLateBinding.LateGet(null, typeFromHandle2, memberName2, obj4, null, null, obj5);
			if (array2[0])
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
				obj = RuntimeHelpers.GetObjectValue(array[0]);
			}
			string text = Conversions.ToString(NewLateBinding.LateGet(instance, null, XC.A(13872), new object[0], null, null, null));
			if (clsFile.NewerVersions(text).Count > 0)
			{
				while (true)
				{
					switch (4)
					{
					case 0:
						continue;
					}
					break;
				}
				Main.Analysis.Errors.Add(new LinkNewerVersionAvailable(RuntimeHelpers.GetObjectValue(obj), Path.GetFileName(text)));
			}
			if (clsFile.IsPathUrl(text) || File.Exists(text))
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
				Main.Analysis.Errors.Add(new LinkBroken(RuntimeHelpers.GetObjectValue(obj), Path.GetFileName(text)));
				return;
			}
		}
	}
}
