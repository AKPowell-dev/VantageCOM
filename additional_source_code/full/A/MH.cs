using System;
using System.CodeDom.Compiler;
using System.Diagnostics;
using ExcelAddIn1;
using Microsoft.Office.Tools.Excel;

namespace A;

[GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "15.0.0.0")]
[DebuggerNonUserCode]
internal sealed class MH
{
	private static ThisAddIn m_A;

	private static ApplicationFactory m_A;

	private static NH m_A;

	internal static ThisAddIn A
	{
		get
		{
			return MH.m_A;
		}
		set
		{
			if (MH.m_A == null)
			{
				while (true)
				{
					switch (7)
					{
					case 0:
						break;
					default:
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						MH.m_A = value;
						return;
					}
				}
			}
			throw new NotSupportedException();
		}
	}

	internal static ApplicationFactory A
	{
		get
		{
			return MH.m_A;
		}
		set
		{
			if (MH.m_A == null)
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
						MH.m_A = value;
						return;
					}
				}
			}
			throw new NotSupportedException();
		}
	}

	internal static NH A
	{
		get
		{
			if (MH.m_A == null)
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
				MH.m_A = new NH(MH.m_A.GetRibbonFactory());
			}
			return MH.m_A;
		}
	}

	private MH()
	{
	}
}
