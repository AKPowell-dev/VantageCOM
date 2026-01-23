using System;
using A;
using MacabacusMacros.Auth;
using MacabacusMacros.ImportExport;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.Links;

public sealed class ImportExcel
{
	internal static void A()
	{
		//IL_0020: Unknown result type (might be due to invalid IL or missing references)
		if (!Access.AllowPowerPointOperation((PlanType)5, (Restriction)1, false))
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
			try
			{
				ExcelToPowerPoint.DefaultImportExport(A(), (Application)null, NG.A.Application);
				return;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
				return;
			}
		}
	}

	internal static void B()
	{
		A((ExportObjectType)0);
	}

	internal static void C()
	{
		A((ExportObjectType)1);
	}

	internal static void D()
	{
		A((ExportObjectType)3);
	}

	internal static void E()
	{
		A((ExportObjectType)2);
	}

	internal static void F()
	{
		A((ExportObjectType)5);
	}

	internal static void G()
	{
		A((ExportObjectType)4);
	}

	private static void A(ExportObjectType A)
	{
		//IL_001f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0020: Unknown result type (might be due to invalid IL or missing references)
		//IL_0021: Unknown result type (might be due to invalid IL or missing references)
		//IL_0026: Unknown result type (might be due to invalid IL or missing references)
		if (!Access.AllowPowerPointOperation((PlanType)5, (Restriction)1, false))
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
			ExcelToPowerPoint.ImportButton(A, A, ImportExcel.A(), (Application)null, NG.A.Application);
			return;
		}
	}

	private static MatchSize A()
	{
		//IL_001c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0021: Unknown result type (might be due to invalid IL or missing references)
		return clsImportExport.GetMatchSize(PB.Settings.ImportMatchDestinationWidth, PB.Settings.ImportMatchDestinationHeight);
	}
}
