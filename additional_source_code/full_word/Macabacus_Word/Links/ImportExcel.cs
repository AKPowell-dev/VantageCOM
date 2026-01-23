using System;
using A;
using MacabacusMacros.Auth;
using MacabacusMacros.ImportExport;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace Macabacus_Word.Links;

public sealed class ImportExcel
{
	internal static void A()
	{
		//IL_000d: Unknown result type (might be due to invalid IL or missing references)
		if (Access.AllowWordOperation((PlanType)5, (Restriction)1, false))
		{
			try
			{
				ExcelToWord.DefaultImportExport(A(), (Application)null, PC.A.Application);
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
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
		//IL_000c: Unknown result type (might be due to invalid IL or missing references)
		//IL_000d: Unknown result type (might be due to invalid IL or missing references)
		//IL_000e: Unknown result type (might be due to invalid IL or missing references)
		if (Access.AllowWordOperation((PlanType)5, (Restriction)1, false))
		{
			ExcelToWord.ImportButton(A, A, ImportExcel.A(), (Application)null, PC.A.Application);
		}
	}

	private static MatchSize A()
	{
		//IL_001c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0021: Unknown result type (might be due to invalid IL or missing references)
		return clsImportExport.GetMatchSize(N.Settings.ImportMatchDestinationWidth, N.Settings.ImportMatchDestinationHeight);
	}
}
