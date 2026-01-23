using System;
using MacabacusMacros;
using MacabacusMacros.Links;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic.CompilerServices;

namespace Macabacus_Word.Links;

public sealed class Update
{
	public static void Source(Shape shp, RefreshInstance refreshInstance, string strSource, bool blnUpdateLastModified)
	{
		A(shp.Anchor, refreshInstance, shp.AlternativeText, strSource, blnUpdateLastModified);
	}

	public static void Source(InlineShape shp, RefreshInstance refreshInstance, string strSource, bool blnUpdateLastModified)
	{
		A(shp.Range, refreshInstance, shp.AlternativeText, strSource, blnUpdateLastModified);
	}

	public static void Source(Table shp, RefreshInstance refreshInstance, string strSource, bool blnUpdateLastModified)
	{
		A(shp.Range, refreshInstance, shp.Descr, strSource, blnUpdateLastModified);
	}

	public static void Source(ContentControl cc, RefreshInstance refreshInstance, string strSource, bool blnUpdateLastModified)
	{
		A(cc.Range, refreshInstance, cc.Tag, strSource, blnUpdateLastModified);
	}

	private static void A(Range A, RefreshInstance B, string C, string D, bool E)
	{
		CustomXML.Update(A.Document, C, CustomXML.XML_NODE_SOURCE, CloudStorage.AddPlaceholdersToPath(D));
		if (!E)
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
			string text = "";
			text = ((B == null) ? Updates.GetLastModifiedTime(D) : B.GetLastModifiedTime(D));
			if (text.Length > 0)
			{
				CustomXML.Update(A.Document, C, CustomXML.XML_NODE_SOURCE_LAST_MOD, text);
			}
			return;
		}
	}

	private static void A(Shape A)
	{
		Update.A(A.Anchor, A.AlternativeText);
	}

	private static void A(InlineShape A)
	{
		Update.A(A.Range, A.AlternativeText);
	}

	private static void A(Table A)
	{
		Update.A(A.Range, A.Descr);
	}

	private static void A(Range A, string B)
	{
		CustomXML.Update(A.Document, B, CustomXML.XML_NODE_UPDATED, Base.LastUpdate());
	}

	public static void User(Shape shp)
	{
		B(shp.Anchor, shp.AlternativeText);
	}

	public static void User(InlineShape shp)
	{
		B(shp.Range, shp.AlternativeText);
	}

	public static void User(Table shp)
	{
		B(shp.Range, shp.Descr);
	}

	public static void User(ContentControl cc)
	{
		B(cc.Range, cc.Tag);
	}

	private static void B(Range A, string B)
	{
		try
		{
			CustomXML.Update(A.Document, B, CustomXML.XML_NODE_USER, A.Application.UserName);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			clsReporting.LogException(ex2);
			ProjectData.ClearProjectError();
		}
	}

	private static void A(Shape A, string B)
	{
		Update.A(A.Anchor, A.AlternativeText, B);
	}

	private static void A(InlineShape A, string B)
	{
		Update.A(A.Range, A.AlternativeText, B);
	}

	private static void A(Table A, string B)
	{
		Update.A(A.Range, A.Descr, B);
	}

	private static void A(Range A, string B, string C)
	{
		CustomXML.Update(A.Document, B, CustomXML.XML_NODE_ADDRESS, C);
	}
}
