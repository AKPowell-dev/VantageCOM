using System;
using System.Collections;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using A;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Comments;

public sealed class Author
{
	public static void Change()
	{
		Application application = MH.A.Application;
		application.ScreenUpdating = false;
		try
		{
			string userName = MH.A.Application.UserName;
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = application.ActiveWorkbook.Worksheets.GetEnumerator();
				IEnumerator enumerator2 = default(IEnumerator);
				while (enumerator.MoveNext())
				{
					Worksheet worksheet = (Worksheet)enumerator.Current;
					try
					{
						enumerator2 = worksheet.Comments.GetEnumerator();
						while (enumerator2.MoveNext())
						{
							Comment obj = (Comment)enumerator2.Current;
							obj.Text(Regex.Replace(obj.Text(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)), VH.A(141720), userName + VH.A(2826)), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
							obj.Shape.TextFrame.Characters(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)).Font.Bold = false;
							obj.Shape.TextFrame.Characters(1, checked(Strings.Len(userName) + 1)).Font.Bold = true;
							_ = null;
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
							break;
						}
					}
					finally
					{
						if (enumerator2 is IDisposable)
						{
							while (true)
							{
								switch (5)
								{
								case 0:
									continue;
								}
								(enumerator2 as IDisposable).Dispose();
								break;
							}
						}
					}
				}
			}
			finally
			{
				if (enumerator is IDisposable)
				{
					while (true)
					{
						switch (6)
						{
						case 0:
							continue;
						}
						(enumerator as IDisposable).Dispose();
						break;
					}
				}
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		application.ScreenUpdating = true;
		application = null;
	}

	public static void Remove()
	{
		Application application = MH.A.Application;
		application.ScreenUpdating = false;
		Range range;
		try
		{
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = application.ActiveWorkbook.Worksheets.GetEnumerator();
				IEnumerator enumerator2 = default(IEnumerator);
				while (enumerator.MoveNext())
				{
					Worksheet worksheet = (Worksheet)enumerator.Current;
					try
					{
						enumerator2 = worksheet.Comments.GetEnumerator();
						while (enumerator2.MoveNext())
						{
							Comment obj = (Comment)enumerator2.Current;
							string text = RemoveFromText(obj);
							range = (Range)obj.Parent;
							obj.Delete();
							range.AddComment(text).Shape.TextFrame.Characters(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)).Font.Bold = false;
							_ = null;
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
							break;
						}
					}
					finally
					{
						if (enumerator2 is IDisposable)
						{
							while (true)
							{
								switch (4)
								{
								case 0:
									continue;
								}
								(enumerator2 as IDisposable).Dispose();
								break;
							}
						}
					}
				}
			}
			finally
			{
				if (enumerator is IDisposable)
				{
					while (true)
					{
						switch (7)
						{
						case 0:
							continue;
						}
						(enumerator as IDisposable).Dispose();
						break;
					}
				}
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		application.ScreenUpdating = true;
		application = null;
		range = null;
	}

	public static string RemoveFromText(Comment cmt)
	{
		string text = Regex.Replace(cmt.Text(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)), VH.A(141743), "");
		if (Operators.CompareString(text, "", TextCompare: false) == 0)
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
			text = VH.A(41385);
		}
		return text;
	}
}
