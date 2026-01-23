using System;
using System.Collections;
using System.IO;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Xml;
using A;
using MacabacusMacros;
using MacabacusMacros.UI;
using Macabacus_Word.Keyboard;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace Macabacus_Word;

public sealed class clsUtilities
{
	public static readonly string DOTM_FILE_NAME = XC.A(42506);

	public static readonly string ADDIN_NAME = XC.A(2438);

	public static readonly int SVG_WD_INLINE_SHAPE_TYPE = 17;

	public static readonly int SVG_WD_MSO_SHAPE_TYPE = 28;

	public static void StartupProcedures(XmlDocument xmlSettings)
	{
		NC.A = new clsSettings(xmlSettings);
		Shortcuts.BuildDictionary();
		Shortcuts.Load();
	}

	public static string MacabacusDotmPath()
	{
		return Path.Combine(clsEnvironment.CommonAppDataPath, DOTM_FILE_NAME);
	}

	public static void InitializeMacabacus()
	{
		Application application = PC.A.Application;
		AddIn addIn = null;
		string text = MacabacusDotmPath();
		Document document = null;
		bool flag = false;
		try
		{
			AddIns addIns = application.AddIns;
			object Index = ADDIN_NAME;
			addIn = addIns[ref Index];
			if (Operators.CompareString(addIn.Path.ToLower(), text.ToLower(), TextCompare: false) != 0)
			{
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
					Forms.ErrorMessage(XC.A(42205) + DOTM_FILE_NAME + XC.A(42214) + text);
					break;
				}
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			if (Conversion.Val(application.Version) < 15.0)
			{
				while (true)
				{
					switch (2)
					{
					case 0:
						continue;
					}
					break;
				}
				IEnumerator enumerator = default(IEnumerator);
				try
				{
					enumerator = application.Documents.GetEnumerator();
					while (true)
					{
						if (enumerator.MoveNext())
						{
							document = (Document)enumerator.Current;
							Microsoft.Office.Interop.Word.Windows windows = document.Windows;
							object Index = 1;
							if (!windows[ref Index].Visible)
							{
								continue;
							}
							while (true)
							{
								switch (2)
								{
								case 0:
									continue;
								}
								flag = true;
								break;
							}
							break;
						}
						while (true)
						{
							switch (1)
							{
							case 0:
								break;
							default:
								goto end_IL_0123;
							}
							continue;
							end_IL_0123:
							break;
						}
						break;
					}
				}
				finally
				{
					if (enumerator is IDisposable)
					{
						while (true)
						{
							switch (3)
							{
							case 0:
								continue;
							}
							(enumerator as IDisposable).Dispose();
							break;
						}
					}
				}
				if (!flag)
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
					Documents documents = application.Documents;
					object Index = RuntimeHelpers.GetObjectValue(Missing.Value);
					object NewTemplate = RuntimeHelpers.GetObjectValue(Missing.Value);
					object DocumentType = RuntimeHelpers.GetObjectValue(Missing.Value);
					object Visible = RuntimeHelpers.GetObjectValue(Missing.Value);
					document = documents.Add(ref Index, ref NewTemplate, ref DocumentType, ref Visible);
				}
				else
				{
					document = null;
				}
			}
			try
			{
				AddIns addIns2 = application.AddIns;
				object Visible = RuntimeHelpers.GetObjectValue(Missing.Value);
				addIn = addIns2.Add(text, ref Visible);
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				ProjectData.ClearProjectError();
			}
			if (document != null)
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
				Document document2 = document;
				object Visible = false;
				object DocumentType = RuntimeHelpers.GetObjectValue(Missing.Value);
				object NewTemplate = RuntimeHelpers.GetObjectValue(Missing.Value);
				document2.Close(ref Visible, ref DocumentType, ref NewTemplate);
				document = null;
			}
			ProjectData.ClearProjectError();
		}
		try
		{
			if (addIn != null)
			{
				while (true)
				{
					switch (6)
					{
					case 0:
						continue;
					}
					if (addIn.Installed)
					{
						break;
					}
					while (true)
					{
						switch (2)
						{
						case 0:
							continue;
						}
						addIn.Installed = true;
						break;
					}
					break;
				}
			}
		}
		catch (Exception ex5)
		{
			ProjectData.SetProjectError(ex5);
			Exception ex6 = ex5;
			ProjectData.ClearProjectError();
		}
		addIn = null;
		application = null;
	}

	public static bool IsSelectionTextInsideShape(Selection sel)
	{
		if (sel.Type == WdSelectionType.wdSelectionNormal)
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			if (sel.Range.StoryType == WdStoryType.wdTextFrameStory)
			{
				while (true)
				{
					switch (1)
					{
					case 0:
						break;
					default:
						return true;
					}
				}
			}
		}
		return false;
	}

	public static bool IsSelectionTextInsideTable(Selection sel)
	{
		int num;
		if (Conversions.ToBoolean(sel.Range.get_Information(WdInformation.wdWithInTable)))
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
			num = ((!IsSelectionTypeOfShape(sel)) ? 1 : 0);
		}
		else
		{
			num = 0;
		}
		if (Conversions.ToBoolean((byte)num != 0))
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					return true;
				}
			}
		}
		return false;
	}

	public static bool IsSelectionTypeOfShape(Selection sel)
	{
		if (!IsSelectionTypeFloatingShape(sel))
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
			if (!IsSelectionTypeNonFloatingShape(sel))
			{
				return false;
			}
			while (true)
			{
				switch (3)
				{
				case 0:
					continue;
				}
				break;
			}
		}
		return true;
	}

	public static bool IsSelectionTypeNonFloatingShape(Selection sel)
	{
		if (sel.Type == WdSelectionType.wdSelectionInlineShape)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					return true;
				}
			}
		}
		return false;
	}

	public static bool IsSelectionTypeFloatingShape(Selection sel)
	{
		if (sel.Type == WdSelectionType.wdSelectionShape)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					return true;
				}
			}
		}
		return false;
	}

	public static bool IsRangeInHeaderFooter(Range rng)
	{
		if (rng.StoryType != WdStoryType.wdPrimaryHeaderStory)
		{
			while (true)
			{
				switch (2)
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
			if (rng.StoryType != WdStoryType.wdPrimaryFooterStory && rng.StoryType != WdStoryType.wdFirstPageHeaderStory)
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
				if (rng.StoryType != WdStoryType.wdFirstPageFooterStory)
				{
					return false;
				}
				while (true)
				{
					switch (5)
					{
					case 0:
						continue;
					}
					break;
				}
			}
		}
		return true;
	}

	public static bool DoesAutoShapeHaveText(Microsoft.Office.Interop.Word.Shape shp)
	{
		bool result;
		try
		{
			if (shp.AutoShapeType == MsoAutoShapeType.msoShapeMixed || shp.TextFrame == null)
			{
				goto IL_003c;
			}
			while (true)
			{
				switch (1)
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
			if (shp.TextFrame.HasText == 0)
			{
				goto IL_003c;
			}
			result = true;
			goto end_IL_0000;
			IL_003c:
			result = false;
			end_IL_0000:;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			result = false;
			ProjectData.ClearProjectError();
		}
		return result;
	}
}
