using System;
using System.Collections;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text;
using System.Xml;
using A;
using MacabacusMacros;
using Macabacus_Word.Colors;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace Macabacus_Word;

public sealed class clsStyles
{
	private struct KC
	{
		public string A;

		public string B;

		public string C;

		public string D;

		public string E;

		public string F;

		public string G;

		public string H;

		public int A;

		public WdListType A;

		public int B;

		public WdListNumberStyle A;

		public ListTemplate A;
	}

	public enum WordStyles
	{
		Headings = 1,
		Lists,
		Text,
		Tables
	}

	private static readonly string m_A = XC.A(20083);

	public static void Hello()
	{
		Interaction.MsgBox(XC.A(40987));
	}

	public static void ApplyHeadingStyle(string strId)
	{
		Application application = PC.A.Application;
		UndoRecord undoRecord = application.UndoRecord;
		application.ScreenUpdating = false;
		undoRecord.StartCustomRecord(XC.A(40998));
		KC b;
		Selection selection;
		try
		{
			selection = application.Selection;
			b = A(NC.A.SettingsXml.SelectSingleNode(XC.A(41037) + strId + XC.A(7149)));
			WdSelectionType type = selection.Type;
			if ((uint)(type - 1) <= 1u)
			{
				while (true)
				{
					switch (7)
					{
					case 0:
						continue;
					}
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					A(selection, b);
					break;
				}
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		undoRecord.EndCustomRecord();
		application.ScreenUpdating = true;
		b = default(KC);
		selection = null;
		undoRecord = null;
		application = null;
	}

	public static void ApplyListStyle(string strId)
	{
		Application application = PC.A.Application;
		UndoRecord undoRecord = application.UndoRecord;
		ListTemplate listTemplate = null;
		application.ScreenUpdating = false;
		undoRecord.StartCustomRecord(XC.A(41100));
		Selection selection;
		XmlNode xmlNode;
		try
		{
			selection = application.Selection;
			string[] array = strId.Split('|');
			string text = array[0];
			string text2 = array[1];
			xmlNode = NC.A.SettingsXml.SelectSingleNode(XC.A(41133) + text + XC.A(7149));
			string text3 = A(xmlNode, XC.A(3725));
			float num = Conversions.ToSingle(A(xmlNode, XC.A(41190)));
			float inches = Conversions.ToSingle(A(xmlNode, XC.A(41219)));
			float inches2 = Conversions.ToSingle(A(xmlNode, XC.A(41242)));
			B(xmlNode);
			WdSelectionType type = selection.Type;
			if ((uint)(type - 1) <= 1u)
			{
				IEnumerator enumerator = default(IEnumerator);
				IEnumerator enumerator2 = default(IEnumerator);
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
					Interaction.MsgBox(((Template)application.ActiveDocument.AttachedTemplate).FullName);
					Template template = (Template)application.ActiveDocument.AttachedTemplate;
					try
					{
						enumerator = template.ListTemplates.GetEnumerator();
						while (enumerator.MoveNext())
						{
							Interaction.MsgBox(((ListTemplate)enumerator.Current).Name);
						}
						while (true)
						{
							switch (5)
							{
							case 0:
								break;
							default:
								goto end_IL_0192;
							}
							continue;
							end_IL_0192:
							break;
						}
					}
					finally
					{
						if (enumerator is IDisposable)
						{
							while (true)
							{
								switch (4)
								{
								case 0:
									continue;
								}
								(enumerator as IDisposable).Dispose();
								break;
							}
						}
					}
					object Index;
					try
					{
						ListTemplates listTemplates = application.ActiveDocument.ListTemplates;
						Index = text3;
						listTemplate = listTemplates[ref Index];
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						ProjectData.ClearProjectError();
					}
					object Name;
					if (listTemplate == null)
					{
						ListTemplates listTemplates2 = application.ActiveDocument.ListTemplates;
						Index = Conversions.ToBoolean(A(xmlNode, XC.A(41267)));
						Name = text3;
						ListTemplate listTemplate2 = listTemplates2.Add(ref Index, ref Name);
						text3 = Conversions.ToString(Name);
						listTemplate = listTemplate2;
						int num2 = 1;
						try
						{
							enumerator2 = xmlNode.SelectNodes(XC.A(41292)).GetEnumerator();
							while (enumerator2.MoveNext())
							{
								XmlNode a = (XmlNode)enumerator2.Current;
								ListLevel listLevel = listTemplate.ListLevels[num2];
								listLevel.Alignment = WdListLevelAlignment.wdListLevelAlignLeft;
								listLevel.TrailingCharacter = WdTrailingCharacter.wdTrailingTab;
								listLevel.NumberStyle = (WdListNumberStyle)Conversions.ToInteger(A(a, XC.A(41321)));
								if (listLevel.NumberStyle == WdListNumberStyle.wdListNumberStyleBullet)
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
									listLevel.NumberFormat = A(a, XC.A(41338));
								}
								else
								{
									listLevel.NumberFormat = A(a, XC.A(41338)).Replace(clsStyles.m_A, num2.ToString());
								}
								checked
								{
									listLevel.NumberPosition = application.InchesToPoints(num * (float)(num2 - 1));
									listLevel.TabPosition = listLevel.NumberPosition + application.InchesToPoints(inches);
									listLevel.TextPosition = listLevel.NumberPosition + application.InchesToPoints(inches2);
									Microsoft.Office.Interop.Word.Font font = listLevel.Font;
									string text4 = A(a, XC.A(41357));
									if (text4.Length > 0)
									{
										while (true)
										{
											switch (1)
											{
											case 0:
												continue;
											}
											break;
										}
										font.Size = Conversions.ToSingle(text4);
									}
									text4 = A(a, XC.A(41374));
									if (text4.Length > 0)
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
										font.TextColor.RGB = clsColors.RGB2Ole(text4);
									}
									text4 = A(a, XC.A(41393));
									if (text4.Length > 0)
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
										font.Bold = Conversions.ToInteger(text4);
									}
									text4 = A(a, XC.A(41410));
									if (text4.Length > 0)
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
										font.Italic = Conversions.ToInteger(text4);
									}
									text4 = A(a, XC.A(41431));
									if (text4.Length > 0)
									{
										while (true)
										{
											switch (3)
											{
											case 0:
												continue;
											}
											break;
										}
										font.Name = text4;
									}
									text4 = A(a, XC.A(41448));
									if (text4.Length > 0)
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
										font.Shading.BackgroundPatternColor = Helpers.ColorToWdColor(clsColors.RGB2Color(text4));
									}
									font = null;
									listLevel = null;
									num2++;
								}
							}
							while (true)
							{
								switch (6)
								{
								case 0:
									break;
								default:
									goto end_IL_04e6;
								}
								continue;
								end_IL_04e6:
								break;
							}
						}
						finally
						{
							if (enumerator2 is IDisposable)
							{
								while (true)
								{
									switch (2)
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
					int listLevelNumber = Conversions.ToInteger(xmlNode.SelectSingleNode(XC.A(41471) + text2 + XC.A(7149)).Attributes[XC.A(41512)].Value);
					ListFormat listFormat = selection.Range.ListFormat;
					ListTemplate listTemplate3 = listTemplate;
					Name = RuntimeHelpers.GetObjectValue(Missing.Value);
					Index = RuntimeHelpers.GetObjectValue(Missing.Value);
					object DefaultListBehavior = RuntimeHelpers.GetObjectValue(Missing.Value);
					listFormat.ApplyListTemplate(listTemplate3, ref Name, ref Index, ref DefaultListBehavior);
					listFormat.ListLevelNumber = listLevelNumber;
					_ = null;
					break;
				}
			}
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			Interaction.MsgBox(ex4.Message);
			ProjectData.ClearProjectError();
		}
		undoRecord.EndCustomRecord();
		application.ScreenUpdating = true;
		selection = null;
		undoRecord = null;
		application = null;
		xmlNode = null;
	}

	public static void ApplyTextStyle(string strId)
	{
		Application application = PC.A.Application;
		UndoRecord undoRecord = application.UndoRecord;
		application.ScreenUpdating = false;
		undoRecord.StartCustomRecord(XC.A(41523));
		KC b;
		Selection selection;
		XmlNode xmlNode;
		try
		{
			selection = PC.A.Application.Selection;
			xmlNode = NC.A.SettingsXml.SelectSingleNode(XC.A(41556) + strId + XC.A(7149));
			b = C(xmlNode);
			WdSelectionType type = selection.Type;
			if ((uint)(type - 1) <= 1u)
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
					A(selection, b);
					selection.Font.Superscript = Conversions.ToInteger(xmlNode.Attributes[XC.A(41613)].Value);
					break;
				}
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		undoRecord.EndCustomRecord();
		application.ScreenUpdating = true;
		b = default(KC);
		selection = null;
		undoRecord = null;
		application = null;
		xmlNode = null;
	}

	public static void ApplyTableStyle(string strId)
	{
	}

	private static KC A(XmlNode A)
	{
		KC result = default(KC);
		try
		{
			result = D(A);
			return result;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		return result;
	}

	private static KC B(XmlNode A)
	{
		KC result = default(KC);
		try
		{
			result = D(A);
			return result;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		return result;
	}

	private static KC C(XmlNode A)
	{
		KC result = default(KC);
		try
		{
			result = D(A);
			string text = clsStyles.A(A, XC.A(41613));
			if (text.Length > 0)
			{
				while (true)
				{
					switch (7)
					{
					case 0:
						continue;
					}
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					result.A = Conversions.ToInteger(text);
					break;
				}
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		return result;
	}

	private static KC D(XmlNode A)
	{
		return new KC
		{
			A = clsStyles.A(A, XC.A(41431)),
			B = clsStyles.A(A, XC.A(41357)),
			C = clsStyles.A(A, XC.A(41636)),
			E = clsStyles.A(A, XC.A(41647)),
			F = clsStyles.A(A, XC.A(41656)),
			G = clsStyles.A(A, XC.A(41669)),
			D = clsStyles.A(A, XC.A(41688)),
			H = clsStyles.A(A, XC.A(41707))
		};
	}

	private static void A(Selection A, KC B)
	{
		Microsoft.Office.Interop.Word.Font font = A.Font;
		if (B.A.Length > 0)
		{
			font.Name = B.A;
		}
		if (B.B.Length > 0)
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
			font.Size = Conversions.ToSingle(B.B);
		}
		if (B.E.Length > 0)
		{
			font.Bold = Conversions.ToInteger(B.E);
		}
		if (B.F.Length > 0)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					continue;
				}
				break;
			}
			font.Italic = Conversions.ToInteger(B.F);
		}
		if (B.G.Length > 0)
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
			font.Underline = (WdUnderline)Conversions.ToInteger(B.G);
		}
		if (B.C.Length > 0)
		{
			font.Color = Helpers.ColorToWdColor(clsColors.RGB2Color(B.C));
		}
		if (B.D.Length > 0)
		{
			font.Shading.BackgroundPatternColor = Helpers.ColorToWdColor(clsColors.RGB2Color(B.D));
		}
		if (B.H.Length > 0)
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
			font.AllCaps = Conversions.ToInteger(B.H);
		}
		font = null;
	}

	private static string A(XmlNode A, string B)
	{
		return A.Attributes[B].Value;
	}

	public static string MenuHeadingStyles()
	{
		return A(WordStyles.Headings);
	}

	public static string MenuListStyles()
	{
		StringBuilder stringBuilder = new StringBuilder(XC.A(36369));
		int num = 1;
		string text = XC.A(41722);
		XmlNodeList xmlNodeList;
		XmlNodeList xmlNodeList2;
		try
		{
			xmlNodeList = NC.A.SettingsXml.SelectNodes(XC.A(41751));
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = xmlNodeList.GetEnumerator();
				IEnumerator enumerator2 = default(IEnumerator);
				while (enumerator.MoveNext())
				{
					XmlNode xmlNode = (XmlNode)enumerator.Current;
					string text2 = A(xmlNode, XC.A(21468));
					xmlNodeList2 = xmlNode.SelectNodes(XC.A(41292));
					stringBuilder.Append(XC.A(41796) + text2 + XC.A(41841) + A(xmlNode, XC.A(3725)) + XC.A(41860));
					try
					{
						enumerator2 = xmlNodeList2.GetEnumerator();
						while (enumerator2.MoveNext())
						{
							XmlNode a = (XmlNode)enumerator2.Current;
							string text3 = A(a, XC.A(21468));
							string text4 = A(a, XC.A(3725));
							string text5;
							if (num < 10)
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
								text5 = XC.A(41869) + Conversions.ToString(num) + XC.A(41880) + text4;
							}
							else
							{
								text5 = Conversions.ToString(num) + XC.A(41880) + text4;
							}
							stringBuilder.Append(XC.A(41885) + text + text3 + XC.A(36548) + text5 + XC.A(41910) + text + XC.A(41935) + text2 + XC.A(19662) + text3 + XC.A(41860));
							num = checked(num + 1);
						}
						while (true)
						{
							switch (6)
							{
							case 0:
								break;
							default:
								goto end_IL_022d;
							}
							continue;
							end_IL_022d:
							break;
						}
					}
					finally
					{
						if (enumerator2 is IDisposable)
						{
							while (true)
							{
								switch (3)
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
				while (true)
				{
					switch (5)
					{
					case 0:
						break;
					default:
						goto end_IL_0267;
					}
					continue;
					end_IL_0267:
					break;
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
		xmlNodeList = null;
		xmlNodeList2 = null;
		stringBuilder.Append(XC.A(37850));
		return stringBuilder.ToString();
	}

	public static string MenuTextStyles()
	{
		return A(WordStyles.Text);
	}

	public static string MenuTableStyles()
	{
		return A(WordStyles.Tables);
	}

	private static string A(WordStyles A)
	{
		StringBuilder stringBuilder = new StringBuilder(XC.A(36369));
		int num = 1;
		string text = B(A);
		string text2 = C(A);
		XmlNodeList xmlNodeList;
		try
		{
			xmlNodeList = NC.A.SettingsXml.SelectNodes(XC.A(41950) + text + XC.A(41955));
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = xmlNodeList.GetEnumerator();
				while (enumerator.MoveNext())
				{
					XmlNode obj = (XmlNode)enumerator.Current;
					string value = obj.Attributes[XC.A(21468)].Value;
					string value2 = obj.Attributes[XC.A(3725)].Value;
					string text3;
					if (num < 10)
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
						text3 = XC.A(41869) + Conversions.ToString(num) + XC.A(41880) + value2;
					}
					else
					{
						text3 = Conversions.ToString(num) + XC.A(41880) + value2;
					}
					stringBuilder.Append(XC.A(41885) + text2 + value + XC.A(36548) + text3 + XC.A(41910) + text2 + XC.A(41935) + value + XC.A(41860));
					num = checked(num + 1);
				}
				while (true)
				{
					switch (5)
					{
					case 0:
						break;
					default:
						goto end_IL_019e;
					}
					continue;
					end_IL_019e:
					break;
				}
			}
			finally
			{
				if (enumerator is IDisposable)
				{
					while (true)
					{
						switch (4)
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
		xmlNodeList = null;
		stringBuilder.Append(XC.A(37850));
		return stringBuilder.ToString();
	}

	private static string B(WordStyles A)
	{
		string text = default(string);
		return A switch
		{
			WordStyles.Headings => XC.A(41968), 
			WordStyles.Lists => XC.A(42003), 
			WordStyles.Text => XC.A(42032), 
			WordStyles.Tables => XC.A(42061), 
			_ => text, 
		};
	}

	private static string C(WordStyles A)
	{
		string text = default(string);
		return A switch
		{
			WordStyles.Headings => XC.A(42092), 
			WordStyles.Lists => XC.A(41722), 
			WordStyles.Text => XC.A(42127), 
			WordStyles.Tables => XC.A(42156), 
			_ => text, 
		};
	}

	public static void StyleCycle1()
	{
		Interaction.MsgBox(XC.A(42187));
	}

	public static void StyleCycle2()
	{
		Interaction.MsgBox(XC.A(42190));
	}

	public static void StyleCycle3()
	{
		Interaction.MsgBox(XC.A(42193));
	}

	public static void StyleCycle4()
	{
		Interaction.MsgBox(XC.A(42196));
	}

	public static void StyleCycle5()
	{
		Interaction.MsgBox(XC.A(42199));
	}

	public static void StyleCycle6()
	{
		Interaction.MsgBox(XC.A(42202));
	}
}
