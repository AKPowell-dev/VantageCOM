using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Xml;
using A;
using MacabacusMacros.UI;
using Macabacus_Word.Links;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic.CompilerServices;

namespace Macabacus_Word.DocBuilder;

public sealed class Base
{
	[CompilerGenerated]
	internal sealed class XB
	{
		public ContentControl A;

		public XB(XB A)
		{
			if (A == null)
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				this.A = A.A;
				return;
			}
		}

		[SpecialName]
		internal bool A(BaseQuestion A)
		{
			return Operators.CompareString(A.ContentControl.ID, this.A.ParentContentControl.ID, TextCompare: false) == 0;
		}
	}

	[CompilerGenerated]
	private static Dictionary<string, string> A;

	[CompilerGenerated]
	private static bool A;

	private static object A = false;

	private static Dictionary<string, string> Fields
	{
		[CompilerGenerated]
		get
		{
			return Base.A;
		}
		[CompilerGenerated]
		set
		{
			Base.A = value;
		}
	} = null;

	public static bool AutoFieldPreview
	{
		[CompilerGenerated]
		get
		{
			return Base.A;
		}
		[CompilerGenerated]
		set
		{
			Base.A = value;
		}
	} = false;

	public static void InspectDocument(Document doc, bool blnManual)
	{
		List<BaseQuestion> list = new List<BaseQuestion>();
		Dictionary<ContentControl, string> listAutoPopFields = new Dictionary<ContentControl, string>();
		if (Fields == null)
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
			Fields = new Dictionary<string, string>();
			try
			{
				XmlNodeList xmlNodeList = NC.A.SettingsXml.DocumentElement.SelectNodes(XC.A(21407));
				IEnumerator enumerator = default(IEnumerator);
				try
				{
					enumerator = xmlNodeList.GetEnumerator();
					while (enumerator.MoveNext())
					{
						XmlNode xmlNode = (XmlNode)enumerator.Current;
						Fields.Add(XC.A(21392) + xmlNode.Attributes[XC.A(21468)].Value, xmlNode.ChildNodes.Item(0).InnerText);
					}
				}
				finally
				{
					if (enumerator is IDisposable)
					{
						while (true)
						{
							switch (5)
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
			finally
			{
				XmlNodeList xmlNodeList = null;
			}
		}
		list = GetQuestions(doc, ref listAutoPopFields);
		if (listAutoPopFields.Count > 0)
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
			Microsoft.Office.Interop.Word.Application application = PC.A.Application;
			UndoRecord undoRecord = application.UndoRecord;
			undoRecord.StartCustomRecord(XC.A(21473));
			application.ScreenUpdating = false;
			using (Dictionary<ContentControl, string>.Enumerator enumerator2 = listAutoPopFields.GetEnumerator())
			{
				while (enumerator2.MoveNext())
				{
					KeyValuePair<ContentControl, string> current = enumerator2.Current;
					try
					{
						if (current.Value.Length > 0)
						{
							ContentControl key = current.Key;
							key.LockContents = false;
							key.LockContentControl = false;
							key.Range.Text = current.Value;
							key.Delete();
							_ = null;
						}
					}
					catch (Exception ex3)
					{
						ProjectData.SetProjectError(ex3);
						Exception ex4 = ex3;
						ProjectData.ClearProjectError();
					}
				}
				while (true)
				{
					switch (2)
					{
					case 0:
						break;
					default:
						goto end_IL_01d3;
					}
					continue;
					end_IL_01d3:
					break;
				}
			}
			application.ScreenUpdating = true;
			undoRecord.EndCustomRecord();
			application = null;
			undoRecord = null;
		}
		if (list.Count > 0)
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
			Pane.Show(list);
			if (!blnManual)
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
				Pane.B();
			}
		}
		else if (blnManual)
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
			Forms.InfoMessage(XC.A(21524));
			Pane.B();
		}
		listAutoPopFields = null;
		list = null;
	}

	public static List<BaseQuestion> GetQuestions(Document doc, ref Dictionary<ContentControl, string> listAutoPopFields)
	{
		List<BaseQuestion> list = new List<BaseQuestion>();
		_ = doc.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.StoryType;
		bool? flag = null;
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = doc.StoryRanges.GetEnumerator();
			IEnumerator enumerator2 = default(IEnumerator);
			XB xB = default(XB);
			while (enumerator.MoveNext())
			{
				Range range = (Range)enumerator.Current;
				do
				{
					try
					{
						enumerator2 = range.ContentControls.GetEnumerator();
						while (enumerator2.MoveNext())
						{
							xB = new XB(xB);
							xB.A = (ContentControl)enumerator2.Current;
							int num;
							if (Conversions.ToBoolean(Operators.NotObject(A)) && !flag.HasValue)
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
								num = (Common.IsLinked(xB.A) ? 1 : 0);
							}
							else
							{
								num = 0;
							}
							if (Conversions.ToBoolean((byte)num != 0))
							{
								flag = object.Equals(xB.A.Appearance, WdContentControlAppearance.wdContentControlHidden);
							}
							if (xB.A.Type != WdContentControlType.wdContentControlRichText)
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
								if (xB.A.Type != WdContentControlType.wdContentControlText)
								{
									continue;
								}
								while (true)
								{
									switch (4)
									{
									case 0:
										continue;
									}
									break;
								}
							}
							string text = xB.A.Tag.ToLower();
							checked
							{
								if (Operators.CompareString(text, XC.A(21597), TextCompare: false) != 0)
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
									if (Operators.CompareString(text, XC.A(21606), TextCompare: false) != 0)
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
										if (Operators.CompareString(text, XC.A(21615), TextCompare: false) != 0)
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
											if (text.StartsWith(XC.A(21687)))
											{
												list.Add(new TextInput(xB.A, list.Count + 1));
											}
											else
											{
												if (!text.StartsWith(XC.A(21392)))
												{
													continue;
												}
												while (true)
												{
													switch (6)
													{
													case 0:
														continue;
													}
													break;
												}
												string value = string.Empty;
												if (!Fields.TryGetValue(text, out value))
												{
													continue;
												}
												if (!AutoFieldPreview)
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
													if (listAutoPopFields == null)
													{
														continue;
													}
													while (true)
													{
														switch (7)
														{
														case 0:
															continue;
														}
														break;
													}
													listAutoPopFields.Add(xB.A, value);
												}
												else
												{
													list.Add(new AutofillField(xB.A, list.Count + 1, value));
												}
											}
											continue;
										}
										try
										{
											BaseQuestion baseQuestion = list.First(xB.A);
											MultipleChoice multipleChoice = (MultipleChoice)baseQuestion;
											multipleChoice.Choices.Add(new Choice(baseQuestion, xB.A, xB.A.Title, multipleChoice.Choices.Count));
											int num2 = multipleChoice.Choices.Count - 2;
											for (int i = 0; i <= num2; i++)
											{
												multipleChoice.Choices[i].CornerRadius = new CornerRadius(0.0);
											}
											while (true)
											{
												switch (6)
												{
												case 0:
													continue;
												}
												multipleChoice = null;
												baseQuestion = null;
												break;
											}
										}
										catch (Exception ex)
										{
											ProjectData.SetProjectError(ex);
											Exception ex2 = ex;
											throw new InvalidTemplateException(XC.A(21632));
										}
									}
									else
									{
										list.Add(new MultipleChoice(xB.A, list.Count + 1));
									}
								}
								else
								{
									list.Add(new YesNoQuestion(xB.A, list.Count + 1));
								}
							}
						}
						while (true)
						{
							switch (2)
							{
							case 0:
								break;
							default:
								goto end_IL_0393;
							}
							continue;
							end_IL_0393:
							break;
						}
					}
					finally
					{
						if (enumerator2 is IDisposable)
						{
							while (true)
							{
								switch (7)
								{
								case 0:
									continue;
								}
								(enumerator2 as IDisposable).Dispose();
								break;
							}
						}
					}
					range = range.NextStoryRange;
				}
				while (range != null);
				while (true)
				{
					switch (1)
					{
					case 0:
						continue;
					}
					break;
				}
			}
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					goto end_IL_03e7;
				}
				continue;
				end_IL_03e7:
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
		if (object.Equals(flag, false))
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
			Forms.InfoMessage(XC.A(21702));
			A = true;
		}
		return list;
	}
}
