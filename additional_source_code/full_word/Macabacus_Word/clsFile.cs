using System;
using System.Collections;
using System.IO;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Windows.Forms;
using A;
using MacabacusMacros;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic.CompilerServices;

namespace Macabacus_Word;

public sealed class clsFile
{
	public static void SaveAll()
	{
		int try0000_dispatch = -1;
		int num2 = default(int);
		int num = default(int);
		int num3 = default(int);
		SaveFileDialog saveFileDialog = default(SaveFileDialog);
		Document document = default(Document);
		bool flag = default(bool);
		Microsoft.Office.Interop.Word.Application application = default(Microsoft.Office.Interop.Word.Application);
		Document document2 = default(Document);
		IEnumerator enumerator = default(IEnumerator);
		while (true)
		{
			try
			{
				/*Note: ILSpy has introduced the following switch to emulate a goto from catch-block to try-block*/;
				Document document3;
				SaveFileDialog saveFileDialog2;
				object FileFormat;
				object LockComments;
				object Password;
				object AddToRecentFiles;
				object WritePassword;
				object ReadOnlyRecommended;
				object EmbedTrueTypeFonts;
				object SaveNativePictureFormat;
				object SaveFormsData;
				object SaveAsAOCELetter;
				object Encoding;
				object InsertLineBreaks;
				object AllowSubstitutions;
				object LineEnding;
				object AddBiDiMarks;
				Microsoft.Office.Interop.Word.Windows windows;
				object CompatibilityMode;
				Microsoft.Office.Interop.Word.Windows windows2;
				object FileName;
				switch (try0000_dispatch)
				{
				default:
					num2 = 1;
					if (!Licensing.AllowRestrictedMode())
					{
						goto end_IL_0000;
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					goto IL_001f;
				case 873:
					{
						num = num2;
						switch (num3)
						{
						case 1:
							break;
						default:
							goto end_IL_0000_2;
						}
						int num4 = num + 1;
						num = 0;
						switch (num4)
						{
						case 1:
							break;
						case 3:
							goto IL_001f;
						case 4:
							goto IL_0026;
						case 5:
							goto IL_002b;
						case 6:
							goto IL_003d;
						case 7:
							goto IL_0047;
						case 8:
							goto IL_006c;
						case 9:
							goto IL_0097;
						case 10:
							goto IL_00b9;
						case 11:
							goto IL_00bf;
						case 12:
							goto IL_00c9;
						case 13:
							goto IL_00df;
						case 14:
							goto IL_00f0;
						case 15:
							goto IL_0106;
						case 16:
							goto IL_0122;
						case 17:
							goto IL_023e;
						case 18:
							goto IL_0249;
						case 20:
							goto IL_024e;
						case 21:
							goto IL_0258;
						case 19:
						case 22:
							goto IL_0263;
						case 23:
							goto IL_027e;
						case 24:
							goto IL_0296;
						case 25:
							goto IL_02a7;
						case 26:
							goto IL_02c4;
						case 27:
							goto IL_02c7;
						case 28:
							goto end_IL_0000_3;
						default:
							goto end_IL_0000_2;
						case 2:
						case 29:
							goto end_IL_0000;
						}
						goto default;
					}
					IL_00df:
					num2 = 13;
					saveFileDialog.FileName = document.Name;
					goto IL_00f0;
					IL_00f0:
					num2 = 14;
					saveFileDialog.Filter = XC.A(1916);
					goto IL_0106;
					IL_0106:
					num2 = 15;
					if (saveFileDialog.ShowDialog() == DialogResult.OK)
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
						goto IL_0122;
					}
					goto IL_0249;
					IL_0263:
					num2 = 22;
					goto IL_0266;
					IL_001f:
					ProjectData.ClearProjectError();
					num3 = 1;
					goto IL_0026;
					IL_0026:
					num2 = 4;
					flag = false;
					goto IL_002b;
					IL_002b:
					num2 = 5;
					application = PC.A.Application;
					goto IL_003d;
					IL_003d:
					num2 = 6;
					document2 = application.ActiveDocument;
					goto IL_0047;
					IL_0047:
					num2 = 7;
					enumerator = application.Documents.GetEnumerator();
					goto IL_0266;
					IL_0266:
					if (enumerator.MoveNext())
					{
						document = (Document)enumerator.Current;
						goto IL_006c;
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
					goto IL_027e;
					IL_0122:
					num2 = 16;
					document3 = document;
					FileName = (saveFileDialog2 = saveFileDialog).FileName;
					FileFormat = RuntimeHelpers.GetObjectValue(Missing.Value);
					LockComments = RuntimeHelpers.GetObjectValue(Missing.Value);
					Password = RuntimeHelpers.GetObjectValue(Missing.Value);
					AddToRecentFiles = RuntimeHelpers.GetObjectValue(Missing.Value);
					WritePassword = RuntimeHelpers.GetObjectValue(Missing.Value);
					ReadOnlyRecommended = RuntimeHelpers.GetObjectValue(Missing.Value);
					EmbedTrueTypeFonts = RuntimeHelpers.GetObjectValue(Missing.Value);
					SaveNativePictureFormat = RuntimeHelpers.GetObjectValue(Missing.Value);
					SaveFormsData = RuntimeHelpers.GetObjectValue(Missing.Value);
					SaveAsAOCELetter = RuntimeHelpers.GetObjectValue(Missing.Value);
					Encoding = RuntimeHelpers.GetObjectValue(Missing.Value);
					InsertLineBreaks = RuntimeHelpers.GetObjectValue(Missing.Value);
					AllowSubstitutions = RuntimeHelpers.GetObjectValue(Missing.Value);
					LineEnding = RuntimeHelpers.GetObjectValue(Missing.Value);
					AddBiDiMarks = RuntimeHelpers.GetObjectValue(Missing.Value);
					CompatibilityMode = RuntimeHelpers.GetObjectValue(Missing.Value);
					document3.SaveAs2(ref FileName, ref FileFormat, ref LockComments, ref Password, ref AddToRecentFiles, ref WritePassword, ref ReadOnlyRecommended, ref EmbedTrueTypeFonts, ref SaveNativePictureFormat, ref SaveFormsData, ref SaveAsAOCELetter, ref Encoding, ref InsertLineBreaks, ref AllowSubstitutions, ref LineEnding, ref AddBiDiMarks, ref CompatibilityMode);
					saveFileDialog2.FileName = Conversions.ToString(FileName);
					goto IL_023e;
					IL_027e:
					num2 = 23;
					if (enumerator is IDisposable)
					{
						(enumerator as IDisposable).Dispose();
					}
					goto IL_0296;
					IL_0249:
					saveFileDialog = null;
					goto IL_0263;
					IL_0296:
					num2 = 24;
					if (flag)
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
						goto IL_02a7;
					}
					goto IL_02c4;
					IL_024e:
					num2 = 20;
					document.Save();
					goto IL_0258;
					IL_02a7:
					num2 = 25;
					windows = document2.Windows;
					CompatibilityMode = 1;
					windows[ref CompatibilityMode].Activate();
					goto IL_02c4;
					IL_02c4:
					application = null;
					goto IL_02c7;
					IL_02c7:
					num2 = 27;
					document2 = null;
					break;
					IL_023e:
					num2 = 17;
					document.Saved = true;
					goto IL_0249;
					IL_006c:
					num2 = 8;
					if (Operators.CompareString(document.Name, document.FullName, TextCompare: false) == 0)
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
						goto IL_0097;
					}
					goto IL_024e;
					IL_0258:
					num2 = 21;
					document.Saved = true;
					goto IL_0263;
					IL_0097:
					num2 = 9;
					windows2 = document.Windows;
					FileName = 1;
					windows2[ref FileName].Activate();
					goto IL_00b9;
					IL_00b9:
					num2 = 10;
					flag = true;
					goto IL_00bf;
					IL_00bf:
					num2 = 11;
					saveFileDialog = new SaveFileDialog();
					goto IL_00c9;
					IL_00c9:
					num2 = 12;
					saveFileDialog.DefaultExt = XC.A(1907);
					goto IL_00df;
					end_IL_0000_3:
					break;
				}
				num2 = 28;
				A(XC.A(1975));
				break;
				end_IL_0000_2:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 873;
				continue;
			}
			throw ProjectData.CreateProjectError(-2146828237);
			continue;
			end_IL_0000:
			break;
		}
		if (num != 0)
		{
			ProjectData.ClearProjectError();
		}
	}

	public static void SaveUp()
	{
		if (!Licensing.AllowRestrictedMode())
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			A(PC.A.Application.ActiveDocument);
			return;
		}
	}

	private static void A(Document A)
	{
		Microsoft.Office.Interop.Word.Application application = PC.A.Application;
		try
		{
			if (application.Documents.Count == 0)
			{
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
					throw new Exception();
				}
			}
			if (A.Path.Length == 0)
			{
				((_Application)application).get_FileDialog(MsoFileDialogType.msoFileDialogSaveAs).Show();
				throw new Exception();
			}
			string text = clsFile.BaseName(A.Name);
			string extension = Path.GetExtension(A.Name);
			int num = clsFile.VersionNumber(A.Name);
			if (num == 0)
			{
				while (true)
				{
					switch (6)
					{
					case 0:
						continue;
					}
					Forms.WarningMessage(XC.A(1992));
					throw new Exception();
				}
			}
			num = checked(num + 1);
			string text2 = text + Conversions.ToString(num) + extension;
			string text3 = A.Path + XC.A(2144) + text2;
			if (!clsFile.IsPathUrl(text3))
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
				if (File.Exists(text3))
				{
					while (true)
					{
						switch (6)
						{
						case 0:
							break;
						default:
							Forms.WarningMessage(XC.A(2147));
							throw new Exception();
						}
					}
				}
			}
			if (clsFile.NewerVersions(text3).Count > 0)
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
				if (MessageBox.Show(XC.A(2248), XC.A(2438), MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
				{
					while (true)
					{
						switch (6)
						{
						case 0:
							break;
						default:
							throw new Exception();
						}
					}
				}
			}
			try
			{
				object FileName = text3;
				object FileFormat = WdSaveFormat.wdFormatDocumentDefault;
				object LockComments = RuntimeHelpers.GetObjectValue(Missing.Value);
				object Password = RuntimeHelpers.GetObjectValue(Missing.Value);
				object AddToRecentFiles = RuntimeHelpers.GetObjectValue(Missing.Value);
				object WritePassword = RuntimeHelpers.GetObjectValue(Missing.Value);
				object ReadOnlyRecommended = RuntimeHelpers.GetObjectValue(Missing.Value);
				object EmbedTrueTypeFonts = RuntimeHelpers.GetObjectValue(Missing.Value);
				object SaveNativePictureFormat = RuntimeHelpers.GetObjectValue(Missing.Value);
				object SaveFormsData = RuntimeHelpers.GetObjectValue(Missing.Value);
				object SaveAsAOCELetter = RuntimeHelpers.GetObjectValue(Missing.Value);
				object Encoding = RuntimeHelpers.GetObjectValue(Missing.Value);
				object InsertLineBreaks = RuntimeHelpers.GetObjectValue(Missing.Value);
				object AllowSubstitutions = RuntimeHelpers.GetObjectValue(Missing.Value);
				object LineEnding = RuntimeHelpers.GetObjectValue(Missing.Value);
				object AddBiDiMarks = RuntimeHelpers.GetObjectValue(Missing.Value);
				object CompatibilityMode = RuntimeHelpers.GetObjectValue(Missing.Value);
				A.SaveAs2(ref FileName, ref FileFormat, ref LockComments, ref Password, ref AddToRecentFiles, ref WritePassword, ref ReadOnlyRecommended, ref EmbedTrueTypeFonts, ref SaveNativePictureFormat, ref SaveFormsData, ref SaveAsAOCELetter, ref Encoding, ref InsertLineBreaks, ref AllowSubstitutions, ref LineEnding, ref AddBiDiMarks, ref CompatibilityMode);
				text3 = Conversions.ToString(FileName);
				A.Saved = true;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				Forms.ErrorMessage(XC.A(2457) + ex2.Message);
				ProjectData.ClearProjectError();
			}
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			ProjectData.ClearProjectError();
		}
		application = null;
		clsFile.A(XC.A(2572));
	}

	public static bool IsVersionInFileName(Document doc)
	{
		return (Operators.CompareString(doc.FullName, doc.Name, TextCompare: false) != 0) & clsFile.FileNameVersionRegex().IsMatch(doc.Name);
	}

	public static void CloseOthers(Document presThis = null)
	{
		if (!Licensing.AllowRestrictedMode())
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			Microsoft.Office.Interop.Word.Application application = PC.A.Application;
			object RouteDocument;
			Document document;
			try
			{
				if (presThis == null)
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
					presThis = application.ActiveDocument;
				}
				Documents documents = application.Documents;
				for (int i = documents.Count; i >= 1; i = checked(i + -1))
				{
					Documents documents2 = documents;
					object Index = i;
					document = documents2[ref Index];
					if (document == presThis)
					{
						continue;
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
					if (!document.Saved)
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
						Microsoft.Office.Interop.Word.Windows windows = document.Windows;
						Index = 1;
						windows[ref Index].Activate();
					}
					Document document2 = document;
					Index = RuntimeHelpers.GetObjectValue(Missing.Value);
					object OriginalFormat = RuntimeHelpers.GetObjectValue(Missing.Value);
					RouteDocument = RuntimeHelpers.GetObjectValue(Missing.Value);
					document2.Close(ref Index, ref OriginalFormat, ref RouteDocument);
				}
				while (true)
				{
					switch (4)
					{
					case 0:
						continue;
					}
					documents = null;
					break;
				}
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			Microsoft.Office.Interop.Word.Windows windows2 = presThis.Windows;
			RouteDocument = 1;
			windows2[ref RouteDocument].Activate();
			presThis = null;
			document = null;
			application = null;
			A(XC.A(2587));
			return;
		}
	}

	public static void Reopen(Document doc = null)
	{
		int try0000_dispatch = -1;
		int num2 = default(int);
		int num = default(int);
		int num3 = default(int);
		bool flag = default(bool);
		Microsoft.Office.Interop.Word.Application application = default(Microsoft.Office.Interop.Word.Application);
		string text = default(string);
		while (true)
		{
			try
			{
				/*Note: ILSpy has introduced the following switch to emulate a goto from catch-block to try-block*/;
				Document document;
				Documents documents;
				object RouteDocument;
				object OriginalFormat;
				object SaveChanges;
				object AddToRecentFiles;
				object PasswordDocument;
				object PasswordTemplate;
				object Revert;
				object WritePasswordDocument;
				object WritePasswordTemplate;
				object Format;
				object Encoding;
				object Visible;
				object OpenAndRepair;
				object DocumentDirection;
				object NoEncodingDialog;
				object XMLTransform;
				switch (try0000_dispatch)
				{
				default:
					num2 = 1;
					if (!Licensing.AllowRestrictedMode())
					{
						goto end_IL_0000;
					}
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
					goto IL_001f;
				case 741:
					{
						num = num2;
						switch (num3)
						{
						case 1:
							break;
						default:
							goto end_IL_0000_2;
						}
						int num4 = num + 1;
						num = 0;
						switch (num4)
						{
						case 1:
							break;
						case 3:
							goto IL_001f;
						case 4:
							goto IL_0026;
						case 5:
							goto IL_0035;
						case 6:
							goto IL_003a;
						case 7:
							goto IL_0049;
						case 8:
							goto IL_0057;
						case 9:
							goto IL_005f;
						case 10:
							goto IL_0087;
						case 11:
							goto IL_009e;
						case 12:
							goto IL_00eb;
						case 14:
							goto IL_00f7;
						case 13:
						case 15:
							goto IL_00fd;
						case 16:
							goto IL_0111;
						case 17:
							goto IL_011e;
						case 18:
							goto IL_0153;
						case 19:
							goto IL_0259;
						case 20:
							goto IL_026b;
						case 21:
							goto end_IL_0000_3;
						default:
							goto end_IL_0000_2;
						case 2:
						case 22:
							goto end_IL_0000;
						}
						goto default;
					}
					IL_00eb:
					num2 = 12;
					doc.Saved = true;
					goto IL_00fd;
					IL_00f7:
					num2 = 14;
					flag = true;
					goto IL_00fd;
					IL_00fd:
					num2 = 15;
					if (!flag)
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
						goto IL_0111;
					}
					goto IL_0259;
					IL_026b:
					num2 = 20;
					doc = null;
					break;
					IL_001f:
					ProjectData.ClearProjectError();
					num3 = 1;
					goto IL_0026;
					IL_0026:
					num2 = 4;
					application = PC.A.Application;
					goto IL_0035;
					IL_0035:
					num2 = 5;
					flag = false;
					goto IL_003a;
					IL_003a:
					num2 = 6;
					if (doc == null)
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
						goto IL_0049;
					}
					goto IL_0057;
					IL_0111:
					num2 = 16;
					text = doc.FullName;
					goto IL_011e;
					IL_0049:
					num2 = 7;
					doc = application.ActiveDocument;
					goto IL_0057;
					IL_0057:
					num2 = 8;
					if (doc == null)
					{
						break;
					}
					goto IL_005f;
					IL_005f:
					num2 = 9;
					if (Operators.CompareString(doc.Name, doc.FullName, TextCompare: false) != 0)
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
						goto IL_0087;
					}
					goto IL_026b;
					IL_011e:
					num2 = 17;
					document = doc;
					SaveChanges = RuntimeHelpers.GetObjectValue(Missing.Value);
					OriginalFormat = RuntimeHelpers.GetObjectValue(Missing.Value);
					RouteDocument = RuntimeHelpers.GetObjectValue(Missing.Value);
					document.Close(ref SaveChanges, ref OriginalFormat, ref RouteDocument);
					goto IL_0153;
					IL_0087:
					num2 = 10;
					if (!doc.Saved)
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
						goto IL_009e;
					}
					goto IL_00fd;
					IL_0153:
					num2 = 18;
					documents = application.Documents;
					RouteDocument = text;
					OriginalFormat = RuntimeHelpers.GetObjectValue(Missing.Value);
					SaveChanges = RuntimeHelpers.GetObjectValue(Missing.Value);
					AddToRecentFiles = RuntimeHelpers.GetObjectValue(Missing.Value);
					PasswordDocument = RuntimeHelpers.GetObjectValue(Missing.Value);
					PasswordTemplate = RuntimeHelpers.GetObjectValue(Missing.Value);
					Revert = RuntimeHelpers.GetObjectValue(Missing.Value);
					WritePasswordDocument = RuntimeHelpers.GetObjectValue(Missing.Value);
					WritePasswordTemplate = RuntimeHelpers.GetObjectValue(Missing.Value);
					Format = RuntimeHelpers.GetObjectValue(Missing.Value);
					Encoding = RuntimeHelpers.GetObjectValue(Missing.Value);
					Visible = RuntimeHelpers.GetObjectValue(Missing.Value);
					OpenAndRepair = RuntimeHelpers.GetObjectValue(Missing.Value);
					DocumentDirection = RuntimeHelpers.GetObjectValue(Missing.Value);
					NoEncodingDialog = RuntimeHelpers.GetObjectValue(Missing.Value);
					XMLTransform = RuntimeHelpers.GetObjectValue(Missing.Value);
					documents.Open(ref RouteDocument, ref OriginalFormat, ref SaveChanges, ref AddToRecentFiles, ref PasswordDocument, ref PasswordTemplate, ref Revert, ref WritePasswordDocument, ref WritePasswordTemplate, ref Format, ref Encoding, ref Visible, ref OpenAndRepair, ref DocumentDirection, ref NoEncodingDialog, ref XMLTransform);
					text = Conversions.ToString(RouteDocument);
					goto IL_0259;
					IL_009e:
					num2 = 11;
					if (MessageBox.Show(XC.A(2612) + doc.Name + XC.A(2701), XC.A(2438), MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation) == DialogResult.OK)
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
						goto IL_00eb;
					}
					goto IL_00f7;
					IL_0259:
					num2 = 19;
					A(XC.A(2704));
					goto IL_026b;
					end_IL_0000_3:
					break;
				}
				num2 = 21;
				application = null;
				break;
				end_IL_0000_2:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 741;
				continue;
			}
			throw ProjectData.CreateProjectError(-2146828237);
			continue;
			end_IL_0000:
			break;
		}
		if (num == 0)
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
			ProjectData.ClearProjectError();
			return;
		}
	}

	public static Document Duplicate(Document docOld = null)
	{
		if (!Licensing.AllowRestrictedMode())
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
					return null;
				}
			}
		}
		Microsoft.Office.Interop.Word.Application application = PC.A.Application;
		Document result = null;
		if (docOld == null)
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
			docOld = application.ActiveDocument;
		}
		if (docOld.Path.Length > 0)
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
			Documents documents = application.Documents;
			object Template = docOld.FullName;
			object NewTemplate = RuntimeHelpers.GetObjectValue(Missing.Value);
			object DocumentType = RuntimeHelpers.GetObjectValue(Missing.Value);
			object Visible = RuntimeHelpers.GetObjectValue(Missing.Value);
			result = documents.Add(ref Template, ref NewTemplate, ref DocumentType, ref Visible);
			A(XC.A(2735));
		}
		else
		{
			Forms.WarningMessage(XC.A(2772));
		}
		docOld = null;
		application = null;
		return result;
	}

	public static Document OpenDocumentQuietly(Microsoft.Office.Interop.Word.Application wdApp, string strPath)
	{
		Documents documents = wdApp.Documents;
		object FileName = strPath;
		object ConfirmConversions = RuntimeHelpers.GetObjectValue(Missing.Value);
		object ReadOnly = true;
		object AddToRecentFiles = RuntimeHelpers.GetObjectValue(Missing.Value);
		object PasswordDocument = RuntimeHelpers.GetObjectValue(Missing.Value);
		object PasswordTemplate = RuntimeHelpers.GetObjectValue(Missing.Value);
		object Revert = RuntimeHelpers.GetObjectValue(Missing.Value);
		object WritePasswordDocument = RuntimeHelpers.GetObjectValue(Missing.Value);
		object WritePasswordTemplate = RuntimeHelpers.GetObjectValue(Missing.Value);
		object Format = RuntimeHelpers.GetObjectValue(Missing.Value);
		object Encoding = RuntimeHelpers.GetObjectValue(Missing.Value);
		object Visible = false;
		object OpenAndRepair = RuntimeHelpers.GetObjectValue(Missing.Value);
		object DocumentDirection = RuntimeHelpers.GetObjectValue(Missing.Value);
		object NoEncodingDialog = RuntimeHelpers.GetObjectValue(Missing.Value);
		object XMLTransform = RuntimeHelpers.GetObjectValue(Missing.Value);
		Document result = documents.Open(ref FileName, ref ConfirmConversions, ref ReadOnly, ref AddToRecentFiles, ref PasswordDocument, ref PasswordTemplate, ref Revert, ref WritePasswordDocument, ref WritePasswordTemplate, ref Format, ref Encoding, ref Visible, ref OpenAndRepair, ref DocumentDirection, ref NoEncodingDialog, ref XMLTransform);
		strPath = Conversions.ToString(FileName);
		return result;
	}

	public static void OpenFolder(Document doc = null)
	{
		if (!Licensing.AllowRestrictedMode())
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			if (doc == null)
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
				Microsoft.Office.Interop.Word.Application application = PC.A.Application;
				if (application.Documents.Count > 0)
				{
					doc = application.ActiveDocument;
				}
				application = null;
			}
			if (doc != null)
			{
				try
				{
					clsFile.OpenExplorerToFile(doc.FullName);
					A(XC.A(2885));
					return;
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					clsReporting.LogException(ex2);
					ProjectData.ClearProjectError();
					return;
				}
			}
			Forms.WarningMessage(XC.A(2914));
			return;
		}
	}

	public static void CopyPath(Document doc = null)
	{
		if (!Licensing.AllowRestrictedMode())
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
			try
			{
				if (doc == null)
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
					doc = PC.A.Application.ActiveDocument;
				}
				if (doc.Path.Length > 0)
				{
					while (true)
					{
						switch (6)
						{
						case 0:
							continue;
						}
						clsClipboard.SetText(doc.FullName);
						A(XC.A(2955));
						break;
					}
				}
				else
				{
					Forms.WarningMessage(XC.A(2974));
				}
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			doc = null;
			return;
		}
	}

	private static void A(string A)
	{
		clsReporting.LogActivity((ActivityApp)3, (ActivityCategory)8, A);
	}
}
