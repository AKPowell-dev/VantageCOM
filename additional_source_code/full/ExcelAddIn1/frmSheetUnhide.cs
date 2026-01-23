using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.CompilerServices;
using System.Windows.Forms;
using A;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1;

[DesignerGenerated]
public sealed class frmSheetUnhide : Form
{
	private IContainer m_A;

	[AccessedThroughProperty("lstSheets")]
	[CompilerGenerated]
	private System.Windows.Forms.ListBox m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("Label1")]
	private Label m_A;

	[AccessedThroughProperty("btnCancel")]
	[CompilerGenerated]
	private Button m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("btnOk")]
	private Button B;

	private Microsoft.Office.Interop.Excel.Application m_A;

	private Dictionary<int, object> m_A;

	internal virtual System.Windows.Forms.ListBox lstSheets
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_A = value;
		}
	}

	internal virtual Label Label1
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_A = value;
		}
	}

	internal virtual Button btnCancel
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_A = value;
		}
	}

	internal virtual Button btnOk
	{
		[CompilerGenerated]
		get
		{
			return B;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			B = value;
		}
	}

	public frmSheetUnhide()
	{
		base.FormClosing += A;
		A();
		this.m_A = MH.A.Application;
		this.m_A = new Dictionary<int, object>();
		int num = 0;
		lstSheets.BeginUpdate();
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = this.m_A.ActiveWorkbook.Sheets.GetEnumerator();
			while (enumerator.MoveNext())
			{
				object objectValue = RuntimeHelpers.GetObjectValue(enumerator.Current);
				if (!Operators.ConditionalCompareObjectEqual(NewLateBinding.LateGet(objectValue, null, VH.A(41367), new object[0], null, null, null), XlSheetVisibility.xlSheetHidden, TextCompare: false))
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					if (!Operators.ConditionalCompareObjectEqual(NewLateBinding.LateGet(objectValue, null, VH.A(41367), new object[0], null, null, null), XlSheetVisibility.xlSheetVeryHidden, TextCompare: false))
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
						break;
					}
				}
				this.m_A.Add(num, RuntimeHelpers.GetObjectValue(objectValue));
				lstSheets.Items.Add(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, null, VH.A(19019), new object[0], null, null, null)));
				num = checked(num + 1);
			}
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					goto end_IL_014b;
				}
				continue;
				end_IL_014b:
				break;
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
		lstSheets.EndUpdate();
	}

	[DebuggerNonUserCode]
	protected override void Dispose(bool disposing)
	{
		try
		{
			if (!disposing)
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
				if (this.m_A == null)
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
					this.m_A.Dispose();
					return;
				}
			}
		}
		finally
		{
			base.Dispose(disposing);
		}
	}

	[DebuggerStepThrough]
	private void A()
	{
		lstSheets = new System.Windows.Forms.ListBox();
		Label1 = new Label();
		btnCancel = new Button();
		btnOk = new Button();
		SuspendLayout();
		lstSheets.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
		lstSheets.Font = new System.Drawing.Font(VH.A(50021), 9f);
		lstSheets.FormattingEnabled = true;
		lstSheets.IntegralHeight = false;
		lstSheets.ItemHeight = 15;
		lstSheets.Location = new System.Drawing.Point(8, 26);
		lstSheets.Name = VH.A(204919);
		lstSheets.SelectionMode = SelectionMode.MultiExtended;
		lstSheets.Size = new Size(237, 108);
		lstSheets.TabIndex = 0;
		Label1.AutoSize = true;
		Label1.Location = new System.Drawing.Point(5, 8);
		Label1.Name = VH.A(204938);
		Label1.Size = new Size(73, 13);
		Label1.TabIndex = 1;
		Label1.Text = VH.A(204951);
		btnCancel.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
		btnCancel.DialogResult = DialogResult.Cancel;
		btnCancel.Font = new System.Drawing.Font(VH.A(50021), 9f);
		btnCancel.Location = new System.Drawing.Point(181, 141);
		btnCancel.Name = VH.A(204978);
		btnCancel.Size = new Size(64, 22);
		btnCancel.TabIndex = 2;
		btnCancel.Text = VH.A(180569);
		btnCancel.UseVisualStyleBackColor = true;
		btnOk.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
		btnOk.DialogResult = DialogResult.OK;
		btnOk.Font = new System.Drawing.Font(VH.A(50021), 9f);
		btnOk.Location = new System.Drawing.Point(111, 141);
		btnOk.Name = VH.A(204997);
		btnOk.Size = new Size(64, 22);
		btnOk.TabIndex = 3;
		btnOk.Text = VH.A(205008);
		btnOk.UseVisualStyleBackColor = true;
		base.AcceptButton = btnOk;
		base.AutoScaleDimensions = new SizeF(6f, 13f);
		base.AutoScaleMode = AutoScaleMode.Font;
		base.CancelButton = btnCancel;
		base.ClientSize = new Size(252, 170);
		base.Controls.Add(btnOk);
		base.Controls.Add(btnCancel);
		base.Controls.Add(Label1);
		base.Controls.Add(lstSheets);
		base.MaximizeBox = false;
		base.MinimizeBox = false;
		MinimumSize = new Size(268, 209);
		base.Name = VH.A(205013);
		base.ShowIcon = false;
		base.ShowInTaskbar = false;
		base.StartPosition = FormStartPosition.CenterParent;
		Text = VH.A(205042);
		ResumeLayout(performLayout: false);
		PerformLayout();
	}

	private void A(object A, FormClosingEventArgs B)
	{
		object obj = null;
		if (base.DialogResult == DialogResult.OK)
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			this.m_A.ScreenUpdating = false;
			try
			{
				IEnumerator enumerator = default(IEnumerator);
				try
				{
					enumerator = lstSheets.SelectedIndices.GetEnumerator();
					while (enumerator.MoveNext())
					{
						int key = Conversions.ToInteger(enumerator.Current);
						obj = RuntimeHelpers.GetObjectValue(this.m_A[key]);
						NewLateBinding.LateSet(obj, null, VH.A(41367), new object[1] { XlSheetVisibility.xlSheetVisible }, null, null);
					}
					while (true)
					{
						switch (7)
						{
						case 0:
							break;
						default:
							goto end_IL_0098;
						}
						continue;
						end_IL_0098:
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
				NewLateBinding.LateCall(obj, null, VH.A(39985), new object[0], null, null, null, IgnoreReturn: true);
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			this.m_A.ScreenUpdating = true;
		}
		JH.A(RuntimeHelpers.GetObjectValue(obj));
		JH.A((object)this.m_A);
		this.m_A = null;
	}
}
