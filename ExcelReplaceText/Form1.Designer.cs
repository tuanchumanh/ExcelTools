namespace ExcelReplaceText
{
	partial class Form1
	{
		/// <summary>
		/// 必要なデザイナー変数です。
		/// </summary>
		private System.ComponentModel.IContainer components = null;

		/// <summary>
		/// 使用中のリソースをすべてクリーンアップします。
		/// </summary>
		/// <param name="disposing">マネージ リソースが破棄される場合 true、破棄されない場合は false です。</param>
		protected override void Dispose(bool disposing)
		{
			if (disposing && (components != null))
			{
				components.Dispose();
			}
			base.Dispose(disposing);
		}

		#region Windows フォーム デザイナーで生成されたコード

		/// <summary>
		/// デザイナー サポートに必要なメソッドです。このメソッドの内容を
		/// コード エディターで変更しないでください。
		/// </summary>
		private void InitializeComponent()
		{
			this.btnReplace = new System.Windows.Forms.Button();
			this.btnBrowse = new System.Windows.Forms.Button();
			this.dataGridView1 = new System.Windows.Forms.DataGridView();
			this.Find = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.Replace = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.listBox1 = new System.Windows.Forms.ListBox();
			this.cmbSheets = new System.Windows.Forms.ComboBox();
			this.btnClose = new System.Windows.Forms.Button();
			this.checkBox1 = new System.Windows.Forms.CheckBox();
			((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
			this.SuspendLayout();
			// 
			// btnReplace
			// 
			this.btnReplace.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btnReplace.Location = new System.Drawing.Point(546, 449);
			this.btnReplace.Name = "btnReplace";
			this.btnReplace.Size = new System.Drawing.Size(75, 23);
			this.btnReplace.TabIndex = 1;
			this.btnReplace.Text = "Replace";
			this.btnReplace.UseVisualStyleBackColor = true;
			this.btnReplace.Click += new System.EventHandler(this.btnReplace_Click);
			// 
			// btnBrowse
			// 
			this.btnBrowse.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnBrowse.Location = new System.Drawing.Point(13, 449);
			this.btnBrowse.Name = "btnBrowse";
			this.btnBrowse.Size = new System.Drawing.Size(75, 23);
			this.btnBrowse.TabIndex = 2;
			this.btnBrowse.Text = "Browse...";
			this.btnBrowse.UseVisualStyleBackColor = true;
			this.btnBrowse.Click += new System.EventHandler(this.btnBrowse_Click);
			// 
			// dataGridView1
			// 
			this.dataGridView1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
			this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Find,
            this.Replace});
			this.dataGridView1.Location = new System.Drawing.Point(13, 13);
			this.dataGridView1.Name = "dataGridView1";
			this.dataGridView1.RowTemplate.Height = 21;
			this.dataGridView1.Size = new System.Drawing.Size(608, 150);
			this.dataGridView1.TabIndex = 3;
			this.dataGridView1.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellValueChanged);
			this.dataGridView1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dataGridView1_KeyDown);
			// 
			// Find
			// 
			this.Find.HeaderText = "Find";
			this.Find.Name = "Find";
			this.Find.Width = 250;
			// 
			// Replace
			// 
			this.Replace.HeaderText = "Replace";
			this.Replace.Name = "Replace";
			this.Replace.Width = 250;
			// 
			// listBox1
			// 
			this.listBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.listBox1.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawVariable;
			this.listBox1.FormattingEnabled = true;
			this.listBox1.Location = new System.Drawing.Point(13, 169);
			this.listBox1.Name = "listBox1";
			this.listBox1.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
			this.listBox1.Size = new System.Drawing.Size(608, 251);
			this.listBox1.TabIndex = 4;
			this.listBox1.DrawItem += new System.Windows.Forms.DrawItemEventHandler(this.listBox1_DrawItem);
			this.listBox1.MeasureItem += new System.Windows.Forms.MeasureItemEventHandler(this.listBox1_MeasureItem);
			// 
			// cmbSheets
			// 
			this.cmbSheets.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.cmbSheets.FormattingEnabled = true;
			this.cmbSheets.Location = new System.Drawing.Point(13, 423);
			this.cmbSheets.Name = "cmbSheets";
			this.cmbSheets.Size = new System.Drawing.Size(269, 20);
			this.cmbSheets.TabIndex = 5;
			this.cmbSheets.SelectedIndexChanged += new System.EventHandler(this.cmbSheets_SelectedIndexChanged);
			// 
			// btnClose
			// 
			this.btnClose.Enabled = false;
			this.btnClose.Location = new System.Drawing.Point(94, 449);
			this.btnClose.Name = "btnClose";
			this.btnClose.Size = new System.Drawing.Size(75, 23);
			this.btnClose.TabIndex = 6;
			this.btnClose.Text = "Close File";
			this.btnClose.UseVisualStyleBackColor = true;
			this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
			// 
			// checkBox1
			// 
			this.checkBox1.AutoSize = true;
			this.checkBox1.Location = new System.Drawing.Point(467, 453);
			this.checkBox1.Name = "checkBox1";
			this.checkBox1.Size = new System.Drawing.Size(73, 16);
			this.checkBox1.TabIndex = 7;
			this.checkBox1.Text = "Overwrite";
			this.checkBox1.UseVisualStyleBackColor = true;
			this.checkBox1.CheckedChanged += new System.EventHandler(this.checkBox1_CheckedChanged);
			// 
			// Form1
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(633, 484);
			this.Controls.Add(this.checkBox1);
			this.Controls.Add(this.btnClose);
			this.Controls.Add(this.cmbSheets);
			this.Controls.Add(this.listBox1);
			this.Controls.Add(this.dataGridView1);
			this.Controls.Add(this.btnBrowse);
			this.Controls.Add(this.btnReplace);
			this.MinimumSize = new System.Drawing.Size(641, 511);
			this.Name = "Form1";
			this.Text = "Find & Replace";
			((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.Button btnReplace;
		private System.Windows.Forms.Button btnBrowse;
		private System.Windows.Forms.DataGridView dataGridView1;
		private System.Windows.Forms.DataGridViewTextBoxColumn Find;
		private System.Windows.Forms.DataGridViewTextBoxColumn Replace;
		private System.Windows.Forms.ListBox listBox1;
		private System.Windows.Forms.ComboBox cmbSheets;
		private System.Windows.Forms.Button btnClose;
		private System.Windows.Forms.CheckBox checkBox1;
	}
}

