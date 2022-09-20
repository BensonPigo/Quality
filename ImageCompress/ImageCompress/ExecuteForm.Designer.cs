
namespace ImageCompress
{
    partial class ExecuteForm
    {
        /// <summary>
        /// 設計工具所需的變數。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清除任何使用中的資源。
        /// </summary>
        /// <param name="disposing">如果應該處置受控資源則為 true，否則為 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form 設計工具產生的程式碼

        /// <summary>
        /// 此為設計工具支援所需的方法 - 請勿使用程式碼編輯器修改
        /// 這個方法的內容。
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.panel1 = new Sci.Win.UI.Panel();
            this.panel2 = new Sci.Win.UI.Panel();
            this.displayConnString = new Sci.Win.UI.DisplayBox();
            this.label1 = new Sci.Win.UI.Label();
            this.btnGetSchema = new Sci.Win.UI.Button();
            this.brnCompress = new Sci.Win.UI.Button();
            this.panel3 = new Sci.Win.UI.Panel();
            this.panel4 = new Sci.Win.UI.Panel();
            this.panel5 = new Sci.Win.UI.Panel();
            this.grid = new Sci.Win.UI.Grid();
            this.listControlBindingSource = new Sci.Win.UI.ListControlBindingSource(this.components);
            this.panel2.SuspendLayout();
            this.panel5.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.grid)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.listControlBindingSource)).BeginInit();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(1159, 19);
            this.panel1.TabIndex = 0;
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.displayConnString);
            this.panel2.Controls.Add(this.label1);
            this.panel2.Controls.Add(this.btnGetSchema);
            this.panel2.Controls.Add(this.brnCompress);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panel2.Location = new System.Drawing.Point(0, 400);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(1159, 50);
            this.panel2.TabIndex = 1;
            // 
            // displayConnString
            // 
            this.displayConnString.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.displayConnString.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(183)))), ((int)(((byte)(227)))), ((int)(((byte)(255)))));
            this.displayConnString.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(255)))));
            this.displayConnString.Location = new System.Drawing.Point(294, 15);
            this.displayConnString.Name = "displayConnString";
            this.displayConnString.Size = new System.Drawing.Size(633, 23);
            this.displayConnString.TabIndex = 3;
            // 
            // label1
            // 
            this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.label1.Location = new System.Drawing.Point(159, 14);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(126, 23);
            this.label1.TabIndex = 2;
            this.label1.Text = "Connection string";
            // 
            // btnGetSchema
            // 
            this.btnGetSchema.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnGetSchema.Location = new System.Drawing.Point(18, 10);
            this.btnGetSchema.Name = "btnGetSchema";
            this.btnGetSchema.Size = new System.Drawing.Size(138, 30);
            this.btnGetSchema.TabIndex = 1;
            this.btnGetSchema.Text = "Get Schema";
            this.btnGetSchema.UseVisualStyleBackColor = true;
            this.btnGetSchema.Click += new System.EventHandler(this.btnGetSchema_Click);
            // 
            // brnCompress
            // 
            this.brnCompress.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.brnCompress.Location = new System.Drawing.Point(1046, 8);
            this.brnCompress.Name = "brnCompress";
            this.brnCompress.Size = new System.Drawing.Size(96, 30);
            this.brnCompress.TabIndex = 0;
            this.brnCompress.Text = "Compress";
            this.brnCompress.UseVisualStyleBackColor = true;
            this.brnCompress.Click += new System.EventHandler(this.brnCompress_Click);
            // 
            // panel3
            // 
            this.panel3.Dock = System.Windows.Forms.DockStyle.Left;
            this.panel3.Location = new System.Drawing.Point(0, 19);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(17, 381);
            this.panel3.TabIndex = 2;
            // 
            // panel4
            // 
            this.panel4.Dock = System.Windows.Forms.DockStyle.Right;
            this.panel4.Location = new System.Drawing.Point(1142, 19);
            this.panel4.Name = "panel4";
            this.panel4.Size = new System.Drawing.Size(17, 381);
            this.panel4.TabIndex = 3;
            // 
            // panel5
            // 
            this.panel5.Controls.Add(this.grid);
            this.panel5.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel5.Location = new System.Drawing.Point(17, 19);
            this.panel5.Name = "panel5";
            this.panel5.Size = new System.Drawing.Size(1125, 381);
            this.panel5.TabIndex = 4;
            // 
            // grid
            // 
            this.grid.AllowUserToAddRows = false;
            this.grid.AllowUserToDeleteRows = false;
            this.grid.AllowUserToResizeRows = false;
            this.grid.BackgroundColor = System.Drawing.SystemColors.Control;
            this.grid.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.EnableWithoutHeaderText;
            this.grid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.grid.DataSource = this.listControlBindingSource;
            this.grid.Dock = System.Windows.Forms.DockStyle.Fill;
            this.grid.EditingEnter = Ict.Win.UI.DataGridViewEditingEnter.NextCellOrNextRow;
            this.grid.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter;
            this.grid.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.grid.GridColor = System.Drawing.Color.FromArgb(((int)(((byte)(161)))), ((int)(((byte)(162)))), ((int)(((byte)(163)))));
            this.grid.Location = new System.Drawing.Point(0, 0);
            this.grid.Name = "grid";
            this.grid.RowTemplate.DefaultCellStyle.SelectionBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(196)))), ((int)(((byte)(228)))), ((int)(((byte)(255)))));
            this.grid.RowTemplate.DefaultCellStyle.SelectionForeColor = System.Drawing.Color.Black;
            this.grid.RowTemplate.Height = 24;
            this.grid.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.grid.ShowCellToolTips = false;
            this.grid.Size = new System.Drawing.Size(1125, 381);
            this.grid.TabIndex = 0;
            // 
            // ExecuteForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1159, 450);
            this.Controls.Add(this.panel5);
            this.Controls.Add(this.panel4);
            this.Controls.Add(this.panel3);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.panel1);
            this.IsToolbarVisible = false;
            this.Name = "ExecuteForm";
            this.OnLineHelpID = "Sci.Win.Tems.Base";
            this.Text = "Image Compress Forme";
            this.Controls.SetChildIndex(this.panel1, 0);
            this.Controls.SetChildIndex(this.panel2, 0);
            this.Controls.SetChildIndex(this.panel3, 0);
            this.Controls.SetChildIndex(this.panel4, 0);
            this.Controls.SetChildIndex(this.panel5, 0);
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            this.panel5.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.grid)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.listControlBindingSource)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private Sci.Win.UI.Panel panel1;
        private Sci.Win.UI.Panel panel2;
        private Sci.Win.UI.Panel panel3;
        private Sci.Win.UI.Panel panel4;
        private Sci.Win.UI.Panel panel5;
        private Sci.Win.UI.Grid grid;
        private Sci.Win.UI.ListControlBindingSource listControlBindingSource;
        private Sci.Win.UI.Button brnCompress;
        private Sci.Win.UI.Button btnGetSchema;
        private Sci.Win.UI.DisplayBox displayConnString;
        private Sci.Win.UI.Label label1;
    }
}

