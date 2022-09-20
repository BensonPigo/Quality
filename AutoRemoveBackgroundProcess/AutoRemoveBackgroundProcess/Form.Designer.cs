
namespace AutoRemoveBackgroundProcess
{
    partial class Form
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
            this.ChkExcel = new Sci.Win.UI.CheckBox();
            this.numExcel = new Sci.Win.UI.NumericBox();
            this.btnRemove = new Sci.Win.UI.Button();
            this.label1 = new Sci.Win.UI.Label();
            this.label2 = new Sci.Win.UI.Label();
            this.btnSettingSave = new Sci.Win.UI.Button();
            this.btnClose = new Sci.Win.UI.Button();
            this.SuspendLayout();
            // 
            // ChkExcel
            // 
            this.ChkExcel.AutoSize = true;
            this.ChkExcel.Checked = true;
            this.ChkExcel.CheckState = System.Windows.Forms.CheckState.Checked;
            this.ChkExcel.Font = new System.Drawing.Font("新細明體", 12F);
            this.ChkExcel.Location = new System.Drawing.Point(26, 75);
            this.ChkExcel.Name = "ChkExcel";
            this.ChkExcel.Size = new System.Drawing.Size(88, 20);
            this.ChkExcel.TabIndex = 0;
            this.ChkExcel.Text = "Excel.exe";
            this.ChkExcel.UseVisualStyleBackColor = true;
            // 
            // numExcel
            // 
            this.numExcel.Location = new System.Drawing.Point(188, 73);
            this.numExcel.Minimum = new decimal(new int[] {
            10,
            0,
            0,
            0});
            this.numExcel.Name = "numExcel";
            this.numExcel.NullValue = new decimal(new int[] {
            0,
            0,
            0,
            0});
            this.numExcel.Size = new System.Drawing.Size(87, 22);
            this.numExcel.TabIndex = 1;
            this.numExcel.Value = new decimal(new int[] {
            0,
            0,
            0,
            0});
            // 
            // btnRemove
            // 
            this.btnRemove.Location = new System.Drawing.Point(372, 65);
            this.btnRemove.Name = "btnRemove";
            this.btnRemove.Size = new System.Drawing.Size(80, 30);
            this.btnRemove.TabIndex = 2;
            this.btnRemove.Text = "Remove";
            this.btnRemove.UseVisualStyleBackColor = true;
            this.btnRemove.Click += new System.EventHandler(this.btnRemove_Click);
            // 
            // label1
            // 
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F);
            this.label1.Location = new System.Drawing.Point(26, 20);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(117, 34);
            this.label1.TabIndex = 3;
            this.label1.Text = "Process Name";
            // 
            // label2
            // 
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F);
            this.label2.Location = new System.Drawing.Point(188, 20);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(87, 34);
            this.label2.TabIndex = 4;
            this.label2.Text = "Idle Time";
            // 
            // btnSettingSave
            // 
            this.btnSettingSave.Location = new System.Drawing.Point(372, 20);
            this.btnSettingSave.Name = "btnSettingSave";
            this.btnSettingSave.Size = new System.Drawing.Size(80, 30);
            this.btnSettingSave.TabIndex = 5;
            this.btnSettingSave.Text = "Setting Save";
            this.btnSettingSave.UseVisualStyleBackColor = true;
            this.btnSettingSave.Click += new System.EventHandler(this.btnSettingSave_Click);
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(372, 119);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(80, 30);
            this.btnClose.TabIndex = 6;
            this.btnClose.Text = "Close";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // Form
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(486, 176);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.btnSettingSave);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnRemove);
            this.Controls.Add(this.numExcel);
            this.Controls.Add(this.ChkExcel);
            this.Name = "Form";
            this.Text = "Auto Remove Background Process";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private Sci.Win.UI.CheckBox ChkExcel;
        private Sci.Win.UI.NumericBox numExcel;
        private Sci.Win.UI.Button btnRemove;
        private Sci.Win.UI.Label label1;
        private Sci.Win.UI.Label label2;
        private Sci.Win.UI.Button btnSettingSave;
        private Sci.Win.UI.Button btnClose;
    }
}

