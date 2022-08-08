using Ict;
using Ict.Win;
using Sci;
using Sci.Data;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Transactions;
using System.Windows.Forms;

namespace ImageCompress
{
    public partial class ExecuteForm : Sci.Win.Tems.Base
    {
        public ExecuteForm()
        {
            InitializeComponent();
        }

        protected override void OnFormLoaded()
        {
            base.OnFormLoaded();

            this.EditMode = true;
            this.grid.IsEditingReadOnly = false;
            this.Helper.Controls.Grid.Generator(this.grid)
              .CheckBox("Selected", header: string.Empty, width: Widths.AnsiChars(3), iseditable: true, trueValue: 1, falseValue: 0)
              .Text("TableName", header: $"Table", width: Widths.AnsiChars(40), iseditingreadonly: true)
              .Text("ColumnName", header: $"Column", width: Widths.AnsiChars(40), iseditingreadonly: true)
              .Text("Pkey", header: $"Pkey", width: Widths.AnsiChars(40), iseditingreadonly: true)
              ;

            DataTable source = this.GetDatatable();
            this.listControlBindingSource.DataSource = source;
        }


        private DataTable GetDatatable()
        {
            DataTable dt;
            List<SqlParameter> parameters = new List<SqlParameter>();

            string cmd = $@"
USE PMSFile


SELECT  c.TABLE_NAME, c.COLUMN_NAME
	--,c.DATA_TYPE, c.Column_default, c.character_maximum_length, c.numeric_precision, c.is_nullable,CASE WHEN pk.COLUMN_NAME IS NOT NULL THEN 'PRIMARY KEY' ELSE '' END AS KeyType
INTO #PKey
FROM INFORMATION_SCHEMA.COLUMNS c
LEFT JOIN (
            SELECT ku.TABLE_CATALOG,ku.TABLE_SCHEMA,ku.TABLE_NAME,ku.COLUMN_NAME
            FROM INFORMATION_SCHEMA.TABLE_CONSTRAINTS AS tc
            INNER JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE AS ku
                ON tc.CONSTRAINT_TYPE = 'PRIMARY KEY' 
                AND tc.CONSTRAINT_NAME = ku.CONSTRAINT_NAME
         )   pk 
ON  c.TABLE_CATALOG = pk.TABLE_CATALOG
            AND c.TABLE_SCHEMA = pk.TABLE_SCHEMA
            AND c.TABLE_NAME = pk.TABLE_NAME
            AND c.COLUMN_NAME = pk.COLUMN_NAME
where pk.COLUMN_NAME IS NOT NULL
ORDER BY c.TABLE_SCHEMA,c.TABLE_NAME, c.ORDINAL_POSITION 

select Selected = 0
    ,TableName = TABLE_NAME
	,ColumnName = COLUMN_NAME
	,Pkey = Pkey.val
from INFORMATION_SCHEMA.COLUMNS o
OUTER APPLY(
	SELECT val = STUFF((
		SELECT DISTINCT ',' + COLUMN_NAME
		from #PKey p
		where p.TABLE_NAME = o.TABLE_NAME
		FOR XML PATH('')
		),1,1,'')
)Pkey
where TABLE_CATALOG='PMSFile' AND DATA_TYPE='varbinary'

DROP TABLE #PKey
";
            string where = string.Empty;

            DualResult r = DBProxy.Current.Select("PMSFile", cmd, parameters, out dt);

            if (!r)
            {
                this.ShowErr(r);
                this.grid.DataSource = null;
                return null;
            }
            else
            {
                this.grid.DataSource = dt;
            }

            return dt;
        }

        private void brnCompress_Click(object sender, EventArgs e)
        {
            DataTable dataTable;
            if (!MyUtility.Check.Empty(this.grid.DataSource))
            {
                dataTable = (DataTable)this.grid.DataSource;
            }
            else
            {
                return;
            }

            DataRow[] selecteds = dataTable.Select("Selected=1");

            if (selecteds.Length == 0)
            {
                MyUtility.Msg.InfoBox("Please select data first.");
                return;
            }

            this.ShowWaitMessage("Processing...");
            DataTable d = selecteds.CopyToDataTable();

            List<DatabaseSchema> tableList = DataTableToList.ConvertToClassList<DatabaseSchema>(d).ToList();
            bool result = Compress(tableList);
            if (result)
            {
                MyUtility.Msg.InfoBox("Success!");
            }
            this.HideWaitMessage();
        }

        private bool Compress(List<DatabaseSchema> tableList)
        {
            DualResult r;

            var tables = tableList.Select(o => o.TableName).Distinct().ToList();

            // By Table，逐一進行
            // 步驟包含：取出DB圖片、壓縮、寫回原欄位、UPDATE DB
            foreach (var table in tables)
            {
                DataTable dt;
                string updateSQL = string.Empty;

                var imgColumns = tableList.Where(o => o.TableName == table).Select(o => o.ColumnName).ToList();
                var pKeyColumns = tableList.Where(o => o.TableName == table).FirstOrDefault().Pkey.Split(',').Where(o => !MyUtility.Check.Empty(o)).ToList();
                string strColumns = tableList.Where(o => o.TableName == table).Select(o => o.ColumnName).JoinToString(",");

                // 取出圖片
                string sql = $@"
select *
from {table} WITH(NOLOCK)

";
                r = DBProxy.Current.Select("PMSFile", sql, null, out dt);
                if (!r)
                {
                    this.ShowErr(r);
                    return false;
                }

                if (dt != null && dt.Rows.Count > 0)
                {
                    foreach (DataRow dr in dt.AsEnumerable())
                    {
                        List<SqlParameter> para = new List<SqlParameter>();
                        // 每一個Row，的每一個圖片欄位都抓出來，壓到500kb
                        foreach (var column in imgColumns)
                        {
                            if (dr[column] == DBNull.Value)
                            {
                                continue;
                            }
                            byte[] oldImgmp = (byte[])dr[column];
                            // 開始壓縮
                            byte[] newImg = ImageHelper.ImageCompress(oldImgmp);

                            // 塞回去原欄位
                            dr[column] = newImg;
                        }

                        // 開始組合update  SQL
                        List<string> imgString = new List<string>();
                        List<string> pkeyString = new List<string>();

                        foreach (var column in imgColumns)
                        {
                            if (dr[column] == DBNull.Value)
                            {
                                continue;
                            }
                            imgString.Add($@"{column}=@{column}");
                            para.Add(new SqlParameter(column, (byte[])dr[column]));
                        }

                        foreach (var column in pKeyColumns)
                        {
                            pkeyString.Add($@" {column}=@{column} ");
                            para.Add(new SqlParameter(column, dr[column]));
                        }

                        // 該Table有圖片才做
                        if (imgString.Count > 0)
                        {
                            updateSQL = $@"
update {table}
set  {imgString.JoinToString(", ")}
where 1=1 AND {pkeyString.JoinToString("and")}
";
                            // 開始執行
                            using (TransactionScope transactionScope = new TransactionScope())
                            {
                                r = DBProxy.Current.Execute("PMSFile", updateSQL, para);
                                if (!r)
                                {
                                    transactionScope.Dispose();
                                    this.ShowErr(r);
                                    return false;
                                }

                                transactionScope.Complete();
                                transactionScope.Dispose();
                            }
                        }
                    }
                }
            }

            return true;
        }

        public class DatabaseSchema
        {
            public string TableName { get; set; }
            public string ColumnName { get; set; }
            public string Pkey { get; set; }
        }

        private void btnGetSchema_Click(object sender, EventArgs e)
        {
            DataTable source = this.GetDatatable();
            this.listControlBindingSource.DataSource = source;
        }
    }
}
