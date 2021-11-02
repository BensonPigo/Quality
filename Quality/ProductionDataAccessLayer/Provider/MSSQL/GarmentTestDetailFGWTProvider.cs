using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using ProductionDataAccessLayer.Interface;
using ADOHelper.Template.MSSQL;
using ADOHelper.Utility;
using DatabaseObject.ProductionDB;
using DatabaseObject.ViewModel;
using System.Linq;
using ToolKit;
using System.Data.SqlClient;

namespace ProductionDataAccessLayer.Provider.MSSQL
{
    public class GarmentTestDetailFGWTProvider : SQLDAL, IGarmentTestDetailFGWTProvider
    {
        #region 底層連線
        public GarmentTestDetailFGWTProvider(string ConString) : base(ConString) { }
        public GarmentTestDetailFGWTProvider(SQLDataTransaction tra) : base(tra) { }
        #endregion

        #region CRUD Base

        public IList<GarmentTest_Detail_FGWT_ViewModel> Get_GarmentTest_Detail_FGWT(string ID, string No)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@ID", DbType.String, ID } ,
                { "@No", DbType.String, No } ,
            };
            string sqlcmd = @"
select 
[ID]
,[No]
,[Location]
,[LocationText]= CASE WHEN Location='B' THEN 'Bottom'
						WHEN Location='T' THEN 'Top'
						WHEN Location='S' THEN 'Top+Bottom'
						ELSE ''
					END
,[Type]
,[TestDetail]
,[BeforeWash]
,[SizeSpec]
,[AfterWash]
,[Shrinkage]
,[Scale]
,[Criteria]
,[Criteria2]
,[SystemType]
,[EditType] = case 
	when (Criteria is not null  or Criteria2 is not null) and IsInPercentage.value !=1 then '1'
	when Scale is not null and IsInPercentage.value !=1 then '2'
	else '3' end
,[IsInPercentage] = IsInPercentage.value
,[Seq]
,[Result] = IIF(Scale IS NOT NULL
    ,IIF(Scale='4-5' OR Scale ='5','Pass',IIF(Scale='','','Fail'))
    ,IIF( (BeforeWash IS NOT NULL AND AfterWash IS NOT NULL AND Criteria IS NOT NULL AND Shrinkage IS NOT NULL)
          or (Type = 'spirality: Garment - in percentage (average)')
          or (Type = 'spirality: Garment - in percentage (average) (Top Method A)')
          or (Type = 'spirality: Garment - in percentage (average) (Top Method B)')
          or (Type = 'spirality: Garment - in percentage (average) (Bottom Method A)')
          or (Type = 'spirality: Garment - in percentage (average) (Bottom Method B)')
   ,( IIF( TestDetail = '%' OR TestDetail = 'Range%'   
   -- % 為ISP20201331舊資料、Range% 為ISP20201606加上的新資料，兩者都視作百分比
      ---- 百分比 判斷方式
      ,IIF( ISNULL(Criteria,0)  <= ISNULL(Shrinkage,0) AND ISNULL(Shrinkage,0) <= ISNULL(Criteria2,0)
       , 'Pass'
       , 'Fail'
      )
      ---- 非百分比 判斷方式
      ,IIF( ISNULL(AfterWash,0) - ISNULL(BeforeWash,0) <= ISNULL(Criteria,0)
       ,'Pass'
       ,'Fail'
      )
    )
   )
   ,''
 )
)
from GarmentTest_Detail_FGWT f
outer apply(
	select value =
	cast((
	case when f.Type = 'spirality: Garment - in percentage (average)'
			or f.Type = 'spirality: Garment - in percentage (average) (Top Method A)'
			or f.Type = 'spirality: Garment - in percentage (average) (Top Method B)' 
			or f.Type = 'spirality: Garment - in percentage (average) (Bottom Method A)'
			or f.Type = 'spirality: Garment - in percentage (average) (Bottom Method B)' then 1
			else 0 end) as bit)
)IsInPercentage
where ID = @ID
and No = @No
order by Seq asc , Location
";
            return ExecuteList<GarmentTest_Detail_FGWT_ViewModel>(CommandType.Text, sqlcmd, objParameter);
        }

        public bool Chk_FGWTExists(GarmentTest_Detail_ViewModel source)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@ID", DbType.Int64, source.ID } ,
                { "@No", DbType.String, source.No } ,
            };

            string sqlcmd = @"
SELECT 1 FROM GarmentTest_Detail_FGWT WHERE ID = @ID and No = @No
";
            return Convert.ToInt32(ExecuteNonQuery(CommandType.Text, sqlcmd, objParameter)) > 0;
        }

        public bool Create_FGWT(GarmentTest Master, GarmentTest_Detail_ViewModel source)
        {
            bool result = true;
            string washType = string.Empty;
            string fibresType = string.Empty;
            if (source.LineDry != null && source.LineDry == true)
            {
                washType = "Line";
            }
            else if (source.TumbleDry != null && source.TumbleDry == true)
            {
                washType = "Tumnle";
            }
            else if (source.HandWash != null && source.HandWash == true)
            {
                washType = "Hand";
            }

            if (source.Above50NaturalFibres != null && source.Above50NaturalFibres == true)
            {
                fibresType = "Natural";
            }
            else if (source.Above50SyntheticFibres != null && source.Above50SyntheticFibres == true)
            {
                fibresType = "Synthetic";
            }

            SQLParameterCollection objParameter_Loctions = new SQLParameterCollection
            {
                { "@StyleID", DbType.String, Master.StyleID} ,
                { "@BrandID", DbType.String, Master.BrandID } ,
                { "@SeasonID", DbType.String, Master.SeasonID} ,
            };

            string sql_locations = $@"
SELECT locations = STUFF(
	(
        select DISTINCT ',' + sl.Location
	    from Style s
	    INNER JOIN Style_Location sl ON s.Ukey = sl.StyleUkey 
	    where s.id = @StyleID AND s.BrandID = @BrandID AND s.SeasonID = @SeasonID
	    FOR XML PATH('')
	) 
,1,1,'')";
            DataTable dtLocations = ExecuteDataTableByServiceConn(CommandType.Text, sql_locations, objParameter_Loctions);
            List<string> locations = dtLocations.Rows[0]["locations"].ToString().Split(',').ToList();

            List<GarmentTest_Detail_FGWT> fGWTs = new List<GarmentTest_Detail_FGWT>();
            bool containsT = locations.Contains("T");
            bool containsB = locations.Contains("B");

            // 若只有B則寫入Bottom的項目+ALL的項目，若只有T則寫入TOP的項目+ALL的項目，若有B和T則寫入Top+ Bottom的項目+ALL的項目
            // 若為Hand只寫入第88項
            if (washType == "Hand" && source.MtlTypeID != "WOVEN")
            {
                fGWTs = GetDefaultFGWT(false, false, false, source.MtlTypeID, washType, fibresType);
            }
            else
            {
                if (containsT && containsB)
                {
                    fGWTs = GetDefaultFGWT(false, false, true, source.MtlTypeID, washType, fibresType);
                }
                else if (containsT)
                {
                    fGWTs = GetDefaultFGWT(containsT, false, false, source.MtlTypeID, washType, fibresType);
                }
                else
                {
                    fGWTs = GetDefaultFGWT(false, containsB, false, source.MtlTypeID, washType, fibresType);
                }
            }

            long? garmentTest_Detail_ID = source.ID;
            long? garmentTest_Detail_No = source.No;

            StringBuilder insertCmd = new StringBuilder();
            SQLParameterCollection parameters = new SQLParameterCollection();
            int idx = 0;

            // 組合INSERT SQL
            foreach (var fGWT in fGWTs)
            {
                string location = string.Empty;

                switch (fGWT.Location)
                {
                    case "Top":
                        location = "T";
                        break;
                    case "Bottom":
                        location = "B";
                        break;
                    case "Top+Bottom":
                        location = "S";
                        break;
                    default:
                        break;
                }

                if (string.IsNullOrEmpty(fGWT.Scale))
                {
                    if (fGWT.TestDetail.ToUpper() == "CM")
                    {
                        insertCmd.Append($@"

INSERT INTO GarmentTest_Detail_FGWT
           (ID, No, Location, Type ,TestDetail ,Criteria, SystemType, Seq)
     VALUES
           ( {garmentTest_Detail_ID}
           , {garmentTest_Detail_No}
           , @Location{idx}
           , @Type{idx}
           , @TestDetail{idx}
           , @Criteria{idx}
           , @SystemType{idx}
           , @Seq{idx})

");
                    }
                    else
                    {
                        if (fGWT.Type.ToUpper() == "DIMENSIONAL CHANGE: FLAT MADE-UP TEXTILE ARTICLES A) OVERALL LENGTH" || fGWT.Type.ToUpper() == "DIMENSIONAL CHANGE: FLAT MADE-UP TEXTILE ARTICLES B) OVERALL WIDTH")
                        {
                            insertCmd.Append($@"

INSERT INTO GarmentTest_Detail_FGWT
           (ID, No, Location, Type ,TestDetail, SystemType, Seq )
     VALUES
           ( {garmentTest_Detail_ID}
           , {garmentTest_Detail_No}
           , @Location{idx}
           , @Type{idx}
           , @TestDetail{idx} 
           , @SystemType{idx}
           , @Seq{idx})

");
                        }
                        else
                        {
                            insertCmd.Append($@"

INSERT INTO GarmentTest_Detail_FGWT
           (ID, No, Location, Type ,TestDetail ,Criteria ,Criteria2, SystemType, Seq )
     VALUES
           ( {garmentTest_Detail_ID}
           , {garmentTest_Detail_No}
           , @Location{idx}
           , @Type{idx}
           , @TestDetail{idx}
           , @Criteria{idx} 
           , @Criteria2_{idx}
           , @SystemType{idx}
           , @Seq{idx})

");
                        }
                    }
                }
                else
                {
                    insertCmd.Append($@"

INSERT INTO GarmentTest_Detail_FGWT
           (ID, No, Location, Type ,Scale,TestDetail, SystemType, Seq)
     VALUES
           ( {garmentTest_Detail_ID}
           , {garmentTest_Detail_No}
           , @Location{idx}
           , @Type{idx}
           , ''
           , @TestDetail{idx}
           , @SystemType{idx}
           , @Seq{idx})

");
                }

                parameters.Add(new SqlParameter($"@Location{idx}", location));
                parameters.Add(new SqlParameter($"@Type{idx}", fGWT.Type));
                parameters.Add(new SqlParameter($"@TestDetail{idx}", fGWT.TestDetail));
                parameters.Add(new SqlParameter($"@Criteria{idx}", fGWT.Criteria));
                parameters.Add(new SqlParameter($"@Criteria2_{idx}", fGWT.Criteria2));
                parameters.Add(new SqlParameter($"@SystemType{idx}", fGWT.SystemType));
                parameters.Add(new SqlParameter($"@Seq{idx}", fGWT.Seq));
                idx++;
            }

            // 找不到才Insert
            if (Chk_FGWTExists(source) == false)
            {
                result = Convert.ToInt32(ExecuteNonQuery(CommandType.Text, insertCmd.ToString() + this.UpdateGarmentTest_Detail_FGWTShrinkage(source.ID, source.No), parameters)) > 0;
            }

            Update_Spirality_FGWT(source.ID.ToString(), source.No.ToString());

            return result;
        }

        public bool Update_FGWT(List<GarmentTest_Detail_FGWT_ViewModel> source)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();
            int idx = 0;
            string sqlcmd = string.Empty;


            foreach (var item in source)
            {
                // Key
                objParameter.Add(new SqlParameter($"@ID{idx}", item.ID));
                objParameter.Add(new SqlParameter($"@No{idx}", item.No));
                objParameter.Add(new SqlParameter($"@Location{idx}", string.IsNullOrEmpty(item.Location) ? string.Empty : item.Location));
                objParameter.Add(new SqlParameter($"@Type{idx}", item.Type));

                objParameter.Add(new SqlParameter($"@BeforeWash{idx}", item.BeforeWash));
                objParameter.Add(new SqlParameter($"@SizeSpec{idx}", item.SizeSpec));
                objParameter.Add(new SqlParameter($"@AfterWash{idx}", item.AfterWash));
                objParameter.Add(new SqlParameter($"@Shrinkage{idx}", item.Shrinkage));
                objParameter.Add(new SqlParameter($"@Scale{idx}", item.Scale));

                sqlcmd += $@"
                update gf
                	set gf.[BeforeWash] = @BeforeWash{idx},
                		gf.[SizeSpec]  = @SizeSpec{idx},
                		gf.[AfterWash]	= @AfterWash{idx},
                		gf.[Shrinkage]	= @Shrinkage{idx},
                		gf.[Scale]	= @Scale{idx} 
                from GarmentTest_Detail_FGWT gf
                outer apply (
                	select distinct
                		[Location] = iif (slC.cnt > 1, 'S', sl.Location)
                	from GarmentTest g
                	inner join Style s on g.StyleID = s.ID and g.BrandID = s.BrandID and g.SeasonID = s.SeasonID
                	inner join Style_Location sl on s.Ukey = sl.StyleUkey
                	outer apply (
                		select cnt = count(*)
                		from Style_Location sl 
                		where s.Ukey = sl.StyleUkey
                		and sl.Location in ('B', 'T')
                	)slC
                	where gf.ID = g.ID
                )sl 
                where gf.ID = @ID{idx} and gf.No = @No{idx} and gf.Location = @Location{idx} and gf.Type = @Type{idx}

                " + Environment.NewLine;
                idx++;
            }

            bool bolResult = Convert.ToInt32(ExecuteNonQuery(CommandType.Text, sqlcmd, objParameter)) > 0;
            Update_Spirality_FGWT(source.FirstOrDefault().ID.ToString(), source.FirstOrDefault().No.ToString());

            return bolResult;
        }

        private void Update_Spirality_FGWT(string ID, string No)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@ID", DbType.String, ID} ,
                { "@No", DbType.String, No } ,
            };

            string sqlcmd = $@"
update t
set t.Shrinkage = s.MethodA
from GarmentTest_Detail_FGWT t
inner join Garment_Detail_Spirality s on t.ID= s.ID and t.No = s.No
and s.Location = 'T'
where t.ID = @ID and t.No = @No
and t.Type = 'spirality: Garment - in percentage (average) (Top Method A)'

update t
set t.Shrinkage = s.MethodB
from GarmentTest_Detail_FGWT t
inner join Garment_Detail_Spirality s on t.ID= s.ID and t.No = s.No
and s.Location ='T'
where t.ID = @ID and t.No = @No
and t.Type = 'spirality: Garment - in percentage (average) (Top Method B)'

update t
set t.Shrinkage = s.MethodA
from GarmentTest_Detail_FGWT t
inner join Garment_Detail_Spirality s on t.ID= s.ID and t.No = s.No
and s.Location = 'B'
where t.ID = @ID and t.No = @No
and t.Type = 'spirality: Garment - in percentage (average) (Bottom Method A)'

update t
set t.Shrinkage = s.MethodB
from GarmentTest_Detail_FGWT t
inner join Garment_Detail_Spirality s on t.ID= s.ID and t.No = s.No
and s.Location = 'B'
where t.ID = @ID and t.No = @No
and t.Type = 'spirality: Garment - in percentage (average) (Bottom Method B)'
";
            ExecuteNonQuery(CommandType.Text, sqlcmd, objParameter);
        }

        private string UpdateGarmentTest_Detail_FGWTShrinkage(long? ID, long? NO)
        {
            SQLParameterCollection objParameter_Spirality = new SQLParameterCollection
            {
                { "@ID", ID } ,
                { "@No", NO } ,
            };
            string sql_Spirality = $@"
select * from Garment_Detail_Spirality where ID = @ID and No = @No
";
            DataTable dtSpirality = ExecuteDataTableByServiceConn(CommandType.Text, sql_Spirality, objParameter_Spirality);
            string strResut = string.Empty;

            if (dtSpirality != null)
            {
                foreach (DataRow dr in dtSpirality.Rows)
                {
                    if (dr["Location"].ToString().ToLower() == "t")
                    {
                        strResut += $@"
update GarmentTest_Detail_FGWT set Shrinkage = {dr["MethodA"]} where id = {dr["ID"]} and No = {dr["No"]} and type = 'spirality: Garment - in percentage (average) (Top Method A)'
update GarmentTest_Detail_FGWT set Shrinkage = {dr["MethodB"]} where id = {dr["ID"]} and No = {dr["No"]} and type = 'spirality: Garment - in percentage (average) (Top Method B)'
";
                    }
                    else if (dr["Location"].ToString().ToLower() == "b")
                    {
                        strResut += $@"
update GarmentTest_Detail_FGWT set Shrinkage = {dr["MethodA"]} where id = {dr["ID"]} and No = {dr["No"]} and type = 'spirality: Garment - in percentage (average) (Bottom Method A)'
update GarmentTest_Detail_FGWT set Shrinkage = {dr["MethodB"]} where id = {dr["ID"]} and No = {dr["No"]} and type = 'spirality: Garment - in percentage (average) (Bottom Method B)'
";
                    }
                }
            }

            return strResut;
        }

        /// <summary>
        /// 取得預設FGWT
        /// </summary>
        /// <param name="isTop">是否TOP</param>
        /// <param name="isBottom">>是否Bottom</param>
        /// <param name="isTop_Bottom">>是否TOP & Bottom</param>
        /// <param name="mtlTypeID">mtlTypeID</param>
        /// <param name="washType">washType</param>
        /// <param name="fibresType">fibresType</param>
        /// <param name="isAll">>是否All</param>
        /// <returns>預設清單</returns>
        public List<GarmentTest_Detail_FGWT> GetDefaultFGWT(bool isTop, bool isBottom, bool isTop_Bottom, string mtlTypeID, string washType, string fibresType, bool isAll = true)
        {
            string sqlWhere = string.Empty;

            if (!string.IsNullOrEmpty(mtlTypeID))
            {
                sqlWhere += $" and MtlTypeID = '{mtlTypeID}' ";
            }

            if (mtlTypeID == "KNIT")
            {
                if (!string.IsNullOrEmpty(washType))
                {
                    string washing = string.Empty;

                    switch (washType)
                    {
                        case "Hand":
                            washing = "HandWash";
                            break;
                        case "Line":
                            washing = "LineDry";
                            break;
                        case "Tumnle":
                            washing = "TumbleDry";
                            break;
                        default:
                            break;
                    }

                    sqlWhere += $" and Washing = '{washing}' ";
                }

                if (!string.IsNullOrEmpty(fibresType))
                {
                    string fabricComposition = string.Empty;

                    switch (fibresType)
                    {
                        case "Natural":
                            fabricComposition = "Above50NaturaFibres";
                            break;
                        case "Synthetic":
                            fabricComposition = "Above50SyntheticFibres";
                            break;
                        default:
                            break;
                    }

                    sqlWhere += $" and FabricComposition = '{fabricComposition}' ";
                }
            }

            if (isAll || isBottom || isTop || isTop_Bottom)
            {
                List<string> listLocation = new List<string>();

                if (isAll)
                {
                    listLocation.Add("''");
                }

                if (isBottom)
                {
                    listLocation.Add("'B'");
                }

                if (isTop)
                {
                    listLocation.Add("'T'");
                }

                if (isTop_Bottom)
                {
                    listLocation.Add("'S'");
                }

                sqlWhere += $" and Location in ({listLocation.JoinToString(",")}) ";
            }

            string sqlGetDefaultFGWT = $@"
select  Seq,
        [Location] = case when Location = 'T' then 'Top'
                          when Location = 'B' then 'Bottom'
                          when Location = 'S' then 'Top+Bottom'
                          else '' end,
        ReportType,
        SystemType,
        Scale ,
        TestDetail,
        Criteria = ISNULL(Criteria,0),
        Criteria2 = ISNULL(Criteria2,0)
from    Adidas_FGWT with (nolock)
where 1 = 1 {sqlWhere}
";

            DataTable dtResult = ExecuteDataTableByServiceConn(CommandType.Text, sqlGetDefaultFGWT, new SQLParameterCollection());

            List<GarmentTest_Detail_FGWT> defaultFGWTList = dtResult.AsEnumerable().Select(s => new GarmentTest_Detail_FGWT
            {
                Seq = Convert.ToInt32(s["Seq"]),
                Location = s["Location"].ToString(),
                Type = s["ReportType"].ToString(),
                SystemType = s["SystemType"].ToString(),
                Scale = s["Scale"].ToString(),
                TestDetail = s["TestDetail"].ToString(),
                Criteria = Convert.ToDecimal(s["Criteria"]),
                Criteria2 = Convert.ToDecimal(s["Criteria2"]),
            }).ToList();

            return defaultFGWTList;
        }
	#endregion
    }
}
