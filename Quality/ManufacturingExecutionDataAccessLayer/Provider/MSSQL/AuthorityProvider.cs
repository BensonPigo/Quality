using ADOHelper.Template.MSSQL;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ManufacturingExecutionDataAccessLayer.Interface;
using ADOHelper.Utility;
using DatabaseObject.ResultModel;
using System.Web.Mvc;
using System.Data;
using DatabaseObject.RequestModel;

namespace ManufacturingExecutionDataAccessLayer.Provider.MSSQL
{
    public class AuthorityProvider : SQLDAL, IAuthorityProvider
    {
        #region 底層連線
        public AuthorityProvider(string ConString) : base(ConString) { }
        public AuthorityProvider(SQLDataTransaction tra) : base(tra) { }
        #endregion
        
        public IList<UserList_Browse> Get_User_List_Browse()
        {
            StringBuilder SbSql = new StringBuilder();

            SbSql.Append($@"
select   UserID = p.ID
		,[Name] = IIF(isnull(pp1.Name,'') = '', mp1.name,pp1.name) 
		,p.Position
from Quality_Pass1 p
left join ManufacturingExecution.dbo.Pass1 mp1 ON p.ID= mp1.ID
left join Production.dbo.Pass1 pp1 on p.ID = pp1.id
WHERE p.ID != 'SCIMIS'

");

            return ExecuteList<UserList_Browse>(CommandType.Text, SbSql.ToString(), new SQLParameterCollection());
        }
       
        public IList<UserList_Authority> Get_User_List_Head(string UserID)
        {
            StringBuilder SbSql = new StringBuilder();

            SbSql.Append($@"
select TOp 1 UserID = p.ID
		,[Name] = IIF(isnull(pp1.Name,'') = '', mp1.name,pp1.name)
		,[Password] = IIF(isnull(pp1.Password,'') = '', mp1.Password, pp1.Password)
		,[Email] = IIF(isnull(pp1.Email,'')='', mp1.Email, pp1.Email)
		,p.Position
from Quality_Pass1 p
left join ManufacturingExecution.dbo.Pass1 mp1 ON p.ID= mp1.ID
left join Production.dbo.Pass1 pp1 on p.ID = pp1.id
inner join Quality_Position pos on pos.ID=p.Position
where p.ID='{UserID}'


");

            return ExecuteList<UserList_Authority>(CommandType.Text, SbSql.ToString(), new SQLParameterCollection());
        }
       
        public IList<Module_Detail> Get_User_List_Detail(string UserID)
        {
            StringBuilder SbSql = new StringBuilder();

            SbSql.Append($@"
select   MenuID = m.ID
		,m.ModuleName
		,m.FunctionName
		,p2.Used
from Quality_Pass1 p
left join ManufacturingExecution.dbo.Pass1 mp1 ON p.ID= mp1.ID
left join Production.dbo.Pass1 pp1 on p.ID = pp1.idinner join Quality_Position pos on pos.ID=p.Position
inner join Quality_Pass2 p2 on p2.PositionID=pos.ID
inner join Quality_Menu m on p2.MenuID=m.id
where p.ID='{UserID}'
order by m.ModuleSeq, m.FunctionSeq

");

            return ExecuteList<Module_Detail>(CommandType.Text, SbSql.ToString(), new SQLParameterCollection());
        }

        public bool Update_User_List_Detail(UserList_Authority_Request Req)
        {
            StringBuilder SbSql = new StringBuilder();

            SbSql.Append($@"
UPDATE Quality_Pass1
SET Position='{Req.Position}'
WHERE Id='{Req.UserID}'
;
");

            /*表身只唯讀， 因此註解這段
            foreach (var data in Req.DataList)
            {
                SbSql.Append($@"
UPDATE p2
SET p2.Used='{(data.Used ? "1" : "0")}'
from Quality_Pass1 p
inner join InternalWeb.dbo.Pass1 p1 on p.ID=p1.ID
inner join Quality_Position pos on pos.ID=p.Position
inner join Quality_Pass2 p2 on p2.PositionID=pos.ID AND p2.FactoryID=pos.Factory
inner join Quality_Menu m on p2.MenuID=m.id
where p.ID = '{Req.UserID}'
AND p2.MenuID = {data.MenuID}
;
");
            }
            */
            bool result = Convert.ToInt32(ExecuteNonQuery(CommandType.Text, SbSql.ToString(), new SQLParameterCollection())) > 0;
            return result;
        }

        public IList<Quality_Position> Get_Position_Head(string Position)
        {
            StringBuilder SbSql = new StringBuilder();

            SbSql.Append($@"
        select Position = ID
                ,Description
                ,IsAdmin
                ,Junk
        from Quality_Position
        where ID='{Position}'

        ");

            return ExecuteList<Quality_Position>(CommandType.Text, SbSql.ToString(), new SQLParameterCollection());
        }

        public IList<Module_Detail> Get_Position_Detail(string Position)
        {
            StringBuilder SbSql = new StringBuilder();

            SQLParameterCollection objParameter = new SQLParameterCollection();

            if (string.IsNullOrEmpty(Position))
            {
                // Click New時要帶出全部
                SbSql.Append($@"
select   MenuID = m.ID
	    ,m.ModuleName
	    ,m.FunctionName
	    , Used = Cast(1 as bit)
from Quality_Menu m
where m.Junk = 0
order by m.ModuleSeq, m.FunctionSeq


");
            }
            else
            {
                objParameter.Add("@ID", DbType.String, Position);

                SbSql.Append($@"
select   MenuID = m.ID
	    ,m.ModuleName
	    ,m.FunctionName
	    ,p2.Used
from Quality_Menu m
inner join Quality_Pass2 p2 ON p2.MenuID = m.id
inner join Quality_Position p on p.ID=p2.PositionID
where p.ID = @ID
order by m.ModuleSeq, m.FunctionSeq

");
            }

            return ExecuteList<Module_Detail>(CommandType.Text, SbSql.ToString(), objParameter);
        }

        public bool Update_Position_Detail(Quality_Position_Request Req)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            objParameter.Add("@ID", DbType.String, Req.Position);
            objParameter.Add("@Description", DbType.String, Req.Description);
            objParameter.Add("@IsAdmin", DbType.Boolean, Req.IsAdmin);
            objParameter.Add("@Junk", DbType.Boolean, Req.Junk);

            SbSql.Append($@"
UPDATE Quality_Position
SET Description = @Description
    ,IsAdmin = @IsAdmin
    ,Junk = @Junk
WHERE ID=@ID 
;
");

            foreach (var data in Req.DataList)
            {
                SbSql.Append($@"
UPDATE p2
SET p2.Used='{(data.Used ? "1" : "0")}'
from Quality_Pass2 p2
inner join Quality_Position p on p.ID=p2.PositionID 
inner join Quality_Menu m ON p2.MenuID = m.id
where p.ID=@ID AND m.ID='{data.MenuID}'
;
");
            }
            bool result = Convert.ToInt32(ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter)) > 0;
            return result;
        }
       
        public int Check_Position_Exists(Quality_Position_Request Req)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            objParameter.Add("@ID", DbType.String, Req.Position);

            SbSql.Append($@"
        SELECT COUNT(1)
        FROM Quality_Position
        WHERE ID = @ID
        ");

            return Convert.ToInt32(ExecuteScalar(CommandType.Text, SbSql.ToString(), objParameter));
        }
        
        public bool Create_Position_Detail(Quality_Position_Request Req)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            objParameter.Add("@ID", DbType.String, Req.Position);
            objParameter.Add("@Description", DbType.String, Req.Description);
            objParameter.Add("@IsAdmin", DbType.Boolean, Req.IsAdmin);
            objParameter.Add("@Junk", DbType.Boolean, Req.Junk);

            SbSql.Append($@"
        INSERT INTO dbo.Quality_Position
                   (ID
                   ,Description
                   ,IsAdmin
                   ,Junk)
             VALUES
                   (@ID
                   ,@Description
                   ,@IsAdmin
                   ,@Junk)
        ;
        ");

            foreach (var data in Req.DataList)
            {
                SbSql.Append($@"
        INSERT INTO dbo.Quality_Pass2
                   (PositionID
                   ,MenuID
                   ,Used)
             VALUES
                   (@ID
                   ,{data.MenuID}
                   ,{(data.Used ? "1" : "0")})
        ;
        ");
            }
            bool result = Convert.ToInt32(ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter)) > 0;
            return result;
        }
        
        public IList<UserList_Browse> GetAllUser(string FactoryID)
        {
            StringBuilder SbSql = new StringBuilder();

            SbSql.Append($@"
select DISTINCT 
[Select] = Cast(0 as bit)
, UserID = p.ID
,p.Name
,Position = ''--po.ID
from Production.dbo.Pass1 p 
left join Quality_Pass1 p1 ON p.ID=p1.ID 
WHERE 1=1
{(string.IsNullOrEmpty(FactoryID) ? "" : $"AND p.Factory='{FactoryID}' ")}
AND NOT EXISTS
(
	SELECT 1 
	FROM Quality_Pass1 p1
	WHERE p.ID=p1.ID
) 
");

            return ExecuteList<UserList_Browse>(CommandType.Text, SbSql.ToString(), new SQLParameterCollection());
        }
       
        public bool ImportUsers(List<UserList_Browse> DataList)
        {
            StringBuilder SbSql = new StringBuilder();

            foreach (var data in DataList)
            {
                SbSql.Append($@"
INSERT INTO dbo.Quality_Pass1
           (ID
           ,Position)
     VALUES
           ('{data.UserID}'
           ,'{data.Position}'
            )
;
");
            }

            bool result = Convert.ToInt32(ExecuteNonQuery(CommandType.Text, SbSql.ToString(), new SQLParameterCollection())) > 0;
            return result;
        }
        
        public IList<SelectListItem> GetPositionList(string FactoryID)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();

            SbSql.Append($@"

/*取得Position下拉選單資料來源*/

SELECT [Text] = ID, [Value]= ID
FROM Quality_Position
WHERE FactoryID='{FactoryID}'
");


            return ExecuteList<SelectListItem>(CommandType.Text, SbSql.ToString(), objParameter);
        }
       
        public IList<Quality_Position> Get_Position_Browse(string FactoryID)
        {
            StringBuilder SbSql = new StringBuilder();

            SbSql.Append($@"
select Position = ID
	,IsAdmin
	,Description
	,Junk
from Quality_Position 
where 1=1
");

            return ExecuteList<Quality_Position>(CommandType.Text, SbSql.ToString(), new SQLParameterCollection());
        }
    }
}
