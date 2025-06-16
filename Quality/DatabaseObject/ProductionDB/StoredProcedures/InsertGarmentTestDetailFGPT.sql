CREATE PROCEDURE [dbo].[InsertGarmentTestDetailFGPT]
    @ID varchar(50),
    @No int,
    @Location varchar(10),
    @Type varchar(50),
    @TestName varchar(50),
    @TestDetail varchar(50),
    @Criteria int,
    @TestResult varchar(50),
    @TestUnit varchar(50),
    @Seq int,
    @TypeSelection_VersionID int,
    @TypeSelection_Seq int
AS
BEGIN
    INSERT INTO [GarmentTest_Detail_FGPT](
        ID, No, Location, [Type], TestName, TestDetail,
        Criteria, TestResult, TestUnit, Seq,
        TypeSelection_VersionID, TypeSelection_Seq)
    VALUES(
        @ID, @No, @Location, @Type, @TestName, @TestDetail,
        @Criteria, @TestResult, @TestUnit, @Seq,
        @TypeSelection_VersionID, @TypeSelection_Seq);
END
GO
