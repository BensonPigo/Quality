CREATE PROCEDURE [dbo].[InsertGarmentTestDetailApperance]
    @ID varchar(50),
    @No int,
    @Type varchar(50),
    @Wash1 varchar(50),
    @Wash2 varchar(50),
    @Wash3 varchar(50),
    @Comment varchar(255),
    @Seq int,
    @Wash4 varchar(50),
    @Wash5 varchar(50)
AS
BEGIN
    INSERT INTO [GarmentTest_Detail_Apperance](
        ID, No, [Type], Wash1, Wash2, Wash3,
        Comment, Seq, Wash4, Wash5)
    VALUES(
        @ID, @No, @Type, @Wash1, @Wash2, @Wash3,
        @Comment, @Seq, @Wash4, @Wash5);
END
GO
