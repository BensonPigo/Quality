CREATE PROCEDURE [dbo].[InsertGarmentTestDetailShrinkage]
    @ID varchar(50),
    @No int,
    @Location varchar(10),
    @Type varchar(50),
    @BeforeWash varchar(50),
    @SizeSpec varchar(50),
    @AfterWash1 varchar(50),
    @Shrinkage1 varchar(50),
    @AfterWash2 varchar(50),
    @Shrinkage2 varchar(50),
    @AfterWash3 varchar(50),
    @Shrinkage3 varchar(50),
    @Seq int
AS
BEGIN
    INSERT INTO [GarmentTest_Detail_Shrinkage](
        ID, No, Location, [Type], BeforeWash, SizeSpec,
        AfterWash1, Shrinkage1, AfterWash2, Shrinkage2,
        AfterWash3, Shrinkage3, Seq)
    VALUES(
        @ID, @No, @Location, @Type, @BeforeWash, @SizeSpec,
        @AfterWash1, @Shrinkage1, @AfterWash2, @Shrinkage2,
        @AfterWash3, @Shrinkage3, @Seq);
END
GO
