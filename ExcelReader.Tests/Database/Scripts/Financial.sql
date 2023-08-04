USE [ExcelReader]
GO

IF OBJECT_ID('[Financial]') IS NOT NULL
BEGIN
	PRINT 'Table [Financial] already exists... dropping...'
	DROP TABLE [Financial]
END
GO

IF OBJECT_ID('[Financial]') IS NULL
BEGIN
	PRINT 'Creating Table [Financial]'

	CREATE TABLE [Financial] (
		[Id] [int] IDENTITY(1,1) NOT NULL,
		[Segment] nvarchar(100) NOT NULL,
		[Country] nvarchar(100) NOT NULL,
		[Units Sold] decimal(24,6),
		[Manufacturing Price] decimal(24,6),
		[Sales] decimal(24,6),
		[Profit] decimal(24,6),
		[Date] date null,
		[Month Number] integer NOT NULL,
		[DateCreated] DATETIME NOT NULL DEFAULT GETDATE(),
		[DateModified] DATETIME NOT NULL DEFAULT GETDATE()
	)
END
GO
