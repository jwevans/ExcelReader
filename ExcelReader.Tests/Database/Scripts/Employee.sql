USE [ExcelReader]
GO

IF OBJECT_ID('[Employee]') IS NOT NULL
BEGIN
	PRINT 'Table [Employee] already exists... dropping...'
	DROP TABLE [Employee]
END
GO

IF OBJECT_ID('[Employee]') IS NULL
BEGIN
	PRINT 'Creating Table [Employee]'

	CREATE TABLE [Employee] (
		[Id] [int] IDENTITY(1,1) NOT NULL,
		[EEID] nvarchar(100) NOT NULL,
		[Job Title] nvarchar(100) NOT NULL,
		[Department] nvarchar(100) NOT NULL,		
		[Hire Date] date null,
		[DateCreated] DATETIME NOT NULL DEFAULT GETDATE(),
		[DateModified] DATETIME NOT NULL DEFAULT GETDATE()
	)
END
GO
