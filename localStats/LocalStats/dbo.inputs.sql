CREATE TABLE [dbo].[inputs] (
    [ID]     INT            IDENTITY (1, 1) NOT NULL,
    [local]  NVARCHAR (MAX) NOT NULL,
    [utente] INT            NOT NULL,
    [qty]    INT            NOT NULL,
    CONSTRAINT [PK_dbo.inputs] PRIMARY KEY CLUSTERED ([ID] ASC)
);

