IF EXISTS(SELECT name FROM sysobjects                
      WHERE name = 'VID_VW_FE_NotaCredito_R_EXTRA' AND type = 'V')
   DROP VIEW VID_VW_FE_NotaCredito_R_EXTRA
GO--                                                 

CREATE
 VIEW [dbo].[VID_VW_FE_NotaCredito_R_EXTRA]
AS
	SELECT 0																									[Folio_Sii]
		,''																										[TpoDocRef]
		,''																										[FolioRef]
		,''																										[FchRef]
		,''																										[CodRef]
		,''																										[RazonRef]
		,''																										[ObjType]
		,0																										[DocEntry]
GO
