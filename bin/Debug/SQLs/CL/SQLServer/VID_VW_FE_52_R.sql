IF EXISTS(SELECT name FROM sysobjects                
      WHERE name = 'VID_VW_FE_52_R' AND type = 'V')
   DROP VIEW VID_VW_FE_52_R
GO--                                                 

CREATE VIEW [dbo].[VID_VW_FE_52_R]
AS
	
	SELECT
		 T0.FolioNum												[Folio_Sii]
		,'801'														[TpoDocRef]
		,LEFT(T0.NumAtCard, 18)										[FolioRef]
		,REPLACE(CONVERT(CHAR(10),T0.TaxDate,102),'.','-')			[FchRef]
		,0															[CodRef]
		,''															[RazonRef]
		,T0.DocEntry												[DocEntry]
		,T0.ObjType													[ObjType]
	FROM ODLN T0
	JOIN NNM1 N0 ON N0.Series = T0.Series
	, [@VID_FEPARAM] PA0
	WHERE 1 = 1
		AND ((ISNULL(T0.FolioNum, 0) <> 0 AND ISNULL(PA0.U_Distrib,'N') = 'Y') OR (ISNULL(T0.FolioNum, 0) = 0 AND ISNULL(PA0.U_FPortal,'N') = 'Y') OR (ISNULL(T0.FolioNum, 0) <> 0 AND ISNULL(PA0.U_FPortal,'N') = 'N' AND ISNULL(PA0.U_Distrib,'N') = 'N'))
		AND ISNULL(T0.NumAtCard, '') <> ''
		AND T0.DocSubType = '--'
		AND UPPER(LEFT(N0.BeginStr, 1)) = 'E'
GO
