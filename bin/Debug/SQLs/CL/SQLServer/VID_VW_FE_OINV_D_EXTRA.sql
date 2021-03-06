IF EXISTS(SELECT name FROM sysobjects                
      WHERE name = 'VID_VW_FE_OINV_D_EXTRA' AND type = 'V')
   DROP VIEW VID_VW_FE_OINV_D_EXTRA
GO--                                                 

CREATE
 VIEW [dbo].[VID_VW_FE_OINV_D_EXTRA]
AS
	SELECT T0.DocEntry																																		[DocEntry]
		  ,T0.ObjType																																		[ObjType]
		  ,T1.VisOrder																																		[LineaOrden]
		  ,1																																				[LineaOrden2]
		  ,''																																				[DET_EXTRA1]
		  ,''																																				[DET_EXTRA2]
		  ,''																																				[DET_EXTRA3]
		  ,''																																				[DET_EXTRA4]
		  ,''																																				[DET_EXTRA5]
		
	FROM	  OINV			   T0
		 JOIN INV1			   T1 ON T1.DocEntry	= T0.DocEntry
		 JOIN NNM1			   N0 ON N0.Series	 	= T0.Series
		 , [@VID_FEPARAM] PA0
	WHERE 1 = 1
		AND ((ISNULL(T0.FolioNum, 0) <> 0 AND ISNULL(PA0.U_Distrib,'N') = 'Y') OR (ISNULL(T0.FolioNum, 0) = 0 AND ISNULL(PA0.U_FPortal,'N') = 'Y') OR (ISNULL(T0.FolioNum, 0) <> 0 AND ISNULL(PA0.U_FPortal,'N') = 'N' AND ISNULL(PA0.U_Distrib,'N') = 'N'))
		AND UPPER(LEFT(N0.BeginStr, 1)) = 'E'

GO
