IF EXISTS(SELECT name FROM sysobjects
      WHERE name = 'VID_SP_FE_LibroDiario' AND type = 'P')
   DROP PROCEDURE VID_SP_FE_LibroDiario
GO--
--EXEC VID_SP_FE_LibroDiario '132' --'2016-05'
CREATE PROCEDURE [dbo].[VID_SP_FE_LibroDiario]
 @Periodo VARCHAR(10)
--WITH ENCRYPTION
AS
BEGIN
	SELECT REPLACE(A0.TaxIdNum,'.','') [Identificacion/RutContribuyente]
		  ,REPLACE(CONVERT(CHAR(7), T0.F_RefDate,102),'.','-') [Identificacion/PeriodoTributario/Inicial]
		  ,REPLACE(CONVERT(CHAR(7), T0.F_RefDate,102),'.','-') [Identificacion/PeriodoTributario/Final]
		  ,REPLACE(CONVERT(CHAR(10), J0.RefDate, 102),'.','-')	[RegistroDiario/FechaContable]
		  ,CAST(SUM(CASE WHEN J1.Line_ID = 0 THEN 1 ELSE 0 END) AS VARCHAR(20))	[RegistroDiario/CantidadComprobantes]
		  ,CAST(COUNT(*) AS VARCHAR(20)) [RegistroDiario/CantidadMovimientos]
		  ,LTRIM(STR(SUM(J1.Debit),18,0)) [RegistroDiario/SumaValorComprobante]
		  ,(SELECT CAST(COUNT(*) AS VARCHAR(20)) FROM OJDT WHERE RefDate BETWEEN T0.F_RefDate AND T0.T_RefDate) [Cierre/CantidadComprobantes]
		  ,(SELECT CAST(COUNT(*) AS VARCHAR(20)) FROM OJDT A JOIN JDT1 B ON B.TransId = A.TransId WHERE A.RefDate BETWEEN T0.F_RefDate AND T0.T_RefDate) [Cierre/CantidadMovimientos]
		  ,(SELECT LTRIM(STR(SUM(B.Debit),18,0)) FROM OJDT A JOIN JDT1 B ON B.TransId = A.TransId WHERE A.RefDate BETWEEN T0.F_RefDate AND T0.T_RefDate) [Cierre/SumaValorComprobante]
		  ,(SELECT LTRIM(STR(SUM(B.Debit),18,0)) FROM OJDT A JOIN JDT1 B ON B.TransId = A.TransId JOIN OFPR C ON A.FinncPriod = C.AbsEntry WHERE A.RefDate <= T0.T_RefDate AND C.Category = T0.Category) [Cierre/ValorAcumulado]
	  FROM OFPR T0
	  JOIN OJDT J0 ON J0.RefDate BETWEEN T0.F_RefDate AND T0.T_RefDate
	  JOIN JDT1 J1 ON J1.TransId = J0.TransId
	  , OADM A0
	 WHERE T0.AbsEntry = @Periodo
	   AND J0.TransType NOT IN ('-2','-3')
	 GROUP BY
		   A0.TaxIdNum
		  ,REPLACE(CONVERT(CHAR(10), J0.RefDate, 102),'.','-')
		  ,T0.F_RefDate
		  ,T0.T_RefDate
		  ,T0.Category
END