--DROP PROCEDURE VID_SP_FE_Boleta_D;
CREATE PROCEDURE VID_SP_FE_Boleta_D
(
     IN DocEntry		Integer
    ,IN TipoDoc			VarChar(10)
    ,IN ObjType			VarChar(10)
)
LANGUAGE SqlScript
AS
BEGIN
	SELECT ROW_NUMBER() OVER(ORDER BY T0."LineaOrden", T0."LineaOrden2")								"NroLinDet" 
		,T0."DiscSum"																			"DescuentoMonto"
		,T0."DiscPrcnt"																			"DescuentoPct"
		,T0."Indicador_Exento"																	"IndExe"
		,T0."LineTotal"																			"MontoItem"
		,T0."ItemCode"																			"NmbItem"
		,T0."Price"																				"PrcItem"
		,0.0																					"PrcRef"
		,T0."Quantity"																			"QtyItem"
		,0.0																					"QtyRef"
		,IFNULL(T0."RecargoMonto",0.0)															"RecargoMonto"
		,0.0																					"RecargoPct"
		,T0."DET_UNIDAD_MEDIDA"																	"UnmdItem"
		,T0."ItemCode"																			"VlrCodigo"
		,T0."Dscription_Larga"																	"DscItem"
		
		,T0."CodImpAdic"																		"CodImpAdic"
		,T0."MontoImptoAdic"																	"MntImpAdic"
	FROM VID_VW_FE_OINV_D		 T0
	WHERE 1 = 1
		AND T0."DocEntry" = :DocEntry
		AND T0."ObjType" = :ObjType;
END;
