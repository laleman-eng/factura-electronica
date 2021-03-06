--DROP VIEW VID_VW_FE_52_D
CREATE VIEW VID_VW_FE_52_D
AS
	SELECT
		 T0."FolioNum"																																		"FolioNum"
		,CASE
			WHEN T1."VatSum" = 0 THEN 1
			ELSE 2
		 END																																				"Indicador_Exento"
		,CASE T0."DocType" WHEN 'S' THEN 'Servicio'
								  ELSE LEFT(T1."ItemCode", 50)
		 END																																				"ItemCode"
		,VID_FN_FE_LimpiaCaracteres(LEFT(T1."Dscription", 80))																								"Dscription"
		,VID_FN_FE_LimpiaCaracteres(LEFT(T1."Dscription", 250))																								"Dscription_Larga"
		,LEFT(IFNULL(T1."unitMsr", ''),4)																													"DET_UNIDAD_MEDIDA"
		,CASE WHEN T0."DocType" = 'S' THEN 1
									ELSE T1."Quantity"
		 END																																				"Quantity"
		,CASE IFNULL(T1."Currency", 'CLP') WHEN 'CLP' THEN T1."PriceBefDi"
										 WHEN '$'	THEN T1."PriceBefDi"
													ELSE ROUND(T1."PriceBefDi" * T1."Rate", 4)							  
		 END																																				"Price"
		,CASE WHEN T1."DiscPrcnt" >= 0 THEN T1."DiscPrcnt" ELSE 0.0 END																						"DiscPrcnt"
		,CASE WHEN IFNULL(T1."DiscPrcnt",0.0) <= 0.0 THEN 0.0
		 ELSE 
		 ROUND((CASE WHEN IFNULL(T1."Currency", 'CLP') IN ('CLP', '$') THEN T1."PriceBefDi"
																	 ELSE ROUND(T1."PriceBefDi" * T1."Rate", 2)
				END) * CASE WHEN T0."DocType" = 'S' THEN 1
												  ELSE T1."Quantity"
					   END, 0) - T1."LineTotal"	 
		 END 																																				"DiscSum"
		,CASE WHEN IFNULL(T1."DiscPrcnt",0.0) >= 0.0 THEN 0.0
		 ELSE T1."LineTotal" -
		 ROUND((CASE WHEN IFNULL(T1."Currency", 'CLP') IN ('CLP', '$') THEN T1."PriceBefDi"
																	 ELSE ROUND(T1."PriceBefDi" * T1."Rate", 2)
				END) * CASE WHEN T0."DocType" = 'S' THEN 1
												  ELSE T1."Quantity"
					   END, 0)
		 END 																																				"RecargoMonto"
		,ROUND(T1."LineTotal", 0)																															"LineTotal"
		,IFNULL(F0."U_CodImpto", '')																														"CodImpAdic"
		,LEFT(IFNULL(U0."U_NAME", ''), 30)																													"Usuario"
		,IFNULL(T4."TaxSum", 0.0)																															"MontoImptoAdic"
		,IFNULL(F0."U_Porc",0.0)																															"PorcImptoAdic"
		,T1."Rate"																																			"Rate"
		,T1."Currency"																																		"Currency"
		,T1."VisOrder"																																		"LineaOrden"
		,1																																					"LineaOrden2"
		,T0."ObjType"																																		"ObjType"
		,T0."DocEntry"																																		"DocEntry"
	FROM	  "OWTR"		   T0
		 JOIN "WTR1"		   T1 ON T1."DocEntry"	= T0."DocEntry"
		 JOIN "OUSR"		   U0 ON U0."INTERNAL_K"= T0."UserSign"
		 JOIN "NNM1"		   N0 ON N0."Series"	= T0."Series"
								 AND N0."ObjectCode"= T0."ObjType"
	LEFT JOIN "WTR4"		   T4 ON T4."DocEntry"	= T0."DocEntry"
								 AND T4."LineNum"	= T1."LineNum"
								 AND T4."StaCode" NOT IN ('IVA', 'IVA_EXE')
	LEFT JOIN "@VID_FEIMPADIC" F0 ON F0."Code"		= T4."StaCode"
	WHERE 1 = 1
		--AND IFNULL(T0."FolioNum", 0) <> 0
		AND UPPER(LEFT(N0."BeginStr", 1)) = 'E'
		
	UNION ALL
	
	SELECT
		 T0."FolioNum"																																		"FolioNum"
		,2																																					"Indicador_Exento"
		,'Texto'																																			"ItemCode"
		,VID_FN_FE_LimpiaCaracteres(LEFT(CAST(T10."LineText" AS VARCHAR(254)), 80))																			"Dscription"
		,VID_FN_FE_LimpiaCaracteres(LEFT(CAST(T10."LineText" AS VARCHAR(254)), 250))																		"Dscription_Larga"
		,''																																					"DET_UNIDAD_MEDIDA"
		,1.0																																				"Quantity"
		,0.0																																				"Price"
		,0.0																																				"DiscPrcnt"
		,0.0																																				"DiscSum"
		,0.0																																				"RecargoMonto"
		,0.0																																				"LineTotal"
		,''																																					"CodImpAdic"
		,LEFT(IFNULL(U0."U_NAME", ''), 30)																													"Usuario"
		,0.0																																				"MontoImptoAdic"
		,0.0																																				"PorcImptoAdic"
		,0.0																																				"Rate"
		,''																																					"Currency"
		,T10."AftLineNum"																																	"LineaOrden"
		,2																																					"LineaOrden2"
		,T0."ObjType"																																		"ObjType"
		,T0."DocEntry"																																		"DocEntry"
	FROM "OWTR"  T0
	JOIN "WTR10" T10 ON T10."DocEntry"	= T0."DocEntry"
	JOIN "OUSR"  U0  ON U0."INTERNAL_K"	= T0."UserSign"
	JOIN "NNM1"  N0  ON N0."Series"		= T0."Series"
				  AND N0."ObjectCode"	= T0."ObjType"
	WHERE 1 = 1
		--AND IFNULL(T0."FolioNum", 0) <> 0
		AND UPPER(LEFT(N0."BeginStr", 1)) = 'E'
	
	UNION ALL
	
	SELECT
		 T0."FolioNum"																																		"FolioNum"
		,CASE
			WHEN T1."VatSum" = 0 THEN 1
			ELSE 2
		 END																																				"Indicador_Exento"
		,CASE T0."DocType" WHEN 'S' THEN 'Servicio'
								  ELSE LEFT(T1."ItemCode", 50)
		 END																																				"ItemCode"
		,VID_FN_FE_LimpiaCaracteres(LEFT(T1."Dscription", 80))																								"Dscription"
		,VID_FN_FE_LimpiaCaracteres(LEFT(T1."Dscription", 250))																								"Dscription_Larga"
		,LEFT(IFNULL(T1."unitMsr", ''),4)																													"DET_UNIDAD_MEDIDA"
		,CASE WHEN T0."DocType" = 'S' THEN 1
									ELSE T1."Quantity"
		 END																																				"Quantity"
		,CASE IFNULL(T1."Currency", 'CLP') WHEN 'CLP' THEN T1."PriceBefDi"
										 WHEN '$'	THEN T1."PriceBefDi"
													ELSE ROUND(T1."PriceBefDi" * T1."Rate", 4)							  
		 END																																				"Price"
		,CASE WHEN T1."DiscPrcnt" >= 0.0 THEN T1."DiscPrcnt" ELSE 0.0 END																					"DiscPrcnt"
		,CASE WHEN IFNULL(T1."DiscPrcnt",0.0) <= 0.0 THEN 0.0
		ELSE
		ROUND((CASE WHEN IFNULL(T1."Currency", 'CLP') IN ('CLP', '$') THEN T1."PriceBefDi"
																	 ELSE ROUND(T1."PriceBefDi" * T1."Rate", 2)
				END) * CASE WHEN T0."DocType" = 'S' THEN 1
												  ELSE T1."Quantity"
					   END, 0) - T1."LineTotal"
		 END 																																				"DiscSum"
		,CASE WHEN IFNULL(T1."DiscPrcnt",0.0) >= 0.0 THEN 0.0
		ELSE T1."LineTotal" -
		ROUND((CASE WHEN IFNULL(T1."Currency", 'CLP') IN ('CLP', '$') THEN T1."PriceBefDi"
																	 ELSE ROUND(T1."PriceBefDi" * T1."Rate", 2)
				END) * CASE WHEN T0."DocType" = 'S' THEN 1
												  ELSE T1."Quantity"
					   END, 0)
		 END 																																				"RecargoMonto"
		,ROUND(T1."LineTotal", 0)																															"LineTotal"
		,IFNULL(F0."U_CodImpto", '')																														"CodImpAdic"
		,LEFT(IFNULL(U0."U_NAME", ''), 30)																													"Usuario"
		,IFNULL(T4."TaxSum", 0.0)																															"MontoImptoAdic"
		,IFNULL(F0."U_Porc",0.0)																															"PorcImptoAdic"
		,T1."Rate"																																			"Rate"
		,T1."Currency"																																		"Currency"
		,T1."VisOrder"																																		"LineaOrden"
		,1																																					"LineaOrden2"
		,T0."ObjType"																																		"ObjType"
		,T0."DocEntry"																																		"DocEntry"
	FROM	  "ORPD"			   T0
		 JOIN "RPD1"			   T1 ON T1."DocEntry"	= T0."DocEntry"
		 JOIN "OUSR"			   U0 ON U0."INTERNAL_K"= T0."UserSign"
		 JOIN "NNM1"			   N0 ON N0."Series"	= T0."Series"
								 AND N0."ObjectCode"= T0."ObjType"
	LEFT JOIN "RPD4"			   T4 ON T4."DocEntry"	= T0."DocEntry"
								 AND T4."LineNum"	= T1."LineNum"
								 AND T4."StaCode" NOT IN ('IVA', 'IVA_EXE')
	LEFT JOIN "@VID_FEIMPADIC" F0 ON F0."Code"		= T4."StaCode"
	WHERE 1 = 1
		--AND IFNULL(T0."FolioNum", 0) <> 0
		AND UPPER(LEFT(N0."BeginStr", 1)) = 'E'
		
	UNION ALL
	
	SELECT
		 T0."FolioNum"																																		"FolioNum"
		,2																																					"Indicador_Exento"
		,'Texto'																																			"ItemCode"
		,VID_FN_FE_LimpiaCaracteres(LEFT(CAST(T10."LineText" AS VARCHAR(254)), 80))																			"Dscription"
		,VID_FN_FE_LimpiaCaracteres(LEFT(CAST(T10."LineText" AS VARCHAR(254)), 250))																		"Dscription_Larga"
		,''																																					"DET_UNIDAD_MEDIDA"
		,1.0																																				"Quantity"
		,0.0																																				"Price"
		,0.0																																				"DiscPrcnt"
		,0.0																																				"DiscSum"
		,0.0																																				"RecargoMonto"
		,0.0																																				"LineTotal"
		,''																																					"CodImpAdic"
		,LEFT(IFNULL(U0."U_NAME", ''), 30)																													"Usuario"
		,0.0																																				"MontoImptoAdic"
		,0.0																																				"PorcImptoAdic"
		,0.0																																				"Rate"
		,''																																					"Currency"
		,T10."AftLineNum"																																	"LineaOrden"
		,2																																					"LineaOrden2"
		,T0."ObjType"																																		"ObjType"
		,T0."DocEntry"																																		"DocEntry"
	FROM "ORPD"  T0
	JOIN "RPD10" T10 ON T10."DocEntry"	= T0."DocEntry"
	JOIN "OUSR"  U0  ON U0."INTERNAL_K"	= T0."UserSign"
	JOIN "NNM1"  N0  ON N0."Series"		= T0."Series"
				  AND N0."ObjectCode"	= T0."ObjType"
	WHERE 1 = 1
		--AND IFNULL(T0."FolioNum", 0) <> 0
		AND UPPER(LEFT(N0."BeginStr", 1)) = 'E'

	UNION ALL
	
	SELECT
		 T0."FolioNum"																																		"FolioNum"
		,CASE
			WHEN T1."VatSum" = 0 THEN 1
			ELSE 2
		 END																																				"Indicador_Exento"
		,CASE T0."DocType" WHEN 'S' THEN 'Servicio'
								  ELSE LEFT(T1."ItemCode", 50)
		 END																																				"ItemCode"
		,VID_FN_FE_LimpiaCaracteres(LEFT(T1."Dscription", 80))																								"Dscription"
		,VID_FN_FE_LimpiaCaracteres(LEFT(T1."Dscription", 250))																								"Dscription_Larga"
		,LEFT(IFNULL(T1."unitMsr", ''),4)																													"DET_UNIDAD_MEDIDA"
		,CASE WHEN T0."DocType" = 'S' THEN 1
									ELSE T1."Quantity"
		 END																																				"Quantity"
		,CASE IFNULL(T1."Currency", 'CLP') WHEN 'CLP' THEN T1."PriceBefDi"
										 WHEN '$'   THEN T1."PriceBefDi"
													ELSE ROUND(T1."PriceBefDi" * T1."Rate", 4)							  
		 END																																				"Price"
		,CASE WHEN T1."DiscPrcnt" >= 0 THEN T1."DiscPrcnt" ELSE 0.0 END																						"DiscPrcnt"
		,CASE WHEN IFNULL(T1."DiscPrcnt",0.0) <= 0.0 THEN 0.0
		ELSE
		ROUND((CASE WHEN IFNULL(T1."Currency", 'CLP') IN ('CLP', '$') THEN T1."PriceBefDi"
																	 ELSE ROUND(T1."PriceBefDi" * T1."Rate", 2)
				END) * CASE WHEN T0."DocType" = 'S' THEN 1
												  ELSE T1."Quantity"
					   END, 0) - T1."LineTotal"
		 END																																				"DiscSum"
		,CASE WHEN IFNULL(T1."DiscPrcnt",0.0) >= 0.0 THEN 0.0
		ELSE T1."LineTotal" -
		ROUND((CASE WHEN IFNULL(T1."Currency", 'CLP') IN ('CLP', '$') THEN T1."PriceBefDi"
																	 ELSE ROUND(T1."PriceBefDi" * T1."Rate", 2)
				END) * CASE WHEN T0."DocType" = 'S' THEN 1
												  ELSE T1."Quantity"
					   END, 0)
		 END																																				"RecargoMonto"
		,ROUND(T1."LineTotal", 0)																															"LineTotal"
		,IFNULL(F0."U_CodImpto", '')																														"CodImpAdic"
		,LEFT(IFNULL(U0."U_NAME", ''), 30)																													"Usuario"
		,IFNULL(T4."TaxSum", 0.0)																															"MontoImptoAdic"
		,IFNULL(F0."U_Porc",0.0)																															"PorcImptoAdic"
		,T1."Rate"																																			"Rate"
		,T1."Currency"																																		"Currency"
		,T1."VisOrder"																																		"LineaOrden"
		,1																																					"LineaOrden2"
		,T0."ObjType"																																		"ObjType"
		,T0."DocEntry"																																		"DocEntry"
	FROM	  "ODLN"			   T0
		 JOIN "DLN1"			   T1 ON T1."DocEntry"	= T0."DocEntry"
		 JOIN "OUSR"			   U0 ON U0."INTERNAL_K"= T0."UserSign"
		 JOIN "NNM1"			   N0 ON N0."Series"	= T0."Series"
								 AND N0."ObjectCode"	= T0."ObjType"
	LEFT JOIN "DLN4"			   T4 ON T4."DocEntry"	= T0."DocEntry"
								 AND T4."LineNum"		= T1."LineNum"
								 AND T4."StaCode" NOT IN ('IVA', 'IVA_EXE')
	LEFT JOIN "@VID_FEIMPADIC" F0 ON F0."Code"	= T4."StaCode"
	WHERE 1 = 1
		--AND IFNULL(T0."FolioNum", 0) <> 0
		AND UPPER(LEFT(N0."BeginStr", 1)) = 'E'
		
	UNION ALL
	
	SELECT
		 T0."FolioNum"																																		"FolioNum"
		,2																																					"Indicador_Exento"
		,'Texto'																																			"ItemCode"
		,VID_FN_FE_LimpiaCaracteres(LEFT(CAST(T10."LineText" AS VARCHAR(254)), 80))																			"Dscription"
		,VID_FN_FE_LimpiaCaracteres(LEFT(CAST(T10."LineText" AS VARCHAR(254)), 250))																		"Dscription_Larga"
		,''																																					"DET_UNIDAD_MEDIDA"
		,1.0																																				"Quantity"
		,0.0																																				"Price"
		,0.0																																				"DiscPrcnt"
		,0.0																																				"DiscSum"
		,0.0																																				"RecargoMonto"
		,0.																																					"LineTotal"
		,''																																					"CodImpAdic"
		,LEFT(IFNULL(U0."U_NAME", ''), 30)																													"Usuario"
		,0.0																																				"MontoImptoAdic"
		,0.0																																				"PorcImptoAdic"
		,0.0																																				"Rate"
		,''																																					"Currency"
		,T10."AftLineNum"																																	"LineaOrden"
		,2																																					"LineaOrden2"
		,T0."ObjType"																																		"ObjType"
		,T0."DocEntry"																																		"DocEntry"
	FROM "ODLN"  T0
	JOIN "DLN10" T10 ON T10."DocEntry"	= T0."DocEntry"
	JOIN "OUSR"  U0  ON U0."INTERNAL_K"	= T0."UserSign"
	JOIN "NNM1"  N0  ON N0."Series"		= T0."Series"
				  AND N0."ObjectCode"	= T0."ObjType"
	WHERE 1 = 1
		---AND IFNULL(T0."FolioNum", 0) <> 0
		AND UPPER(LEFT(N0."BeginStr", 1)) = 'E';
