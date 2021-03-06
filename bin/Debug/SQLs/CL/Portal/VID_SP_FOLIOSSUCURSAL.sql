IF EXISTS(SELECT name FROM sysobjects
      WHERE name = 'VID_SP_FOLIOSSUCURSAL' AND type = 'P')
	DROP PROCEDURE [dbo].[VID_SP_FOLIOSSUCURSAL]
GO--

CREATE PROCEDURE [dbo].[VID_SP_FOLIOSSUCURSAL](    
	@DocEntry numeric(18,0),
	@TipoDoc  varchar(10),
	@CAFDesde int,
	@CAFHasta int,
	@CAFFecha varchar(10),
	@Sucursal nvarchar(30),
	@Desde	   numeric(18, 0),
	@Hasta    numeric(18, 0),
	@CantAsig int
)      
as 
BEGIN   
	Declare    
	   @Ws_Mensaje char(200)    
	    
	   Select @Ws_Mensaje = ' '    
	    
	set dateformat dmy    
	  
	If Not Exists(Select 1 
				  From VID_FEASIGFOLSUC
				  Where DocEntry = @DocEntry 
					And TipoDoc = @TipoDoc
					And CAFDesde = @CAFDesde
					And CAFHasta = @CAFHasta
					And CAFFecha = @CAFFecha
					And Sucursal = @Sucursal
					And Desde = @Desde
					And Hasta = @Hasta
					And CantAsig = @CantAsig )     
	   Begin    
	  INSERT INTO  VID_FEASIGFOLSUC     
		   ( DocEntry
			,TipoDoc
			,CAFDesde
			,CAFHasta
			,CAFFecha
			,Sucursal
			,Desde
			,Hasta
			,CantAsig
			,CreateDate
		   ) VALUES (
		     @DocEntry
			,@TipoDoc
			,@CAFDesde
			,@CAFHasta
			,@CAFFecha
			,@Sucursal
			,@Desde
			,@Hasta
			,@CantAsig
			,getdate()
		   )       
	 
		 --Select @DYT_ID_TRASPASO = @@IDENTITY  
		 SET @Ws_Mensaje = 'INSERT'   
	   End    
	       
	Else    
	   Begin    
		  --UPDATE  Faet_Erp_Encabezado_Doc     
			 --SET  CAB_NUME_IDENTIFI  = @CAB_NUME_IDENTIFI     
			 --, CAB_FEC_EMISION   = @CAB_FEC_EMISION      
			 --, CAB_EXTRA5 = @CAB_EXTRA5
			 --, CAB_DESC_TOTAL = @CAB_DESC_TOTAL
		  -- Where CAB_EMPRESA   = @CAB_EMPRESA    
		  -- And CAB_DIVISION      = @CAB_DIVISION     
		  -- And CAB_UNIDAD    = @CAB_UNIDAD    
		  -- And CAB_COD_TP_FACTURA  = @CAB_COD_TP_FACTURA    
		  -- And CAB_FOL_DOCTO_INT   = @CAB_FOL_DOCTO_INT     
		  -- And DYT_ID_TRASPASO   = @DYT_ID_TRASPASO   
		  -- And IndTipoLibro   = 'V'  
	   SET @Ws_Mensaje = 'UPDATE'  
	 End    
	    
	    
	 Select @DocEntry 'DocEntry'
			, @TipoDoc    'TipoDoc'
		,@Ws_Mensaje         Mensaje
END
