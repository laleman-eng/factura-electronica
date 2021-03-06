IF EXISTS(SELECT name FROM sysobjects
      WHERE name = 'VID_SP_EXISTEFOLIO' AND type = 'P')
	DROP PROCEDURE [dbo].[VID_SP_EXISTEFOLIO]
GO--

CREATE PROCEDURE [dbo].[VID_SP_EXISTEFOLIO](    
	@TipoDoc  varchar(10),
	@FolioNum int
)      
as 
BEGIN   
	Declare    
	   @Ws_Mensaje char(200)    
	    
	   Select @Ws_Mensaje = ' '    
	    
	set dateformat dmy    
	  
	If Not Exists(Select 1 
				  From Faet_Erp_Encabezado_Doc with(nolock)
				  Where CAB_COD_TP_FACTURA  = @TipoDoc
				  And CAB_FOL_DOCTO_INT     = @FolioNum)
	   Begin    
		 SET @Ws_Mensaje = 'N'   
	   End    
	       
	Else    
	   Begin    
	   SET @Ws_Mensaje = 'Y'  
	 End    
	    
	    
	 Select @Ws_Mensaje  Mensaje

END

