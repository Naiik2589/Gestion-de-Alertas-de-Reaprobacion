# Gestion-de-Alertas-de-Reaprobacion
Codigo de AppScrib que gestiona la creacion de eventos en el calendar de google y envia correos segun fechas establecidas, las fechas de alertas y reaprobacion se calculan con formulas en la hoja de calculo y luego con esas fechas se automatiza la creacion de los eventos segun periodo.



FORMULAS DE GOOGLE SHEET

CALCULA LA SIGUIENTE FECHA DE REAPROBACION SEGUN PERIODO:
=SI(F2="Trimestral"; 
   SI(DIA(FIN.MES(E2; 3)) < DIA(E2); 
      FIN.MES(E2; 3); 
      FIN.MES(E2; 3) - DIA(FIN.MES(E2; 3)) + DIA(E2));
 SI(F2="Semestral"; 
   SI(DIA(FIN.MES(E2; 6)) < DIA(E2); 
      FIN.MES(E2; 6); 
      FIN.MES(E2; 6) - DIA(FIN.MES(E2; 6)) + DIA(E2));
 SI(F2="Anual"; 
   SI(DIA(FIN.MES(E2; 12)) < DIA(E2); 
      FIN.MES(E2; 12); 
      FIN.MES(E2; 12) - DIA(FIN.MES(E2; 12)) + DIA(E2));
 "Periodo no válido")))


 CALCULA LA SIGUIENTE FECHA DE LA ALERTA (1 MES ANTES DE LA REAPROBACION):
 =SI(DIA(G2) > DIA(FIN.MES(FECHA(AÑO(G2); MES(G2) - 1; 1); 0)); 
   FIN.MES(FECHA(AÑO(G2); MES(G2) - 1; 1); 0); 
   FECHA(AÑO(G2); MES(G2) - 1; DIA(G2)))
