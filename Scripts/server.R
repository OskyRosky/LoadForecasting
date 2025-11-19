############################
#   Contenido del server   # 
############################


server <- function(input, output, session) {
  
  
  ##############################################
  #         Análisis de los ingresos           #
  ##############################################
  
  # Tabla con ingresos
  
  output$conceptos.1 = renderReactable({
    
    reactable(conceptos)
  })
  
  ##############################################
  #        Presentación de los ingresos        #
  ##############################################
  
  
  output$picture.ingresos <- renderImage({
    return(list(src = "C:/Users/oscar/Desktop/Ingresos - SIGAF/Dashboard --- server --- fiscontinua/Scripts y tablas/www/ingresos.png",contentType = "image/png",alt = "R"))
  }, 
  deleteFile = FALSE
  )
  

  ##############################################
  #         Indicadores de los ingresos        #
  ##############################################
  
  output$index1 <- renderValueBox({
    
    ind1 <- round(indicador.1/Millones,1)
    
    ind1 <-  ind1 %>% 
      mutate( `Ingresos`= number(ind1$`Ingresos`, accuracy = .1, big.mark = ".", decimal.mark = ",")
              
      )
    
    valueBox(value= ind1, subtitle=paste0("(en millones)"),
             icon = icon("money"), color = "purple", width = 3)
  })
  
  output$index2 <- renderValueBox({
    
    ind2 <- round(indicador.2,1)
    
    ind2 <-  ind2 %>% 
      mutate( `Ingresos`= number(ind2$`Ingresos`, accuracy = .1, big.mark = ".", decimal.mark = ",")
              
      )
    
    valueBox(value= ind2, subtitle=paste0("(en porcentaje)"),
             icon = icon("money"), color = "green", width = 3)
  })
  
  output$index3 <- renderValueBox({
    
    ind3 <- round(indicador.3,1)
    
    ind3 <-  ind3 %>% 
      mutate( `Ingresos`= number(ind3$`Ingresos`, accuracy = .1, big.mark = ".", decimal.mark = ",")
              
      )
    
    valueBox(value= ind3, subtitle=paste0("(en porcentaje)"),
             icon = icon("money"), color = "green", width = 3)
  })
  
  output$index4 <- renderValueBox({
    
    ind4 <- round(indicador.4,1)
    
    
    valueBox(value= ind4, subtitle=paste0("(en porcentaje)"),
             icon = icon("money"), color = "blue", width = 3)
  })
  
  output$index5 <- renderValueBox({
    
    ind5 <- round(indicador.5,1)
    
    
    valueBox(value= ind5, subtitle=paste0("(en porcentaje)"),
             icon = icon("money"), color = "blue", width = 3)
  })
  
  
  output$index6 <- renderValueBox({
    
    ind6 <- round(indicador.6,1)
    
    
    valueBox(value= ind6, subtitle=paste0("(en porcentaje)"),
             icon = icon("weed"), color = "green", width = 3)
  })
  
  ##############################################
  #   Graficos de los tipos de presupuestos    #
  ##############################################
  
#  output$picture <- renderImage({
#    return(list(src = "C:/Users/oscar/Desktop/Regla fiscal/RF --- REVOLUTION/RF_FINAL/Scripts_tablas/www/ppt-rf.png",contentType = "image/png",alt = "R"))
#  }, 
#  deleteFile = FALSE
#  )
  
  
  ###########################################
  #      Evolucion de los presupuestos      #
  ###########################################
  
  ####################
  #        Anual     #
  ####################
  
  output$HCIA_1 <- renderHighchart({
    
    Ingresos_anual <- tabla_1
    
    Ingresos_anual <- Ingresos_anual %>%
      mutate(
        `Inicial` = round(Inicial/Millones,1),
        `Actual`= round(Actual/Millones,1),
        `Ajustado` = round(Ajustado/Millones,1),
        `Acumulado` = round(Acumulado/Millones,1) #,
      #   `Ajustado especial` = round(`Ajustado especial`/Millones,1)
      ) 
    
    Ingresos_anual  <- highchart() %>% 
      hc_title(text = "",
               margin = 20, align = "center",
               style = list(color = "#129", useHTML = TRUE)) %>% 
      hc_subtitle(text = "",
                  align = "right",
                  style = list(color = "#634", fontWeight = "bold")) %>%
      hc_credits(enabled = TRUE, # add credits
                 text = "" # href = "www.cgr.go.cr"
      ) %>%
      hc_legend(align = "left", verticalAlign = "top",
                layout = "vertical", x = 0, y = 100) %>%
      hc_tooltip(crosshairs = TRUE, backgroundColor = "#FCFFC5",
                 shared = TRUE, borderWidth = 5) %>% 
      hc_exporting(enabled = TRUE) %>% 
      hc_xAxis(categories = Ingresos_anual$Año) %>% 
      hc_add_series(name = "Inicial", data = Ingresos_anual$Inicial) %>% 
      hc_add_series(name = "Actual", data = Ingresos_anual$Actual) %>% 
      hc_add_series(name = "Ajustado", data = Ingresos_anual$Ajustado) %>% 
   #   hc_add_series(name = "Ajustado especial", data = Ingresos_anual$`Ajustado especial`) %>% 
      hc_add_series(name = "Ingresos acumulados", data = Ingresos_anual$Acumulado, color="red") %>% 
      hc_yAxis(title = list(text = "Millones de colones"), labels = list(format = "{value}"))%>% 
      hc_xAxis(title = list(text = "Años") )  %>% 
      
      hc_chart(zoomType = "xy")
    
    Ingresos_anual
    
  })
  
  ####################
  #      Mensual     #
  ####################
  
  output$HCIM_1 <- renderHighchart({
    
    Ingresos_mensual <- tabla_2
    
    Ingresos_mensual <- Ingresos_mensual %>%
      dplyr::mutate(
        Ingresos = round(Ingresos/Millones,1)
      )
    
    Ingresos_mensual  <- highchart() %>% 
      hc_title(text = "",
               margin = 20, align = "center",
               style = list(color = "#129", useHTML = TRUE)) %>% 
      hc_subtitle(text = "",
                  align = "right",
                  style = list(color = "#634", fontWeight = "bold")) %>%
      hc_credits(enabled = TRUE, # add credits
                 text = "" # href = "www.cgr.go.cr"
      ) %>%
      hc_legend(align = "left", verticalAlign = "top",
                layout = "vertical", x = 0, y = 100) %>%
      hc_exporting(enabled = TRUE) %>% 
      hc_xAxis(categories = Ingresos_mensual$Fecha) %>% 
      hc_add_series(name = "Ingresos", data = Ingresos_mensual$Ingresos) %>% 
      hc_yAxis(title = list(text = "Millones de colones"), labels = list(format = "{value}"))%>% 
      hc_xAxis(title = list(text = "Año-mes") )  %>% 
      hc_chart(zoomType = "xy")
    
    Ingresos_mensual
    
  })
  
  
  output$HCIM_2 <- renderHighchart({
    
    Ingresos_mensual_var <- tabla_2_var
    

    Ingresos_mensual_var  <- highchart() %>% 
      hc_title(text = "",
               margin = 20, align = "center",
               style = list(color = "#129", useHTML = TRUE)) %>% 
      hc_subtitle(text = "",
                  align = "right",
                  style = list(color = "#634", fontWeight = "bold")) %>%
      hc_credits(enabled = TRUE, # add credits
                 text = "" # href = "www.cgr.go.cr"
      ) %>%
      hc_legend(align = "left", verticalAlign = "top",
                layout = "vertical", x = 0, y = 100) %>%
      hc_exporting(enabled = TRUE) %>% 
      hc_xAxis(categories = Ingresos_mensual_var$Fecha) %>% 
      hc_add_series(name = "Variación interanual", data = Ingresos_mensual_var$var.Ingresos) %>% 
      hc_add_series(name = "Variación acumulada", data = Ingresos_mensual_var$var.acum_12) %>% 
      hc_add_series(name = "Variación acumulada al mes", data = Ingresos_mensual_var$var.cum_ano_Ingresos) %>%
      hc_yAxis(title = list(text = "Porcentaje"), labels = list(format = "{value}"))%>% 
      hc_xAxis(title = list(text = "Año-mes") )  %>% 
      hc_chart(zoomType = "xy")
    
    Ingresos_mensual_var
    
  })
  
  output$HCIM_3 <- renderHighchart({
    
    tabla.evo.mensual.1_acum <- tabla.evo.mensual.1_acum
    
    tabla.evo.mensual.1_acum <-tabla.evo.mensual.1_acum %>% mutate(
      
                                                                  cum_ingresos =  round(cum_ingresos/1000000,1)
    )
    
    
    t.e.m.1   <- highchart() %>% 
      hc_title(text = "",
               margin = 20, align = "center",
               style = list(color = "#129", useHTML = TRUE)) %>% 
      hc_subtitle(text = "",
                  align = "right",
                  style = list(color = "#634", fontWeight = "bold")) %>%
      hc_credits(enabled = TRUE, # add credits
                 text = "" # href = "www.cgr.go.cr"
      ) %>%
      hc_legend(align = "left", verticalAlign = "top",
                layout = "vertical", x = 0, y = 100) %>%
      hc_exporting(enabled = TRUE) %>% 
      hc_xAxis(categories = tabla.evo.mensual.1_acum$Fecha) %>% 
      hc_add_series(name = "Ingresos", data = tabla.evo.mensual.1_acum$cum_ingresos) %>% 
      hc_yAxis(title = list(text = "Millones de colones"), labels = list(format = "{value}"))%>% 
      hc_xAxis(title = list(text = "Año-mes") )  %>% 
      hc_chart(zoomType = "xy")
    
    
    t.e.m.1
    
  })
  
  output$HCIM_4 <- renderHighchart({
    
    tabla.evo.mensual.2_acum <- tabla.evo.mensual.2_acum
    
    
    t.e.m.2   <- highchart() %>% 
      hc_title(text = "",
               margin = 20, align = "center",
               style = list(color = "#129", useHTML = TRUE)) %>% 
      hc_subtitle(text = "",
                  align = "right",
                  style = list(color = "#634", fontWeight = "bold")) %>%
      hc_credits(enabled = TRUE, # add credits
                 text = "" # href = "www.cgr.go.cr"
      ) %>%
      hc_legend(align = "left", verticalAlign = "top",
                layout = "vertical", x = 0, y = 100) %>%
      hc_exporting(enabled = TRUE) %>% 
      hc_xAxis(categories = tabla.evo.mensual.2_acum$Fecha) %>% 
      hc_add_series(name = "Carga tributaria", data = tabla.evo.mensual.2_acum$`Carga tributaria`) %>% 
      hc_yAxis(title = list(text = "Porcentaje del PIB."), labels = list(format = "{value}"))%>% 
      hc_xAxis(title = list(text = "Año-mes") )  %>% 
      hc_chart(zoomType = "xy")
    t.e.m.2
    
  })
  
  output$HCIM_5 <- renderHighchart({
    
    tabla.evo.mensual.3_acum <- tabla.evo.mensual.3_acum
    
    
    t.e.m.3   <- highchart() %>% 
      hc_title(text = "",
               margin = 20, align = "center",
               style = list(color = "#129", useHTML = TRUE)) %>% 
      hc_subtitle(text = "",
                  align = "right",
                  style = list(color = "#634", fontWeight = "bold")) %>%
      hc_credits(enabled = TRUE, # add credits
                 text = "" # href = "www.cgr.go.cr"
      ) %>%
      hc_legend(align = "left", verticalAlign = "top",
                layout = "vertical", x = 0, y = 100) %>%
      hc_exporting(enabled = TRUE) %>% 
      hc_xAxis(categories = tabla.evo.mensual.3_acum$Fecha) %>% 
      hc_add_series(name = "% Ejecución", data = tabla.evo.mensual.3_acum$`Ejecucion`) %>% 
      hc_yAxis(title = list(text = "Porcentaje de ejecucción acumulada al mes"), labels = list(format = "{value}"))%>% 
      hc_xAxis(title = list(text = "Año-mes") )  %>% 
      hc_chart(zoomType = "xy")
    t.e.m.3
    
    
  })
  
  
  ###################################
  #           Impuestos             # 
  ###################################
  
  output$HC_I <- renderHighchart({
    
#    filtros <- eventReactive(input$select_2, {
#      Impuestos
#    })
    
    
    
    Impuesto <- Impuestos %>% dplyr::group_by(Año, mes.cod, mes) %>% 
      
          dplyr::summarise ("Ingresos"  = sum(!!input$variable, na.rm = TRUE))  %>%
          dplyr::mutate(
                         Fecha = paste(Año,mes,sep ="-")
                         ) %>% 
         dplyr::arrange(Año,mes.cod) 
      
        
 #   Impuesto <- data.frame(Impuesto)
    
    g_Impuesto  <- highchart() %>% 
      hc_title(text = "",
               margin = 20, align = "center",
               style = list(color = "#129", useHTML = TRUE)) %>% 
      hc_subtitle(text = "",
                  align = "right",
                  style = list(color = "#634", fontWeight = "bold")) %>%
      hc_credits(enabled = TRUE, # add credits
                 text = "" # href = "www.cgr.go.cr"
      ) %>%
      hc_legend(align = "left", verticalAlign = "top",
                layout = "vertical", x = 0, y = 100) %>%
      hc_exporting(enabled = TRUE) %>% 
      hc_xAxis(categories = Impuesto$Fecha) %>% 
      hc_add_series(name = "Ingresos", data = Impuesto$Ingresos) %>% 
      hc_yAxis(title = list(text = "En colones"), labels = list(format = "{value}"))%>% 
      hc_xAxis(title = list(text = "Año-mes"))  %>% 
      hc_chart(zoomType = "xy")
    
        g_Impuesto
    
  })
  
  
  output$HC_VI <- renderHighchart({
    
    #    filtros <- eventReactive(input$select_2, {
    #      Impuestos
    #    })
    
    
    Impuesto <- Impuestos %>% dplyr::group_by(Año, mes.cod, mes) %>% 
      
      #     dplyr::summarise ("Ingresos"  = sum(`Impuesto a los Ingresos y Utilidades-ISR`, na.rm = TRUE))  %>%
      dplyr::summarise ("Ingresos"  = sum(!!input$variable, na.rm = TRUE))  %>%
      mutate(
        Fecha = paste(Año,mes,sep ="-"))   #   mensual$
    
    
    Impuesto <- as.data.frame(Impuesto)
    
    
    Impuesto <- Impuesto %>%  mutate( 
                      var.Ingresos = round((Ingresos/lag(Ingresos,n = 12)-1)*100,2),
                      acum_12 = roll_sum(Ingresos, 12, align = "right", fill = NA),
                      var.acum_12  = round((acum_12/lag(acum_12,n = 12)-1)*100,1)
                    )

    Millones <- 1000000
    
    Impuesto <- Impuesto %>%
      dplyr::mutate(
        Ingresos = round(Ingresos/Millones,1)
      )
    
    Impuesto <-   Impuesto %>% dplyr::group_by(Año)  %>% dplyr::mutate(
      cum_ano_Ingresos     = cumsum(Ingresos)
    )  
    Impuesto <- data.frame(Impuesto) 
    
    Impuesto <- Impuesto %>% dplyr::mutate(var.cum_ano_Ingresos =  round((cum_ano_Ingresos/lag(cum_ano_Ingresos,n = 12)-1)*100,1) )
    
    Impuesto  <- highchart() %>% 
      hc_title(text = "",
               margin = 20, align = "center",
               style = list(color = "#129", useHTML = TRUE)) %>% 
      hc_subtitle(text = "",
                  align = "right",
                  style = list(color = "#634", fontWeight = "bold")) %>%
      hc_credits(enabled = TRUE, # add credits
                 text = "" # href = "www.cgr.go.cr"
      ) %>%
      hc_legend(align = "left", verticalAlign = "top",
                layout = "vertical", x = 0, y = 100) %>%
      hc_exporting(enabled = TRUE) %>% 
      hc_xAxis(categories = Impuesto$Fecha) %>% 
      hc_add_series(name = "Variación interanual", data = Impuesto$var.Ingresos) %>% 
      hc_add_series(name = "Variación acumulada", data = Impuesto$var.acum_12) %>% 
      hc_add_series(name = "Variación acumulada al mes", data = Impuesto$var.cum_ano_Ingresos) %>%
      hc_yAxis(title = list(text = "Porcentaje"), labels = list(format = "{value}"))%>% 
      hc_xAxis(title = list(text = "Año-mes"))  %>% 
      hc_chart(zoomType = "xy")
    
    Impuesto
    
  })
  
 
  
  
  ###################################
  #           Avanzado              # 
  ###################################
  
  filtro_clase <- reactive({
    Ingresos_mensual %>%
      filter(clase == input$f_clase)
    
  })
  
  
  filtro_subclase <- reactive({
    if(input$f_subclase == "Todo"){
      filtro_clase()
    } else {
      filtro_clase() %>% 
        filter(subclase == input$f_subclase)
    }
  })
  
  
  filtro_grupo <- reactive({
    if(input$f_grupo == "Todo"){
      filtro_subclase()
    } else {
      filtro_subclase() %>% 
        filter(grupo == input$f_grupo)
    }
  })
  
  filtro_subgrupo <- reactive({
    if(input$f_subgrupo == "Todo"){
      filtro_grupo()
    } else {
      filtro_grupo() %>% 
        filter(subgrupo == input$f_subgrupo)
    }
  })
  
  filtro_partida <- reactive({
    if(input$f_partida == "Todo"){
      filtro_subgrupo()
    } else {
      filtro_subgrupo() %>% 
        filter(partida == input$f_partida)
    }
  })
  
  filtro_subpartida <- reactive({
    if(input$f_subpartida == "Todo"){
      filtro_partida()
    } else {
      filtro_partida() %>% 
        filter(subpartida == input$f_subpartida)
    }
  })
  
  filtro_renglon <- reactive({
    if(input$f_renglon == "Todo"){
      filtro_subpartida()
    } else {
      filtro_subpartida() %>% 
        filter(renglon == input$f_renglon)
    }
  })
  
  filtro_subrenglon <- reactive({
    if(input$f_subrenglon == "Todo"){
      filtro_renglon()
    } else {
      filtro_renglon() %>% 
        filter(subrenglon == input$f_subrenglon)
    }
  })
  
  # Ultimo # 
  
  filtros <- eventReactive(input$select_2, {
    filtro_subrenglon()
  })
  
  
  output$HC_P <- renderHighchart({
    
    Ingresos_mensual <- filtros() %>% dplyr::group_by(Año, mes.cod, mes) %>% 
      dplyr::summarise ("Ingresos"  = sum(Ingresos, na.rm = TRUE))  %>%
      mutate(
        Fecha = paste(Año,mes,sep ="-"))  %>%  
      dplyr::filter(Ingresos>0)
    
    
    Millones <- 1000000
    
    Ingresos_mensual <- Ingresos_mensual %>%
      dplyr::mutate(
        Ingresos = round(Ingresos/Millones,1)
      )
    
    Ingresos_mensual  <- highchart() %>% 
      hc_title(text = "",
               margin = 20, align = "center",
               style = list(color = "#129", useHTML = TRUE)) %>% 
      hc_subtitle(text = "",
                  align = "right",
                  style = list(color = "#634", fontWeight = "bold")) %>%
      hc_credits(enabled = TRUE, # add credits
                 text = "" # href = "www.cgr.go.cr"
      ) %>%
      hc_legend(align = "left", verticalAlign = "top",
                layout = "vertical", x = 0, y = 100) %>%
      hc_exporting(enabled = TRUE) %>% 
      hc_xAxis(categories = Ingresos_mensual$Fecha) %>% 
      hc_add_series(name = "Ingresos", data = Ingresos_mensual$Ingresos) %>% 
      hc_yAxis(title = list(text = "Millones de colones"), labels = list(format = "{value}"))%>% 
      hc_xAxis(title = list(text = "Año-mes"))  %>% 
      hc_chart(zoomType = "xy")
    
    Ingresos_mensual
    
  })
  
  output$HC_VP <- renderHighchart({
    
    Ingresos_mensual <- filtros() %>% dplyr::group_by(Año, mes.cod, mes) %>% 
      dplyr::summarise ("Ingresos"  = sum(Ingresos, na.rm = TRUE))  %>%
      mutate(
        Fecha = paste(Año,mes,sep ="-"))   %>%  
      dplyr::filter(Ingresos>0)
    
    Ingresos_mensual <- Ingresos_mensual %>%  mutate( 
                            var.Ingresos = round((Ingresos/lag(Ingresos,n = 12)-1)*100,2),
                            acum_12 = roll_sum(Ingresos, 12, align = "right", fill = NA),
                          var.acum_12  = round((acum_12/lag(acum_12,n = 12)-1)*100,1)
                        )
    
    
    Millones <- 1000000
    
    Ingresos_mensual <- Ingresos_mensual %>%
      dplyr::mutate(
        Ingresos = round(Ingresos/Millones,1)
      )
    
    Ingresos_mensual <-   Ingresos_mensual %>% dplyr::group_by(Año)  %>% dplyr::mutate(
      cum_ano_Ingresos     = cumsum(Ingresos)
    )  
    Ingresos_mensual <- data.frame(Ingresos_mensual) 
    
    Ingresos_mensual <- Ingresos_mensual %>% dplyr::mutate(var.cum_ano_Ingresos =  round((cum_ano_Ingresos/lag(cum_ano_Ingresos,n = 12)-1)*100,1) )
    
    
    Ingresos_mensual  <- highchart() %>% 
      hc_title(text = "",
               margin = 20, align = "center",
               style = list(color = "#129", useHTML = TRUE)) %>% 
      hc_subtitle(text = "",
                  align = "right",
                  style = list(color = "#634", fontWeight = "bold")) %>%
      hc_credits(enabled = TRUE, # add credits
                 text = "" # href = "www.cgr.go.cr"
      ) %>%
      hc_legend(align = "left", verticalAlign = "top",
                layout = "vertical", x = 0, y = 100) %>%
      hc_exporting(enabled = TRUE) %>% 
      hc_xAxis(categories = Ingresos_mensual$Fecha) %>% 
      hc_add_series(name = "Variación interanual", data = Ingresos_mensual$var.Ingresos) %>% 
      hc_add_series(name = "Variación acumulada", data = Ingresos_mensual$var.acum_12) %>% 
      hc_add_series(name = "Variación acumulada al mes", data = Ingresos_mensual$var.cum_ano_Ingresos) %>% 
      hc_yAxis(title = list(text = "Porcentaje"), labels = list(format = "{value}"))%>% 
      hc_xAxis(title = list(text = "Año-mes"))  %>% 
      
      hc_chart(zoomType = "xy")
    
    Ingresos_mensual
    
  })
  
##   #############################
##   #############################
##   # Análisis por clasificador #
##   #############################
##   #############################
##   
##   #######
##   #Clase#
##   #######
##   
##   # Indicadores 
##   
##   # Tabla con ingresos
##   
##   output$clasi_clase = renderReactable({
##     
##     reactable(tabla_clasi_c_1_1)
##   })
##   
##   #Gráfico mensual con ingresos
##   
##   output$hc_clasi_clase <- renderHighchart({
##     
##     Ingresos_clase <- tabla_clasi_c_1_1
##     
##     
##     Ingresos_clase  <- highchart() %>% 
##       hc_title(text = "",
##                margin = 20, align = "center",
##                style = list(color = "#129", useHTML = TRUE)) %>% 
##       hc_subtitle(text = "",
##                   align = "right",
##                   style = list(color = "#634", fontWeight = "bold")) %>%
##       hc_credits(enabled = TRUE, # add credits
##                  text = "" # href = "www.cgr.go.cr"
##       ) %>%
##       hc_legend(align = "left", verticalAlign = "top",
##                 layout = "vertical", x = 0, y = 100) %>%
##       hc_tooltip(crosshairs = TRUE, backgroundColor = "#FCFFC5",
##                  shared = TRUE, borderWidth = 5) %>% 
##       hc_exporting(enabled = TRUE) %>% 
##       hc_xAxis(categories = Ingresos_clase$Fecha) %>% 
##       hc_add_series(name = "Financiamiento", data = Ingresos_clase$FINANCIAMIENTO) %>% 
##       hc_add_series(name = "Ingresos corrientes", data = Ingresos_clase$`INGRESOS CORRIENTES`) %>% 
##       hc_add_series(name = "Ingresos de capital", data = Ingresos_clase$`INGRESOS DE CAPITAL`) %>% 
##       hc_chart(zoomType = "xy")
##     
##     Ingresos_clase
##     
##   })
##   
##   ##########
##   #Subclase#
##   ##########
##   
##   # Indicadores 
##   
##   # Tabla con ingresos
##   
##   output$clasi_subclase = renderReactable({
##     
##     reactable(tabla_clasi_c_2_1)
##   })
##   
##   #Gráfico mensual con ingresos
##   
##   output$hc_clasi_subclase <- renderHighchart({
##     
##     Ingresos_subclase <- tabla_clasi_c_2_1
##     
##     
##     Ingresos_subclase  <- highchart() %>% 
##       hc_title(text = "",
##                margin = 20, align = "center",
##                style = list(color = "#129", useHTML = TRUE)) %>% 
##       hc_subtitle(text = "",
##                   align = "right",
##                   style = list(color = "#634", fontWeight = "bold")) %>%
##       hc_credits(enabled = TRUE, # add credits
##                  text = "" # href = "www.cgr.go.cr"
##       ) %>%
##       hc_legend(align = "left", verticalAlign = "top",
##                 layout = "vertical", x = 0, y = 100) %>%
##       hc_tooltip(crosshairs = TRUE, backgroundColor = "#FCFFC5",
##                  shared = TRUE, borderWidth = 5) %>% 
##       hc_exporting(enabled = TRUE) %>% 
##       hc_xAxis(categories = Ingresos_subclase$Fecha) %>% 
##       hc_add_series(name = "Contribuciones sociales", data = Ingresos_subclase$`CONTRIBUCIONES SOCIALES`) %>% 
##       hc_add_series(name = "Financiamiento externo", data = Ingresos_subclase$`FINANCIAMIENTO EXTERNO`) %>% 
##       hc_chart(zoomType = "xy")
##     
##     Ingresos_subclase
##     
##   })
##   
##   #######
##   #Grupo#
##   #######
##   
##   # Indicadores 
##   
##   # Tabla con ingresos
##   
##   output$clasi_grupo = renderReactable({
##     
##     reactable(tabla_clasi_c_3_1)
##   })
##   
##   #Gráfico mensual con ingresos
##   
##   output$hc_clasi_grupo <- renderHighchart({
##     
##     Ingresos_grupo <- tabla_clasi_c_3_1
##     
##     
##     Ingresos_grupo  <- highchart() %>% 
##       hc_title(text = "",
##                margin = 20, align = "center",
##                style = list(color = "#129", useHTML = TRUE)) %>% 
##       hc_subtitle(text = "",
##                   align = "right",
##                   style = list(color = "#634", fontWeight = "bold")) %>%
##       hc_credits(enabled = TRUE, # add credits
##                  text = "" # href = "www.cgr.go.cr"
##       ) %>%
##       hc_legend(align = "left", verticalAlign = "top",
##                 layout = "vertical", x = 0, y = 100) %>%
##       hc_tooltip(crosshairs = TRUE, backgroundColor = "#FCFFC5",
##                  shared = TRUE, borderWidth = 5) %>% 
##       hc_exporting(enabled = TRUE) %>% 
##       hc_xAxis(categories = Ingresos_grupo$Fecha) %>% 
##       hc_add_series(name = "Contribuciones sociales", data = Ingresos_grupo$`COLOCACIÓN DE TÍTULOS VALORES`) %>% 
##       hc_add_series(name = "Financiamiento externo", data = Ingresos_grupo$`IMPUESTOS A LOS INGRESOS Y UTILIDADES`) %>% 
##       hc_chart(zoomType = "xy")
##     
##     Ingresos_grupo
##     
##   })
##   
##   ##########
##   #Subgrupo#
##   ##########
##   
##   # Indicadores 
##   
##   # Tabla con ingresos
##   
##   output$clasi_subgrupo = renderReactable({
##     
##     reactable(tabla_clasi_c_4_1)
##   })
##   
##   #Gráfico mensual con ingresos
##   
##   output$hc_clasi_subgrupo <- renderHighchart({
##     
##     Ingresos_subgrupo <- tabla_clasi_c_4_1
##     
##     
##     Ingresos_subgrupo  <- highchart() %>% 
##       hc_title(text = "",
##                margin = 20, align = "center",
##                style = list(color = "#129", useHTML = TRUE)) %>% 
##       hc_subtitle(text = "",
##                   align = "right",
##                   style = list(color = "#634", fontWeight = "bold")) %>%
##       hc_credits(enabled = TRUE, # add credits
##                  text = "" # href = "www.cgr.go.cr"
##       ) %>%
##       hc_legend(align = "left", verticalAlign = "top",
##                 layout = "vertical", x = 0, y = 100) %>%
##       hc_tooltip(crosshairs = TRUE, backgroundColor = "#FCFFC5",
##                  shared = TRUE, borderWidth = 5) %>% 
##       hc_exporting(enabled = TRUE) %>% 
##       hc_xAxis(categories = Ingresos_subgrupo$Fecha) %>% 
##       hc_add_series(name = "T-V corto plazo", data = Ingresos_subgrupo$`COLOCACIÓN DE TÍTULOS VALORES DE CORTO PLAZO`) %>% 
##       hc_chart(zoomType = "xy")
##     
##     Ingresos_subgrupo
##     
##   })
##   
##   #########
##   #Partida#
##   #########
##   
##   # Indicadores 
##   
##   # Tabla con ingresos
##   
##   output$clasi_partida = renderReactable({
##     
##     reactable(tabla_clasi_c_5_1)
##   })
##   
##   #Gráfico mensual con ingresos
##   
##   output$hc_clasi_partida <- renderHighchart({
##     
##     Ingresos_partida <- tabla_clasi_c_5_1
##     
##     
##     Ingresos_partida  <- highchart() %>% 
##       hc_title(text = "",
##                margin = 20, align = "center",
##                style = list(color = "#129", useHTML = TRUE)) %>% 
##       hc_subtitle(text = "",
##                   align = "right",
##                   style = list(color = "#634", fontWeight = "bold")) %>%
##       hc_credits(enabled = TRUE, # add credits
##                  text = "" # href = "www.cgr.go.cr"
##       ) %>%
##       hc_legend(align = "left", verticalAlign = "top",
##                 layout = "vertical", x = 0, y = 100) %>%
##       hc_tooltip(crosshairs = TRUE, backgroundColor = "#FCFFC5",
##                  shared = TRUE, borderWidth = 5) %>% 
##       hc_exporting(enabled = TRUE) %>% 
##       hc_xAxis(categories = Ingresos_partida$Fecha) %>% 
##       hc_add_series(name = "Alquileres", data = Ingresos_partida$`ALQUILERES`) %>% 
##       hc_chart(zoomType = "xy")
##     
##     Ingresos_partida
##     
##   })
##   
##   ############
##   #Subpartida#
##   ############
##   
##   # Indicadores 
##   
##   # Tabla con ingresos
##   
##   output$clasi_subpartida = renderReactable({
##     
##     reactable(tabla_clasi_c_6_1)
##   })
##   
##   #Gráfico mensual con ingresos
##   
##   output$hc_clasi_subpartida <- renderHighchart({
##     
##     Ingresos_partida <- tabla_clasi_c_6_1
##     
##     
##     Ingresos_partida  <- highchart() %>% 
##       hc_title(text = "",
##                margin = 20, align = "center",
##                style = list(color = "#129", useHTML = TRUE)) %>% 
##       hc_subtitle(text = "",
##                   align = "right",
##                   style = list(color = "#634", fontWeight = "bold")) %>%
##       hc_credits(enabled = TRUE, # add credits
##                  text = "" # href = "www.cgr.go.cr"
##       ) %>%
##       hc_legend(align = "left", verticalAlign = "top",
##                 layout = "vertical", x = 0, y = 100) %>%
##       hc_tooltip(crosshairs = TRUE, backgroundColor = "#FCFFC5",
##                  shared = TRUE, borderWidth = 5) %>% 
##       hc_exporting(enabled = TRUE) %>% 
##       hc_xAxis(categories = Ingresos_partida$Fecha) %>% 
##       hc_add_series(name = "Protección al consumidor", data = Ingresos_partida$`1% Ley Protección al Consumidor`) %>% 
##       hc_chart(zoomType = "xy")
##     
##     Ingresos_partida
##     
##   })
##   
##   #########
##   #Renglón#
##   #########
##  
##   # Indicadores 
##   
##   # Tabla con ingresos
##   
##   output$clasi_renglon = renderReactable({
##     
##     reactable(tabla_clasi_c_7_1)
##   })
##   
##   #Gráfico mensual con ingresos
##   
##   output$hc_clasi_renglon <- renderHighchart({
##     
##     Ingresos_renglon <- tabla_clasi_c_7_1
##     
##     
##     Ingresos_renglon  <- highchart() %>% 
##       hc_title(text = "",
##                margin = 20, align = "center",
##                style = list(color = "#129", useHTML = TRUE)) %>% 
##       hc_subtitle(text = "",
##                   align = "right",
##                   style = list(color = "#634", fontWeight = "bold")) %>%
##       hc_credits(enabled = TRUE, # add credits
##                  text = "" # href = "www.cgr.go.cr"
##       ) %>%
##       hc_legend(align = "left", verticalAlign = "top",
##                 layout = "vertical", x = 0, y = 100) %>%
##       hc_tooltip(crosshairs = TRUE, backgroundColor = "#FCFFC5",
##                  shared = TRUE, borderWidth = 5) %>% 
##       hc_exporting(enabled = TRUE) %>% 
##       hc_xAxis(categories = Ingresos_renglon$Fecha) %>% 
##       hc_add_series(name = "Alquiler de edificios", data = Ingresos_renglon$`Alquiler de edificios`) %>% 
##       hc_chart(zoomType = "xy")
##     
##     Ingresos_renglon
##     
##   })
##   
##   ############
##   #Subrenglón#
##   ############
##   
##   # Indicadores 
##   
##   # Tabla con ingresos
##   
##   output$clasi_subrenglon = renderReactable({
##     
##     reactable(tabla_clasi_c_8_1)
##   })
##   
##   #Gráfico mensual con ingresos
##   
##   output$hc_clasi_subrenglon <- renderHighchart({
##     
##     Ingresos_subrenglon <- tabla_clasi_c_8_1
##     
##     
##     Ingresos_subrenglon  <- highchart() %>% 
##       hc_title(text = "",
##                margin = 20, align = "center",
##                style = list(color = "#129", useHTML = TRUE)) %>% 
##       hc_subtitle(text = "",
##                   align = "right",
##                   style = list(color = "#634", fontWeight = "bold")) %>%
##       hc_credits(enabled = TRUE, # add credits
##                  text = "" # href = "www.cgr.go.cr"
##       ) %>%
##       hc_legend(align = "left", verticalAlign = "top",
##                 layout = "vertical", x = 0, y = 100) %>%
##       hc_tooltip(crosshairs = TRUE, backgroundColor = "#FCFFC5",
##                  shared = TRUE, borderWidth = 5) %>% 
##       hc_exporting(enabled = TRUE) %>% 
##       hc_xAxis(categories = Ingresos_subrenglon$Fecha) %>% 
##       hc_add_series(name = "Impuesto a los productos de tabaco (Importaciones)", data = Ingresos_subrenglon$`Impuesto a los productos de tabaco (Importaciones)`) %>% 
##       hc_chart(zoomType = "xy")
##     
##     Ingresos_subrenglon
##     
##   })
## 
##   ########################
##   #Fuente de Financiación#
##   ########################
##   
##   # Indicadores 
##   
##   # Tabla con ingresos
##   
##   output$clasi_ff = renderReactable({
##     
##     reactable(tabla_clasi_c_9_1)
##   })
##   
##   #Gráfico mensual con ingresos
  
  ############################################
  #              Pronóstico                  #
  ############################################
  
  ###########################################
  #    Proyeccion del gasto  total mensual  #
  ###########################################
  
  # Gráfico #
  
  output$forecast1 <- renderHighchart({
    
    
     Ingresos_mensuales_totales <- tabla_2  %>% filter(row_number() <= n()-1)
        
 #       Egresos_GC_GC_mensual_anual <- Egresos_GC_GC_mensual_anual[-c(dim(Egresos_GC_GC_mensual_anual)[1]),]
        
        serie1 <- ts(Ingresos_mensuales_totales$Ingresos, frequency=12, start=c(2007,1))
        
        ##########################
        #  ARIMA (1,1,1) (0,1,1) #
        ##########################
        
        modelo_1 <- Arima(serie1, order=c(1,1,1), seasonal=c(0,1,1))
        x1 <- forecast(modelo_1, h = input$horizonte.1, level = 95)
        
        
        ##########################
        #          ETS MAM       #
        ##########################
        
        modelo_2 <- ets(serie1,  model ="MAM",damped = TRUE)
        x2 <- forecast(modelo_2, h = input$horizonte.1, level = 95)
        
        ##########################
        #       Auto-ARIMA       #
        ##########################
        
        modelo_3 <- auto.arima(serie1)
        x3 <- forecast(modelo_3, h = input$horizonte.1, level = 95)
        
        ##########################
        #    Regresion TS       #
        ##########################  
        
        modelo_4 <- tslm(serie1 ~ trend+season) 
        x4 <- forecast(modelo_4, h = input$horizonte.1, level = 95)
        
        
        XF  <-hchart(serie1,name = "Ingresos") %>% 
          
          
          hc_add_series(name = "ARIMA",fitted(modelo_1), type = "line", color="green") %>% 
          hc_add_series(name = "F ARIMA", x1$mean, color="green") %>%
          
          hc_add_series(name = "ETS",fitted(modelo_2), type = "line", color="black") %>% 
          hc_add_series(name = "F ETs", x2$mean, color="black") %>%
          
          hc_add_series(name = "AUTO.ARIMA",fitted(modelo_3), type = "line", color="blue") %>%
          hc_add_series(name = "F AUTO.ARIMA", x3$mean, color="blue") %>%
          
          hc_add_series(name = "Regresión",fitted(modelo_4), type = "line", color="grey") %>%
          hc_add_series(name = "F Regresión", x4$mean, color="grey") %>%
          
          hc_title(text = "",
                   margin = 20, align = "center",
                   style = list(color = "#black", useHTML = TRUE, fontWeight = "bold")) %>%
          hc_yAxis(title = list(text = "En colones,"), labels = list(format = "{value}"))%>% 
          hc_xAxis(title = list(text = "Años") )  %>% 
          
          hc_tooltip(crosshairs = T,valueDecimals = 1, shared = TRUE, borderWidth = 5) %>%
          hc_chart(
            zoomType = "xy"
          )
        
        XF 
    
  })
  
  #################################################
  # Talba con estadísticos de bondad y de ajuste  #
  #################################################
  
  output$GoF.1 <- renderReactable({
    
    Ingresos_mensuales_totales <- tabla_2  %>% filter(row_number() <= n()-1)
    
    #       Egresos_GC_GC_mensual_anual <- Egresos_GC_GC_mensual_anual[-c(dim(Egresos_GC_GC_mensual_anual)[1]),]
    
    
    serie1 <- ts(Ingresos_mensuales_totales$Ingresos, frequency=12, start=c(2007,1))
    
    ##########################
    #  ARIMA (1,1,1) (0,1,1) #
    ##########################
    
    modelo_1 <- Arima(serie1, order=c(1,1,1), seasonal=c(0,1,1))
    x1 <- forecast(modelo_1, h = 12, level = 95)
    
    
    ##########################
    #          ETS MAM       #
    ##########################
    
    modelo_2 <- ets(serie1,  model ="MAM",damped = TRUE)
    x2 <- forecast(modelo_2, h = 12, level = 95)
    
    ##########################
    #       Auto-ARIMA       #
    ##########################
    
    modelo_3 <- auto.arima(serie1)
    x3 <- forecast(modelo_3, h = 12, level = 95)
    
    ##########################
    #    Regresion TS       #
    ##########################  
    
    modelo_4 <- tslm(serie1 ~ trend+season) 
    x4 <- forecast(modelo_4, h = 12, level = 95)
    
    REGRESION <-round(c(accuracy(modelo_4)[c(2,3,5)],CV(modelo_4)[c(2,4)]),3)
    ETS.MAM<-round(c(accuracy(modelo_2)[c(2,3,5)],modelo_2$aic,modelo_2$bic),3)
    ARIMA<-round(c(accuracy(modelo_1)[c(2,3,5)],modelo_1$aic,modelo_1$bic),3)
    AUTO.ARIMA<-round(c(accuracy(modelo_3)[c(2,3,5)],modelo_3$aic,modelo_3$bic),3)
    
    cuadro<-rbind(Medida = c("RSME","MAE","MAPE","AIC","BIC"),
                  REGRESION,ETS.MAM,ARIMA,AUTO.ARIMA)
    
    colnames(cuadro)<-cuadro[1,]
    cuadro <- cuadro[-1,]
    
    reactable(cuadro)
    
  })
  
  ##########################
  #    PRONOSTICO ARIMA    #
  ##########################
  
  output$tabla.forcast.1.1 <- renderReactable({ 
    
    Ingresos_mensuales_totales <- tabla_2  %>% filter(row_number() <= n()-1)
    
    #       Egresos_GC_GC_mensual_anual <- Egresos_GC_GC_mensual_anual[-c(dim(Egresos_GC_GC_mensual_anual)[1]),]
    
    serie1 <- ts(Ingresos_mensuales_totales$Ingresos, frequency=12, start=c(2007,1))
    
    modelo <- Arima(serie1, order=c(1,1,1), seasonal=c(0,1,1))
    
    x1 <- forecast(modelo, h = input$horizonte.1, level = 95)
    
    
    pronosticos <- x1 %>%
      as.data.frame() %>%
      dplyr::rename(
        Proyección =`Point Forecast`,
        `Límite inferior` =`Lo 95`,
        `Límite superior` =`Hi 95`
      ) %>% mutate(
        Proyección = round(Proyección/Millones,1),
        `Límite inferior` = round(`Límite inferior`/Millones,1),
        `Límite superior` = round(`Límite superior`/Millones,1)
      )
    
    pronosticos <- pronosticos %>%
      mutate(
        `Proyección` = number(pronosticos$`Proyección`,           accuracy = .1, big.mark = ".", decimal.mark = ","),
        `Límite inferior` = number(pronosticos$`Límite inferior`, accuracy = .1, big.mark = ".", decimal.mark = ","),
        `Límite superior` = number(pronosticos$`Límite superior`, accuracy = .1, big.mark = ".", decimal.mark = ",")
      )
    
    reactable(pronosticos) 
    
  })
  
  ##########################
  #    PRONOSTICO ETS      #
  ##########################
  
  output$tabla.forcast.1.2 <- renderReactable({ 
    
    Ingresos_mensuales_totales <- tabla_2  %>% filter(row_number() <= n()-1)
    
    #       Egresos_GC_GC_mensual_anual <- Egresos_GC_GC_mensual_anual[-c(dim(Egresos_GC_GC_mensual_anual)[1]),]
    
    serie1 <- ts(Ingresos_mensuales_totales$Ingresos, frequency=12, start=c(2007,1))
    
    ##########################
    #          ETS MAM       #
    ##########################
    
    modelo_2 <- ets(serie1,  model ="MAM",damped = TRUE)
    x2 <- forecast(modelo_2, h = 12, level = 95)
    
    
    pronosticos <- x2 %>%
      as.data.frame() %>%
      dplyr::rename(
        Proyección =`Point Forecast`,
        `Límite inferior` =`Lo 95`,
        `Límite superior` =`Hi 95`
      ) %>% mutate(
        Proyección = round(Proyección/Millones,1),
        `Límite inferior` = round(`Límite inferior`/Millones,1),
        `Límite superior` = round(`Límite superior`/Millones,1)
      )
    
    pronosticos <- pronosticos %>%
      mutate(
        `Proyección` = number(pronosticos$`Proyección`,           accuracy = .1, big.mark = ".", decimal.mark = ","),
        `Límite inferior` = number(pronosticos$`Límite inferior`, accuracy = .1, big.mark = ".", decimal.mark = ","),
        `Límite superior` = number(pronosticos$`Límite superior`, accuracy = .1, big.mark = ".", decimal.mark = ",")
      )
    
    reactable(pronosticos) 
    
  })
   
  ##############################
  #    PRONOSTICO AUTO.ARIMA   #
  ##############################
 
  output$tabla.forcast.1.3 <- renderReactable({ 
    
    Ingresos_mensuales_totales <- tabla_2  %>% filter(row_number() <= n()-1)
    
    #       Egresos_GC_GC_mensual_anual <- Egresos_GC_GC_mensual_anual[-c(dim(Egresos_GC_GC_mensual_anual)[1]),]
    
    serie1 <- ts(Ingresos_mensuales_totales$Ingresos, frequency=12, start=c(2007,1))
    
    ##########################
    #          ETS MAM       #
    ##########################
    
    modelo_3 <- auto.arima(serie1)
    x3 <- forecast(modelo_3, h = 12, level = 95)
    
    
    pronosticos <- x3 %>%
      as.data.frame() %>%
      dplyr::rename(
        Proyección =`Point Forecast`,
        `Límite inferior` =`Lo 95`,
        `Límite superior` =`Hi 95`
      ) %>% mutate(
        Proyección = round(Proyección/Millones,1),
        `Límite inferior` = round(`Límite inferior`/Millones,1),
        `Límite superior` = round(`Límite superior`/Millones,1)
      )
    
    pronosticos <- pronosticos %>%
      mutate(
        `Proyección` = number(pronosticos$`Proyección`,           accuracy = .1, big.mark = ".", decimal.mark = ","),
        `Límite inferior` = number(pronosticos$`Límite inferior`, accuracy = .1, big.mark = ".", decimal.mark = ","),
        `Límite superior` = number(pronosticos$`Límite superior`, accuracy = .1, big.mark = ".", decimal.mark = ",")
      )
    
    
    reactable(pronosticos)
    
  }) 
  
  ##############################
  #    PRONOSTICO REGRESION    #
  ##############################
  
  output$tabla.forcast.1.4 <- renderReactable({ 
    
    Ingresos_mensuales_totales <- tabla_2  %>% filter(row_number() <= n()-1)
    
    #       Egresos_GC_GC_mensual_anual <- Egresos_GC_GC_mensual_anual[-c(dim(Egresos_GC_GC_mensual_anual)[1]),]
    
    serie1 <- ts(Ingresos_mensuales_totales$Ingresos, frequency=12, start=c(2007,1))
    
    ##########################
    #          ETS MAM       #
    ##########################
    
    modelo_4 <- tslm(serie1 ~ trend+season) 
    x4 <- forecast(modelo_4, h = 12, level = 95)
    
    
    pronosticos <- x4 %>%
      as.data.frame() %>%
      dplyr::rename(
        Proyección =`Point Forecast`,
        `Límite inferior` =`Lo 95`,
        `Límite superior` =`Hi 95`
      ) %>% mutate(
        Proyección = round(Proyección/Millones,1),
        `Límite inferior` = round(`Límite inferior`/Millones,1),
        `Límite superior` = round(`Límite superior`/Millones,1)
      )
    
    pronosticos <- pronosticos %>%
      mutate(
        `Proyección` = number(pronosticos$`Proyección`,           accuracy = .1, big.mark = ".", decimal.mark = ","),
        `Límite inferior` = number(pronosticos$`Límite inferior`, accuracy = .1, big.mark = ".", decimal.mark = ","),
        `Límite superior` = number(pronosticos$`Límite superior`, accuracy = .1, big.mark = ".", decimal.mark = ",")
      )
    
    reactable(pronosticos)
    
  }) 
  
  
###  ###############################################
###  #    Proyeccion del gasto según clasificador  #
###  ###############################################
###  
###  # Gráfico #
###  
###  output$forecast2 <- renderHighchart({
###    
###    
###    Ingresos_mensual <- Ingresos_mensual %>%
###      dplyr::filter(n_Nivel==input$clasi.2_2)  %>% 
###      dplyr::group_by(Año, mes.cod, mes) %>% 
###      dplyr::summarise ("Ingresos"  = sum(Ingresos, na.rm = TRUE))  %>%
###      mutate(
###        Fecha = paste(Año,mes,sep ="-"))  %>%  
###      dplyr::filter(Ingresos>0)
###    
###    Ingresos_mensual <- Ingresos_mensual %>%
###      dplyr::mutate(
###        Ingresos = round(Ingresos/Millones,1)
###      ) 
### 
###    Ingresos_mensual <-  as.data.frame(Ingresos_mensual) %>% filter(row_number() <= n()-1)
###    
###    serie1 <- ts(Ingresos_mensual$Ingresos, frequency=12, start=c(2007,1))
###    
###    ##########################
###    #  ARIMA (1,1,1) (0,1,1) #
###    ##########################
###    
###    modelo_1 <- Arima(serie1, order=c(1,1,1), seasonal=c(0,1,1))
###    x1 <- forecast(modelo_1, h = input$horizonte.1, level = 95)
###    
###    
###    ##########################
###    #          ETS MAM       #
###    ##########################
###    
###    modelo_2 <- ets(serie1,  model ="MAM",damped = TRUE)
###    x2 <- forecast(modelo_2, h = input$horizonte.1, level = 95)
###    
###    ##########################
###    #       Auto-ARIMA       #
###    ##########################
###    
###    modelo_3 <- auto.arima(serie1)
###    x3 <- forecast(modelo_3, h = input$horizonte.1, level = 95)
###    
###    ##########################
###    #    Regresion TS       #
###    ##########################  
###    
###    modelo_4 <- tslm(serie1 ~ trend+season) 
###    x4 <- forecast(modelo_4, h = input$horizonte.1, level = 95)
###    
###    
###    XF  <-hchart(serie1,name = "Ingresos") %>% 
###      
###      
###      hc_add_series(name = "ARIMA",fitted(modelo_1), type = "line", color="green") %>% 
###      hc_add_series(name = "F ARIMA", x1$mean, color="green") %>%
###      
###      hc_add_series(name = "ETS",fitted(modelo_2), type = "line", color="black") %>% 
###      hc_add_series(name = "F ETs", x2$mean, color="black") %>%
###      
###      hc_add_series(name = "AUTO.ARIMA",fitted(modelo_3), type = "line", color="blue") %>%
###      hc_add_series(name = "F AUTO.ARIMA", x3$mean, color="blue") %>%
###      
###      hc_add_series(name = "Regresión",fitted(modelo_4), type = "line", color="grey") %>%
###      hc_add_series(name = "F Regresión", x4$mean, color="grey") %>%
###      
###      hc_title(text = "",
###               margin = 20, align = "center",
###               style = list(color = "#black", useHTML = TRUE, fontWeight = "bold")) %>%
###      hc_yAxis(title = list(text = "Colones"),
###               labels = list(format = "{value}"))%>% 
###      hc_xAxis(title = list(text = "Anos") )  %>% 
###      
###      hc_tooltip(crosshairs = T,valueDecimals = 1, shared = TRUE, borderWidth = 5) %>%
###      hc_chart(
###        zoomType = "xy"
###      )
###    
###    XF 
###    
###  })
###  
###  
###  #################################################################
###  # Talba con estadísticos de bondad y de ajuste por clasificador #
###  #################################################################
###  
###  output$GoF.2 <- renderReactable({
###    
###    Ingresos_mensual <- Ingresos_mensual %>%
###      dplyr::filter(n_Nivel==input$clasi.2_2)  %>% 
###      dplyr::group_by(Año, mes.cod, mes) %>% 
###      dplyr::summarise ("Ingresos"  = sum(Ingresos, na.rm = TRUE))  %>%
###      mutate(
###        Fecha = paste(Año,mes,sep ="-"))  %>%  
###      dplyr::filter(Ingresos>0)
###    
###    Ingresos_mensual <- Ingresos_mensual %>%
###      dplyr::mutate(
###        Ingresos = round(Ingresos/Millones,1)
###      ) 
###    
###    Ingresos_mensual <-  as.data.frame(Ingresos_mensual) %>% filter(row_number() <= n()-1)
###    
###    serie1 <- ts(Ingresos_mensual$Ingresos, frequency=12, start=c(2007,1))
###    
###    ##########################
###    #  ARIMA (1,1,1) (0,1,1) #
###    ##########################
###    
###    modelo_1 <- Arima(serie1, order=c(1,1,1), seasonal=c(0,1,1))
###    x1 <- forecast(modelo_1, h = input$horizonte.1, level = 95)
###    
###    
###    ##########################
###    #          ETS MAM       #
###    ##########################
###    
###    modelo_2 <- ets(serie1,  model ="MAM",damped = TRUE)
###    x2 <- forecast(modelo_2, h = input$horizonte.1, level = 95)
###    
###    ##########################
###    #       Auto-ARIMA       #
###    ##########################
###    
###    modelo_3 <- auto.arima(serie1)
###    x3 <- forecast(modelo_3, h = input$horizonte.1, level = 95)
###    
###    ##########################
###    #    Regresion TS       #
###    ##########################  
###    
###    modelo_4 <- tslm(serie1 ~ trend+season) 
###    x4 <- forecast(modelo_4, h = input$horizonte.1, level = 95)
###    
###    REGRESION <-round(c(accuracy(modelo_4)[c(2,3,5)],CV(modelo_4)[c(2,4)]),3)
###    ETS.MAM<-round(c(accuracy(modelo_2)[c(2,3,5)],modelo_2$aic,modelo_2$bic),3)
###    ARIMA<-round(c(accuracy(modelo_1)[c(2,3,5)],modelo_1$aic,modelo_1$bic),3)
###    AUTO.ARIMA<-round(c(accuracy(modelo_3)[c(2,3,5)],modelo_3$aic,modelo_3$bic),3)
###    
###    cuadro<-rbind(Medida = c("RSME","MAE","MAPE","AIC","BIC"),
###                  REGRESION,ETS.MAM,ARIMA,AUTO.ARIMA)
###    
###    colnames(cuadro)<-cuadro[1,]
###    cuadro <- cuadro[-1,]
###    
###    reactable(cuadro)
###    
###  }) 
###
#########   
###
###  ##########################
###  #    PRONOSTICO ARIMA    #
###  ##########################
###  
###  output$tabla.forcast.2.1 <- renderReactable({ 
###    
###    Ingresos_mensual <- Ingresos_mensual %>%
###      dplyr::filter(n_Nivel==input$clasi.2_2)  %>% 
###      dplyr::group_by(Año, mes.cod, mes) %>% 
###      dplyr::summarise ("Ingresos"  = sum(Ingresos, na.rm = TRUE))  %>%
###      mutate(
###        Fecha = paste(Año,mes,sep ="-"))  %>%  
###      dplyr::filter(Ingresos>0)
###    
###    
###    Ingresos_mensual <-  as.data.frame(Ingresos_mensual) %>% filter(row_number() <= n()-1)
###    
###    serie1 <- ts(Ingresos_mensual$Ingresos, frequency=12, start=c(2007,1))
###    
###    modelo <- Arima(serie1, order=c(1,1,1), seasonal=c(0,1,1))
###    
###    x1 <- forecast(modelo, h = input$horizonte.1, level = 95)
###    
###    
###    pronosticos <- x1 %>%
###      as.data.frame() %>%
###      dplyr::rename(
###        Proyeccion =`Point Forecast`,
###        `Limite inferior` =`Lo 95`,
###        `Limite superior` =`Hi 95`
###      ) %>% mutate(
###        Proyeccion = round(Proyeccion/Millones,1),
###        `Limite inferior` = round(`Limite inferior`/Millones,1),
###        `Limite superior` = round(`Limite superior`/Millones,1)
###      )
###    
###    pronosticos <- pronosticos %>%
###      mutate(
###        `Proyeccion` = number(pronosticos$`Proyeccion`,           accuracy = .1, big.mark = ".", decimal.mark = ","),
###        `Limite inferior` = number(pronosticos$`Limite inferior`, accuracy = .1, big.mark = ".", decimal.mark = ","),
###        `Limite superior` = number(pronosticos$`Limite superior`, accuracy = .1, big.mark = ".", decimal.mark = ",")
###      )
###    
###    reactable(pronosticos) 
###    
###  })
###  
###  ##########################
###  #    PRONOSTICO ETS      #
###  ##########################
###  
###  output$tabla.forcast.2.2 <- renderReactable({ 
###    
###    Ingresos_mensual <- Ingresos_mensual %>%
###      dplyr::filter(n_Nivel==input$clasi.2_2)  %>% 
###      dplyr::group_by(Año, mes.cod, mes) %>% 
###      dplyr::summarise ("Ingresos"  = sum(Ingresos, na.rm = TRUE))  %>%
###      mutate(
###        Fecha = paste(Año,mes,sep ="-"))  %>%  
###      dplyr::filter(Ingresos>0)
###    
###
###    
###    Ingresos_mensual <-  as.data.frame(Ingresos_mensual) %>% filter(row_number() <= n()-1)
###    
###    serie1 <- ts(Ingresos_mensual$Ingresos, frequency=12, start=c(2007,1))
###    
###    ##########################
###    #          ETS MAM       #
###    ##########################
###    
###    modelo_2 <- ets(serie1,  model ="MAM",damped = TRUE)
###    x2 <- forecast(modelo_2, h = 24, level = 95)
###    
###    
###    pronosticos <- x2 %>%
###      as.data.frame() %>%
###      dplyr::rename(
###        Proyeccion =`Point Forecast`,
###        `Limite inferior` =`Lo 95`,
###        `Limite superior` =`Hi 95`
###      ) %>% mutate(
###        Proyeccion = round(Proyeccion/Millones,1),
###        `Limite inferior` = round(`Limite inferior`/Millones,1),
###        `Limite superior` = round(`Limite superior`/Millones,1)
###      )
###    
###    pronosticos <- pronosticos %>%
###      mutate(
###        `Proyeccion` = number(pronosticos$`Proyeccion`,           accuracy = .1, big.mark = ".", decimal.mark = ","),
###        `Limite inferior` = number(pronosticos$`Limite inferior`, accuracy = .1, big.mark = ".", decimal.mark = ","),
###        `Limite superior` = number(pronosticos$`Limite superior`, accuracy = .1, big.mark = ".", decimal.mark = ",")
###      )
###    
###    reactable(pronosticos) 
###    
###  })
###  
###  ##############################
###  #    PRONOSTICO AUTO.ARIMA   #
###  ##############################
###  
###  output$tabla.forcast.2.3 <- renderReactable({ 
###    
###    Ingresos_mensual <- Ingresos_mensual %>%
###      dplyr::filter(n_Nivel==input$clasi.2_2)  %>% 
###      dplyr::group_by(Año, mes.cod, mes) %>% 
###      dplyr::summarise ("Ingresos"  = sum(Ingresos, na.rm = TRUE))  %>%
###      mutate(
###        Fecha = paste(Año,mes,sep ="-"))  %>%  
###      dplyr::filter(Ingresos>0)
###    
###    Ingresos_mensual <-  as.data.frame(Ingresos_mensual) %>% filter(row_number() <= n()-1)
###    
###    serie1 <- ts(Ingresos_mensual$Ingresos, frequency=12, start=c(2007,1))
###    
###    ##########################
###    #          ETS MAM       #
###    ##########################
###    
###    modelo_3 <- auto.arima(serie1)
###    x3 <- forecast(modelo_3, h = 24, level = 95)
###    
###    
###    pronosticos <- x3 %>%
###      as.data.frame() %>%
###      dplyr::rename(
###        Proyeccion =`Point Forecast`,
###        `Limite inferior` =`Lo 95`,
###        `Limite superior` =`Hi 95`
###      ) %>% mutate(
###        Proyeccion = round(Proyeccion/Millones,1),
###        `Limite inferior` = round(`Limite inferior`/Millones,1),
###        `Limite superior` = round(`Limite superior`/Millones,1)
###      )
###    
###    pronosticos <- pronosticos %>%
###      mutate(
###        `Proyeccion` = number(pronosticos$`Proyeccion`,           accuracy = .1, big.mark = ".", decimal.mark = ","),
###        `Limite inferior` = number(pronosticos$`Limite inferior`, accuracy = .1, big.mark = ".", decimal.mark = ","),
###        `Limite superior` = number(pronosticos$`Limite superior`, accuracy = .1, big.mark = ".", decimal.mark = ",")
###      )
###    
###    
###    reactable(pronosticos)
###    
###  }) 
###  
###  ##############################
###  #    PRONOSTICO REGRESION    #
###  ##############################
###  
###  output$tabla.forcast.2.4 <- renderReactable({ 
###    
###    Ingresos_mensual <- Ingresos_mensual %>%
###      dplyr::filter(n_Nivel==input$clasi.2_2)  %>% 
###      dplyr::group_by(Año, mes.cod, mes) %>% 
###      dplyr::summarise ("Ingresos"  = sum(Ingresos, na.rm = TRUE))  %>%
###      mutate(
###        Fecha = paste(Año,mes,sep ="-"))  %>%  
###      dplyr::filter(Ingresos>0)
###    
###    Ingresos_mensual <-  as.data.frame(Ingresos_mensual) %>% filter(row_number() <= n()-1)
###    
###    serie1 <- ts(Ingresos_mensual$Ingresos, frequency=12, start=c(2007,1))
###    
###    ##########################
###    #          ETS MAM       #
###    ##########################
###    
###    modelo_4 <- tslm(serie1 ~ trend+season) 
###    x4 <- forecast(modelo_4, h = 24, level = 95)
###    
###    
###    pronosticos <- x4 %>%
###      as.data.frame() %>%
###      dplyr::rename(
###        Proyeccion =`Point Forecast`,
###        `Limite inferior` =`Lo 95`,
###        `Limite superior` =`Hi 95`
###      ) %>% mutate(
###        Proyeccion = round(Proyeccion/Millones,1),
###        `Limite inferior` = round(`Limite inferior`/Millones,1),
###        `Limite superior` = round(`Limite superior`/Millones,1)
###      )
###    
###    pronosticos <- pronosticos %>%
###      mutate(
###        `Proyeccion` = number(pronosticos$`Proyeccion`,           accuracy = .1, big.mark = ".", decimal.mark = ","),
###        `Limite inferior` = number(pronosticos$`Limite inferior`, accuracy = .1, big.mark = ".", decimal.mark = ","),
###        `Limite superior` = number(pronosticos$`Limite superior`, accuracy = .1, big.mark = ".", decimal.mark = ",")
###      )
###    
###    reactable(pronosticos)
###    
###  }) 
  
  #################################
  #            Impuestos          #
  #################################
  
  
  output$forecast3 <- renderHighchart({
    
    
    Impuesto <- Impuestos %>% dplyr::group_by(Año, mes.cod, mes) %>% 
      
      dplyr::summarise ("Ingresos"  = sum(!!input$variable2, na.rm = TRUE))  %>%
      dplyr::mutate(
        Fecha = paste(Año,mes,sep ="-")
      )
    
     # %>%      #   dplyr::filter(Ingresos>0)
    
    Impuesto$Ingresos[Impuesto$Ingresos <= 0] <- 1000
    
    
    Impuesto <- as.data.frame(Impuesto)  %>% filter(row_number() <= n()-1)
    
    serie1 <- ts(Impuesto$Ingresos, frequency=12, start=c(2007,1))
    
    ##########################
    #  ARIMA (1,1,1) (0,1,1) #
    ##########################
    
    modelo_1 <- Arima(serie1, order=c(1,1,1), seasonal=c(0,1,1))
    x1 <- forecast(modelo_1, h = input$horizonte.1, level = 95)
    
    
    ##########################
    #          ETS MAM       #
    ##########################
    
    modelo_2 <- ets(serie1,  model ="MAM",damped = TRUE)
    x2 <- forecast(modelo_2, h = input$horizonte.1, level = 95)
    
    ##########################
    #       Auto-ARIMA       #
    ##########################
    
    modelo_3 <- auto.arima(serie1)
    x3 <- forecast(modelo_3, h = input$horizonte.1, level = 95)
    
    ##########################
    #    Regresion TS       #
    ##########################  
    
    modelo_4 <- tslm(serie1 ~ trend+season) 
    x4 <- forecast(modelo_4, h = input$horizonte.1, level = 95)
    
    
    XF  <-hchart(serie1,name = "Ingresos") %>% 
      
      
      hc_add_series(name = "ARIMA",fitted(modelo_1), type = "line", color="green") %>% 
      hc_add_series(name = "F ARIMA", x1$mean, color="green") %>%
      
      hc_add_series(name = "ETS",fitted(modelo_2), type = "line", color="black") %>% 
      hc_add_series(name = "F ETs", x2$mean, color="black") %>%
      
      hc_add_series(name = "AUTO.ARIMA",fitted(modelo_3), type = "line", color="blue") %>%
      hc_add_series(name = "F AUTO.ARIMA", x3$mean, color="blue") %>%
      
      hc_add_series(name = "Regresión",fitted(modelo_4), type = "line", color="grey") %>%
      hc_add_series(name = "F Regresión", x4$mean, color="grey") %>%
      
      hc_title(text = "",
               margin = 20, align = "center",
               style = list(color = "#black", useHTML = TRUE, fontWeight = "bold")) %>%
      hc_yAxis(title = list(text = "En colones."),     labels = list(format = "{value}"))%>% 
      hc_xAxis(title = list(text = "Años") )  %>% 
      
      hc_tooltip(crosshairs = T,valueDecimals = 1, shared = TRUE, borderWidth = 5) %>%
      hc_chart(
        zoomType = "xy"
      )
    
    XF 
    
    
  })
  

  #################################################################
  # Talba con estadísticos de bondad y de ajuste por impuestos   #
  #################################################################
  
  output$GoF.3 <- renderReactable({
    
    Impuesto <- Impuestos %>% dplyr::group_by(Año, mes.cod, mes) %>% 
      
      dplyr::summarise ("Ingresos"  = sum(!!input$variable2, na.rm = TRUE))  %>%
      dplyr::mutate(
        Fecha = paste(Año,mes,sep ="-")
      ) %>%  
      dplyr::filter(Ingresos>0)
    
    
    Impuesto <- as.data.frame(Impuesto)  %>% filter(row_number() <= n()-1)
    
    serie1 <- ts(Impuesto$Ingresos, frequency=12, start=c(2007,1))
    
    ##########################
    #  ARIMA (1,1,1) (0,1,1) #
    ##########################
    
    modelo_1 <- Arima(serie1, order=c(1,1,1), seasonal=c(0,1,1))
    x1 <- forecast(modelo_1, h = input$horizonte.1, level = 95)
    
    
    ##########################
    #          ETS MAM       #
    ##########################
    
    modelo_2 <- ets(serie1,  model ="MAM",damped = TRUE)
    x2 <- forecast(modelo_2, h = input$horizonte.1, level = 95)
    
    ##########################
    #       Auto-ARIMA       #
    ##########################
    
    modelo_3 <- auto.arima(serie1)
    x3 <- forecast(modelo_3, h = input$horizonte.1, level = 95)
    
    ##########################
    #    Regresion TS       #
    ##########################  
    
    modelo_4 <- tslm(serie1 ~ trend+season) 
    x4 <- forecast(modelo_4, h = input$horizonte.1, level = 95)
    
    REGRESION <-round(c(accuracy(modelo_4)[c(2,3,5)],CV(modelo_4)[c(2,4)]),3)
    ETS.MAM<-round(c(accuracy(modelo_2)[c(2,3,5)],modelo_2$aic,modelo_2$bic),3)
    ARIMA<-round(c(accuracy(modelo_1)[c(2,3,5)],modelo_1$aic,modelo_1$bic),3)
    AUTO.ARIMA<-round(c(accuracy(modelo_3)[c(2,3,5)],modelo_3$aic,modelo_3$bic),3)
    
    cuadro<-rbind(Medida = c("RSME","MAE","MAPE","AIC","BIC"),
                  REGRESION,ETS.MAM,ARIMA,AUTO.ARIMA)
    
    colnames(cuadro)<-cuadro[1,]
    cuadro <- cuadro[-1,]
    
    reactable(cuadro)
    
  }) 
  
  
  ##########################
  #    PRONOSTICO ARIMA    #
  ##########################
  
  output$tabla.forcast.3.1 <- renderReactable({ 
    Impuesto <- Impuestos %>% dplyr::group_by(Año, mes.cod, mes) %>% 
      
      dplyr::summarise ("Ingresos"  = sum(!!input$variable2, na.rm = TRUE))  %>%
      dplyr::mutate(
        Fecha = paste(Año,mes,sep ="-")
      ) %>%  
      dplyr::filter(Ingresos>0)
    
    
    Impuesto <- as.data.frame(Impuesto)  %>% filter(row_number() <= n()-1)
    
    serie1 <- ts(Impuesto$Ingresos, frequency=12, start=c(2007,1))
    
    modelo <- Arima(serie1, order=c(1,1,1), seasonal=c(0,1,1))
    
    x1 <- forecast(modelo, h = input$horizonte.1, level = 95)
    
    
    pronosticos <- x1 %>%
      as.data.frame() %>%
      dplyr::rename(
        Proyección =`Point Forecast`,
        `Límite inferior` =`Lo 95`,
        `Límite superior` =`Hi 95`
      ) %>% mutate(
        Proyección = round(Proyección/Millones,1),
        `Límite inferior` = round(`Límite inferior`/Millones,1),
        `Límite superior` = round(`Límite superior`/Millones,1)
      )
    
    pronosticos <- pronosticos %>%
      mutate(
        `Proyección` = number(pronosticos$`Proyección`,           accuracy = .1, big.mark = ".", decimal.mark = ","),
        `Límite inferior` = number(pronosticos$`Límite inferior`, accuracy = .1, big.mark = ".", decimal.mark = ","),
        `Límite superior` = number(pronosticos$`Límite superior`, accuracy = .1, big.mark = ".", decimal.mark = ",")
      )
    
    reactable(pronosticos) 
    
  })
  
  ##########################
  #    PRONOSTICO ETS      #
  ##########################
  
  output$tabla.forcast.3.2 <- renderReactable({ 
    
    Impuesto <- Impuestos %>% dplyr::group_by(Año, mes.cod, mes) %>% 
      
      dplyr::summarise ("Ingresos"  = sum(!!input$variable2, na.rm = TRUE))  %>%
      dplyr::mutate(
        Fecha = paste(Año,mes,sep ="-")
      ) %>%  
      dplyr::filter(Ingresos>0)
    
    
    Impuesto <- as.data.frame(Impuesto)  %>% filter(row_number() <= n()-1)
    
    serie1 <- ts(Impuesto$Ingresos, frequency=input$horizonte.1, start=c(2007,1))
    
    ##########################
    #          ETS MAM       #
    ##########################
    
    modelo_2 <- ets(serie1,  model ="MAM",damped = TRUE)
    x2 <- forecast(modelo_2, h = input$horizonte.1, level = 95)
    
    
    pronosticos <- x2 %>%
      as.data.frame() %>%
      dplyr::rename(
        Proyección =`Point Forecast`,
        `Límite inferior` =`Lo 95`,
        `Límite superior` =`Hi 95`
      ) %>% mutate(
        Proyección = round(Proyección/Millones,1),
        `Límite inferior` = round(`Límite inferior`/Millones,1),
        `Límite superior` = round(`Límite superior`/Millones,1)
      )
    
    pronosticos <- pronosticos %>%
      mutate(
        `Proyección` = number(pronosticos$`Proyección`,           accuracy = .1, big.mark = ".", decimal.mark = ","),
        `Límite inferior` = number(pronosticos$`Límite inferior`, accuracy = .1, big.mark = ".", decimal.mark = ","),
        `Límite superior` = number(pronosticos$`Límite superior`, accuracy = .1, big.mark = ".", decimal.mark = ",")
      )
    
    reactable(pronosticos) 
    
  })
  
  ##############################
  #    PRONOSTICO AUTO.ARIMA   #
  ##############################
  
  output$tabla.forcast.3.3 <- renderReactable({ 
    
    Impuesto <- Impuestos %>% dplyr::group_by(Año, mes.cod, mes) %>% 
      
      dplyr::summarise ("Ingresos"  = sum(!!input$variable2, na.rm = TRUE))  %>%
      dplyr::mutate(
        Fecha = paste(Año,mes,sep ="-")
      ) %>%  
      dplyr::filter(Ingresos>0) 
    
    
    Impuesto <- as.data.frame(Impuesto)  %>% filter(row_number() <= n()-1)
    
    serie1 <- ts(Impuesto$Ingresos, frequency=input$horizonte.1, start=c(2007,1))
    
    ##########################
    #          ETS MAM       #
    ##########################
    
    modelo_3 <- auto.arima(serie1)
    x3 <- forecast(modelo_3, h = input$horizonte.1, level = 95)
    
    
    pronosticos <- x3 %>%
      as.data.frame() %>%
      dplyr::rename(
        Proyección =`Point Forecast`,
        `Límite inferior` =`Lo 95`,
        `Límite superior` =`Hi 95`
      ) %>% mutate(
        Proyección = round(Proyección/Millones,1),
        `Límite inferior` = round(`Límite inferior`/Millones,1),
        `Límite superior` = round(`Límite superior`/Millones,1)
      )
    
    pronosticos <- pronosticos %>%
      mutate(
        `Proyección` = number(pronosticos$`Proyección`,           accuracy = .1, big.mark = ".", decimal.mark = ","),
        `Límite inferior` = number(pronosticos$`Límite inferior`, accuracy = .1, big.mark = ".", decimal.mark = ","),
        `Límite superior` = number(pronosticos$`Límite superior`, accuracy = .1, big.mark = ".", decimal.mark = ",")
      )
    
    
    reactable(pronosticos)
    
  }) 
  
  ##############################
  #    PRONOSTICO REGRESION    #
  ##############################
  
  output$tabla.forcast.3.4 <- renderReactable({ 
    
    Impuesto <- Impuestos %>% dplyr::group_by(Año, mes.cod, mes) %>% 
      
      dplyr::summarise ("Ingresos"  = sum(!!input$variable2, na.rm = TRUE))  %>%
      dplyr::mutate(
        Fecha = paste(Año,mes,sep ="-")
      ) %>%  
      dplyr::filter(Ingresos>0)
    
    
    Impuesto <- as.data.frame(Impuesto)  %>% filter(row_number() <= n()-1)
    
    serie1 <- ts(Impuesto$Ingresos, frequency=input$horizonte.1, start=c(2007,1))
    
    ##########################
    #          ETS MAM       #
    ##########################
    
    modelo_4 <- tslm(serie1 ~ trend+season) 
    x4 <- forecast(modelo_4, h = input$horizonte.1, level = 95)
    
    
    pronosticos <- x4 %>%
      as.data.frame() %>%
      dplyr::rename(
        Proyección =`Point Forecast`,
        `Límite inferior` =`Lo 95`,
        `Límite superior` =`Hi 95`
      ) %>% mutate(
        Proyección = round(Proyección/Millones,1),
        `Límite inferior` = round(`Límite inferior`/Millones,1),
        `Límite superior` = round(`Límite superior`/Millones,1)
      )
    
    pronosticos <- pronosticos %>%
      mutate(
        `Proyección` = number(pronosticos$`Proyección`,           accuracy = .1, big.mark = ".", decimal.mark = ","),
        `Límite inferior` = number(pronosticos$`Límite inferior`, accuracy = .1, big.mark = ".", decimal.mark = ","),
        `Límite superior` = number(pronosticos$`Límite superior`, accuracy = .1, big.mark = ".", decimal.mark = ",")
      )
    
    reactable(pronosticos)
    
  }) 
  
  ############################################
  #   Filtros para el módulo   Avanzado      # 
  ############################################
  
  filtro_clase_2 <- reactive({
    Ingresos_mensual %>%
      filter(clase == input$f_clase_2)
    
  })
  
  
  filtro_subclase_2 <- reactive({
    if(input$f_subclase_2 == "Todo"){
      filtro_clase_2()
    } else {
      filtro_clase_2() %>% 
        filter(subclase == input$f_subclase_2)
    }
  })
  
  
  filtro_grupo_2 <- reactive({
    if(input$f_grupo_2 == "Todo"){
      filtro_subclase_2()
    } else {
      filtro_subclase_2() %>% 
        filter(grupo == input$f_grupo_2)
    }
  })
  
  filtro_subgrupo_2 <- reactive({
    if(input$f_subgrupo == "Todo"){
      filtro_grupo_2()
    } else {
      filtro_grupo_2() %>% 
        filter(subgrupo == input$f_subgrupo_2)
    }
  })
  
  filtro_partida_2 <- reactive({
    if(input$f_partida_2 == "Todo"){
      filtro_subgrupo_2()
    } else {
      filtro_subgrupo_2() %>% 
        filter(partida == input$f_partida_2)
    }
  })
  
  filtro_subpartida_2 <- reactive({
    if(input$f_subpartida_2 == "Todo"){
      filtro_partida_2()
    } else {
      filtro_partida_2() %>% 
        filter(subpartida == input$f_subpartida_2)
    }
  })
  
  filtro_renglon_2 <- reactive({
    if(input$f_renglon_2 == "Todo"){
      filtro_subpartida_2()
    } else {
      filtro_subpartida_2() %>% 
        filter(renglon == input$f_renglon_2)
    }
  })
  
  filtro_subrenglon_2 <- reactive({
    if(input$f_subrenglon_2 == "Todo"){
      filtro_renglon_2()
    } else {
      filtro_renglon_2() %>% 
        filter(subrenglon == input$f_subrenglon_2)
    }
  })
  
  # Ultimo # 
  
  filtros_2 <- eventReactive(input$select_3, {
    filtro_subrenglon_2()
  })
  
  
  output$forecast4 <- renderHighchart({
    
    
    Ingresos_mensual <- filtros_2() %>% dplyr::group_by(Año, mes.cod, mes) %>% 
      dplyr::summarise ("Ingresos"  = sum(Ingresos, na.rm = TRUE))  %>%
      mutate(
        Fecha = paste(Año,mes,sep ="-"))  %>%  
      dplyr::filter(Ingresos>0)
    
    
    Millones <- 1000000
    
    Ingresos_mensual <- Ingresos_mensual %>%
      dplyr::mutate(
        Ingresos = round(Ingresos/Millones,1)
      )
    
    
    Ingresos_mensual <- as.data.frame(Ingresos_mensual) %>% filter(row_number() <= n()-1)
    
    
    serie1 <- ts(Ingresos_mensual$Ingresos, frequency=12, start=c(2007,1))
    
    ##########################
    #  ARIMA (1,1,1) (0,1,1) #
    ##########################
    
    modelo_1 <- Arima(serie1, order=c(1,1,1), seasonal=c(0,1,1))
    x1 <- forecast(modelo_1, h = input$horizonte.4, level = 95)
    
    
    ##########################
    #          ETS MAM       #
    ##########################
    
    modelo_2 <- ets(serie1,  model ="MAM",damped = TRUE)
    x2 <- forecast(modelo_2, h = input$horizonte.4, level = 95)
    
    ##########################
    #       Auto-ARIMA       #
    ##########################
    
    modelo_3 <- auto.arima(serie1)
    x3 <- forecast(modelo_3, h = input$horizonte.4, level = 95)
    
    ##########################
    #    Regresion TS       #
    ##########################  
    
    modelo_4 <- tslm(serie1 ~ trend+season) 
    x4 <- forecast(modelo_4, h = input$horizonte.4, level = 95)
    
    
    XF  <-hchart(serie1,name = "Ingresos") %>% 
      
      
      hc_add_series(name = "ARIMA",fitted(modelo_1), type = "line", color="green") %>% 
      hc_add_series(name = "F ARIMA", x1$mean, color="green") %>%
      
      hc_add_series(name = "ETS",fitted(modelo_2), type = "line", color="black") %>% 
      hc_add_series(name = "F ETs", x2$mean, color="black") %>%
      
      hc_add_series(name = "AUTO.ARIMA",fitted(modelo_3), type = "line", color="blue") %>%
      hc_add_series(name = "F AUTO.ARIMA", x3$mean, color="blue") %>%
      
      hc_add_series(name = "Regresión",fitted(modelo_4), type = "line", color="grey") %>%
      hc_add_series(name = "F Regresión", x4$mean, color="grey") %>%
      
      hc_title(text = "",
               margin = 20, align = "center",
               style = list(color = "#black", useHTML = TRUE, fontWeight = "bold")) %>%
      hc_yAxis(title = list(text = "Millones de colones"),     labels = list(format = "{value}"))%>% 
      hc_xAxis(title = list(text = "Años") )  %>% 
      
      
      hc_tooltip(crosshairs = T,valueDecimals = 1, shared = TRUE, borderWidth = 5) %>%
      hc_chart(
        zoomType = "xy"
      )
    
    XF 
    
    
  })
  
  
  ################################################################
  # Talba con estadísticos de bondad y de ajuste por impuestos   #
  #################################################################
  
  output$GoF.4 <- renderReactable({
    
    
    Ingresos_mensual <- filtros_2() %>% dplyr::group_by(Año, mes.cod, mes) %>% 
      dplyr::summarise ("Ingresos"  = sum(Ingresos, na.rm = TRUE))  %>%
      mutate(
        Fecha = paste(Año,mes,sep ="-"))  %>%  
      dplyr::filter(Ingresos>0)
    
    Millones <- 1000000
    
    Ingresos_mensual <- Ingresos_mensual %>%
      dplyr::mutate(
        Ingresos = round(Ingresos/Millones,1)
      )
    
    
    Ingresos_mensual <- as.data.frame(Ingresos_mensual) %>% filter(row_number() <= n()-1)
    
    
    serie1 <- ts(Ingresos_mensual$Ingresos, frequency=12, start=c(2007,1))
    
    ##########################
    #  ARIMA (1,1,1) (0,1,1) #
    ##########################
    
    modelo_1 <- Arima(serie1, order=c(1,1,1), seasonal=c(0,1,1))
    x1 <- forecast(modelo_1, h = input$horizonte.4, level = 95)
    
    
    ##########################
    #          ETS MAM       #
    ##########################
    
    modelo_2 <- ets(serie1,  model ="MAM",damped = TRUE)
    x2 <- forecast(modelo_2, h = input$horizonte.4, level = 95)
    
    ##########################
    #       Auto-ARIMA       #
    ##########################
    
    modelo_3 <- auto.arima(serie1)
    x3 <- forecast(modelo_3, h = input$horizonte.4, level = 95)
    
    ##########################
    #    Regresion TS       #
    ##########################  
    
    modelo_4 <- tslm(serie1 ~ trend+season) 
    x4 <- forecast(modelo_4, h = input$horizonte.4, level = 95)
    
    REGRESION <-round(c(accuracy(modelo_4)[c(2,3,5)],CV(modelo_4)[c(2,4)]),3)
    ETS.MAM<-round(c(accuracy(modelo_2)[c(2,3,5)],modelo_2$aic,modelo_2$bic),3)
    ARIMA<-round(c(accuracy(modelo_1)[c(2,3,5)],modelo_1$aic,modelo_1$bic),3)
    AUTO.ARIMA<-round(c(accuracy(modelo_3)[c(2,3,5)],modelo_3$aic,modelo_3$bic),3)
    
    cuadro<-rbind(Medida = c("RSME","MAE","MAPE","AIC","BIC"),
                  REGRESION,ETS.MAM,ARIMA,AUTO.ARIMA)
    
    colnames(cuadro)<-cuadro[1,]
    cuadro <- cuadro[-1,]
    
    reactable(cuadro)
    
  })   
  

  ##########################
  #    PRONOSTICO ARIMA    #
  ##########################
  
  output$tabla.forcast.4.1 <- renderReactable({ 
    Ingresos_mensual <- filtros_2() %>% dplyr::group_by(Año, mes.cod, mes) %>% 
      dplyr::summarise ("Ingresos"  = sum(Ingresos, na.rm = TRUE))  %>%
      mutate(
        Fecha = paste(Año,mes,sep ="-"))  %>%  
      dplyr::filter(Ingresos>0)
    
    Millones <- 1000000
    
    Ingresos_mensual <- as.data.frame(Ingresos_mensual) %>% filter(row_number() <= n()-1)
    
    
    serie1 <- ts(Ingresos_mensual$Ingresos, frequency=12, start=c(2007,1))
    modelo <- Arima(serie1, order=c(1,1,1), seasonal=c(0,1,1))
    
    x1 <- forecast(modelo, h = input$horizonte.4, level = 95)
    
    
    pronosticos <- x1 %>%
      as.data.frame() %>%
      dplyr::rename(
        Proyección =`Point Forecast`,
        `Límite inferior` =`Lo 95`,
        `Límite superior` =`Hi 95`
      ) %>% mutate(
        Proyección = round(Proyección/Millones,1),
        `Límite inferior` = round(`Límite inferior`/Millones,1),
        `Límite superior` = round(`Límite superior`/Millones,1)
      )
    
    pronosticos <- pronosticos %>%
      mutate(
        `Proyección` = number(pronosticos$`Proyección`,           accuracy = .1, big.mark = ".", decimal.mark = ","),
        `Límite inferior` = number(pronosticos$`Límite inferior`, accuracy = .1, big.mark = ".", decimal.mark = ","),
        `Límite superior` = number(pronosticos$`Límite superior`, accuracy = .1, big.mark = ".", decimal.mark = ",")
      )
    
    reactable(pronosticos) 
    
  })
  
  ##########################
  #    PRONOSTICO ETS      #
  ##########################
  
  output$tabla.forcast.4.2 <- renderReactable({ 
    
    Ingresos_mensual <- filtros_2() %>% dplyr::group_by(Año, mes.cod, mes) %>% 
      dplyr::summarise ("Ingresos"  = sum(Ingresos, na.rm = TRUE))  %>%
      mutate(
        Fecha = paste(Año,mes,sep ="-"))  %>%  
      dplyr::filter(Ingresos>0)
    
    Millones <- 1000000
    
    
    Ingresos_mensual <- as.data.frame(Ingresos_mensual) %>% filter(row_number() <= n()-1)
    
    
    serie1 <- ts(Ingresos_mensual$Ingresos, frequency=input$horizonte.4, start=c(2007,1))
    
    ##########################
    #          ETS MAM       #
    ##########################
    
    modelo_2 <- ets(serie1,  model ="MAM",damped = TRUE)
    x2 <- forecast(modelo_2, h = input$horizonte.4, level = 95)
    
    
    pronosticos <- x2 %>%
      as.data.frame() %>%
      dplyr::rename(
        Proyección =`Point Forecast`,
        `Límite inferior` =`Lo 95`,
        `Límite superior` =`Hi 95`
      ) %>% mutate(
        Proyección = round(Proyección/Millones,1),
        `Límite inferior` = round(`Límite inferior`/Millones,1),
        `Límite superior` = round(`Límite superior`/Millones,1)
      )
    
    pronosticos <- pronosticos %>%
      mutate(
        `Proyección` = number(pronosticos$`Proyección`,           accuracy = .1, big.mark = ".", decimal.mark = ","),
        `Límite inferior` = number(pronosticos$`Límite inferior`, accuracy = .1, big.mark = ".", decimal.mark = ","),
        `Límite superior` = number(pronosticos$`Límite superior`, accuracy = .1, big.mark = ".", decimal.mark = ",")
      )
    
    reactable(pronosticos) 
    
  })
  
  ##############################
  #    PRONOSTICO AUTO.ARIMA   #
  ##############################
  
  output$tabla.forcast.4.3 <- renderReactable({ 
    
    Ingresos_mensual <- filtros_2() %>% dplyr::group_by(Año, mes.cod, mes) %>% 
      dplyr::summarise ("Ingresos"  = sum(Ingresos, na.rm = TRUE))  %>%
      mutate(
        Fecha = paste(Año,mes,sep ="-"))  %>%  
      dplyr::filter(Ingresos>0)
    
    Millones <- 1000000
    
    
    Ingresos_mensual <- as.data.frame(Ingresos_mensual) %>% filter(row_number() <= n()-1)
    
    
    serie1 <- ts(Ingresos_mensual$Ingresos, frequency=input$horizonte.4, start=c(2007,1))
    
    ##########################
    #          ETS MAM       #
    ##########################
    
    modelo_3 <- auto.arima(serie1)
    x3 <- forecast(modelo_3, h = input$horizonte.4, level = 95)
    
    
    pronosticos <- x3 %>%
      as.data.frame() %>%
      dplyr::rename(
        Proyección =`Point Forecast`,
        `Límite inferior` =`Lo 95`,
        `Límite superior` =`Hi 95`
      ) %>% mutate(
        Proyección = round(Proyección/Millones,1),
        `Límite inferior` = round(`Límite inferior`/Millones,1),
        `Límite superior` = round(`Límite superior`/Millones,1)
      )
    
    pronosticos <- pronosticos %>%
      mutate(
        `Proyección` = number(pronosticos$`Proyección`,           accuracy = .1, big.mark = ".", decimal.mark = ","),
        `Límite inferior` = number(pronosticos$`Límite inferior`, accuracy = .1, big.mark = ".", decimal.mark = ","),
        `Límite superior` = number(pronosticos$`Límite superior`, accuracy = .1, big.mark = ".", decimal.mark = ",")
      )
    
    
    reactable(pronosticos)
    
  }) 
  
  ##############################
  #    PRONOSTICO REGRESION    #
  ##############################
  
  output$tabla.forcast.4.4 <- renderReactable({ 
    
    Ingresos_mensual <- filtros_2() %>% dplyr::group_by(Año, mes.cod, mes) %>% 
      dplyr::summarise ("Ingresos"  = sum(Ingresos, na.rm = TRUE))  %>%
      mutate(
        Fecha = paste(Año,mes,sep ="-"))  %>%  
      dplyr::filter(Ingresos>0)
    
    Millones <- 1000000
    
    
    Ingresos_mensual <- as.data.frame(Ingresos_mensual) %>% filter(row_number() <= n()-1)
    
    
    serie1 <- ts(Ingresos_mensual$Ingresos, frequency=input$horizonte.4, start=c(2007,1))
    
    ##########################
    #          ETS MAM       #
    ##########################
    
    modelo_4 <- tslm(serie1 ~ trend+season) 
    x4 <- forecast(modelo_4, h = input$horizonte.4, level = 95)
    
    
    pronosticos <- x4 %>%
      as.data.frame() %>%
      dplyr::rename(
        Proyección =`Point Forecast`,
        `Límite inferior` =`Lo 95`,
        `Límite superior` =`Hi 95`
      ) %>% mutate(
        Proyección = round(Proyección/Millones,1),
        `Límite inferior` = round(`Límite inferior`/Millones,1),
        `Límite superior` = round(`Límite superior`/Millones,1)
      )
    
    pronosticos <- pronosticos %>%
      mutate(
        `Proyección` = number(pronosticos$`Proyección`,           accuracy = .1, big.mark = ".", decimal.mark = ","),
        `Límite inferior` = number(pronosticos$`Límite inferior`, accuracy = .1, big.mark = ".", decimal.mark = ","),
        `Límite superior` = number(pronosticos$`Límite superior`, accuracy = .1, big.mark = ".", decimal.mark = ",")
      )
    
    reactable(pronosticos)
    
  }) 
  
  
  ############################################
  #              Estacionalidad              #
  ############################################
  
  ################
  #    General   #
  ################
  
  #Paramétro mensual 
  
  output$Estacionalidad_1 <- renderHighchart({
    
    Estacionalidad <- Estacionalidad %>% dplyr::mutate( `Estacionalidad` = round(`Estacionalidad`*100,1),
                                                        `Estacionalidad promedio` = round(`Estacionalidad promedio`*100,1)
                                                        
    )
    
    HC_Estacionalidad  <- highchart() %>% 
      hc_title(text = "",
               margin = 20, align = "center",
               style = list(color = "#129", useHTML = TRUE)) %>% 
      hc_subtitle(text = "",
                  align = "right",
                  style = list(color = "#634", fontWeight = "bold")) %>%
      hc_credits(enabled = TRUE, # add credits
                 text = "" # href = "www.cgr.go.cr"
      ) %>%
      hc_legend(align = "left", verticalAlign = "top",
                layout = "vertical", x = 0, y = 100) %>%
      hc_exporting(enabled = TRUE) %>% 
      hc_xAxis(categories = Estacionalidad$Fecha) %>% 
      hc_add_series(name = "Estacionalidad", data = Estacionalidad$Estacionalidad) %>% 
      hc_add_series(name = "Estacionalidad promedio", data = Estacionalidad$`Estacionalidad promedio`) %>% 
      hc_chart(zoomType = "xy")
    
    HC_Estacionalidad
    
  })
  
  #Parámetro acumulado

  output$Estacionalidad_2 <- renderHighchart({  
  
    Estacionalidad <- Estacionalidad %>% dplyr::mutate( `Estacionalidad_acumulada` = round(`Estacionalidad_acumulada`*100,1),
                                                        `Estacionalidad acumulada promedio` = round(`Estacionalidad acumulada promedio`*100,1)
                                                  
    ) 
    
    HC_Estacionalidad_acumulada  <- highchart() %>% 
      hc_title(text = "",
               margin = 20, align = "center",
               style = list(color = "#129", useHTML = TRUE)) %>% 
      hc_subtitle(text = "",
                  align = "right",
                  style = list(color = "#634", fontWeight = "bold")) %>%
      hc_credits(enabled = TRUE, # add credits
                 text = "" # href = "www.cgr.go.cr"
      ) %>%
      hc_legend(align = "left", verticalAlign = "top",
                layout = "vertical", x = 0, y = 100) %>%
      hc_exporting(enabled = TRUE) %>% 
      hc_xAxis(categories = Estacionalidad$Fecha) %>% 
      hc_add_series(name = "Estacionalidad acumulada", data = Estacionalidad$`Estacionalidad_acumulada`) %>% 
      hc_add_series(name = "Estacionalidad acumulada promedio", data = Estacionalidad$`Estacionalidad acumulada promedio`) %>% 
      hc_chart(zoomType = "xy")
    
    
    HC_Estacionalidad_acumulada   
    
  })
  
  # Gráfico de control
  
  output$Estacionalidad_3 <- renderHighchart({  
    
    Estacionalidad <- Estacionalidad %>% dplyr::mutate( Estacionalidad = round(`Estacionalidad`*100,1),
                                                        'Limite inferior' = 15.8 ,
                                                        'Limite superior' = 0.8
    ) 
    
    HC_Estacionalidad_mensual  <- highchart() %>% 
      hc_title(text = "",
               margin = 20, align = "center",
               style = list(color = "#129", useHTML = TRUE)) %>% 
      hc_subtitle(text = "",
                  align = "right",
                  style = list(color = "#634", fontWeight = "bold")) %>%
      hc_credits(enabled = TRUE, # add credits
                 text = "" # href = "www.cgr.go.cr"
      ) %>%
      hc_legend(align = "left", verticalAlign = "top",
                layout = "vertical", x = 0, y = 100) %>%
      hc_exporting(enabled = TRUE) %>% 
      hc_xAxis(categories = Estacionalidad$Fecha) %>% 
      hc_add_series(name = "Estacionalidad", data = Estacionalidad$`Estacionalidad`) %>% 
      hc_add_series(name = "Limite inferior", data = Estacionalidad$`Limite inferior`) %>% 
      hc_add_series(name = "Limite superior", data = Estacionalidad$`Limite superior`) %>% 
      hc_chart(zoomType = "xy")
    
    
    HC_Estacionalidad_mensual   
    
  })  
  
  #############################
  #    Ingresos Tributarios   #
  #############################
  
  #Paramétro mensual 
  
  output$Estacionalidad_IT_1 <- renderHighchart({
    
    IT_Estacionalidad <- IT_Estacionalidad %>% dplyr::mutate( `Estacionalidad` = round(`Estacionalidad`*100,1),
                                                        `Estacionalidad promedio` = round(`Estacionalidad promedio`*100,1)
                                                        
    )
    
    HC_Estacionalidad  <- highchart() %>% 
      hc_title(text = "",
               margin = 20, align = "center",
               style = list(color = "#129", useHTML = TRUE)) %>% 
      hc_subtitle(text = "",
                  align = "right",
                  style = list(color = "#634", fontWeight = "bold")) %>%
      hc_credits(enabled = TRUE, # add credits
                 text = "" # href = "www.cgr.go.cr"
      ) %>%
      hc_legend(align = "left", verticalAlign = "top",
                layout = "vertical", x = 0, y = 100) %>%
      hc_exporting(enabled = TRUE) %>% 
      hc_xAxis(categories = Estacionalidad$Fecha) %>% 
      hc_add_series(name = "Estacionalidad", data = IT_Estacionalidad$Estacionalidad) %>% 
      hc_add_series(name = "Estacionalidad promedio", data = IT_Estacionalidad$`Estacionalidad promedio`) %>% 
      hc_chart(zoomType = "xy")
    
    HC_Estacionalidad
    
  })
  
  #Parámetro acumulado
  
  output$Estacionalidad_IT_2 <- renderHighchart({  
    
    IT_Estacionalidad <- IT_Estacionalidad %>% dplyr::mutate( `Estacionalidad_acumulada` = round(`Estacionalidad_acumulada`*100,1),
                                                        `Estacionalidad acumulada promedio` = round(`Estacionalidad acumulada promedio`*100,1)
                                                        
    ) 
    
    HC_Estacionalidad_acumulada  <- highchart() %>% 
      hc_title(text = "",
               margin = 20, align = "center",
               style = list(color = "#129", useHTML = TRUE)) %>% 
      hc_subtitle(text = "",
                  align = "right",
                  style = list(color = "#634", fontWeight = "bold")) %>%
      hc_credits(enabled = TRUE, # add credits
                 text = "" # href = "www.cgr.go.cr"
      ) %>%
      hc_legend(align = "left", verticalAlign = "top",
                layout = "vertical", x = 0, y = 100) %>%
      hc_exporting(enabled = TRUE) %>% 
      hc_xAxis(categories = Estacionalidad$Fecha) %>% 
      hc_add_series(name = "Estacionalidad acumulada", data = IT_Estacionalidad$`Estacionalidad_acumulada`) %>% 
      hc_add_series(name = "Estacionalidad acumulada promedio", data = IT_Estacionalidad$`Estacionalidad acumulada promedio`) %>% 
      hc_chart(zoomType = "xy")
    
    
    HC_Estacionalidad_acumulada   
    
  })
  
  # Gráfico de control
  
  output$Estacionalidad_IT_3 <- renderHighchart({  
    
    IT_Estacionalidad <- IT_Estacionalidad %>% dplyr::mutate( Estacionalidad = round(`Estacionalidad`*100,1),
                                                        'Limite inferior' = 15.6 ,
                                                        'Limite superior' = 0.9
    ) 
    
    HC_Estacionalidad_mensual  <- highchart() %>% 
      hc_title(text = "",
               margin = 20, align = "center",
               style = list(color = "#129", useHTML = TRUE)) %>% 
      hc_subtitle(text = "",
                  align = "right",
                  style = list(color = "#634", fontWeight = "bold")) %>%
      hc_credits(enabled = TRUE, # add credits
                 text = "" # href = "www.cgr.go.cr"
      ) %>%
      hc_legend(align = "left", verticalAlign = "top",
                layout = "vertical", x = 0, y = 100) %>%
      hc_exporting(enabled = TRUE) %>% 
      hc_xAxis(categories = Estacionalidad$Fecha) %>% 
      hc_add_series(name = "Estacionalidad", data = IT_Estacionalidad$`Estacionalidad`) %>% 
      hc_add_series(name = "Limite inferior", data = IT_Estacionalidad$`Limite inferior`) %>% 
      hc_add_series(name = "Limite superior", data = IT_Estacionalidad$`Limite superior`) %>% 
      hc_chart(zoomType = "xy")
    
    
    HC_Estacionalidad_mensual   
    
  })  
  
  #############################
  #         Impuestos         #
  #############################
  
  # Avolución del impuesto
  
  
  output$Estacionalidad_Impuestos_1 <- renderHighchart({
    
    #    filtros <- eventReactive(input$select_2, {
    #      Impuestos
    #    })
    
    
    Impuesto <- Impuestos %>% dplyr::group_by(Año, mes.cod, mes) %>% 
      
      dplyr::summarise ("Ingresos"  = sum(!!input$variable5, na.rm = TRUE))  %>%
      dplyr::mutate(
        Fecha = paste(Año,mes,sep ="-")
      ) 
    
    #   Impuesto <- data.frame(Impuesto)
    
    g_Impuesto  <- highchart() %>% 
      hc_title(text = "",
               margin = 20, align = "center",
               style = list(color = "#129", useHTML = TRUE)) %>% 
      hc_subtitle(text = "",
                  align = "right",
                  style = list(color = "#634", fontWeight = "bold")) %>%
      hc_credits(enabled = TRUE, # add credits
                 text = "" # href = "www.cgr.go.cr"
      ) %>%
      hc_legend(align = "left", verticalAlign = "top",
                layout = "vertical", x = 0, y = 100) %>%
      hc_exporting(enabled = TRUE) %>% 
      hc_xAxis(categories = Impuesto$Fecha) %>% 
      hc_add_series(name = "Ingresos", data = Impuesto$Ingresos) %>% 
      hc_yAxis(title = list(text = "Millones de colones"),     labels = list(format = "{value}"))%>% 
      hc_xAxis(title = list(text = "Año-mes") )  %>% 
      hc_chart(zoomType = "xy")
    
    g_Impuesto
    
  })
  
  # Paramétro mensual 
  
  output$Estacionalidad_Impuestos_2 <- renderHighchart({
    
    #    filtros <- eventReactive(input$select_2, {
    #      Impuestos
    #    })
    
    
    Estacionalidad_Impuesto <- Impuestos %>%  dplyr::filter(cod_fuentefinanciacion_3!=0)  %>%  
      dplyr::group_by(Año, mes.cod, mes) %>% 
      
      dplyr::summarise ("Ingresos"  = sum(!!input$variable5, na.rm = TRUE))  %>%
      dplyr::mutate(
        Fecha = paste(Año,mes,sep ="-")
      ) 
    
    Estacionalidad_Impuesto <- Estacionalidad_Impuesto %>%
      dplyr::group_by(Año) %>%                                                            
      dplyr::mutate( 
        Ingreso_acumulado = cumsum(Ingresos), 
        Sumarecaudacion = sum(`Ingresos`),
        Estacionalidad = Ingresos/Sumarecaudacion,
        Estacionalidad_acumulada = cumsum(Estacionalidad)
      )  %>%
      dplyr::arrange(Año,mes.cod)
    
    Estacioanalidades_promedio <-  Estacionalidad_Impuesto %>% 
      dplyr::group_by(mes) %>%  
      dplyr::summarise('Estacionalidad promedio' =  mean(Estacionalidad, na.rm = TRUE),
                       'Estacionalidad acumulada promedio' = mean(Estacionalidad_acumulada, na.rm = TRUE))
    
    
    
    Estacionalidad_Impuesto <-  Estacionalidad_Impuesto %>%  dplyr::full_join(Estacioanalidades_promedio) 
    Estacionalidad_Impuesto <- as.data.frame(Estacionalidad_Impuesto)
    
    
    Estacionalidad_Impuesto <- Estacionalidad_Impuesto %>% dplyr::mutate( `Estacionalidad` = round(`Estacionalidad`*100,1),
                                                                          `Estacionalidad promedio` = round(`Estacionalidad promedio`*100,1)
                                                                          
    )
    
    HC_Estacionalidad  <- highchart() %>% 
      hc_title(text = "",
               margin = 20, align = "center",
               style = list(color = "#129", useHTML = TRUE)) %>% 
      hc_subtitle(text = "",
                  align = "right",
                  style = list(color = "#634", fontWeight = "bold")) %>%
      hc_credits(enabled = TRUE, # add credits
                 text = "" # href = "www.cgr.go.cr"
      ) %>%
      hc_legend(align = "left", verticalAlign = "top",
                layout = "vertical", x = 0, y = 100) %>%
      hc_exporting(enabled = TRUE) %>% 
      hc_xAxis(categories = Estacionalidad$Fecha) %>% 
      hc_add_series(name = "Estacionalidad", data = Estacionalidad_Impuesto$Estacionalidad) %>% 
      hc_add_series(name = "Estacionalidad promedio", data = Estacionalidad_Impuesto$`Estacionalidad promedio`) %>% 
      hc_chart(zoomType = "xy")
    
    HC_Estacionalidad
  
    
  })
    
  # Parámetro acumulado
  
  output$Estacionalidad_Impuestos_3 <- renderHighchart({
    
    #    filtros <- eventReactive(input$select_2, {
    #      Impuestos
    #    })
    
    
    Estacionalidad_Impuesto <- Impuestos %>%  dplyr::filter(cod_fuentefinanciacion_3!=0)  %>%  
      dplyr::group_by(Año, mes.cod, mes) %>% 
      
      dplyr::summarise ("Ingresos"  = sum(!!input$variable5, na.rm = TRUE))  %>%
      dplyr::mutate(
        Fecha = paste(Año,mes,sep ="-")
      ) 
    
    Estacionalidad_Impuesto <- Estacionalidad_Impuesto %>%
      dplyr::group_by(Año) %>%                                                            
      dplyr::mutate( 
        Ingreso_acumulado = cumsum(Ingresos), 
        Sumarecaudacion = sum(`Ingresos`),
        Estacionalidad = Ingresos/Sumarecaudacion,
        Estacionalidad_acumulada = cumsum(Estacionalidad)
      )  %>%
      dplyr::arrange(Año,mes.cod)
    
    Estacioanalidades_promedio <-  Estacionalidad_Impuesto %>% 
      dplyr::group_by(mes) %>%  
      dplyr::summarise('Estacionalidad promedio' =  mean(Estacionalidad, na.rm = TRUE),
                       'Estacionalidad acumulada promedio' = mean(Estacionalidad_acumulada, na.rm = TRUE))
    
    Estacionalidad_Impuesto <-  Estacionalidad_Impuesto %>%  dplyr::full_join(Estacioanalidades_promedio) 
    Estacionalidad_Impuesto <- as.data.frame(Estacionalidad_Impuesto)
    
    Estacionalidad_Impuesto <- Estacionalidad_Impuesto %>% dplyr::mutate( `Estacionalidad_acumulada` = round(`Estacionalidad_acumulada`*100,1),
                                                                          `Estacionalidad acumulada promedio` = round(`Estacionalidad acumulada promedio`*100,1))
    
    HC_Estacionalidad_acumulada  <- highchart() %>% 
      hc_title(text = "",
               margin = 20, align = "center",
               style = list(color = "#129", useHTML = TRUE)) %>% 
      hc_subtitle(text = "",
                  align = "right",
                  style = list(color = "#634", fontWeight = "bold")) %>%
      hc_credits(enabled = TRUE, # add credits
                 text = "" # href = "www.cgr.go.cr"
      ) %>%
      hc_legend(align = "left", verticalAlign = "top",
                layout = "vertical", x = 0, y = 100) %>%
      hc_exporting(enabled = TRUE) %>% 
      hc_xAxis(categories = Estacionalidad$Fecha) %>% 
      hc_add_series(name = "Estacionalidad acumulada", data = Estacionalidad$`Estacionalidad_acumulada`) %>% 
      hc_add_series(name = "Estacionalidad acumulada promedio", data = Estacionalidad$`Estacionalidad acumulada promedio`) %>% 
      hc_chart(zoomType = "xy")
    
    
    HC_Estacionalidad_acumulada        
    
 
  })
  
  # Gráfico de control
  
  ###########################################
  #              Descargar archivos         # 
  ###########################################

  
  #### ARchivo de los ingresos
  
  observeEvent(input$show2, {
    showModal(modalDialog(
      title = "Descargar archivo de los ingresos en el tiempo ",br(),
      "Seleccione el tipo de archivo de descarga", br(),
      br(),
      downloadButton("download2.1","Archivo .csv"),
      br(),
      br(),
      downloadButton("download2.2","Archivo .txt"),
      footer = modalButton("Cerrar"),
      easyClose = TRUE)
    )
    
  })
  
  
  
  output$download2.1 <- downloadHandler(
    
    
    filename = function() {
      paste("Ingresos-", Sys.Date(), ".csv", sep="")
    },
    
    content = function(file) {
      write.csv(Ingresos_0, file)
    }
  )
  
  output$download2.2 <- downloadHandler(
    
    filename = function() {
      paste("Ingresos-", Sys.Date(), ".txt", sep="")
    },
    content = function(file) {
      write.table(Ingresos_0, file)
    }
  )
  
    
  #### ARchivo de impuestos
  
  observeEvent(input$show1, {
    showModal(modalDialog(
      title = "Descargar archivo de los Impuestos en el tiempo ",br(),
      "Seleccione el tipo de archivo de descarga", br(),
      br(),
      downloadButton("download1","Archivo .csv"),
      br(),
      br(),
      downloadButton("download2","Archivo .txt"),
      footer = modalButton("Cerrar"),
      easyClose = TRUE)
    )
    
  })
  
  output$download1 <- downloadHandler(
    
    
    filename = function() {
      paste("Impuesto-", Sys.Date(), ".csv", sep="")
    },
    
    content = function(file) {
      write.csv(Impuestos, file)
    }
  )
  
  output$download2 <- downloadHandler(
    filename = function() {
      paste("Impuesto-", Sys.Date(), ".txt", sep="")
    },
    content = function(file) {
      write.table(Impuestos, file)
    }
  )
  
  
  
  
}




